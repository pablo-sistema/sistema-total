import io
from datetime import date, timedelta
from urllib.parse import quote

from fastapi import FastAPI, Request, Form, HTTPException
from fastapi.responses import HTMLResponse, RedirectResponse, StreamingResponse
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles

from sqlalchemy.orm import joinedload
from sqlalchemy import func
from sqlalchemy.exc import IntegrityError

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib import colors

from database import engine, Base, SessionLocal
from models import Propiedad, Inquilino, Contrato, Cargo, Pago, Lectura

app = FastAPI()
app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")
Base.metadata.create_all(bind=engine)

DIA_CORTE = 25

# ----------------------------
# Helpers fechas
# ----------------------------
def ultimo_dia_mes(anio: int, mes: int) -> int:
    d = date(anio, mes, 28) + timedelta(days=4)
    return (d - timedelta(days=d.day)).day

def sumar_mes(fecha: date) -> date:
    anio = fecha.year + (1 if fecha.month == 12 else 0)
    mes = 1 if fecha.month == 12 else fecha.month + 1
    dia = min(fecha.day, ultimo_dia_mes(anio, mes))
    return date(anio, mes, dia)

def fecha_corte(fecha_inicio: date) -> date:
    if fecha_inicio.day <= DIA_CORTE:
        return date(fecha_inicio.year, fecha_inicio.month, DIA_CORTE)
    prox = sumar_mes(date(fecha_inicio.year, fecha_inicio.month, 1))
    return date(prox.year, prox.month, DIA_CORTE)

def calcular_prorrata_mensual(monto_mensual: float, fecha_inicio: date):
    corte = fecha_corte(fecha_inicio)
    dias = (corte - fecha_inicio).days
    if dias < 0:
        dias = 0
    diario = float(monto_mensual) / 30.0
    monto = round(dias * diario, 2)
    return dias, monto, corte

def calcular_alquiler_diario(monto_mensual: float, fecha_inicio: date, fecha_fin: date):
    # Cobro por días: (monto_mensual/30) * días
    if fecha_fin < fecha_inicio:
        raise HTTPException(400, "Fecha fin no puede ser menor que fecha inicio")
    dias = (fecha_fin - fecha_inicio).days + 1  # inclusivo
    diario = float(monto_mensual) / 30.0
    total = round(diario * dias, 2)
    return dias, total


# ----------------------------
# Cargos automáticos (mensual día 25)
# ----------------------------
def generar_cargos_mensuales(db, hoy: date):
    """Crea ALQUILER_MENSUAL el día 25 (o después) SOLO para contratos tipo 'mensual' activos."""
    if hoy.day < DIA_CORTE:
        return

    venc = date(hoy.year, hoy.month, DIA_CORTE)
    periodo = f"{venc.year:04d}-{venc.month:02d}"

    contratos = db.query(Contrato).filter(
        Contrato.estado == "Activo",
        (Contrato.tipo_alquiler == None) | (Contrato.tipo_alquiler == "mensual")
    ).all()

    for contrato in contratos:
        existe = db.query(Cargo).filter(
            Cargo.contrato_id == contrato.id,
            Cargo.concepto == "ALQUILER_MENSUAL",
            Cargo.periodo == periodo
        ).first()
        if existe:
            continue

        db.add(Cargo(
            contrato_id=contrato.id,
            concepto="ALQUILER_MENSUAL",
            periodo=periodo,
            monto=float(contrato.monto_mensual or 0.0),
            vencimiento=venc,
            estado="Pendiente",
            pagado_acumulado=0.0
        ))


# ----------------------------
# Cargo saldo/estado automático
# ----------------------------
def recalcular_cargo(db, cargo_id: int):
    cargo = db.query(Cargo).filter(Cargo.id == cargo_id).first()
    if not cargo:
        return

    total_pagado = db.query(func.coalesce(func.sum(Pago.monto), 0.0)).filter(
        Pago.cargo_id == cargo_id
    ).scalar() or 0.0

    cargo.pagado_acumulado = float(total_pagado)

    total = float(cargo.monto or 0.0)
    if cargo.pagado_acumulado <= 0:
        cargo.estado = "Pendiente"
    elif cargo.pagado_acumulado < total:
        cargo.estado = "Parcial"
    else:
        cargo.estado = "Pagado"


# ----------------------------
# Recibo PDF
# ----------------------------
def serie_recibo(pago_id: int, fecha_pago: date) -> str:
    return f"R-{fecha_pago.year}-{pago_id:06d}"

def generar_recibo_pdf_pro(pago: Pago) -> bytes:
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    w, h = A4

    M = 16 * mm
    top = h - M
    left = M
    right = w - M

    c.setFillColor(colors.HexColor("#111827"))
    c.rect(left, top - 28*mm, right - left, 28*mm, fill=1, stroke=0)

    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 18)
    c.drawString(left + 6*mm, top - 12*mm, "RECIBO DE PAGO - CREDIMAS")

    serie = serie_recibo(pago.id, pago.fecha_pago)
    c.setStrokeColor(colors.white)
    c.roundRect(right - 58*mm, top - 22*mm, 56*mm, 16*mm, 3*mm, stroke=1, fill=0)
    c.setFont("Helvetica-Bold", 11)
    c.drawRightString(right - 4*mm, top - 12*mm, serie)

    c.setFillColor(colors.black)
    c.setStrokeColor(colors.black)

    y = top - 36*mm

    inq = pago.contrato.inquilino
    prop = pago.contrato.propiedad
    cargo = pago.cargo

    c.setFont("Helvetica-Bold", 11)
    c.drawString(left, y, "Cliente / Contrato")
    y -= 6*mm

    c.setFont("Helvetica", 10)
    c.drawString(left, y, f"Inquilino: {inq.nombre}  (DNI: {inq.dni})")
    y -= 5*mm
    c.drawString(left, y, f"Propiedad: {prop.tipo} - {prop.numero}")
    y -= 5*mm
    c.drawString(left, y, f"Fecha pago: {pago.fecha_pago}   Método: {pago.metodo}")
    y -= 8*mm

    c.setStrokeColor(colors.HexColor("#D1D5DB"))
    c.line(left, y, right, y)
    c.setStrokeColor(colors.black)
    y -= 10*mm

    c.setFont("Helvetica-Bold", 11)
    c.drawString(left, y, "Detalle")
    y -= 7*mm

    c.setFillColor(colors.HexColor("#F3F4F6"))
    c.rect(left, y - 6*mm, right - left, 8*mm, fill=1, stroke=0)
    c.setFillColor(colors.black)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(left + 2*mm, y - 3*mm, "Concepto")
    c.drawString(left + 70*mm, y - 3*mm, "Periodo")
    c.drawString(left + 100*mm, y - 3*mm, "Venc.")
    c.drawRightString(right - 2*mm, y - 3*mm, "Importe")
    y -= 10*mm

    concepto = cargo.concepto if cargo else "-"
    periodo = cargo.periodo if cargo else "-"
    venc = str(cargo.vencimiento) if cargo else "-"

    c.setFont("Helvetica", 10)
    c.drawString(left + 2*mm, y, str(concepto))
    c.drawString(left + 70*mm, y, str(periodo))
    c.drawString(left + 100*mm, y, str(venc))
    c.drawRightString(right - 2*mm, y, f"S/ {pago.monto:.2f}")
    y -= 12*mm

    c.setStrokeColor(colors.HexColor("#111827"))
    c.line(left, y, right, y)
    y -= 8*mm
    c.setFont("Helvetica-Bold", 12)
    c.drawRightString(right - 2*mm, y, f"TOTAL PAGADO: S/ {pago.monto:.2f}")

    c.setFont("Helvetica-Oblique", 8)
    c.setFillColor(colors.HexColor("#6B7280"))
    c.drawString(left, 12*mm, "Documento generado automáticamente por el sistema.")
    c.setFillColor(colors.black)

    c.showPage()
    c.save()
    buffer.seek(0)
    return buffer.read()


# =========================================================
# DASHBOARD
# =========================================================
@app.get("/", response_class=HTMLResponse)
@app.get("/dashboard", response_class=HTMLResponse)
def dashboard(request: Request):
    db = SessionLocal()
    try:
        hoy = date.today()
        periodo_actual = f"{hoy.year:04d}-{hoy.month:02d}"

        # genera cargos mensuales si corresponde (día 25 o más)
        generar_cargos_mensuales(db, hoy)
        db.commit()

        # KPIs base
        total_propiedades = db.query(func.count(Propiedad.id)).scalar() or 0
        ocupadas = db.query(func.count(Propiedad.id)).filter(Propiedad.estado == "ocupado").scalar() or 0
        libres = db.query(func.count(Propiedad.id)).filter(Propiedad.estado == "libre").scalar() or 0

        total_inquilinos = db.query(func.count(Inquilino.id)).scalar() or 0
        contratos_activos = db.query(func.count(Contrato.id)).filter(Contrato.estado == "Activo").scalar() or 0

        # ====== COBRANZA DEL PERIODO ACTUAL ======
        # total pendiente del periodo actual
        pendiente_mes = db.query(
            func.coalesce(
                func.sum(Cargo.monto - func.coalesce(Cargo.pagado_acumulado, 0.0)),
                0.0
            )
        ).filter(
            Cargo.periodo == periodo_actual,
            Cargo.estado.in_(["Pendiente", "Parcial"])
        ).scalar() or 0.0

        # total cobrado en el mes actual (por fecha de pago)
        inicio_mes = date(hoy.year, hoy.month, 1)
        cobrado_mes = db.query(
            func.coalesce(func.sum(Pago.monto), 0.0)
        ).filter(
            Pago.fecha_pago >= inicio_mes,
            Pago.fecha_pago <= hoy
        ).scalar() or 0.0

        # morosos reales
        morosos = db.query(func.count(Cargo.id)).filter(
            Cargo.estado.in_(["Pendiente", "Parcial"]),
            Cargo.vencimiento < hoy
        ).scalar() or 0

        # cargos vencidos
        cargos_morosos = db.query(Cargo).options(
            joinedload(Cargo.contrato).joinedload(Contrato.inquilino),
            joinedload(Cargo.contrato).joinedload(Contrato.propiedad),
        ).filter(
            Cargo.estado.in_(["Pendiente", "Parcial"]),
            Cargo.vencimiento < hoy
        ).order_by(Cargo.vencimiento.asc()).limit(20).all()

        # próximos vencimientos
        proximos_vencimientos = db.query(Cargo).options(
            joinedload(Cargo.contrato).joinedload(Contrato.inquilino),
            joinedload(Cargo.contrato).joinedload(Contrato.propiedad),
        ).filter(
            Cargo.estado.in_(["Pendiente", "Parcial"]),
            Cargo.vencimiento >= hoy
        ).order_by(Cargo.vencimiento.asc()).limit(15).all()

        # top deudores
        top_deudores = db.query(
            Inquilino.id.label("inquilino_id"),
            Inquilino.nombre.label("nombre"),
            Inquilino.whatsapp.label("whatsapp"),
            func.coalesce(
                func.sum(Cargo.monto - func.coalesce(Cargo.pagado_acumulado, 0.0)),
                0.0
            ).label("deuda")
        ).join(Contrato, Contrato.inquilino_id == Inquilino.id) \
         .join(Cargo, Cargo.contrato_id == Contrato.id) \
         .filter(Cargo.estado.in_(["Pendiente", "Parcial"])) \
         .group_by(Inquilino.id, Inquilino.nombre, Inquilino.whatsapp) \
         .order_by(
             func.sum(Cargo.monto - func.coalesce(Cargo.pagado_acumulado, 0.0)).desc()
         ) \
         .limit(10).all()

        # pagos recientes
        pagos_recientes = db.query(Pago).options(
            joinedload(Pago.contrato).joinedload(Contrato.inquilino),
            joinedload(Pago.contrato).joinedload(Contrato.propiedad),
            joinedload(Pago.cargo)
        ).order_by(Pago.fecha_pago.desc(), Pago.id.desc()).limit(10).all()

        return templates.TemplateResponse("dashboard.html", {
            "request": request,
            "hoy": hoy,
            "periodo_actual": periodo_actual,

            "total_propiedades": total_propiedades,
            "ocupadas": ocupadas,
            "libres": libres,
            "total_inquilinos": total_inquilinos,
            "contratos_activos": contratos_activos,

            "pendiente_mes": round(float(pendiente_mes), 2),
            "cobrado_mes": round(float(cobrado_mes), 2),
            "morosos": morosos,

            "cargos_morosos": cargos_morosos,
            "proximos_vencimientos": proximos_vencimientos,
            "top_deudores": top_deudores,
            "pagos_recientes": pagos_recientes,
        })
    finally:
        db.close()


# =========================================================
# PROPIEDADES (CRUD)
# =========================================================
@app.get("/propiedades", response_class=HTMLResponse)
def propiedades_listar(request: Request):
    db = SessionLocal()
    try:
        propiedades = db.query(Propiedad).order_by(Propiedad.id.desc()).all()
        return templates.TemplateResponse("index.html", {"request": request, "propiedades": propiedades})
    finally:
        db.close()

@app.post("/propiedades/crear")
def propiedades_crear(tipo: str = Form(...), numero: str = Form(...), precio: float = Form(...)):
    db = SessionLocal()
    try:
        db.add(Propiedad(tipo=tipo, numero=numero, precio=float(precio), estado="libre"))
        db.commit()
        return RedirectResponse(url="/propiedades", status_code=303)
    finally:
        db.close()

@app.get("/propiedades/{propiedad_id}/editar", response_class=HTMLResponse)
def propiedades_editar(request: Request, propiedad_id: int):
    db = SessionLocal()
    try:
        p = db.query(Propiedad).filter(Propiedad.id == propiedad_id).first()
        if not p:
            raise HTTPException(404, "Propiedad no encontrada")
        return templates.TemplateResponse("propiedad_editar.html", {"request": request, "p": p})
    finally:
        db.close()

@app.post("/propiedades/{propiedad_id}/actualizar")
def propiedades_actualizar(propiedad_id: int, tipo: str = Form(...), numero: str = Form(...), precio: float = Form(...), estado: str = Form(...)):
    db = SessionLocal()
    try:
        p = db.query(Propiedad).filter(Propiedad.id == propiedad_id).first()
        if not p:
            raise HTTPException(404, "Propiedad no encontrada")
        p.tipo = tipo
        p.numero = numero
        p.precio = float(precio)
        p.estado = estado
        db.commit()
        return RedirectResponse(url="/propiedades", status_code=303)
    finally:
        db.close()

@app.post("/propiedades/{propiedad_id}/eliminar")
def propiedades_eliminar(propiedad_id: int):
    db = SessionLocal()
    try:
        p = db.query(Propiedad).filter(Propiedad.id == propiedad_id).first()
        if not p:
            raise HTTPException(404, "Propiedad no encontrada")
        if p.estado == "ocupado":
            raise HTTPException(400, "No se puede eliminar una propiedad ocupada")
        db.delete(p)
        db.commit()
        return RedirectResponse(url="/propiedades", status_code=303)
    finally:
        db.close()


# =========================================================
# INQUILINOS (CRUD)
# =========================================================
@app.get("/inquilinos", response_class=HTMLResponse)
def inquilinos_listar(request: Request):
    db = SessionLocal()
    try:
        inquilinos = db.query(Inquilino).order_by(Inquilino.id.desc()).all()
        return templates.TemplateResponse("inquilinos.html", {"request": request, "inquilinos": inquilinos})
    finally:
        db.close()

@app.post("/inquilinos/crear")
def inquilinos_crear(
    dni: str = Form(...),
    nombre: str = Form(...),
    telefono: str = Form(""),
    whatsapp: str = Form(""),
    estado: str = Form("Activo")
):
    db = SessionLocal()
    try:
        dni = (dni or "").strip()
        nombre = (nombre or "").strip()
        telefono = (telefono or "").strip()
        whatsapp = (whatsapp or "").strip()
        estado = (estado or "Activo").strip()

        if not dni:
            raise HTTPException(400, "El DNI es obligatorio")

        if not nombre:
            raise HTTPException(400, "El nombre es obligatorio")

        # validar repetido antes de insertar
        existe = db.query(Inquilino).filter(Inquilino.dni == dni).first()
        if existe:
            raise HTTPException(400, f"Ya existe un inquilino con DNI {dni}")

        nuevo = Inquilino(
            dni=dni,
            nombre=nombre,
            telefono=telefono,
            whatsapp=whatsapp,
            estado=estado
        )

        db.add(nuevo)
        db.commit()
        return RedirectResponse(url="/inquilinos", status_code=303)

    except IntegrityError:
        db.rollback()
        raise HTTPException(400, f"Ya existe un inquilino con DNI {dni}")
    finally:
        db.close()

@app.get("/inquilinos/{inquilino_id}/editar", response_class=HTMLResponse)
def inquilinos_editar(request: Request, inquilino_id: int):
    db = SessionLocal()
    try:
        i = db.query(Inquilino).filter(Inquilino.id == inquilino_id).first()
        if not i:
            raise HTTPException(404, "Inquilino no encontrado")
        return templates.TemplateResponse("inquilino_editar.html", {"request": request, "i": i})
    finally:
        db.close()

@app.post("/inquilinos/{inquilino_id}/actualizar")
def inquilinos_actualizar(
    inquilino_id: int,
    dni: str = Form(...),
    nombre: str = Form(...),
    telefono: str = Form(""),
    whatsapp: str = Form(""),
    estado: str = Form("Activo")
):
    db = SessionLocal()
    try:
        i = db.query(Inquilino).filter(Inquilino.id == inquilino_id).first()
        if not i:
            raise HTTPException(404, "Inquilino no encontrado")

        dni = (dni or "").strip()
        nombre = (nombre or "").strip()
        telefono = (telefono or "").strip()
        whatsapp = (whatsapp or "").strip()
        estado = (estado or "Activo").strip()

        if not dni:
            raise HTTPException(400, "El DNI es obligatorio")

        if not nombre:
            raise HTTPException(400, "El nombre es obligatorio")

        existe = db.query(Inquilino).filter(
            Inquilino.dni == dni,
            Inquilino.id != inquilino_id
        ).first()
        if existe:
            raise HTTPException(400, f"Ya existe otro inquilino con DNI {dni}")

        i.dni = dni
        i.nombre = nombre
        i.telefono = telefono
        i.whatsapp = whatsapp
        i.estado = estado

        db.commit()
        return RedirectResponse(url="/inquilinos", status_code=303)

    except IntegrityError:
        db.rollback()
        raise HTTPException(400, f"Ya existe otro inquilino con DNI {dni}")
    finally:
        db.close()

@app.post("/inquilinos/{inquilino_id}/eliminar")
def inquilinos_eliminar(inquilino_id: int):
    db = SessionLocal()
    try:
        existe = db.query(Contrato).filter(Contrato.inquilino_id == inquilino_id, Contrato.estado == "Activo").first()
        if existe:
            raise HTTPException(400, "No se puede eliminar: el inquilino tiene contrato activo")
        i = db.query(Inquilino).filter(Inquilino.id == inquilino_id).first()
        if not i:
            raise HTTPException(404, "Inquilino no encontrado")
        db.delete(i)
        db.commit()
        return RedirectResponse(url="/inquilinos", status_code=303)
    finally:
        db.close()


# =========================================================
# CONTRATOS (CRUD) + mensual/diario
# =========================================================
@app.get("/contratos", response_class=HTMLResponse)
def contratos_listar(request: Request):
    db = SessionLocal()
    try:
        contratos = db.query(Contrato).options(
            joinedload(Contrato.propiedad),
            joinedload(Contrato.inquilino),
        ).order_by(Contrato.id.desc()).all()

        propiedades = db.query(Propiedad).order_by(Propiedad.id.desc()).all()
        inquilinos = db.query(Inquilino).order_by(Inquilino.id.desc()).all()

        return templates.TemplateResponse("contratos.html", {
            "request": request,
            "contratos": contratos,
            "propiedades": propiedades,
            "inquilinos": inquilinos
        })
    finally:
        db.close()

@app.post("/contratos/crear")
def contratos_crear(
    inquilino_id: int = Form(...),
    propiedad_id: int = Form(...),
    fecha_inicio: date = Form(...),
    fecha_fin: date = Form(...),
    monto_mensual: float = Form(...),
    tipo_alquiler: str = Form("mensual"),
):
    db = SessionLocal()
    try:
        prop = db.query(Propiedad).filter(Propiedad.id == propiedad_id).first()
        if not prop:
            raise HTTPException(404, "Propiedad no encontrada")
        if prop.estado != "libre":
            raise HTTPException(400, "La propiedad no está libre")

        contrato = Contrato(
            inquilino_id=inquilino_id,
            propiedad_id=propiedad_id,
            fecha_inicio=fecha_inicio,
            fecha_fin=fecha_fin,
            monto_mensual=float(monto_mensual),
            estado="Activo",
            tipo_alquiler=tipo_alquiler
        )
        db.add(contrato)
        db.flush()

        prop.estado = "ocupado"

        # CARGO INICIAL:
        if tipo_alquiler == "mensual":
            # prorrata hasta el 25
            _, monto_prorrata, venc = calcular_prorrata_mensual(float(monto_mensual), fecha_inicio)
            periodo = f"{venc.year:04d}-{venc.month:02d}"
            if monto_prorrata > 0:
                db.add(Cargo(
                    contrato_id=contrato.id,
                    concepto="ALQUILER_PRORRATA",
                    periodo=periodo,
                    monto=float(monto_prorrata),
                    vencimiento=venc,
                    estado="Pendiente",
                    pagado_acumulado=0.0
                ))
        else:
            # diario: total por días hasta fecha_fin
            dias, total = calcular_alquiler_diario(float(monto_mensual), fecha_inicio, fecha_fin)
            periodo = f"{fecha_inicio.isoformat()}_{fecha_fin.isoformat()}"
            db.add(Cargo(
                contrato_id=contrato.id,
                concepto=f"ALQUILER_DIARIO ({dias} días)",
                periodo=periodo,
                monto=float(total),
                vencimiento=fecha_fin,
                estado="Pendiente",
                pagado_acumulado=0.0
            ))

        db.commit()
        return RedirectResponse(url="/contratos", status_code=303)
    finally:
        db.close()

@app.get("/contratos/{contrato_id}/editar", response_class=HTMLResponse)
def contratos_editar(request: Request, contrato_id: int):
    db = SessionLocal()
    try:
        c = db.query(Contrato).options(
            joinedload(Contrato.propiedad),
            joinedload(Contrato.inquilino),
        ).filter(Contrato.id == contrato_id).first()
        if not c:
            raise HTTPException(404, "Contrato no encontrado")

        propiedades = db.query(Propiedad).order_by(Propiedad.id.desc()).all()
        inquilinos = db.query(Inquilino).order_by(Inquilino.id.desc()).all()

        return templates.TemplateResponse("contrato_editar.html", {
            "request": request,
            "c": c,
            "propiedades": propiedades,
            "inquilinos": inquilinos
        })
    finally:
        db.close()

@app.post("/contratos/{contrato_id}/actualizar")
def contratos_actualizar(
    contrato_id: int,
    inquilino_id: int = Form(...),
    propiedad_id: int = Form(...),
    fecha_inicio: date = Form(...),
    fecha_fin: date = Form(...),
    monto_mensual: float = Form(...),
    estado: str = Form("Activo"),
    tipo_alquiler: str = Form("mensual"),
):
    db = SessionLocal()
    try:
        c = db.query(Contrato).filter(Contrato.id == contrato_id).first()
        if not c:
            raise HTTPException(404, "Contrato no encontrado")

        # Cambio de propiedad (liberar/ocupar)
        if c.propiedad_id != propiedad_id:
            prop_nueva = db.query(Propiedad).filter(Propiedad.id == propiedad_id).first()
            if not prop_nueva:
                raise HTTPException(404, "Propiedad nueva no encontrada")
            if prop_nueva.estado != "libre":
                raise HTTPException(400, "La propiedad nueva no está libre")

            prop_vieja = db.query(Propiedad).filter(Propiedad.id == c.propiedad_id).first()
            if prop_vieja:
                prop_vieja.estado = "libre"
            prop_nueva.estado = "ocupado"
            c.propiedad_id = propiedad_id

        c.inquilino_id = inquilino_id
        c.fecha_inicio = fecha_inicio
        c.fecha_fin = fecha_fin
        c.monto_mensual = float(monto_mensual)
        c.estado = estado
        c.tipo_alquiler = tipo_alquiler

        db.commit()
        return RedirectResponse(url="/contratos", status_code=303)
    finally:
        db.close()

@app.post("/contratos/{contrato_id}/eliminar")
def contratos_eliminar(contrato_id: int):
    db = SessionLocal()
    try:
        c = db.query(Contrato).filter(Contrato.id == contrato_id).first()
        if not c:
            raise HTTPException(404, "Contrato no encontrado")

        prop = db.query(Propiedad).filter(Propiedad.id == c.propiedad_id).first()
        if prop:
            prop.estado = "libre"

        db.query(Pago).filter(Pago.contrato_id == contrato_id).delete()
        db.query(Cargo).filter(Cargo.contrato_id == contrato_id).delete()
        db.query(Lectura).filter(Lectura.contrato_id == contrato_id).delete()

        db.delete(c)
        db.commit()
        return RedirectResponse(url="/contratos", status_code=303)
    finally:
        db.close()


# =========================================================
# CARGOS (CRUD)
# =========================================================
@app.get("/cargos", response_class=HTMLResponse)
def cargos_listar(request: Request):
    db = SessionLocal()
    try:
        cargos = db.query(Cargo).options(
            joinedload(Cargo.contrato).joinedload(Contrato.inquilino),
            joinedload(Cargo.contrato).joinedload(Contrato.propiedad),
        ).order_by(Cargo.vencimiento.desc()).all()

        contratos = db.query(Contrato).options(
            joinedload(Contrato.inquilino),
            joinedload(Contrato.propiedad),
        ).filter(Contrato.estado == "Activo").order_by(Contrato.id.desc()).all()

        return templates.TemplateResponse("cargos.html", {"request": request, "cargos": cargos, "contratos": contratos})
    finally:
        db.close()

@app.post("/cargos/crear")
def cargos_crear(contrato_id: int = Form(...), concepto: str = Form(...), periodo: str = Form(...), monto: float = Form(...), vencimiento: date = Form(...)):
    db = SessionLocal()
    try:
        db.add(Cargo(
            contrato_id=contrato_id,
            concepto=concepto,
            periodo=periodo,
            monto=round(float(monto), 2),
            vencimiento=vencimiento,
            estado="Pendiente",
            pagado_acumulado=0.0
        ))
        db.commit()
        return RedirectResponse(url="/cargos", status_code=303)
    finally:
        db.close()

@app.get("/cargos/{cargo_id}/editar", response_class=HTMLResponse)
def cargos_editar(request: Request, cargo_id: int):
    db = SessionLocal()
    try:
        cargo = db.query(Cargo).filter(Cargo.id == cargo_id).first()
        if not cargo:
            raise HTTPException(404, "Cargo no encontrado")

        contratos = db.query(Contrato).options(
            joinedload(Contrato.inquilino),
            joinedload(Contrato.propiedad),
        ).filter(Contrato.estado == "Activo").order_by(Contrato.id.desc()).all()

        return templates.TemplateResponse("cargo_editar.html", {"request": request, "cargo": cargo, "contratos": contratos})
    finally:
        db.close()

@app.post("/cargos/{cargo_id}/actualizar")
def cargos_actualizar(cargo_id: int, contrato_id: int = Form(...), concepto: str = Form(...), periodo: str = Form(...), monto: float = Form(...), vencimiento: date = Form(...)):
    db = SessionLocal()
    try:
        cargo = db.query(Cargo).filter(Cargo.id == cargo_id).first()
        if not cargo:
            raise HTTPException(404, "Cargo no encontrado")

        cargo.contrato_id = contrato_id
        cargo.concepto = concepto
        cargo.periodo = periodo
        cargo.monto = round(float(monto), 2)
        cargo.vencimiento = vencimiento

        recalcular_cargo(db, cargo_id)

        db.commit()
        return RedirectResponse(url="/cargos", status_code=303)
    finally:
        db.close()

@app.post("/cargos/{cargo_id}/eliminar")
def cargos_eliminar(cargo_id: int):
    db = SessionLocal()
    try:
        existe_pago = db.query(Pago).filter(Pago.cargo_id == cargo_id).first()
        if existe_pago:
            raise HTTPException(400, "No se puede eliminar: el cargo tiene pagos registrados")

        cargo = db.query(Cargo).filter(Cargo.id == cargo_id).first()
        if not cargo:
            raise HTTPException(404, "Cargo no encontrado")

        db.delete(cargo)
        db.commit()
        return RedirectResponse(url="/cargos", status_code=303)
    finally:
        db.close()

@app.post("/generar_cargos_mensuales")
def generar_cargos_manual():
    db = SessionLocal()
    try:
        hoy = date.today()
        generar_cargos_mensuales(db, hoy)
        db.commit()
        return RedirectResponse(url="/cargos", status_code=303)
    finally:
        db.close()


# =========================================================
# PAGOS (POR CONTRATO + PERIODO)
# =========================================================

def obtener_periodos_pendientes(db):
    """
    Devuelve lista agrupada por contrato + periodo con total pendiente.
    """
    rows = (
        db.query(Cargo)
        .options(
            joinedload(Cargo.contrato).joinedload(Contrato.inquilino),
            joinedload(Cargo.contrato).joinedload(Contrato.propiedad),
        )
        .filter(Cargo.estado.in_(["Pendiente", "Parcial"]))
        .order_by(Cargo.contrato_id.asc(), Cargo.periodo.asc(), Cargo.vencimiento.asc(), Cargo.id.asc())
        .all()
    )

    agrupado = {}
    for c in rows:
        saldo = float(c.monto or 0) - float(c.pagado_acumulado or 0)
        if saldo <= 0:
            continue

        key = (c.contrato_id, c.periodo)
        if key not in agrupado:
            agrupado[key] = {
                "contrato_id": c.contrato_id,
                "periodo": c.periodo,
                "inquilino": c.contrato.inquilino.nombre if c.contrato and c.contrato.inquilino else "-",
                "dni": c.contrato.inquilino.dni if c.contrato and c.contrato.inquilino else "-",
                "propiedad": f"{c.contrato.propiedad.tipo} {c.contrato.propiedad.numero}" if c.contrato and c.contrato.propiedad else "-",
                "total_pendiente": 0.0,
                "cargos": []
            }

        agrupado[key]["total_pendiente"] += saldo
        agrupado[key]["cargos"].append({
            "id": c.id,
            "concepto": c.concepto,
            "saldo": round(saldo, 2),
            "vencimiento": c.vencimiento
        })

    return list(agrupado.values())


@app.get("/pagos", response_class=HTMLResponse)
def pagos_listar(request: Request):
    db = SessionLocal()
    try:
        pagos = (
            db.query(Pago)
            .options(
                joinedload(Pago.contrato).joinedload(Contrato.inquilino),
                joinedload(Pago.contrato).joinedload(Contrato.propiedad),
                joinedload(Pago.cargo),
            )
            .order_by(Pago.fecha_pago.desc(), Pago.id.desc())
            .all()
        )

        periodos_pendientes = obtener_periodos_pendientes(db)

        return templates.TemplateResponse("pagos.html", {
            "request": request,
            "pagos": pagos,
            "periodos_pendientes": periodos_pendientes
        })
    finally:
        db.close()


@app.post("/pagos/crear")
def pagos_crear(
    contrato_id: int = Form(...),
    periodo: str = Form(...),
    monto_total: float = Form(...),
    fecha_pago: str = Form(...),
    metodo: str = Form(...),
):
    db = SessionLocal()
    try:
        monto_disponible = round(float(monto_total), 2)
        if monto_disponible <= 0:
            raise HTTPException(400, "El monto debe ser mayor a cero")

        cargos = (
            db.query(Cargo)
            .filter(
                Cargo.contrato_id == contrato_id,
                Cargo.periodo == periodo,
                Cargo.estado.in_(["Pendiente", "Parcial"])
            )
            .order_by(Cargo.vencimiento.asc(), Cargo.id.asc())
            .all()
        )

        if not cargos:
            raise HTTPException(404, "No hay cargos pendientes para ese contrato y periodo")

        total_pendiente = 0.0
        for c in cargos:
            saldo = float(c.monto or 0) - float(c.pagado_acumulado or 0)
            if saldo > 0:
                total_pendiente += saldo

        total_pendiente = round(total_pendiente, 2)

        if monto_disponible > total_pendiente:
            raise HTTPException(400, f"El monto excede el total pendiente del periodo. Máximo: S/ {total_pendiente:.2f}")

        fecha = date.fromisoformat(fecha_pago)

        # Repartir pago por orden de vencimiento
        for c in cargos:
            saldo = round(float(c.monto or 0) - float(c.pagado_acumulado or 0), 2)
            if saldo <= 0:
                continue

            aplicar = min(monto_disponible, saldo)
            if aplicar <= 0:
                break

            p = Pago(
                contrato_id=contrato_id,
                cargo_id=c.id,
                monto=round(aplicar, 2),
                fecha_pago=fecha,
                metodo=metodo
            )
            db.add(p)
            db.flush()

            recalcular_cargo(db, c.id)
            monto_disponible = round(monto_disponible - aplicar, 2)

            if monto_disponible <= 0:
                break

        db.commit()
        return RedirectResponse(url="/pagos", status_code=303)
    finally:
        db.close()


@app.get("/pagos/{pago_id}/editar", response_class=HTMLResponse)
def pagos_editar(request: Request, pago_id: int):
    db = SessionLocal()
    try:
        pago = db.query(Pago).options(
            joinedload(Pago.cargo),
            joinedload(Pago.contrato).joinedload(Contrato.inquilino),
            joinedload(Pago.contrato).joinedload(Contrato.propiedad),
        ).filter(Pago.id == pago_id).first()

        if not pago:
            raise HTTPException(404, "Pago no encontrado")

        return templates.TemplateResponse("pago_editar.html", {
            "request": request,
            "pago": pago
        })
    finally:
        db.close()


@app.post("/pagos/{pago_id}/actualizar")
def pagos_actualizar(
    pago_id: int,
    monto: float = Form(...),
    fecha_pago: str = Form(...),
    metodo: str = Form(...),
):
    db = SessionLocal()
    try:
        pago = db.query(Pago).filter(Pago.id == pago_id).first()
        if not pago:
            raise HTTPException(404, "Pago no encontrado")

        cargo = db.query(Cargo).filter(Cargo.id == pago.cargo_id).first()
        if not cargo:
            raise HTTPException(404, "Cargo no encontrado")

        # saldo permitido = saldo actual + pago actual
        saldo_actual = round(float(cargo.monto or 0) - float(cargo.pagado_acumulado or 0), 2)
        maximo = round(saldo_actual + float(pago.monto or 0), 2)

        if float(monto) <= 0:
            raise HTTPException(400, "El monto debe ser mayor a cero")

        if float(monto) > maximo:
            raise HTTPException(400, f"El monto no puede ser mayor a S/ {maximo:.2f}")

        pago.monto = round(float(monto), 2)
        pago.fecha_pago = date.fromisoformat(fecha_pago)
        pago.metodo = metodo

        db.flush()
        recalcular_cargo(db, cargo.id)

        db.commit()
        return RedirectResponse(url="/pagos", status_code=303)
    finally:
        db.close()


@app.post("/pagos/{pago_id}/eliminar")
def pagos_eliminar(pago_id: int):
    db = SessionLocal()
    try:
        pago = db.query(Pago).filter(Pago.id == pago_id).first()
        if not pago:
            raise HTTPException(404, "Pago no encontrado")

        cargo_id = pago.cargo_id

        db.delete(pago)
        db.flush()

        if cargo_id:
            recalcular_cargo(db, cargo_id)

        db.commit()
        return RedirectResponse(url="/pagos", status_code=303)
    finally:
        db.close()


@app.get("/pagos/{pago_id}/recibo.pdf")
def descargar_recibo_pdf(pago_id: int):
    db = SessionLocal()
    try:
        pago = db.query(Pago).options(
            joinedload(Pago.contrato).joinedload(Contrato.inquilino),
            joinedload(Pago.contrato).joinedload(Contrato.propiedad),
            joinedload(Pago.cargo),
        ).filter(Pago.id == pago_id).first()

        if not pago:
            raise HTTPException(404, "Pago no encontrado")

        pdf_bytes = generar_recibo_pdf_pro(pago)
        filename = f"recibo_{serie_recibo(pago.id, pago.fecha_pago)}.pdf"

        return StreamingResponse(
            io.BytesIO(pdf_bytes),
            media_type="application/pdf",
            headers={"Content-Disposition": f'inline; filename="{filename}"'}
        )
    finally:
        db.close()

# =========================================================
# LECTURAS (CRUD) -> genera cargo del servicio
# =========================================================
@app.get("/lecturas", response_class=HTMLResponse)
def lecturas_listar(request: Request):
    db = SessionLocal()
    try:
        hoy = date.today()
        periodo = f"{hoy.year:04d}-{hoy.month:02d}"

        contratos = db.query(Contrato).options(
            joinedload(Contrato.inquilino),
            joinedload(Contrato.propiedad),
        ).filter(Contrato.estado == "Activo").all()

        lecturas = db.query(Lectura).options(
            joinedload(Lectura.contrato).joinedload(Contrato.inquilino),
            joinedload(Lectura.contrato).joinedload(Contrato.propiedad),
        ).order_by(Lectura.fecha_registro.desc()).all()

        return templates.TemplateResponse("lecturas.html", {
            "request": request,
            "contratos": contratos,
            "lecturas": lecturas,
            "periodo": periodo
        })
    finally:
        db.close()

@app.post("/lecturas/crear")
def lecturas_crear(
    contrato_id: int = Form(...),
    servicio: str = Form(...),
    periodo: str = Form(...),
    lectura_anterior: float = Form(...),
    lectura_actual: float = Form(...),
    tarifa: float = Form(...),
):
    db = SessionLocal()
    try:
        if float(lectura_actual) < float(lectura_anterior):
            raise HTTPException(400, "La lectura actual no puede ser menor que la anterior")

        consumo = round(float(lectura_actual) - float(lectura_anterior), 2)
        monto = round(consumo * float(tarifa), 2)

        existe = db.query(Lectura).filter(
            Lectura.contrato_id == contrato_id,
            Lectura.servicio == servicio,
            Lectura.periodo == periodo
        ).first()
        if existe:
            raise HTTPException(400, f"Ya existe lectura {servicio} para ese contrato en {periodo}")

        lectura = Lectura(
            contrato_id=contrato_id,
            servicio=servicio,
            periodo=periodo,
            lectura_anterior=float(lectura_anterior),
            lectura_actual=float(lectura_actual),
            consumo=consumo,
            tarifa=float(tarifa),
            monto=monto,
            fecha_registro=date.today()
        )
        db.add(lectura)

        y, m = periodo.split("-")
        venc = date(int(y), int(m), DIA_CORTE)

        db.add(Cargo(
            contrato_id=contrato_id,
            concepto=servicio,
            periodo=periodo,
            monto=monto,
            vencimiento=venc,
            estado="Pendiente",
            pagado_acumulado=0.0
        ))

        db.commit()
        return RedirectResponse(url="/lecturas", status_code=303)
    finally:
        db.close()

@app.post("/lecturas/{lectura_id}/eliminar")
def lecturas_eliminar(lectura_id: int):
    db = SessionLocal()
    try:
        l = db.query(Lectura).filter(Lectura.id == lectura_id).first()
        if not l:
            raise HTTPException(404, "Lectura no encontrada")
        db.delete(l)
        db.commit()
        return RedirectResponse(url="/lecturas", status_code=303)
    finally:
        db.close()


# (Opcional) para que no falle tu base.html si existe /logout
@app.get("/logout")
def logout():
    return RedirectResponse(url="/", status_code=303)

from urllib.parse import quote

# =========================================================
# ESTADO DE CUENTA PDF
# =========================================================
def estado_cuenta_pdf(inq: Inquilino, cargos: list[Cargo]) -> bytes:
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    w, h = A4
    M = 16 * mm
    y = h - M

    c.setFont("Helvetica-Bold", 16)
    c.drawString(M, y, "ESTADO DE CUENTA - INQUILINO")
    y -= 8 * mm

    c.setFont("Helvetica", 10)
    c.drawString(M, y, f"Fecha emisión: {date.today().isoformat()}")
    y -= 6 * mm

    c.setFont("Helvetica-Bold", 11)
    c.drawString(M, y, f"Inquilino: {inq.nombre}")
    y -= 5 * mm
    c.setFont("Helvetica", 10)
    c.drawString(M, y, f"DNI: {inq.dni}    WhatsApp: {inq.whatsapp or '-'}")
    y -= 8 * mm

    # Encabezado tabla
    c.setFillColor(colors.HexColor("#F3F4F6"))
    c.rect(M, y - 6*mm, w - 2*M, 8*mm, fill=1, stroke=0)
    c.setFillColor(colors.black)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(M + 2*mm, y - 3*mm, "Venc.")
    c.drawString(M + 25*mm, y - 3*mm, "Propiedad")
    c.drawString(M + 60*mm, y - 3*mm, "Concepto")
    c.drawString(M + 100*mm, y - 3*mm, "Periodo")
    c.drawRightString(w - M - 55*mm, y - 3*mm, "Monto")
    c.drawRightString(w - M - 28*mm, y - 3*mm, "Pagado")
    c.drawRightString(w - M - 2*mm, y - 3*mm, "Saldo")
    y -= 10*mm

    total_monto = 0.0
    total_pagado = 0.0
    total_saldo = 0.0

    c.setFont("Helvetica", 9)

    for item in cargos:
        saldo = float(item.monto or 0.0) - float(item.pagado_acumulado or 0.0)
        if saldo < 0:
            saldo = 0.0

        if y < 25 * mm:
            c.showPage()
            y = h - M

        prop = "-"
        if item.contrato and item.contrato.propiedad:
            prop = item.contrato.propiedad.numero

        c.drawString(M + 2*mm, y, str(item.vencimiento))
        c.drawString(M + 25*mm, y, str(prop))
        c.drawString(M + 60*mm, y, str(item.concepto))
        c.drawString(M + 100*mm, y, str(item.periodo))
        c.drawRightString(w - M - 55*mm, y, f"S/ {float(item.monto or 0):.2f}")
        c.drawRightString(w - M - 28*mm, y, f"S/ {float(item.pagado_acumulado or 0):.2f}")
        c.drawRightString(w - M - 2*mm, y, f"S/ {saldo:.2f}")
        y -= 6*mm

        total_monto += float(item.monto or 0.0)
        total_pagado += float(item.pagado_acumulado or 0.0)
        total_saldo += float(saldo)

    y -= 4*mm
    c.setStrokeColor(colors.HexColor("#111827"))
    c.line(M, y, w - M, y)
    y -= 8*mm

    c.setFont("Helvetica-Bold", 11)
    c.drawRightString(w - M, y, f"TOTAL MONTO: S/ {total_monto:.2f}")
    y -= 6*mm
    c.drawRightString(w - M, y, f"TOTAL PAGADO: S/ {total_pagado:.2f}")
    y -= 6*mm
    c.drawRightString(w - M, y, f"TOTAL SALDO: S/ {total_saldo:.2f}")

    c.setFont("Helvetica-Oblique", 8)
    c.setFillColor(colors.HexColor("#6B7280"))
    c.drawString(M, 12*mm, "Documento generado automáticamente por el sistema.")
    c.setFillColor(colors.black)

    c.showPage()
    c.save()
    buffer.seek(0)
    return buffer.read()


@app.get("/inquilinos/{inquilino_id}/estado_cuenta.pdf")
def descargar_estado_cuenta(inquilino_id: int):
    db = SessionLocal()
    try:
        inq = db.query(Inquilino).filter(Inquilino.id == inquilino_id).first()
        if not inq:
            raise HTTPException(404, "Inquilino no encontrado")

        cargos = (
            db.query(Cargo)
            .join(Contrato, Cargo.contrato_id == Contrato.id)
            .options(
                joinedload(Cargo.contrato).joinedload(Contrato.propiedad),
                joinedload(Cargo.contrato).joinedload(Contrato.inquilino),
            )
            .filter(
                Contrato.inquilino_id == inquilino_id,
                Cargo.estado.in_(["Pendiente", "Parcial"])
            )
            .order_by(Cargo.vencimiento.asc())
            .all()
        )

        pdf_bytes = estado_cuenta_pdf(inq, cargos)
        filename = f"estado_cuenta_{inq.dni}_{date.today().isoformat()}.pdf"

        return StreamingResponse(
            io.BytesIO(pdf_bytes),
            media_type="application/pdf",
            headers={"Content-Disposition": f'inline; filename="{filename}"'}
        )
    finally:
        db.close()


# =========================================================
# WHATSAPP RECORDATORIO
# =========================================================
def normalizar_numero_pe(numero: str) -> str:
    n = "".join([c for c in (numero or "") if c.isdigit()])
    if not n:
        return ""
    if len(n) == 9:
        return "51" + n
    return n


@app.get("/whatsapp/recordatorio/{inquilino_id}")
def whatsapp_recordatorio(inquilino_id: int):
    db = SessionLocal()
    try:
        inq = db.query(Inquilino).filter(Inquilino.id == inquilino_id).first()
        if not inq:
            raise HTTPException(404, "Inquilino no encontrado")

        deuda = db.query(
            func.coalesce(
                func.sum(Cargo.monto - func.coalesce(Cargo.pagado_acumulado, 0.0)),
                0.0
            )
        ).join(Contrato, Cargo.contrato_id == Contrato.id).filter(
            Contrato.inquilino_id == inquilino_id,
            Cargo.estado.in_(["Pendiente", "Parcial"])
        ).scalar() or 0.0

        numero = normalizar_numero_pe(inq.whatsapp or inq.telefono or "")
        if not numero:
            raise HTTPException(400, "El inquilino no tiene WhatsApp o teléfono registrado")

        mensaje = (
            f"Hola {inq.nombre}, te saluda CREDIMAS. "
            f"Tu saldo pendiente es S/ {float(deuda):.2f}. "
            f"Por favor regularizar tu pago. Gracias."
        )

        url = f"https://wa.me/{numero}?text={quote(mensaje)}"
        return RedirectResponse(url=url, status_code=302)
    finally:
        db.close()

@app.get("/whatsapp/recibo/{pago_id}")
def whatsapp_recibo(pago_id: int, request: Request):
    db = SessionLocal()
    try:
        pago = db.query(Pago).options(
            joinedload(Pago.contrato).joinedload(Contrato.inquilino),
            joinedload(Pago.contrato).joinedload(Contrato.propiedad),
            joinedload(Pago.cargo),
        ).filter(Pago.id == pago_id).first()

        if not pago:
            raise HTTPException(404, "Pago no encontrado")

        inq = pago.contrato.inquilino
        numero = normalizar_numero_pe(inq.whatsapp or inq.telefono or "")
        if not numero:
            raise HTTPException(400, "El inquilino no tiene WhatsApp o teléfono registrado")

        base_url = str(request.base_url).rstrip("/")
        link_recibo = f"{base_url}/pagos/{pago.id}/recibo.pdf"

        mensaje = (
            f"Hola {inq.nombre}, se registró su pago correctamente.\n"
            f"Monto pagado: S/ {float(pago.monto):.2f}\n"
            f"Fecha: {pago.fecha_pago}\n"
            f"Método: {pago.metodo}\n"
            f"Recibo: {link_recibo}\n"
            f"Gracias por su pago.\n"
            f"CREDIMAS"
        )

        url = f"https://wa.me/{numero}?text={quote(mensaje)}"
        return RedirectResponse(url=url, status_code=302)

    finally:
        db.close()


