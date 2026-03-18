import io
from datetime import date, timedelta, datetime
from urllib.parse import quote
from functools import wraps

from fastapi import FastAPI, Request, Form, HTTPException
from fastapi.responses import HTMLResponse, RedirectResponse, StreamingResponse, JSONResponse

from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles

from starlette.middleware.sessions import SessionMiddleware

from collections import defaultdict
from openpyxl import Workbook
from fastapi.responses import StreamingResponse

from sqlalchemy.orm import joinedload
from sqlalchemy import func, case
from sqlalchemy.exc import IntegrityError

from passlib.context import CryptContext

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib import colors
from urllib.parse import quote

from database import engine, Base, SessionLocal
from models import Propiedad, Inquilino, Contrato, Cargo, Pago, Lectura, BitacoraCobranza, ConfiguracionEmpresa, Base, Gasto
from models_auth import User, AuditLog

app = FastAPI()

app.add_middleware(
    SessionMiddleware,
    secret_key="credimas-clave-segura-2026",
    session_cookie="credimas_session",
    max_age=60 * 60 * 8,  # 8 horas
    same_site="lax",
    https_only=False
)

app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="templates")

pwd_context = CryptContext(schemes=["pbkdf2_sha256"], deprecated="auto")

Base.metadata.create_all(bind=engine)

DIA_CORTE = 25
TARIFA_DIARIA_NOCHE = 50.0

# ----------------------------
# AUTH / SESION
# ----------------------------
def hash_password(password: str) -> str:
    return pwd_context.hash(password)

def verify_password(password: str, password_hash: str) -> bool:
    return pwd_context.verify(password, password_hash)

def get_current_user(request: Request):
    user_id = request.session.get("user_id")
    if not user_id:
        return None

    db = SessionLocal()
    try:
        user = db.query(User).filter(User.id == user_id).first()
        if not user:
            return None
        if user.estado != "Activo":
            return None
        return user
    finally:
        db.close()

def login_required(route_function):
    @wraps(route_function)
    def wrapper(*args, **kwargs):
        request = kwargs.get("request")
        if request is None:
            for arg in args:
                if isinstance(arg, Request):
                    request = arg
                    break

        if request is None:
            raise HTTPException(status_code=500, detail="Request no encontrado")

        user_id = request.session.get("user_id")
        if not user_id:
            return RedirectResponse(url="/login", status_code=303)

        db = SessionLocal()
        try:
            user = db.query(User).filter(User.id == user_id).first()
            if not user or user.estado != "Activo":
                request.session.clear()
                return RedirectResponse(url="/login", status_code=303)
        finally:
            db.close()

        return route_function(*args, **kwargs)
    return wrapper
def role_required(*roles_permitidos):
    def decorator(route_function):
        @wraps(route_function)
        def wrapper(*args, **kwargs):
            request = kwargs.get("request")
            if request is None:
                for arg in args:
                    if isinstance(arg, Request):
                        request = arg
                        break

            if request is None:
                raise HTTPException(status_code=500, detail="Request no encontrado")

            user_id = request.session.get("user_id")
            user_rol = request.session.get("user_rol")

            if not user_id:
                return RedirectResponse(url="/login", status_code=303)

            if user_rol not in roles_permitidos:
                return RedirectResponse(url="/dashboard", status_code=303)

            return route_function(*args, **kwargs)
        return wrapper
    return decorator
def get_client_ip(request: Request) -> str:
    forwarded = request.headers.get("x-forwarded-for")
    if forwarded:
        return forwarded.split(",")[0].strip()

    if request.client:
        return request.client.host

    return ""


def registrar_auditoria(
    db,
    request: Request,
    action: str,
    module: str,
    detail: str = "",
    user_id: int = None
):
    if user_id is None:
        user_id = request.session.get("user_id")

    log = AuditLog(
        user_id=user_id,
        action=action,
        module=module,
        detail=detail,
        ip=get_client_ip(request)
    )
    db.add(log)

def crear_admin_inicial():
    db = SessionLocal()
    try:
        admin = db.query(User).filter(User.username == "admin").first()
        if not admin:
            nuevo = User(
                nombre="Administrador General",
                username="admin",
                password_hash=hash_password("admin123"),
                rol="Administrador",
                estado="Activo"
            )
            db.add(nuevo)
            db.commit()
    finally:
        db.close()

crear_admin_inicial()


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


def calcular_total_noche(fecha_inicio: date, fecha_fin: date, tarifa: float):
    noches = (fecha_fin - fecha_inicio).days

    if noches <= 0:
        raise HTTPException(
            status_code=400,
            detail="La fecha fin debe ser mayor que la fecha inicio"
        )

    total = round(float(noches) * float(tarifa), 2)
    return noches, total


def calcular_total_diario(fecha_inicio: date, fecha_fin: date, tarifa: float):
    dias = (fecha_fin - fecha_inicio).days + 1

    if dias <= 0:
        raise HTTPException(
            status_code=400,
            detail="La fecha fin debe ser igual o mayor que la fecha inicio"
        )

    total = round(float(dias) * float(tarifa), 2)
    return dias, total


def calcular_total_diario_noche(
    fecha_inicio: date,
    fecha_fin: date,
    tarifa: float,
    tipo_alquiler: str = "noche"
):
    tipo = (tipo_alquiler or "noche").strip().lower()

    if tipo == "noche":
        return calcular_total_noche(fecha_inicio, fecha_fin, tarifa)

    if tipo == "diario":
        return calcular_total_diario(fecha_inicio, fecha_fin, tarifa)

    raise HTTPException(
        status_code=400,
        detail="Tipo de alquiler inválido para cálculo diario/noche"
    )


def validar_rango_fechas(fecha_inicio: date, fecha_fin: date):
    if fecha_fin < fecha_inicio:
        raise HTTPException(400, "La fecha fin no puede ser menor que la fecha inicio")


def actualizar_estado_propiedad_por_contrato(db, contrato: Contrato):
    prop = db.query(Propiedad).filter(Propiedad.id == contrato.propiedad_id).first()
    if not prop:
        return

    hay_activo = db.query(Contrato).filter(
        Contrato.propiedad_id == contrato.propiedad_id,
        Contrato.estado == "Activo",
        Contrato.id != contrato.id
    ).first()

    if contrato.estado == "Activo" or hay_activo:
        prop.estado = "ocupado"
    else:
        prop.estado = "libre"

# ----------------------------
# Cargos automáticos (mensual día 25)
# ----------------------------
def generar_cargos_mensuales(db, hoy: date):
    """Crea ALQUILER_MENSUAL solo para contratos mensuales activos."""
    if hoy.day < DIA_CORTE:
        return

    venc = date(hoy.year, hoy.month, DIA_CORTE)
    periodo = f"{venc.year:04d}-{venc.month:02d}"

    contratos = db.query(Contrato).filter(
        Contrato.estado == "Activo",
        (Contrato.tipo_alquiler == None) | (Contrato.tipo_alquiler == "mensual"),
        Contrato.fecha_inicio <= hoy,
        Contrato.fecha_fin >= hoy
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


import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors


def generar_recibo_pdf_pro(pago, config, formato: str = "a4") -> bytes:
    buffer = io.BytesIO()

    formato = (formato or "a4").strip().lower()

    # =========================
    # CONFIGURACIÓN GENERAL
    # =========================
    nombre_negocio = config.nombre_comercial or "CREDIMAS"
    moneda = config.moneda or "S/"
    mensaje_recibo = config.mensaje_recibo or "Gracias por su pago."
    direccion = config.direccion or ""
    telefono = config.telefono or ""

    # =========================
    # TAMAÑO DE PÁGINA
    # =========================
    if formato == "ticket":
        w = 80 * mm
        h = 220 * mm
        pagesize = (w, h)
        M = 5 * mm
    else:
        w, h = A4
        pagesize = A4
        M = 16 * mm

    c = canvas.Canvas(buffer, pagesize=pagesize)

    # =========================
    # DATOS RELACIONADOS
    # =========================
    contrato = pago.contrato
    inq = contrato.inquilino if contrato else None
    prop = contrato.propiedad if contrato else None
    cargo = pago.cargo

    # =========================
    # FORMATO TICKET
    # =========================
    if formato == "ticket":
        y = h - 8 * mm

        # Título
        c.setFont("Helvetica-Bold", 11)
        c.drawCentredString(w / 2, y, nombre_negocio)
        y -= 5 * mm

        if direccion:
            c.setFont("Helvetica", 7)
            c.drawCentredString(w / 2, y, direccion)
            y -= 4 * mm

        if telefono:
            c.drawCentredString(w / 2, y, f"Tel: {telefono}")
            y -= 4 * mm

        y -= 2 * mm

        c.setFont("Helvetica-Bold", 9)
        c.drawCentredString(w / 2, y, "RECIBO DE PAGO")
        y -= 6 * mm

        # Datos
        c.setFont("Helvetica", 8)

        c.drawString(M, y, f"Fecha: {pago.fecha_pago}")
        y -= 5 * mm

        c.drawString(M, y, f"Inquilino:")
        y -= 4 * mm
        c.drawString(M, y, f"{inq.nombre if inq else '-'}")
        y -= 5 * mm

        c.drawString(M, y, f"Unidad: {prop.tipo if prop else ''} {prop.numero if prop else ''}")
        y -= 5 * mm

        c.drawString(M, y, f"Concepto:")
        y -= 4 * mm
        c.drawString(M, y, f"{cargo.concepto if cargo else '-'}")
        y -= 5 * mm

        c.drawString(M, y, f"Periodo: {cargo.periodo if cargo else '-'}")
        y -= 5 * mm

        c.drawString(M, y, f"Metodo: {pago.metodo or '-'}")
        y -= 6 * mm

        # Línea
        c.line(M, y, w - M, y)
        y -= 6 * mm

        # Total
        c.setFont("Helvetica-Bold", 10)
        c.drawString(M, y, f"TOTAL: {moneda} {float(pago.monto or 0):.2f}")
        y -= 8 * mm

        # Mensaje
        c.setFont("Helvetica", 8)
        c.drawCentredString(w / 2, y, mensaje_recibo)

    # =========================
    # FORMATO A4 (TU DISEÑO PRO)
    # =========================
    else:
        top = h - M
        left = M
        right = w - M

        # HEADER
        c.setFillColor(colors.HexColor("#111827"))
        c.rect(left, top - 34*mm, right - left, 34*mm, fill=1, stroke=0)

        c.setFillColor(colors.white)
        c.setFont("Helvetica-Bold", 17)
        c.drawString(left + 6*mm, top - 11*mm, nombre_negocio)

        c.setFont("Helvetica", 9)
        if direccion:
            c.drawString(left + 6*mm, top - 18*mm, direccion)
        if telefono:
            c.drawString(left + 6*mm, top - 23*mm, f"Tel: {telefono}")

        c.setFont("Helvetica-Bold", 11)
        c.drawRightString(right - 6*mm, top - 14*mm, "RECIBO DE PAGO")

        y = top - 45*mm
        c.setFillColor(colors.black)

        c.setFont("Helvetica", 10)

        c.drawString(left, y, f"Inquilino: {inq.nombre if inq else '-'}")
        y -= 7*mm

        c.drawString(left, y, f"Unidad: {prop.tipo if prop else ''} {prop.numero if prop else ''}")
        y -= 7*mm

        c.drawString(left, y, f"Concepto: {cargo.concepto if cargo else '-'}")
        y -= 7*mm

        c.drawString(left, y, f"Periodo: {cargo.periodo if cargo else '-'}")
        y -= 7*mm

        c.drawString(left, y, f"Fecha: {pago.fecha_pago}")
        y -= 7*mm

        c.drawString(left, y, f"Metodo: {pago.metodo or '-'}")
        y -= 10*mm

        c.setFont("Helvetica-Bold", 13)
        c.drawString(left, y, f"TOTAL PAGADO: {moneda} {float(pago.monto or 0):.2f}")
        y -= 10*mm

        c.setFont("Helvetica", 9)
        c.drawString(left, y, mensaje_recibo)

    c.showPage()
    c.save()

    buffer.seek(0)
    return buffer.read()
# =========================================================
# LOGIN / LOGOUT
# =========================================================
@app.get("/login", response_class=HTMLResponse)
def login_view(request: Request):
    if request.session.get("user_id"):
        return RedirectResponse(url="/dashboard", status_code=303)

    return templates.TemplateResponse(
        "login.html",
        {
            "request": request,
            "error": None
        }
    )

@app.post("/login", response_class=HTMLResponse)
def login_post(
    request: Request,
    username: str = Form(...),
    password: str = Form(...)
):
    db = SessionLocal()
    try:
        user = db.query(User).filter(User.username == username.strip()).first()

        if not user:
            return templates.TemplateResponse(
                "login.html",
                {
                    "request": request,
                    "error": "Usuario o contraseña incorrectos"
                }
            )

        if user.estado != "Activo":
            return templates.TemplateResponse(
                "login.html",
                {
                    "request": request,
                    "error": "Tu usuario está inactivo"
                }
            )

        if not verify_password(password, user.password_hash):
            return templates.TemplateResponse(
                "login.html",
                {
                    "request": request,
                    "error": "Usuario o contraseña incorrectos"
                }
            )

        request.session["user_id"] = user.id
        request.session["user_nombre"] = user.nombre
        request.session["user_username"] = user.username
        request.session["user_rol"] = user.rol

        registrar_auditoria(
            db,
            request,
            action="LOGIN",
            module="AUTH",
            detail=f"Inicio de sesión de usuario {user.username}",
            user_id=user.id
        )
        db.commit()

        if user.debe_cambiar_password == 1:
            return RedirectResponse(url="/cambiar-password", status_code=303)

        return RedirectResponse(url="/dashboard", status_code=303)

    finally:
        db.close()
@app.get("/logout")
def logout(request: Request):
    db = SessionLocal()
    try:
        user_id = request.session.get("user_id")
        username = request.session.get("user_username", "")

        if user_id:
            try:
                registrar_auditoria(
                    db,
                    request,
                    action="LOGOUT",
                    module="AUTH",
                    detail=f"Cierre de sesión de usuario {username}",
                    user_id=user_id
                )
                db.commit()
            except Exception:
                db.rollback()  # 🔥 evita que rompa logout

    finally:
        db.close()

    # 🔥 limpiar sesión SIEMPRE
    request.session.clear()

    return RedirectResponse(url="/login", status_code=303)


# =========================================================
# DASHBOARD
# =========================================================
@app.get("/", response_class=HTMLResponse)
def inicio(request: Request):
    return RedirectResponse(url="/dashboard", status_code=303)
@app.get("/dashboard", response_class=HTMLResponse)
@login_required
def dashboard(
    request: Request,
    ok_pago: int = 0,
    ultimo_pago_id: int = 0
):
    db = SessionLocal()
    try:
        hoy = date.today()
        inicio_mes = date(hoy.year, hoy.month, 1)

        ingresos_hoy = db.query(
            func.coalesce(func.sum(Pago.monto), 0.0)
        ).filter(
            Pago.fecha_pago == hoy
        ).scalar() or 0.0

        ingresos_mes = db.query(
            func.coalesce(func.sum(Pago.monto), 0.0)
        ).filter(
            Pago.fecha_pago >= inicio_mes,
            Pago.fecha_pago <= hoy
        ).scalar() or 0.0

        deuda_total = db.query(
            func.coalesce(
                func.sum(Cargo.monto - func.coalesce(Cargo.pagado_acumulado, 0.0)),
                0.0
            )
        ).filter(
            Cargo.estado.in_(["Pendiente", "Parcial"])
        ).scalar() or 0.0

        pagos_hoy = db.query(
            func.count(Pago.id)
        ).filter(
            Pago.fecha_pago == hoy
        ).scalar() or 0


        periodo_actual = f"{hoy.year:04d}-{hoy.month:02d}"

        generar_cargos_mensuales(db, hoy)
        db.commit()

        total_propiedades = db.query(func.count(Propiedad.id)).scalar() or 0
        ocupadas = db.query(func.count(Propiedad.id)).filter(Propiedad.estado == "ocupado").scalar() or 0
        libres = db.query(func.count(Propiedad.id)).filter(Propiedad.estado == "libre").scalar() or 0

        total_inquilinos = db.query(func.count(Inquilino.id)).scalar() or 0
        contratos_activos = db.query(func.count(Contrato.id)).filter(Contrato.estado == "Activo").scalar() or 0

        pendiente_mes = db.query(
            func.coalesce(
                func.sum(Cargo.monto - func.coalesce(Cargo.pagado_acumulado, 0.0)),
                0.0
            )
        ).filter(
            Cargo.periodo == periodo_actual,
            Cargo.estado.in_(["Pendiente", "Parcial"])
        ).scalar() or 0.0

        inicio_mes = date(hoy.year, hoy.month, 1)
        cobrado_mes = db.query(
            func.coalesce(func.sum(Pago.monto), 0.0)
        ).filter(
            Pago.fecha_pago >= inicio_mes,
            Pago.fecha_pago <= hoy
        ).scalar() or 0.0

        morosos = db.query(func.count(Cargo.id)).filter(
            Cargo.estado.in_(["Pendiente", "Parcial"]),
            Cargo.vencimiento < hoy
        ).scalar() or 0

        cargos_morosos = db.query(Cargo).options(
            joinedload(Cargo.contrato).joinedload(Contrato.inquilino),
            joinedload(Cargo.contrato).joinedload(Contrato.propiedad),
        ).filter(
            Cargo.estado.in_(["Pendiente", "Parcial"]),
            Cargo.vencimiento < hoy
        ).order_by(Cargo.vencimiento.asc()).limit(20).all()

        proximos_vencimientos = db.query(Cargo).options(
            joinedload(Cargo.contrato).joinedload(Contrato.inquilino),
            joinedload(Cargo.contrato).joinedload(Contrato.propiedad),
        ).filter(
            Cargo.estado.in_(["Pendiente", "Parcial"]),
            Cargo.vencimiento >= hoy
        ).order_by(Cargo.vencimiento.asc()).limit(15).all()

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

        pagos_recientes = db.query(Pago).options(
            joinedload(Pago.contrato).joinedload(Contrato.inquilino),
            joinedload(Pago.contrato).joinedload(Contrato.propiedad),
            joinedload(Pago.cargo)
        ).order_by(Pago.fecha_pago.desc(), Pago.id.desc()).limit(10).all()

        propiedades = db.query(Propiedad).order_by(Propiedad.numero.asc()).all()
        propiedades_tablero = []

        for prop in propiedades:
            contrato_activo = db.query(Contrato).options(
                joinedload(Contrato.inquilino)
            ).filter(
                Contrato.propiedad_id == prop.id,
                Contrato.estado == "Activo"
            ).first()

            ultimo_pago = None
            ultimo_pago_id = None
            monto_contrato = 0.0

            moroso = False
            inquilino_nombre = None
            tipo_alquiler = None
            fecha_inicio = None
            fecha_fin = None
            deuda_total = 0.0
            deuda_alquiler = 0.0
            deuda_agua = 0.0
            deuda_luz = 0.0
            deuda_otros = 0.0
            vence_pronto = False
            pago_parcial = False
            estado_financiero = "sin_deuda"
            periodos_pendientes = []
            tiene_ampliacion_pendiente = False
            ultima_lectura_agua = 0.0
            ultima_lectura_luz = 0.0
            fecha_lectura_agua = "-"
            fecha_lectura_luz = "-"

            numero_txt = str(prop.numero or "").strip()
            tipo_txt = str(prop.tipo or "").strip().lower()

            if tipo_txt == "local":
                grupo = "Locales"
            elif tipo_txt == "cochera":
                grupo = "Cocheras"
            elif numero_txt.isdigit() and len(numero_txt) >= 3:
                piso = numero_txt[0]
                grupo = f"Piso {piso}"
            else:
                grupo = "Otros"

            if contrato_activo:
                ultimo_pago = db.query(Pago).filter(
                    Pago.contrato_id == contrato_activo.id
                ).order_by(Pago.fecha_pago.desc(), Pago.id.desc()).first()

                if ultimo_pago:
                    ultimo_pago_id = ultimo_pago.id

                monto_contrato = float(contrato_activo.monto_mensual or 0.0)
                inquilino_nombre = contrato_activo.inquilino.nombre if contrato_activo.inquilino else None
                tipo_alquiler = contrato_activo.tipo_alquiler
                fecha_inicio = contrato_activo.fecha_inicio
                fecha_fin = contrato_activo.fecha_fin

                deuda_total = db.query(
                    func.coalesce(
                        func.sum(Cargo.monto - func.coalesce(Cargo.pagado_acumulado, 0.0)),
                        0.0
                    )
                ).filter(
                    Cargo.contrato_id == contrato_activo.id,
                    Cargo.estado.in_(["Pendiente", "Parcial"])
                ).scalar() or 0.0

                deuda_alquiler = db.query(
                    func.coalesce(
                        func.sum(Cargo.monto - func.coalesce(Cargo.pagado_acumulado, 0.0)),
                        0.0
                    )
                ).filter(
                    Cargo.contrato_id == contrato_activo.id,
                    Cargo.estado.in_(["Pendiente", "Parcial"]),
                    Cargo.concepto.in_(["ALQUILER_PRORRATA", "ALQUILER_MENSUAL", "ALQUILER_MENSUAL_AMPLIACION"]) |
                    Cargo.concepto.like("ALQUILER_DIARIO%") |
                    Cargo.concepto.like("ALQUILER_NOCHE%")
                ).scalar() or 0.0

                deuda_agua = db.query(
                    func.coalesce(
                        func.sum(Cargo.monto - func.coalesce(Cargo.pagado_acumulado, 0.0)),
                        0.0
                    )
                ).filter(
                    Cargo.contrato_id == contrato_activo.id,
                    Cargo.estado.in_(["Pendiente", "Parcial"]),
                    Cargo.concepto == "Agua"
                ).scalar() or 0.0

                deuda_luz = db.query(
                    func.coalesce(
                        func.sum(Cargo.monto - func.coalesce(Cargo.pagado_acumulado, 0.0)),
                        0.0
                    )
                ).filter(
                    Cargo.contrato_id == contrato_activo.id,
                    Cargo.estado.in_(["Pendiente", "Parcial"]),
                    Cargo.concepto == "Luz"
                ).scalar() or 0.0

                deuda_otros = db.query(
                    func.coalesce(
                        func.sum(Cargo.monto - func.coalesce(Cargo.pagado_acumulado, 0.0)),
                        0.0
                    )
                ).filter(
                    Cargo.contrato_id == contrato_activo.id,
                    Cargo.estado.in_(["Pendiente", "Parcial"]),
                    ~Cargo.concepto.in_(["ALQUILER_PRORRATA", "ALQUILER_MENSUAL", "ALQUILER_MENSUAL_AMPLIACION", "Agua", "Luz"]),
                    ~Cargo.concepto.like("ALQUILER_DIARIO%"),
                    ~Cargo.concepto.like("ALQUILER_NOCHE%")
                ).scalar() or 0.0

                ampliacion_pendiente = db.query(func.count(Cargo.id)).filter(
                    Cargo.contrato_id == contrato_activo.id,
                    Cargo.estado.in_(["Pendiente", "Parcial"]),
                    Cargo.concepto.like("ALQUILER_DIARIO_AMPLIACION%") |
                    Cargo.concepto.like("ALQUILER_NOCHE_AMPLIACION%") |
                    (Cargo.concepto == "ALQUILER_MENSUAL_AMPLIACION")
                ).scalar() or 0

                tiene_ampliacion_pendiente = ampliacion_pendiente > 0

                vencidos = db.query(func.count(Cargo.id)).filter(
                    Cargo.contrato_id == contrato_activo.id,
                    Cargo.estado.in_(["Pendiente", "Parcial"]),
                    Cargo.vencimiento < hoy
                ).scalar() or 0

                proximos_cargos = db.query(func.count(Cargo.id)).filter(
                    Cargo.contrato_id == contrato_activo.id,
                    Cargo.estado.in_(["Pendiente", "Parcial"]),
                    Cargo.vencimiento >= hoy,
                    Cargo.vencimiento <= (hoy + timedelta(days=3))
                ).scalar() or 0

                parciales = db.query(func.count(Cargo.id)).filter(
                    Cargo.contrato_id == contrato_activo.id,
                    Cargo.estado == "Parcial"
                ).scalar() or 0

                contrato_por_vencer = False
                if contrato_activo.fecha_fin:
                    contrato_por_vencer = hoy <= contrato_activo.fecha_fin <= (hoy + timedelta(days=3))

                moroso = vencidos > 0
                vence_pronto = (proximos_cargos > 0) or contrato_por_vencer
                pago_parcial = parciales > 0

                if moroso:
                    estado_financiero = "moroso"
                elif tiene_ampliacion_pendiente:
                    estado_financiero = "ampliacion"
                elif pago_parcial:
                    estado_financiero = "parcial"
                elif deuda_total > 0:
                    estado_financiero = "pendiente"
                else:
                    estado_financiero = "al_dia"

                cargos_pend = (
                    db.query(Cargo)
                    .filter(
                        Cargo.contrato_id == contrato_activo.id,
                        Cargo.estado.in_(["Pendiente", "Parcial"])
                    )
                    .order_by(Cargo.vencimiento.asc(), Cargo.periodo.asc())
                    .all()
                )

                periodos_unicos = []
                vistos = set()

                for cg in cargos_pend:
                    if cg.periodo not in vistos:
                        vistos.add(cg.periodo)
                        periodos_unicos.append(cg.periodo)

                periodos_pendientes = periodos_unicos

                ultima_agua = db.query(Lectura).filter(
                    Lectura.contrato_id == contrato_activo.id,
                    Lectura.servicio == "Agua"
                ).order_by(Lectura.fecha_registro.desc(), Lectura.id.desc()).first()

                ultima_luz = db.query(Lectura).filter(
                    Lectura.contrato_id == contrato_activo.id,
                    Lectura.servicio == "Luz"
                ).order_by(Lectura.fecha_registro.desc(), Lectura.id.desc()).first()

                ultima_lectura_agua = float(ultima_agua.lectura_actual) if ultima_agua else 0.0
                ultima_lectura_luz = float(ultima_luz.lectura_actual) if ultima_luz else 0.0
                fecha_lectura_agua = str(ultima_agua.fecha_registro) if ultima_agua else "-"
                fecha_lectura_luz = str(ultima_luz.fecha_registro) if ultima_luz else "-"

            propiedades_tablero.append({
                "id": prop.id,
                "numero": prop.numero,
                "tipo": prop.tipo,
                "grupo": grupo,
                "precio": float(prop.precio or 0),
                "estado": (prop.estado or "libre").lower(),
                "inquilino": inquilino_nombre,
                "tipo_alquiler": tipo_alquiler,
                "fecha_inicio": str(fecha_inicio) if fecha_inicio else "",
                "fecha_fin": str(fecha_fin) if fecha_fin else "",
                "moroso": moroso,
                "ampliacion_pendiente": tiene_ampliacion_pendiente,
                "deuda_total": float(deuda_total or 0.0),
                "deuda_alquiler": float(deuda_alquiler or 0.0),
                "deuda_agua": float(deuda_agua or 0.0),
                "deuda_luz": float(deuda_luz or 0.0),
                "deuda_otros": float(deuda_otros or 0.0),
                "vence_pronto": vence_pronto,
                "pago_parcial": pago_parcial,
                "estado_financiero": estado_financiero,
                "contrato_id": contrato_activo.id if contrato_activo else None,
                "ultimo_pago_id": ultimo_pago_id,
                "periodos_pendientes": periodos_pendientes,
                "ultima_lectura_agua": ultima_lectura_agua,
                "ultima_lectura_luz": ultima_lectura_luz,
                "fecha_lectura_agua": fecha_lectura_agua,
                "fecha_lectura_luz": fecha_lectura_luz,
                "monto_contrato": monto_contrato,
            })

        def ordenar_grupo(nombre):
            if nombre.startswith("Piso "):
                try:
                    return (0, int(nombre.replace("Piso ", "").strip()))
                except:
                    return (0, 999)
            elif nombre == "Locales":
                return (1, 0)
            elif nombre == "Cocheras":
                return (2, 0)
            else:
                return (3, 0)

        grupos_tablero = sorted(
            list({item["grupo"] for item in propiedades_tablero}),
            key=ordenar_grupo
        )

        resumen_grupos = {}
        for grupo in grupos_tablero:
            items_grupo = [item for item in propiedades_tablero if item["grupo"] == grupo]

            total = len(items_grupo)
            ocupadas_g = sum(1 for item in items_grupo if item["estado"] == "ocupado")
            libres_g = sum(1 for item in items_grupo if item["estado"] == "libre")
            morosos_g = sum(1 for item in items_grupo if item["moroso"])
            revision_g = sum(1 for item in items_grupo if item["estado"] == "revision")

            resumen_grupos[grupo] = {
                "total": total,
                "ocupadas": ocupadas_g,
                "libres": libres_g,
                "morosos": morosos_g,
                "revision": revision_g,
            }

        inquilinos_dashboard = db.query(Inquilino).filter(
            Inquilino.estado == "Activo"
        ).order_by(Inquilino.nombre.asc()).all()
        ingresos_hoy = db.query(
            func.coalesce(func.sum(Pago.monto), 0.0)
        ).filter(
            Pago.fecha_pago == hoy
        ).scalar() or 0.0

        inicio_mes = date(hoy.year, hoy.month, 1)

        ingresos_mes = db.query(
            func.coalesce(func.sum(Pago.monto), 0.0)
        ).filter(
            Pago.fecha_pago >= inicio_mes,
            Pago.fecha_pago <= hoy
        ).scalar() or 0.0

        deuda_total = db.query(
            func.coalesce(
                func.sum(Cargo.monto - func.coalesce(Cargo.pagado_acumulado, 0.0)),
                0.0
            )
        ).filter(
            Cargo.estado.in_(["Pendiente", "Parcial"])
        ).scalar() or 0.0

        pagos_hoy = db.query(
            func.count(Pago.id)
        ).filter(
            Pago.fecha_pago == hoy
        ).scalar() or 0


        # =========================
        # COBRANZA AUTOMÁTICA
        # =========================
        vencen_hoy = db.query(func.count(func.distinct(Cargo.contrato_id))).filter(
            Cargo.estado.in_(["Pendiente", "Parcial"]),
            Cargo.vencimiento == hoy
        ).scalar() or 0

        pagos_parciales = db.query(func.count(func.distinct(Cargo.contrato_id))).filter(
            Cargo.estado == "Parcial"
        ).scalar() or 0

        contratos_al_dia = db.query(func.count(Contrato.id)).filter(
            Contrato.estado == "Activo"
        ).all()

        contratos_con_deuda_ids = {
            x[0] for x in db.query(Cargo.contrato_id).filter(
                Cargo.estado.in_(["Pendiente", "Parcial"])
            ).distinct().all()
        }

        al_dia = 0
        for c in db.query(Contrato).filter(Contrato.estado == "Activo").all():
            if c.id not in contratos_con_deuda_ids:
                al_dia += 1

        cobranza_items = []
        contratos_cobranza = db.query(Contrato).options(
            joinedload(Contrato.inquilino),
            joinedload(Contrato.propiedad)
        ).filter(
            Contrato.estado == "Activo"
        ).all()

        for contrato in contratos_cobranza:
            cargos_pendientes = db.query(Cargo).filter(
                Cargo.contrato_id == contrato.id,
                Cargo.estado.in_(["Pendiente", "Parcial"])
            ).order_by(Cargo.vencimiento.asc()).all()

            if not cargos_pendientes:
                continue

            deuda_alquiler = 0.0
            deuda_agua = 0.0
            deuda_luz = 0.0
            deuda_otros = 0.0
            tiene_moroso = False
            tiene_parcial = False
            vence_hoy_flag = False

            for c in cargos_pendientes:
                saldo = round(float(c.monto or 0) - float(c.pagado_acumulado or 0), 2)
                if saldo <= 0:
                    continue

                concepto = (c.concepto or "").upper()

                if c.vencimiento and c.vencimiento < hoy:
                    tiene_moroso = True
                if c.vencimiento == hoy:
                    vence_hoy_flag = True
                if c.estado == "Parcial":
                    tiene_parcial = True

                if concepto in ["ALQUILER_PRORRATA", "ALQUILER_MENSUAL"] or concepto.startswith("ALQUILER_DIARIO") or concepto.startswith("ALQUILER_NOCHE"):
                    deuda_alquiler += saldo
                elif c.concepto == "Agua":
                    deuda_agua += saldo
                elif c.concepto == "Luz":
                    deuda_luz += saldo
                else:
                    deuda_otros += saldo

            deuda_total_cobranza = round(deuda_alquiler + deuda_agua + deuda_luz + deuda_otros, 2)
            if deuda_total_cobranza <= 0:
                continue

            estado_cobranza = "Pendiente"
            if tiene_moroso:
                estado_cobranza = "Moroso"
            elif tiene_parcial:
                estado_cobranza = "Pago parcial"
            elif vence_hoy_flag:
                estado_cobranza = "Vence hoy"

            cobranza_items.append({
                "contrato_id": contrato.id,
                "habitacion": contrato.propiedad.numero if contrato.propiedad else "-",
                "tipo": contrato.propiedad.tipo if contrato.propiedad else "-",
                "inquilino": contrato.inquilino.nombre if contrato.inquilino else "-",
                "estado": estado_cobranza,
                "deuda_total": deuda_total_cobranza,
                "deuda_alquiler": round(deuda_alquiler, 2),
                "deuda_agua": round(deuda_agua, 2),
                "deuda_luz": round(deuda_luz, 2),
                "deuda_otros": round(deuda_otros, 2),
            })
        
        # ---------------------------------
        # Últimos pagos de hoy para dashboard
        # ---------------------------------
        pagos_hoy_dashboard_db = db.query(Pago).options(
            joinedload(Pago.contrato).joinedload(Contrato.inquilino),
            joinedload(Pago.contrato).joinedload(Contrato.propiedad)
        ).filter(
            Pago.fecha_pago == hoy
        ).order_by(Pago.id.desc()).limit(8).all()

        ultimos_pagos_hoy = []

        for p in pagos_hoy_dashboard_db:
            propiedad = p.contrato.propiedad if p.contrato else None
            inquilino = p.contrato.inquilino if p.contrato else None
            cargo = db.query(Cargo).filter(Cargo.id == p.cargo_id).first() if p.cargo_id else None

            ultimos_pagos_hoy.append({
                "habitacion": propiedad.numero if propiedad else "-",
                "tipo": propiedad.tipo if propiedad else "-",
                "inquilino": inquilino.nombre if inquilino else "-",
                "concepto": cargo.concepto if cargo else "-",
                "periodo": cargo.periodo if cargo else "-",
                "monto": round(float(p.monto or 0.0), 2),
                "metodo": p.metodo if hasattr(p, "metodo") else "-",
                "fecha": p.fecha_pago
            })
        contratos_por_vencer = db.query(func.count(Contrato.id)).filter(
            Contrato.estado == "Activo",
            Contrato.fecha_fin != None,
            Contrato.fecha_fin >= hoy,
            Contrato.fecha_fin <= (hoy + timedelta(days=5))
        ).scalar() or 0

        unidades_con_deuda = db.query(func.count(func.distinct(Cargo.contrato_id))).filter(
            Cargo.estado.in_(["Pendiente", "Parcial"])
        ).scalar() or 0

        unidades_con_agua_pendiente = db.query(func.count(func.distinct(Cargo.contrato_id))).filter(
            Cargo.estado.in_(["Pendiente", "Parcial"]),
            Cargo.concepto == "Agua"
        ).scalar() or 0

        unidades_con_luz_pendiente = db.query(func.count(func.distinct(Cargo.contrato_id))).filter(
            Cargo.estado.in_(["Pendiente", "Parcial"]),
            Cargo.concepto == "Luz"
        ).scalar() or 0



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

            "propiedades_tablero": propiedades_tablero,
            "grupos_tablero": grupos_tablero,
            "resumen_grupos": resumen_grupos,
            "inquilinos_dashboard": inquilinos_dashboard,
            "ingresos_hoy": round(float(ingresos_hoy), 2),
            "ingresos_mes": round(float(ingresos_mes), 2),
            "deuda_total": round(float(deuda_total), 2),
            "pagos_hoy": pagos_hoy,
            "ingresos_hoy": round(float(ingresos_hoy), 2),
            "ingresos_mes": round(float(ingresos_mes), 2),
            "deuda_total": round(float(deuda_total), 2),
            "pagos_hoy": pagos_hoy,
            "ultimos_pagos_hoy": ultimos_pagos_hoy,
            "contratos_por_vencer": contratos_por_vencer,
            "unidades_con_deuda": unidades_con_deuda,
            "unidades_con_agua_pendiente": unidades_con_agua_pendiente,
            "unidades_con_luz_pendiente": unidades_con_luz_pendiente,
            "ok_pago": ok_pago,
            "ultimo_pago_id": ultimo_pago_id,
            "vencen_hoy": vencen_hoy,
            "pagos_parciales": pagos_parciales,
            "al_dia": al_dia,
            "cobranza_items": cobranza_items,
        })
    finally:
        db.close()
@app.post("/dashboard/registrar_lectura")
def dashboard_registrar_lectura(
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

        existe_cargo = db.query(Cargo).filter(
            Cargo.contrato_id == contrato_id,
            Cargo.concepto == servicio,
            Cargo.periodo == periodo
        ).first()
        if existe_cargo:
            raise HTTPException(400, f"Ya existe un cargo {servicio} para ese contrato en {periodo}")

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
        return RedirectResponse(url="/dashboard", status_code=303)
    finally:
        db.close()


@app.post("/dashboard/finalizar_contrato")
@login_required
def dashboard_finalizar_contrato(
    request: Request,
    contrato_id: int = Form(...)
):
    db = SessionLocal()
    try:
        contrato = db.query(Contrato).filter(Contrato.id == contrato_id).first()
        if not contrato:
            return RedirectResponse(url="/dashboard?error=contrato_no_encontrado", status_code=303)

        contrato.estado = "Finalizado"

        # liberar propiedad
        propiedad = db.query(Propiedad).filter(Propiedad.id == contrato.propiedad_id).first()
        if propiedad:
            propiedad.estado = "libre"

        # obtener cargos del contrato
        cargos = db.query(Cargo).filter(Cargo.contrato_id == contrato.id).all()

        for c in cargos:
            if c.estado in ["Pendiente", "Parcial"]:
                c.estado = "Anulado"

        db.commit()

        return RedirectResponse(url="/dashboard?ok=contrato_finalizado", status_code=303)

    except Exception as e:
        db.rollback()
        print("ERROR FINALIZAR CONTRATO:", e)
        return RedirectResponse(url="/dashboard?error=finalizar_contrato", status_code=303)

    finally:
        db.close()


# =========================================================
# PROPIEDADES (CRUD)
# =========================================================
@app.get("/propiedades", response_class=HTMLResponse)
@login_required
def propiedades_listar(request: Request):
    db = SessionLocal()
    try:
        propiedades = db.query(Propiedad).order_by(Propiedad.id.desc()).all()
        return templates.TemplateResponse(
            "index.html",
            {
                "request": request,
                "propiedades": propiedades
            }
        )
    finally:
        db.close()


@app.post("/propiedades/crear")
@login_required
@role_required("Administrador")
def propiedades_crear(
    request: Request,
    tipo: str = Form(...),
    numero: str = Form(...),
    precio: float = Form(...)
):
    db = SessionLocal()
    try:
        db.add(Propiedad(tipo=tipo, numero=numero, precio=float(precio), estado="libre"))
        db.commit()
        return RedirectResponse(url="/propiedades", status_code=303)
    finally:
        db.close()


@app.get("/propiedades/{propiedad_id}/editar", response_class=HTMLResponse)
@login_required
@role_required("Administrador")
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
@login_required
@role_required("Administrador")
def propiedades_actualizar(
    request: Request,
    propiedad_id: int,
    tipo: str = Form(...),
    numero: str = Form(...),
    precio: float = Form(...),
    estado: str = Form(...)
):
    db = SessionLocal()
    try:
        print("=== ACTUALIZAR PROPIEDAD ===")
        print("propiedad_id:", propiedad_id)

        p = db.query(Propiedad).filter(Propiedad.id == propiedad_id).first()
        print("propiedad encontrada:", p is not None)

        if not p:
            raise HTTPException(404, "Propiedad no encontrada")

        tipo = (tipo or "").strip()
        numero = (numero or "").strip()
        estado = (estado or "").strip().lower()

        if not tipo:
            raise HTTPException(400, "El tipo es obligatorio")
        if not numero:
            raise HTTPException(400, "El número es obligatorio")
        if float(precio or 0) <= 0:
            raise HTTPException(400, "El precio debe ser mayor a 0")
        if estado not in ["libre", "ocupado", "revision"]:
            raise HTTPException(400, "Estado inválido")

        existe_otra = db.query(Propiedad).filter(
            Propiedad.numero == numero,
            Propiedad.id != propiedad_id
        ).first()
        if existe_otra:
            raise HTTPException(400, f"Ya existe otra propiedad con número {numero}")

        p.tipo = tipo
        p.numero = numero
        p.precio = float(precio)
        p.estado = estado

        db.commit()
        return RedirectResponse(url="/propiedades", status_code=303)

    except Exception as e:
        db.rollback()
        import traceback
        print("=== ERROR ACTUALIZAR PROPIEDAD ===")
        print(type(e).__name__, str(e))
        traceback.print_exc()
        raise

    finally:
        db.close()


@app.post("/propiedades/{propiedad_id}/eliminar")
@login_required
@role_required("Administrador")
def propiedades_eliminar(request: Request, propiedad_id: int):
    db = SessionLocal()
    try:
        print("=== ELIMINAR PROPIEDAD ===")
        print("propiedad_id:", propiedad_id)

        p = db.query(Propiedad).filter(Propiedad.id == propiedad_id).first()
        print("propiedad encontrada:", p is not None)

        if not p:
            raise HTTPException(404, "Propiedad no encontrada")

        if (p.estado or "").lower() == "ocupado":
            raise HTTPException(400, "No se puede eliminar una propiedad ocupada")

        tiene_historial = db.query(Contrato).filter(
            Contrato.propiedad_id == propiedad_id
        ).first()
        print("tiene_historial:", tiene_historial is not None)

        if tiene_historial:
            raise HTTPException(400, "No se puede eliminar la propiedad porque tiene contratos registrados")

        db.delete(p)
        db.commit()
        return RedirectResponse(url="/propiedades", status_code=303)

    except Exception as e:
        db.rollback()
        import traceback
        print("=== ERROR ELIMINAR PROPIEDAD ===")
        print(type(e).__name__, str(e))
        traceback.print_exc()
        raise

    finally:
        db.close()


# =========================================================
# INQUILINOS (CRUD)
# =========================================================
@app.get("/inquilinos", response_class=HTMLResponse)
@login_required
def inquilinos_listar(request: Request):
    db = SessionLocal()
    try:
        inquilinos = db.query(Inquilino).order_by(Inquilino.id.desc()).all()
        return templates.TemplateResponse(
            "inquilinos.html",
            {
                "request": request,
                "inquilinos": inquilinos
            }
        )
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
@app.get("/inquilinos/{inquilino_id}/detalle", response_class=HTMLResponse)
def inquilino_detalle(request: Request, inquilino_id: int):
    db = SessionLocal()
    try:
        inquilino = db.query(Inquilino).filter(Inquilino.id == inquilino_id).first()
        if not inquilino:
            raise HTTPException(404, "Inquilino no encontrado")

        contratos = db.query(Contrato).options(
            joinedload(Contrato.propiedad)
        ).filter(
            Contrato.inquilino_id == inquilino_id
        ).order_by(Contrato.fecha_inicio.desc(), Contrato.id.desc()).all()

        cargos_pendientes = db.query(Cargo).options(
            joinedload(Cargo.contrato).joinedload(Contrato.propiedad)
        ).join(Contrato, Cargo.contrato_id == Contrato.id).filter(
            Contrato.inquilino_id == inquilino_id,
            Cargo.estado.in_(["Pendiente", "Parcial"])
        ).order_by(Cargo.vencimiento.asc(), Cargo.id.asc()).all()

        deuda_total = 0.0
        for c in cargos_pendientes:
            deuda_total += round(float(c.monto or 0) - float(c.pagado_acumulado or 0), 2)

        pagos_recientes = db.query(Pago).options(
            joinedload(Pago.contrato).joinedload(Contrato.propiedad)
        ).join(Contrato, Pago.contrato_id == Contrato.id).filter(
            Contrato.inquilino_id == inquilino_id
        ).order_by(Pago.fecha_pago.desc(), Pago.id.desc()).limit(10).all()

        bitacora = db.query(BitacoraCobranza).filter(
            BitacoraCobranza.inquilino_id == inquilino_id
        ).order_by(BitacoraCobranza.fecha_envio.desc(), BitacoraCobranza.id.desc()).limit(10).all()

        return templates.TemplateResponse("inquilino_detalle.html", {
            "request": request,
            "inquilino": inquilino,
            "contratos": contratos,
            "cargos_pendientes": cargos_pendientes,
            "deuda_total": round(float(deuda_total), 2),
            "pagos_recientes": pagos_recientes,
            "bitacora": bitacora,
        })
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

@app.post("/inquilinos/{inquilino_id}/desactivar")
@login_required
@role_required("Administrador")
def inquilinos_desactivar(request: Request, inquilino_id: int):
    db = SessionLocal()
    try:
        i = db.query(Inquilino).filter(Inquilino.id == inquilino_id).first()
        if not i:
            raise HTTPException(404, "Inquilino no encontrado")

        i.estado = "Inactivo"
        db.commit()
        return RedirectResponse(url="/inquilinos", status_code=303)
    finally:
        db.close()

@app.post("/inquilinos/{inquilino_id}/activar")
@login_required
@role_required("Administrador")
def inquilinos_activar(request: Request, inquilino_id: int):
    db = SessionLocal()
    try:
        i = db.query(Inquilino).filter(Inquilino.id == inquilino_id).first()
        if not i:
            raise HTTPException(404, "Inquilino no encontrado")

        i.estado = "Activo"
        db.commit()
        return RedirectResponse(url="/inquilinos", status_code=303)
    finally:
        db.close()

@app.post("/inquilinos/{inquilino_id}/eliminar")
@login_required
@role_required("Administrador")
def inquilinos_eliminar(request: Request, inquilino_id: int):
    db = SessionLocal()
    try:
        i = db.query(Inquilino).filter(Inquilino.id == inquilino_id).first()
        if not i:
            raise HTTPException(404, "Inquilino no encontrado")

        tiene_historial = db.query(Contrato).filter(
            Contrato.inquilino_id == inquilino_id
        ).first()

        if tiene_historial:
            raise HTTPException(
                400,
                "No se puede eliminar el inquilino porque tiene contratos registrados. Solo puedes desactivarlo."
            )

        db.delete(i)
        db.commit()
        return RedirectResponse(url="/inquilinos", status_code=303)

    except Exception:
        db.rollback()
        import traceback
        traceback.print_exc()
        raise
    finally:
        db.close()
# =========================================================
# CONTRATOS (CRUD) + mensual/diario/noche
# =========================================================
@app.get("/contratos", response_class=HTMLResponse)
@login_required
def contratos_listar(request: Request):
    db = SessionLocal()
    try:
        contratos = db.query(Contrato).options(
            joinedload(Contrato.propiedad),
            joinedload(Contrato.inquilino),
        ).order_by(Contrato.id.desc()).all()

        propiedades = db.query(Propiedad).order_by(Propiedad.id.desc()).all()
        inquilinos = db.query(Inquilino).order_by(Inquilino.id.desc()).all()

        return templates.TemplateResponse(
            "contratos.html",
            {
                "request": request,
                "contratos": contratos,
                "propiedades": propiedades,
                "inquilinos": inquilinos
            }
        )
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
        validar_rango_fechas(fecha_inicio, fecha_fin)

        tipo_alquiler = (tipo_alquiler or "mensual").strip().lower()

        prop = db.query(Propiedad).filter(Propiedad.id == propiedad_id).first()
        if not prop:
            raise HTTPException(404, "Propiedad no encontrada")

        inquilino = db.query(Inquilino).filter(Inquilino.id == inquilino_id).first()
        if not inquilino:
            raise HTTPException(404, "Inquilino no encontrado")

        existe_activo = db.query(Contrato).filter(
            Contrato.propiedad_id == propiedad_id,
            Contrato.estado == "Activo"
        ).first()
        if existe_activo or prop.estado != "libre":
            raise HTTPException(400, "La propiedad no está libre")

        monto_mensual_final = float(monto_mensual or 0.0)
        if tipo_alquiler in ["diario", "noche"]:
            monto_mensual_final = TARIFA_DIARIA_NOCHE

        contrato = Contrato(
            inquilino_id=inquilino_id,
            propiedad_id=propiedad_id,
            fecha_inicio=fecha_inicio,
            fecha_fin=fecha_fin,
            monto_mensual=monto_mensual_final,
            estado="Activo",
            tipo_alquiler=tipo_alquiler
        )
        db.add(contrato)
        db.flush()

        prop.estado = "ocupado"

        if tipo_alquiler == "mensual":
            _, monto_prorrata, venc = calcular_prorrata_mensual(float(monto_mensual_final), fecha_inicio)
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
        elif tipo_alquiler in ["diario", "noche"]:
            cantidad, total = calcular_total_diario_noche(
                fecha_inicio,
                fecha_fin,
                monto_mensual_final,
                tipo_alquiler
            )

            periodo = f"{fecha_inicio.isoformat()}_{fecha_fin.isoformat()}"

            if tipo_alquiler == "diario":
                concepto_base = "ALQUILER_DIARIO"
                unidad_texto = "día" if cantidad == 1 else "días"
            else:
                concepto_base = "ALQUILER_NOCHE"
                unidad_texto = "noche" if cantidad == 1 else "noches"

            db.add(Cargo(
                contrato_id=contrato.id,
                concepto=f"{concepto_base} ({cantidad} {unidad_texto})",
                periodo=periodo,
                monto=float(total),
                vencimiento=fecha_fin,
                estado="Pendiente",
                pagado_acumulado=0.0
            ))
        else:
            raise HTTPException(400, "Tipo de alquiler no válido")

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


from datetime import date
from fastapi import Form, HTTPException
from fastapi.responses import RedirectResponse

@app.post("/contratos/{contrato_id}/actualizar")
@login_required
@role_required("Administrador")
def contratos_actualizar(
    request: Request,
    contrato_id: int,
    inquilino_id: int = Form(...),
    propiedad_id: int = Form(...),
    fecha_inicio: date = Form(...),
    fecha_fin: date = Form(...),
    monto_mensual: float = Form(...),
    tipo_alquiler: str = Form("mensual"),
    estado: str = Form("Activo")
):
    db = SessionLocal()
    try:
        contrato = db.query(Contrato).filter(Contrato.id == contrato_id).first()
        if not contrato:
            raise HTTPException(status_code=404, detail="Contrato no encontrado")

        if fecha_fin < fecha_inicio:
            raise HTTPException(status_code=400, detail="La fecha fin no puede ser menor que la fecha inicio")

        tipo_alquiler = (tipo_alquiler or "mensual").strip().lower()
        estado = (estado or "Activo").strip()

        contrato.inquilino_id = inquilino_id
        contrato.propiedad_id = propiedad_id
        contrato.fecha_inicio = fecha_inicio
        contrato.fecha_fin = fecha_fin
        contrato.monto_mensual = float(monto_mensual or 0.0)
        contrato.tipo_alquiler = tipo_alquiler
        contrato.estado = estado

        db.commit()
        return RedirectResponse(url="/contratos", status_code=303)

    except Exception as e:
        db.rollback()
        import traceback
        print("ERROR REAL contratos_actualizar:")
        traceback.print_exc()
        raise

    finally:
        db.close()
@app.post("/contratos/{contrato_id}/eliminar")
@login_required
@role_required("Administrador")
def contratos_eliminar(request: Request, contrato_id: int):
    db = SessionLocal()
    try:
        c = db.query(Contrato).filter(Contrato.id == contrato_id).first()
        if not c:
            raise HTTPException(404, "Contrato no encontrado")

        prop = db.query(Propiedad).filter(Propiedad.id == c.propiedad_id).first()

        if prop:
            prop.estado = "libre"

        db.query(Pago).filter(Pago.contrato_id == contrato_id).delete(synchronize_session=False)
        db.query(Cargo).filter(Cargo.contrato_id == contrato_id).delete(synchronize_session=False)
        db.query(Lectura).filter(Lectura.contrato_id == contrato_id).delete(synchronize_session=False)

        db.delete(c)

        if prop:
            otro_activo = db.query(Contrato).filter(
                Contrato.propiedad_id == prop.id,
                Contrato.estado == "Activo",
                Contrato.id != contrato_id
            ).first()
            prop.estado = "ocupado" if otro_activo else "libre"

        db.commit()
        return RedirectResponse(url="/contratos", status_code=303)

    except Exception:
        db.rollback()
        import traceback
        traceback.print_exc()
        raise
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
@login_required
@role_required("Administrador")
def cargos_crear(contrato_id: int = Form(...), concepto: str = Form(...), periodo: str = Form(...), monto: float = Form(...), vencimiento: date = Form(...)):
    db = SessionLocal()
    try:
        existe = db.query(Cargo).filter(
            Cargo.contrato_id == contrato_id,
            Cargo.concepto == concepto,
            Cargo.periodo == periodo
        ).first()
        if existe:
            raise HTTPException(400, "Ya existe un cargo con ese concepto y periodo para el contrato")

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
@login_required
@role_required("Administrador")
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
@login_required
@role_required("Administrador")
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
@login_required
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

        return templates.TemplateResponse(
            "pagos.html",
            {
                "request": request,
                "pagos": pagos,
                "periodos_pendientes": periodos_pendientes
            }
        )
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

        ultimo_pago_id = None

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

            ultimo_pago_id = p.id

            recalcular_cargo(db, c.id)
            monto_disponible = round(monto_disponible - aplicar, 2)

            if monto_disponible <= 0:
                break

        db.commit()

        if ultimo_pago_id:
            return RedirectResponse(
                url=f"/dashboard?ok_pago=1&ultimo_pago_id={ultimo_pago_id}",
                status_code=303
            )

        return RedirectResponse(url="/dashboard", status_code=303)
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
@login_required
@role_required("Administrador")
def pagos_eliminar(request: Request, pago_id: int):
    db = SessionLocal()
    try:
        print("=== ELIMINAR PAGO ===")
        print("pago_id:", pago_id)

        pago = db.query(Pago).filter(Pago.id == pago_id).first()
        print("pago encontrado:", pago is not None)

        if not pago:
            raise HTTPException(404, "Pago no encontrado")

        cargo_id = pago.cargo_id
        print("cargo_id:", cargo_id)

        db.delete(pago)
        db.flush()
        print("pago eliminado en flush")

        if cargo_id:
            cargo = db.query(Cargo).filter(Cargo.id == cargo_id).first()
            print("cargo encontrado:", cargo is not None)

            if cargo:
                recalcular_cargo(db, cargo_id)
                print("cargo recalculado")

        db.commit()
        print("commit ok")
        return RedirectResponse(url="/pagos", status_code=303)

    except Exception as e:
        db.rollback()
        import traceback
        print("=== ERROR ELIMINAR PAGO ===")
        print("tipo:", type(e).__name__)
        print("mensaje:", str(e))
        traceback.print_exc()
        raise

    finally:
        db.close()

@app.get("/pagos/{pago_id}/recibo.pdf")
def descargar_recibo_pdf(pago_id: int, formato: str = "a4"):
    db = SessionLocal()
    try:
        pago = db.query(Pago).options(
            joinedload(Pago.contrato).joinedload(Contrato.inquilino),
            joinedload(Pago.contrato).joinedload(Contrato.propiedad),
            joinedload(Pago.cargo),
        ).filter(Pago.id == pago_id).first()

        if not pago:
            raise HTTPException(404, "Pago no encontrado")

        config = db.query(ConfiguracionEmpresa).first()

        if not config:
            config = ConfiguracionEmpresa(
                nombre_comercial="CREDIMAS",
                razon_social="",
                ruc="",
                direccion="",
                telefono="",
                whatsapp="",
                correo="",
                ciudad="",
                moneda="S/",
                mensaje_recibo="Gracias por su pago.",
                pie_pagina="Documento generado por CREDIMAS.",
                logo_url="",
                dia_corte=25
            )

        pdf_bytes = generar_recibo_pdf_pro(pago, config, formato=formato)
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
@login_required
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

        existe_cargo = db.query(Cargo).filter(
            Cargo.contrato_id == contrato_id,
            Cargo.concepto == servicio,
            Cargo.periodo == periodo
        ).first()
        if existe_cargo:
            raise HTTPException(400, f"Ya existe un cargo {servicio} para ese contrato en {periodo}")

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

        cargo = db.query(Cargo).filter(
            Cargo.contrato_id == l.contrato_id,
            Cargo.concepto == l.servicio,
            Cargo.periodo == l.periodo
        ).first()
        if cargo:
            tiene_pagos = db.query(Pago).filter(Pago.cargo_id == cargo.id).first()
            if tiene_pagos:
                raise HTTPException(400, "No se puede eliminar la lectura porque su cargo ya tiene pagos registrados")
            db.delete(cargo)

        db.delete(l)
        db.commit()
        return RedirectResponse(url="/lecturas", status_code=303)
    finally:
        db.close()
import traceback

@app.post("/dashboard/crear_contrato")
@login_required
def dashboard_crear_contrato(
    request: Request,
    propiedad_id: int = Form(...),
    inquilino_id: int = Form(...),
    fecha_inicio: date = Form(...),
    fecha_fin: date = Form(...),
    tipo_alquiler: str = Form("mensual"),
    monto_mensual: float = Form(0),
):
    db = SessionLocal()
    try:
        print("DEBUG crear_contrato ->", {
            "propiedad_id": propiedad_id,
            "inquilino_id": inquilino_id,
            "fecha_inicio": fecha_inicio,
            "fecha_fin": fecha_fin,
            "tipo_alquiler": tipo_alquiler,
            "monto_mensual": monto_mensual,
            "user_rol": request.session.get("user_rol"),
        })

        validar_rango_fechas(fecha_inicio, fecha_fin)

        tipo_alquiler = (tipo_alquiler or "mensual").strip().lower()

        if tipo_alquiler not in ["mensual", "diario", "noche"]:
            raise HTTPException(status_code=400, detail="Tipo de alquiler no válido")

        prop = db.query(Propiedad).filter(Propiedad.id == propiedad_id).first()
        if not prop:
            raise HTTPException(status_code=404, detail="Propiedad no encontrada")

        inquilino = db.query(Inquilino).filter(Inquilino.id == inquilino_id).first()
        if not inquilino:
            raise HTTPException(status_code=404, detail="Inquilino no encontrado")

        existe_activo = db.query(Contrato).filter(
            Contrato.propiedad_id == propiedad_id,
            Contrato.estado == "Activo"
        ).first()

        if existe_activo or (prop.estado or "").lower() != "libre":
            raise HTTPException(status_code=400, detail="La propiedad ya no está libre")

        if tipo_alquiler in ["diario", "noche"] and fecha_fin <= fecha_inicio:
            raise HTTPException(status_code=400, detail="La fecha fin debe ser mayor que la fecha inicio")

        # =========================
        # MONTO SEGURO SEGÚN ROL
        # =========================
        user_rol = request.session.get("user_rol", "")

        if user_rol == "Administrador":
            monto_final = float(monto_mensual or 0.0)
        else:
            # Operador NO puede alterar precios
            if tipo_alquiler == "mensual":
                monto_final = float(prop.precio or 0.0)
            elif tipo_alquiler in ["diario", "noche"]:
                monto_final = float(TARIFA_DIARIA_NOCHE)
            else:
                raise HTTPException(status_code=400, detail="Tipo de alquiler no válido")

        if monto_final <= 0:
            raise HTTPException(status_code=400, detail="El monto debe ser mayor a 0")

        contrato = Contrato(
            inquilino_id=inquilino_id,
            propiedad_id=propiedad_id,
            fecha_inicio=fecha_inicio,
            fecha_fin=fecha_fin,
            monto_mensual=monto_final,
            estado="Activo",
            tipo_alquiler=tipo_alquiler
        )
        db.add(contrato)
        db.flush()

        prop.estado = "ocupado"

        if tipo_alquiler == "mensual":
            _, monto_prorrata, venc = calcular_prorrata_mensual(monto_final, fecha_inicio)
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

        elif tipo_alquiler in ["diario", "noche"]:
            cantidad, total = calcular_total_diario_noche(
                fecha_inicio,
                fecha_fin,
                monto_final,
                tipo_alquiler
            )
            periodo = f"{fecha_inicio.isoformat()}_{fecha_fin.isoformat()}"

            if tipo_alquiler == "diario":
                unidad_texto = "día" if cantidad == 1 else "días"
                concepto_base = "ALQUILER_DIARIO"
            else:
                unidad_texto = "noche" if cantidad == 1 else "noches"
                concepto_base = "ALQUILER_NOCHE"

            db.add(Cargo(
                contrato_id=contrato.id,
                concepto=f"{concepto_base} ({cantidad} {unidad_texto})",
                periodo=periodo,
                monto=float(total),
                vencimiento=fecha_fin,
                estado="Pendiente",
                pagado_acumulado=0.0
            ))

        db.commit()
        return RedirectResponse(url="/dashboard?ok=contrato_creado", status_code=303)

    except HTTPException:
        db.rollback()
        raise

    except Exception as e:
        db.rollback()
        print("ERROR CREAR CONTRATO:")
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Error interno al crear contrato: {str(e)}")

    finally:
        db.close()
@app.post("/dashboard/registrar_pago")
def dashboard_registrar_pago(
    contrato_id: int = Form(...),
    periodo: str = Form(...),
    monto_total: float = Form(...),
    fecha_pago: str = Form(...),
    metodo: str = Form(...),
    tipo_pago: str = Form("todo"),
):
    db = SessionLocal()
    try:
        monto_disponible = round(float(monto_total), 2)
        if monto_disponible <= 0:
            raise HTTPException(400, "El monto debe ser mayor a cero")

        tipo_pago = (tipo_pago or "todo").strip().lower()

        query = db.query(Cargo).filter(
            Cargo.contrato_id == contrato_id,
            Cargo.periodo == periodo,
            Cargo.estado.in_(["Pendiente", "Parcial"])
        )

        if tipo_pago == "alquiler":
            query = query.filter(
                Cargo.concepto.in_(["ALQUILER_PRORRATA", "ALQUILER_MENSUAL"]) |
                Cargo.concepto.like("ALQUILER_DIARIO%") |
                Cargo.concepto.like("ALQUILER_NOCHE%")
            )
        elif tipo_pago == "agua":
            query = query.filter(Cargo.concepto == "Agua")
        elif tipo_pago == "luz":
            query = query.filter(Cargo.concepto == "Luz")
        elif tipo_pago == "otros":
            query = query.filter(
                ~Cargo.concepto.in_(["ALQUILER_PRORRATA", "ALQUILER_MENSUAL", "Agua", "Luz"]),
                ~Cargo.concepto.like("ALQUILER_DIARIO%"),
                ~Cargo.concepto.like("ALQUILER_NOCHE%")
            )

        cargos = query.order_by(Cargo.vencimiento.asc(), Cargo.id.asc()).all()

        if not cargos:
            raise HTTPException(404, "No hay cargos pendientes para esa selección")

        total_pendiente = 0.0
        for c in cargos:
            saldo = round(float(c.monto or 0) - float(c.pagado_acumulado or 0), 2)
            if saldo > 0:
                total_pendiente += saldo

        total_pendiente = round(total_pendiente, 2)

        if monto_disponible > total_pendiente:
            raise HTTPException(
                400,
                f"El monto excede el total pendiente seleccionado. Máximo: S/ {total_pendiente:.2f}"
            )

        try:
            if "/" in fecha_pago:
                fecha = datetime.strptime(fecha_pago, "%d/%m/%Y").date()
            else:
                fecha = date.fromisoformat(fecha_pago)
        except Exception:
            raise HTTPException(400, "La fecha de pago es inválida")

        ultimo_pago_id = None

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

            ultimo_pago_id = p.id

            recalcular_cargo(db, c.id)
            monto_disponible = round(monto_disponible - aplicar, 2)

            if monto_disponible <= 0:
                break

        db.commit()

        if ultimo_pago_id:
            return RedirectResponse(
                url=f"/dashboard?ok_pago=1&ultimo_pago_id={ultimo_pago_id}",
                status_code=303
            )

        return RedirectResponse(url="/dashboard", status_code=303)

    except HTTPException:
        db.rollback()
        raise
    finally:
        db.close()
# =========================================================
# REPORTES FINANCIEROS
# =========================================================
@app.get("/reportes", response_class=HTMLResponse)
@login_required
@role_required("Administrador")
def reportes(
    request: Request,
    desde: str = "",
    hasta: str = "",
    propiedad_id: int = 0,
    concepto: str = ""
):
    db = SessionLocal()
    try:
        hoy = date.today()
        inicio_mes = date(hoy.year, hoy.month, 1)

        fecha_desde = None
        fecha_hasta = None

        if desde:
            try:
                fecha_desde = datetime.strptime(desde, "%Y-%m-%d").date()
            except:
                fecha_desde = None

        if hasta:
            try:
                fecha_hasta = datetime.strptime(hasta, "%Y-%m-%d").date()
            except:
                fecha_hasta = None

        # -------------------------
        # Query base de pagos
        # -------------------------
        q_pagos = db.query(Pago).join(Cargo, Pago.cargo_id == Cargo.id)

        if fecha_desde:
            q_pagos = q_pagos.filter(Pago.fecha_pago >= fecha_desde)
        if fecha_hasta:
            q_pagos = q_pagos.filter(Pago.fecha_pago <= fecha_hasta)
        if propiedad_id:
            q_pagos = q_pagos.join(Contrato, Pago.contrato_id == Contrato.id).filter(
                Contrato.propiedad_id == propiedad_id
            )
        if concepto:
            q_pagos = q_pagos.filter(Cargo.concepto == concepto)

        pagos_filtrados = q_pagos.all()
        ingresos_filtrados = round(sum(float(p.monto or 0) for p in pagos_filtrados), 2)

        # -------------------------
        # Gastos filtrados
        # -------------------------
        q_gastos = db.query(Gasto).options(joinedload(Gasto.propiedad))

        if fecha_desde:
            q_gastos = q_gastos.filter(Gasto.fecha >= fecha_desde)
        if fecha_hasta:
            q_gastos = q_gastos.filter(Gasto.fecha <= fecha_hasta)
        if propiedad_id:
            q_gastos = q_gastos.filter(Gasto.propiedad_id == propiedad_id)

        gastos_filtrados_db = q_gastos.order_by(Gasto.fecha.desc(), Gasto.id.desc()).all()
        gastos_filtrados = round(sum(float(g.monto or 0) for g in gastos_filtrados_db), 2)

        # -------------------------
        # Tarjetas generales ingresos
        # -------------------------
        ingresos_hoy = db.query(
            func.coalesce(func.sum(Pago.monto), 0.0)
        ).filter(
            Pago.fecha_pago == hoy
        ).scalar() or 0.0

        ingresos_mes = db.query(
            func.coalesce(func.sum(Pago.monto), 0.0)
        ).filter(
            Pago.fecha_pago >= inicio_mes,
            Pago.fecha_pago <= hoy
        ).scalar() or 0.0

        # -------------------------
        # Tarjetas generales gastos
        # -------------------------
        gastos_hoy = db.query(
            func.coalesce(func.sum(Gasto.monto), 0.0)
        ).filter(
            Gasto.fecha == hoy
        ).scalar() or 0.0

        gastos_mes = db.query(
            func.coalesce(func.sum(Gasto.monto), 0.0)
        ).filter(
            Gasto.fecha >= inicio_mes,
            Gasto.fecha <= hoy
        ).scalar() or 0.0

        utilidad_mes = round(float(ingresos_mes) - float(gastos_mes), 2)
        utilidad_filtrada = round(float(ingresos_filtrados) - float(gastos_filtrados), 2)

        deuda_total = db.query(
            func.coalesce(
                func.sum(Cargo.monto - func.coalesce(Cargo.pagado_acumulado, 0.0)),
                0.0
            )
        ).filter(
            Cargo.estado.in_(["Pendiente", "Parcial"])
        ).scalar() or 0.0

        total_morosos = db.query(func.count(func.distinct(Cargo.contrato_id))).filter(
            Cargo.estado.in_(["Pendiente", "Parcial"]),
            Cargo.vencimiento < hoy
        ).scalar() or 0

        ocupadas = db.query(func.count(Propiedad.id)).filter(
            func.lower(Propiedad.estado) == "ocupado"
        ).scalar() or 0

        libres = db.query(func.count(Propiedad.id)).filter(
            func.lower(Propiedad.estado) == "libre"
        ).scalar() or 0

        # -------------------------
        # Ingresos por concepto con filtro actual
        # -------------------------
        total_alquiler_q = db.query(
            func.coalesce(func.sum(Pago.monto), 0.0)
        ).join(Cargo, Pago.cargo_id == Cargo.id)

        if fecha_desde:
            total_alquiler_q = total_alquiler_q.filter(Pago.fecha_pago >= fecha_desde)
        if fecha_hasta:
            total_alquiler_q = total_alquiler_q.filter(Pago.fecha_pago <= fecha_hasta)
        if propiedad_id:
            total_alquiler_q = total_alquiler_q.join(Contrato, Pago.contrato_id == Contrato.id).filter(
                Contrato.propiedad_id == propiedad_id
            )

        total_alquiler = total_alquiler_q.filter(
            (Cargo.concepto == "ALQUILER_PRORRATA") |
            (Cargo.concepto == "ALQUILER_MENSUAL") |
            (Cargo.concepto.like("ALQUILER_DIARIO%")) |
            (Cargo.concepto.like("ALQUILER_NOCHE%"))
        ).scalar() or 0.0

        q_agua = db.query(func.coalesce(func.sum(Pago.monto), 0.0)).join(Cargo, Pago.cargo_id == Cargo.id)
        q_luz = db.query(func.coalesce(func.sum(Pago.monto), 0.0)).join(Cargo, Pago.cargo_id == Cargo.id)
        q_total = db.query(func.coalesce(func.sum(Pago.monto), 0.0))

        if fecha_desde:
            q_agua = q_agua.filter(Pago.fecha_pago >= fecha_desde)
            q_luz = q_luz.filter(Pago.fecha_pago >= fecha_desde)
            q_total = q_total.filter(Pago.fecha_pago >= fecha_desde)
        if fecha_hasta:
            q_agua = q_agua.filter(Pago.fecha_pago <= fecha_hasta)
            q_luz = q_luz.filter(Pago.fecha_pago <= fecha_hasta)
            q_total = q_total.filter(Pago.fecha_pago <= fecha_hasta)
        if propiedad_id:
            q_agua = q_agua.join(Contrato, Pago.contrato_id == Contrato.id).filter(Contrato.propiedad_id == propiedad_id)
            q_luz = q_luz.join(Contrato, Pago.contrato_id == Contrato.id).filter(Contrato.propiedad_id == propiedad_id)
            q_total = q_total.join(Contrato, Pago.contrato_id == Contrato.id).filter(Contrato.propiedad_id == propiedad_id)

        total_agua = q_agua.filter(Cargo.concepto == "Agua").scalar() or 0.0
        total_luz = q_luz.filter(Cargo.concepto == "Luz").scalar() or 0.0
        total_general = q_total.scalar() or 0.0

        total_otros = round(float(total_general) - float(total_alquiler) - float(total_agua) - float(total_luz), 2)
        if total_otros < 0:
            total_otros = 0.0

        ingresos_por_concepto = [
            {"concepto": "Alquiler", "total": round(float(total_alquiler), 2)},
            {"concepto": "Agua", "total": round(float(total_agua), 2)},
            {"concepto": "Luz", "total": round(float(total_luz), 2)},
            {"concepto": "Otros", "total": round(float(total_otros), 2)},
        ]

        # -------------------------
        # Gastos por categoría
        # -------------------------
        gastos_por_categoria_raw = db.query(
            Gasto.categoria,
            func.coalesce(func.sum(Gasto.monto), 0.0)
        )

        if fecha_desde:
            gastos_por_categoria_raw = gastos_por_categoria_raw.filter(Gasto.fecha >= fecha_desde)
        if fecha_hasta:
            gastos_por_categoria_raw = gastos_por_categoria_raw.filter(Gasto.fecha <= fecha_hasta)
        if propiedad_id:
            gastos_por_categoria_raw = gastos_por_categoria_raw.filter(Gasto.propiedad_id == propiedad_id)

        gastos_por_categoria_raw = gastos_por_categoria_raw.group_by(Gasto.categoria).all()

        gastos_por_categoria = [
            {"categoria": x[0] or "Sin categoría", "total": round(float(x[1] or 0.0), 2)}
            for x in gastos_por_categoria_raw
        ]

        # -------------------------
        # Morosos
        # -------------------------
        cargos_morosos = db.query(Cargo).options(
            joinedload(Cargo.contrato).joinedload(Contrato.inquilino),
            joinedload(Cargo.contrato).joinedload(Contrato.propiedad)
        ).filter(
            Cargo.estado.in_(["Pendiente", "Parcial"]),
            Cargo.vencimiento < hoy
        ).order_by(Cargo.vencimiento.asc()).all()

        if propiedad_id:
            cargos_morosos = [c for c in cargos_morosos if c.contrato and c.contrato.propiedad_id == propiedad_id]

        if concepto:
            cargos_morosos = [c for c in cargos_morosos if c.concepto == concepto]

        lista_morosos = []
        for c in cargos_morosos:
            saldo = round(float(c.monto or 0.0) - float(c.pagado_acumulado or 0.0), 2)
            if saldo <= 0:
                continue

            propiedad = c.contrato.propiedad if c.contrato else None
            inquilino = c.contrato.inquilino if c.contrato else None

            lista_morosos.append({
                "habitacion": propiedad.numero if propiedad else "-",
                "tipo": propiedad.tipo if propiedad else "-",
                "inquilino": inquilino.nombre if inquilino else "-",
                "concepto": c.concepto,
                "periodo": c.periodo,
                "vencimiento": c.vencimiento,
                "saldo": saldo,
                "estado": c.estado
            })

        # -------------------------
        # Pagos detallados
        # -------------------------
        q_detalle_pagos = db.query(Pago).options(
            joinedload(Pago.contrato).joinedload(Contrato.inquilino),
            joinedload(Pago.contrato).joinedload(Contrato.propiedad),
            joinedload(Pago.cargo)
        ).order_by(Pago.fecha_pago.desc(), Pago.id.desc())

        if fecha_desde:
            q_detalle_pagos = q_detalle_pagos.filter(Pago.fecha_pago >= fecha_desde)
        if fecha_hasta:
            q_detalle_pagos = q_detalle_pagos.filter(Pago.fecha_pago <= fecha_hasta)
        if propiedad_id:
            q_detalle_pagos = q_detalle_pagos.join(Contrato, Pago.contrato_id == Contrato.id).filter(
                Contrato.propiedad_id == propiedad_id
            )
        if concepto:
            q_detalle_pagos = q_detalle_pagos.join(Cargo, Pago.cargo_id == Cargo.id).filter(
                Cargo.concepto == concepto
            )

        pagos_detallados_db = q_detalle_pagos.all()

        pagos_detallados = []
        for p in pagos_detallados_db:
            propiedad = p.contrato.propiedad if p.contrato else None
            inquilino = p.contrato.inquilino if p.contrato else None
            cargo = p.cargo if hasattr(p, "cargo") else None

            pagos_detallados.append({
                "fecha": p.fecha_pago,
                "habitacion": propiedad.numero if propiedad else "-",
                "tipo": propiedad.tipo if propiedad else "-",
                "inquilino": inquilino.nombre if inquilino else "-",
                "concepto": cargo.concepto if cargo else "-",
                "periodo": cargo.periodo if cargo else "-",
                "monto": round(float(p.monto or 0.0), 2),
                "metodo": p.metodo if hasattr(p, "metodo") else "-"
            })

        # -------------------------
        # Pagos de hoy por cuarto
        # -------------------------
        pagos_hoy_db = db.query(Pago).options(
            joinedload(Pago.contrato).joinedload(Contrato.propiedad)
        ).filter(
            Pago.fecha_pago == hoy
        ).all()

        habitaciones_pagaron_hoy = len({
            p.contrato.propiedad_id
            for p in pagos_hoy_db
            if p.contrato and p.contrato.propiedad_id
        })

        pagos_hoy_query = db.query(Pago).options(
            joinedload(Pago.contrato).joinedload(Contrato.inquilino),
            joinedload(Pago.contrato).joinedload(Contrato.propiedad)
        ).filter(
            Pago.fecha_pago == hoy
        ).order_by(Pago.id.desc())

        pagos_hoy_db = pagos_hoy_query.all()

        pagos_hoy_detallados = []
        total_cobrado_hoy = 0.0

        for p in pagos_hoy_db:
            propiedad = p.contrato.propiedad if p.contrato else None
            inquilino = p.contrato.inquilino if p.contrato else None
            cargo = db.query(Cargo).filter(Cargo.id == p.cargo_id).first() if p.cargo_id else None

            monto_pago = round(float(p.monto or 0.0), 2)
            total_cobrado_hoy += monto_pago

            pagos_hoy_detallados.append({
                "fecha": p.fecha_pago,
                "habitacion": propiedad.numero if propiedad else "-",
                "tipo": propiedad.tipo if propiedad else "-",
                "inquilino": inquilino.nombre if inquilino else "-",
                "concepto": cargo.concepto if cargo else "-",
                "periodo": cargo.periodo if cargo else "-",
                "monto": monto_pago,
                "metodo": p.metodo if hasattr(p, "metodo") else "-"
            })

        # -------------------------
        # Resumen por propiedad
        # -------------------------
        propiedades = db.query(Propiedad).order_by(Propiedad.numero.asc()).all()
        resumen_propiedades = []

        for prop in propiedades:
            if propiedad_id and prop.id != propiedad_id:
                continue

            contratos_ids = [x.id for x in db.query(Contrato.id).filter(Contrato.propiedad_id == prop.id).all()]

            if contratos_ids:
                q_cobrado = db.query(func.coalesce(func.sum(Pago.monto), 0.0)).filter(Pago.contrato_id.in_(contratos_ids))
                if fecha_desde:
                    q_cobrado = q_cobrado.filter(Pago.fecha_pago >= fecha_desde)
                if fecha_hasta:
                    q_cobrado = q_cobrado.filter(Pago.fecha_pago <= fecha_hasta)

                total_cobrado = q_cobrado.scalar() or 0.0

                q_pendiente = db.query(
                    func.coalesce(
                        func.sum(Cargo.monto - func.coalesce(Cargo.pagado_acumulado, 0.0)),
                        0.0
                    )
                ).filter(
                    Cargo.contrato_id.in_(contratos_ids),
                    Cargo.estado.in_(["Pendiente", "Parcial"])
                )

                if concepto:
                    q_pendiente = q_pendiente.filter(Cargo.concepto == concepto)

                total_pendiente = q_pendiente.scalar() or 0.0
            else:
                total_cobrado = 0.0
                total_pendiente = 0.0

            q_gastos_prop = db.query(func.coalesce(func.sum(Gasto.monto), 0.0)).filter(
                Gasto.propiedad_id == prop.id
            )
            if fecha_desde:
                q_gastos_prop = q_gastos_prop.filter(Gasto.fecha >= fecha_desde)
            if fecha_hasta:
                q_gastos_prop = q_gastos_prop.filter(Gasto.fecha <= fecha_hasta)

            total_gastos_prop = q_gastos_prop.scalar() or 0.0

            resumen_propiedades.append({
                "id": prop.id,
                "numero": prop.numero,
                "tipo": prop.tipo,
                "estado": prop.estado,
                "total_cobrado": round(float(total_cobrado), 2),
                "total_pendiente": round(float(total_pendiente), 2),
                "total_gastos": round(float(total_gastos_prop), 2),
            })

        propiedades_filtro = db.query(Propiedad).order_by(Propiedad.numero.asc()).all()

        # -------------------------
        # Gráfico mensual ingresos vs gastos
        # -------------------------
        chart_labels = []
        chart_values = []
        chart_gastos_values = []

        meses_nombres = {
            1: "Ene", 2: "Feb", 3: "Mar", 4: "Abr",
            5: "May", 6: "Jun", 7: "Jul", 8: "Ago",
            9: "Sep", 10: "Oct", 11: "Nov", 12: "Dic"
        }

        for mes in range(1, 13):
            inicio = date(hoy.year, mes, 1)
            if mes == 12:
                fin = date(hoy.year + 1, 1, 1) - timedelta(days=1)
            else:
                fin = date(hoy.year, mes + 1, 1) - timedelta(days=1)

            total_ing_mes_q = db.query(func.coalesce(func.sum(Pago.monto), 0.0)).filter(
                Pago.fecha_pago >= inicio,
                Pago.fecha_pago <= fin
            )

            total_gasto_mes_q = db.query(func.coalesce(func.sum(Gasto.monto), 0.0)).filter(
                Gasto.fecha >= inicio,
                Gasto.fecha <= fin
            )

            if propiedad_id:
                total_ing_mes_q = total_ing_mes_q.join(Contrato, Pago.contrato_id == Contrato.id).filter(
                    Contrato.propiedad_id == propiedad_id
                )
                total_gasto_mes_q = total_gasto_mes_q.filter(
                    Gasto.propiedad_id == propiedad_id
                )

            total_ing_mes = total_ing_mes_q.scalar() or 0.0
            total_gasto_mes = total_gasto_mes_q.scalar() or 0.0

            chart_labels.append(meses_nombres[mes])
            chart_values.append(round(float(total_ing_mes), 2))
            chart_gastos_values.append(round(float(total_gasto_mes), 2))

        return templates.TemplateResponse("reportes.html", {
            "request": request,
            "ingresos_hoy": round(float(ingresos_hoy), 2),
            "ingresos_mes": round(float(ingresos_mes), 2),
            "gastos_hoy": round(float(gastos_hoy), 2),
            "gastos_mes": round(float(gastos_mes), 2),
            "utilidad_mes": round(float(utilidad_mes), 2),
            "utilidad_filtrada": round(float(utilidad_filtrada), 2),
            "gastos_filtrados": round(float(gastos_filtrados), 2),
            "deuda_total": round(float(deuda_total), 2),
            "total_morosos": total_morosos,
            "ocupadas": ocupadas,
            "libres": libres,
            "ingresos_filtrados": ingresos_filtrados,
            "ingresos_por_concepto": ingresos_por_concepto,
            "gastos_por_categoria": gastos_por_categoria,
            "lista_morosos": lista_morosos,
            "resumen_propiedades": resumen_propiedades,
            "propiedades_filtro": propiedades_filtro,
            "desde": desde,
            "hasta": hasta,
            "propiedad_id": propiedad_id,
            "concepto": concepto,
            "pagos_detallados": pagos_detallados,
            "habitaciones_pagaron_hoy": habitaciones_pagaron_hoy,
            "pagos_hoy_detallados": pagos_hoy_detallados,
            "total_cobrado_hoy": round(float(total_cobrado_hoy), 2),
            "chart_labels": chart_labels,
            "chart_values": chart_values,
            "chart_gastos_values": chart_gastos_values,
        })
    finally:
        db.close()


# =========================================================
# ESTADO DE CUENTA PDF
# =========================================================
def estado_cuenta_pdf(inq: Inquilino, cargos: list[Cargo], config) -> bytes:
    nombre_negocio = config.nombre_comercial or "CREDIMAS"
    razon_social = config.razon_social or ""
    ruc = config.ruc or ""
    direccion = config.direccion or ""
    telefono = config.telefono or ""
    whatsapp_empresa = config.whatsapp or ""
    moneda = config.moneda or "S/"
    pie_pagina = config.pie_pagina or "Documento generado por CREDIMAS."

    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    w, h = A4
    M = 16 * mm
    y = h - M

    # =========================
    # CABECERA
    # =========================
    c.setFillColor(colors.HexColor("#111827"))
    c.rect(M, y - 28*mm, w - 2*M, 28*mm, fill=1, stroke=0)

    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 16)
    c.drawString(M + 5*mm, y - 9*mm, nombre_negocio)

    c.setFont("Helvetica-Bold", 11)
    c.drawString(M + 5*mm, y - 16*mm, "ESTADO DE CUENTA - INQUILINO")

    info_y = y - 22*mm
    c.setFont("Helvetica", 8)

    if razon_social:
        c.drawString(M + 5*mm, info_y, razon_social)
        info_y -= 4*mm

    if ruc:
        c.drawString(M + 5*mm, info_y, f"RUC: {ruc}")
        info_y -= 4*mm

    if direccion:
        c.drawString(M + 5*mm, info_y, direccion)
        info_y -= 4*mm

    contacto_empresa = ""
    if telefono:
        contacto_empresa += f"Tel: {telefono}"
    if whatsapp_empresa:
        if contacto_empresa:
            contacto_empresa += " | "
        contacto_empresa += f"WhatsApp: {whatsapp_empresa}"

    if contacto_empresa:
        c.drawString(M + 5*mm, info_y, contacto_empresa)

    c.setFillColor(colors.black)
    y -= 34 * mm

    # =========================
    # DATOS DEL CLIENTE
    # =========================
    c.setFont("Helvetica", 10)
    c.drawString(M, y, f"Fecha emisión: {date.today().isoformat()}")
    y -= 6 * mm

    c.setFont("Helvetica-Bold", 11)
    c.drawString(M, y, f"Inquilino: {inq.nombre}")
    y -= 5 * mm

    c.setFont("Helvetica", 10)
    c.drawString(M, y, f"DNI: {inq.dni}    WhatsApp: {inq.whatsapp or '-'}")
    y -= 8 * mm

    # =========================
    # ENCABEZADO TABLA
    # =========================
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

    # =========================
    # DETALLE DE CARGOS
    # =========================
    for item in cargos:
        saldo = float(item.monto or 0.0) - float(item.pagado_acumulado or 0.0)
        if saldo < 0:
            saldo = 0.0

        if y < 25 * mm:
            c.showPage()
            y = h - M

            # Repite cabecera simple en nueva página
            c.setFont("Helvetica-Bold", 13)
            c.drawString(M, y, f"{nombre_negocio} - ESTADO DE CUENTA")
            y -= 8*mm

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

            c.setFont("Helvetica", 9)

        prop = "-"
        if item.contrato and item.contrato.propiedad:
            prop = item.contrato.propiedad.numero

        c.drawString(M + 2*mm, y, str(item.vencimiento))
        c.drawString(M + 25*mm, y, str(prop))
        c.drawString(M + 60*mm, y, str(item.concepto))
        c.drawString(M + 100*mm, y, str(item.periodo))
        c.drawRightString(w - M - 55*mm, y, f"{moneda} {float(item.monto or 0):.2f}")
        c.drawRightString(w - M - 28*mm, y, f"{moneda} {float(item.pagado_acumulado or 0):.2f}")
        c.drawRightString(w - M - 2*mm, y, f"{moneda} {saldo:.2f}")
        y -= 6*mm

        total_monto += float(item.monto or 0.0)
        total_pagado += float(item.pagado_acumulado or 0.0)
        total_saldo += float(saldo)

    # =========================
    # TOTALES
    # =========================
    y -= 4*mm
    c.setStrokeColor(colors.HexColor("#111827"))
    c.line(M, y, w - M, y)
    y -= 8*mm

    c.setFont("Helvetica-Bold", 11)
    c.drawRightString(w - M, y, f"TOTAL MONTO: {moneda} {total_monto:.2f}")
    y -= 6*mm
    c.drawRightString(w - M, y, f"TOTAL PAGADO: {moneda} {total_pagado:.2f}")
    y -= 6*mm
    c.drawRightString(w - M, y, f"TOTAL SALDO: {moneda} {total_saldo:.2f}")

    # =========================
    # PIE
    # =========================
    c.setFont("Helvetica-Oblique", 8)
    c.setFillColor(colors.HexColor("#6B7280"))
    c.drawString(M, 12*mm, pie_pagina)

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

        config = db.query(ConfiguracionEmpresa).first()

        if not config:
            config = ConfiguracionEmpresa(
                nombre_comercial="CREDIMAS",
                razon_social="",
                ruc="",
                direccion="",
                telefono="",
                whatsapp="",
                correo="",
                ciudad="",
                moneda="S/",
                mensaje_recibo="Gracias por su pago.",
                pie_pagina="Documento generado por CREDIMAS.",
                logo_url="",
                dia_corte=25
            )

        pdf_bytes = estado_cuenta_pdf(inq, cargos, config)
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

        config = db.query(ConfiguracionEmpresa).first()

        if not config:
            config = ConfiguracionEmpresa(
                nombre_comercial="CREDIMAS",
                razon_social="",
                ruc="",
                direccion="",
                telefono="",
                whatsapp="",
                correo="",
                ciudad="",
                moneda="S/",
                mensaje_recibo="Gracias por su pago.",
                pie_pagina="Documento generado por CREDIMAS.",
                logo_url="",
                dia_corte=25
            )

        nombre_negocio = config.nombre_comercial or "CREDIMAS"
        moneda = config.moneda or "S/"
        mensaje_recibo = config.mensaje_recibo or "Gracias por su pago."

        inq = pago.contrato.inquilino
        prop = pago.contrato.propiedad

        numero = normalizar_numero_pe(inq.whatsapp or inq.telefono or "")
        if not numero:
            raise HTTPException(400, "El inquilino no tiene WhatsApp o teléfono registrado")

        base_url = str(request.base_url).rstrip("/")
        link_recibo = f"{base_url}/pagos/{pago.id}/recibo.pdf"

        mensaje = (
            f"{nombre_negocio}\n\n"
            f"Hola {inq.nombre}, se registró su pago correctamente.\n\n"
            f"Unidad: {prop.tipo} {prop.numero}\n"
            f"Monto pagado: {moneda} {float(pago.monto or 0):.2f}\n"
            f"Fecha: {pago.fecha_pago}\n"
            f"Método: {pago.metodo}\n"
            f"Recibo: {link_recibo}\n\n"
            f"{mensaje_recibo}"
        )

        url = f"https://wa.me/{numero}?text={quote(mensaje)}"
        return RedirectResponse(url=url, status_code=302)

    finally:
        db.close()
@app.get("/propiedades/{propiedad_id}/detalle", response_class=HTMLResponse)
def propiedad_detalle(request: Request, propiedad_id: int):
    db = SessionLocal()
    try:
        propiedad = db.query(Propiedad).filter(Propiedad.id == propiedad_id).first()
        if not propiedad:
            raise HTTPException(404, "Propiedad no encontrada")

        contrato = db.query(Contrato).options(
            joinedload(Contrato.inquilino),
            joinedload(Contrato.propiedad)
        ).filter(
            Contrato.propiedad_id == propiedad_id,
            Contrato.estado == "Activo"
        ).first()

        deuda_total = 0.0
        total_pagado = 0.0
        cargos_pendientes = []
        pagos_recientes = []

        if contrato:
            deuda_total = db.query(
                func.coalesce(
                    func.sum(Cargo.monto - func.coalesce(Cargo.pagado_acumulado, 0.0)),
                    0.0
                )
            ).filter(
                Cargo.contrato_id == contrato.id,
                Cargo.estado.in_(["Pendiente", "Parcial"])
            ).scalar() or 0.0

            total_pagado = db.query(
                func.coalesce(func.sum(Pago.monto), 0.0)
            ).filter(
                Pago.contrato_id == contrato.id
            ).scalar() or 0.0

            cargos_pendientes = db.query(Cargo).filter(
                Cargo.contrato_id == contrato.id,
                Cargo.estado.in_(["Pendiente", "Parcial"])
            ).order_by(Cargo.vencimiento.asc()).all()

            pagos_recientes = db.query(Pago).filter(
                Pago.contrato_id == contrato.id
            ).order_by(Pago.fecha_pago.desc(), Pago.id.desc()).limit(10).all()

        return templates.TemplateResponse("propiedad_detalle.html", {
            "request": request,
            "propiedad": propiedad,
            "contrato": contrato,
            "deuda_total": float(deuda_total or 0.0),
            "total_pagado": float(total_pagado or 0.0),
            "cargos_pendientes": cargos_pendientes,
            "pagos_recientes": pagos_recientes,
        })
    finally:
        db.close()
@app.get("/reportes/exportar")
def exportar_reportes():
    db = SessionLocal()
    try:
        wb = Workbook()
        ws = wb.active
        ws.title = "Reporte financiero"

        ws.append(["REPORTE FINANCIERO CREDIMAS"])
        ws.append([])

        ws.append(["Habitación", "Tipo", "Estado", "Total cobrado", "Total pendiente"])

        propiedades = db.query(Propiedad).order_by(Propiedad.numero.asc()).all()

        for prop in propiedades:
            contratos_ids = [x.id for x in db.query(Contrato.id).filter(Contrato.propiedad_id == prop.id).all()]

            if contratos_ids:
                total_cobrado = db.query(
                    func.coalesce(func.sum(Pago.monto), 0.0)
                ).filter(
                    Pago.contrato_id.in_(contratos_ids)
                ).scalar() or 0.0

                total_pendiente = db.query(
                    func.coalesce(
                        func.sum(Cargo.monto - func.coalesce(Cargo.pagado_acumulado, 0.0)),
                        0.0
                    )
                ).filter(
                    Cargo.contrato_id.in_(contratos_ids),
                    Cargo.estado.in_(["Pendiente", "Parcial"])
                ).scalar() or 0.0
            else:
                total_cobrado = 0.0
                total_pendiente = 0.0

            ws.append([
                prop.numero,
                prop.tipo,
                prop.estado,
                round(float(total_cobrado), 2),
                round(float(total_pendiente), 2),
            ])

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        return StreamingResponse(
            output,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=reporte_financiero.xlsx"}
        )
    finally:
        db.close()
@app.get("/pagos/{pago_id}/whatsapp")
def enviar_recibo_whatsapp(pago_id: int, request: Request):
    db = SessionLocal()
    try:
        pago = db.query(Pago).options(
            joinedload(Pago.contrato).joinedload(Contrato.inquilino),
            joinedload(Pago.contrato).joinedload(Contrato.propiedad),
            joinedload(Pago.cargo),
        ).filter(Pago.id == pago_id).first()

        if not pago:
            raise HTTPException(404, "Pago no encontrado")

        config = db.query(ConfiguracionEmpresa).first()

        if not config:
            config = ConfiguracionEmpresa(
                nombre_comercial="CREDIMAS",
                razon_social="",
                ruc="",
                direccion="",
                telefono="",
                whatsapp="",
                correo="",
                ciudad="",
                moneda="S/",
                mensaje_recibo="Gracias por su pago.",
                pie_pagina="Documento generado por CREDIMAS.",
                logo_url="",
                dia_corte=25
            )

        nombre_negocio = config.nombre_comercial or "CREDIMAS"
        moneda = config.moneda or "S/"
        mensaje_recibo = config.mensaje_recibo or "Gracias por su pago."

        inq = pago.contrato.inquilino if pago.contrato else None
        prop = pago.contrato.propiedad if pago.contrato else None

        if not inq:
            raise HTTPException(400, "El pago no tiene inquilino asociado")

        numero = normalizar_numero_pe(inq.whatsapp or inq.telefono or "")
        if not numero:
            raise HTTPException(400, "El inquilino no tiene WhatsApp o teléfono registrado")

        base_url = str(request.base_url).rstrip("/")
        link_recibo = f"{base_url}/pagos/{pago.id}/recibo.pdf"

        mensaje = (
            f"{nombre_negocio}\n\n"
            f"Hola {inq.nombre}, se registró su pago correctamente.\n\n"
            f"Unidad: {prop.tipo if prop else '-'} {prop.numero if prop else '-'}\n"
            f"Monto pagado: {moneda} {float(pago.monto or 0):.2f}\n"
            f"Fecha: {pago.fecha_pago}\n"
            f"Método: {pago.metodo}\n"
            f"Recibo: {link_recibo}\n\n"
            f"{mensaje_recibo}"
        )

        url = f"https://wa.me/{numero}?text={quote(mensaje)}"
        return RedirectResponse(url=url, status_code=302)

    finally:
        db.close()
@app.get("/contratos/{contrato_id}/whatsapp_recordatorio")
def whatsapp_recordatorio_deuda(contrato_id: int):
    db = SessionLocal()
    try:
        hoy = date.today()

        contrato = db.query(Contrato).options(
            joinedload(Contrato.inquilino),
            joinedload(Contrato.propiedad)
        ).filter(Contrato.id == contrato_id).first()

        if not contrato:
            raise HTTPException(404, "Contrato no encontrado")

        inq = contrato.inquilino
        prop = contrato.propiedad

        if not inq:
            raise HTTPException(400, "El contrato no tiene inquilino asociado")

        telefono = (inq.whatsapp or inq.telefono or "").strip()
        telefono = ''.join(ch for ch in telefono if ch.isdigit())

        if not telefono:
            raise HTTPException(400, "El inquilino no tiene número de WhatsApp registrado")

        if not telefono.startswith("51"):
            telefono = "51" + telefono

        cargos = db.query(Cargo).filter(
            Cargo.contrato_id == contrato_id,
            Cargo.estado.in_(["Pendiente", "Parcial"])
        ).order_by(Cargo.vencimiento.asc(), Cargo.id.asc()).all()

        deuda_alquiler = 0.0
        deuda_agua = 0.0
        deuda_luz = 0.0
        deuda_otros = 0.0

        tiene_moroso = False
        tiene_parcial = False
        vence_hoy = False
        proximo_vencimiento = None

        for c in cargos:
            saldo = round(float(c.monto or 0) - float(c.pagado_acumulado or 0), 2)
            if saldo <= 0:
                continue

            if c.vencimiento:
                if proximo_vencimiento is None or c.vencimiento < proximo_vencimiento:
                    proximo_vencimiento = c.vencimiento
                if c.vencimiento < hoy:
                    tiene_moroso = True
                if c.vencimiento == hoy:
                    vence_hoy = True

            if c.estado == "Parcial":
                tiene_parcial = True

            concepto = (c.concepto or "").upper()

            if concepto in ["ALQUILER_PRORRATA", "ALQUILER_MENSUAL"] or concepto.startswith("ALQUILER_DIARIO") or concepto.startswith("ALQUILER_NOCHE"):
                deuda_alquiler += saldo
            elif c.concepto == "Agua":
                deuda_agua += saldo
            elif c.concepto == "Luz":
                deuda_luz += saldo
            else:
                deuda_otros += saldo

        deuda_total = round(deuda_alquiler + deuda_agua + deuda_luz + deuda_otros, 2)

        if deuda_total <= 0:
            raise HTTPException(400, "Este contrato no tiene deuda pendiente")

        detalle = []
        if deuda_alquiler > 0:
            detalle.append(f"- Alquiler: S/ {deuda_alquiler:.2f}")
        if deuda_agua > 0:
            detalle.append(f"- Agua: S/ {deuda_agua:.2f}")
        if deuda_luz > 0:
            detalle.append(f"- Luz: S/ {deuda_luz:.2f}")
        if deuda_otros > 0:
            detalle.append(f"- Otros: S/ {deuda_otros:.2f}")

        detalle_txt = "\n".join(detalle)

        if tiene_moroso:
            tipo_mensaje = "Moroso"
            mensaje = (
                f"Estimado(a) {inq.nombre}, le escribimos de CREDIMAS.\n\n"
                f"La unidad {prop.numero} registra una deuda vencida de S/ {deuda_total:.2f}.\n\n"
                f"Detalle pendiente:\n{detalle_txt}\n\n"
                f"Le solicitamos regularizar el pago a la brevedad para evitar inconvenientes.\n"
                f"Gracias."
            )
        elif tiene_parcial:
            tipo_mensaje = "Pago parcial"
            mensaje = (
                f"Hola {inq.nombre}, le escribimos de CREDIMAS.\n\n"
                f"Registramos un pago parcial de la unidad {prop.numero}, pero aún mantiene un saldo pendiente de S/ {deuda_total:.2f}.\n\n"
                f"Detalle:\n{detalle_txt}\n\n"
                f"Le agradeceremos completar el pago.\n"
                f"Gracias."
            )
        elif vence_hoy:
            tipo_mensaje = "Vence hoy"
            mensaje = (
                f"Hola {inq.nombre}, le escribimos de CREDIMAS.\n\n"
                f"Le recordamos que hoy vence el pago pendiente de la unidad {prop.numero} por un total de S/ {deuda_total:.2f}.\n\n"
                f"Detalle:\n{detalle_txt}\n\n"
                f"Le agradeceremos regularizarlo el día de hoy.\n"
                f"Muchas gracias."
            )
        else:
            tipo_mensaje = "Pendiente"
            fecha_txt = str(proximo_vencimiento) if proximo_vencimiento else "próximamente"
            mensaje = (
                f"Hola {inq.nombre}, le escribimos de CREDIMAS.\n\n"
                f"La unidad {prop.numero} mantiene un saldo pendiente de S/ {deuda_total:.2f} con vencimiento {fecha_txt}.\n\n"
                f"Detalle:\n{detalle_txt}\n\n"
                f"Le enviamos este recordatorio para su regularización.\n"
                f"Gracias."
            )

        bitacora = BitacoraCobranza(
            contrato_id=contrato.id,
            inquilino_id=inq.id,
            fecha_envio=datetime.now(),
            tipo_mensaje=tipo_mensaje,
            deuda_total=deuda_total,
            telefono=telefono,
            detalle=mensaje
        )
        db.add(bitacora)
        db.commit()

        url = f"https://wa.me/{telefono}?text={quote(mensaje)}"
        return RedirectResponse(url=url, status_code=302)

    finally:
        db.close()
@app.get("/cobranza/bitacora", response_class=HTMLResponse)
@login_required
@role_required("Administrador")
def cobranza_bitacora(
    request: Request,
    fecha: str = "",
):
    db = SessionLocal()
    try:
        query = db.query(BitacoraCobranza).options(
            joinedload(BitacoraCobranza.contrato).joinedload(Contrato.propiedad),
            joinedload(BitacoraCobranza.inquilino)
        )

        if fecha:
            try:
                fecha_obj = date.fromisoformat(fecha)
                query = query.filter(func.date(BitacoraCobranza.fecha_envio) == fecha_obj)
            except:
                pass

        items = query.order_by(
            BitacoraCobranza.fecha_envio.desc(),
            BitacoraCobranza.id.desc()
        ).all()

        return templates.TemplateResponse("cobranza_bitacora.html", {
            "request": request,
            "items": items,
            "fecha": fecha
        })
    finally:
        db.close()
@app.get("/contratos/{contrato_id}/detalle", response_class=HTMLResponse)
def contrato_detalle(request: Request, contrato_id: int):
    db = SessionLocal()
    try:
        contrato = db.query(Contrato).options(
            joinedload(Contrato.inquilino),
            joinedload(Contrato.propiedad)
        ).filter(Contrato.id == contrato_id).first()

        if not contrato:
            raise HTTPException(404, "Contrato no encontrado")

        cargos_pendientes = db.query(Cargo).filter(
            Cargo.contrato_id == contrato_id,
            Cargo.estado.in_(["Pendiente", "Parcial"])
        ).order_by(Cargo.vencimiento.asc(), Cargo.id.asc()).all()

        deuda_total = 0.0
        for c in cargos_pendientes:
            deuda_total += round(float(c.monto or 0) - float(c.pagado_acumulado or 0), 2)

        pagos_recientes = db.query(Pago).filter(
            Pago.contrato_id == contrato_id
        ).order_by(Pago.fecha_pago.desc(), Pago.id.desc()).limit(10).all()

        lecturas = db.query(Lectura).filter(
            Lectura.contrato_id == contrato_id
        ).order_by(Lectura.fecha_registro.desc(), Lectura.id.desc()).limit(10).all()

        return templates.TemplateResponse("contrato_detalle.html", {
            "request": request,
            "contrato": contrato,
            "cargos_pendientes": cargos_pendientes,
            "deuda_total": round(float(deuda_total), 2),
            "pagos_recientes": pagos_recientes,
            "lecturas": lecturas,
        })
    finally:
        db.close()
@app.get("/configuracion", response_class=HTMLResponse)
@login_required
def configuracion_empresa(request: Request):
    db = SessionLocal()
    try:
        config = db.query(ConfiguracionEmpresa).first()

        if not config:
            config = ConfiguracionEmpresa(
                nombre_comercial="CREDIMAS",
                razon_social="",
                ruc="",
                direccion="",
                telefono="",
                whatsapp="",
                correo="",
                ciudad="",
                moneda="S/",
                mensaje_recibo="Gracias por su pago.",
                pie_pagina="Documento generado por CREDIMAS.",
                logo_url="",
                dia_corte=25
            )
            db.add(config)
            db.commit()
            db.refresh(config)

        return templates.TemplateResponse("configuracion.html", {
            "request": request,
            "config": config
        })
    finally:
        db.close()


@app.post("/configuracion/guardar")
@login_required
def guardar_configuracion_empresa(
    request: Request,
    nombre_comercial: str = Form(""),
    razon_social: str = Form(""),
    ruc: str = Form(""),
    direccion: str = Form(""),
    telefono: str = Form(""),
    whatsapp: str = Form(""),
    correo: str = Form(""),
    ciudad: str = Form(""),
    moneda: str = Form("S/"),
    mensaje_recibo: str = Form(""),
    pie_pagina: str = Form(""),
    logo_url: str = Form(""),
    dia_corte: int = Form(25),
):
    db = SessionLocal()
    try:
        print("=== GUARDAR CONFIGURACION ===")
        print("nombre_comercial:", nombre_comercial)
        print("razon_social:", razon_social)
        print("ruc:", ruc)
        print("direccion:", direccion)
        print("telefono:", telefono)
        print("whatsapp:", whatsapp)
        print("correo:", correo)
        print("ciudad:", ciudad)
        print("moneda:", moneda)
        print("mensaje_recibo:", mensaje_recibo)
        print("pie_pagina:", pie_pagina)
        print("logo_url:", logo_url)
        print("dia_corte:", dia_corte)

        if int(dia_corte) < 1 or int(dia_corte) > 31:
            raise HTTPException(400, "El día de corte debe estar entre 1 y 31")

        config = db.query(ConfiguracionEmpresa).first()

        if not config:
            config = ConfiguracionEmpresa()
            db.add(config)
            db.flush()

        config.nombre_comercial = (nombre_comercial or "").strip()
        config.razon_social = (razon_social or "").strip()
        config.ruc = (ruc or "").strip()
        config.direccion = (direccion or "").strip()
        config.telefono = (telefono or "").strip()
        config.whatsapp = (whatsapp or "").strip()
        config.correo = (correo or "").strip()
        config.ciudad = (ciudad or "").strip()
        config.moneda = (moneda or "S/").strip()
        config.mensaje_recibo = (mensaje_recibo or "").strip()
        config.pie_pagina = (pie_pagina or "").strip()
        config.logo_url = (logo_url or "").strip()
        config.dia_corte = int(dia_corte)

        db.commit()
        return RedirectResponse(url="/configuracion", status_code=303)

    except Exception as e:
        db.rollback()
        import traceback
        print("=== ERROR GUARDAR CONFIGURACION ===")
        print(type(e).__name__, str(e))
        traceback.print_exc()
        raise

    finally:
        db.close()
def generar_contrato_pdf_pro(contrato: Contrato, config) -> bytes:
    nombre_negocio = config.nombre_comercial or "CREDIMAS"
    razon_social = config.razon_social or ""
    ruc = config.ruc or ""
    direccion = config.direccion or ""
    telefono = config.telefono or ""
    whatsapp_empresa = config.whatsapp or ""
    moneda = config.moneda or "S/"
    pie_pagina = config.pie_pagina or "Documento generado por CREDIMAS."

    inq = contrato.inquilino
    prop = contrato.propiedad

    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    w, h = A4
    M = 16 * mm
    y = h - M

    # CABECERA
    c.setFillColor(colors.HexColor("#111827"))
    c.rect(M, y - 30*mm, w - 2*M, 30*mm, fill=1, stroke=0)

    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 16)
    c.drawString(M + 5*mm, y - 10*mm, nombre_negocio)

    c.setFont("Helvetica-Bold", 11)
    c.drawString(M + 5*mm, y - 17*mm, "CONTRATO DE ALQUILER")

    info_y = y - 23*mm
    c.setFont("Helvetica", 8)

    if razon_social:
        c.drawString(M + 5*mm, info_y, razon_social)
        info_y -= 4*mm

    if ruc:
        c.drawString(M + 5*mm, info_y, f"RUC: {ruc}")
        info_y -= 4*mm

    if direccion:
        c.drawString(M + 5*mm, info_y, direccion)
        info_y -= 4*mm

    contacto = ""
    if telefono:
        contacto += f"Tel: {telefono}"
    if whatsapp_empresa:
        if contacto:
            contacto += " | "
        contacto += f"WhatsApp: {whatsapp_empresa}"

    if contacto:
        c.drawString(M + 5*mm, info_y, contacto)

    c.setFont("Helvetica-Bold", 11)
    c.drawRightString(w - M - 5*mm, y - 12*mm, f"Contrato #{contrato.id}")
    c.setFont("Helvetica", 8)
    c.drawRightString(w - M - 5*mm, y - 18*mm, f"Fecha: {date.today().isoformat()}")

    c.setFillColor(colors.black)
    y -= 38*mm

    # TITULO
    c.setFont("Helvetica-Bold", 12)
    c.drawString(M, y, "DATOS DEL CONTRATO")
    y -= 8*mm

    c.setFont("Helvetica", 10)
    lineas = [
        f"Inquilino: {inq.nombre if inq else '-'}",
        f"DNI: {inq.dni if inq else '-'}",
        f"Teléfono: {inq.telefono if inq and inq.telefono else '-'}",
        f"WhatsApp: {inq.whatsapp if inq and inq.whatsapp else '-'}",
        f"Propiedad: {prop.tipo if prop else '-'} {prop.numero if prop else '-'}",
        f"Precio base: {moneda} {float(prop.precio or 0):.2f}" if prop else f"Precio base: {moneda} 0.00",
        f"Tipo de alquiler: {contrato.tipo_alquiler or '-'}",
        f"Monto mensual: {moneda} {float(contrato.monto_mensual or 0):.2f}",
        f"Fecha inicio: {contrato.fecha_inicio}",
        f"Fecha fin: {contrato.fecha_fin}",
        f"Estado: {contrato.estado or '-'}",
    ]

    for linea in lineas:
        c.drawString(M, y, linea)
        y -= 6*mm

    y -= 4*mm
    c.setStrokeColor(colors.HexColor("#D1D5DB"))
    c.line(M, y, w - M, y)
    y -= 8*mm

    # TEXTO CONTRACTUAL
    c.setFont("Helvetica-Bold", 12)
    c.drawString(M, y, "CLÁUSULAS BÁSICAS")
    y -= 8*mm

    texto = [
        f"1. El arrendador da en alquiler la unidad {prop.tipo if prop else '-'} {prop.numero if prop else '-'} al inquilino antes identificado.",
        f"2. El monto pactado del alquiler es de {moneda} {float(contrato.monto_mensual or 0):.2f}.",
        f"3. El contrato rige desde {contrato.fecha_inicio} hasta {contrato.fecha_fin}.",
        "4. El inquilino se compromete a pagar puntualmente el alquiler y los consumos o cargos adicionales que correspondan.",
        "5. El incumplimiento de pago podrá generar cobranza, recargos administrativos o resolución del contrato según corresponda.",
        "6. Ambas partes declaran estar conformes con las condiciones registradas en este documento.",
    ]

    c.setFont("Helvetica", 10)
    for t in texto:
        partes = partir_texto_pdf(t, 95)
        for p in partes:
            if y < 35*mm:
                c.showPage()
                y = h - M
                c.setFont("Helvetica", 10)
            c.drawString(M, y, p)
            y -= 6*mm
        y -= 2*mm

    # FIRMAS
    y -= 8*mm
    if y < 45*mm:
        c.showPage()
        y = h - M

    c.setStrokeColor(colors.black)
    c.line(M + 10*mm, y, M + 70*mm, y)
    c.line(w - M - 70*mm, y, w - M - 10*mm, y)

    c.setFont("Helvetica", 9)
    c.drawCentredString(M + 40*mm, y - 5*mm, "Arrendador")
    c.drawCentredString(w - M - 40*mm, y - 5*mm, "Inquilino")

    # PIE
    c.setFont("Helvetica-Oblique", 8)
    c.setFillColor(colors.HexColor("#6B7280"))
    c.drawString(M, 12*mm, pie_pagina)

    c.setFillColor(colors.black)
    c.showPage()
    c.save()
    buffer.seek(0)
    return buffer.read()
def partir_texto_pdf(texto: str, max_chars: int = 95):
    palabras = texto.split()
    lineas = []
    actual = ""

    for palabra in palabras:
        prueba = f"{actual} {palabra}".strip()
        if len(prueba) <= max_chars:
            actual = prueba
        else:
            if actual:
                lineas.append(actual)
            actual = palabra

    if actual:
        lineas.append(actual)

    return lineas

@app.get("/contratos/{contrato_id}/contrato.pdf")
def descargar_contrato_pdf(contrato_id: int):
    db = SessionLocal()
    try:
        contrato = db.query(Contrato).options(
            joinedload(Contrato.inquilino),
            joinedload(Contrato.propiedad)
        ).filter(Contrato.id == contrato_id).first()

        if not contrato:
            raise HTTPException(404, "Contrato no encontrado")

        config = db.query(ConfiguracionEmpresa).first()

        if not config:
            config = ConfiguracionEmpresa(
                nombre_comercial="CREDIMAS",
                razon_social="",
                ruc="",
                direccion="",
                telefono="",
                whatsapp="",
                correo="",
                ciudad="",
                moneda="S/",
                mensaje_recibo="Gracias por su pago.",
                pie_pagina="Documento generado por CREDIMAS.",
                logo_url="",
                dia_corte=25
            )

        pdf_bytes = generar_contrato_pdf_pro(contrato, config)
        filename = f"contrato_{contrato.id}_{date.today().isoformat()}.pdf"

        return StreamingResponse(
            io.BytesIO(pdf_bytes),
            media_type="application/pdf",
            headers={"Content-Disposition": f'inline; filename="{filename}"'}
        )
    finally:
        db.close()
@app.get("/cobranza", response_class=HTMLResponse)
@login_required
@role_required("Administrador")
def cobranza_panel(request: Request):
    db = SessionLocal()
    try:
        hoy = date.today()

        contratos = db.query(Contrato).options(
            joinedload(Contrato.inquilino),
            joinedload(Contrato.propiedad)
        ).filter(
            Contrato.estado == "Activo"
        ).all()

        cobranza_items = []

        for contrato in contratos:
            inq = contrato.inquilino
            prop = contrato.propiedad

            cargos = db.query(Cargo).filter(
                Cargo.contrato_id == contrato.id,
                Cargo.estado.in_(["Pendiente", "Parcial"])
            ).order_by(Cargo.vencimiento.asc(), Cargo.id.asc()).all()

            if not cargos:
                continue

            deuda_alquiler = 0.0
            deuda_agua = 0.0
            deuda_luz = 0.0
            deuda_otros = 0.0

            tiene_moroso = False
            tiene_parcial = False
            vence_hoy = False
            proximo_vencimiento = None

            for c in cargos:
                saldo = round(float(c.monto or 0) - float(c.pagado_acumulado or 0), 2)
                if saldo <= 0:
                    continue

                if c.vencimiento:
                    if proximo_vencimiento is None or c.vencimiento < proximo_vencimiento:
                        proximo_vencimiento = c.vencimiento
                    if c.vencimiento < hoy:
                        tiene_moroso = True
                    if c.vencimiento == hoy:
                        vence_hoy = True

                if c.estado == "Parcial":
                    tiene_parcial = True

                concepto = (c.concepto or "").upper()
                if concepto in ["ALQUILER_PRORRATA", "ALQUILER_MENSUAL"] or concepto.startswith("ALQUILER_DIARIO") or concepto.startswith("ALQUILER_NOCHE"):
                    deuda_alquiler += saldo
                elif c.concepto == "Agua":
                    deuda_agua += saldo
                elif c.concepto == "Luz":
                    deuda_luz += saldo
                else:
                    deuda_otros += saldo

            deuda_total = round(deuda_alquiler + deuda_agua + deuda_luz + deuda_otros, 2)
            if deuda_total <= 0:
                continue

            if tiene_moroso:
                tipo_alerta = "Moroso"
            elif tiene_parcial:
                tipo_alerta = "Pago parcial"
            elif vence_hoy:
                tipo_alerta = "Vence hoy"
            else:
                tipo_alerta = "Pendiente"

            cobranza_items.append({
                "contrato_id": contrato.id,
                "inquilino_id": inq.id if inq else None,
                "inquilino": inq.nombre if inq else "-",
                "telefono": inq.whatsapp if inq and inq.whatsapp else (inq.telefono if inq else "-"),
                "propiedad": f"{prop.tipo} {prop.numero}" if prop else "-",
                "propiedad_numero": prop.numero if prop else "-",
                "propiedad_tipo": prop.tipo if prop else "-",
                "tipo_alerta": tipo_alerta,
                "deuda_total": round(deuda_total, 2),
                "deuda_alquiler": round(deuda_alquiler, 2),
                "deuda_agua": round(deuda_agua, 2),
                "deuda_luz": round(deuda_luz, 2),
                "deuda_otros": round(deuda_otros, 2),
                "proximo_vencimiento": proximo_vencimiento,
            })

        total_casos = len(cobranza_items)
        total_morosos = sum(1 for x in cobranza_items if x["tipo_alerta"] == "Moroso")
        total_vence_hoy = sum(1 for x in cobranza_items if x["tipo_alerta"] == "Vence hoy")
        deuda_total_general = round(sum(x["deuda_total"] for x in cobranza_items), 2)

        return templates.TemplateResponse("cobranza.html", {
            "request": request,
            "items": cobranza_items,
            "total_casos": total_casos,
            "total_morosos": total_morosos,
            "total_vence_hoy": total_vence_hoy,
            "deuda_total_general": deuda_total_general,
        })
    finally:
        db.close()
@app.get("/gastos", response_class=HTMLResponse)
@login_required
def gastos(
    request: Request,
    desde: str = "",
    hasta: str = "",
    categoria: str = "",
):
    db = SessionLocal()
    try:
        query = db.query(Gasto).options(
            joinedload(Gasto.propiedad)
        )

        if desde:
            try:
                query = query.filter(Gasto.fecha >= date.fromisoformat(desde))
            except:
                pass

        if hasta:
            try:
                query = query.filter(Gasto.fecha <= date.fromisoformat(hasta))
            except:
                pass

        if categoria:
            query = query.filter(Gasto.categoria == categoria)

        items = query.order_by(Gasto.fecha.desc(), Gasto.id.desc()).all()

        total_gastado = sum(float(x.monto or 0) for x in items)
        total_hoy = sum(float(x.monto or 0) for x in items if x.fecha == date.today())

        propiedades = db.query(Propiedad).order_by(Propiedad.numero.asc()).all()

        return templates.TemplateResponse("gastos.html", {
            "request": request,
            "items": items,
            "propiedades": propiedades,
            "desde": desde,
            "hasta": hasta,
            "categoria": categoria,
            "total_gastado": round(total_gastado, 2),
            "total_hoy": round(total_hoy, 2),
        })
    finally:
        db.close()


@app.post("/gastos/crear")
def crear_gasto(
    fecha: str = Form(...),
    categoria: str = Form(...),
    concepto: str = Form(...),
    monto: float = Form(...),
    metodo: str = Form(""),
    observacion: str = Form(""),
    propiedad_id: int = Form(0),
):
    db = SessionLocal()
    try:
        gasto = Gasto(
            fecha=date.fromisoformat(fecha),
            categoria=categoria,
            concepto=concepto,
            monto=float(monto or 0),
            metodo=metodo,
            observacion=observacion,
            propiedad_id=propiedad_id if propiedad_id and propiedad_id > 0 else None
        )
        db.add(gasto)
        db.commit()
        return RedirectResponse(url="/gastos", status_code=303)
    finally:
        db.close()


@app.get("/gastos/{gasto_id}/editar", response_class=HTMLResponse)
def editar_gasto(request: Request, gasto_id: int):
    db = SessionLocal()
    try:
        gasto = db.query(Gasto).filter(Gasto.id == gasto_id).first()
        if not gasto:
            raise HTTPException(404, "Gasto no encontrado")

        propiedades = db.query(Propiedad).order_by(Propiedad.numero.asc()).all()

        return templates.TemplateResponse("gasto_editar.html", {
            "request": request,
            "gasto": gasto,
            "propiedades": propiedades,
        })
    finally:
        db.close()


@app.post("/gastos/{gasto_id}/actualizar")
def actualizar_gasto(
    gasto_id: int,
    fecha: str = Form(...),
    categoria: str = Form(...),
    concepto: str = Form(...),
    monto: float = Form(...),
    metodo: str = Form(""),
    observacion: str = Form(""),
    propiedad_id: int = Form(0),
):
    db = SessionLocal()
    try:
        gasto = db.query(Gasto).filter(Gasto.id == gasto_id).first()
        if not gasto:
            raise HTTPException(404, "Gasto no encontrado")

        gasto.fecha = date.fromisoformat(fecha)
        gasto.categoria = categoria
        gasto.concepto = concepto
        gasto.monto = float(monto or 0)
        gasto.metodo = metodo
        gasto.observacion = observacion
        gasto.propiedad_id = propiedad_id if propiedad_id and propiedad_id > 0 else None

        db.commit()
        return RedirectResponse(url="/gastos", status_code=303)
    finally:
        db.close()


@app.post("/gastos/{gasto_id}/eliminar")
def eliminar_gasto(gasto_id: int):
    db = SessionLocal()
    try:
        gasto = db.query(Gasto).filter(Gasto.id == gasto_id).first()
        if not gasto:
            raise HTTPException(404, "Gasto no encontrado")

        db.delete(gasto)
        db.commit()
        return RedirectResponse(url="/gastos", status_code=303)
    finally:
        db.close()
# =========================================================
# USUARIOS
# =========================================================
@app.get("/usuarios", response_class=HTMLResponse)
@login_required
@role_required("Administrador")
def usuarios_listar(request: Request, q: str = ""):
    db = SessionLocal()
    try:
        query = db.query(User)

        if q and q.strip():
            texto = f"%{q.strip()}%"
            query = query.filter(
                (User.nombre.ilike(texto)) |
                (User.username.ilike(texto)) |
                (User.rol.ilike(texto)) |
                (User.estado.ilike(texto))
            )

        usuarios = query.order_by(User.id.desc()).all()

        return templates.TemplateResponse(
            "usuarios.html",
            {
                "request": request,
                "usuarios": usuarios,
                "q": q,
                "ok": request.query_params.get("ok"),
                "error": request.query_params.get("error"),
            }
        )
    finally:
        db.close()


@app.post("/usuarios/crear")
@login_required
@role_required("Administrador")
def usuarios_crear(
    request: Request,
    nombre: str = Form(...),
    username: str = Form(...),
    password: str = Form(...),
    rol: str = Form(...),
    estado: str = Form(...)
):
    db = SessionLocal()
    try:
        nombre = nombre.strip()
        username_limpio = username.strip()
        password = password.strip()
        rol = rol.strip()
        estado = estado.strip()

        if not nombre or len(nombre) < 3:
            return RedirectResponse(url="/usuarios?error=nombre_invalido", status_code=303)

        if not username_limpio or len(username_limpio) < 3:
            return RedirectResponse(url="/usuarios?error=username_invalido", status_code=303)

        if not password or len(password) < 4:
            return RedirectResponse(url="/usuarios?error=password_invalida", status_code=303)

        if rol not in ["Administrador", "Operador"]:
            return RedirectResponse(url="/usuarios?error=rol_invalido", status_code=303)

        if estado not in ["Activo", "Inactivo"]:
            return RedirectResponse(url="/usuarios?error=estado_invalido", status_code=303)

        existe = db.query(User).filter(User.username == username_limpio).first()
        if existe:
            return RedirectResponse(url="/usuarios?error=usuario_existe", status_code=303)

        nuevo = User(
            nombre=nombre,
            username=username_limpio,
            password_hash=hash_password(password),
            rol=rol,
            estado=estado,
            debe_cambiar_password=1
        )
        db.add(nuevo)
        db.commit()
        db.refresh(nuevo)

        registrar_auditoria(
            db,
            request,
            action="CREAR",
            module="USUARIOS",
            detail=f"Creó usuario {nuevo.username} con rol {nuevo.rol} y estado {nuevo.estado}"
        )
        db.commit()

        return RedirectResponse(url="/usuarios?ok=creado", status_code=303)
    finally:
        db.close()

@app.post("/usuarios/{user_id}/estado")
@login_required
@role_required("Administrador")
def usuarios_cambiar_estado(
    user_id: int,
    estado: str = Form(...),
    request: Request = None
):
    db = SessionLocal()
    try:
        usuario = db.query(User).filter(User.id == user_id).first()
        if not usuario:
            return RedirectResponse(url="/usuarios?error=no_encontrado", status_code=303)

        if request.session.get("user_id") == usuario.id and estado == "Inactivo":
            return RedirectResponse(url="/usuarios?error=no_puedes_desactivarte", status_code=303)

        estado_anterior = usuario.estado
        usuario.estado = estado.strip()
        db.commit()

        registrar_auditoria(
            db,
            request,
            action="CAMBIAR_ESTADO",
            module="USUARIOS",
            detail=f"Cambió estado de {usuario.username} de {estado_anterior} a {usuario.estado}"
        )
        db.commit()

        return RedirectResponse(url="/usuarios?ok=estado", status_code=303)
    finally:
        db.close()

@app.post("/usuarios/{user_id}/rol")
@login_required
@role_required("Administrador")
def usuarios_cambiar_rol(
    user_id: int,
    rol: str = Form(...),
    request: Request = None
):
    db = SessionLocal()
    try:
        usuario = db.query(User).filter(User.id == user_id).first()
        if not usuario:
            return RedirectResponse(url="/usuarios?error=no_encontrado", status_code=303)

        if request.session.get("user_id") == usuario.id and rol != "Administrador":
            total_admins_activos = db.query(User).filter(
                User.rol == "Administrador",
                User.estado == "Activo"
            ).count()

            if total_admins_activos <= 1:
                return RedirectResponse(url="/usuarios?error=ultimo_admin", status_code=303)

        rol_anterior = usuario.rol
        usuario.rol = rol.strip()
        db.commit()

        if request.session.get("user_id") == usuario.id:
            request.session["user_rol"] = usuario.rol

        registrar_auditoria(
            db,
            request,
            action="CAMBIAR_ROL",
            module="USUARIOS",
            detail=f"Cambió rol de {usuario.username} de {rol_anterior} a {usuario.rol}"
        )
        db.commit()

        return RedirectResponse(url="/usuarios?ok=rol", status_code=303)
    finally:
        db.close()


@app.post("/usuarios/{user_id}/reset-password")
@login_required
@role_required("Administrador")
def usuarios_reset_password(
    user_id: int,
    nueva_password: str = Form(...),
    request: Request = None
):
    db = SessionLocal()
    try:
        usuario = db.query(User).filter(User.id == user_id).first()
        if not usuario:
            return RedirectResponse(url="/usuarios?error=no_encontrado", status_code=303)

        nueva_password = nueva_password.strip()
        if len(nueva_password) < 4:
            return RedirectResponse(url="/usuarios?error=password_invalida", status_code=303)

        usuario.password_hash = hash_password(nueva_password)
        db.commit()

        registrar_auditoria(
            db,
            request,
            action="RESET_PASSWORD",
            module="USUARIOS",
            detail=f"Restableció contraseña del usuario {usuario.username}"
        )
        db.commit()

        return RedirectResponse(url="/usuarios?ok=clave", status_code=303)
    finally:
        db.close()
@app.get("/auditoria", response_class=HTMLResponse)
@login_required
@role_required("Administrador")
def auditoria_listar(request: Request):
    db = SessionLocal()
    try:
        logs = db.query(AuditLog).options(
            joinedload(AuditLog.user)
        ).order_by(AuditLog.id.desc()).limit(300).all()

        return templates.TemplateResponse(
            "auditoria.html",
            {
                "request": request,
                "logs": logs
            }
        )
    finally:
        db.close()
@app.get("/cambiar-password", response_class=HTMLResponse)
@login_required
def cambiar_password_form(request: Request):
    return templates.TemplateResponse(
        "cambiar_password.html",
        {"request": request}
    )


@app.post("/cambiar-password")
@login_required
def cambiar_password_guardar(
    request: Request,
    password: str = Form(...)
):
    db = SessionLocal()
    try:
        user_id = request.session.get("user_id")
        user = db.query(User).filter(User.id == user_id).first()

        if not user:
            return RedirectResponse("/login", status_code=303)

        if len(password) < 4:
            return templates.TemplateResponse(
                "cambiar_password.html",
                {"request": request, "error": "La contraseña debe tener mínimo 4 caracteres"}
            )

        user.password_hash = hash_password(password)
        user.debe_cambiar_password = 0

        db.commit()

        return RedirectResponse("/dashboard", status_code=303)

    finally:
        db.close()

import traceback
from fastapi.responses import JSONResponse
from sqlalchemy.exc import IntegrityError

@app.post("/inquilinos/ajax-crear")
@login_required
def crear_inquilino_ajax(
    request: Request,
    nombre: str = Form(...),
    dni: str = Form(...),
    telefono: str = Form(""),
    whatsapp: str = Form("")
):
    db = SessionLocal()
    try:
        nombre = nombre.strip()
        dni = dni.strip()
        telefono = telefono.strip()
        whatsapp = whatsapp.strip()

        if not nombre:
            return JSONResponse(
                status_code=400,
                content={"ok": False, "error": "El nombre es obligatorio"}
            )

        if not dni:
            return JSONResponse(
                status_code=400,
                content={"ok": False, "error": "El DNI es obligatorio"}
            )

        existe_dni = db.query(Inquilino).filter(Inquilino.dni == dni).first()
        if existe_dni:
            return JSONResponse(
                status_code=400,
                content={"ok": False, "error": "Ya existe un inquilino con ese DNI"}
            )

        nuevo = Inquilino(
            nombre=nombre,
            dni=dni,
            telefono=telefono,
            whatsapp=whatsapp,
            estado="Activo"
        )

        db.add(nuevo)
        db.commit()
        db.refresh(nuevo)

        return JSONResponse(
            status_code=200,
            content={
                "ok": True,
                "id": nuevo.id,
                "nombre": nuevo.nombre,
                "dni": nuevo.dni,
                "telefono": nuevo.telefono or ""
            }
        )

    except IntegrityError as e:
        db.rollback()
        print("ERROR INTEGRITY AJAX INQUILINO:")
        traceback.print_exc()
        return JSONResponse(
            status_code=400,
            content={"ok": False, "error": f"IntegrityError: {str(e)}"}
        )

    except Exception as e:
        db.rollback()
        print("ERROR GENERAL AJAX INQUILINO:")
        traceback.print_exc()
        return JSONResponse(
            status_code=500,
            content={"ok": False, "error": f"{type(e).__name__}: {str(e)}"}
        )

    finally:
        db.close()

import traceback
from datetime import datetime

from datetime import datetime, timedelta

@app.post("/dashboard/editar_contrato")
@login_required
@role_required("Administrador")
def dashboard_editar_contrato(
    request: Request,
    contrato_id: int = Form(...),
    tipo_alquiler: str = Form(...),
    monto_mensual: str = Form(...),
    fecha_inicio: str = Form(...),
    fecha_fin: str = Form(...)
):
    db = SessionLocal()
    try:
        contrato = db.query(Contrato).filter(Contrato.id == contrato_id).first()
        if not contrato:
            raise HTTPException(status_code=404, detail="Contrato no encontrado")

        tipo_alquiler = (tipo_alquiler or "mensual").strip().lower()
        if tipo_alquiler not in ["mensual", "diario", "noche"]:
            raise HTTPException(status_code=400, detail="Tipo de alquiler inválido")

        monto = float(monto_mensual or 0)
        if monto <= 0:
            raise HTTPException(status_code=400, detail="El monto debe ser mayor a 0")

        fecha_inicio_new = datetime.strptime(fecha_inicio, "%Y-%m-%d").date()
        fecha_fin_new = datetime.strptime(fecha_fin, "%Y-%m-%d").date()

        if fecha_fin_new < fecha_inicio_new:
            raise HTTPException(status_code=400, detail="Fechas inválidas")

        # valores anteriores
        fecha_inicio_old = contrato.fecha_inicio
        fecha_fin_old = contrato.fecha_fin
        tipo_old = (contrato.tipo_alquiler or "").strip().lower()
        monto_old = float(contrato.monto_mensual or 0)

        # actualizar contrato
        contrato.tipo_alquiler = tipo_alquiler
        contrato.monto_mensual = monto
        contrato.fecha_inicio = fecha_inicio_new
        contrato.fecha_fin = fecha_fin_new

        # =========================
        # 1) ACTUALIZAR CARGOS EXISTENTES PENDIENTES/PARCIALES
        # =========================
        cargos = db.query(Cargo).filter(
            Cargo.contrato_id == contrato.id,
            Cargo.estado.in_(["Pendiente", "Parcial"])
        ).all()

        for c in cargos:
            pagado = float(c.pagado_acumulado or 0)

            if tipo_alquiler == "mensual":
                # Solo reajustamos cargos mensuales/prorrata pendientes
                concepto = (c.concepto or "").upper()

                if concepto in ["ALQUILER_MENSUAL", "ALQUILER_MENSUAL_AMPLIACION"]:
                    c.monto = monto

                elif concepto == "ALQUILER_PRORRATA":
                    _, monto_prorrata, venc = calcular_prorrata_mensual(monto, fecha_inicio_new)
                    c.monto = float(monto_prorrata)
                    c.vencimiento = venc
                    c.periodo = f"{venc.year:04d}-{venc.month:02d}"

            elif tipo_alquiler in ["diario", "noche"]:
                concepto = (c.concepto or "").upper()

                if concepto.startswith("ALQUILER_DIARIO") or concepto.startswith("ALQUILER_NOCHE"):
                    cantidad, total = calcular_total_diario_noche(
                        fecha_inicio_new,
                        fecha_fin_new,
                        monto,
                        tipo_alquiler
                    )
                    c.monto = total
                    c.vencimiento = fecha_fin_new
                    c.periodo = f"{fecha_inicio_new}_{fecha_fin_new}"

                    if tipo_alquiler == "diario":
                        c.concepto = f"ALQUILER_DIARIO ({cantidad} {'día' if cantidad == 1 else 'días'})"
                    else:
                        c.concepto = f"ALQUILER_NOCHE ({cantidad} {'noche' if cantidad == 1 else 'noches'})"

            # recalcular estado
            if pagado <= 0:
                c.estado = "Pendiente"
            elif pagado < c.monto:
                c.estado = "Parcial"
            else:
                c.estado = "Pagado"

        # =========================
        # 2) GENERAR CARGO POR AMPLIACIÓN
        # =========================
        if fecha_fin_old and fecha_fin_new > fecha_fin_old:

            # ---------- NOCHE ----------
            if tipo_alquiler == "noche":
                fecha_inicio_extra = fecha_fin_old
                fecha_fin_extra = fecha_fin_new

                if fecha_fin_extra > fecha_inicio_extra:
                    cantidad, total = calcular_total_diario_noche(
                        fecha_inicio_extra,
                        fecha_fin_extra,
                        monto,
                        "noche"
                    )

                    periodo_extra = f"{fecha_inicio_extra}_{fecha_fin_extra}"

                    existe = db.query(Cargo).filter(
                        Cargo.contrato_id == contrato.id,
                        Cargo.periodo == periodo_extra,
                        Cargo.concepto.like("ALQUILER_NOCHE_AMPLIACION%")
                    ).first()

                    if not existe:
                        db.add(Cargo(
                            contrato_id=contrato.id,
                            concepto=f"ALQUILER_NOCHE_AMPLIACION ({cantidad} {'noche' if cantidad == 1 else 'noches'})",
                            periodo=periodo_extra,
                            monto=total,
                            vencimiento=fecha_fin_extra,
                            estado="Pendiente",
                            pagado_acumulado=0.0
                        ))

            # ---------- DIARIO ----------
            elif tipo_alquiler == "diario":
                fecha_inicio_extra = fecha_fin_old + timedelta(days=1)
                fecha_fin_extra = fecha_fin_new

                if fecha_fin_extra >= fecha_inicio_extra:
                    cantidad, total = calcular_total_diario_noche(
                        fecha_inicio_extra,
                        fecha_fin_extra,
                        monto,
                        "diario"
                    )

                    periodo_extra = f"{fecha_inicio_extra}_{fecha_fin_extra}"

                    existe = db.query(Cargo).filter(
                        Cargo.contrato_id == contrato.id,
                        Cargo.periodo == periodo_extra,
                        Cargo.concepto.like("ALQUILER_DIARIO_AMPLIACION%")
                    ).first()

                    if not existe:
                        db.add(Cargo(
                            contrato_id=contrato.id,
                            concepto=f"ALQUILER_DIARIO_AMPLIACION ({cantidad} {'día' if cantidad == 1 else 'días'})",
                            periodo=periodo_extra,
                            monto=total,
                            vencimiento=fecha_fin_extra,
                            estado="Pendiente",
                            pagado_acumulado=0.0
                        ))

            # ---------- MENSUAL ----------
            elif tipo_alquiler == "mensual":
                fecha_cursor = fecha_fin_old

                while fecha_cursor < fecha_fin_new:
                    siguiente_mes = (fecha_cursor.replace(day=1) + timedelta(days=32)).replace(day=1)
                    periodo = f"{siguiente_mes.year}-{str(siguiente_mes.month).zfill(2)}"

                    existe = db.query(Cargo).filter(
                        Cargo.contrato_id == contrato.id,
                        Cargo.periodo == periodo,
                        Cargo.concepto == "ALQUILER_MENSUAL_AMPLIACION"
                    ).first()

                    if not existe:
                        db.add(Cargo(
                            contrato_id=contrato.id,
                            concepto="ALQUILER_MENSUAL_AMPLIACION",
                            periodo=periodo,
                            monto=monto,
                            vencimiento=siguiente_mes,
                            estado="Pendiente",
                            pagado_acumulado=0.0
                        ))

                    fecha_cursor = siguiente_mes

        registrar_auditoria(
            db,
            request,
            action="EDITAR_CONTRATO",
            module="CONTRATOS",
            detail=(
                f"Contrato {contrato.id} actualizado | "
                f"tipo: {tipo_old} -> {tipo_alquiler}, "
                f"monto: {monto_old:.2f} -> {monto:.2f}, "
                f"inicio: {fecha_inicio_old} -> {fecha_inicio_new}, "
                f"fin: {fecha_fin_old} -> {fecha_fin_new}"
            )
        )

        db.commit()
        return RedirectResponse("/dashboard?ok=1", status_code=303)

    except Exception:
        db.rollback()
        import traceback
        traceback.print_exc()
        raise

    finally:
        db.close()
@app.get("/contratos/{contrato_id}/whatsapp_ampliacion")
def whatsapp_ampliacion_contrato(contrato_id: int):
    db = SessionLocal()
    try:
        contrato = db.query(Contrato).options(
            joinedload(Contrato.inquilino),
            joinedload(Contrato.propiedad)
        ).filter(Contrato.id == contrato_id).first()

        if not contrato:
            raise HTTPException(404, "Contrato no encontrado")

        inq = contrato.inquilino
        prop = contrato.propiedad

        if not inq:
            raise HTTPException(400, "El contrato no tiene inquilino asociado")

        telefono = normalizar_numero_pe(inq.whatsapp or inq.telefono or "")
        if not telefono:
            raise HTTPException(400, "El inquilino no tiene número de WhatsApp registrado")

        cargos_ampliacion = db.query(Cargo).filter(
            Cargo.contrato_id == contrato_id,
            Cargo.estado.in_(["Pendiente", "Parcial"]),
            (
                Cargo.concepto.like("ALQUILER_DIARIO_AMPLIACION%") |
                Cargo.concepto.like("ALQUILER_NOCHE_AMPLIACION%") |
                (Cargo.concepto == "ALQUILER_MENSUAL_AMPLIACION")
            )
        ).order_by(Cargo.id.desc()).all()

        if not cargos_ampliacion:
            raise HTTPException(400, "No hay ampliaciones pendientes para este contrato")

        total = 0.0
        detalle = []

        for c in cargos_ampliacion:
            saldo = round(float(c.monto or 0) - float(c.pagado_acumulado or 0), 2)
            if saldo <= 0:
                continue

            total += saldo
            detalle.append(f"- {c.concepto}: S/ {saldo:.2f}")

        if total <= 0:
            raise HTTPException(400, "No hay saldo pendiente por ampliación")

        detalle_txt = "\n".join(detalle)

        mensaje = (
            f"Hola {inq.nombre}, le escribimos de CREDIMAS.\n\n"
            f"Se registró una ampliación de su contrato/estadía en la unidad {prop.numero if prop else '-'}.\n\n"
            f"Detalle pendiente:\n{detalle_txt}\n\n"
            f"Total pendiente: S/ {total:.2f}\n\n"
            f"Por favor regularizar el pago. Gracias."
        )

        url = f"https://wa.me/{telefono}?text={quote(mensaje)}"
        return RedirectResponse(url=url, status_code=302)

    finally:
        db.close()