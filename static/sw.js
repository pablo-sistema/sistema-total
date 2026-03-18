self.addEventListener("install", (event) => {
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  console.log("Service Worker activado");
});

self.addEventListener("fetch", (event) => {
});