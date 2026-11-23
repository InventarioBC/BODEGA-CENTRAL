# 📦 Bodega Central — Control de Abastecimiento, Bultos e Inventario
**Departamento Eléctrico | Guzmán y Larraín**

Aplicación web progresiva (PWA) y móvil para la consulta de existencias, control visual de bultos/pallets con códigos QR, registro de fotografías y preparación de despachos para obras.

---

## 🚀 Acceso en Internet (GitHub Pages)
Una vez activado GitHub Pages en este repositorio, la aplicación estará disponible 24/7 en:
👉 `https://<tu-usuario>.github.io/<nombre-repositorio>/`

* ✅ **Sin servidores locales:** Carga instantánea desde cualquier celular, tablet o PC.
* ✅ **Instalable:** Compatible con pantalla de inicio en Android (estilo Play Store) e iPhone (Safari).
* ✅ **Bultos y Pallets:** Visualización de fotografías de bultos, racks y contenido de materiales.
* ✅ **Descarga App Android:** Acceso directo al instalador `BodegaCentral.apk`.

---

## 📱 Descarga de la Aplicación Android (.APK)
Puedes descargar directamente el archivo instalador oficial para teléfonos y tablets Android:
* [Descargar BodegaCentral.apk](BodegaCentral.apk)

---

## 🛠️ Tecnologías Utilizadas
* **Frontend:** HTML5, Bootstrap 5, Bootstrap Icons, JavaScript Vanilla (sin dependencias pesadas).
* **PWA:** Service Worker con soporte offline y manifiesto web.
* **Backend Local:** Python 3 (servidor HTTP multiproceso, OCR Tesseract, integración con planillas Excel `Abastecimiento.xlsm`).
* **Android:** Java / Android SDK WebView Wrapper con acceso a cámara nativa y selector de archivos.
