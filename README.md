# 📧 Envío Masivo de Correos — Google Sheets + Apps Script

Herramienta para enviar correos personalizados de forma masiva (*mail merge*) usando una hoja de **Google Sheets** como lista de destinatarios y un **borrador de Gmail** como plantilla. Cada destinatario recibe el mensaje con sus propios datos (nombre, dependencia, etc.), con soporte para **CC** y **CCO**.

---

## ✨ Características

- Envío personalizado por fila mediante marcadores tipo `{{Nombre}}`.
- La plantilla se redacta como **borrador en Gmail**, conservando formato, imágenes en línea y adjuntos.
- Columnas opcionales de **CC** (copia) y **CCO** (copia oculta), con uno o varios correos por celda.
- Marca cada fila como `ENVIADO` con fecha y hora, lo que permite **reanudar sin duplicar**.
- Verifica la **cuota diaria** de Gmail y se detiene antes de superarla.
- Menú propio dentro de la hoja, con opción de **correo de prueba**.
- Sin dependencias externas ni llaves de API: usa la autorización nativa de Google.

---

## 📁 Contenido del repositorio

| Archivo | Descripción |
|---|---|
| `EnvioMasivo_CC_CCO.gs` | Código de Google Apps Script (versión con CC y CCO). |
| `Contactos_EnvioMasivo_CC_CCO.xlsx` | Hoja plantilla con encabezados y ejemplos ficticios. |
| `Manual_Envio_Masivo_Correos.docx` | Manual de uso y configuración para usuarios finales. |

---

## ✅ Requisitos

- Cuenta de Google (Gmail o Google Workspace).
- Acceso desde computadora con navegador (la configuración no se puede hacer desde el celular).
- La hoja debe estar en **formato nativo de Google Sheets** (no en modo `.XLSX`).

---

## 🚀 Instalación y configuración

1. **Prepara la hoja.** Sube `Contactos_EnvioMasivo_CC_CCO.xlsx` a Google Drive y ábrela con Google Sheets. Si aparece la etiqueta `.XLSX` junto al nombre, ve a **Archivo → Guardar como Hojas de cálculo de Google** y trabaja sobre la copia.
2. **Redacta el borrador.** En Gmail, crea un borrador usando marcadores como `{{Nombre}}` o `{{Dependencia}}`. Guárdalo sin enviar.
3. **Instala el código.** En la hoja: **Extensiones → Apps Script**. Borra el contenido y pega `EnvioMasivo_CC_CCO.gs`. Guarda (`Ctrl + S`).
4. **Configura.** En el objeto `CONFIG`, ajusta `ASUNTO_BORRADOR` con el asunto **exacto** de tu borrador. Opcionalmente, `NOMBRE_REMITENTE`, `RESPONDER_A` y `CORREO_PRUEBA`.
5. **Recarga la hoja** (`F5`). Aparecerá el menú **📧 Envío masivo**.
6. **Autoriza** los permisos de Gmail la primera vez.

---

## 📨 Uso

1. Llena la hoja `Contactos` con tus destinatarios.
2. Menú **📧 Envío masivo → Enviar correo de prueba** para revisar el formato.
3. Menú **📧 Envío masivo → Enviar correos pendientes** para el envío real.

### Estructura de columnas

| Columna | Uso | Obligatoria |
|---|---|---|
| `Correo` | Destinatario principal | Sí |
| `Nombre`, `Dependencia`, … | Campos de personalización (`{{Nombre}}`, etc.) | Recomendadas |
| `CC` | Copia (uno o varios correos separados por coma) | Opcional |
| `CCO` | Copia oculta (uno o varios correos separados por coma) | Opcional |
| `Estado` | La llena el sistema (`ENVIADO` / `ERROR`) | — |
| `FechaEnvio` | La llena el sistema | — |

---

## 🎛️ Personalización

- **Asunto por fila:** agrega una columna `Asunto`, pon `{{Asunto}}` como asunto del borrador y deja `ASUNTO_BORRADOR: '{{Asunto}}'` en el código.
- **Más campos:** crea cualquier columna nueva y úsala como `{{NombreDeLaColumna}}`; no requiere tocar el código.
- **CC visible vs. CCO oculto:** los destinatarios en `CC` se ven entre sí; los de `CCO` quedan ocultos.

---

## 📊 Límites de envío

| Tipo de cuenta | Cuota diaria aproximada |
|---|---|
| Gmail gratuito | ~100 correos/día |
| Google Workspace | ~1,500 correos/día |

> Cada destinatario cuenta, **incluidos los de CC y CCO**. El script se detiene antes de superar la cuota y permite reanudar al día siguiente sin duplicar.

---

## 🛠️ Solución de problemas

| Síntoma | Causa / Solución |
|---|---|
| No aparece el menú **Extensiones** | La hoja está en modo `.XLSX`. Conviértela a Google Sheets. |
| Error `Cannot call SpreadsheetApp.getUi() from this context` | Ejecutaste el código a mano con ▶. Recarga la hoja y usa el menú **📧 Envío masivo**. |
| No aparece el menú **📧 Envío masivo** | Recarga con `F5`; verifica que el código esté guardado y ligado a la hoja. |
| No encuentra el borrador | `ASUNTO_BORRADOR` debe coincidir exactamente con el asunto del borrador. |
| Un `{{campo}}` aparece literal | El nombre no coincide con un encabezado (revisa mayúsculas y acentos). |

---

## 🔒 Privacidad y datos personales

Este repositorio contiene únicamente **datos de ejemplo ficticios**. **No subas hojas con correos o datos reales** de personas a un repositorio público: pueden constituir datos personales sujetos a protección. Para versionar listas reales, usa un repositorio **privado**. El archivo `.gitignore` incluido excluye por defecto los archivos con datos reales.

---

## 👤 Autor

Desarrollado por [@vincentjbs](https://github.com/vincentjbs).

## 📄 Licencia

Distribuido bajo licencia MIT. Puedes usarlo, modificarlo y compartirlo libremente.
