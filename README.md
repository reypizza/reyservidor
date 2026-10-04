# Servidor de Rey Pizza · CAIX, S.A.

Este es el servidor propio de la app. **Hoy hace tres cosas** (Fase 0) y las demás se van sumando módulo por módulo (ver `PLAN.md`):

1. **Puerta única.** La app le habla a este servidor. Lo que todavía no está migrado, él se lo pasa a Apps Script (igual que hacía el teléfono) y devuelve la respuesta tal cual. Nadie tiene que cambiar nada.
2. **Mantiene despierto a Apps Script.** Cada 4 minutos le toca la puerta, para que la primera consulta del día no tarde 7 segundos.
3. **Mide la velocidad real** de cada consulta (sin guardar PINs ni datos) y la muestra en `/velocidad`.

Si este servidor se cae, **la app sigue funcionando** con Apps Script (respaldo automático).

---

## Qué necesita (una sola vez)

- Una cuenta en **GitHub** (github.com, gratis).
- Una cuenta en **Render** (render.com). Pide tarjeta; antes de crear nada le muestra el precio mensual.
- La dirección `/exec` de su Apps Script (la que está en `config.js`: `https://script.google.com/macros/s/…/exec`).

## Paso a paso

**1. Subir el proyecto a GitHub** (se puede hacer desde una tablet o el teléfono)
1. En el celular o tablet, abra el zip **rey-servidor-github.zip** con la app **Archivos** y toque **Extraer**. Quedan unos 17 archivos sueltos (sin carpetas).
2. En GitHub: **+** (arriba a la derecha) → **New repository** → nombre `rey-servidor` → **Private** → **Create repository**.
3. En la página del repositorio vacío: toque el enlace **uploading an existing file** → **choose your files** → seleccione **todos** los archivos extraídos → espere a que terminen de subir → **Commit changes**.
4. Compruebe que en la lista del repositorio se vean los archivos directamente (`index.js`, `server.js`, `Dockerfile`, `render.yaml`…), **sin ninguna carpeta**. `render.yaml` y `Dockerfile` tienen que estar ahí.

**2. Crear el servidor en Render**
1. En Render: **New +** → **Blueprint** → conecte GitHub → escoja `rey-servidor`.
2. Render lee `render.yaml` y le pregunta una cosa: **APPS_SCRIPT_URL**. Pegue la dirección `/exec`.
3. Revise el precio que muestra y confirme con **Apply**. Tarda unos 5 minutos.

**3. Comprobar que quedó bien**
- Abra `https://rey-servidor.onrender.com/salud` (Render le muestra la dirección exacta arriba, en la página del servicio). Debe decir `"ok":true` y `"baseDatos":true`.
- Para ver las velocidades: en Render → su servicio → **Environment** → copie el valor de `ADMIN_CLAVE` y abra `https://…onrender.com/velocidad?clave=SU_CLAVE`.

**4. Hacer que la app lo use** (mejor de noche, sin gente trabajando)
1. Abra `config.js` de la app y déjelo así:
   ```js
   window.REY_API = "https://rey-servidor.onrender.com";
   window.REY_API_RESPALDO = "https://script.google.com/macros/s/…/exec";   // la de siempre
   ```
2. Suba la app a Netlify como siempre. Los teléfonos se actualizan solos.
3. Entre a **Configuración → Velocidad del servidor → Medir ahora** y compare con las cifras de antes. El `ping` debe bajar de ~2 s a una fracción de segundo. Si el resto de las consultas **no** mejoran o empeoran, vuelva a poner `REY_API` con la dirección de Google (paso de deshacer abajo).

**Para deshacer:** en `config.js` ponga otra vez la dirección de Google en `REY_API` y borre la línea de `REY_API_RESPALDO`. Suba a Netlify. Los datos nunca se tocan.

**5. Dirección propia (opcional)** · Render → su servicio → **Settings → Custom Domains** → `api.caixsa.com`. Render le dice qué registro `CNAME` crear donde administra el dominio. Después use `https://api.caixsa.com` en `config.js`.

## Mantenimiento

| Qué | Cómo |
|---|---|
| Actualizar el servidor | En el repositorio: **Add file → Upload files** y suba los archivos nuevos (los que tengan el mismo nombre se reemplazan). Render publica solo. |
| Ver si algo falla | Render → su servicio → **Logs**. Ahí nunca aparecen PINs ni datos de las consultas. |
| Respaldos de la base de datos | Render → `rey-db` → **Backups**. Confirme ahí que estén activos en su plan **antes** de migrar datos reales (Fase 1 en adelante). |
| Costo | Se ve en Render → **Billing**. Son dos piezas: el servidor y la base de datos. |
| Cambiar la clave de `/velocidad` | Render → **Environment** → `ADMIN_CLAVE`. |

## Variables de entorno

| Variable | Para qué | Por defecto |
|---|---|---|
| `APPS_SCRIPT_URL` | Dirección `/exec` de Apps Script | (obligatoria) |
| `DATABASE_URL` | Base de datos PostgreSQL | la pone Render |
| `ORIGENES` | Páginas que pueden usar el servidor | `https://app.caixsa.com` |
| `ADMIN_CLAVE` | Clave de `/velocidad` (sin ella, esa página queda cerrada) | — |
| `TRUST_PROXY` | `1` en Render (para limitar por la IP real) | `0` |
| `WARM_MINUTES` | Cada cuántos minutos despierta a Apps Script (`0` = nunca) | `4` |
| `LIMITE_ENTRAR_POR_MINUTO` | Intentos de entrar por IP por minuto | `40` |
| `LIMITE_POR_MINUTO` | Consultas por IP por minuto | `400` |

## Seguridad en pocas palabras

- La conexión es HTTPS (Render pone el certificado).
- El servidor **no guarda ni imprime** lo que viene en la consulta (ahí va el PIN). Solo anota el nombre de la consulta y cuánto tardó.
- Solo `app.caixsa.com` puede leer las respuestas desde un navegador.
- Límite por IP: más estricto para entrar, así nadie prueba PINs sin parar. Además, Apps Script sigue bloqueando a quien falle 8 veces con el mismo nombre.
