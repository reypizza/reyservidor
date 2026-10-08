# Servidor de Rey Pizza · CAIX, S.A.

Este es el servidor propio de la app. **Hoy hace seis cosas** (Fases 0, 1, 2 y 3) y las demás se van sumando módulo por módulo (ver `PLAN.md`):

1. **Puerta única.** La app le habla a este servidor. Lo que todavía no está migrado, él se lo pasa a Apps Script (igual que hacía el teléfono) y devuelve la respuesta tal cual. Nadie tiene que cambiar nada.
2. **Mantiene despierto a Apps Script.** Cada 4 minutos le toca la puerta, para que la primera consulta del día no tarde 7 segundos.
3. **Mide la velocidad real** de cada consulta (sin guardar PINs ni datos) y la muestra en `/velocidad`.
4. **Conoce a los usuarios** (Fase 1): aprende de la hoja «Usuarios» cada 2 minutos, guarda los PIN ya transformados (nunca legibles) y contesta él solo la primera consulta al abrir la app (`listaEntrada`), sin esperar a Google.
5. **Contesta las lecturas de indicadores** (Fase 2): mantiene una copia al día de las hojas de indicadores y, con ella, contesta él solo las tarjetas de Inicio del administrador (`resumenAreas`) y el Tablero (`tableroKpi`), sin esperar a Google. Todo lo que *cambia* indicadores sigue pasando por Apps Script, y la hoja sigue siendo la que manda.
6. **Corre el código de la app aquí mismo** (Fase 3, apagada hasta que usted la encienda): el mismo `Codigo.gs` de Apps Script, sin cambios, sobre una copia de todas las pestañas. Primero en **modo sombra** (Google sigue contestando y se compara en silencio) y, cuando las respuestas coinciden, en **modo nativa** (las lecturas las contesta el servidor en milésimas de segundo; lo que guarda datos sigue yendo a Google).

Si este servidor se cae, **la app sigue funcionando** con Apps Script (respaldo automático).

---

## Qué necesita (una sola vez)

- Una cuenta en **GitHub** (github.com, gratis).
- Una cuenta en **Render** (render.com). Pide tarjeta; antes de crear nada le muestra el precio mensual.
- La dirección `/exec` de su Apps Script (la que está en `config.js`: `https://script.google.com/macros/s/…/exec`).

## Paso a paso

**1. Subir el proyecto a GitHub** (se puede hacer desde una tablet o el teléfono)
1. En el celular o tablet, abra el zip **rey-servidor-github.zip** con la app **Archivos** y toque **Extraer**. Quedan unos 34 archivos sueltos (sin carpetas).
2. En GitHub: **+** (arriba a la derecha) → **New repository** → nombre `rey-servidor` → **Private** → **Create repository**.
3. En la página del repositorio vacío: toque el enlace **uploading an existing file** → **choose your files** → seleccione **todos** los archivos extraídos (**no** el .zip) → espere a que terminen de subir → **Commit changes**.
4. Compruebe que en la lista del repositorio se vean los archivos directamente (`index.js`, `server.js`, `Dockerfile`, `render.yaml`, `Codigo.gs`…), **sin ninguna carpeta**.

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
| Actualizar el servidor | Suba los archivos nuevos a GitHub (mismo repositorio). Render publica solo. |
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
| `SERVER_TOKEN` | Clave secreta compartida con Apps Script (Fase 1) | la inventa Render |
| `PIN_PEPPER` | Clave secreta para transformar los PIN (Fase 1) | la inventa Render |
| `SYNC_MINUTES` | Cada cuántos minutos aprende de la hoja Usuarios | `2` |
| `LECTURAS_MODO` | `apagado` (por defecto), `sombra`, `auto` (recomendado) o `nativa` (ver Fase 3) | `apagado` |
| `MIN_GRADUAR` | modo `auto`: cuántas comparaciones seguidas iguales necesita una pantalla para contestarse aquí | `12` |
| `MUESTRA_VERIFICACION` | En modo `nativa`, qué fracción de respuestas se compara después con Google (0 a 1) | `0.05` |
| `SYNC_KPI_SEGUNDOS` | Cada cuántos segundos revisa la hoja de indicadores (con la app en uso; `0` = nunca) | `60` |
| `LIMITE_ENTRAR_POR_MINUTO` | Intentos de entrar por IP por minuto | `40` |
| `LIMITE_POR_MINUTO` | Consultas por IP por minuto | `400` |

## Seguridad en pocas palabras

- La conexión es HTTPS (Render pone el certificado).
- El servidor **no guarda ni imprime** lo que viene en la consulta (ahí va el PIN). Solo anota el nombre de la consulta y cuánto tardó.
- Solo `app.caixsa.com` puede leer las respuestas desde un navegador.
- Límite por IP: más estricto para entrar, así nadie prueba PINs sin parar. Además, Apps Script sigue bloqueando a quien falle 8 veces con el mismo nombre.

---

## Fase 1: que el servidor conozca a los usuarios (una sola vez)

Es una conexión de dos pasos entre Render y Apps Script, con una clave secreta que **solo ustedes dos conocen**. El orden importa poco: mientras no esté todo conectado, el servidor sigue pasando todo a Apps Script y nada se rompe.

**A. Apps Script** (con el `Codigo.gs` nuevo)
1. Pegue el `Codigo.gs` nuevo en el proyecto de Apps Script y guarde.
2. **Implementar → Administrar implementaciones →** el lápiz → **Versión: Nueva versión → Implementar**.

**B. GitHub y Render**
1. Extraiga el zip nuevo y, en su repositorio, **Add file → Upload files** → seleccione **todos** los archivos → **Commit changes**. (Los que ya existen se reemplazan; los nuevos se agregan.)
2. Render publica solo (unos 3 a 5 minutos). Cuando termine, en **Environment** deben aparecer dos variables nuevas: `SERVER_TOKEN` y `PIN_PEPPER`.
   - Si no aparecen: **Add Environment Variable** y cree cada una con una clave larga inventada (30 letras y números mezclados). `PIN_PEPPER` no se copia a ningún lado.

**C. Copiar la clave a Apps Script**
1. En Render → **Environment** → toque el ojito de `SERVER_TOKEN` → copie el valor.
2. En Apps Script: **Configuración del proyecto** (el engranaje) → **Propiedades del script → Agregar propiedad**. Nombre: `SERVER_TOKEN`. Valor: lo que copió. **Guardar propiedades**.

**D. Comprobar** (espere 2 minutos)
- Abra `https://rey-servidor.onrender.com/estado?clave=SU_ADMIN_CLAVE`. Debe decir `"total"` y `"activos"` con el número de personas de su hoja Usuarios, `"error": null` y una fecha en `"ultimaOk"`. Si `"error"` dice «No autorizado», la clave de Apps Script no es idéntica a la de Render.
- En la app: **Configuración → Velocidad del servidor**. `listaEntrada` debe bajar de ~1,5 s a ~0,15 s.

**Si algo sale mal:** borre la propiedad `SERVER_TOKEN` en Apps Script. La puerta se cierra y el servidor vuelve a pasar todo a Google.

**Qué viaja y qué se guarda:** los PIN viajan de Apps Script al servidor por HTTPS, con esa clave. El servidor los transforma con `PIN_PEPPER` y no guarda ninguno legible. `/estado` nunca muestra PIN ni nombres.

---

## Fase 2: indicadores (no hay nada que conectar)

Usa la misma clave `SERVER_TOKEN` de la Fase 1, así que solo hay que **actualizar los dos lados**:
1. Pegar el `Codigo.gs` nuevo en Apps Script y crear una **nueva versión** de la implementación (agrega la puerta de exportación de indicadores).
2. Subir a GitHub **todos** los archivos del zip nuevo, ya extraídos (se reemplazan los que cambiaron y se agregan los nuevos). Render publica solo.

**Comprobar** (espere 2 minutos): `https://rey-servidor.onrender.com/estado?clave=SU_ADMIN_CLAVE` ahora trae, además de `usuarios`, un bloque `indicadores` con `indicadores` (cuántos hay en su hoja), `filas` (cuántos registros), `error: null` y una fecha en `ultimaOk`. Después, en **Velocidad del servidor** de la app, `resumenAreas` debe bajar de ~3 s a décimas de segundo, y en `/velocidad` debe salir con `"nativa": true`.

**Cómo se mantiene al día:** después de cada cambio que pasa por el servidor (guardar un indicador, anotar valores…) la copia se actualiza de inmediato, y la siguiente lectura espera a eso. Si alguien edita la hoja a mano, se entera en un minuto con la app en uso (5 minutos si nadie la usa). Cada 6 horas vuelve a traer todo. Si por cualquier motivo la copia tiene más de 3 minutos sin actualizarse, **contesta Apps Script como antes**, nunca datos viejos.

**Si algo se ve raro:** `POST /resincronizar?clave=SU_ADMIN_CLAVE` trae todo de nuevo (desde una herramienta como `curl`; no se puede desde la barra del navegador). Para apagar solo esto, borre `SERVER_TOKEN` en Apps Script: todo vuelve a pasar a Google.

---

## Fase 3: todas las lecturas aquí (se enciende cuando usted quiera)

**Importante:** el servidor lleva su propia copia de `Codigo.gs`. Cada vez que le cambie algo a `Codigo.gs` en Apps Script, **suba también el mismo archivo a GitHub** (reemplaza al anterior). Si las dos versiones no coinciden (el número `VERSION_CODIGO` al inicio del archivo), el servidor **no contesta nada por su cuenta** y todo pasa a Google. Eso evita respuestas con reglas viejas.

**Cómo encenderla (por pasos, sin riesgo):**
1. Publique el `Codigo.gs` nuevo en Apps Script (nueva versión) y suba los archivos del zip a GitHub. En ese momento **todavía no cambia nada**: el modo por defecto es «apagado».
2. En Render → su servicio → **Environment → Add Environment Variable**: `LECTURAS_MODO` = `sombra`. Render reinicia. El servidor trae una copia de todas las pestañas (1 a 3 minutos la primera vez; después queda guardada, cifrada, en la base de datos).
3. Use la app unos días como siempre. Abra `https://rey-servidor.onrender.com/comparaciones?clave=SU_ADMIN_CLAVE`: dice, por función, cuántas veces la respuesta del servidor fue **idéntica** a la de Google, con los tiempos de cada uno, y la última diferencia si hubo.
4. **Recomendado:** cambie `LECTURAS_MODO` a `auto`. Cada pantalla empieza contestándola Google y comparándose aquí; cuando sus últimas 12 comparaciones salen iguales, pasa sola a contestarse aquí (rápido). Una sola diferencia la regresa a Google. En `/estado?clave=…` se ve la lista `graduadas`. (O, si prefiere decidirlo usted, cambie a `nativa` cuando casi todo coincida.)
5. Con `nativa`, Las lecturas pasan a contestarse aquí. Lo que guarda datos (registrar, borrar…) se detecta solo y sigue yendo a Google; la copia se actualiza sola después de cada cambio.

**Salvavidas:** si algo falla, hay datos viejos o el código no coincide, contesta Google. Si una función da respuestas distintas 3 veces, se deja de usar aquí 30 minutos. Para apagarlo todo: `LECTURAS_MODO` = `apagado` (o bórrela).

**Qué NO hace todavía:** guardar datos, enviar correos, generar PDF ni subir archivos a Drive siguen en Apps Script. La hoja de Google sigue siendo la que manda.

**Seguridad:** la copia incluye la hoja Usuarios (con PIN). En la base de datos está **cifrada** con `PIN_PEPPER` (sin esa clave no se puede leer); en la memoria del servidor está sin cifrar mientras corre, igual que en Google.
