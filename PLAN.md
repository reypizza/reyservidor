# Hoja de ruta: de Apps Script a servidor propio

**Regla de oro:** cada módulo tiene **una sola fuente de verdad**. Mientras un módulo vive en las hojas, Apps Script manda; cuando se migra, manda la base de datos y las hojas pasan a ser copia para reportes. Nunca los dos a la vez.

| Fase | Qué se hace | Qué gana | Estado |
|---|---|---|---|
| **0** | Puerta única + despertador de Apps Script + medición + respaldo automático en la app | `ping` instantáneo; la 1.ª consulta del día ya no tarda 7 s (por comprobar); se ve dónde se va el tiempo | **Lista** |
| **1** | Cimientos de usuarios: Apps Script entrega la lista de usuarios por una puerta interna con clave; el servidor guarda los PIN transformados, reconoce a cada persona igual que `quien` (mismos mensajes y reglas) y contesta `listaEntrada` (lo primero que pide la app al abrir) | Abrir la app ya no espera ~1,5 s a Google para mostrar la pantalla del PIN. Además, base para que lo migrado se autentique sin Google | **Lista** (por conectar: ver README) |
| **2** | **Indicadores, lecturas**: el servidor mantiene una copia al día de las hojas de indicadores y contesta `resumenAreas` (tarjetas de Inicio) y `tableroKpi` sin Google; lo que cambia indicadores sigue pasando por Apps Script. Siguientes en esta misma fase: `pantallaKpi` (lo que ve cada gerente), `listaKpis*`, `resultadosKpi*` | Tarjetas de Inicio y Tablero de ~3 s a décimas | **2a lista** (resumen y tablero); 2b y 2c pendientes |
| **3** | **Todas las lecturas aquí**: el mismo `Codigo.gs` corre en el servidor sobre una copia de todas las pestañas (modo sombra → modo nativa). Lo que guarda datos sigue en Google. Es el primer paso que acelera TODA la app (entrar, datos, notificaciones, listas) | `entrar` y las pantallas de ~5 s a milésimas | **Lista para probar en modo sombra** |
| **3b** | Sin internet en el teléfono (cola local con la misma llave `rid`) | Anotar aunque se caiga la red | Pendiente |
| **4** | **Caja**: ingresos, gastos, validación y confirmación en banco | Lo más delicado (dinero): después de probar el método con indicadores | Pendiente |
| **5** | **Inventario y bodega**: conteos, solicitudes, traslados, diferencias, compras | El módulo más grande | Pendiente |
| **6** | Lo que quede (calendarios, planilla, presupuestos…) y se apaga Apps Script | Un solo sistema | Pendiente |

## Decisiones por tomar antes de la Fase 3 (sin internet)

- **Hora de un registro hecho sin red:** se guarda la hora del teléfono marcada «sin conexión»; el servidor guarda además la hora a la que llegó. Nadie puede cambiarla a mano.
- **Dos teléfonos cambian lo mismo sin red:** para cuentas (caja, inventario) se acepta **todo** lo registrado y se avisa de la diferencia; para configuraciones, gana lo último que llegue.

## Cómo se migra cada módulo (siempre igual)

1. Se importan los datos de las hojas a PostgreSQL (script repetible).
2. Se escriben las funciones en `src/nativas.js` y se prueban contra los mismos casos que Apps Script.
3. Se compara el resultado nuevo con el de Apps Script con datos reales (sin que la app lo use todavía).
4. Se enciende: la app agrega esas funciones a `REY_NATIVAS` (para que no caigan al respaldo) y las hojas pasan a ser copia.
5. Marcha atrás posible durante unos días: se apaga la función nativa y vuelve Apps Script (las hojas siguen al día por la copia).
