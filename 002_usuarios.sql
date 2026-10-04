-- Usuarios: copia de la hoja «Usuarios» (por ahora la hoja sigue mandando; el servidor la lee cada pocos minutos).
-- El PIN NUNCA se guarda legible: solo su transformación (scrypt + sal propia + clave secreta del servidor).
CREATE TABLE usuarios (
  id             bigserial PRIMARY KEY,
  nombre         text    NOT NULL,
  nombre_lc      text    NOT NULL UNIQUE,
  pin_hash       text    NOT NULL,                 -- vacío si la persona no tiene PIN
  pin_largo      integer NOT NULL DEFAULT 0,       -- cuántos dígitos tiene (la app avisa si el código es corto)
  rol            text    NOT NULL,
  activo         boolean NOT NULL,
  sucursal       text    NOT NULL DEFAULT '',
  confirma       boolean NOT NULL DEFAULT false,
  correo         text    NOT NULL DEFAULT '',
  actualizado_en timestamptz NOT NULL DEFAULT now()
);

-- Intentos fallidos por nombre (igual que Apps Script: 8 en 10 minutos cierran la entrada de ese nombre).
CREATE TABLE intentos (
  nombre_lc text PRIMARY KEY,
  n         integer     NOT NULL,
  hasta     timestamptz NOT NULL
);

-- Cuándo se aprendió por última vez de la hoja, y si hubo problema.
CREATE TABLE sincronizacion (
  clave          text PRIMARY KEY,
  huella         text,
  ultima_ok      timestamptz,
  ultimo_intento timestamptz,
  error          text
);
