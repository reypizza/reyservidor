-- Copia de las hojas «Indicadores» y «Indicadores registros». La hoja sigue mandando; esto se mantiene al día solo
-- (cada minuto con actividad, y de inmediato tras cada cambio que pasa por este servidor) y sirve para contestar lecturas rápido.
CREATE TABLE indicadores (
  orden integer PRIMARY KEY,           -- el orden de la hoja (las pantallas lo usan)
  id    text    NOT NULL,
  doc   jsonb   NOT NULL
);
CREATE TABLE indicadores_registros (
  fila integer PRIMARY KEY,            -- el número de fila de la hoja (así se sabe cuáles son «las últimas»)
  id   text    NOT NULL,               -- vacío en filas sueltas: cuentan en la ventana, pero no se usan
  doc  jsonb   NOT NULL
);
ALTER TABLE sincronizacion
  ADD COLUMN cursor_fila integer NOT NULL DEFAULT 2,     -- desde qué fila falta traer
  ADD COLUMN ultima_fila integer NOT NULL DEFAULT 0,     -- cuántas filas tenía la hoja la última vez
  ADD COLUMN completa_en timestamptz;                    -- la última vez que se volvió a traer TODO
