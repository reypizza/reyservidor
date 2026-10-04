-- Base del servidor. Cada migración corre una sola vez y queda anotada en «migraciones».

-- Una bitácora liviana de velocidad: qué consulta, cuánto tardó, si salió bien.
-- NUNCA guarda los datos de la consulta (ahí viene el PIN).
CREATE TABLE llamadas (
  id        bigserial PRIMARY KEY,
  en        timestamptz NOT NULL DEFAULT now(),
  fn        text        NOT NULL,
  nativa    boolean     NOT NULL,          -- true = la contestó este servidor; false = se pasó a Apps Script
  ok        boolean     NOT NULL,
  ms_total  integer     NOT NULL,          -- lo que tardó la puerta completa
  ms_origen integer                        -- lo que tardó Apps Script (si se le preguntó)
);
CREATE INDEX llamadas_en_idx ON llamadas (en);
CREATE INDEX llamadas_fn_en_idx ON llamadas (fn, en);

-- Para lo que cambia datos: si el teléfono repite una consulta por mala conexión, no se guarda dos veces.
CREATE TABLE solicitudes (
  rid          text PRIMARY KEY,
  fn           text        NOT NULL,
  estado       text        NOT NULL CHECK (estado IN ('en_curso', 'ok', 'error')),
  resultado    jsonb,
  creada_en    timestamptz NOT NULL DEFAULT now(),
  terminada_en timestamptz
);
CREATE INDEX solicitudes_creada_idx ON solicitudes (creada_en);
