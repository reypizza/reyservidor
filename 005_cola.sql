-- Guardado rápido: lo que el servidor guardó primero en su copia y todavía hay que reenviar (o ya se reenvió) a la hoja de Google.
CREATE TABLE cola_escrituras (
  id          bigserial PRIMARY KEY,
  rid         text,
  fn          text        NOT NULL,
  ops         jsonb       NOT NULL,                 -- las celdas exactas que cambió: {t:'v', h:hoja, r:fila, c:columna, a:antes, v:ahora} y {t:'f', ...formato}
  resultado   jsonb,                                -- lo que se le contestó a la app (para contestar lo mismo si repite la misma llave)
  estado      text        NOT NULL DEFAULT 'pendiente',   -- pendiente | aplicada | conflicto
  intentos    integer     NOT NULL DEFAULT 0,
  creada_en   timestamptz NOT NULL DEFAULT now(),
  aplicada_en timestamptz,
  detalle     text
);
CREATE INDEX cola_escrituras_estado_idx ON cola_escrituras (estado, id);
CREATE UNIQUE INDEX cola_escrituras_rid_idx ON cola_escrituras (rid) WHERE rid IS NOT NULL;
-- Llaves (rid) de gastos o ingresos que se mandaron a Google por el camino de siempre: si la app repite esa llave, se repite en Google
-- (que la reconoce), nunca por el camino rápido (que no la conoce y lo guardaría dos veces).
CREATE TABLE rid_google (
  rid  text PRIMARY KEY,
  fn   text NOT NULL,
  en   timestamptz NOT NULL DEFAULT now()
);
