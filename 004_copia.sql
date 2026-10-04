-- La copia completa de las pestañas (cifrada: ahí viajan los PIN), y el registro de cuánto coincide el servidor con Google.
CREATE TABLE copia_hojas (
  nombre      text PRIMARY KEY,
  datos       bytea NOT NULL,                -- JSON comprimido y cifrado con una clave que solo vive en el servidor (PIN_PEPPER)
  actualizado timestamptz NOT NULL DEFAULT now()
);
CREATE TABLE comparaciones (
  id        bigserial PRIMARY KEY,
  en        timestamptz NOT NULL DEFAULT now(),
  fn        text    NOT NULL,
  igual     boolean NOT NULL,
  ms_google integer,
  ms_propio integer,
  detalle   text                              -- dónde y cómo difieren (corto)
);
CREATE INDEX comparaciones_fn_idx ON comparaciones (fn, en);
