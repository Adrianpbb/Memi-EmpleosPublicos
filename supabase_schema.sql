-- Ejecutar UNA vez en Supabase > SQL Editor.
-- Reemplaza la tabla anterior (los datos se regeneran en cada corrida).
drop table if exists public.empleos_publicos cascade;
drop function if exists public.reset_empleos_publicos();

create table public.empleos_publicos (
  id_convocatoria   text primary key,  -- 'ep-<id>' portal | 'ext-<hash>' sitio externo
  cargo             text,
  ministerio        text,
  institucion       text,
  numero_vacantes   int,
  area_trabajo      text,
  region            text,
  ciudad            text,
  tipo_vacante      text,
  renta_bruta       text,
  renta_numerica    int,
  fecha_inicio      date,
  fecha_cierre      date,
  dias_restantes    int,
  tipo_postulacion  text,
  cargo_profesional text,
  primer_empleo     boolean,
  url_oferta        text,
  estado            text,
  extraido_en       timestamptz not null
);

create index on public.empleos_publicos (fecha_cierre);
create index on public.empleos_publicos (region);

-- RLS activo: la service_role key (scraper) lo ignora; lectura pública no habilitada.
alter table public.empleos_publicos enable row level security;

-- Vista para Looker Studio con campos calculados al día (dias_restantes siempre fresco)
create or replace view public.v_empleos_dashboard as
select e.*,
  greatest(e.fecha_cierre - current_date, 0) as dias_restantes_hoy,
  case
    when e.fecha_cierre - current_date <= 2 then '🔴 Cierra pronto (≤ 48h)'
    when e.fecha_cierre - current_date <= 5 then '🟡 Cierra esta semana'
    else '🟢 Tiempo suficiente'
  end as semaforo_cierre,
  case
    when e.renta_numerica = 0 then 'Sin información'
    when e.renta_numerica < 1000000 then 'Menos de $1M'
    when e.renta_numerica <= 2000000 then '$1M a $2M'
    when e.renta_numerica <= 3000000 then '$2M a $3M'
    else 'Sobre $3M'
  end as rango_salarial
from public.empleos_publicos e;
