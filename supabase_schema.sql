-- Ejecutar UNA vez en Supabase > SQL Editor. Reemplaza tabla y vista (los datos se regeneran en cada corrida).
drop view if exists public.v_empleos_dashboard;
drop table if exists public.empleos_publicos cascade;

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
  formacion         text,   -- requisitos académicos (detalle de la oferta)
  especializacion   text,
  experiencia       text,   -- requisitos laborales (texto)
  anios_exp_min     int,    -- menor nº de años mencionado en experiencia
  anios_exp_max     int,    -- mayor nº de años mencionado en experiencia
  carreras          text,   -- carreras detectadas en formación (coma)
  grado             int,    -- grado EUS mencionado en la oferta (null si no aparece)
  estamento         text,   -- Planta / Contrata / Honorarios / Código del Trabajo / Suplencia/Reemplazo / Otro
  postulantes       int,    -- el portal no publica la cantidad: siempre null
  extraido_en       timestamptz not null
);
create index on public.empleos_publicos (fecha_cierre);
create index on public.empleos_publicos (region);
alter table public.empleos_publicos enable row level security;

-- Vista para Looker Studio (security_invoker: respeta RLS de la tabla)
create view public.v_empleos_dashboard with (security_invoker = true) as
select t.*,
  case when t.perfil <> 'Otros' then 'Afín a mi perfil' else 'Otros' end as afin,
  greatest(t.fecha_cierre - current_date, 0) as dias_restantes_hoy,
  case
    when t.fecha_cierre - current_date <= 2 then '🔴 Cierra pronto (≤ 48h)'
    when t.fecha_cierre - current_date <= 5 then '🟡 Cierra esta semana'
    else '🟢 Tiempo suficiente'
  end as semaforo_cierre,
  case
    when t.renta_numerica = 0 then 'Sin información'
    when t.renta_numerica < 1000000 then 'Menos de $1M'
    when t.renta_numerica <= 2000000 then '$1M a $2M'
    when t.renta_numerica <= 3000000 then '$2M a $3M'
    else 'Sobre $3M'
  end as rango_salarial
from (
  select e.*,
    case
      when e.region ilike '%arica%' then 'CL-AP' when e.region ilike '%tarap%' then 'CL-TA'
      when e.region ilike '%antofa%' then 'CL-AN' when e.region ilike '%atacama%' then 'CL-AT'
      when e.region ilike '%coquimbo%' then 'CL-CO' when e.region ilike '%valpara%' then 'CL-VS'
      when e.region ilike '%metropolitana%' then 'CL-RM' when e.region ilike '%higgins%' then 'CL-LI'
      when e.region ilike '%maule%' then 'CL-ML' when e.region ilike '%uble%' then 'CL-NB'
      when e.region ilike '%bio%' or e.region ilike '%bío%' then 'CL-BI' when e.region ilike '%araucan%' then 'CL-AR'
      when e.region ilike '%los r%' then 'CL-LR' when e.region ilike '%los lagos%' then 'CL-LL'
      when e.region ilike '%ays%' then 'CL-AI' when e.region ilike '%magallanes%' then 'CL-MA'
    end as region_iso,
    -- Perfil del usuario: Administrador Público, mg. Ciencia de Datos, dipl. Control de Gestión (editable)
    case
      when e.cargo ilike any (array['%reemplazo%','%auxiliar%','%chofer%','%administrativ%','%técnico%','%tecnico%','%médic%','%medic%','%enfermer%','%matron%','%kinesi%','%paramédic%','%odontol%','%psicólog%','%docente%','%profesor%','%educador%','%abogad%','%conductor%','%secretari%']) then 'Otros'
      when e.cargo ilike any (array['%datos%','%estadíst%','%estadist%','%inteligencia%','%analítica%','%analitica%','%información%','%informacion%','%business%','%científico%','%cientifico%','%bi %','%estudios%']) then 'Datos y analítica'
      when e.cargo ilike any (array['%control de gestión%','%control de gestion%','%gestión%','%gestion%','%planificaci%','%presupuest%','%desempeño%','%desempeno%','%indicador%','%evaluaci%','%calidad%']) then 'Control de gestión y presupuesto'
      when e.cargo ilike any (array['%financier%','%contab%','%auditor%','%tesorer%','%control interno%','%adquisicion%','%compras%']) then 'Finanzas y auditoría'
      when e.cargo ilike any (array['%administrador público%','%administrador publico%','%administradora pública%','%políticas públicas%','%politicas publicas%','%analista%','%profesional%','%jefe%','%jefa%','%director%','%subdirector%','%coordinador%','%encargado%','%proyectos%']) then 'Gestión y políticas públicas'
      else 'Otros'
    end as perfil
  from public.empleos_publicos e
) t;
