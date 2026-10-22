-- Ejecutar en Supabase > SQL Editor para agregar la fecha/hora de la última postulación registrada en el portal
alter table public.empleos_publicos add column if not exists ultima_postulacion timestamp;
