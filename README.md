# Memi-EmpleosPublicos
Scraper de empleospublicos.cl vía API JSON -> Supabase (GitHub Actions, lunes y viernes) -> Looker Studio.

## Puesta en marcha
1. Supabase: crear proyecto y ejecutar `supabase_schema.sql` en SQL Editor.
2. GitHub > Settings > Secrets > Actions: `SUPABASE_URL` y `SUPABASE_KEY` (service_role key).
3. Actions > "Scraper Empleos Públicos" > Run workflow.
4. Looker Studio: conector PostgreSQL (Supabase > Connect > Session pooler) sobre la vista `v_empleos_dashboard`.

Nota: ~20% de las ofertas apuntan a sitios externos (JUNJI, UBB, etc.); llevan id `ext-<hash>`.
