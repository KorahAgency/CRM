delete from public.korah_crm_state
where id in (
  'korah-crm-pipe',
  'korah-crm-imp',
  'korah-scripts',
  'korah-hist',
  'korah-crm-v2-pipe',
  'korah-crm-v2-imp',
  'korah-crm-v2-scripts',
  'korah-crm-v2-hist'
);
