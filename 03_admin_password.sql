-- =====================================================================
-- Password reset by the owner (no email needed).
-- Run once in Supabase SQL Editor, after 01_schema.sql.
-- From the app: Account > User management > "New password".
-- If the OWNER forgets his own password, run in the SQL Editor:
--   select public.admin_set_password('owner@example.com', 'NewPass123');
-- =====================================================================
create or replace function public.admin_set_password(p_email text, p_password text)
returns void language plpgsql security definer
set search_path = public, extensions, auth as $$
declare v_id uuid;
begin
  if auth.uid() is not null and coalesce(public.my_role(), '') <> 'owner' then
    raise exception 'owner only';
  end if;
  if length(coalesce(p_password, '')) < 6 then
    raise exception 'password too short';
  end if;
  select id into v_id from auth.users where lower(email) = lower(p_email);
  if v_id is null then raise exception 'no user with email %', p_email; end if;
  update auth.users
     set encrypted_password = crypt(p_password, gen_salt('bf')), updated_at = now()
   where id = v_id;
  -- sign the user out everywhere so the old password stops working immediately
  begin delete from auth.sessions where user_id = v_id; exception when others then null; end;
end $$;

revoke all on function public.admin_set_password(text, text) from public, anon;
grant execute on function public.admin_set_password(text, text) to authenticated;
