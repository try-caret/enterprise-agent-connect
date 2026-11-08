-- Work that every process of the service would otherwise do at once (a Graph subscription's sync,
-- a pass over the running sessions): whoever holds the name's lease does it. A process that dies
-- holding one leaves it to lapse at expires_at.
create table leases (
  name text primary key,
  token uuid not null,
  expires_at timestamptz not null
);
alter table leases enable row level security;
alter table leases force row level security;
create policy service on leases for all to current_user using (true) with check (true);
