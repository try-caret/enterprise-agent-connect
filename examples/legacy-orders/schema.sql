-- A stand-in for a company's legacy orders database, and the user agents reach it as.
-- Postgres runs it on first start (docker-entrypoint-initdb.d).
create table customers (
  code text primary key,
  name text not null
);

create table orders (
  order_no integer primary key,
  customer_code text not null references customers (code),
  status text not null default 'open',
  total numeric(10, 2) not null,
  created_on date not null default current_date
);

create table order_notes (
  id serial primary key,
  order_no integer not null references orders (order_no),
  note text not null,
  created_at timestamptz not null default now()
);

insert into customers values ('ACME', 'Acme Corp'), ('GLOBEX', 'Globex Inc');
insert into orders values
  (1001, 'ACME', 'shipped', 1250.00, '2026-09-02'),
  (1002, 'ACME', 'open', 89.90, '2026-10-20'),
  (1003, 'GLOBEX', 'open', 430.00, '2026-10-21');

-- Only what the connector's queries need: reads, adding notes, changing an order's status.
create role agents login password 'agents-example';
grant select on customers, orders, order_notes to agents;
grant insert on order_notes to agents;
grant usage on sequence order_notes_id_seq to agents;
grant update (status) on orders to agents;
