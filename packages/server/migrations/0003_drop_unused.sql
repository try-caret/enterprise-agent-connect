-- Never read or written by the service.
alter table connections drop column access_ceiling;
alter table grants drop column access_level;
