drop table rpt_bills_detail_monthend;
drop table rpt_ifp_bills_detail_monthend;
drop table rpt_ifp_osfinint_monthend;
drop table rpt_run_date_monthend;
drop table rpt_bank_monthend;
drop table rpt_customer_monthend;

create table rpt_bills_detail_monthend as select * from rpt_bills_detail;
create table rpt_ifp_bills_detail_monthend as select * from rpt_ifp_bills_detail;
create table rpt_ifp_osfinint_monthend as select * from rpt_ifp_osfinint_run_date;
create table rpt_run_date_monthend as select * from rpt_run_date;
create table rpt_bank_monthend as select * from rpt_bank;
create table rpt_customer_monthend as select * from rpt_customer;

COMMIT;
QUIT;
