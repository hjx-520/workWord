truncate table RPT_IS9_PROVISION;
insert into RPT_IS9_PROVISION select substr(deal_ref,5,10), tot_provision from RPT_IS9_PROVISION_STG;
COMMIT;
QUIT;

