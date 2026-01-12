--update RPT_SWAP_CURVE set start_date = start_date + 9, end_date = end_date + 9;
--update RPT_SWAP_CURVE set start_date = start_date - 163, end_date = end_date - 163;
update RPT_SWAP_CURVE set start_date = (select monthend_date from rpt_run_date_monthend);
COMMIT;
QUIT;

