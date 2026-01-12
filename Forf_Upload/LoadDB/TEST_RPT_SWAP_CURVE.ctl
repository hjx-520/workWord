Load Data
INFILE "RPT_SWAP_CURVE.csv"
BADFILE "RPT_SWAP_CURVE.BAD"
DISCARDFILE "RPT_SWAP_CURVE.DSC"
TRUNCATE
INTO TABLE TEST_RPT_SWAP_CURVE
FIELDS TERMINATED BY "," OPTIONALLY ENCLOSED BY '"'
TRAILING NULLCOLS
(
  Curve_name "TRIM (:Curve_name)",
  Currency_Code,
  Start_Date  Date "yyyyMMdd",
  End_Date  Date "yyyyMMdd",
  Discount_Factor
)
