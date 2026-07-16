/* Adapted from utl-monarch-and-altair-personal-slc-convert-many-pdf-files-with-tables-to-excel-workbooks-by-date.sas
   The DATA step below is byte-identical to the source repo's multi-output
   SELECT/WHEN split (it reads the same inline data and routes rows to the
   same three date-named datasets). The three PROC PRINT calls are added
   here only so the run produces visible output -- the original relies on
   the author's own %array/%do_over macro library (not included in the
   repo) to loop over the datasets and write them to PDF/Excel instead. */

data date2025_09_14  date2025_10_04  date2025_12_15;
  input
    date$11. age lunch dinner;
  select (date);
    when ('2025-09-14') output date2025_09_14;
    when ('2025-10-04') output date2025_10_04;
    when ('2025-12-15') output date2025_12_15;
    otherwise;
  end;
cards4;
2025-09-14 13 26 34
2025-09-14 13 15 28
2025-09-14 14 22 22
2025-10-04 14 19 12
2025-10-04 14 13 32
2025-10-04 12 17 13
2025-12-15 14 19 12
2025-12-15 14 13 32
2025-12-15 12 17 13
;;;;
run;

proc print data=date2025_09_14; run;
proc print data=date2025_10_04; run;
proc print data=date2025_12_15; run;
