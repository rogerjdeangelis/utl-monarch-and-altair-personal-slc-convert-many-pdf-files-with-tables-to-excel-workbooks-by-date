/* Adapted from utl-monarch-and-altair-personal-slc-convert-many-pdf-files-with-tables-to-excel-workbooks-by-date.sas
   Rebuilds the same three date-named datasets the source script builds,
   then reproduces its "PDF metadata" step: a PROC SQL query against
   dictionary.tables that pattern-matches member names and builds a
   pdfname column with CATS(), exactly like the source's
     select cats('d:/pdf/',memname,'.pdf') as pdfname
     from dictionary.tables where memname eqt "DATE"
   The source's original WHERE clause used the PROC SQL truncated-string
   operator EQT ("equal to truncated strings" -- SAS 9.4 SQL Procedure
   User's Guide, sql-expression, "Operators and Order of Evaluation",
   Group 7) to match names starting with "DATE". That construct is
   filed separately as a Jenner compatibility gap; this bundle swaps in
   the equivalent LIKE pattern so the rest of the author's SQL/CATS
   logic can be demonstrated end to end. The source's `libname xls
   excel "d:/xls/pdf.xlsx";` (writing straight to an Excel workbook) is
   dropped -- the query result is left as a plain WORK table instead of
   an external Excel libname. */

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

proc sql;
  create
     table pdfsmeta as
  select
     cats('d:/pdf/',memname,'.pdf') as pdfname
  from
     dictionary.tables
  where
      memname like "DATE%"
;quit;

proc print data=pdfsmeta;
run;
