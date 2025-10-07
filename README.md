# utl-monarch-and-altair-personal-slc-convert-many-pdf-files-with-tables-to-excel-workbooks-by-date
Monarch and Altair personal slc convert many pdf files with tables to excel workbooks by date

    %let pgm=utl-monarch-and-altair-personal-slc-convert-many-pdf-files-with-tables-to-excel-workbooks-by-date;

    %stop_submission;

    Monarch and Altair personal slc convert many pdf files with tables to excel workbooks by date

    github
    https://github.com/rogerjdeangelis/utl-monarch-and-altair-personal-slc-convert-many-pdf-files-with-tables-to-excel-workbooks-by-date

    community.altair
    https://community.altair.com/discussion/17417/does_monarch-data-prep-studio-allow-itemized-exportation?tab=all&utm_source=community-search&utm_medium=organic-search&utm_term=monarch%20excel

    Coverting an arbiraery number of pfd file to excel.

    Process

    Create excel sheet with the paths to your pdf files (just add more.

    -----------------------+
    | A1| fx    | PDFNAME  |
    -------------------------------+
    [_] |          A               |
    -------------------------------|
     1  | PDFNAME                  |
     -- |----------+---------------|
     2  | d:/pdf/date2025_09_14.pdf|
     -- |--------------------------|
     3  | d:/pdf/date2025_10_04.pdf|
     -- |--------------------------|
     4  | d:/pdf/date2025_12_15.pdf|
     -- +--------------------------+
     [PDFSMETA]

    Use R tm package with the english corpus
    to convert pdf to text

    -------------------------------------
    | DATE       | AGE | LUNCH| DINNER  |
    |----------+-+-----+------+---------|
    | 2025-09-14 | 13  | 26   | 34      |
    |------------+-----+------+---------|
    | 2025-09-14 | 13  | 15   | 28      |
    |------------+-----+------+---------|
    | 2025-09-14 | 14  | 22   | 22      |
    -------------------------------------

    Use r Openxlsx to parse text lines to
    excel workbooks columns

    -----------------------+
    | A1| fx    | DATE     |
    ----------------------------------------+
    [_] |          A |   B |   C  |    D    |
    ----------------------------------------|
     1  | DATE       | AGE | LUNCH| DINNER  |
     -- |----------+-+-----+------+---------|
     2  | 2025-09-14 | 13  | 26   | 34      |
     -- |------------+-----+------+---------|
     3  | 2025-09-14 | 13  | 15   | 28      |
     -- |------------+-----+------+---------|
     4  | 2025-09-14 | 14  | 22   | 22      |
    ----------------------------------------|
    [2025-09-14]

    /*                   _
    (_)_ __  _ __  _   _| |_
    | | `_ \| `_ \| | | | __|
    | | | | | |_) | |_| | |_
    |_|_| |_| .__/ \__,_|\__|
            |_|
    */
    /*--- create 3 pdfs ---*/


           INPUT PDFS                                    OUTPUT
           ====                                          ======

    d:/xls/pdf.xls META DATA
                                               d:/xls/ date2025_09_14.xlsx
    -----------------------+
    | A1| fx    | PDFNAME  |                   -----------------------+
    -------------------------------+           | A1| fx    | DATE     |
    [_] |          A               |           ----------------------------------------+
    -------------------------------|           [_] |          A |   B |   C  |    D    |
     1  | PDFNAME                  |           ----------------------------------------|
     -- |----------+---------------|            1  | DATE       | AGE | LUNCH| DINNER  |
     2  | d:/pdf/date2025_09_14.pdf|            -- |----------+-+-----+------+---------|
     -- |--------------------------|            2  | 2025-09-14 | 13  | 26   | 34      |
     3  | d:/pdf/date2025_10_04.pdf|            -- |------------+-----+------+---------|
     -- |--------------------------|            3  | 2025-09-14 | 13  | 15   | 28      |
     4  | d:/pdf/date2025_12_15.pdf|            -- |------------+-----+------+---------|
     -- +--------------------------+            4  | | 14  | 22   | 22      |
     [PDFSMETA]                                ----------------------------------------|
                                               [2025-09-14]

    PDFS                                       d:/xls/ date2025_10_04.xlsx
    ===                                        -----------------------+
                                               | A1| fx    | DATE     |
    d:/pdf/date2025_09_14.pdf loosklike        ----------------------------------------+
                                               [_] |          A |   B |   C  |    D    |
    ---------------------------------|         ----------------------------------------|
    | DATE       | AGE| LUNCH| DINNER|          1  | DATE       | AGE | LUNCH| DINNER  |
    |----------+-+----+------+-------|          -- |------------+-----+------+---------|
    | 2025-09-14 | 13 | 26   | 34    |          2  | 2025-10-04 | 14  | 19   | 12      |
    |------------+----+------+-------|          -- |------------+-----+------+---------|
    | 2025-09-14 | 13 | 15   | 28    |          3  | 2025-10-04 | 14  | 13   | 32      |
    |------------+----+------+-------|          -- |------------+-----+------+---------|
    | 2025-09-14 | 14 | 22   | 22    |          4  | 2025-10-04 | 12  | 17   | 13      |
    ---------------------------------|          ---------------------------------------|
                                                [date2025_10_04]
    d:/pdf/date2025_10_04.pdf
                                               d:/xls/ date2025_12_15.xlsx
    ---------------------------------|         -----------------------+
    | DATE       | AGE| LUNCH| DINNER|         | A1| fx    | DATE     |
    |------------+----+------+-------|         ----------------------------------------+
    | 2025-10-04 | 14 | 19   | 12    |         [_] |          A |   B |   C  |    D    |
    |------------+----+------+-------|         ----------------------------------------|
    | 2025-10-04 | 14 | 13   | 32    |          1  | DATE       | AGE | LUNCH| DINNER  |
    |------------+----+------+-------|          -- |------------+-----+------+---------|
    | 2025-10-04 | 12 | 17   | 13    |          2  | 2025-12-15 | 14  | 19   | 12      |
    ---------------------------------|          -- |------------+-----+------+---------|
                                                3  | 2025-12-15 | 14  | 14   | 32      |
    d:/pdf/date2025_12_15.pdf                   -- |------------+-----+------+---------|
                                                4  | 2025-12-15 | 12  | 17   | 13      |
    d:/xls/ date2025_12_15.xlsx                 ---------------------------------------|
    ---------------------------------|          [2025-12-15]
    | DATE       | AGE| LUNCH| DINNER|
    |------------+----+------+-------|
    | 2025-12-15 | 14 | 19   | 12    |
    |------------+----+------+-------|
    | 2025-12-15 | 14 | 14   | 32    |
    |------------+----+------+-------|
    | 2025-12-15 | 12 | 17   | 13    |
    ---------------------------------|

    /*                   _
    (_)_ __  _ __  _   _| |_
    | | `_ \| `_ \| | | | __|
    | | | | | |_) | |_| | |_
    |_|_| |_| .__/ \__,_|\__|
            |_|
    */

    %utlfkil(%sysfunc(pathname(WPSWBHTM))); /*-- disable precode --*/
    &_init_; /*-- enable listing output and set options          --*/

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
    run;quit;

    /*
    | | ___   __ _
    | |/ _ \ / _` |
    | | (_) | (_| |
    |_|\___/ \__, |
             |___/
    */
    /*--- CREATE THE THREE PDFS ---*/

    %array(pdfs,values=date2025_09_14 date2025_10_04 date2025_12_15);

    title;footnote;
    %do_over(pdfs,phrase=%str(
       ods pdf file="d:/pdf/?.pdf";
       proc print data=?;
       run;quit;
       ods pdf close;));



    87       ODS _ALL_ CLOSE;
    288       FILENAME WPSWBHTM TEMP;
    NOTE: Writing HTML(WBHTML) BODY file d:\wpswrk\_TD19596\#LN00013
    289       ODS HTML(ID=WBHTML) BODY=WPSWBHTM GPATH="d:\wpswrk\_TD19596";
    290       %utlfkil(%sysfunc(pathname(WPSWBHTM))); /*-- disable precode --*/
    291       &_init_; /*-- enable listing output and set options          --*/
    292
    293       data date2025_09_14  date2025_10_04  date2025_12_15;
    294         input
    295           date$11. age lunch dinner;
    296         select (date);
    297           when ('2025-09-14') output date2025_09_14;
    298           when ('2025-10-04') output date2025_10_04;
    299           when ('2025-12-15') output date2025_12_15;
    300           otherwise;
    301         end;
    302       cards4;

    NOTE: Data set "WORK.date2025_09_14" has 3 observation(s) and 4 variable(s)
    NOTE: Data set "WORK.date2025_10_04" has 3 observation(s) and 4 variable(s)
    NOTE: Data set "WORK.date2025_12_15" has 3 observation(s) and 4 variable(s)
    NOTE: The data step took :
          real time : 0.037
          cpu time  : 0.015


    303       2025-09-14 13 26 34
    304       2025-09-14 13 15 28
    305       2025-09-14 14 22 22
    306       2025-10-04 14 19 12
    307       2025-10-04 14 13 32
    308       2025-10-04 12 17 13
    309       2025-12-15 14 19 12
    310       2025-12-15 14 13 32
    311       2025-12-15 12 17 13
    312       ;;;;
    313       run;quit;
    314
    315
    316       /*--- CREATE THE THREE PDFS ---*/
    317
    318       %array(pdfs,values=date2025_09_14 date2025_10_04 date2025_12_15);
    319
    320       title;footnote;
    321       %do_over(pdfs,phrase=%str(
    NOTE: View opening spill file for output observations.
    322          ods pdf file="d:/pdf/?.pdf";
    323          proc print data=?;
    324          run;quit;
    325          ods pdf close;));
    NOTE: Writing file d:\pdf\date2025_09_14.pdf
    NOTE: 3 observations were read from "WORK.date2025_09_14"
    NOTE: Procedure print step took :
          real time : 0.257
          cpu time  : 0.328


    NOTE: Writing file d:\pdf\date2025_10_04.pdf
    NOTE: 3 observations were read from "WORK.date2025_10_04"
    NOTE: Procedure print step took :
          real time : 0.029
          cpu time  : 0.078


    NOTE: Writing file d:\pdf\date2025_12_15.pdf
    NOTE: 3 observations were read from "WORK.date2025_12_15"
    NOTE: Procedure print step took :
          real time : 0.015
          cpu time  : 0.046


    326
    327
    328       quit; run;
    329       ODS _ALL_ CLOSE;
    330       FILENAME WPSWBHTM CLEAR;


    /*--- CREATE THE PDF META DATA WORKBOOK ---*/

    %utlfkil(%sysfunc(pathname(WPSWBHTM))); /*-- disable precode --*/
    &_init_; /*-- enable listing output and set options          --*/

    %utlfkil(d:/xls/pdf.xlsx);            /*--- incase you rerun ---*/

    libname xls excel "d:/xls/pdf.xlsx";

    proc datasets lib=xls;                /*--- incase you rerun ---*/
     delete pdfmeta;
    run;quit;

    proc sql;
      create
         table xls.pdfsmeta as
      select
         cats('d:/pdf/',memname,'.pdf') as pdfname
      from
         dictionary.tables
      where
          memname eqt "DATE"
    ;quit;

    proc print data=xls.pdfsmeta;
    run;quit;

    libname xls clear;

    run;quit;

    /*
     _ __  _ __ ___   ___ ___  ___ ___
    | `_ \| `__/ _ \ / __/ _ \/ __/ __|
    | |_) | | | (_) | (_|  __/\__ \__ \
    | .__/|_|  \___/ \___\___||___/___/
    |_|
    */

    %utlfkil(%sysfunc(pathname(WPSWBHTM))); /*-- disable precode --*/
    &_init_; /*-- enable listing output and set options          --*/

    options set=RHOME "D:\d451";
    proc r;
    submit;
    library("tm")
    library(openxlsx)

    convert_pdfs_to_dataframes <- function(pdf_files) {

      for (file in pdf_files) {
        # Extract the base name without extension for the dataframe name
        df_name <- tools::file_path_sans_ext(basename(file))

        Rpdf <- readPDF(control = list(text = "-layout"))
        corpus <- VCorpus(URISource(file), readerControl = list(reader = Rpdf))
        text_content <- content(content(corpus)[[1]])
        # Split by newlines and convert to dataframe
        lines_vector <- unlist(strsplit(text_content, "\n"))
        want <- data.frame(lines = lines_vector, stringsAsFactors = FALSE)

        # Convert to proper dataframe with columns
        dfout <- read.table(text = want$lines, header = TRUE, stringsAsFactors = FALSE)
        assign(df_name, dfout, envir = .GlobalEnv)

        # create excel workbooks
        wb <- createWorkbook()
        addWorksheet(wb, df_name)
        writeData(wb, df_name, dfout)
        saveWorkbook(wb, paste0("d:/xls/",df_name, ".xlsx"),
          overwrite = TRUE)
      }
    }

    df <- read.xlsx(xlsxFile = "d:/xls/pdf.xlsx")
    pdf_files <- df$PDFNAME

    convert_pdfs_to_dataframes(pdf_files)

    endsubmit;

    run;quit;

    /*
    | | ___   __ _
    | |/ _ \ / _` |
    | | (_) | (_| |
    |_|\___/ \__, |
             |___/
    */

    331       ODS _ALL_ CLOSE;
    332       FILENAME WPSWBHTM TEMP;
    NOTE: Writing HTML(WBHTML) BODY file d:\wpswrk\_TD19596\#LN00015
    333       ODS HTML(ID=WBHTML) BODY=WPSWBHTM GPATH="d:\wpswrk\_TD19596";
    334       %utlfkil(%sysfunc(pathname(WPSWBHTM))); /*-- disable precode --*/
    335       &_init_; /*-- enable listing output and set options          --*/
    336
    337       options set=RHOME "D:\d451";
    338       proc r;
    339       submit;
    340       library("tm")
    341       library(openxlsx)
    342
    343       convert_pdfs_to_dataframes <- function(pdf_files) {
    344
    345         for (file in pdf_files) {
    346           # Extract the base name without extension for the dataframe name
    347           df_name <- tools::file_path_sans_ext(basename(file))
    348
    349           Rpdf <- readPDF(control = list(text = "-layout"))
    350           corpus <- VCorpus(URISource(file), readerControl = list(reader = Rpdf))
    351           text_content <- content(content(corpus)[[1]])
    352           # Split by newlines and convert to dataframe
    353           lines_vector <- unlist(strsplit(text_content, "\n"))
    354           want <- data.frame(lines = lines_vector, stringsAsFactors = FALSE)
    355
    356           # Convert to proper dataframe with columns
    357           dfout <- read.table(text = want$lines, header = TRUE, stringsAsFactors = FALSE)
    358           assign(df_name, dfout, envir = .GlobalEnv)
    359
    360           # create excel workbooks
    361           wb <- createWorkbook()
    362           addWorksheet(wb, df_name)
    363           writeData(wb, df_name, dfout)
    364           saveWorkbook(wb, paste0("d:/xls/",df_name, ".xlsx"),
    365             overwrite = TRUE)
    366         }
    367       }
    368
    369       df <- read.xlsx(xlsxFile = "d:/xls/pdf.xlsx")
    370       pdf_files <- df$PDFNAME
    371
    372       convert_pdfs_to_dataframes(pdf_files)
    373
    374       endsubmit;
    NOTE: Using R version 4.5.1 (2025-06-13 ucrt) from d:\r451

    NOTE: Submitting statements to R:

    Loading required package: NLP
    > library("tm")
    > library(openxlsx)
    >
    > convert_pdfs_to_dataframes <- function(pdf_files) {
    +
    +   for (file in pdf_files) {
    +     # Extract the base name without extension for the dataframe name
    +     df_name <- tools::file_path_sans_ext(basename(file))
    +
    +     Rpdf <- readPDF(control = list(text = "-layout"))
    +     corpus <- VCorpus(URISource(file), readerControl = list(reader = Rpdf))
    +     text_content <- content(content(corpus)[[1]])
    +     # Split by newlines and convert to dataframe
    +     lines_vector <- unlist(strsplit(text_content, "\n"))
    +     want <- data.frame(lines = lines_vector, stringsAsFactors = FALSE)
    +
    +     # Convert to proper dataframe with columns
    +     dfout <- read.table(text = want$lines, header = TRUE, stringsAsFactors = FALSE)
    +     assign(df_name, dfout, envir = .GlobalEnv)
    +
    +     # create excel workbooks
    +     wb <- createWorkbook()
    +     addWorksheet(wb, df_name)
    +     writeData(wb, df_name, dfout)
    +     saveWorkbook(wb, paste0("d:/xls/",df_name, ".xlsx"),
    +       overwrite = TRUE)
    +   }
    + }
    >
    > df <- read.xlsx(xlsxFile = "d:/xls/pdf.xlsx")
    > pdf_files <- df$PDFNAME
    >
    > convert_pdfs_to_dataframes(pdf_files)

    NOTE: Processing of R statements complete

    >
    375
    376       run;quit;
    NOTE: Procedure r step took :
          real time : 1.259
          cpu time  : 0.000


    377       quit; run;
    378       ODS _ALL_ CLOSE;
    379       FILENAME WPSWBHTM CLEAR;

    /*           _               _
      ___  _   _| |_ _ __  _   _| |_
     / _ \| | | | __| `_ \| | | | __|
    | (_) | |_| | |_| |_) | |_| | |_
     \___/ \__,_|\__| .__/ \__,_|\__|
                    |_|
    */

    d:/xls/ date2025_09_14.xlsx

    -----------------------+
    | A1| fx    | DATE     |
    -------------------------------------------------------+
    [_] |          A         |    B    |   C     |    D    |
    -------------------------------------------------------|
     1  | DATE               | AGE     | LUNCH   | DINNER  |
     -- |----------+---------+---------+---------+---------|
     2  | 2025-09-14         | 13      | 26      | 34      |
     -- |--------------------+---------+---------+---------|
     3  | 2025-09-14         | 13      | 15      | 28      |
     -- |--------------------+---------+---------+---------|
     4  | 2025-09-14         | 14      | 22      | 22      |
    -------------------------------------------------------|
    [2025-09-14]

    d:/xls/ date2025_10_04.xlsx

    -----------------------+
    | A1| fx    | DATE     |
    -------------------------------------------------------+
    [_] |          A         |    B    |   C     |    D    |
    -------------------------------------------------------|
     1  | DATE               | AGE     | LUNCH   | DINNER  |
     -- |--------------------+---------+---------+---------|
     2  | 2025-10-04         | 14      | 19      | 12      |
     -- |--------------------+---------+---------+---------|
     3  | 2025-10-04         | 14      | 13      | 32      |
     -- |--------------------+---------+---------+---------|
     4  | 2025-10-04         | 12      | 17      | 13      |
     ------------------------------------------------------|
    [date2025_10_04]

    d:/xls/ date2025_12_15.xlsx

    -----------------------+
    | A1| fx    | DATE     |
    -------------------------------------------------------+
    [_] |          A         |    B    |   C     |    D    |
    -------------------------------------------------------|
     1  | DATE               | AGE     | LUNCH   | DINNER  |
     -- |--------------------+---------+---------+---------|
     2  | 2025-12-15         | 14      | 19      | 12      |
     -- |--------------------+---------+---------+---------|
     3  | 2025-12-15         | 14      | 14      | 32      |
     -- |--------------------+---------+---------+---------|
     4  | 2025-12-15         | 12      | 17      | 13      |
     ------------------------------------------------------|
     [2025-12-15]

    /*              _
      ___ _ __   __| |
     / _ \ `_ \ / _` |
    |  __/ | | | (_| |
     \___|_| |_|\__,_|

    */

