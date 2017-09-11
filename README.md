# utl_excel_add_to_sheet
Add a side by side report or graph to an existing worksheet. General excel layout.

    ```  Adding second report to existing sheet (generic excel layout R )  ```
    ```    ```
    ```    ```
    ```  T1006790 SAS Forum: SAS/WPS/R Two reports one starting at B3 and the second at G3 (side by side)  ```
    ```    ```
    ```    ```
    ```     WORKING CODE  ```
    ```    ```
    ```        OPEN WORKBOOK  ```
    ```    ```
    ```          writeWorksheet(wb,   males, sheet = "sex", startRow = 3,startCol = 2, header = TRUE);  ```
    ```          writeWorksheet(wb, females, sheet = "sex", startRow = 3,startCol = 7, header = TRUE);  ```
    ```    ```
    ```        CLOSED WORKBOOK  ```
    ```    ```
    ```          writeWorksheet(wb,   males, sheet = "sex", startRow = 3,startCol = 2, header = TRUE);  ```
    ```    ```
    ```          *  close workbook - leaver R;  ```
    ```    ```
    ```          * add second report starting in row=3 column=5  ```
    ```          writeWorksheet(wb, females, sheet = "sex", startRow = 3,startCol = 5, header = TRUE);  ```
    ```    ```
    ```     See end of message for side by side table and graph  ```
    ```    ```
    ```  see  ```
    ```  https://goo.gl/AuFkyU  ```
    ```  https://communities.sas.com/t5/ODS-and-Base-Reporting/Excel-print-output-across-the-page/m-p/394597  ```
    ```    ```
    ```  see  ```
    ```  https://goo.gl/4VfdMR  ```
    ```  https://communities.sas.com/t5/ODS-and-Base-Reporting/How-to-print-report-in-a-specific-cell-in-excel/m-p/378853  ```
    ```    ```
    ```    ```
    ```  I am using SAS 9.3 version and I am trying to output some data to excel  ```
    ```  starting from a particular row and column number. In the attached excel  ```
    ```  screenshot I have two reports. The first one is starting at Column B Row  ```
    ```  3 and the second report is starting at Column G Row 3 and so forth.  ```
    ```    ```
    ```  Could someone show how to direct SAS to start printing the report from a  ```
    ```  particular row and column cell in excel. Thanks for your help!  ```
    ```    ```
    ```    ```
    ```  HAVE  ```
    ```  ====  ```
    ```    ```
    ```    Up to 40 obs SD1.HAVE total obs=10  ```
    ```    ```
    ```    Obs    NAME       SEX  ```
    ```    ```
    ```      1    Alfred      M  ```
    ```      2    Alice       F  ```
    ```      3    Barbara     F  ```
    ```      4    Carol       F  ```
    ```      5    Henry       M  ```
    ```      6    James       M  ```
    ```      7    Jane        F  ```
    ```      8    Janet       F  ```
    ```      9    Jeffrey     M  ```
    ```     10    John        M  ```
    ```    ```
    ```    ```
    ```    ```
    ```    ```
    ```   WANT excel sheet with males starting at A3 and females at G3  ```
    ```  ==============================================================  ```
    ```    ```
    ```  d:/xls/sex_fm.xlsx  ```
    ```    ```
    ```      +---------------------+-----------------------------------+  ```
    ```      |  A  |  B    |  C    |  D  |  E  |  F    |  G    |  H    |  ```
    ```      +---------------------+-----------------------------------+  ```
    ```  1   |     |       |       |     |     |       |       |       |  ```
    ```      |-----+-------+-------|-----+-----+-------+-------+-------|  ```
    ```  2   |     |       |       |     |     |       |       |       |  ```
    ```      |-----+-------+-------+-----+-----+-------+-------+-------+  ```
    ```  3   |     |NAME   |SEX    |     |     |       |NAME   |SEX    |  ```
    ```      |-----+-------+-------|-----+-----+-------+-------+-------|  ```
    ```  4   |     |Alfred |M      |     |     |       |Alice  |F      |  ```
    ```      |-----+-------+-------+-----+-----+-------+-------+-------+  ```
    ```  5   |     |Alex   |M      |     |     |       |Barbara|F      |  ```
    ```      |-----+-------+-------+-----+-----+-------+-------+-------+  ```
    ```  6   |     |Bob    |M      |     |     |       |Carol  |F      |  ```
    ```      |-----+-------+-------+-----+-----+-------+-------+-------+  ```
    ```  7   |     |Chris  |M      |     |     |       |Jane   |F      |  ```
    ```      |-----+-------+-------+-----+-----+-------+-------+-------+  ```
    ```  8   |     |Henry  |M      |     |     |       |       |       |  ```
    ```      -----------------------------------------------------------  ```
    ```  ...  ```
    ```    ```
    ```  [SEX]  ```
    ```    ```
    ```  *                _              _       _  ```
    ```   _ __ ___   __ _| | _____    __| | __ _| |_ __ _  ```
    ```  | '_ ` _ \ / _` | |/ / _ \  / _` |/ _` | __/ _` |  ```
    ```  | | | | | | (_| |   <  __/ | (_| | (_| | || (_| |  ```
    ```  |_| |_| |_|\__,_|_|\_\___|  \__,_|\__,_|\__\__,_|  ```
    ```    ```
    ```  ;  ```
    ```  options validvarname=upcase;  ```
    ```  libname sd1 "d:/sd1";  ```
    ```  data sd1.have;  ```
    ```    set sashelp.class(keep=name sex obs=10);  ```
    ```  run;quit;  ```
    ```    ```
    ```  *          _       _   _  ```
    ```   ___  ___ | |_   _| |_(_) ___  _ __  ```
    ```  / __|/ _ \| | | | | __| |/ _ \| '_ \  ```
    ```  \__ \ (_) | | |_| | |_| | (_) | | | |  ```
    ```  |___/\___/|_|\__,_|\__|_|\___/|_| |_|  ```
    ```  ;  ```
    ```    ```
    ```  %utl_submit_r64('  ```
    ```  source("c:/Program Files/R/R-3.3.2/etc/Rprofile.site",echo=T);  ```
    ```  library(haven);  ```
    ```  library(XLConnect);  ```
    ```  have<-read_sas("d:/sd1/have.sas7bdat");  ```
    ```  have;  ```
    ```  males<-have[have$SEX=="M",];  ```
    ```  females<-have[have$SEX=="F",];  ```
    ```  wb <- loadWorkbook("d:/xls/sex_mf.xlsx", create = TRUE);  ```
    ```  createSheet(wb, name = "sex");  ```
    ```  writeWorksheet(wb, males, sheet = "sex", startRow = 3,startCol = 2, header = TRUE);  ```
    ```  writeWorksheet(wb, females, sheet = "sex", startRow = 3,startCol = 7, header = TRUE);  ```
    ```  saveWorkbook(wb);  ```
    ```  ');  ```
    ```    ```
    ```    ```
    ```  *     _                    _                       _    _                 _  ```
    ```    ___| | ___  ___  ___  __| |  __      _____  _ __| | _| |__   ___   ___ | | __  ```
    ```   / __| |/ _ \/ __|/ _ \/ _` |  \ \ /\ / / _ \| '__| |/ / '_ \ / _ \ / _ \| |/ /  ```
    ```  | (__| | (_) \__ \  __/ (_| |   \ V  V / (_) | |  |   <| |_) | (_) | (_) |   <  ```
    ```   \___|_|\___/|___/\___|\__,_|    \_/\_/ \___/|_|  |_|\_\_.__/ \___/ \___/|_|\_\  ```
    ```    ```
    ```  ;  ```
    ```    ```
    ```  %utlfkil(d:/xls/sex_1.xlsx);  ```
    ```    ```
    ```  %utl_submit_r64('  ```
    ```  source("c:/Program Files/R/R-3.3.2/etc/Rprofile.site",echo=T);  ```
    ```  library(haven);  ```
    ```  library(XLConnect);  ```
    ```  have<-read_sas("d:/sd1/have.sas7bdat");  ```
    ```  have;  ```
    ```  males<-have[have$SEX=="M",];  ```
    ```  wb <- loadWorkbook("d:/xls/sex_1.xlsx", create = TRUE);  ```
    ```  createSheet(wb, name = "samesheet");  ```
    ```  writeWorksheet(wb, males, sheet = "samesheet", startRow = 3,startCol = 2, header = TRUE);  ```
    ```  saveWorkbook(wb);  ```
    ```  ');  ```
    ```    ```
    ```  %utl_submit_r64('  ```
    ```  source("c:/Program Files/R/R-3.3.2/etc/Rprofile.site",echo=T);  ```
    ```  library(haven);  ```
    ```  library(XLConnect);  ```
    ```  have<-read_sas("d:/sd1/have.sas7bdat");  ```
    ```  have;  ```
    ```  females<-have[have$SEX=="F",];  ```
    ```  wb <- loadWorkbook("d:/xls/sex_1.xlsx", create = FALSE);  ```
    ```  writeWorksheet(wb, females, sheet = "samesheet", startRow = 3,startCol = 10, header = TRUE);  ```
    ```  saveWorkbook(wb);  ```
    ```  ');  ```
    ```    ```
    ```    ```
    ```  *                          _                            _  ```
    ```   _ __ ___ _ __   ___  _ __| |_     __ _ _ __ __ _ _ __ | |__  ```
    ```  | '__/ _ \ '_ \ / _ \| '__| __|   / _` | '__/ _` | '_ \| '_ \  ```
    ```  | | |  __/ |_) | (_) | |  | |_   | (_| | | | (_| | |_) | | | |  ```
    ```  |_|  \___| .__/ \___/|_|   \__|   \__, |_|  \__,_| .__/|_| |_|  ```
    ```           |_|                      |___/          |_|  ```
    ```  ;  ```
    ```    ```
    ```  T1003920 General Excel Layout for report and graph (side by side)  ```
    ```    ```
    ```    ```
    ```  This is  works in unix and windows without any Microsoft products?  ```
    ```    ```
    ```  Puts the table and graph side by side. R puts the SAS objects side by side.  ```
    ```    ```
    ```  for output see  ```
    ```  https://www.dropbox.com/s/6shxcdum3mi1xgn/sbysout.xlsx?dl=0  ```
    ```    ```
    ```    ```
    ```  HAVE  Excel sheet1 with 'proc report' and png formatted histogram  ```
    ```  ==================================================================  ```
    ```    ```
    ```  EXCEL:  D:/XLS/SBYS.XLSX  ```
    ```    ```
    ```          A           B            C  ```
    ```  ROW  ---------|-----------|---------------  ```
    ```    ```
    ```  1    Country     Product     Actual Sales  ```
    ```    ```
    ```  2    CANADA      BED           $47,729.00  ```
    ```  3                CHAIR         $50,239.00  ```
    ```  4                DESK          $52,187.00  ```
    ```  5                SOFA          $50,135.00  ```
    ```  6                TABLE         $46,700.00  ```
    ```  7    GERMANY     BED           $46,134.00  ```
    ```  8               CHAIR          $47,105.00  ```
    ```  9                DESK          $48,502.00  ```
    ```    ```
    ```  ------  ```
    ```  SHEET1  ```
    ```  ------  ```
    ```    ```
    ```  PNG : D:/PNG/SBYS.PNG  ```
    ```    ```
    ```  PNG Graphic Histogram  ```
    ```    ```
    ```  Frequency  ```
    ```  6 +  ```
    ```    |  ```
    ```  4 +                 *****   *****  ```
    ```    |                 *****   *****  ```
    ```  2 +  *****   *****  *****   *****  ```
    ```    |  *****   *****  *****   *****  ```
    ```    --------------------------------  ```
    ```       BED     CHAIR   DESK   TABLE  ```
    ```    ```
    ```                    PRODUCT  ```
    ```    ```
    ```    ```
    ```  WANT ( new excel workbook with sheet1 with side by side report and histogram)  ```
    ```  =============================================================================  ```
    ```    ```
    ```  EXCEL : D:/XLS/SBYSOUT.XLSX  ```
    ```    ```
    ```    ```
    ```  WANT (NEW EXCEL SHEET WITH 'proc report'(not png) as histogram  ```
    ```    ```
    ```    ```
    ```  EXCEL   A           B            C  ```
    ```  ROW  ---------|-----------|---------------  ```
    ```                                                 Frequency  ```
    ```  1    Country     Product     Actual Sales      6 +  ```
    ```                                                   |  ```
    ```  2    CANADA      BED           $47,729.00      4 +                 *****   *****  ```
    ```  3                CHAIR         $50,239.00        |                 *****   *****  ```
    ```  4                DESK          $52,187.00      2 +  *****   *****  *****   *****  ```
    ```  5                SOFA          $50,135.00        |  *****   *****  *****   *****  ```
    ```  6                TABLE         $46,700.00        --------------------------------  ```
    ```  7    GERMANY     BED           $46,134.00           BED     CHAIR   DESK   TABLE  ```
    ```  8               CHAIR          $47,105.00  ```
    ```  9                DESK          $48,502.00                      PRODUCT  ```
    ```    ```
    ```  ------  ```
    ```  SHEET1  ```
    ```  ------  ```
    ```    ```
    ```    ```
    ```  WORKING CODE  ```
    ```    ```
    ```     SAS  ```
    ```         ods excel options(sheet_name="sheet1" start_at="A1");  ```
    ```         proc report data=sashelp.prdsale;  ```
    ```         ods graphics on / width=4in imagefmt=png imagename="sbys";  ```
    ```         proc sgplot data=sashelp.class;  ```
    ```    ```
    ```     R  ```
    ```         wb <- loadWorkbook("d:/xls/sbys.xlsx", create = FALSE);  ```
    ```         writeWorksheet(wbout,sheetin,"sheetout",startCol=1,header=T);  ```
    ```         addImage(wbout, filename = "d:/png/sbys.png", name = "sheetout",originalSize = TRUE);  ```
    ```    ```
    ```    ```
    ```  FULL SOLUTION  ```
    ```    ```
    ```  * CREATE TWO SAS REPORTS (in sheet1 and sheet2)  ```
    ```    ```
    ```  %utlfkil(d:\xls\sbys.xlsx);  ```
    ```  %utlfkil(d:\png\sbs.png);  ```
    ```    ```
    ```  ods excel file="d:\xls\sbys.xlsx";  ```
    ```    ```
    ```  ods excel options(sheet_name="sheet1" start_at="A1");  ```
    ```    ```
    ```  proc report data=sashelp.prdsale;  ```
    ```  column country product actual;  ```
    ```  define country / group;  ```
    ```  define product / group;  ```
    ```  rbreak after / summarize;  ```
    ```  run;quit;  ```
    ```    ```
    ```  Ods excel close;  ```
    ```    ```
    ```  ods listing style=journal;  ```
    ```  ods listing  gpath='d:/png';  ```
    ```  ods graphics on / width=4in imagefmt=png imagename="sbys";  ```
    ```    ```
    ```  proc sgplot data=sashelp.class;  ```
    ```  format age 2.;  ```
    ```  vbar age /datalabel;  ```
    ```  run;quit;  ```
    ```    ```
    ```  ods graphics off;  ```
    ```    ```
    ```    ```
    ```  %utl_submit_r64('  ```
    ```  source("c:/Program Files/R/R-3.3.2/etc/Rprofile.site",echo=T);  ```
    ```  library(XLConnect);  ```
    ```  wb <- loadWorkbook("d:/xls/sbys.xlsx", create = FALSE);  ```
    ```  sheetin = readWorksheet(wb, sheet = "sheet1");  ```
    ```  wbout <- loadWorkbook("d:/xls/sbysout.xlsx", create = TRUE);  ```
    ```  createSheet(wbout, name = "sheetout");  ```
    ```  writeWorksheet(wbout,sheetin,"sheetout",startCol=1,header=T);  ```
    ```  createName(wbout, name = "sheetout", formula = "sheetout!$E$2");  ```
    ```  addImage(wbout, filename = "d:/png/sbys.png", name = "sheetout",  ```
    ```           originalSize = TRUE);  ```
    ```  saveWorkbook(wb);  ```
    ```  saveWorkbook(wbout);  ```
    ```  ');  ```
    ```    ```
