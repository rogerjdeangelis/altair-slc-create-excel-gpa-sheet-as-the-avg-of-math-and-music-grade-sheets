%let pgm=altair-slc-create-excel-gpa-sheet-as-the-avg-of-math-and-music-grade-sheets;

%stop_submission;

Altair slc create excel gpa sheet as the avg of math and music grade sheets

Although this is a trivial example, it opens up powerfull sql processing in Python and R

Too long to podt to a list, see github
https://github.com/rogerjdeangelis/altair-slc-create-excel-gpa-sheet-as-the-avg-of-math-and-music-grade-sheets

  SOLUTIONS

    1 slc proc r
    2 slc proc python
    3 enhanced sqlite functions

PROBLEM ADD SHEET GPA

-------------------------------------------------------   ------------------------------------
|                      INPUT                          |   |            OTPUT                 |
-------------------------------------------------------   ------------------------------------
d:\xls\class.xlsx                                         d:\xls\class.xlsx
SHEET:MATH                   SHEET:MUSIC                  SHEET:GPA
-------------------------+    ------------------------+   --------------------+
| A1| fx       | STUDENT |   | A1| fx     | STUDEN    |   | A1| fx    |STUDENT|
--------------------------   --------------------------   ------------------------------------
[_] |    A     |    B    |   [_] |    A     |    B    |   [_] |    A   |  B   |   C   |  D   |
--------------------------   --------------------------   ------------------------------------
 1  | STUDENT  | GRADE   |    1  | STUDENT  | GRADE   |    1  | STUDENT| MATH | MUSIC | MEAN |
 -- |----------+---------+    -- |----------+---------+    -- |--------+------+-------+------+
 2  |  MARY    |  64     |    2  |  JOSH    |  74     |    2  |  MARY  |  64  |  63   | 63.5 |
 -- |----------+---------+    -- |----------+---------+    -- |--------+------+-------+------+
 3  |  JOHN    |  70     |    3  |  JANE    |  70     |    3  |  ALEX  |  83  |  83   | 83   |
 -- |----------+---------+    -- |----------+---------+    -- |--------+------+-------+------+
 4  |  MILE    |  83     |    4  |  ALEX    |  83     |   [GPA]
 -- |----------+---------+    -- |----------+---------+
 5  |  ALEX    |  83     |    5  |  MARY    |  63     |   JOIN MATH AND MUSIC ON STUDENT NAME
 -- |----------+---------+    -- |----------+---------+   AND AVERAGE MAT AND MUSIC GRADE
 6  |  JOHN    |  93     |    6  |  MIKE    |  93     |
 -- |----------+---------+    -- |----------+---------+
[MATH]                        [MUSIC]

/*                   _
(_)_ __  _ __  _   _| |_
| | `_ \| `_ \| | | | __|
| | | | | |_) | |_| | |_
|_|_| |_| .__/ \__,_|\__|
        |_|
*/

proc datasets lib=workx kill ;
run;

%utlfkil(d:\xls\class.xlsx);

libname xls excel "d:\xls\class.xlsx";

data xls.math;
  input student$ grade;
cards4;
MARY 64
JOHN 70
MILE 83
ALEX 83
JOHN 93
;;;;
run;

data xls.music;
  input student$ grade;
cards4;
JOSH 74
JANE 70
ALEX 83
MARY 63
MIKE 93
;;;;
run;

libname xls clear;

/*************************************************************************************************************************/
/* TWO SHEETS IN WORKBOOK d:\xls\class.xlsx                                                                              */
/*                                                                                                                       */
/* SHEET:MATH                   SHEET:MUSIC                                                                              */
/* -------------------------+    ------------------------+                                                               */
/* | A1| fx       | STUDENT |   | A1| fx     | STUDEN    |                                                               */
/* --------------------------   --------------------------                                                               */
/* [_] |    A     |    B    |   [_] |    A     |    B    |                                                               */
/* --------------------------   --------------------------                                                               */
/*  1  | STUDENT  | GRADE   |    1  | STUDENT  | GRADE   |                                                               */
/*  -- |----------+---------+    -- |----------+---------+                                                               */
/*  2  |  MARY    |  64     |    2  |  JOSH    |  74     |                                                               */
/*  -- |----------+---------+    -- |----------+---------+                                                               */
/*  3  |  JOHN    |  70     |    3  |  JANE    |  70     |                                                               */
/*  -- |----------+---------+    -- |----------+---------+                                                               */
/*  4  |  MILE    |  83     |    4  |  ALEX    |  83     |                                                               */
/*  -- |----------+---------+    -- |----------+---------+                                                               */
/*  5  |  ALEX    |  83     |    5  |  MARY    |  63     |                                                               */
/*  -- |----------+---------+    -- |----------+---------+                                                               */
/*  6  |  JOHN    |  93     |    6  |  MIKE    |  93     |                                                               */
/*  -- |----------+---------+    -- |----------+---------+                                                               */
/* [MATH]                        [MUSIC]                                                                                 */
/*************************************************************************************************************************/

/*       _
/ |  ___| | ___   _ __  _ __ ___   ___   _ __
| | / __| |/ __| | `_ \| `__/ _ \ / __| | `__|
| | \__ \ | (__  | |_) | | | (_) | (__  | |
|_| |___/_|\___| | .__/|_|  \___/ \___| |_|
                 |_|
*/

options set=RHOME "C:\Progra~1\R\R-4.5.2\bin\r";
proc r;
submit;
library(sqldf)
library(openxlsx)
options(sqldf.dll = "d:/dll/sqlean.dll")

wb <- loadWorkbook("d:/xls/class.xlsx")

math  <- read.xlsx(wb, sheet = "math")
music <- read.xlsx(wb, sheet = "music")

gpa=sqldf('
  select
     l.student
    ,l.grade as math
    ,r.grade as music
    ,(l.grade + r.grade)/2 as mean
  from
     math as l join music as r
  on
     l.student = r.student
  ')

gpa
addWorksheet(wb, "gpa")
writeData(wb, sheet = "gpa", x = gpa)
saveWorkbook(wb,"d:/xls/class.xlsx", overwrite = TRUE)
endsubmit;
import data=workx.gpa r=gpa;
run;

/**************************************************************************************************************************/
/*  d:\xls\class.xlsx                      |                                                                              */
/* SHEET:GPA                               |   SLC DATTASET                                                               */
/* --------------------+                   |                                                                              */
/* | A1| fx    |STUDENT|                   |   WORKX.GPA                                                                  */
/* ------------------------------------    |                                                                              */
/* [_] |    A   |  B   |   C   |  D   |    |   Obs    STUDENT    MATH    MUSIC    MEAN                                    */
/* ------------------------------------    |                                                                              */
/*  1  | STUDENT| MATH | MUSIC | MEAN |    |    1      MARY       64       63     63.5                                    */
/*  -- |--------+------+-------+------+    |    2      ALEX       83       83     83.0                                    */
/*  2  |  MARY  |  64  |  63   | 63.5 |    |                                                                              */
/*  -- |--------+------+-------+------+    |                                                                              */
/*  3  |  ALEX  |  83  |  83   | 83   |    |                                                                              */
/*  -- |--------+------+-------+------+    |                                                                              */
/*  [GPA]                                  |                                                                              */
/**************************************************************************************************************************/

/*
| | ___   __ _
| |/ _ \ / _` |
| | (_) | (_| |
|_|\___/ \__, |
         |___/
*/

1                                          Altair SLC      17:43 Tuesday, February 17, 2026

NOTE: Copyright 2002-2025 World Programming, an Altair Company
NOTE: Altair SLC 2026 (05.26.01.00.000758)
      Licensed to Roger DeAngelis
NOTE: This session is executing on the X64_WIN11PRO platform and is running in 64 bit mode

NOTE: AUTOEXEC processing beginning; file is C:\wpsoto\autoexec.sas
NOTE: AUTOEXEC source line
1       +  ï»¿ods _all_ close;
           ^
ERROR: Expected a statement keyword : found "?"
NOTE: Library workx assigned as follows:
      Engine:        SAS7BDAT
      Physical Name: d:\wpswrkx

NOTE: Library slchelp assigned as follows:
      Engine:        WPD
      Physical Name: C:\Progra~1\Altair\SLC\2026\sashelp

NOTE: Library worksas assigned as follows:
      Engine:        SAS7BDAT
      Physical Name: d:\worksas

NOTE: Library workwpd assigned as follows:
      Engine:        WPD
      Physical Name: d:\workwpd


LOG:  17:43:27
NOTE: 1 record was written to file PRINT

NOTE: The data step took :
      real time : 0.031
      cpu time  : 0.015


NOTE: AUTOEXEC processing completed

1         options set=RHOME "C:\Progra~1\R\R-4.5.2\bin\r";
2         proc r;
3         submit;
4         library(sqldf)
5         library(openxlsx)
6         options(sqldf.dll = "d:/dll/sqlean.dll")
7
8         wb <- loadWorkbook("d:/xls/class.xlsx")
9
10        math  <- read.xlsx(wb, sheet = "math")
11        music <- read.xlsx(wb, sheet = "music")
12
13        gpa=sqldf('
14          select
15             l.student
16            ,l.grade as math
17            ,r.grade as music
18            ,(l.grade + r.grade)/2 as mean
19          from
20             math as l join music as r
21          on
22             l.student = r.student
23          ')
24
25        gpa
26        addWorksheet(wb, "gpa")
27        writeData(wb, sheet = "gpa", x = gpa)
28        saveWorkbook(wb,"d:/xls/class.xlsx", overwrite = TRUE)
29        endsubmit;
NOTE: Using R version 4.5.2 (2025-10-31 ucrt) from C:\Program Files\R\R-4.5.2

NOTE: Submitting statements to R:

> library(sqldf)
Loading required package: gsubfn
Loading required package: proto
Loading required package: RSQLite
> library(openxlsx)
> options(sqldf.dll = "d:/dll/sqlean.dll")
>
> wb <- loadWorkbook("d:/xls/class.xlsx")
>
> math  <- read.xlsx(wb, sheet = "math")
> music <- read.xlsx(wb, sheet = "music")
>
> gpa=sqldf('
+   select
+      l.student
+     ,l.grade as math
+     ,r.grade as music
+     ,(l.grade + r.grade)/2 as mean
+   from
+      math as l join music as r
+   on
+      l.student = r.student
+   ')
>
> gpa
> addWorksheet(wb, "gpa")
Error in addWorksheet(wb, "gpa") :
  A worksheet by the name 'gpa' already exists! Sheet names must be unique case-insensitive.
> writeData(wb, sheet = "gpa", x = gpa)
> saveWorkbook(wb,"d:/xls/class.xlsx", overwrite = TRUE)

NOTE: Processing of R statements complete

30        import data=workx.gpa r=gpa;
NOTE: Creating data set 'WORKX.gpa' from R data frame 'gpa'
NOTE: Column names modified during import of 'gpa'
NOTE: Data set "WORKX.gpa" has 2 observation(s) and 4 variable(s)

31        run;
NOTE: Procedure r step took :
      real time : 1.910
      cpu time  : 0.015


ERROR: Error printed on page 1

NOTE: Submitted statements took :
      real time : 2.020
      cpu time  : 0.093

/*___        _                                      _   _
|___ \   ___| | ___   _ __  _ __ ___    _ __  _   _| |_| |__   ___  _ __
  __) | / __| |/ __| | `_ \| `__/ _ \  | `_ \| | | | __| `_ \ / _ \| `_ \
 / __/  \__ \ | (__  | |_) | | | (_) | | |_) | |_| | |_| | | | (_) | | | |
|_____| |___/_|\___| | .__/|_|  \___/  | .__/ \__, |\__|_| |_|\___/|_| |_|
                     |_|               |_|    |___/
*/
/*--- NEED TO RECREATE EXCEL WORKBOOK BECAUSE WORKBOOK ALREAD HAVE SHEET GPA ---*/

proc datasets lib=workx kill ;
run;

%utlfkil(d:\xls\class.xlsx);

libname xls excel "d:\xls\class.xlsx";

data xls.math;
  input student$ grade;
cards4;
MARY 64
JOHN 70
MILE 83
ALEX 83
JOHN 93
;;;;
run;

data xls.music;
  input student$ grade;
cards4;
JOSH 74
JANE 70
ALEX 83
MARY 63
MIKE 93
;;;;
run;

libname xls clear;

/*           _   _
 _ __  _   _| |_| |__   ___  _ __
| `_ \| | | | __| `_ \ / _ \| `_ \
| |_) | |_| | |_| | | | (_) | | | |
| .__/ \__, |\__|_| |_|\___/|_| |_|
|_|    |___/
*/


options set=PYTHONHOME "D:\py314";
proc python;
submit;
import pyarrow
import pandas as pd
import sqlite3
from pandasql import sqldf
import pandas as pd

# Create connection and load sqlean.dll
conn = sqlite3.connect(":memory:")
conn.enable_load_extension(True)
conn.load_extension(r"d:/dll/sqlean.dll")

# Read Excel sheets
wb_file = r"d:/xls/class.xlsx"
math_df = pd.read_excel(wb_file, sheet_name="math")
music_df = pd.read_excel(wb_file, sheet_name="music")

# SQL query using pandasql (equivalent to sqldf with sqlean.dll)
# Note: pandasql uses SQLite under the hood, similar to your r sqldf setup
query = '''
  SELECT
     l.student,
     l.grade as math,
     r.grade as music,
     (l.grade + r.grade)/2 as mean
  FROM math_df as l
  JOIN music_df as r
  ON l.student = r.student
'''

gpa = sqldf(query, locals())
print(gpa)

# Write GPA to new sheet (preserves existing math/music sheets)
with pd.ExcelWriter(wb_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    gpa.to_excel(writer, sheet_name='gpa', index=False)

print("GPA sheet added successfully to d:/xls/class.xlsx!")

gpa.to_parquet('d:/wpswrkx/gpa.parquet', engine='pyarrow')
endsubmit;
run;quit;


/*--- ONLY USE PYTHON 310 TO COVERT PARQUET FILES TO SLC DATASETS ---*/
/*--- MORE RECENT VERSIONS OF PYTHON DO NOT SUPPORT SLC IMPOTT    ---*/
/*--- PY310 CREATE SAS DATASET FROM THE PARQUET FILE              ---*/


options set=PYTHONHOME "D:\py310";
proc python;
submit;
import pyarrow
import pandas as pd
class_paq = pd.read_parquet('d:/wpswrkx/gpa.parquet', engine='pyarrow')
print(class_paq)
endsubmit;
import python=class_paq data=workx.class_paq;
run;quit;

proc print data=workx.class_paq;
title "From python310";
run;quit;


/**************************************************************************************************************************/
/*  d:\xls\class.xlsx                      |                                                                              */
/* SHEET:GPA                               |   SLC DATTASET                                                               */
/* --------------------+                   |                                                                              */
/* | A1| fx    |STUDENT|                   |   WORKX.class_paq                                                            */
/* ------------------------------------    |                                                                              */
/* [_] |    A   |  B   |   C   |  D   |    |   Obs    STUDENT    MATH    MUSIC    MEAN                                    */
/* ------------------------------------    |                                                                              */
/*  1  | STUDENT| MATH | MUSIC | MEAN |    |    1      MARY       64       63     63.5                                    */
/*  -- |--------+------+-------+------+    |    2      ALEX       83       83     83.0                                    */
/*  2  |  MARY  |  64  |  63   | 63.5 |    |                                                                              */
/*  -- |--------+------+-------+------+    |                                                                              */
/*  3  |  ALEX  |  83  |  83   | 83   |    |                                                                              */
/*  -- |--------+------+-------+------+    |                                                                              */
/*  [GPA]                                  |                                                                              */
/**************************************************************************************************************************/

/*
| | ___   __ _
| |/ _ \ / _` |
| | (_) | (_| |
|_|\___/ \__, |
         |___/
*/

1                                          Altair SLC      18:41 Tuesday, February 17, 2026

NOTE: Copyright 2002-2025 World Programming, an Altair Company
NOTE: Altair SLC 2026 (05.26.01.00.000758)
      Licensed to Roger DeAngelis
NOTE: This session is executing on the X64_WIN11PRO platform and is running in 64 bit mode

NOTE: AUTOEXEC processing beginning; file is C:\wpsoto\autoexec.sas
NOTE: AUTOEXEC source line
1       +  ï»¿ods _all_ close;
           ^
ERROR: Expected a statement keyword : found "?"
NOTE: Library workx assigned as follows:
      Engine:        SAS7BDAT
      Physical Name: d:\wpswrkx

NOTE: Library slchelp assigned as follows:
      Engine:        WPD
      Physical Name: C:\Progra~1\Altair\SLC\2026\sashelp

NOTE: Library worksas assigned as follows:
      Engine:        SAS7BDAT
      Physical Name: d:\worksas

NOTE: Library workwpd assigned as follows:
      Engine:        WPD
      Physical Name: d:\workwpd


LOG:  18:41:46
NOTE: 1 record was written to file PRINT

NOTE: The data step took :
      real time : 0.040
      cpu time  : 0.015


NOTE: AUTOEXEC processing completed

1         options set=PYTHONHOME "D:\py314";
2         proc python;
3         submit;
4         import pyarrow
5         import pandas as pd
6         import sqlite3
7         from pandasql import sqldf
8         import pandas as pd
9
10        # Create connection and load sqlean.dll
11        conn = sqlite3.connect(":memory:")
12        conn.enable_load_extension(True)
13        conn.load_extension(r"d:/dll/sqlean.dll")
14
15        # Read Excel sheets
16        wb_file = r"d:/xls/class.xlsx"
17        math_df = pd.read_excel(wb_file, sheet_name="math")
18        music_df = pd.read_excel(wb_file, sheet_name="music")
19
20        # SQL query using pandasql (equivalent to sqldf with sqlean.dll)
21        # Note: pandasql uses SQLite under the hood, similar to your r sqldf setup
22        query = '''
23          SELECT
24             l.student,
25             l.grade as math,
26             r.grade as music,
27             (l.grade + r.grade)/2 as mean
28          FROM math_df as l
29          JOIN music_df as r
30          ON l.student = r.student
31        '''
32
33        gpa = sqldf(query, locals())
34        print(gpa)
35
36        # Write GPA to new sheet (preserves existing math/music sheets)
37        with pd.ExcelWriter(wb_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
38            gpa.to_excel(writer, sheet_name='gpa', index=False)
39
40        print("GPA sheet added successfully to d:/xls/class.xlsx!")
41
42        gpa.to_parquet('d:/wpswrkx/gpa.parquet', engine='pyarrow')
43        endsubmit;

NOTE: Submitting statements to Python:


44        run;quit;
NOTE: Procedure python step took :
      real time : 1.442
      cpu time  : 0.015


45
46
47        /*--- ONLY USE PYTHON 310 TO COVERT PARQUET FILES TO SLC DATASETS ---*/
48        /*--- MORE RECENT VERSIONS OF PYTHON DO NOT SUPPORT SLC IMPOTT    ---*/
49        /*--- PY310 CREATE SAS DATASET FROM THE PARQUET FILE              ---*/
50
51
52        options set=PYTHONHOME "D:\py310";
53        proc python;
54        submit;
55        import pyarrow
56        import pandas as pd
57        class_paq = pd.read_parquet('d:/wpswrkx/gpa.parquet', engine='pyarrow')
58        print(class_paq)
59        endsubmit;

NOTE: Submitting statements to Python:


60        import python=class_paq data=workx.class_paq;
NOTE: Creating data set 'WORKX.class_paq' from Python data frame 'class_paq'
NOTE: Data set "WORKX.class_paq" has 2 observation(s) and 4 variable(s)

61        run;quit;
NOTE: Procedure python step took :
      real time : 0.911
      cpu time  : 0.031


62
63        proc print data=workx.class_paq;
64        title "From python310";
65        run;quit;
NOTE: 2 observations were read from "WORKX.class_paq"
NOTE: Procedure print step took :
      real time : 0.021
      cpu time  : 0.000


66
67
ERROR: Error printed on page 1

NOTE: Submitted statements took :
      real time : 2.486
      cpu time  : 0.109

d:/dll/sqlean.dll
/**************************************************************************************************************************************/
/*   NOTE SOME OF THESE FUNCTION ARE ONLY AVAILABLE ON THE SQLITE COMMAND LINE (IE MODE)                                              */
/*               name          define_free    json_group_array      math_ceil           radians            stddev       time_fmt_iso  */
/*                 ->              degrees   json_group_object       math_cos            random        stddev_pop      time_fmt_time  */
/*                ->>           dense_rank         json_insert      math_cosh        randomblob       stddev_samp           time_get  */
/*                abs           difference         json_object   math_degrees              rank             stdev       time_get_day  */
/*               acos         dlevenshtein          json_patch       math_exp          readfile         strfilter      time_get_hour  */
/*              acosh                dur_h         json_pretty     math_floor            regexp          strftime   time_get_isoweek  */
/*                age                dur_m          json_quote        math_ln    regexp_capture        string_agg   time_get_isoyear  */
/*               asin               dur_ms         json_remove       math_log       regexp_like            strpos    time_get_minute  */
/*              asinh               dur_ns        json_replace     math_log10    regexp_replace            substr     time_get_month  */
/*               atan                dur_s            json_set      math_log2     regexp_substr         substring      time_get_nano  */
/*              atan2               dur_us           json_type       math_mod            repeat           subtype    time_get_second  */
/*              atanh        edit_distance          json_valid        math_pi           replace               sum   time_get_weekday  */
/*               atn2               encode               jsonb       math_pow         replicate           symlink      time_get_year  */
/*                avg                 eval         jsonb_array   math_radians           reverse               tan   time_get_yearday  */
/*         bit_length                  exp       jsonb_extract     math_round             right              tanh         time_micro  */
/*             blake3        fileio_append   jsonb_group_array       math_sin          rightstr      text_bitsize         time_milli  */
/*               bm25         fileio_mkdir  jsonb_group_object      math_sinh             round     text_casefold          time_nano  */
/*              btrim          fileio_mode        jsonb_insert      math_sqrt        row_number       text_concat           time_now  */
/*           casefold          fileio_read        jsonb_object       math_tan              rpad     text_contains         time_parse  */
/*         caverphone       fileio_symlink         jsonb_patch      math_tanh          rsoundex        text_count         time_round  */
/*               ceil         fileio_write        jsonb_remove     math_trunc        rtreecheck   text_has_prefix         time_since  */
/*            ceiling          first_value       jsonb_replace            max        rtreedepth   text_has_suffix           time_sub  */
/*            changes                floor           jsonb_set            md5         rtreenode        text_index      time_to_micro  */
/*               char               format           julianday         median             rtrim         text_join      time_to_milli  */
/*        char_length       fts3_tokenizer                 lag            min       script_code   text_last_index       time_to_nano  */
/*   character_length                 fts5   last_insert_rowid          mkdir              sha1         text_left       time_to_unix  */
/*          charindex      fts5_get_locale          last_value            mod            sha256       text_length         time_trunc  */
/*           coalesce          fts5_locale                lead          .mode            sha384         text_like          time_unix  */
/*             concat       fts5_source_id                left          nlike            sha512        text_lower         time_until  */
/*          concat_ws          fuzzy_caver             leftstr         nlower              sign         text_lpad           timediff  */
/*                cos         fuzzy_damlev              length            now               sin        text_ltrim       to_timestamp  */
/*               cosh       fuzzy_editdist         levenshtein      nth_value              sinh       text_repeat              total  */
/*                cot        fuzzy_hamming                like          ntile           snippet      text_replace      total_changes  */
/*               coth        fuzzy_jarowin          likelihood         nullif           soundex      text_reverse          translate  */
/*              count          fuzzy_leven              likely         nupper        split_part        text_right           translit  */
/*      crypto_blake3        fuzzy_osadist                  ln   octet_length    sqlean_version         text_rpad               trim  */
/*      crypto_decode       fuzzy_phonetic      load_extension        offsets        sqlite_log        text_rtrim              trunc  */
/*      crypto_encode       fuzzy_rsoundex                 log       optimize  sqlite_source_id         text_size             typeof  */
/*         crypto_md5         fuzzy_script               log10   osa_distance    sqlite_version        text_slice           unaccent  */
/*        crypto_sha1        fuzzy_soundex                log2           padc              sqrt        text_split           undefine  */
/*      crypto_sha256       fuzzy_translit               lower           padl            square    text_substring              unhex  */
/*      crypto_sha384      gen_random_uuid      lower_quartile           padr       starts_with        text_title            unicode  */
/*      crypto_sha512                 glob                lpad   percent_rank      stats_median    text_translate    unicode_version  */
/*          cume_dist         group_concat              lsmode     percentile         stats_p25         text_trim          unixepoch  */
/*       current_date              hamming               ltrim  percentile_25         stats_p75        text_upper           unlikely  */
/*       current_time                  hex           make_date  percentile_75         stats_p90              time              upper  */
/*  current_timestamp            highlight      make_timestamp  percentile_90         stats_p95          time_add     upper_quartile  */
/*               date               ifnull               match  percentile_95         stats_p99     time_add_date              uuid4  */
/*           date_add                  iif           matchinfo  percentile_99        stats_perc        time_after              uuid7  */
/*          date_part                instr           math_acos  phonetic_hash      stats_stddev       time_before         uuid_blob   */
/*         date_trunc         jaro_winkler          math_acosh             pi  stats_stddev_pop      time_compare          uuid_str   */
/*           datetime                 json           math_asin            pow stats_stddev_samp         time_date           var_pop   */
/*             decode           json_array          math_asinh          power         stats_var        time_equal          var_samp   */
/*             define    json_array_length           math_atan         printf     stats_var_pop     time_fmt_date          variance   */
/*       define_cache  json_error_position          math_atan2         proper    stats_var_samp time_fmt_datetime         writefile   */
/*       json_extract           math_atanh          quote                                                                  zeroblob   */
/**************************************************************************************************************************************/

/*              _
  ___ _ __   __| |
 / _ \ `_ \ / _` |
|  __/ | | | (_| |
 \___|_| |_|\__,_|

*/
