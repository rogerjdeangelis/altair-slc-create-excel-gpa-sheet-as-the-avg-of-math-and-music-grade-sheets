/* GPA sheet as the average of the MATH and MUSIC grade sheets.                */
/* Build the two grade tables from the same data as the workbook, then join    */
/* them on student name and average the two grades — the join + average logic  */
/* the proc r (sqldf) and proc python (pandasql) solutions perform, expressed  */
/* directly in PROC SQL so it runs as plain SAS.                               */

data math;
  input student $ grade;
cards4;
MARY 64
JOHN 70
MILE 83
ALEX 83
JOHN 93
;;;;
run;

data music;
  input student $ grade;
cards4;
JOSH 74
JANE 70
ALEX 83
MARY 63
MIKE 93
;;;;
run;

proc sql;
  create table gpa as
  select
     l.student
    ,l.grade as math
    ,r.grade as music
    ,(l.grade + r.grade)/2 as mean
  from
     math as l
  join
     music as r
  on
     l.student = r.student
  order by l.student;
quit;

proc print data=gpa noobs;
  title "GPA = average of MATH and MUSIC grades, joined on STUDENT";
run;
