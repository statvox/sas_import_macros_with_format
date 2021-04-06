/*******************************************************/
/*******************************************************/
/* Program name :  			macro_import_with_type.sas   			   */
/*******************************************************/
/* Author : 				Monica TURINICI				   */
/* Date of creation : 		01/01/2000				   */
/*******************************************************/


/* The macro takes as input
A/ a XLSX file named name.xlsx (note the 
absence of quotations '"')
B/ a dictionary xlsx file containing three columns, 
in this exact order
 1/ first column: the variable name
 2/ second column: the variable type: "char","date","num" 
 3/ the format column: the input,informat and format 
of each variable are set to this, e.g. $50, best12. and so on

Note: first line is considered to be used by the column names
and thus DISCARDED, do not use reference files that have
usable information in the first line.

Output: a SAS database with columns of desired formats and 
thus corresponding types is put in "outfile" 

Other input: 
- sheet_from_type : the name of the sheet
in the typefile from which to take the "informat" information
*/

/* the following macro takes an input SAS set
a variable from it and formats the variable to
the indicated format in place, that is
the new table will have the same name
and the new  variable the same name as before, only
the format is changed.
*/
%macro format_var_from_set(input_set=,oldvar=,newformat=);
	%put &newformat.;
	data tmp;
		set &input_set.;
		format newvar &newformat.;
		informat newvar &newformat.;
		if missing(input(&oldvar.,&newformat.)) then
			newvar=put(&oldvar.,&newformat.);
		else newvar=input(&oldvar.,&newformat.);	
	run;
	data tmp;
		set tmp;
		drop &oldvar.;
	run;
	
	data &input_set.;
		set tmp;
		rename newvar=&oldvar.;
	run;
%mend format_var_from_set;



%macro import_xlsx_with_format(input_file=,input_sheet=,
format_file=,format_sheet="formats",outfile=,);

	/* import input file */
	proc import datafile="&input_file." 
	DBMS=xlsx out=raw_input replace;
	sheet="&input_sheet."; GETNAMES=YES;
	run;

	/* use the "contents" procedure to obtain 
	a table of all variables names, in the input file 
	together with the position of each variable */
	proc contents data = raw_input
	 noprint out=data_info (keep = name varnum);
	run;

	/* import the file containing type/format information
	A = name of variable
	B = type
	C = format/informat/input
	*/
	proc import datafile="&format_file." 
	out=type_tmp(rename=A=name rename=B=var_type 
	rename=C=var_format) dbms=xlsx replace;
		RANGE="&format_sheet.$A2:0"; GETNAMES=NO;
	run;

	/* Merge the previous two: the new dataset 
	"merge_type_var" contains all the values 
	from the left table, plus matched values from the 
	right table or missing values in the case of no match.
	*/
	PROC SQL;
	CREATE TABLE merge_type_var0 AS
	SELECT data_info.*, type_tmp.var_type,type_tmp.var_format 
	FROM data_info LEFT JOIN type_tmp
	ON type_tmp.name=data_info.name ;
	QUIT;

	/* in case there are some missing values not present in reference, we set it
	as character with a large number of values */
	data merge_type_var;
	   set merge_type_var0;
	      if missing(var_format) then var_format='$250.';
	run;

	/* sort in order in which they apprear in the database */
	proc sort data=merge_type_var; by varnum; run;

	/* execute many times a call to the procedure that changes the format of a variable */
	data _null_;
	  set merge_type_var;
	  call execute('%nrstr(%format_var_from_set(input_set=raw_input,oldvar='||name||',newformat='||var_format||'))');
	run;

	data &outfile.;
		set raw_input;
	run;
%mend import_xlsx_with_format;


/* test macro */
/*
%import_xlsx_with_format(
input_file=myfile.xlsx,
input_sheet=mysheet,
format_file=formatfile.xlsx,
format_sheet=sheets,outfile=test);
*/


/*****************************************************/
/* The macro takes as input
A/ a input data file with delimiter=";" 
(can be changed using the macro argument delimiter=";" 
if necesary) 
B/ a dictionary xlsx file containing three columns, 
in this exact order
 1/ first column: the variable name
 2/ second column: the variable type: "char","date","num" 
 3/ the format column: the input,informat and format 
of each variable are set to this, e.g. $50, best12. and so on

Note: first line is considered to be used by the column names
and thus DISCARDED, do not use reference files that have
usable information in the first line.

Output: a SAS database with columns of desired types 
is put in "outfile" 

Other input: 
- sheet_from_type : the name of the sheet
in the typefile from which to take the "informat" information
- delimiter: if needed choose the delimited of the input file

Note : the input file can be a CSV or any text file
*/
%macro import_csv_with_format(input_file=,typefile=,
sheet_from_type="formats",outfile=,input_delimiter=";");

/* import the input file to obtain the list of 
variables; this will be used later */
options obs=2;/* only read two lines */
proc import datafile=&input_file. 
out=tmp_input_csv dbms=dlm replace;
delimiter=&input_delimiter.;
getnames=yes;
run;
options obs=max;
/* use the "contents" procedure to obtain 
a table of all variables names, in the input file 
together with the position of each variable */
proc contents data = tmp_input_csv
 noprint out=data_info (keep = name varnum);
run;

/* import the file containing type/format information
A will be the name of variable
B = type
C = format/informat/input
*/
proc import datafile=&typefile. 
out=type_tmp(rename=A=name rename=B=var_type rename=C=var_format) dbms=xlsx replace;
	RANGE="&sheet_from_type.$A2:0"; GETNAMES=NO;
run;


/* Merge the previous two: the new dataset 
"merge_type_var" contains all the values 
from the left table, plus matched values from the 
right table or missing values in the case of no match.
*/
PROC SQL;
CREATE TABLE merge_type_var0 AS
SELECT data_info.*, type_tmp.var_type,type_tmp.var_format 
FROM data_info LEFT JOIN type_tmp
ON type_tmp.name=data_info.name ;
QUIT;

/* in case there are some missing values not present in reference, we set it
as character with a large number of values */
data merge_type_var;
   set merge_type_var0;
      if missing(var_format) then var_format='$250.';
run;


/* sort in order in which they apprear in the database */
proc sort data=merge_type_var; by varnum; run;

/* proc contents merge_type_var; run;*/

/* now we can construct the input, informat and format statements */
/* format and informat will be the same */
proc sql noprint;
select cat(name,' ',var_format) into : format_statement 
separated by ' ' from merge_type_var ; 
quit;
/* input statement has ":" between objects */
proc sql noprint;
select cats(name,':',var_format) into : input_statement 
separated by " " from merge_type_var ; 
quit;
/* final step: import the file */
DATA &outfile.;
    INFILE &input_file.       
        TERMSTR=CRLF
        DLM=';' FIRSTOBS=2
        MISSOVER
        DSD ;
	FORMAT &format_statement.;
	INFORMAT &format_statement.;
	INPUT &input_statement. ;
RUN;
%mend import_csv_with_format;
       /* LRECL=3976
        TERMSTR=CRLF
        DLM=';'
        MISSOVER
        DSD ;*/
/*test macro:
%import_csv_with_format(input_file="myfile.csv",
typefile="formats_file.xlsx",
sheet_from_type=sheet1,outfile=test,input_delimiter=";");
*/

