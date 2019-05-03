;AHK_L v2.0 alpha 100-52515e2
;Tigerlily 5/1/19

#SingleInstance
SetWorkingDir( A_ScriptDir )
SetTitleMatchMode( 2 )

;		ACCOUNT DETAILS		;


AccountNameLong	:= "Universal Technical Institute"
AccountNameShort	:= "UTI"
AccountName		:= "UTI"

PreviousYear	:= 2018
CurrentYear	:= 2019
NextYear		:= 2020



;		DROP DOWN LIST FOR MONTH SELECTION	GUI		;
Gui := GuiCreate( "+AlwaysOnTop -SysMenu", "UTI Report Builder" )
Gui.Add( "Text", ,"Choose the month you want to report on:" )

Gui.Add( "DropDownList", "x22 y42 w250"
	, "January|February|March|April|May|June|July|August|September|October|November|December")
	.OnEvent( "Change", Func("OnSubmit"))

Gui.Show()

While ( ReportingMonth = "" )
	Sleep(100)

MsgBox(	Text		:= "A window will open up shortly that allows you to select " AccountName "'s previous month SEO Report to add/"
				.  "update with " ReportingMonth " data."
,		Title	:= "Please Wait...then open "
,		Options	:=  "0x40040 T10" ) ; System Modal (always on top) with (i) info icon - 10 second timeout

;		START NEW EXCEL INSTANCE		;
Xl := ComObjCreate( "Excel.Application" )


;		USER SELECT FILE TO OPEN		;
fileToOpen := Xl.GetOpenFilename(  FileFilter	:= "Excel Files [*.xls*], *.xls*"
						,	FilterIndex	:=  1
						,	Title		:= "OPEN " AccountName " MONTHLY REPORT"
						,	MultiSelect	:=  False )

if ( fileToOpen != 0 )
{		
		Xl.Workbooks.Open( fileToOpen, UpdateLinks := 0 )
		Xl.Visible := true
		Xl_pID := WinGetPID("ahk_id " Xl.hWnd)
		WinMaximize("ahk_id " Xl.hWnd)
}
else if ( fileToOpen = 0 )
{
	SoundPlay( "*-1" )
	
	MsgBox(	Text		:= "No file was selected.`n`n" AccountName " SEO Monthly Report Builder will close."
	,		Title	:= "REPORT NOT OPENED"
	,		Options	:=  0x40010) ; System Modal (always on top) with STOP hand icon
	
	Xl := ""	
	ExitApp
}

SplitPath( fileToOpen
	    , WBFileName
	    , WBdir
	    , WBext
	    , WBnameNoExt
	    , WBdrive )

Xl.DisplayAlerts := False


;		REPORTING WORKSHEETS		;
wsEXEC	:= Xl.Worksheets( "EXECUTIVE SUMMARY" )
wsMILE	:= Xl.Worksheets( "SITE MILESTONES" )

wsTRAFFIC	:= Xl.Worksheets( "TOTAL TRAFFIC REPORT" )
wsAUTO	:= Xl.Worksheets( "AUTOMOTIVE" )
wsDIESEL	:= Xl.Worksheets( "DIESEL" )
wsMOTO	:= Xl.Worksheets( "MOTORCYCLE" )
wsMARINE	:= Xl.Worksheets( "MARINE" )
wsBLOG	:= Xl.Worksheets( "BLOG" )
wsPOS	:= Xl.Worksheets( "POSITION REPORT" )

;		DATA WORKSHEETS		;
wsGA		:= Xl.Worksheets( "Org Entrance & Goal LPs (All)" )
wsGAm	:= Xl.Worksheets( "Org Entran & Goals LPs (Mobile)" )
wsSTAT	:= Xl.Worksheets( "STAT Keywords" )

;		DATA WORKBOOK FILES		;
wb_GA	:= "GA.csv"
wb_GAm	:= "GAm.csv"
wb_STAT	:= "STAT.csv"


;		UPDATE HEADERS	& MONTHS		;
wsEXEC.Range( "B2:L2" ).Value		:= ReportingMonth_CAPS " 2019 / EXECUTIVE SUMMARY / " StrUpper( AccountNameLong )
wsMILE.Range( "B2:D2" ).Value		:= ReportingMonth_CAPS " 2019 / SEO MILESTONES / "    StrUpper( AccountNameLong )
wsTRAFFIC.Range( "B2:P2" ).Value	:= ReportingMonth_CAPS " 2019 / ENTRANCES REPORT / "  StrUpper( AccountNameLong ) " / TOTAL"
wsAUTO.Range( "B2:P2" ).Value		:= ReportingMonth_CAPS " 2019 / ENTRANCES REPORT / "  StrUpper( AccountNameLong ) " / AUTOMOTIVE"
wsDIESEL.Range( "B2:P2" ).Value	:= ReportingMonth_CAPS " 2019 / ENTRANCES REPORT / "  StrUpper( AccountNameLong ) " / DIESEL"
wsMOTO.Range( "B2:P2" ).Value		:= ReportingMonth_CAPS " 2019 / ENTRANCES REPORT / "  StrUpper( AccountNameLong ) " / MOTORCYCLE"
wsMARINE.Range( "B2:P2" ).Value	:= ReportingMonth_CAPS " 2019 / ENTRANCES REPORT / "  StrUpper( AccountNameLong ) " / MARINE"
wsBLOG.Range( "B2:P2" ).Value		:= ReportingMonth_CAPS " 2019 / ENTRANCES REPORT / "  StrUpper( AccountNameLong ) " / BLOG"
wsPOS.Range( "B2:L2" ).Value		:= ReportingMonth_CAPS " 2019 / POSITION REPORT / "   StrUpper( AccountNameLong ) 

wsTRAFFIC.Range( "H159" ).Value	:= "'" ReportingMonth_CAPS " " CurrentYear
wsAUTO.Range( "H184" ).Value		:= "'" ReportingMonth_CAPS " " CurrentYear
wsDIESEL.Range( "H184" ).Value	:= "'" ReportingMonth_CAPS " " CurrentYear
wsMOTO.Range( "H184" ).Value		:= "'" ReportingMonth_CAPS " " CurrentYear
wsMARINE.Range( "H184" ).Value	:= "'" ReportingMonth_CAPS " " CurrentYear
wsBLOG.Range( "H184" ).Value		:= "'" ReportingMonth_CAPS " " CurrentYear


;		TRANSFER GA DATA (ALL)		;
Try Xl.Workbooks.Open( FileName	:= WBdir "\" wb_GA
				 , UpdateLinks	:= 0 )
catch 
{
	SoundPlay( "*-1" )
	
	MsgBox(	Text		:=	 wb_GA " file could not be found in the same folder as the report you opened.`n`n"
					.	"Please move the " wb_GA " data export file into`n`n"
					.	 WBDir "`n`n"
					.	 AccountName " SEO Monthly Report Builder will now close."
	,		Title	:=	"REPORT CREATION FAILURE - UNABLE TO FIND DATA FILE"
	,		Options	:=	0x1010 )
	try Xl.Quit
	Xl := ""
	ExitApp
}


Xl.ActiveSheet.Rows( "1:6" ).Delete ; Delete Top 6 rows

Xl.ActiveSheet.Range( "A1" ).Select
Xl.ActiveCell.CurrentRegion.Select 

LastRow := Xl.Selection.Rows.Count
Xl.ActiveSheet.Rows( LastRow ).Delete ; Delete last row of table

Xl.ActiveSheet.Range( "A1" ).Select
Xl.ActiveCell.CurrentRegion.Select

Xl.ActiveSheet.ListObjects.Add( 1, Xl.Range( Xl.Selection.Address ), , 0 ).Name := "Table1" ; Create table 


TableRangeValues	:= Xl.Selection.Value
TableRange		:= Xl.Selection.Address[0,0]
TableRows			:= Xl.Selection.Rows.Count
TableColumns		:= Xl.Selection.Columns.Count


;		STORE DATA TO GO TO GA TAB		;

GA_srcData := Xl.ActiveSheet.Range( "A2:D" TableRows ).Value


;		GET TOP 10 URLs - All		;
fromWB			:=  wb_GA
toWB				:=  WBnameNoExt
toWS				:= "TOTAL TRAFFIC REPORT"
toRange_URL		:= "B162:B171"
toRange_Entrances	:= "H162:H171"

Top10_URL_All		:= Xl.ActiveSheet.Range( "A2:A11" ).Value
Top10_Entrances_All	:= Xl.ActiveSheet.Range( "C2:C11" ).Value

Xl.Windows( toWB ).Activate
Xl.Workbooks( toWB ).Activate	
Xl.Worksheets( toWS ).Activate	

Xl.ActiveSheet.Range( toRange_URL ).Value		:= Top10_URL_All
Xl.ActiveSheet.Range( toRange_Entrances ).Value	:= Top10_Entrances_All


;		TRANSFER GA ENTRANCES & GOAL COMPLETION DATA		;
wsGA.Activate

Xl.ActiveSheet.Range( "B5" ).Select
Xl.ActiveCell.CurrentRegion.Select 

GA_Table_Address		:= Xl.Selection.Address[0,0]
GA_Table_Rows			:= Xl.Selection.Rows.Count
Table_LastRow_Address	:= StrSplit( GA_Table_Address, "J" )[2]

NewData_GA_Address		:= "B" ( Table_LastRow_Address + 1 ) ":E" ( Table_LastRow_Address + TableRows - 1 )

Xl.ActiveSheet.Range( NewData_GA_Address ).Value := GA_srcData
Xl.ActiveSheet.Range( NewData_GA_Address ).NumberFormat := "General"

;		GET TOP 10 URLs from each CATEGORY		;

ContainsString		:= "automotive"
fromWB			:=  wb_GA
toWB				:=  WBnameNoExt
toWS				:= "AUTOMOTIVE"
toRange_URL		:= "B187:B196"
toRange_Entrances	:= "H187:H196"

Top10_URL( ContainsString, fromWB, toWB, toWS, toRange_URL, toRange_Entrances )


ContainsString		:= "diesel"
fromWB			:=  wb_GA
toWB				:=  WBnameNoExt
toWS				:= "DIESEL"
toRange_URL		:= "B187:B196"
toRange_Entrances	:= "H187:H196"

Top10_URL( ContainsString, fromWB, toWB, toWS, toRange_URL, toRange_Entrances )


ContainsString		:= "motorcycle"
fromWB			:=  wb_GA
toWB				:=  WBnameNoExt
toWS				:= "MOTORCYCLE"
toRange_URL		:= "B187:B196"
toRange_Entrances	:= "H187:H196"

Top10_URL( ContainsString, fromWB, toWB, toWS, toRange_URL, toRange_Entrances )


ContainsString		:= "marine"
fromWB			:=  wb_GA
toWB				:=  WBnameNoExt
toWS				:= "MARINE"
toRange_URL		:= "B187:B196"
toRange_Entrances	:= "H187:H196"


Top10_URL( ContainsString, fromWB, toWB, toWS, toRange_URL, toRange_Entrances )


ContainsString		:= "blog"
fromWB			:=  wb_GA
toWB				:=  WBnameNoExt
toWS				:= "BLOG"
toRange_URL		:= "B187:B196"
toRange_Entrances	:= "H187:H196"

Top10_URL( ContainsString, fromWB, toWB, toWS, toRange_URL, toRange_Entrances )

Xl.Windows(fromWB).Activate
Xl.Workbooks(fromWB).Activate
Xl.Workbooks(fromWB).Close


;		TRANSFER GA DATA (MOBILE)		;
Try Xl.Workbooks.Open( FileName	:= WBdir "\" wb_GAm
				 , UpdateLinks	:= 0 )
catch 
{
	SoundPlay( "*-1" )
	
	MsgBox(	Text		:=	 wb_GAm " file could not be found in the same folder as the report you opened.`n`n"
					.	"Please move the " wb_GAm " data export file into`n`n"
					.	 WBDir "`n`n"
					.	 AccountName " SEO Monthly Report Builder will now close."
	,		Title	:=	"REPORT CREATION FAILURE - UNABLE TO FIND DATA FILE"
	,		Options	:=	0x1010)
	try Xl.Quit
	Xl := ""
	ExitApp
}


Xl.ActiveSheet.Rows( "1:6" ).Delete ; Delete Top 6 rows

Xl.ActiveSheet.Range( "A1" ).Select
Xl.ActiveCell.CurrentRegion.Select 

LastRow := Xl.Selection.Rows.Count
Xl.ActiveSheet.Rows( LastRow ).Delete ; Delete last row of table

Xl.ActiveSheet.Range( "A1" ).Select
Xl.ActiveCell.CurrentRegion.Select

Xl.ActiveSheet.ListObjects.Add( 1, Xl.Range( Xl.Selection.Address ), , 0 ).Name := "Table1" ; Create table 


TableRangeValues	:= Xl.Selection.Value
TableRange		:= Xl.Selection.Address[0,0]
TableRows			:= Xl.Selection.Rows.Count
TableColumns		:= Xl.Selection.Columns.Count


;		STORE DATA TO GO TO GA MOBILE TAB		;

GAm_srcData := Xl.ActiveSheet.Range( "A2:D" TableRows ).Value


fromWB	:=  wb_GAm
toWB		:=  WBnameNoExt

Xl.Windows( toWB ).Activate
Xl.Workbooks( toWB ).Activate		


;		TRANSFER GA MOBILE ENTRANCES & GOAL COMPLETION DATA		;
wsGAm.Activate

Xl.ActiveSheet.Range( "B5" ).Select
Xl.ActiveCell.CurrentRegion.Select 

GA_Table_Address		:= Xl.Selection.Address[0,0]
GA_Table_Rows			:= Xl.Selection.Rows.Count
Table_LastRow_Address	:= StrSplit( GA_Table_Address, "J" )[2]

NewData_GA_Address		:= "B" ( Table_LastRow_Address + 1 ) ":E" ( Table_LastRow_Address + TableRows - 1 )


Xl.ActiveSheet.Range( NewData_GA_Address ).Value			:= GAm_srcData
Xl.ActiveSheet.Range( NewData_GA_Address ).NumberFormat	:= "General"

Xl.Windows( fromWB ).Activate
Xl.Workbooks( fromWB ).Activate
Xl.Workbooks( fromWB ).Close


;		UPDATE FORMAT & NUMBERS		;

;		TRAFFIC Tab		;
tab 		:= wsTRAFFIC

FIRSTrow_TRAFFIC	:= 94
JANrow_TRAFFIC		:= 111
LASTrow_TRAFFIC	:= 151

PREVrow_TRAFFIC	:= ( JANrow_TRAFFIC + ReportingMonthDate - 2 )
CURRENTrow_TRAFFIC	:= ( JANrow_TRAFFIC + ReportingMonthDate - 1 )
NEXTrow_TRAFFIC	:= ( JANrow_TRAFFIC + ReportingMonthDate     )

PREVYEARrow_TRAFFIC	:= ( CURRENTrow_TRAFFIC - 12 )

tab.Activate

tab.Rows( FIRSTrow_TRAFFIC	":"  LASTrow_TRAFFIC	).EntireRow.Hidden := True
tab.Rows( PREVYEARrow_TRAFFIC	":"  CURRENTrow_TRAFFIC	).EntireRow.Hidden := False


;		AUTO Tab		;
tab 		:= wsAUTO

FIRSTrow_AUTO	:= 119
JANrow_AUTO	:= 143
LASTrow_AUTO	:= 176

PREVrow_AUTO		:= ( JANrow_AUTO + ReportingMonthDate - 2 )
CURRENTrow_AUTO	:= ( JANrow_AUTO + ReportingMonthDate - 1 )
NEXTrow_AUTO		:= ( JANrow_AUTO + ReportingMonthDate     )

PREVYEARrow_AUTO	:= ( CURRENTrow_AUTO - 12 )

tab.Activate

tab.Rows( FIRSTrow_AUTO		":"  LASTrow_AUTO	 ).EntireRow.Hidden := True
tab.Rows( PREVYEARrow_AUTO	":"  CURRENTrow_AUTO ).EntireRow.Hidden := False


;		DIESEL Tab		;
tab 		:= wsDIESEL

FIRSTrow_DIESEL	:= 119
JANrow_DIESEL		:= 143
LASTrow_DIESEL		:= 176

PREVrow_DIESEL		:= ( JANrow_DIESEL + ReportingMonthDate - 2 )
CURRENTrow_DIESEL	:= ( JANrow_DIESEL + ReportingMonthDate - 1 )
NEXTrow_DIESEL		:= ( JANrow_DIESEL + ReportingMonthDate     )

PREVYEARrow_DIESEL	:= ( CURRENTrow_DIESEL - 12 )

tab.Activate

tab.Rows( FIRSTrow_DIESEL	":"  LASTrow_DIESEL		).EntireRow.Hidden := True
tab.Rows( PREVYEARrow_DIESEL	":"  CURRENTrow_DIESEL	).EntireRow.Hidden := False


;		MOTO Tab		;
tab 		:= wsMOTO

FIRSTrow_MOTO		:= 119
JANrow_MOTO		:= 143
LASTrow_MOTO		:= 176

PREVrow_MOTO		:= ( JANrow_MOTO + ReportingMonthDate - 2 )
CURRENTrow_MOTO	:= ( JANrow_MOTO + ReportingMonthDate - 1 )
NEXTrow_MOTO		:= ( JANrow_MOTO + ReportingMonthDate     )

PREVYEARrow_MOTO	:= ( CURRENTrow_MOTO - 12 )

tab.Activate

tab.Rows( FIRSTrow_MOTO		":"  LASTrow_MOTO	 ).EntireRow.Hidden := True
tab.Rows( PREVYEARrow_MOTO	":"  CURRENTrow_MOTO ).EntireRow.Hidden := False


;		MARINE Tab		;
tab 		:= wsMARINE

FIRSTrow_MARINE	:= 119
JANrow_MARINE		:= 143
LASTrow_MARINE		:= 176

PREVrow_MARINE		:= ( JANrow_MARINE + ReportingMonthDate - 2 )
CURRENTrow_MARINE	:= ( JANrow_MARINE + ReportingMonthDate - 1 )
NEXTrow_MARINE		:= ( JANrow_MARINE + ReportingMonthDate     )

PREVYEARrow_MARINE	:= ( CURRENTrow_MARINE - 12 )

tab.Activate

tab.Rows( FIRSTrow_MARINE	":"  LASTrow_MARINE		).EntireRow.Hidden := True
tab.Rows( PREVYEARrow_MARINE	":"  CURRENTrow_MARINE	).EntireRow.Hidden := False


;		BLOG Tab		;
tab 		:= wsBLOG

FIRSTrow_BLOG		:= 119
JANrow_BLOG		:= 143
LASTrow_BLOG		:= 176

PREVrow_BLOG		:= ( JANrow_BLOG + ReportingMonthDate - 2 )
CURRENTrow_BLOG	:= ( JANrow_BLOG + ReportingMonthDate - 1 )
NEXTrow_BLOG		:= ( JANrow_BLOG + ReportingMonthDate     )

PREVYEARrow_BLOG	:= ( CURRENTrow_BLOG - 12 )

tab.Activate

tab.Rows( FIRSTrow_BLOG		":"  LASTrow_BLOG	 ).EntireRow.Hidden := True
tab.Rows( PREVYEARrow_BLOG	":"  CURRENTrow_BLOG ).EntireRow.Hidden := False


;		POSITION REPORT Tab		;
tab := wsPOS


;		TABLE: BRAND & NONBRAND KEYWORDS - GOOGLE MOBILE SEARCH 		;
FIRSTrow_BNB1		:= 132
JANrow_BNB1		:= 133
LASTrow_BNB1		:= 175

PREVrow_BNB1		:= ( JANrow_BNB1 + ReportingMonthDate - 2 )
CURRENTrow_BNB1	:= ( JANrow_BNB1 + ReportingMonthDate - 1 )
NEXTrow_BNB1		:= ( JANrow_BNB1 + ReportingMonthDate     )

PREVYEARrow_BNB1	:= ( CURRENTrow_BNB1 - 12 )

tab.Activate

tab.Rows( FIRSTrow_BNB1	":"  LASTrow_BNB1	 ).EntireRow.Hidden := True
tab.Rows( JANrow_BNB1	":"  CURRENTrow_BNB1 ).EntireRow.Hidden := False

tab.Range( "C" PREVrow_BNB1 ":F" PREVrow_BNB1 ).Value := tab.Range( "C" PREVrow_BNB1 ":F" PREVrow_BNB1 ).Value


;		TABLE: AVERAGE GOOGLE RANK TRENDS BY CATEGORY (NON-BRANDED)		;
FIRSTrow_BNB2		:= 243
JANrow_BNB2		:= 244
LASTrow_BNB2		:= 285

PREVrow_BNB2		:= ( JANrow_BNB2 + ReportingMonthDate - 2 )
CURRENTrow_BNB2	:= ( JANrow_BNB2 + ReportingMonthDate - 1 )
NEXTrow_BNB2		:= ( JANrow_BNB2 + ReportingMonthDate     )

PREVYEARrow_BNB2	:= ( CURRENTrow_BNB2 - 12 )

tab.Rows( FIRSTrow_BNB2	":"  LASTrow_BNB2	 ).EntireRow.Hidden := True
tab.Rows( JANrow_BNB2	":"  CURRENTrow_BNB2 ).EntireRow.Hidden := False

tab.Range( "C" PREVrow_BNB2 ":F" PREVrow_BNB2 ).Value := tab.Range( "C" PREVrow_BNB2 ":F" PREVrow_BNB2 ).Value


;		CAPTURE LAST MONTH'S STAT DATA		;

PREVIOUS_Page1_UTI_Keywords := Round( wsPOS.Range( "C37" ).Value )
PREVIOUS_Page1to3_UTI_Keywords := Round( wsPOS.Range( "F37" ).Value )

;		TRANSFER STAT DATA		;
tab := wsSTAT

tab.Activate

;		DELETE OLD STAT DATA PREVIOUS MONTH TABLE	;
tab.Range( "AB5" ).Select
Xl.ActiveCell.CurrentRegion.Select 
Xl.Selection.Offset( 2, 0 ).Select
Xl.Selection.Clear

;		MOVE OLD STAT DATA TO PREVIOUS MONTH TABLE		;
tab.Range( "B5" ).Select
Xl.ActiveCell.CurrentRegion.Select 

TableRange_STAT		:= Xl.Selection.Address[0,0]
TableRows_STAT			:= Xl.Selection.Rows.Count
TableColumns_STAT		:= Xl.Selection.Columns.Count


STAT_Table_Address 			:= Xl.Selection.Address[0,0]

STAT_Table_LastRow_Address	:= StrSplit( STAT_Table_Address, "Z" )[2]
STAT_Table_Rows			:= Xl.Selection.Rows.Count

STAT_DataForPrevMonth 		:= tab.Range( "B5:P" STAT_Table_LastRow_Address ).Value

OldData_STAT_Address		:= "AB5:AP" STAT_Table_LastRow_Address

tab.Range( OldData_STAT_Address ).Value := STAT_DataForPrevMonth

tab.Range( "AQ5:AZ5" ).AutoFill( tab.Range( "AQ5:AZ" STAT_Table_LastRow_Address ) )

;		FIX ANNOYING NUMBER FORMAT ISSUE		;
tab.Range( "AD5:AP" STAT_Table_LastRow_Address ).NumberFormat := "General"
tab.Range( "AB5:AC" STAT_Table_LastRow_Address ).NumberFormat := "m/d/yyyy"

;		DELETE OLD STAT DATA CURRENT MONTH TABLE	;
tab.Range( "B5" ).Select
Xl.ActiveCell.CurrentRegion.Select 
Xl.Selection.Offset( 2, 0 ).Select
Xl.Selection.Clear

;		TRANSFER STAT DATA		;
Try Xl.Workbooks.Open( FileName	:= WBdir "\" wb_STAT
				 , UpdateLinks	:= 0 )
catch 
{
	SoundPlay( "*-1" )
	
	MsgBox(	Text		:=	 wb_STAT " file could not be found in the same folder as the report you opened.`n`n"
					.	"Please move the " wb_STAT " data export file into`n`n"
					.	 WBDir "`n`n"
					.	 AccountName " SEO Monthly Report Builder will now close."
	,		Title	:=	"REPORT CREATION FAILURE - UNABLE TO FIND DATA FILE"
	,		Options	:=	0x1010 )
	try Xl.Quit
	Xl := ""
	ExitApp
}

;		STORE DATA TO GO TO GA TAB		;


Xl.ActiveSheet.Range( "A1" ).Select
Xl.ActiveCell.CurrentRegion.Select 

TableRows_STAT	:= Xl.Selection.Rows.Count

STAT_srcData	:= Xl.ActiveSheet.Range( "A2:O" TableRows_STAT ).Value

Xl.ActiveSheet.Range( "A1" ).Select
Xl.ActiveCell.CurrentRegion.Select

Xl.ActiveSheet.ListObjects.Add( 1, Xl.Range( Xl.Selection.Address ), , 0 ).Name := "Table1" ; Create table 
Xl.Selection.RemoveDuplicates( 4,1 )

Xl.ActiveSheet.Range( "A1" ).Select
Xl.ActiveCell.CurrentRegion.Select

TotalTrackedKeywords := ( Xl.Selection.Rows.Count - 1 )

Xl.ActiveWorkbook.Close

tab.Range( "B5:P" ( TableRows_STAT + 3 ) ).Value := STAT_srcData

;		FIX ANNOYING NUMBER FORMAT ISSUE		;
tab.Range( "D5:P" ( TableRows_STAT + 3 ) ).NumberFormat := "General"
tab.Range( "B5:C" ( TableRows_STAT + 3 ) ).NumberFormat := "m/d/yyyy"

tab.Range( "Q5:Z5" ).AutoFill( tab.Range( "Q5:Z" STAT_Table_LastRow_Address ) )

;		SUMMARIES		;


;		TRAFFIC REPORT		;
tab 		:= wsTRAFFIC

FIRSTrow_TRAFFIC1	:= 43
LASTrow_TRAFFIC1	:= 54

PREVrow_TRAFFIC1	:= ( FIRSTrow_TRAFFIC1 + ReportingMonthDate - 2 )
CURRENTrow_TRAFFIC1	:= ( FIRSTrow_TRAFFIC1 + ReportingMonthDate - 1 )
NEXTrow_TRAFFIC1	:= ( FIRSTrow_TRAFFIC1 + ReportingMonthDate     )

tab.Activate

CurrentValue	:= tab.Range( "E" CURRENTrow_TRAFFIC1 ).Value
PreviousValue	:= tab.Range( "E" PREVrow_TRAFFIC1 ).Value

PercentChange( CurrentValue, PreviousValue )

ChangePercentage_TRAFFIC1_EntrMOM	:= ChangePercentage
Incr_or_Decr_TRAFFIC1_EntrMOM		:= Incr_or_Decr
Impr_or_Decl_TRAFFIC1_EntrMOM		:= Impr_or_Decl
Up_or_Down_TRAFFIC1_EntrMOM		:= Up_or_Down
Plus_or_Minus_TRAFFIC1_EntrMOM	:= Plus_or_Minus

CurrentValue	:= tab.Range( "E" CURRENTrow_TRAFFIC1 ).Value
PreviousValue	:= tab.Range( "D" CURRENTrow_TRAFFIC1 ).Value

PercentChange( CurrentValue, PreviousValue )

ChangePercentage_TRAFFIC1_EntrYOY	:= ChangePercentage
Incr_or_Decr_TRAFFIC1_EntrYOY		:= Incr_or_Decr
Impr_or_Decl_TRAFFIC1_EntrYOY		:= Impr_or_Decl
Up_or_Down_TRAFFIC1_EntrYOY		:= Up_or_Down
Plus_or_Minus_TRAFFIC1_EntrYOY	:= Plus_or_Minus


CurrentValue	:= tab.Range( "J" CURRENTrow_TRAFFIC1 ).Value
PreviousValue	:= tab.Range( "J" PREVrow_TRAFFIC1 ).Value

PercentChange( CurrentValue, PreviousValue )

ChangePercentage_TRAFFIC1_InqrMOM	:= ChangePercentage
Incr_or_Decr_TRAFFIC1_InqrMOM		:= Incr_or_Decr
Impr_or_Decl_TRAFFIC1_InqrMOM		:= Impr_or_Decl
Up_or_Down_TRAFFIC1_InqrMOM		:= Up_or_Down
Plus_or_Minus_TRAFFIC1_InqrMOM	:= Plus_or_Minus

CurrentValue	:= tab.Range( "J" CURRENTrow_TRAFFIC1 ).Value
PreviousValue	:= tab.Range( "I" CURRENTrow_TRAFFIC1 ).Value

PercentChange( CurrentValue, PreviousValue )

ChangePercentage_TRAFFIC1_InqrYOY	:= ChangePercentage
Incr_or_Decr_TRAFFIC1_InqrYOY		:= Incr_or_Decr
Impr_or_Decl_TRAFFIC1_InqrYOY		:= Impr_or_Decl
Up_or_Down_TRAFFIC1_InqrYOY		:= Up_or_Down
Plus_or_Minus_TRAFFIC1_InqrYOY	:= Plus_or_Minus

wsTRAFFIC.Range( "B38" ).Value :=
(
"UTI’s Total Organic Entrances " Incr_or_Decr_TRAFFIC1_EntrMOM "d " Plus_or_Minus_TRAFFIC1_EntrMOM . ChangePercentage_TRAFFIC1_EntrMOM " MoM, and " Incr_or_Decr_TRAFFIC1_EntrYOY "d " Plus_or_Minus_TRAFFIC1_EntrYOY . ChangePercentage_TRAFFIC1_EntrYOY " YoY.
UTI’s Total Organic Inquiries " Incr_or_Decr_TRAFFIC1_InqrMOM "d " Plus_or_Minus_TRAFFIC1_InqrMOM . ChangePercentage_TRAFFIC1_InqrMOM " MoM, and " Incr_or_Decr_TRAFFIC1_InqrYOY "d " Plus_or_Minus_TRAFFIC1_InqrYOY . ChangePercentage_TRAFFIC1_InqrYOY " YoY.

Location pages helped improve traffic YoY due to an increase in the total amount of ranking keywords for location related phrases (ex: trade school in long beach, technical school Houston, mechanic schools). 
When we look at the average rank for the Location pages YoY, average page rank decreased, but the total number of ranking keywords those pages grew because the keyword profile diversified to include more nonbranded terms, which led to more organic clicks over time.
This tells us creating additional, informational Location and Campus based pages, like what is planned for Q3, will help increase incremental traffic by helping widen the range of keywords ranking for those pages.
To further improve Location Page performance, iProspect recommends doing an audit of Location Page meta data to help improve CTR and position, as well as identify any Google My Business Location pages for areas of opportunity towards the end of March."
)

FIRSTrow_TRAFFIC	:= 94
JANrow_TRAFFIC		:= 111
LASTrow_TRAFFIC	:= 151

PREVrow_TRAFFIC	:= ( JANrow_TRAFFIC + ReportingMonthDate - 2 )
CURRENTrow_TRAFFIC	:= ( JANrow_TRAFFIC + ReportingMonthDate - 1 )
NEXTrow_TRAFFIC	:= ( JANrow_TRAFFIC + ReportingMonthDate     )

PREVYEARrow_TRAFFIC	:= ( CURRENTrow_TRAFFIC - 12 )

CurrentValue	:= tab.Range( "C" CURRENTrow_TRAFFIC ).Value
PreviousValue	:= tab.Range( "C" PREVrow_TRAFFIC ).Value

PercentChange( CurrentValue, PreviousValue )

ChangePercentage_TRAFFICm_InqrMOM	:= ChangePercentage
Incr_or_Decr_TRAFFICm_InqrMOM		:= Incr_or_Decr
Impr_or_Decl_TRAFFICm_InqrMOM		:= Impr_or_Decl
Up_or_Down_TRAFFICm_InqrMOM		:= Up_or_Down
Plus_or_Minus_TRAFFICm_InqrMOM	:= Plus_or_Minus

CurrentValue	:= tab.Range( "C" CURRENTrow_TRAFFIC ).Value
PreviousValue	:= tab.Range( "C" PREVYEARrow_TRAFFIC ).Value

PercentChange( CurrentValue, PreviousValue )

ChangePercentage_TRAFFICm_InqrYOY	:= ChangePercentage
Incr_or_Decr_TRAFFICm_InqrYOY		:= Incr_or_Decr
Impr_or_Decl_TRAFFICm_InqrYOY		:= Impr_or_Decl
Up_or_Down_TRAFFICm_InqrYOY		:= Up_or_Down
Plus_or_Minus_TRAFFICm_InqrYOY	:= Plus_or_Minus

CURRENTPercent_AllOrganicEntr := ( tab.Range( "F" CURRENTrow_TRAFFIC	).Value * 100 )
PREVPercent_AllOrganicEntr	:= ( tab.Range( "F" PREVrow_TRAFFIC	).Value * 100 )

ChangePercentage := CURRENTPercent_AllOrganicEntr - PREVPercent_AllOrganicEntr

if ( ChangePercentage > 0 )
	{
		Up_or_Down	:= "up"
		Plus_or_Minus	:= "+"
	}
	else if ( ChangePercentage < 0 )
	{	
		Up_or_Down	:= "down"
		Plus_or_Minus	:= "-"
	}
	else if ( ChangePercentage = 0 )
	{
		Up_or_Down	:= ""
		Plus_or_Minus	:= ""
	}

CURRENTPercent_AllOrganicEntr	:= Round( CURRENTPercent_AllOrganicEntr, 1 ) "%"

ChangePercentage			:= Round( Abs( Float( ChangePercentage ) ), 1 ) "%"

wsTRAFFIC.Range( "B89" ).Value :=
(
"UTI's Total Organic Mobile Entrances " Incr_or_Decr_TRAFFICm_InqrMOM "d by " Plus_or_Minus_TRAFFICm_InqrMOM . ChangePercentage_TRAFFICm_InqrMOM " MoM, and " Incr_or_Decr_TRAFFICm_InqrYOY "d by " Plus_or_Minus_TRAFFICm_InqrYOY . ChangePercentage_TRAFFICm_InqrYOY " YoY.

- Organic Mobile Entrances was " CURRENTPercent_AllOrganicEntr " of all Organic Entrances in " ReportingMonth ", " Up_or_Down " " Plus_or_Minus . ChangePercentage " from " PREV_ReportingMonth ".
- Monitor mobile Entrances distribution for major fluctuations to determine if slow page speed is a contributing factor."
)




;		POSITION REPORT		;
tab := wsPOS

Page1_PercentIncr_MOM := tab.Range( "H37" ).Value

if ( Page1_PercentIncr_MOM > 0 )
	{
		Incr_or_Decr			:= "increase"
		Plus_or_MinusPage1_MOM	:= "+"
	}
	else if ( Page1_PercentIncr_MOM < 0 )
	{
		Incr_or_Decr			:= "decrease"
		Plus_or_MinusPage1_MOM	:= "-"
	}
	else if ( Page1_PercentIncr_MOM = 0 )
	{
		Incr_or_Decr			:= ""
		Plus_or_MinusPage1_MOM	:= ""
	}
	
Page1_PercentIncr_MOM := ( Round( ( ( tab.Range( "H37" ).Value ) * 100 ), 1) ) "%"




wsPOS.Range( "B32" ).Value :=
(
"SUMMARY:
There are a total number of " TotalTrackedKeywords " branded and non-branded keywords being tracked for mobile devices.
Any keyword data within the monthly report is reflective of the total number of tracked brand and non-brand tracked keywords.
Keeping a relatively static data set is crucial to measuring ranking performance against SEO efforts.

UTI's known first page rankings " Incr_or_Decr "d " Plus_or_MinusPage1_MOM . Page1_PercentIncr_MOM " MoM, majority of the keywords occurring for Automotive keywords moving from page 2 to page 1, with the average rank for this category improving by 6 positions."
)

CURRENT_Page1_UTI_Keywords	:= Round( tab.Range( "C37" ).Value )
CURRENT_Page1to3_UTI_Keywords	:= Round( tab.Range( "F37" ).Value )

DIFFERENCE_Page1_UTI_Keywords		:= ( CURRENT_Page1_UTI_Keywords	- PREVIOUS_Page1_UTI_Keywords		)
DIFFERENCE_Page1to3_UTI_Keywords	:= ( CURRENT_Page1to3_UTI_Keywords - CURRENT_Page1to3_UTI_Keywords	)


if ( DIFFERENCE_Page1_UTI_Keywords > 0 )
	{
		Up_or_DownP1		:= "up"
		Incr_or_DecrP1		:= "increase"
		Plus_or_MinusP1	:= "+"
	}
	else if ( DIFFERENCE_Page1_UTI_Keywords < 0 )
	{
		Up_or_DownP1		:= "down"
		Incr_or_DecrP1		:= "decrease"
		Plus_or_MinusP1	:= "-"
	}
	else if ( DIFFERENCE_Page1_UTI_Keywords = 0 )
	{
		Up_or_DownP1		:= "maintaining the same amount of keywords"
		Incr_or_DecrP1		:= "stayed the same"
		Plus_or_MinusP1	:= ""
	}
	
if ( DIFFERENCE_Page1to3_UTI_Keywords > 0 )
		Up_or_DownP1to3	:= "up"
	else if ( DIFFERENCE_Page1to3_UTI_Keywords < 0 )
		Up_or_DownP1to3	:= "down"
	else if ( DIFFERENCE_Page1to3_UTI_Keywords = 0 )
		Up_or_DownP1to3	:= "maintaining the same amount"


wsPOS.Range( "B128" ).Value	:=
(
"The number of keywords ranking on the first page for UTI was " CURRENT_Page1_UTI_Keywords " in " ReportingMonth ", " Up_or_DownP1 " from " PREVIOUS_Page1_UTI_Keywords " in " PREV_ReportingMonth ". 
The total number of keywords found ranking on the first three pages was " CURRENT_Page1to3_UTI_Keywords " in " ReportingMonth ", " Up_or_DownP1to3 " from " PREVIOUS_Page1to3_UTI_Keywords " in " PREV_ReportingMonth ".

Majority of these fluctuations were based around Location and Automotive keywords moving from page 2 to page 1. 
No action needed at this time."
)




;		EXECUTIVE SUMMARY	;
wsEXEC.Shapes( "TextBox 2" ).TextFrame.Characters.Text := 
(
"SUMMARY:
- Overall Organic Entrances " Incr_or_Decr_TRAFFIC1_EntrYOY "d YoY (" Plus_or_Minus_TRAFFIC1_EntrYOY . ChangePercentage_TRAFFIC1_EntrYOY ") and Inquiries " Incr_or_Decr_TRAFFIC1_InqrYOY "d YoY (" Plus_or_Minus_TRAFFIC1_InqrYOY . ChangePercentage_TRAFFIC1_InqrYOY "), which was due to an overall lift in:
- organic visibility for non-branded keywords (technical school near me, mechanic training, diesel mechanic schools) 
- traffic to Tuition, Location, Programs, and Scholarships and Grants pages
Location pages helped improve YoY traffic due to an increase in the total amount of ranking keywords for location related phrases (ex: trade school in long beach, technical school Houston, mechanic schools). 
When looking at the average rank for the Location pages YoY, average page rank decreased, but the total number of ranking keywords for those pages grew because the keyword profile diversified to include more nonbranded terms, which led to more organic clicks over time.
This tells us creating additional, informational Location and Campus based pages, like what is planned for Q3, will help increase incremental traffic by helping widen the range of keywords ranking for those pages.

To further improve Location Page performance, iProspect recommends doing an audit of Location Page meta data to help improve CTR and position, as well as identify any Google My Business Location pages for areas of opportunity towards the end of March.

- Overall Organic Mobile Entrances " Impr_or_Decl_TRAFFICm_InqrYOY "d " Plus_or_Minus_TRAFFICm_InqrYOY . ChangePercentage_TRAFFICm_InqrYOY " YoY, with " CURRENTPercent_AllOrganicEntr " of organic traffic coming from mobile devices.
iProspect to follow up with technical team to see when mobile speed updates will be made, considering Mom Organic Mobile Entrances declined by 5%.

- Looking at keyword performance, UTI's tracked keywords saw a " Plus_or_MinusP1 . DIFFERENCE_Page1_UTI_Keywords " MoM " Incr_or_DecrP1 " in page 1 keywords, which seemed to help increase organic entrances YoY and MoM. When looking at the top traffic driving and performing  keywords, Automotive continues to drive the most amount of traffic due to a combination of brand popularity and search interest.
To help grow nonbranded query traffic for all Programs, iProspect plans to conduct a content gap analysis in May.

CURRENT ACTION / IN PROGRESS ITEMS:
- Monitor performance of 4 Financial Aid Pages that were set live on 3/21
- Provide best practice content `"mockups`" for various pages to better adhere to SEO best practices and help improve organic visibility 
- Provide a content an content audit and gap analysis on live and Q3 pipeline content to improve organic visibility 
- Monitor ranking cannibalization to ensure the proper page is ranking and receiving traffic for the most relevant keyword 
- Ensure all Location pages are optimized to ensure the proper geographic page is appearing for localized searches and Google 3 Pack 
- Ensure all Google My Business Locations are claimed, up to date and optimized to appear in localized search and Google 3 Pack 

CURRENT ITEMS COMPLETED:
- UTI keyword book for internal linking
- Q3 content keyword research and mapping
- Q3 Top 20 Content and Meta Optimization
- Homepage and LocalBusiness schema markup
- Redirect chains
- Disavow spammy backlinks"
)




;		AUTOMOTIVE		;
tab 		:= wsAUTO

;		ALL TABLE		;
FIRSTrow_12month	:= 68
LASTrow_12month	:= 79

PREVrow_12month	:= ( FIRSTrow_12month + ReportingMonthDate - 2 )
CURRENTrow_12month	:= ( FIRSTrow_12month + ReportingMonthDate - 1 )
NEXTrow_12month	:= ( FIRSTrow_12month + ReportingMonthDate     )

tab.Activate

CurrentValue	:= tab.Range( "E" CURRENTrow_12month ).Value
PreviousValue	:= tab.Range( "D" CURRENTrow_12month ).Value

PercentChange( CurrentValue, PreviousValue )

ChangePercentage_12month_EntrYOY	:= ChangePercentage
Incr_or_Decr_12month_EntrYOY		:= Incr_or_Decr
Impr_or_Decl_12month_EntrYOY		:= Impr_or_Decl
Up_or_Down_12month_EntrYOY		:= Up_or_Down
Plus_or_Minus_12month_EntrYOY		:= Plus_or_Minus


CurrentValue	:= tab.Range( "J" CURRENTrow_12month ).Value
PreviousValue	:= tab.Range( "I" CURRENTrow_12month ).Value

PercentChange( CurrentValue, PreviousValue )

ChangePercentage_12month_InqrYOY	:= ChangePercentage
Incr_or_Decr_12month_InqrYOY		:= Incr_or_Decr
Impr_or_Decl_12month_InqrYOY		:= Impr_or_Decl
Up_or_Down_12month_InqrYOY		:= Up_or_Down
Plus_or_Minus_12month_InqrYOY		:= Plus_or_Minus

;		MOBILE TABLE		;
FIRSTrow	:= 119
JANrow	:= 143
LASTrow	:= 176

PREVrow		:= ( JANrow + ReportingMonthDate - 2 )
CURRENTrow	:= ( JANrow + ReportingMonthDate - 1 )
NEXTrow		:= ( JANrow + ReportingMonthDate     )

PREVYEARrow	:= ( CURRENTrow - 12 )

tab.Activate

CurrentValue	:= tab.Range( "C" CURRENTrow ).Value
PreviousValue	:= tab.Range( "C" PREVYEARrow ).Value

PercentChange( CurrentValue, PreviousValue )

ChangePercentage_EntrYOY	:= ChangePercentage
Incr_or_Decr_EntrYOY	:= Incr_or_Decr
Impr_or_Decl_EntrYOY	:= Impr_or_Decl
Up_or_Down_EntrYOY		:= Up_or_Down
Plus_or_Minus_EntrYOY	:= Plus_or_Minus

CURRENTPercent_AllOrganicEntr := ( tab.Range( "F" CURRENTrow	).Value * 100 )
PREVPercent_AllOrganicEntr	:= ( tab.Range( "F" PREVrow		).Value * 100 )

ChangePercentage := CURRENTPercent_AllOrganicEntr - PREVPercent_AllOrganicEntr

if ( ChangePercentage > 0 )
{	
	Up_or_Down	:= "up"
	Plus_or_Minus	:= "+"
}	
else if ( ChangePercentage < 0 )
{	
	Up_or_Down	:= "down"
	Plus_or_Minus	:= "-"
}	
else if ( ChangePercentage = 0 )
{	
	Up_or_Down	:= ""
	Plus_or_Minus	:= ""
}	

CURRENTPercent_AllOrganicEntr	:= Round( CURRENTPercent_AllOrganicEntr, 1 ) "%"

ChangePercentage			:= Round( Abs( Float( ChangePercentage ) ), 1 ) "%"


wsAUTO.Shapes( "TextBox 3" ).TextFrame.Characters.Text := 
(
"Summary:

Automotive Organic Entrances " Incr_or_Decr_12month_EntrYOY "d " Plus_or_Minus_12month_EntrYOY . ChangePercentage_12month_EntrYOY " YoY.

Automotive Organic Inquiries " Incr_or_Decr_12month_InqrYOY "d " Plus_or_Minus_12month_InqrYOY . ChangePercentage_12month_InqrYOY " YoY.

Automotive Organic Mobile Entrances " Incr_or_Decr_EntrYOY "d " Plus_or_Minus_EntrYOY . ChangePercentage_EntrYOY " YoY.

Automotive Organic Mobile Entrances is " CURRENTPercent_AllOrganicEntr " of all Organic Entrances in " ReportingMonth ", " Up_or_Down " " Plus_or_Minus . ChangePercentage " from " PREV_ReportingMonth ". 

Automotive is the top category with both the highest monthly Entrances and Inquiries.

UTI is second to ASE.com (auto testing and certification site) in terms of Automotive SOV. iProspect to track if this trend persists, and will conduct a competitive analysis if needed.

Non branded Auto keywords saw an increase in page 1 keywords and average rank for the category overall. This appears to be due to an lift in `"training,`" `"class`" and `"mechanic`" related keywords."
)


;		DIESEL		;
tab 		:= wsDIESEL

;		ALL TABLE		;
FIRSTrow_12month	:= 68
LASTrow_12month	:= 79

PREVrow_12month	:= ( FIRSTrow_12month + ReportingMonthDate - 2 )
CURRENTrow_12month	:= ( FIRSTrow_12month + ReportingMonthDate - 1 )
NEXTrow_12month	:= ( FIRSTrow_12month + ReportingMonthDate     )

tab.Activate

CurrentValue	:= tab.Range( "E" CURRENTrow_12month ).Value
PreviousValue	:= tab.Range( "D" CURRENTrow_12month ).Value

PercentChange( CurrentValue, PreviousValue )

ChangePercentage_12month_EntrYOY	:= ChangePercentage
Incr_or_Decr_12month_EntrYOY		:= Incr_or_Decr
Impr_or_Decl_12month_EntrYOY		:= Impr_or_Decl
Up_or_Down_12month_EntrYOY		:= Up_or_Down
Plus_or_Minus_12month_EntrYOY		:= Plus_or_Minus


CurrentValue	:= tab.Range( "J" CURRENTrow_12month ).Value
PreviousValue	:= tab.Range( "I" CURRENTrow_12month ).Value

PercentChange( CurrentValue, PreviousValue )

ChangePercentage_12month_InqrYOY	:= ChangePercentage
Incr_or_Decr_12month_InqrYOY		:= Incr_or_Decr
Impr_or_Decl_12month_InqrYOY		:= Impr_or_Decl
Up_or_Down_12month_InqrYOY		:= Up_or_Down
Plus_or_Minus_12month_InqrYOY		:= Plus_or_Minus

;		MOBILE TABLE		;
FIRSTrow	:= 119
JANrow	:= 143
LASTrow	:= 176

PREVrow		:= ( JANrow + ReportingMonthDate - 2 )
CURRENTrow	:= ( JANrow + ReportingMonthDate - 1 )
NEXTrow		:= ( JANrow + ReportingMonthDate     )

PREVYEARrow	:= ( CURRENTrow - 12 )

tab.Activate

CurrentValue	:= tab.Range( "C" CURRENTrow ).Value
PreviousValue	:= tab.Range( "C" PREVYEARrow ).Value

PercentChange( CurrentValue, PreviousValue )

ChangePercentage_EntrYOY	:= ChangePercentage
Incr_or_Decr_EntrYOY	:= Incr_or_Decr
Impr_or_Decl_EntrYOY	:= Impr_or_Decl
Up_or_Down_EntrYOY		:= Up_or_Down
Plus_or_Minus_EntrYOY	:= Plus_or_Minus

CURRENTPercent_AllOrganicEntr := ( tab.Range( "F" CURRENTrow	).Value * 100 )
PREVPercent_AllOrganicEntr	:= ( tab.Range( "F" PREVrow		).Value * 100 )

ChangePercentage := CURRENTPercent_AllOrganicEntr - PREVPercent_AllOrganicEntr

if ( ChangePercentage > 0 )
	{
		Up_or_Down	:= "up"
		Plus_or_Minus	:= "+"
	}
	else if ( ChangePercentage < 0 )
	{
		Up_or_Down	:= "down"
		Plus_or_Minus	:= "-"
	}
	else if ( ChangePercentage = 0 )
	{	
		Up_or_Down	:= ""
		Plus_or_Minus	:= ""
	}

CURRENTPercent_AllOrganicEntr	:= Round( CURRENTPercent_AllOrganicEntr, 1 ) "%"

ChangePercentage			:= Round( Abs( Float( ChangePercentage ) ), 1 ) "%"


wsDIESEL.Shapes( "TextBox 3" ).TextFrame.Characters.Text := 
(
"Summary:

Diesel Organic Entrances " Incr_or_Decr_12month_EntrYOY "d " Plus_or_Minus_12month_EntrYOY . ChangePercentage_12month_EntrYOY " YoY.

Diesel Organic Inquiries " Incr_or_Decr_12month_InqrYOY "d " Plus_or_Minus_12month_InqrYOY . ChangePercentage_12month_InqrYOY " YoY.

Diesel Organic Mobile Entrances " Incr_or_Decr_EntrYOY "d " Plus_or_Minus_EntrYOY . ChangePercentage_EntrYOY " YoY.

Diesel Organic Mobile Entrances is " CURRENTPercent_AllOrganicEntr " of all Organic Entrances in " ReportingMonth ", " Up_or_Down " " Plus_or_Minus . ChangePercentage " from " PREV_ReportingMonth ". 

Study.com (online school - similar to coursera) has the largest SOV for tracked Diesel keywords, with UTI and alltrucking.com trailing behind.

Growing trend in `"best of schools`" websites ranking for general terms such as `"diesel mechanic schools`" and `"diesel trade schools`" - iProspect suggests finding other areas of opportunity to grow traffic for Diesel."
)

;		MOTO		;
tab 		:= wsMOTO

;		ALL TABLE		;
FIRSTrow_12month	:= 68
LASTrow_12month	:= 79

PREVrow_12month	:= ( FIRSTrow_12month + ReportingMonthDate - 2 )
CURRENTrow_12month	:= ( FIRSTrow_12month + ReportingMonthDate - 1 )
NEXTrow_12month	:= ( FIRSTrow_12month + ReportingMonthDate     )

tab.Activate

CurrentValue	:= tab.Range( "E" CURRENTrow_12month ).Value
PreviousValue	:= tab.Range( "D" CURRENTrow_12month ).Value

PercentChange( CurrentValue, PreviousValue )

ChangePercentage_12month_EntrYOY	:= ChangePercentage
Incr_or_Decr_12month_EntrYOY		:= Incr_or_Decr
Impr_or_Decl_12month_EntrYOY		:= Impr_or_Decl
Up_or_Down_12month_EntrYOY		:= Up_or_Down
Plus_or_Minus_12month_EntrYOY		:= Plus_or_Minus


CurrentValue	:= tab.Range( "J" CURRENTrow_12month ).Value
PreviousValue	:= tab.Range( "I" CURRENTrow_12month ).Value

PercentChange( CurrentValue, PreviousValue )

ChangePercentage_12month_InqrYOY	:= ChangePercentage
Incr_or_Decr_12month_InqrYOY		:= Incr_or_Decr
Impr_or_Decl_12month_InqrYOY		:= Impr_or_Decl
Up_or_Down_12month_InqrYOY		:= Up_or_Down
Plus_or_Minus_12month_InqrYOY		:= Plus_or_Minus

;		MOBILE TABLE		;
FIRSTrow	:= 119
JANrow	:= 143
LASTrow	:= 176

PREVrow		:= ( JANrow + ReportingMonthDate - 2 )
CURRENTrow	:= ( JANrow + ReportingMonthDate - 1 )
NEXTrow		:= ( JANrow + ReportingMonthDate     )

PREVYEARrow	:= ( CURRENTrow - 12 )

tab.Activate

CurrentValue	:= tab.Range( "C" CURRENTrow ).Value
PreviousValue	:= tab.Range( "C" PREVYEARrow ).Value

PercentChange( CurrentValue, PreviousValue )

ChangePercentage_EntrYOY	:= ChangePercentage
Incr_or_Decr_EntrYOY	:= Incr_or_Decr
Impr_or_Decl_EntrYOY	:= Impr_or_Decl
Up_or_Down_EntrYOY		:= Up_or_Down
Plus_or_Minus_EntrYOY	:= Plus_or_Minus

CURRENTPercent_AllOrganicEntr := ( tab.Range( "F" CURRENTrow	).Value * 100 )
PREVPercent_AllOrganicEntr	:= ( tab.Range( "F" PREVrow		).Value * 100 )

ChangePercentage := CURRENTPercent_AllOrganicEntr - PREVPercent_AllOrganicEntr

if ( ChangePercentage > 0 )
	{
		Up_or_Down	:= "up"
		Plus_or_Minus	:= "+"
	}	
	else if ( ChangePercentage < 0 )
	{	
		Up_or_Down	:= "down"
		Plus_or_Minus	:= "-"
	}		
	else if ( ChangePercentage = 0 )
	{	
		Up_or_Down	:= ""
		Plus_or_Minus	:= ""
	}	
	
CURRENTPercent_AllOrganicEntr	:= Round( CURRENTPercent_AllOrganicEntr, 1 ) "%"

ChangePercentage			:= Round( Abs( Float( ChangePercentage ) ), 1 ) "%"


wsMOTO.Shapes( "TextBox 3" ).TextFrame.Characters.Text := 
(
"Summary:

Motorcycle Organic Entrances " Incr_or_Decr_12month_EntrYOY "d " Plus_or_Minus_12month_EntrYOY . ChangePercentage_12month_EntrYOY " YoY.

Motorcycle Organic Inquiries " Incr_or_Decr_12month_InqrYOY "d " Plus_or_Minus_12month_InqrYOY . ChangePercentage_12month_InqrYOY " YoY.

Motorcycle Organic Mobile Entrances " Incr_or_Decr_EntrYOY "d " Plus_or_Minus_EntrYOY . ChangePercentage_EntrYOY " YoY.

Motorcycle Organic Mobile Entrances is " CURRENTPercent_AllOrganicEntr " of all Organic Entrances in " ReportingMonth ", " Up_or_Down " " Plus_or_Minus . ChangePercentage " from " PREV_ReportingMonth ". 

UTI leads Moto SOV for the tracked keyword set, with btosports.com (moto gear site) and tradeschool.net behind UTI.

By a large margin, the Motorcycle category has the highest percentage of keywords ranking on page 1 and the most competitive average page rank. Page 1 rankings remained flat MoM."
)




;		MARINE		;
tab 		:= wsMARINE

;		ALL TABLE		;
FIRSTrow_12month	:= 68
LASTrow_12month	:= 79

PREVrow_12month	:= ( FIRSTrow_12month + ReportingMonthDate - 2 )
CURRENTrow_12month	:= ( FIRSTrow_12month + ReportingMonthDate - 1 )
NEXTrow_12month	:= ( FIRSTrow_12month + ReportingMonthDate     )

tab.Activate

CurrentValue	:= tab.Range( "E" CURRENTrow_12month ).Value
PreviousValue	:= tab.Range( "D" CURRENTrow_12month ).Value

PercentChange( CurrentValue, PreviousValue )

ChangePercentage_12month_EntrYOY	:= ChangePercentage
Incr_or_Decr_12month_EntrYOY		:= Incr_or_Decr
Impr_or_Decl_12month_EntrYOY		:= Impr_or_Decl
Up_or_Down_12month_EntrYOY		:= Up_or_Down
Plus_or_Minus_12month_EntrYOY		:= Plus_or_Minus


CurrentValue	:= tab.Range( "J" CURRENTrow_12month ).Value
PreviousValue	:= tab.Range( "I" CURRENTrow_12month ).Value

PercentChange( CurrentValue, PreviousValue )

ChangePercentage_12month_InqrYOY	:= ChangePercentage
Incr_or_Decr_12month_InqrYOY		:= Incr_or_Decr
Impr_or_Decl_12month_InqrYOY		:= Impr_or_Decl
Up_or_Down_12month_InqrYOY		:= Up_or_Down
Plus_or_Minus_12month_InqrYOY		:= Plus_or_Minus

;		MOBILE TABLE		;
FIRSTrow	:= 119
JANrow	:= 143
LASTrow	:= 176

PREVrow		:= ( JANrow + ReportingMonthDate - 2 )
CURRENTrow	:= ( JANrow + ReportingMonthDate - 1 )
NEXTrow		:= ( JANrow + ReportingMonthDate     )

PREVYEARrow	:= ( CURRENTrow - 12 )

tab.Activate

CurrentValue	:= tab.Range( "C" CURRENTrow ).Value
PreviousValue	:= tab.Range( "C" PREVYEARrow ).Value

PercentChange( CurrentValue, PreviousValue )

ChangePercentage_EntrYOY	:= ChangePercentage
Incr_or_Decr_EntrYOY	:= Incr_or_Decr
Impr_or_Decl_EntrYOY	:= Impr_or_Decl
Up_or_Down_EntrYOY		:= Up_or_Down
Plus_or_Minus_EntrYOY	:= Plus_or_Minus

CURRENTPercent_AllOrganicEntr := ( tab.Range( "F" CURRENTrow	).Value * 100 )
PREVPercent_AllOrganicEntr	:= ( tab.Range( "F" PREVrow		).Value * 100 )

ChangePercentage := CURRENTPercent_AllOrganicEntr - PREVPercent_AllOrganicEntr

if ( ChangePercentage > 0 )
	{
		Up_or_Down	:= "up"
		Plus_or_Minus	:= "+"
	}	
	else if ( ChangePercentage < 0 )
	{	
		Up_or_Down	:= "down"
		Plus_or_Minus	:= "-"
	}	
	else if ( ChangePercentage = 0 )
	{	
		Up_or_Down	:= ""
		Plus_or_Minus	:= ""
	}	

CURRENTPercent_AllOrganicEntr	:= Round( CURRENTPercent_AllOrganicEntr, 1 ) "%"

ChangePercentage			:= Round( Abs( Float( ChangePercentage ) ), 1 ) "%"


wsMARINE.Shapes( "TextBox 3" ).TextFrame.Characters.Text := 
(
"Summary:

Marine Organic Entrances " Incr_or_Decr_12month_EntrYOY "d " Plus_or_Minus_12month_EntrYOY . ChangePercentage_12month_EntrYOY " YoY.

Marine Organic Inquiries " Incr_or_Decr_12month_InqrYOY "d " Plus_or_Minus_12month_InqrYOY . ChangePercentage_12month_InqrYOY " YoY.

Marine Organic Mobile Entrances " Incr_or_Decr_EntrYOY "d " Plus_or_Minus_EntrYOY . ChangePercentage_EntrYOY " YoY.

Marine Organic Mobile Entrances is " CURRENTPercent_AllOrganicEntr " of all Organic Entrances in " ReportingMonth ", " Up_or_Down " " Plus_or_Minus . ChangePercentage " from " PREV_ReportingMonth ". 

UTI leads Marine SOV for the tracked keyword set, with collegegrad.com and study.com (both college search sites) trailing behind UTI.

Page 1 Marine keywords remained flat MoM, although average rank improved by 1 position which helped improve overall organic traffic."
)




;		BLOG		;
tab 		:= wsBLOG

;		ALL TABLE		;
FIRSTrow_12month	:= 68
LASTrow_12month	:= 79

PREVrow_12month	:= ( FIRSTrow_12month + ReportingMonthDate - 2 )
CURRENTrow_12month	:= ( FIRSTrow_12month + ReportingMonthDate - 1 )
NEXTrow_12month	:= ( FIRSTrow_12month + ReportingMonthDate     )

tab.Activate

CurrentValue	:= tab.Range( "E" CURRENTrow_12month ).Value
PreviousValue	:= tab.Range( "D" CURRENTrow_12month ).Value

PercentChange( CurrentValue, PreviousValue )

ChangePercentage_12month_EntrYOY	:= ChangePercentage
Incr_or_Decr_12month_EntrYOY		:= Incr_or_Decr
Impr_or_Decl_12month_EntrYOY		:= Impr_or_Decl
Up_or_Down_12month_EntrYOY		:= Up_or_Down
Plus_or_Minus_12month_EntrYOY		:= Plus_or_Minus


CurrentValue	:= tab.Range( "J" CURRENTrow_12month ).Value
PreviousValue	:= tab.Range( "I" CURRENTrow_12month ).Value

PercentChange( CurrentValue, PreviousValue )

ChangePercentage_12month_InqrYOY	:= ChangePercentage
Incr_or_Decr_12month_InqrYOY		:= Incr_or_Decr
Impr_or_Decl_12month_InqrYOY		:= Impr_or_Decl
Up_or_Down_12month_InqrYOY		:= Up_or_Down
Plus_or_Minus_12month_InqrYOY		:= Plus_or_Minus

;		MOBILE TABLE		;
FIRSTrow	:= 119
JANrow	:= 143
LASTrow	:= 176

PREVrow		:= ( JANrow + ReportingMonthDate - 2 )
CURRENTrow	:= ( JANrow + ReportingMonthDate - 1 )
NEXTrow		:= ( JANrow + ReportingMonthDate     )

PREVYEARrow	:= ( CURRENTrow - 12 )

tab.Activate

CurrentValue	:= tab.Range( "C" CURRENTrow ).Value
PreviousValue	:= tab.Range( "C" PREVYEARrow ).Value

PercentChange( CurrentValue, PreviousValue )

ChangePercentage_EntrYOY	:= ChangePercentage
Incr_or_Decr_EntrYOY	:= Incr_or_Decr
Impr_or_Decl_EntrYOY	:= Impr_or_Decl
Up_or_Down_EntrYOY		:= Up_or_Down
Plus_or_Minus_EntrYOY	:= Plus_or_Minus

CURRENTPercent_AllOrganicEntr := ( tab.Range( "F" CURRENTrow	).Value * 100 )
PREVPercent_AllOrganicEntr	:= ( tab.Range( "F" PREVrow		).Value * 100 )

ChangePercentage := CURRENTPercent_AllOrganicEntr - PREVPercent_AllOrganicEntr

if ( ChangePercentage > 0 )
	{
		Up_or_Down	:= "up"
		Plus_or_Minus	:= "+"
	}
	else if ( ChangePercentage < 0 )
	{	
		Up_or_Down	:= "down"
		Plus_or_Minus	:= "-"
	}	
	else if ( ChangePercentage = 0 )
	{	
		Up_or_Down	:= ""
		Plus_or_Minus	:= ""
	}	

CURRENTPercent_AllOrganicEntr	:= Round( CURRENTPercent_AllOrganicEntr, 1 ) "%"

ChangePercentage			:= Round( ChangePercentage, 1 ) "%"


wsBLOG.Shapes( "TextBox 3" ).TextFrame.Characters.Text := 
(
"Summary:

Blog Organic Entrances " Incr_or_Decr_12month_EntrYOY "d " Plus_or_Minus_12month_EntrYOY . ChangePercentage_12month_EntrYOY " YoY.

Blog Organic Inquiries " Incr_or_Decr_12month_InqrYOY "d " Plus_or_Minus_12month_InqrYOY . ChangePercentage_12month_InqrYOY " YoY.

Blog Organic Mobile Entrances " Incr_or_Decr_EntrYOY "d " Plus_or_Minus . ChangePercentage_EntrYOY " YoY.

Blog Organic Mobile Entrances is " CURRENTPercent_AllOrganicEntr " of all Organic Entrances in " ReportingMonth ", " Up_or_Down " " Plus_or_Minus . ChangePercentage " from " PREV_ReportingMonth ". 

Plannedparenthood.com has the largest SOV for tracked Blog keywords, with UTI and webmd.com trailing behind in organic visibility.

Despite the slight decrease in average rank for Blog keywords, YoY traffic maintained because majority of those keywords were related to Diesel Program keywords, which were not hyper relevant to the blog pages."
)

;		BUTTON-UP WORKSHEETS		;

wsMILE.Activate
wsMILE.Range( "A1" ).Select
wsTRAFFIC.Activate
wsTRAFFIC.Range( "A1" ).Select
wsAUTO.Activate
wsAUTO.Range( "A1" ).Select
wsDIESEL.Activate
wsDIESEL.Range( "A1" ).Select
wsMOTO.Activate
wsMOTO.Range( "A1" ).Select
wsMARINE.Activate
wsMARINE.Range( "A1" ).Select
wsBLOG.Activate
wsBLOG.Range( "A1" ).Select
wsPOS.Activate
wsPOS.Range( "A1" ).Select
wsGA.Activate
wsGA.Range( "A1" ).Select
wsGAm.Activate
wsGAm.Range( "A1" ).Select
wsSTAT.Activate
wsSTAT.Range( "A1" ).Select
wsEXEC.Activate
wsEXEC.Range( "A1" ).Select


;		SAVE UPDATED REPORT WITH NEW FILENAME		;
WorkbookFile := WBdir "\" AccountName " SEO Report " ReportingMonth " " CurrentYear "." WBext

try
{
	Xl.ActiveWorkbook.SaveCopyAs(WorkbookFile)
	Xl.ActiveWorkbook.Close
	Xl.Workbooks.Open(WorkbookFile)
	Xl.Visible := true

	MsgBox(	Text		:= "Report is complete and saved with updated filename!`n`n"
					.  "Ensure all summaries, graphs and tables have been updated correctly."
		,	Title	:= "REPORT COMPLETE (Saved)"
		,	Options	:=  0x40040) ; System Modal (always on top) with INFO icon
} 
catch 
{
	MsgBox(	Text		:= "Report is complete, but could not be saved automatically.`n`n"
					.  "Ensure all summaries, graphs and tables have been updated correctly. Then Save!"
		,	Title	:= "REPORT COMPLETE"
		,	Options	:=  0x40010) ; System Modal (always on top) with STOP hand icon
}

Xl.DisplayAlerts := True
Xl := ""
RETURN



;		HOTKEYS		;
!Esc::
try Xl.Quit
Xl := ""
ExitApp


;		DATA FUNCTIONS		;
Top10_URL( ContainsString, fromWB, toWB, toWS, toRange_URL, toRange_Entrances ) {
	
Global
	
	Xl.Windows( fromWB ).Activate
	Xl.Workbooks( fromWB ).Activate
	
	Xl.ActiveSheet.Range( "A1:D" TableRows )
	.AutoFilter( Field := 1, Criteria1 := "*/" ContainsString "*" )
	
	Xl.ActiveSheet.Range( "A1" ).Select
	Xl.ActiveCell.CurrentRegion.Select 
	Xl.Selection.SpecialCells( 12 ).Copy( Xl.ActiveSheet.Range("F1") )
	
	Top10_URL			:= Xl.ActiveSheet.Range( "F2:F11" ).Value
	Top10_Entrances	:= Xl.ActiveSheet.Range( "H2:H11" ).Value
	
	Xl.Windows( toWB ).Activate
	Xl.Workbooks( toWB ).Activate	
	Xl.Worksheets( toWS ).Activate	
	
	Xl.Worksheets( toWS ).Range( toRange_URL ).Value		:= Top10_URL
	Xl.Worksheets( toWS ).Range( toRange_Entrances ).Value	:= Top10_Entrances
}


PercentChange( CurrentValue, PreviousValue ) {
	
Global	
	
	try 
	{	
		ChangePercentage :=  Float( ((( CurrentValue - PreviousValue ) / PreviousValue ) * 100 ) )

	
		if ( ChangePercentage > 0 )
		{
			Incr_or_Decr	:= "increase"
			Impr_or_Decl	:= "improve"
			Up_or_Down	:= "up"
			Plus_or_Minus	:= "+"
		}
		else if ( ChangePercentage < 0 )
		{
			Incr_or_Decr	:= "decrease"
			Impr_or_Decl	:= "decline"
			Up_or_Down	:= "down"
			Plus_or_Minus	:= "-"
		}
		else if ( ChangePercentage = 0 )
		{
			Incr_or_Decr	:= "stayed the same"
			Impr_or_Decl	:= ""
			Up_or_Down	:= ""
			Plus_or_Minus	:= ""
		}

	}
	catch e
	{
		if ( e.message = "Divide by zero." )
		{			
			ChangePercentage	:= "***COULDN'T DIVIDE BY ZERO***"
			Incr_or_Decr		:= ""
			Impr_or_Decl		:= ""
			Up_or_Down		:= ""
			Plus_or_Minus		:= ""		
		}
	}
	
	if ( !IsObject( e ) )
	{	
		ChangePercentage	:= Round( Abs( ChangePercentage ), 1 ) "%"
	}	
	else if ( IsObject( e ) )
	{
		e				:= ""
	}
	
}
;		GUI FUNCTIONS		;

OnSubmit( Month ) {
	
Global
	

	
	ReportingMonth		:= ( Month.Text  )
	ReportingMonthDate	:= ( Month.Value )

	PREV_NEXT_Months := [ "January", "February", "March", "April", "May", "June", "July"
					, "August", "September", "October", "November", "December" ]

	PREV_ReportingMonth	:= PREV_NEXT_Months[ ReportingMonthDate - 1 ]
	NEXT_ReportingMonth	:= PREV_NEXT_Months[ ReportingMonthDate + 1 ]	
	
	ReportingMonth_CAPS	:= StrUpper( ReportingMonth )
	
	;		Determine End Date of Month		;
	if		( ReportingMonthDate =  1 )
	||		( ReportingMonthDate =  3 )
	||		( ReportingMonthDate =  5 )
	||		( ReportingMonthDate =  7 )
	||		( ReportingMonthDate =  8 )
	||		( ReportingMonthDate =  10 )
	||		( ReportingMonthDate =  12 )
		
					LastDateOfMonth := 31
	
	else if	( ReportingMonthDate =  4 )
	||		( ReportingMonthDate =  6 )
	||		( ReportingMonthDate =  9 )
	||		( ReportingMonthDate =  11 )
		
					LastDateOfMonth := 30
	
	else if	( ReportingMonthDate =  2 )
		
					LastDateOfMonth := 28	
	
					; account for leap years
					if ( CurrentYear = 2020 ) 
					|| ( CurrentYear = 2024 )
					|| ( CurrentYear = 2028 )
					|| ( CurrentYear = 2032 )
						
						LastDateOfMonth := 29				
	
	Gui.Destroy()	
	
}

