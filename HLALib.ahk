/*
Div's HLA Library
AHK_L v1.1
Last updated 07-2017

Description: Used for custom utilities in HT
For utility to function, this file must be in the same directory as the exe utility
*/

;Directives and Startup
#NoEnv ;Avoids checking empty variables to see if they are environment variables
#NoTrayIcon ;Disables the showing of a tray icon.
#SingleInstance Force ;Replaces the old instance automatically
SetWorkingDir %A_ScriptDir% ;Working directory changed to script residing directory
#include adosql.ahk ;Library to connect to MSSQL database and execute SQL queries. ADOSQL() function is in this library

; Connection Strings to connect to HistoTrac and Fusion databases
global ConnectString_HT ;Need to declare this as a global variable to use it inside functions
ConnectString_HT := "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=HistoTracPROD;Data Source=vm-histotrac;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=2UA2490GTM;Use Encryption for Data=False;Tag with column collation when possible=False;"
global ConnectString_FUSION
ConnectString_FUSION := "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=FUSION;Data Source=vm-hlasql-prod;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=2UA2490GTM;Use Encryption for Data=False;Tag with column collation when possible=False"
;result := ADOSQL(ConnectString_HT, Query_HT) ;This line carries out the the SQL query and outputs it into "result" as a 2d array.
;result[2][1] is the value stored in row 2, column 1. Row 1 is always going to be the column headers.

;Removes duplicate strings (`n delimiter) without affecting list order
remove_dupes(string) 
{
	oList := {}
	Loop, parse, string, `n, `r
		if	!oList[A_LoopField]
			oList[A_LoopField] :=	1, str .=	(!str ? "" : "`n") A_LoopField
	return str
}

;Get sample ID and dates for SAB tests using MRN as a parameter; get_sab_sid(1234567); <sampleID>,<sampleDate>`n 
get_sab_sid(mrnum)
{
	;SQL query column headers -> [MRN] [SSN] [Sample Date] [Sample Number] [Test Type]
	Query_HT := "Select  Patient.HospitalID AS [MR Number], Patient.SSNbr AS [SSN], Sample.SampleDt AS [Sample Date], Sample.SampleNbr AS [Sample Number], Test.TestTypeCd AS [Test Type] FROM  Patient Patient INNER JOIN Sample Sample ON  Sample.PatientId = Patient.PatientId INNER JOIN Test Test ON  Test.SampleId = Sample.SampleId WHERE  Patient.HospitalID = '" . mrnum . "' AND Sample.SampleDt > 0 AND Test.TestTypeCd IN ('A-L-SAI TRTD','A-L-SAII TRTD','A-L-SAI MTH','A-L-SAII MTH') ORDER BY  Sample.SampleDt Asc"
	result := ADOSQL(ConnectString_HT, Query_HT)
	rmi := result.MaxIndex()-1 ; the -1 excludes the the column headers from the output
	if !result[2][1] ; If query returned no data
	{
		error_1 = "get_sab_sid QUERY UNSUCCESSFUL!"
		return error_1
	}
	Loop, %rmi%
	{
		n := A_Index+1
		snum .= result[n][4] "`," result [n][3] "`n"		
	}
	snum := remove_dupes(snum)
	return snum
}

;Get HISTOTRAC SAB results for sample number provided; get_sab_result("17-12345"); This function is mostly useless because there is no standardized format in HT
get_sab_HT(sid) ;sid must be in double quotes, see example above.
{
	;SQL query column headers ->[MRN] [SSN] [Sample Date] [Sample Number] [Test Type] [Specificity] [Mod Risk Spec] [Low Risk Spec] [Reportable Comments]
	Query_HT := "Select  Patient.HospitalID AS [MR Number], Patient.SSNbr AS [SSN], Sample.SampleDt AS [Sample Date], Sample.SampleNbr AS [Sample Number], Test.TestTypeCd AS [Test Type], Test.SpecificityTxt AS [Specificity], Test.ModerateRiskAntibodyTxt AS [Moderate Risk Antibody], Test.LowRiskAntibodyTxt AS [Low Risk Antibody], Test.ReportableCommentsTxt AS [Reportable Comments] FROM  Patient Patient INNER JOIN Sample Sample ON  Sample.PatientId = Patient.PatientId INNER JOIN Test Test ON  Test.SampleId = Sample.SampleId WHERE Sample.SampleNbr = '" . sid . "' AND Test.TestTypeCd IN ('A-L-SAI TRTD','A-L-SAII TRTD','A-L-SAI MTH','A-L-SAII MTH') ORDER BY  Sample.SampleDt Asc"
	result := ADOSQL(ConnectString_HT, Query_HT)
	if !result[2][1] ; If query returned no data
	{
		error_1 = "get_sab_HT QUERY UNSUCCESSFUL!"
		return error_1
	}
	rmi := result.MaxIndex()-1 ; the -1 excludes the the column headers from the output
	Loop, %rmi%
	{
		n := A_Index+1
		sabres .= result[n][6] "|" result [n][7] "|" result [n][8] "|" result [n][9] "`n"
	}
	sabres := remove_dupes(sabres)
	return sabres ;
}

;Get FUSION SAB results for sample number provided; get_sab_result("17-12345");  returns an array [patient name, final assignment, other assignment]; arr[1] = name
get_sab_fusion(sid)
{
	Query_Fusion := "SELECT SAMPLE.SampleIDName, WELL.NC1, WELL.PC1, WELL_RESULT.ResultType, WELL.NC2, WELL.PC2, WELL_RESULT.Value01, WELL_RESULT.Value02 FROM WELL LEFT OUTER JOIN SAMPLE ON WELL.SampleID = SAMPLE.SampleID AND WELL.IsActive = 1 AND SAMPLE.IsActive = 1 LEFT OUTER JOIN WELL_RESULT ON WELL.WellID = WELL_RESULT.WellID WHERE (1 = 1) AND (WELL_RESULT.ResultType IN ('54', '59' , '58')) AND (SAMPLE.SampleIDName LIKE '" . sid . "%') ORDER BY WELL.AnalysisDT DESC, SAMPLE.SampleIDName, WELL.WellPosition, WELL_RESULT.ResultType, WELL_RESULT.Value01"
	result := ADOSQL(ConnectString_FUSION, Query_Fusion)
	if ADOSQL_LastError
		return ADOSQL_LastError
	if !result[2][1] ; If query returned no data
	{
		error_1 = "get_sab_fusion QUERY UNSUCCESSFUL!"
		return error_1
	}
	name := result[2][1]
	StringSplit, lnfn, name, %A_Space%
	rmi := result.MaxIndex()-1 ; the -1 excludes the the column headers from the output
	Loop, %rmi%
	{
		n := A_Index+1
		restype := result[n][4]
		if restype = 54
			sabres .= result[n][7] "`n"
		if restype = 58
			other .= result[n][7] "`n"
	}		
	sabres := remove_dupes(sabres)
	final := [lnfn2,sabres,other]
	return final ;
}


;Get FUSION typing for sample number and locus provided; get_typing_fusion("17-15237","DRB1"); returns an array in this format [sero1,sero2,allele1,allele2]; arr[1] = sero1
get_typing_fusion(sid,loc)
{
	Query_Fusion := "SELECT SAMPLE.SampleIDName, WELL_RESULT.ResultType, WELL_RESULT.Value01, WELL_RESULT.Value02, WELL_RESULT.Value03 FROM WELL LEFT OUTER JOIN SAMPLE ON WELL.SampleID = SAMPLE.SampleID AND WELL.IsActive = 1 AND SAMPLE.IsActive = 1 LEFT OUTER JOIN WELL_RESULT ON WELL.WellID = WELL_RESULT.WellID WHERE (1 = 1) AND (SAMPLE.SampleIDName LIKE '" . sid . "%') AND (WELL_RESULT.ResultType IN ('08','06')) AND (WELL_RESULT.Value03 = '0001') AND (WELL_RESULT.Value02 = '" . loc . "') ORDER BY WELL.AnalysisDT DESC, SAMPLE.SampleIDName, WELL.WellPosition, WELL_RESULT.ResultType, WELL_RESULT.Value01"
	result := ADOSQL(ConnectString_FUSION, Query_Fusion)
	if ADOSQL_LastError
		return ADOSQL_LastError
	if !result[2][1] ; If query returned no data
	{
		error_1 = "get_typing_fusion QUERY UNSUCCESSFUL!"
		return error_1
	}
	rmi := result.MaxIndex()-1 ; the -1 excludes the the column headers from the output
	Loop, %rmi%
	{
		n := A_Index+1
		restype := result[n][2]
		if restype = 06
			allele .= result[n][3] "`n"
		if restype = 08
			sero .= result[n][3] "`n"
	}
	allele := remove_dupes(allele)
	sero := remove_dupes(sero)
	StringSplit, sero, sero, %A_Space%
	StringSplit, allele, allele, %A_Space%
	final := [sero1,sero2,allele1,allele2]
	return final
}

;Get HLA typing from HT provided UNOS# and Locus; get_typing_ht("UNOS123","A1"); <typing_result>
get_typing_ht(unos,loc)
{
	Query_HT := "Select  Patient.HospitalID AS [MR Number], Patient.UNOSId AS [UNOS Id], Patient.mA1EqCd AS [A1 Equivalent], Patient.mA2EqCd AS [A2 Equivalent], Patient.mB1EqCd AS [B1 Equivalent], Patient.mB2EqCd AS [B2 Equivalent], Patient.mBw1EqCd AS [Bw1 Equivalent], Patient.mBw2EqCd AS [Bw2 Equivalent], Patient.mC1EqCd AS [C1 Equivalent], Patient.mC2EqCd AS [C2 Equivalent], Patient.mDRB11EqCd AS [DRB11 Equivalent], Patient.mDRB12EqCd AS [DRB12 Equivalent], Patient.mDRB32EqCd AS [DRB32 Equivalent], Patient.mDRB31EqCd AS [DRB31 Equivalent], Patient.mDRB41EqCd AS [DRB41 Equivalent], Patient.mDRB42EqCd AS [DRB42 Equivalent], Patient.mDRB51EqCd AS [DRB51 Equivalent], Patient.mDRB52EqCd AS [DRB52 Equivalent], Patient.mDQB11EqCd AS [DQB11 Equivalent], Patient.mDQB12EqCd AS [DQB12 Equivalent], Patient.mDPB11cd AS [DPB11], Patient.mDPB12cd AS [DPB12], Patient.mDPA11Cd AS [DPA11], Patient.mDPA12Cd AS [DPA12], Patient.mDQA11Cd AS [DQA11], Patient.mDQA12Cd AS [DQA12] FROM  Patient Patient WHERE  Patient.UNOSId = '" . unos . "'"
	result := ADOSQL(ConnectString_HT, Query_HT)
	if !result[2][1] ; If query returned no data
	{
		error_1 = "get_typing_ht QUERY UNSUCCESSFUL!"
		return error_1
	}
	n := 13
	Loop, 6
	{
		drw .= result[2][n] "`n"
		n := n+1
	}
	dw := remove_dupes(drw)
	Stringsplit, dw, dw, `n ;Codeblock lines 144-150 removes empty DRw fields.
	typing := { "A1":result[2][3], "A2":result[2][4], "B1":result[2][5], "B2":result[2][6], "Bw1":result[2][7], "Bw2":result[2][8], "C1":result[2][9], "C2":result[2][10], "DR1":result[2][11], "DR2":result[2][12], "DQB1":result[2][19], "DQB2":result[2][20], "DQA1":result[2][25], "DQA2":result[2][26], "DPB1":result[2][21], "DPB2":result[2][22], "DPA1":result[2][23], "DPA2":result[2][24], "DW1":dw1, "DW2":dw2 } ;Creates an associative array of typing results in the form of {"LocusType":Value}
	final := typing[loc]
	return final
}

;Get HT previous transplant info for MRN provided; get_prevtx(1234567); <tx_date>|<donor_mrn>|<organ>|<donor_hlatyping>`n
get_prevtx(mrnum)
{
	Query_HT := "Select  Patient.HospitalID AS [MR Number], Patient.lastnm AS [Last Name], Patient.firstnm AS [First Name], Patient.categoryCd AS [Category], TransplantHistory.TransplantDt AS [Event Date], TransplantHistory.EventCd AS [Event], TransplantHistory.DonorNm AS [Donor Name], TransplantHistory.DonorId AS [Donor ID], TransplantHistory.OrganCd AS [Organ], TransplantHistory.DonorMolecularTxt AS [Donor Molecular Typing] FROM  Patient Patient INNER JOIN TransplantHistory TransplantHistory ON  TransplantHistory.PatientId = Patient.PatientId WHERE  Patient.HospitalID = '" . mrnum . "' AND TransplantHistory.EventCd = 'Transplant'"
	result := ADOSQL(ConnectString_HT, Query_HT)
	if !result[2][1] ; If query returned no data
	{
		error_1 = "get_prevtx QUERY UNSUCCESSFUL!"
		return error_1
	}
	rmi := result.MaxIndex()-1 ; the -1 excludes the the column headers from the output
	Loop, %rmi%
	{
		n := A_Index+1
		txdate := result[n][5]
		txmrn := result[n][8]
		txorgan := result[n][9]
		txtyping := result[n][10]
		txstring .= txdate "|" txmrn "|" txorgan "|" txtyping "`n"
	}
	return txstring
}

/*
bla := get_typing_fusion("17-15237","DRB1")
MsgBox % bla
/*
TEST CODE and STRINGS
------------
[MRN][SSN][Sample Date][Sample Number][Test Type][Specificity][Mod Risk Ab][Low Risk Ab][Reportable Comments]
------------
mrnum = 631978
	;SQL query column headers -> [MRN] [SSN] [Sample Date] [Sample Number] [Test Type]
	Query_HT := "Select  Patient.HospitalID AS [MR Number], Patient.SSNbr AS [SSN], Sample.SampleDt AS [Sample Date], Sample.SampleNbr AS [Sample Number], Test.TestTypeCd AS [Test Type], Test.SpecificityTxt AS [Specificity], Test.ModerateRiskAntibodyTxt AS [Moderate Risk Antibody], Test.LowRiskAntibodyTxt AS [Low Risk Antibody], Test.ReportableCommentsTxt AS [Reportable Comments] FROM  Patient Patient INNER JOIN Sample Sample ON  Sample.PatientId = Patient.PatientId INNER JOIN Test Test ON  Test.SampleId = Sample.SampleId WHERE  Patient.HospitalID = '" . mrnum . "' AND Sample.SampleDt > 0 AND Test.TestTypeCd IN ('A-L-SAI TRTD','A-L-SAII TRTD','A-L-SAI MTH','A-L-SAII MTH') ORDER BY  Sample.SampleDt Asc"
	result := ADOSQL(ConnectString_HT, Query_HT)
	rmi := result.MaxIndex()
	Loop, %rmi%
	{
		n := A_Index+1
		snum .= result[n][4] "`," result [n][3] "`n" ;output will contain sample number and sample date separated by a comma
	}
	snum := remove_dupes(snum)
	
	Msgbox % snum
---------------------	
mrnum = 735692
	sid = 15-21863

	Query_HT := "Select  Patient.HospitalID AS [MR Number], Patient.SSNbr AS [SSN], Sample.SampleDt AS [Sample Date], Sample.SampleNbr AS [Sample Number], Test.TestTypeCd AS [Test Type], Test.SpecificityTxt AS [Specificity], Test.ModerateRiskAntibodyTxt AS [Moderate Risk Antibody], Test.LowRiskAntibodyTxt AS [Low Risk Antibody], Test.ReportableCommentsTxt AS [Reportable Comments] FROM  Patient Patient INNER JOIN Sample Sample ON  Sample.PatientId = Patient.PatientId INNER JOIN Test Test ON  Test.SampleId = Sample.SampleId WHERE  Patient.HospitalID = '" . mrnum . "' AND Sample.SampleNbr = '" . sid . "' AND Test.TestTypeCd IN ('A-L-SAI TRTD','A-L-SAII TRTD','A-L-SAI MTH','A-L-SAII MTH') ORDER BY  Sample.SampleDt Asc"
	result := ADOSQL(ConnectString_HT, Query_HT)
	rmi := result.MaxIndex()-1 ; the -1 excludes the the column headers from the output
	Loop, %rmi%
	{
		n := A_Index+1
		sabres .= result[n][6] "|`n|" result [n][7] "|`n|" result [n][8] "|`n|" result [n][9] "|x|`n"
	}

	sabres := remove_dupes(sabres)
	
	msgbox % sabres
--------------
qx := get_sab_sid(631978)
StringSplit, sidr, qx, `,
MsgBox % qx

boo := get_sab_result(sidr1)
MsgBox % boo
----------

"SELECT SAMPLE.SampleIDName, WELL.NC1, WELL.PC1, WELL_RESULT.ResultType, WELL.NC2, WELL.PC2, WELL_RESULT.Value01 FROM WELL LEFT OUTER JOIN SAMPLE ON WELL.SampleID = SAMPLE.SampleID AND WELL.IsActive = 1 AND SAMPLE.IsActive = 1 LEFT OUTER JOIN WELL_RESULT ON WELL.WellID = WELL_RESULT.WellID
WHERE     (1 = 1) AND (WELL_RESULT.ResultType IN ('54', '59')) AND (SAMPLE.SampleIDName LIKE '" . sid . "%') ORDER BY WELL.AnalysisDT DESC, SAMPLE.SampleIDName, WELL.WellPosition, WELL_RESULT.ResultType, WELL_RESULT.Value01"
---------
Transplant HT query 
Select  Patient.HospitalID AS [MR Number], TransplantHistory.TransplantDt AS [Event Date], TransplantHistory.EventCd AS [Event], TransplantHistory.DonorNm AS [Donor Name], TransplantHistory.DonorId AS [Donor ID] FROM  Patient Patient INNER JOIN TransplantHistory TransplantHistory ON  TransplantHistory.PatientId = Patient.PatientId WHERE  TransplantHistory.EventCd = 'Transplant' ORDER BY  TransplantHistory.TransplantDt Desc
---------
bla := get_sab_fusion("17-42831")
MsgBox, %bla%
--------
sidr = 17-15237
locr = DRB1

Query_Fusion := "SELECT     SAMPLE.SampleIDName, WELL_RESULT.ResultType, WELL_RESULT.Value01, WELL_RESULT.Value02, WELL_RESULT.Value03 FROM WELL LEFT OUTER JOIN SAMPLE ON WELL.SampleID = SAMPLE.SampleID AND WELL.IsActive = 1 AND SAMPLE.IsActive = 1 LEFT OUTER JOIN WELL_RESULT ON WELL.WellID = WELL_RESULT.WellID WHERE (1 = 1) AND (SAMPLE.SampleIDName LIKE '" . sidr . "%') AND (WELL_RESULT.ResultType IN ('08','06')) AND (WELL_RESULT.Value03 = N'0001') AND (WELL_RESULT.Value02 = '" . locr . "') ORDER BY WELL.AnalysisDT DESC, SAMPLE.SampleIDName, WELL.WellPosition, WELL_RESULT.ResultType, WELL_RESULT.Value01"
result := ADOSQL(ConnectString_FUSION, Query_Fusion)
rmi := result.MaxIndex()-1 ; the -1 excludes the the column headers from the output
Loop, %rmi%
{
	n := A_Index+1
	restype := result[n][2]
	if restype = 06
		allele .= result[n][3] "`n"
	if restype = 08
		sero .= result[n][3] "`n"
}
allele := remove_dupes(allele)
sero := remove_dupes(sero)
final = %sero%%allele%
Msgbox, %final%
-----
bla := get_typing_ht("AEGC410","DPB2")
------

https://portal.unos.org/DonorNet/DonorRemoteDownload.aspx?don_id=AEGC410