X1 := ComObjGet("T:\Personal DS\Work Stuff\Worksheets_RenalVXM report.xlsx")
m := ComObjMissing()
oWorkbook3 := ComObjGet("C:\Users\dsaini\Downloads\Fusion_all_021215.csv")
oExcelsheet := oWorkbook3.Worksheets(1)
oRange := oExcelsheet.UsedRange

a := X1.Worksheets(2).Range("N6").Text
b := X1.Worksheets(2).Range("U6").Text
bw := X1.Worksheets(2).Range("AB6").Text
cw := X1.Worksheets(2).Range("AI6").Text
dr := X1.Worksheets(2).Range("AP6").Text
dw := X1.Worksheets(2).Range("AW6").Text
dqb := X1.Worksheets(2).Range("BE6").Text
dqa := X1.Worksheets(2).Range("BM6").Text
dpb := X1.Worksheets(2).Range("BT6").Text
dpa := X1.Worksheets(2).Range("CD6").Text

mrn1 := X1.Worksheets(2).Range("S12").Value
mrn2 := X1.Worksheets(2).Range("S15").Value
mrn3 := X1.Worksheets(2).Range("S18").Value
mrn4 := X1.Worksheets(2).Range("S21").Value
mrn5 := X1.Worksheets(2).Range("S24").Value
mrn6 := X1.Worksheets(2).Range("S27").Value
mrn7 := X1.Worksheets(2).Range("S30").Value
mrn8 := X1.Worksheets(2).Range("S33").Value
mrn9 := X1.Worksheets(2).Range("S36").Value
mrn10 := X1.Worksheets(2).Range("S39").Value

aoc := X1.Worksheets(2).Range("BA1").Text
unos := X1.Worksheets(2).Range("AF1").Text



#r::
count = 1
Loop
{
	mrn := Ceil(mrn%count%)

	if !mrn  
		break

	f := oRange.Find[mrn, m, m, xlWhole:=true]
	found := f.text
	a1p := f.offset(0,32).text
	a2p := f.offset(0,33).text
	b1p := f.offset(0,34).text
	b2p := f.offset(0,35).text
	bw1p := f.offset(0,12).text
	bw2p := f.offset(0,13).text
	c1p := f.offset(0,36).text
	c2p := f.offset(0,37).text
	dr1p := f.offset(0,38).text
	dr2p:= f.offset(0,39).text
	dqb1p := f.offset(0,42).text
	dqb2p := f.offset(0,43).text
	dqa1p := f.offset(0,26).text
	dqa2p := f.offset(0,27).text
	dw1p := f.offset(0,40).text
	dw2p := f.offset(0,41).text
	dpb1p := f.offset(0,28).text
	dpb2p := f.offset(0,29).text
	dpa1p := f.offset(0,30).text
	dpa2p := f.offset(0,31).text
	


	WinWait, AS400.rsf - Reflection - IBM 5250 Terminal - \\Remote, 
	IfWinNotActive, AS400.rsf - Reflection - IBM 5250 Terminal - \\Remote, , WinActivate, AS400.rsf - Reflection - IBM 5250 Terminal - \\Remote, 
	WinWaitActive, AS400.rsf - Reflection - IBM 5250 Terminal - \\Remote, 
	Send, 02%mrn%
	Sleep, 750
	Send, {Enter}
	Sleep, 750
	Send, {CTRLDOWN}a{CTRLUP}{CTRLDOWN}c{CTRLUP}
	Sleep, 750
	Send, {F2}
	FileAppend, %Clipboard%, T:\Personal DS\AHK\AS400files\VXMReport_%mrn%.txt
	FileAppend, 
	(


	
	

	AOC#: %aoc%            
	UNOS#: %unos%
	Donor Typing:					Patient Typing:
		A %a%					%a1p% %a2p%
		B %b%					%b1p% %b2p%
		Bw %bw%					%bw1p% %bw2p%
		Cw %cw%					%c1p% %c2p%
		DR %dr%					%dr1p% %dr2p% 
		DRw %dw%					%dw1p% %dw2p%
		DQB %dqb%					%dqb1p% %dqb2p%
		DQA %dqa%					%dqa1p% %dqa2p%
		DPB %dpb%					%dpb1p% %dpb2p%
		DPA %dpa%					%dpa1p% %dpa2p%



	DSAs: ___________________________________________


	Prev TX (y/n):_____ Shared MM:___________________
				

	VXM: __________________ RR#______________________

	
	Comments: _______________________________________


	_________________________________________________



			Tech: _________

			Date: _________

	), T:\Personal DS\AHK\AS400files\VXMReport_%mrn%.txt
	
	
	Run, print T:\Personal DS\AHK\AS400files\VXMReport_%mrn%.txt
	Sleep, 1000
	FileDelete, T:\Personal DS\AHK\AS400files\VXMReport_%mrn%.txt

	count += 1
	if count > 10
		break
	
}


ExitApp
