;SCRIPT OPTIMIZATION block- Makes sure the script will run as an .exe (start)
SetWorkingDir, %A_ScriptDir%
#NoEnv ;Avoids checking empty variables to see if they are environment variables
#NoTrayIcon ;Disables the showing of a tray icon.
#SingleInstance Force ;Replaces the old instance automatically
#Include HLALib.ahk ;Histotrac db connection and queries
;SCRIPT OPTIMIZATION block (end)

;GUI block (start)
Gui, Add, Edit, x30 y20 w60 h20 vmrn,
Gui, Add, Text, x126 y27 w180 h90 vstatus_info, 
Gui, Add, Edit, x330 y20 w120 h20 vorder_name, 
Gui, Add, Button, x346 y67 w100 h30 default glog vshowtime, Show Order
Gui, Add, Text, x30 y5 w59 h14 , MRN
Gui, Add, Text, x331 y6 w100 h14 , Test to be Ordered
Gui, Add, GroupBox, x116 y7 w200 h120 , Patient Info
Gui, Show, x344 y268 h135 w459, Monthly Sample Ordering
Return
;GUI block (end)

;FUNCTIONS block (start);

Is3MonthAgo(unit)
{
	uDate := dateParse(unit)
	month_diff := A_Now
	EnvSub, month_diff, %uDate%, days
	month_diff := floor(month_diff*0.0328767)
	if month_diff >= 3
		is3MonthsAgo = true
	else
		is3MonthsAgo = false
	if !unit
		is3MonthsAgo = true
	return is3MonthsAgo
}

Is6MonthAgo(unit)
{
	uDate := dateParse(unit)
	month_diff := A_Now
	EnvSub, month_diff, %uDate%, days
	month_diff := floor(month_diff*0.0328767)
	if month_diff >= 5
		is6MonthsAgo = true
	else
		is6MonthsAgo = false
	if !unit
		is6MonthsAgo = true
	return is6MonthsAgo
}

dateParse(str)
{
	StringSplit, unit, str, /
	if StrLen(unit3) < 3
		unit3 = 20%unit3%
	if unit1 < 10 
		unit1 = 0%unit1%
	if unit2 < 10
		unit2 = 0%unit2%
	pDate = %unit3%%unit1%%unit2%
	return pDate
}

orderPredict(mrnum)
{
	all_info := get_ab_dates_HT(mrnum)
	ptname := all_info[1]
	if ptname = `,
	{
		final := ["MRN not found","","","",""]
		return final
	}
	sab_dates := all_info[3]
	StringSplit, sabdate, sab_dates, `n
	fc_dates := all_info[4]
	StringSplit, fcdate, fc_dates, `n
	ptstatus := all_info[2]
	if sabdate0 = 2
	{
		final_order := "HLA-LUM SAB I/II MTH"
		final := [ptname,ptstatus,sabdate1,fcdate1,final_order]
		return final
	}
	if ptstatus = 98`%+ PRA
	{
		sab_criteria := Is3MonthAgo(sabdate1)
		if sab_criteria = true
			final_order := "HLA-LUM SAB I/II MTH"
		else
			final_order := "SERUM STORAGE"
	}
	else
	{
		sab_criteria := Is6MonthAgo(sabdate1)
		fc_criteria := Is3MonthAgo(fcdate1)
		if sab_criteria = true
			final_order := "HLA-LUM SAB I/II MTH"
		if sab_criteria = false
		{
			if fc_criteria = true
				final_order := "HLA-FC PRA I/II MTH"
			else
				final_order := "SERUM STORAGE"
		}
	}
	final := [ptname,ptstatus,sabdate1,fcdate1,final_order]
	return final
}
;FUNCTIONS block (end)

log:
Gui,Submit,Nohide
mrnlen := StrLen(mrn)
if (!mrn or mrnlen < 6)
{
	GuiControl,,status_info,Invalid MRN!
	GuiControl,,order_name,""
	Exit
}
displaybox := orderPredict(mrn)
ptname := displaybox[1]
ptstatus := displaybox[2]
last_sab := displaybox[3]
last_fc := displaybox[4]
ordering := displaybox[5]
GuiControl,,status_info,NAME`: %ptname%`n`nSTATUS`: %ptstatus%`n`nLAST SAB DATE`: %last_sab%`n`nLAST FCPRA DATE`: %last_fc%
GuiControl,,order_name,%ordering%
GuiControl,Focus,mrn
Send, ^a
return


GuiClose:
ExitApp
