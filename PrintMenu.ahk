Gui, Font, s10, MS Shell Dlg
Gui, Color, D4D0C8, FFFFFF
Gui, Add, Button, x10 y10 w80 h30,Renal DD Virtual
Gui, Add, Button, x10 y50 w80 h30,Renal DD XM
Gui, Add, Button, x10 y90 w80 h30,HL DD Virtual
Gui, Add, Button, x10 y130 w80 h30,HL DD XM
Gui, Add, Button, x10 y170 w80 h30,Comm Sheet
Gui, Add, Button, x10 y210 w80 h30,Flow XM Worksheet
Gui, Add, Button, x10 y250 w80 h30,SSO Worksheet
Gui, Show, w133 h293, PrintMenu
Return

GuiClose:
ExitApp

ButtonRenalDDVirtual:
Run, print T:\Administrative Forms starting 2006\Deceased Donor Worksheets\Renal DD Virtual Crossmatch Worksheet 04-11-12.doc
Return

ButtonRenalDDXM:
Run, print T:\Administrative Forms starting 2006\Deceased Donor Worksheets\Renal Deceased Donor XM Workup Effective 04-05-12.doc
Return

ButtonHLDDVirtual:
Run, print T:\Administrative Forms starting 2006\Deceased Donor Worksheets\Cardiothoracic DD Virtual Crossmatch Worksheet 03-30-12.doc
Return

ButtonHLDDXM:
Run, print T:\Administrative Forms starting 2006\Deceased Donor Worksheets\Cardiothoracic DD XM Workup Effective 04-05-12.doc
Return

ButtonCommSheet:
Run, print T:\_Deceased Donor Workup\Communication sheet DD.doc
Return

ButtonFlowXMWorksheet:
Run, print T:\Administrative Forms starting 2006\Lab procedures worksheets\Flow Crossmatch Worksheet 11-01-12.doc
Return

ButtonSSOWorksheet:
Run, print T:\Administrative Forms starting 2006\Lab procedures worksheets\HLA typing Solid Organ Cover Sheet 7-15-13.doc
Return
