; PMCC Project Macro (Version 2.12)
; AutoHotkey Version: 1.x
; Language:       English
; Platform:       Win9x/NT
; Author:         Steve Shin <steve.shin@philips.com>
; Modified:       Mike Gamatero <michael.gamatero@philips.com>, Amy Lee <amy.lee@philips.com>
; 0.2 Changes: Default planchecker is now blank, added drop down for project type
; 0.3 Changes: Changed format for time being recorded in Excel log, Added Complexity, ITL Name
; 0.4 Changes: Adjusted ITL Name drop-down list to allow more entries, Changed Project Types
; 0.5 Changes: Even more entries for the ITL Name drop-down, Project folders are now moved to a !Completed directory during check-in, "first run" updated accordingly
; 1.0 Changes: Reformat, reduction.
; 1.1 Changes: Rev Number is a ComboBox, Check Out no longer recopies the created revision in the O Drive, just creates an empty folder
; 1.3 Changes: Moved location of Default Project Folder to shared O: drive, Added button to import CAD folders
; 1.4 Changes: Added Import Frontpage, Import Network Visio buttons; Removed MOC HTML file output

; Modified by Chris Ortiz
; 1.5 Changes: Fine/Replace CMS to PCCI, Server Remapped to to P:, Project folder now in C:PCCI, removed numerous buttons and resized existing buttons, lengthen MOC Box

; Modified by Mike Gamatero
; 2.0 Changes :  Customer ID is now the top folder in project heiarchy. Below Customer ID are the _CAD, _drawing and MOC folders.  Logic added so NO project can be checked in if a Customer ID
;                is checked in through another MOC.  Deleted Import CAD Folders, Import Frontpage, and Cancel Project functionalities.
; 2.1 Changes :  Added Open folder via Customer ID only, without MOC.
; 2.2 Changes :  Upon initial project creation, user is prompted for Address.  MOCID, CustID, and Address stored in proj_ini.txt and Masterfile.txt file
; 2.3 Changes :  Users able to check out more than 1 MOC, provided they are the original holder of the CUSTID folder only!  (not fixed)
; 2.4 Changes :  Added button so we can import pre-sales template into _drawing folder.  
;                Script fixed so _drawing and _CAD folders are not overwritten when multiple MOCs are being checked in/out.  Macro knows if there is an MOC under the same CustID already checked it,
;                it will not pull whatever was from the _drawing and _CAD folders in the P drive.
; 2.5 Changes :  When Import Presales, Import Visio, and Check Out functions are run, folder C:Projects_PCCI\%CustIDNumber%\_drawing is opened.  This is done so planner can see if _drawing folder
;                already has previous drawings
;             :  Added Check for Previous Drawings icon, which opens C:Projects_PCCI\%CustIDNumber%\_drawing folder
;             :  Added presales to project type
; 2.6 Changes :  Added button so we can import config sheets.xls into _drawings folder.  Located at O:\Symbols\10 - PMCC\Config Sheets.xls.  For Preinstall drawings
; 2.7 Changes :  Added button and script so we can automate "Master-Child" text files (via input boxes)
; 2.8 changes :  "Master-Child" folder input is done through O:\Symbols\10 - PMCC\Macro\Master-Child.xlsm.   
;             :  Create MasterChild.txt - Folder input done through excel file.  Excel macro outputs a txt file in C:Projects_PCCI\
;             :  View MasterChild.txt - Views contents in MOC folder
;             :  Transfer MasterChild Txt File to Proj Folder - Transfer MasterChild.txt from to CustomerID level
; Modified by Amy Lee
; 2.9 Changes :  Added Import CAD button to import CAD template to avoid over-writing the Titleblocks from existing projects.
; 2.10 Changes:  Removed Hospital address and site name prompts. Added ITL box. When checking out a project, automatically creates a text file that lists the Cust ID, MOC, Requestor, Rev, Project Type, Planner.
; 2.11 Changes:  Added pop-up calendar for due date, and auto complete for Requestors list.
; 2.12 Changes:  Added autofill for easier check-in process
;--------------------------------------------------------------------

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
FormatTime, MyFormatTime
FormatTime, MyStatusBarTime,,h:mmtt
CheckOutCheckString = Checked Out
CheckInCheckString = Checked In
Menu, Tray, Icon, O:\Symbols\10 - PMCC\Macro\PS.ico


;-----------------First Run------------------------------------ 
IfNotExist, C:\PCCI\
{
   FileCreateDir, C:\PCCI\
}
IfNotExist, C:\PCCI\ReadOnly
{
	FileCreateDir, C:\PCCI\ReadOnly
}
IfNotExist, C:\Projects_PCCI\!Completed
{
	FileCreateDir, C:\Projects_PCCI\!Completed
}
IfNotExist, C:\PCCI\userid.txt
{
   InputBox, FirstTimeUserName, Install, Enter your name:,  
   FileAppend, %FirstTimeUserName%, C:\PCCI\userid.txt
   FileAppend, %FirstTimeUserName%`n, O:\Symbols\10 - PMCC\Macro\alluserid.txt
}
IfNotExist, C:\Projects_PCCI
{
	FileCreateDir, C:\Projects_PCCI
}
IfNotExist, C:\Projects_PCCI\MOC Project Folder2
{
   FilecopyDir, O:\Symbols\10 - PMCC\Project Folder2, C:\Projects_PCCI\MOC Project Folder, 0
}
IfNotExist, C:\local_proj_ini
{
	FileCreateDir, C:\local_proj_ini
}

;-----------------Startup------------------------------------ 
ArrayCount = 0
Loop, Read, O:\Symbols\10 - PMCC\Macro\alluserid.txt  ; This loop retrieves each line from the file, one at a time.
{
   ArrayCount += 1                                                      ; Keep track of how many items are in the array.
   Array%ArrayCount% := A_LoopReadLine                                  ; Store this line in the next array element.
}
ArrayCount2 = 0
Loop, Read, O:\Symbols\10 - PMCC\Macro\ITLList.txt    ; This loop retrieves each line from the file, one at a time.
{
   ArrayCount2 += 1                                                     ; Keep track of how many items are in the array.
   ITLArray%ArrayCount2% := A_LoopReadLine                              ; Store this line in the next array element.
}
FileRead, DefaultUserName, C:\PCCI\userid.txt



;-----------------GUI---------------------------------- 
Gui, Font, , 
Gui, Add, GroupBox, x6 y7 w270 h235 , Project Info
Gui, Add, Text, x26 y30 w80 h20 +Right, Cust ID Number:
Gui, Add, Text, x26 y60 w80 h20 +Right, MOC Number:
Gui, Add, Text, x26 y90 w80 h20 +Right, Requestor:
Gui, Add, Text, x26 y120 w80 h20 +Right, Revision:
Gui, Add, Text, x36 y150 w70 h20 +Right, Planner:
Gui, Add, Text, x36 y180 w70 h20 +Right, Type:

Gui, Add, Text, x36 y215 w70 h20 +Right, Due Date:

Gui, Add, Edit, x116 y30 w145 h20 vCustIDNumber 
Gui, Add, Edit, x116 y60 w145 h20 vMOCNumber 
Gui, Add, ComboBox, x116 y90 w145 h20 R15 vRequestor gAutoComplete, ---|aamir.siddiqui|abhijeet.bhat|adam.brewer|adel.afzal|aj.girardin|alan.babiarz|amy.lee|anthony.glass|anthony.locascio|arleigh.murrell|benjamin.carlos|bert.ferris|bob.aaron|brent.davidson|brian.jackson|bruce.vaal|bryan.morgan|carlos.banda|carlos.casteel|carol.coupet|cathy.marinelli|craig.vandegrift|cristie.nutter|dall.howard|dan.field|dave.codispoti|dave.muller|dave.roden|david.borough|deJuan.simpson|diana.minks|diane.blough|dirk.hummel|dorri.barker|douglas.peterson|duane.kellogg|erica.daley-bruhier|jeff.foster|jered.henry|ed.snyder|eric.huff|florentino.chavez|fotis.papadopoulos|frank.curcio|garry.dragoo|gary.ramey|gary.scott|gerald.forlenza|gary.scott|greg.nies|harry.marquass|harvey.stroyan|herman.rojas|james.barnes|james.flory|jared.lilly|jason.rosenzweig|jason.walter|jeff.foster|jeff.hunter|jeffrey.l.smith|jennifer.globke|jennifer.olszewski|jered.henry|jeremy.plato|jerry.forlenza|jim.cockerham|jim.favaron|jim.laramie|james.pacheco|jared.henry|joan.jenkins|joe.hudson|johan.marte|john.berry|john.dann|john.weiner|jon.crispin|jonas.mckenzie|jorge.lan|jose.gularte|judy.mahoney|kevin.tecce|kim.legrand|lane.boolen|lawrence.blacharski|lewis.ayers|linda.jenny|lisa.crowder|loreto.rodriguez|luis.navarro|matt.notter|mark.massey|michael.hosick|michael.kauffman|mike.durbin|mike.fosco|mike.labranche|mike.rushing|mitch.usrey|mohsin.tejani|nick.saccomanno|oleg.langer|paul.monckton|phil.robertson|randy.nosrati|randy.shafor|ray.west|rich.garber|rob.busby|robbie.spinney|robert.coffelt|robert.mcnulty|robert.sneed|ron.carreira|ron.diecker|scott.keyes|shane.elting|shirish.agnihotri|steve.byrd|steven.meliet|ted.buchel|thomas.willis|tim.bodien|tim.moon|tim.thurston|tom.hoskinson|tony.dimatteo|vahid.salehi|varghese.mathews|victor.avila|wayne.burden|william.gall|
Gui, Add, Edit, x116 y120 w120 h20 vRevNumber
Gui, Add, DropDownList, x116 y150 w120 h21 R7 vPlannerName, %Array1%|%Array2%|%Array3%|%Array4%|%Array5%|%Array6%|%Array7%|%DefaultUserName%
Gui, Add, DropDownList, x116 y180 w120 h21 R4 vProjectType, RF|Install|AsBuilt|PreSales|PreInstall
Gui, Add, DateTime, x116 y215 w80 h20 vDueDate 

Gui, Add, Button, x6 y248 w140 h30 gCheckProjectStatus, Check Project &Status
Gui, Add, Button, x146 y248 w130 h30 gOpenBacklog, Open &Backlog

Gui, Add, Button, x217 y281 w63 h30 gImportCAD, Import &CAD
Gui, Add, Button, x154 y281 w63 h30 gImportVisio, Import &Visio
Gui, Add, Button, x91 y281 w63 h30 gImportPreSales, Import &Presales
Gui, Add, Button, x6 y281 w85 h30 gImportConfigsheetsxls, Import &Configsheets.xls
Gui, Add, Button, x6 y314 w150 h30 gOpenDrawingsFolder, Check For Existing &Drawing 
Gui, Add, Button, x158 y314 w116 h30 gSearchForMOCbyCustID, Check For CustID via &MOC

Gui, Add, GroupBox, x6 y365 w270 h95 , Check Out
Gui, Add, Text, x16 y380 w265 h40 , Create a project in the archive, and copies the revision to C:. Other users cannot check out this project.
Gui, Add, Button, x16 y415 w130 h40 gCheckOut, Check &Out
Gui, Add, Button, x146 y415 w120 h18 gOpenLocalFolder, Open &Local Folder
Gui, Add, Button, x146 y437 w120 h18 gOpenCustIDFolder, Open &Local CustID 

Gui, Add, GroupBox, x6 y465 w270 h100 , Check In
Gui, Add, Text, x16 y482 w250 h40 , Copies the completed project to the archive and updates the Project Log.
Gui, Add, Button, x16 y517 w130 h40 gCheckIn, Check &In
Gui, Add, Button, x146 y517 w120 h18 gOpenArchivedFolder, Open &Archived Folder
Gui, Add, Button, x146 y539 w120 h18 gOpenArchivedCustIDFolder, Open &Archived CustID 

Gui, Add, GroupBox, x6 y570 w270 h90 , Outlook Tools
Gui, Add, Button,  x16 y584 w130 h40 gAutofillByMOC, Autofill

Gui, Add, GroupBox, X6 y665 w270 h60 , Master-Child 
Gui, Add, Text, x16 y680 w260 h50 , Automates "Master-Child" projs creation.
Gui, Add, Button, x16 y695 w220 h18 gCreateMasterChildTxt,  Create &MasterChild.txt(C:\Projects_PCCI)


Gui, Show, x128 y154 h730 w285, PCCI Project Macro (Ver 2.12)

Return

;------Formats the Due Date to MM/DD/YYYY-------------------------------------

DateRoutine:

Year := SubStr(DueDate, 1, 4)
Month := SubStr(DueDate, 5, 2)
Day := SubStr(DueDate, 7, 2)
NewDueDate = %Month%/%Day%/%Year%


Return


;------Auto Complete for the Requestor Field--------------------------------------------

AutoComplete:
If GetKeyState("Delete") or GetKeyState("Backspace")
	Return
SetControlDelay, -1
SetWinDelay, -1
GuiControlGet, h, Hwnd, %A_GuiControl%
ControlGet, haystack, List, , , ahk_id %h%
GuiControlGet, needle, , %A_GuiControl%
lf = `n
StringMid, text, haystack, pos := InStr(lf . haystack, lf . needle)
	, InStr(haystack . lf, lf, false, pos) - pos
If text !=
{
	ControlSetText, , %text%, ahk_id %h%
	ControlSend, , % "{Right " . StrLen(needle) . "}+^{End}", ahk_id %h%
}
Return

;-----------------------------------------------------
ImportCAD:
Gui, Submit, NoHide
If CustIdNumber =
{
   MsgBox, Enter a Cust ID number!
   return
}
If MOCNumber =
{
   MsgBox, Enter an MOC number!
   return
}
If RevNumber =
{
   MsgBox, Enter a revision number!
   return
}
if PlannerName=""
{
   MsgBox, Select a our name!
   Return
}

IfExist, C:\Projects_PCCI\%CustIDNumber%\_CAD\drawing\xref_Titleblock - PCCI 11x17.dwg
{
   MsgBox, 4, Import xref_Titleblock - PCCI 11x17.dwg, CAD Template already exists. OK to overwrite?
   IfMsgBox No
   {
     return
   }
  
   FileCopy, O:\Symbols\10 - PMCC\DWG\Symbols\Equipment Layout\#Floor-Dept name here# Telemetry.dwg, C:\Projects_PCCI\%CustIDNumber%\_CAD\drawing\#Floor-Dept name here# Telemetry.dwg, 1
   FileCopy, O:\Symbols\10 - PMCC\DWG\Symbols\Equipment Layout\xref_Titleblock - PCCI 11x17.dwg, C:\Projects_PCCI\%CustIDNumber%\_CAD\drawing\xref_Titleblock - PCCI 11x17.dwg, 1
   FileCopy, O:\Symbols\10 - PMCC\DWG\Symbols\Equipment Layout\xref_Titleblock - PCCI 18x24.dwg, C:\Projects_PCCI\%CustIDNumber%\_CAD\drawing\xref_Titleblock - PCCI 18x24.dwg, 1
   FileCopy, O:\Symbols\10 - PMCC\DWG\Symbols\Equipment Layout\xref_Titleblock - PCCI 24x36.dwg, C:\Projects_PCCI\%CustIDNumber%\_CAD\drawing\xref_Titleblock - PCCI 24x36.dwg, 1
   FileCopy, O:\Symbols\10 - PMCC\DWG\Symbols\Equipment Layout\Cover Page.dwg, C:\Projects_PCCI\%CustIDNumber%\_CAD\drawing\Cover Page.dwg, 1
   MsgBox, Project Folder for Revision %RevNumber% now contains imported CAD files.
   Run, C:\Projects_PCCI\%CustIDNumber%\_CAD\drawing
   
   return
}

IfNotExist, C:\Projects_PCCI\%CustIDNumber%\_CAD\drawing\xref_Titleblock - PCCI 11x17.dwg
{
   
   FileCopy, O:\Symbols\10 - PMCC\DWG\Symbols\Equipment Layout\#Floor-Dept name here# Telemetry.dwg, C:\Projects_PCCI\%CustIDNumber%\_CAD\drawing\#Floor-Dept name here# Telemetry.dwg, 1
   FileCopy, O:\Symbols\10 - PMCC\DWG\Symbols\Equipment Layout\xref_Titleblock - PCCI 11x17.dwg, C:\Projects_PCCI\%CustIDNumber%\_CAD\drawing\xref_Titleblock - PCCI 11x17.dwg, 1
   FileCopy, O:\Symbols\10 - PMCC\DWG\Symbols\Equipment Layout\xref_Titleblock - PCCI 18x24.dwg, C:\Projects_PCCI\%CustIDNumber%\_CAD\drawing\xref_Titleblock - PCCI 18x24.dwg, 1
   FileCopy, O:\Symbols\10 - PMCC\DWG\Symbols\Equipment Layout\xref_Titleblock - PCCI 24x36.dwg, C:\Projects_PCCI\%CustIDNumber%\_CAD\drawing\xref_Titleblock - PCCI 24x36.dwg, 1
   FileCopy, O:\Symbols\10 - PMCC\DWG\Symbols\Equipment Layout\Cover Page.dwg, C:\Projects_PCCI\%CustIDNumber%\_CAD\drawing\Cover Page.dwg, 1
   MsgBox, Project Folder for Revision %RevNumber% now contains imported CAD files.
   Run, C:\Projects_PCCI\%CustIDNumber%\_CAD\drawing
   
   return
}


;-----------------------------------------------------
ImportVisio:
Gui, Submit, NoHide
If CustIdNumber =
{
   MsgBox, Enter a Cust ID number!
   return
}
If MOCNumber =
{
   MsgBox, Enter an MOC number!
   return
}
If RevNumber =
{
   MsgBox, Enter a revision number!
   return
}
if PlannerName=""
{
   MsgBox, Select a our name!
   Return
}

IfExist, C:\Projects_PCCI\%CustIDNumber%\_drawing\Network Template.vsd
{
   MsgBox, 4, Import Network Template Visio File, Network Template.vsd already exists. OK to overwrite?
   IfMsgBox No
   {
     return
   }
   
   FileCopy, O:\Symbols\10 - PMCC\VSD\Visio Templates for Project Folder\Network Template.vsd, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\_drawing\Network Template.vsd, 1
   MsgBox, Project Folder for Revision %RevNumber% now contains imported Network Template.
   
   return
}

IfNotExist, C:\Projects_PCCI\%CustIDNumber%\%MOCNumber%\%RevNumber%
{
   MsgBox, Project does not exist. Please make sure project is checked out and try again.
   return
}
   
IfExist, C:\Projects_PCCI\%CustIDNumber%\%MOCNumber%\%RevNumber%
{
   FileCopy, O:\Symbols\10 - PMCC\VSD\Visio Templates for Project Folder\Network Template.vsd, C:\Projects_PCCI\%CustIDNumber%\_drawing\Network Template.vsd, 1
   MsgBox, Project Folder for Revision %RevNumber% now contains imported Visio Network Template.
   Run, C:\Projects_PCCI\%CustIDNumber%\_drawing
   return
}




;-----------------------------------------------------
ImportPresales:
Gui, Submit, NoHide
If CustIdNumber =
{
   MsgBox, Enter a Cust ID number!
   return
}
If MOCNumber =
{
   MsgBox, Enter an MOC number!
   return
}
If RevNumber =
{
   MsgBox, Enter a revision number!
   return
}
if PlannerName=""
{
   MsgBox, Select a our name!
   Return
}

IfExist, C:\Projects_PCCI\%CustIDNumber%\_drawing\Pre-sale Template.vsd
{
   MsgBox, 4, Import Network Template Visio File, Pre-sale Template.vsd already exists. OK to overwrite?
   IfMsgBox No
   {
     return
   }
   
   FileCopy, O:\Symbols\10 - PMCC\VSD\Visio Templates for Project Folder\Pre-sale Template.vsd, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\_drawing\Pre-Sale Template.vsd, 1
   MsgBox, Project Folder for Revision %RevNumber% now contains Pre-sale Template.
   
   return
}

IfNotExist, C:\Projects_PCCI\%CustIDNumber%\%MOCNumber%\%RevNumber%
{
   MsgBox, Project does not exist. Please make sure project is checked out and try again.
   return
}
   
IfExist, C:\Projects_PCCI\%CustIDNumber%\%MOCNumber%\%RevNumber%
{
   FileCopy, O:\Symbols\10 - PMCC\VSD\Visio Templates for Project Folder\Pre-Sale Template.vsd, C:\Projects_PCCI\%CustIDNumber%\_drawing\Pre-Sale Template.vsd, 1
   MsgBox, Project Folder for Revision %RevNumber% now contains Pre-sale Template.
   Run, C:\Projects_PCCI\%CustIDNumber%\_drawing
   return
}

;-----------------------------------------------------



;-----------------------------------------------------
ImportConfigsheetsxls:
Gui, Submit, NoHide
If CustIdNumber =
{
   MsgBox, Enter a Cust ID number!
   return
}
If MOCNumber =
{
   MsgBox, Enter an MOC number!
   return
}
If RevNumber =
{
   MsgBox, Enter a revision number!
   return
}
if PlannerName=""
{
   MsgBox, Select a our name!
   Return
}

IfExist, C:\Projects_PCCI\%CustIDNumber%\_drawing\Config Sheets.xls
{
   MsgBox, 4, Import Config Sheets.xls, Config Sheets.xls already exists. OK to overwrite?
   IfMsgBox No
   {
     return
   }
   
   FileCopy, O:\Symbols\10 - PMCC\Config Sheets.xls, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\_drawing\Config Sheets.xls, 1
   MsgBox, Project Folder for Revision %RevNumber% now contains Config Sheets.xls.
   
   return
}

IfNotExist, C:\Projects_PCCI\%CustIDNumber%\%MOCNumber%\%RevNumber%
{
   MsgBox, Project does not exist. Please make sure project is checked out and try again.
   return
}
   
IfExist, C:\Projects_PCCI\%CustIDNumber%\%MOCNumber%\%RevNumber%
{
   FileCopy, O:\Symbols\10 - PMCC\Config Sheets.xls, C:\Projects_PCCI\%CustIDNumber%\_drawing\Config Sheets.xls, 1
   MsgBox, Project Folder for Revision %RevNumber% now contains Config Sheets.xls.
   Run, C:\Projects_PCCI\%CustIDNumber%\_drawing
   return
}

;-----------------------------------------------------
OpenDrawingsFolder:
Gui, Submit, NoHide
If CustIdNumber =
{
   MsgBox, Enter a Cust ID number!
   return
}

Run, C:\Projects_PCCI\%CustIDNumber%\_drawing,, UseErrorLevel

if ErrorLevel
{
   MsgBox, Default folders not found. Could not find Cust ID number.
   return
}
return


;-----------------------------------------------------

CheckProjectStatus:

Gosub, CheckProjectExistence
If ProjectExists = 0   ;exit if project doesn't exist
   return

IfExist, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\Checkoutflag.txt   ;checking for instance that CUST ID is checked out by another MOC
{
	Loop,read, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\Checkoutflag.txt    ;read contents of Checkoutflag.txt and output to message box
	{
		CheckedoutbyCustIDmessage:=A_LoopReadLine
	}		

	MsgBox, %CheckedoutbyCustIDmessage%
}


IfExist, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\%MOCNumber%\log.txt
{
   FileRead, MOCProjectLog, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\%MOCNumber%\log.txt
   ;Drive, Eject ;April Fools Joke
   MsgBox, 64, Status of %MOCNumber%:, %MOCProjectLog%
   return
}

MOCProjectDirectoryLoop = 
Loop, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\%MOCNumber%\*.*, 2
{
   FileGetTime, MOCProjectDirectoryLoopTime, %A_LoopFileFullPath%
   FormatTime, MOCProjectDirectoryLoopTimeFormatted, %MOCProjectDirectoryLoopTime%
   MOCProjectDirectoryLoop = %MOCProjectDirectoryLoop% Revision %A_LoopFileName% exists and was created at %MOCProjectDirectoryLoopTimeFormatted% `n 
}

MsgBox, 64, Status of %MOCNumber%:, %MOCProjectDirectoryLoop%

return




;-----------------------------------------------------
OpenArchivedFolder:
Gosub, CheckProjectExistence
If ProjectExists = 0
   return

Run, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\%MOCNumber%
return

;-----------------------------------------------------
OpenArchivedCustIDFolder:
Gui, Submit, NoHide
If CustIdNumber =
{
   MsgBox, Enter a Cust ID number!
   return
}

Run, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%,, UseErrorLevel

if ErrorLevel
{
   MsgBox, Default folders not found. Could not find Cust ID number.
   return
}
return


;-----------------------------------------------------
OpenLocalFolder:
Gui, Submit, NoHide
If CustIdNumber =
{
   MsgBox, Enter a Cust ID number!
   return
}
If MOCNumber =
{
   MsgBox, Please enter an MOC number.
   return
}

Run, C:\Projects_PCCI\%CustIDNumber%\%MOCNumber%\%RevNumber%\,, UseErrorLevel


if ErrorLevel
{
   MsgBox, Default folders not found. Could not open Project folder.
   return
}
return



;-----------------------------------------------------
OpenCustIDFolder:
Gui, Submit, NoHide
If CustIdNumber =
{
   MsgBox, Enter a Cust ID number!
   return
}

Run, C:\Projects_PCCI\%CustIDNumber%,, UseErrorLevel

if ErrorLevel
{
   MsgBox, Default folders not found. Could not find Cust ID number.
   return
}
return

;-----------------------------------------------------
OpenBacklog:
Run, http://pww.mocwebportal.philips.com/MOC/SitePlanning/MyProjects.aspx
return

;-----------------------------------------------------

LogFileCreate:
Gui, Submit, NoHide

IfExist, P:\Site_Planning\Project_Archive_PCCI\_Project_INI\%MOCNumber%.txt
	{
		MsgBox, Customer ID and MOC already exists in Project_INI folder
		return
	}

IfNotExist, P:\Site_Planning\Project_Archive_PCCI\_Project_INI\%MOCNumber%.txt
	{
		MsgBox, Customer ID and MOC successfully added to Project_INI folder!		
	}
		
FileAppend, `n%CustIDNumber% 'n%MOCNumber%, P:\Site_Planning\Project_Archive_PCCI\_Project_INI\%MOCNumber%.txt
FileAppend, `n%MOCNumber%  :%CustIDNumber%, P:\Site_Planning\Project_Archive_PCCI\_Project_INI\Masterfile\Master_CustID_MOC.txt
MsgBox, Customer ID and MOC successfully added to Master_CUST_ID_MOC file!

return

;------Create text file to local C: Drive for Outlook Macro PC Form Auto Fill----------------------------------------
; for ini txt file in local proj folder
PCFormFileCreate:
Gui, Submit, NoHide

IfExist, C:\local_proj_ini\%MOCNumber%.txt
	{
		FileDelete, C:\local_proj_ini\%MOCNumber%.txt
	}
	
GoSub, DateRoutine
	
FileAppend, `n%CustIDNumber%`n%MOCNumber%`n%RevNumber%`n%ProjectType%`n%Requestor%`n%NewDueDate%, C:\local_proj_ini\%MOCNumber%.txt

return

;-------------------------------------------------------

SearchForMOCbyCustID:
Gui, Submit, NoHide

InputBox, MOCLookup , Lookup by MOC#, Enter MOC#, , ,130 , , , , N/A

IfNotExist, P:\Site_Planning\Project_Archive_PCCI\_Project_INI\%MOCLookup%.txt
	{
		MsgBox, MOC doesn't exist
		return
	}

IfExist, P:\Site_Planning\Project_Archive_PCCI\_Project_INI\%MOCLookup%.txt
	{
		FileReadLine, CustIDLookUp, P:\Site_Planning\Project_Archive_PCCI\_Project_INI\%MOCLookup%.txt, 2
		MsgBox, %MOCLookup% is under Customer ID:  %CustIDLookUp%

	}			
			
Return

;------------Button to autofill from text file to check-in project-------------------------------------------
AutofillByMOC:
Gui, Submit, noHide

InputBox, EnterMOC, Enter MOC#, Enter MOC#, , ,130 , , , , N/A

IfNotExist, C:\local_proj_ini\%EnterMOC%.txt
	{
		MsgBox, MOC doesn't exist
		return
	}
	
IfExist, C:\local_proj_ini\%EnterMOC%.txt
	{
		FileReadLine, CustID, C:\local_proj_ini\%EnterMOC%.txt, 2
		FileReadLine, MOC, C:\local_proj_ini\%EnterMOC%.txt, 3
		FileReadLine, Rev, C:\local_proj_ini\%EnterMOC%.txt, 4
		FileReadLine, Type, C:\local_proj_ini\%EnterMOC%.txt, 5
		
		
		
		IfinString,Type,PreInstall
			{
			Type := 1
			}
		IfinString,Type,Install
			{
			Type := 2
			}
		IfinString,Type,AsBuilt
			{
			Type := 3
			}
		IfinString,Type,PreSales
			{
			Type := 4
			}
			
				
		MsgBox, 0, Autofill, Complete
		
		
		
		GuiControl,, CustIDNumber, %CustID%
		GuiControl,, MOCNumber, %MOC%
		GuiControl,, RevNumber, %rev%
		GuiControl,Choose, ProjectType, %Type%
			
	}
	
Return
	

;-------------------------------------------------------

CheckOut:
Gui, Submit, NoHide
If CustIdNumber =
{
   MsgBox, Enter a Cust ID number!
   return
}
If MOCNumber =
{
   MsgBox, Please enter an MOC number.
   return
}
If RevNumber =
{
   MsgBox, Please enter a revision number.
   return
}
if PlannerName=""
{
   MsgBox, Please select a planner name.
   return
}
if DueDate = 
{
   MsgBox, Please enter a valid due date.
   return
}

IfExist, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\%MOCNumber%\log.txt
   {
     Loop, read, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\%MOCNumber%\log.txt        ; Retrieve the last line from a text file.
     {
       CurrentLogLine := A_LoopReadLine                     ; When loop finishes, this will hold the last line.
     }
     IfInString, CurrentLogLine, %CheckOutCheckString%
     {
       MsgBox, Check Out Error: %CurrentLogLine%
       return
     }
   }

IfExist, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\Checkoutflag.txt  ;This if statement enables user to check multiple MOCs under one CustID, only if user and one who checked custid out is the same.
{
	Loop,read, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\Checkoutflag.txt    ;read contents of Checkoutflag.txt and output to message box
	{
		CheckedoutbyCustIDmessage:=A_LoopReadLine
	}		

	IfInString, CheckedoutbyCustIDmessage, %PlannerName%
	{
		MsgBox, 3, Warning, %CheckedoutbyCustIDmessage% through another MOC.  Continue with Checkout?
		IfMsgBox Yes
		{
			FileDelete, C:\Projects_PCCI\%CustIDNumber%\Checkoutflag.txt
			FileDelete, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\Checkoutflag.txt
		}
		IfMsgBox No
		{
			return
		}
	} 
}


IfExist, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\Checkoutflag.txt
{
	Loop,read, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\Checkoutflag.txt    ;read contents of Checkoutflag.txt and output to message box
	{
		CheckedoutbyCustIDmessage:=A_LoopReadLine
	}		

	MsgBox, %CheckedoutbyCustIDmessage%
   Return
}


IfNotExist, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\%MOCNumber%
{
   MsgBox, 4, Check Out Process, Project does not exist. OK to create project?
   IfMsgBox No
   {
		return
   }
   
   
   FileCopyDir, O:\Symbols\10 - PMCC\Project Folder2\MOC\0, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\%MOCNumber%\%RevNumber%, 1    ;Copy file structure from O drive
   FileCopyDir, O:\Symbols\10 - PMCC\CAD, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\_CAD, 1
   FileCreateDir, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\_drawing



   FormatTime, MyFormatTime
   FileAppend, `nChecked Out by %PlannerName% working on %ProjectType% Rev %RevNumber% at %MyFormatTime%, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\%MOCNumber%\log.txt
  
   FileAppend, `n%CustIDNumber% is Currently Checked Out (by %PlannerName%), P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\Checkoutflag.txt
    
   Gosub, CheckOutCopyFolders   ;Move folders

   MsgBox, Project Revision %RevNumber% has been successfully created and downloaded to your local drive.
   GoSub, LogFileCreate
   GoSub, PCFormFileCreate
   Run, C:\Projects_PCCI\%CustIDNumber%\_drawing
   return
}


IfExist, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\%MOCNumber%
{
   IfExist, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\%MOCNumber%\log.txt
   {
     Loop, read, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\%MOCNumber%\log.txt        ; Retrieve the last line from a text file.
     {
       CurrentLogLine := A_LoopReadLine                     ; When loop finishes, this will hold the last line.
     }
     IfInString, CurrentLogLine, %CheckOutCheckString%
     {
       MsgBox, Check Out Error: %CurrentLogLine%
       return
     }
   }
   
   IfExist, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\%MOCNumber%\%RevNumber%
   {
     MsgBox, 4, Check Out Process, Revision Number %RevNumber% already exists. OK to copy existing archived folder to C:\Projects_PCCI?
     IfMsgBox No 
     {
       return
     }
     
    
     FormatTime, MyFormatTime
     FileAppend, `nChecked Out by %PlannerName% working on %ProjectType% Rev %RevNumber% at %MyFormatTime%, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\%MOCNumber%\log.txt
	 
     FileAppend, `n%CustIDNumber% is Currently Checked Out (by %PlannerName%), P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\Checkoutflag.txt

     
     Gosub, CheckOutCopyFolders   ;Move folders	

     
     MsgBox, Project Revision %RevNumber% has been successfully downloaded from the archive to your local drive.
     GoSub, LogFileCreate
	 GoSub, PCFormFileCreate
     return
   }
   
   IfNotExist, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\%MOCNumber%\%RevNumber%
   {
   
      MOCCheckOutLoop = 
      MOCCheckOutLoopTime = 

      Loop, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\%MOCNumber%\*.*, 2
      {
         if A_Index = 1
         {
            MOCCheckOutLoopTime = %A_LoopFileTimeModified%
            MOCCheckOutLoop = %A_LoopFileFullPath%
         }
         
         if MOCCheckOutLoopTime < %A_LoopFileTimeModified%
         {
            MOCCheckOutLoopTime = %A_LoopFileTimeModified%
            MOCCheckOutLoop = %A_LoopFileFullPath%
         }
      }
     ;MsgBox, %MOCCheckOutLoop% %MOCCheckOutLoopTime%
     MsgBox, 3, Check Out Process, Revision %RevNumber% does not exist. The last edited revision was %MOCCheckOutLoop%.`n`nDo you want to copy the files/folders from the old revision to the new Revision %RevNumber% folder.`nSelecting "NO" will create Revision %RevNumber% using the default folder structure (aka MOC Default).
     
     IfMsgBox Yes
     {
       FileCopyDir, %MOCCheckOutLoop%, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\%MOCNumber%\%RevNumber%
      
       FormatTime, MyFormatTime
       FileAppend, `nChecked Out by %PlannerName% working on %ProjectType% Rev %RevNumber% at %MyFormatTime%, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\%MOCNumber%\log.txt
	   
       FileAppend, `n%CustIDNumber% is Currently Checked Out (by %PlannerName%), P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\Checkoutflag.txt

       Gosub, CheckOutCopyFolders   ;Move folders
      
       
       MsgBox, Project Revision %RevNumber% has been successfully copied from the last revision and downloaded to your local drive.
       GoSub, LogFileCreate
	   GoSub, PCFormFileCreate
       Run, C:\Projects_PCCI\%CustIDNumber%\_drawing
       return
     }
     
     IfMsgBox No
     {
       FileCopyDir, O:\Symbols\10 - PMCC\Project Folder\0, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\%MOCNumber%\%RevNumber%
      
       FormatTime, MyFormatTime
       FileAppend, `nChecked Out by %PlannerName% working on %ProjectType% Rev %RevNumber% at %MyFormatTime%, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\%MOCNumber%\log.txt
       FileAppend, `n%CustIDNumber% is Currently Checked Out (by %PlannerName%), P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\Checkoutflag.txt

       
       FileCopyDir, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\%MOCNumber%\%RevNumber%, C:\Projects_PCCI\%CustIDNumber%\%MOCNumber%\%RevNumber%, 1
       FileCopy, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\%MOCNumber%\log.txt, C:\Projects_PCCI\%CustIDNumber%\%MOCNumber%\log.txt, 1
	   FileCopy, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\Checkoutflag.txt, C:\Projects_PCCI\%CustIDNumber%\Checkoutflag.txt, 1


       MsgBox, Project Revision %RevNumber% has been successfully created and downloaded to your local drive.
       GoSub, LogFileCreate
	   Run, C:\Projects_PCCI\%CustIDNumber%\_drawing
       return
     }
     
     return
     
   }
}


;-----------------------------------------------------
CheckIn:
Gui, Submit, NoHide
if PlannerName=""
{
	 MsgBox, Please select a planner name.
   Return
}
If CustIdNumber =
{
   MsgBox, Enter a Cust ID number!
   return
}


If MOCNumber =
{
   MsgBox, Please enter an MOC number.
   return
}
If RevNumber =
{
   MsgBox, Please enter a revision number.
   return
}
IfNotExist, C:\Projects_PCCI\%CustIDNumber%\%MOCNumber%\%RevNumber%
{
   MsgBox, Project and/or revision does not exist.
   return
}


IfExist, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\%MOCNumber%\log.txt		
{
   Loop, read, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\%MOCNumber%\log.txt        ; Retrieve the last line from a text file.
   	{
    		CurrentLogLine := A_LoopReadLine                     ; When loop finishes, this will hold the last line.
   	}
   		IfInString, CurrentLogLine, %CheckInCheckString%
   	{
     		MsgBox, Check In Error: %CurrentLogLine%
     		return
   	}
}

IfNotExist, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\%MOCNumber%\%RevNumber%
	{
   		MsgBox, 4, Check In Process, OK to copy folder on local drive to the archive?
   		IfMsgBox No
   	{			
     		return
   	}
   
   		GoSub, CheckInCopyAndLog
   		return
	}

IfExist, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\%MOCNumber%\%RevNumber%
{
   	MsgBox, 4, Check In Process, WARNING: Project revision already exists on the archive. Please check your Revision Number. OK to overwrite?
   	IfMsgBox No
   	{
     		return
   	}
   
GoSub, CheckInCopyAndLog
return
}

;-----------------------------------------------------

CheckInCopyAndLog:      ;Procedure to move folders for check in.  

FileRemoveDir, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\%MOCNumber%\%RevNumber%, 1
FileCopyDir, C:\Projects_PCCI\%CustIDNumber%\%MOCNumber%\%RevNumber%, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\%MOCNumber%\%RevNumber%, 1
FileCopyDir, C:\Projects_PCCI\%CustIDNumber%\_drawing, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\_drawing, 1
FileCopyDir, C:\Projects_PCCI\%CustIDNumber%\_CAD, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\_CAD, 1
FileCopy, C:\Projects_PCCI\%CustIDNumber%\MasterChild.txt, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\MasterChild.txt,1

FormatTime, MyFormatTime
FileAppend, `nChecked In by %PlannerName% working on %ProjectType% Rev %RevNumber% at %MyFormatTime%, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\%MOCNumber%\log.txt
FileAppend, `nChecked In by %PlannerName% working on %ProjectType% Rev %RevNumber% at %MyFormatTime%, C:\Projects_PCCI\%CustIDNumber%\%MOCNumber%\log.txt

FormatTime, MyCSVDate, yyyy MM dd  HH:mm
FileAppend, 
(
%MOCNumber%,%ProjectType%,%RevNumber%,%MyCSVDate%,%PlannerName%,,,,,

), P:\Site_Planning\Project_Archive_PCCI\_Project Times (Macro).csv

Filecopy, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\%MOCNumber%\log.txt, C:\Projects_PCCI\%CustIDNumber%\%MOCNumber%\log.txt

Gosub, CountNumberofFolders    ;Function that returns variable "Foldercount".  Counts number of folders in CustIDNumber

If Foldercount > 3     ;  Retain _CAD, _Drawing and checkedout.txt files in CustIDNumber folder if there are other MOCs checked in
{	
	
	FileMoveDir, C:\Projects_PCCI\%CustIDNumber%\%MOCNumber%, C:\Projects_PCCI\!Completed\%CustIDNumber%\%MOCNumber%,1
	FileCopyDir, C:\Projects_PCCI\%CustIDNumber%\_drawing, C:\Projects_PCCI\!Completed\%CustIDNumber%\_drawing,1
	FileCopyDir,  C:\Projects_PCCI\%CustIDNumber%\_CAD, C:\Projects_PCCI\!Completed\%CustIDNumber%\_CAD,1
	FileCopy, C:\Projects_PCCI\%CustIDNumber%\MasterChild.txt, C:\Projects_PCCI\!Completed\%CustIDNumber%\MasterChild.txt,1
	
}
Else	; Move all files if no other MOCs are checked in Cust ID
{
	
	FileMoveDir, C:\Projects_PCCI\%CustIDNumber%, C:\Projects_PCCI\!Completed,1
	FileDelete, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\Checkoutflag.txt
	FileDelete, C:\Projects_PCCI\%CustIDNumber%\Checkoutflag.txt
}
MsgBox, Project Revision %RevNumber% has been successfully archived and moved to the Completed folder.                                                                                                                      Code: %foldercount%
return
;-----------------------------------------------------


CheckOutCopyFolders:    ;Procedure to move folders for check out.  

 Gosub, CountNumberofFolders    ;Function that returns variable "Foldercount".  Counts number of folders in CustIDNumber

       If Foldercount > 0 ; this if statement ignores overwriting the _CAD and _drawing folders on the C: Projects_PCCI folder if there are other MOCs already checked in.  
	     {                             
	        FileCopyDir, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\%MOCNumber%\%RevNumber%, C:\Projects_PCCI\%CustIDNumber%\%MOCNumber%\%RevNumber%, 1
            FileCopy, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\%MOCNumber%\log.txt, C:\Projects_PCCI\%CustIDNumber%\%MOCNumber%\log.txt, 1
            
			FileCopy, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\Checkoutflag.txt, C:\Projects_PCCI\%CustIDNumber%\Checkoutflag.txt, 1
		}
       Else ;copy all if less than 1 MOC
        {
	      	FileCopyDir, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\_CAD, C:\Projects_PCCI\%CustIDNumber%\_CAD, 1
            FileCopyDir, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\_drawing, C:\Projects_PCCI\%CustIDNumber%\_drawing, 1  
			            			
            FileCopyDir, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\%MOCNumber%\%RevNumber%, C:\Projects_PCCI\%CustIDNumber%\%MOCNumber%\%RevNumber%, 1
            FileCopy, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\%MOCNumber%\log.txt, C:\Projects_PCCI\%CustIDNumber%\%MOCNumber%\log.txt, 1
            FileCopy, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\Checkoutflag.txt, C:\Projects_PCCI\%CustIDNumber%\Checkoutflag.txt, 1
			FileCopy, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\MasterChild.txt, C:\Projects_PCCI\%CustIDNumber%\MasterChild.txt, 1

	}

return
;-----------------------------------------------------

CheckProjectExistence:
Gui, Submit, NoHide
ProjectExists := 0

If CustIDNumber =
{
   MsgBox, Please enter a Customer ID Number.
   return
}

If MOCNumber =
{
   MsgBox, Please enter an MOC Number.
   return
}


IfNotExist, P:\Site_Planning\Project_Archive_PCCI\%CustIDNumber%\%MOCNumber%
{
   MsgBox, Project does not exist.
   return
}

ProjectExists := 1
return

;------------------------------------------------------
CountNumberOfFolders:    ;This function counts the number of folders and returns that result
Gui, Submit, NoHide

Foldercount := 0
Loop, C:\Projects_PCCI\%CustIDNumber%\*.*,2  ;looks at folders only
{	
	Foldercount ++
	
}


return 

;------------------------------------------------------


CreateMasterChildTxt:  ;Procedure to open Master-Child.xlms file.  Places MasterChild.txt in C:\Projects_PCCI.
Gui, Submit, NoHide


   MsgBox, Creating MasterChild.txt
   ;This code block opens up O:\Symbols\10 - PMCC\Macro\MasterChild.xlsm
   {
    
     MsgBox,4,, Please fill out excel file, and run excel macro to continue.
      IfMsgBox NO
       {
                MsgBox, MasterChild.txt aborted - Not saved in project folders
                return
       }       
      IfMsgBox YES
       {         
                FileDelete, C:\Projects_PCCI\MasterChild.txt	   
				Run O:\Symbols\10 - PMCC\Macro\MasterChild.xlsm
				MsgBox,4,, Fill out excel with MasterChild info. Create MasterChild.txt by pressing "Master" icon on excel file.  
				IfMsgBox NO
				{
					MsgBox, MasterChild.txt aborted - Not saved in project folders
					return
				}
				IfMsgBox YES
				{ 
				    FileRead, MasterChildContent, C:\Projects_PCCI\MasterChild.txt  ;This code block shows content of file via message box
                    MsgBox,, Contents of MasterChild.txt - local copy only , **This is a local copy in C:/Projects_PCCI/ only and not in the project folder yet**`n`n%MasterChildContent%`n`n`nREMINDER TO CLOSE EXCEL FILE WHEN DONE
					return
				}
				return
       }
   }

return


/*  ------------------------------------------------
Section for Deleted functions and icons

Icons
------------
Gui, Add, Button, x150 y625 w125 h40 gViewMasterChildTxt,  View &MasterChild.txt `n(Customer ID Folder)
Gui, Add, Button, x16 y665 w260 h40 gTransferMasterChildtxt,  Transfer &MasterChild Text File to Proj Folders `n(Customer ID Field Required)






Functions
;--------------------------------------------------------------
ViewMasterChildTxt:  ; Procedure to view contents of MasterChild.txt
Gui, Submit, NoHide
    If CustIdNumber =
	{
		MsgBox, Enter a Cust ID number!
		return
	}

	IfNotExist, C:\Projects_PCCI\%CustIDNumber%
	{
		MsgBox, Project does not exist. Please make sure project is checked out and try again.
		return
	}
	
	FileRead, MasterChildContent, C:\Projects_PCCI\%CustIDNumber%\MasterChild.txt  ;This code block shows content of file via message box
    MsgBox,, Contents of MasterChild.txt, %MasterChildContent%`n`n`nREMINDER TO CLOSE EXCEL FILE WHEN DONE
return
;-----------------------------------------------------


TransferMasterChildtxt:    ;Procedure to create MasterChild txt files and place them in the Customer ID level.  CUSTID must be filled out for macro to run
Gui, Submit, NoHide
If CustIdNumber =
{
   MsgBox, Enter a Cust ID number!
   return
}

IfNotExist, C:\Projects_PCCI\%CustIDNumber%
{
   MsgBox, Project does not exist. Please make sure project is checked out and try again.
   return
}

IfExist, C:\Projects_PCCI\%CustIDNumber%
{

   MsgBox, 4,, Creating MasterChild.txt for Master Customer ID: %CustIDNumber%
   IfMsgBox YES
   {
    Ifnotexist  C:\Projects_PCCI\MasterChild.txt
		{
			MsgBox, MasterChild.txt does not exist
			Return
		}
	Ifexist C:\Projects_PCCI\MasterChild.txt
	    {
		    Filecopy, C:\Projects_PCCI\MasterChild.txt, C:\Projects_PCCI\%CustIDNumber%\MasterChild.txt,1
			MsgBox, MasterChild.txt successfully transferred to C:\Projects_PCCI\%CustIDNumber%\MasterChild.txt
			Run, C:\Projects_PCCI\%CustIDNumber%
			Return
		}
   }
 
   IfMsgBox No
   {
  
     MsgBox, MasterChild.txt aborted - Not saved in project folders
     return
   }       

}

return
   
;-----------------------------------------------------
*/

GuiClose:
ExitApp