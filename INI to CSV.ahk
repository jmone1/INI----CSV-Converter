INIToCSVersion = 10 Sep 2022

;============= AutoHotKey Environment Settings ===========================
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance ignore
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#MaxMem 4095
SetTitleMatchMode, 2 ; As long as it's not 3 it should work, if it's 3 you'd need to alter the title used in the WinGet line below.
DetectHiddenWindows, On
WinGet, BottomMost, IDLast, %A_ScriptFullPath%
If (A_ScriptHwnd != BottomMost)
  ExitApp

;============= Set Vars ===========================
FileEncoding, UTF-8-RAW
;============= Initialise and Setup Menu===========================

;============= Main Menu ================================
Gui Add, TreeView, gShowOptions x5 y12 w165 h260
Gui Add, Button,x5 y310 cBlue gButtonClearLog, X
Gui Add, Text,x35 y314, Clear Log of Events
Gui Add, Edit, x5 y335 w595 h95 vLog
Gui Add, Button, x5 y280 w167 h23 Default, &Cancel

;=============  Standalone Menu ================================
TV_Add("About")
TV_Add("INI <---> CSV")

Gui Add, Tab2, vTab x7 y500 w588 h22 -Wrap +Theme, About|INI <---> CSV
Gui Show, w610 h440, INI <---> CSV

; About
  Gui Tab, 1
  Gui Add, GroupBox, x179 y7 w420 h300, About INI <---> CSV
  Gui Add, Text,x200 y50, Version : %INIToCSVersion%
  Gui Add, Text,, This utility has two funcutions`n1) Convert an INI file to CSV for vewing and editing, and `n2) Convert a CSV back to an INI file `n`n
  GUI Add, Text,, This utility will work on files with a similar structure to INI files `n`n 
  Gui Add, Link,, Thanks to mikeyww and BoBofor assistance with the script <a href="https://www.autohotkey.com/boards/viewtopic.php?f=76&t=108210&sid=99fb0c6b9f7368f07375573f669bfa26">Thread on AutoHotKey Forum</a>

; INI <---> CSV
  Gui Tab, 2
  Gui Add, GroupBox, x179 y7 w420 h300, INI <---> CSV
  Gui, Add, Text,x189 y30, Import / Export
  Gui, Add, Button,gButtonWriteCSV, Convert an INI File to CSV 
  Gui, Add, Button,gButtonWriteINI, Create an INI File from CSV 

ShowOptions:
    TV_GetText(OutputVar, A_EventInfo)
    GuiControl ChooseString, SysTabControl321, % OutputVar
	
Return

;============= Process Choice =============
ButtonClearLog:
  Gui, Submit, NoHide
  guiControlGet, UserInput
  MsgLog = 
  GuiControl,,Log, %MsgLog%
return

;============= INI to CSV =============
ButtonWriteCSV:
  cell := []

  MsgLog = Reading INI File - Please Wait`n%MsgLog%
  GuiControl,,Log, %MsgLog%
  FileSelectFile, SelectedFile, 3, , Open a file, (*.ini; *.lib)
  GuiControl,,Select%A_GuiControl%,%SelectedFile%
  If (SelectedFile = "")
  {
    MsgLog = Operation Cancelled`n%MsgLog%
    GuiControl,,Log, %MsgLog%
    return
  }

  IniRead, sectionList, %SelectedFile%
  For each, sectionName in section := StrSplit(sectionList, "`n") 
  SplitPath, SelectedFile,,,,OutputFileName
  OutputFileName = %OutputFileName%_%A_Now%.csv
  FileAppend ,,%OutputFileName%, UTF-8-RAW
  
  IniRead, sectionList, %SelectedFile%
  For each, sectionName in section := StrSplit(sectionList, "`n") 
  {
    IniRead, sectionText, %SelectedFile%, %sectionName%
    For each, line in StrSplit(sectionText, "`n")
    part := StrSplit(line, "="), cell[part.1, sectionName] := part.2
  }

  MsgLog = `nWritting CSV File - Please Wait`n%MsgLog%
  GuiControl,,Log, %MsgLog%

  FileAppend, Section, %OutputFileName%
  For key in cell
    FileAppend, `,%key%, %OutputFileName%
  For each, sectionName in section 
  {
    FileAppend, `n%sectionName%, %OutputFileName%
    For key, arr in cell
      FileAppend, % "," arr[sectionName], %OutputFileName%
      MsgLog = .%MsgLog%
      GuiControl,,Log, %MsgLog%
  }
  MsgLog = Finished`n%MsgLog%
  GuiControl,,Log, %MsgLog%

  MsgBox, 4,, Would you like to open the CSV?
  IfMsgBox Yes
    Run, %OutputFileName%

return

;============= CSV to INI =============
ButtonWriteINI:
  MsgLog = Reading CSV File - Please Wait`n%MsgLog%
  GuiControl,,Log, %MsgLog%
  FileSelectFile, SelectedFile, 3, , Open a file, *.csv
  GuiControl,,Select%A_GuiControl%,%SelectedFile%
  If (SelectedFile = "")
  {
    MsgLog = Operation Cancelled`n%MsgLog%
    GuiControl,,Log, %MsgLog%
    return
  }

  MsgLog = `nWritting INI File - Please Wait`n%MsgLog%
  GuiControl,,Log, %MsgLog%
  SplitPath, SelectedFile,,,,OutputFileName
  OutputFileName = %OutputFileName%_%A_Now%.ini
  FileAppend ,,%OutputFileName%, UTF-8-RAW


  Loop, read, %SelectedFile%
  {
    MsgLog = .%MsgLog%
    GuiControl,,Log, %MsgLog%
    LineNumber := A_Index
    Loop, parse, A_LoopReadLine, CSV
    {
      If LineNumber = 1
	    Key%A_Index% := A_LoopField
      Else
	  {
	    If A_Index = 1
		  Section = %A_LoopField%
		ThisKey := Key%A_Index%
		ThisValue := A_LoopField
		If ThisValue != 
		  If A_Index > 1
		    IniWrite, %ThisValue%, %OutputFileName%, %Section%, %ThisKey%
	  }
    }
	If LineNumber > 1
      FileAppend ,`n,%OutputFileName%, UTF-8-RAW
  }
  MsgLog = Finished`n%MsgLog%
  GuiControl,,Log, %MsgLog%
return

;=================== APP END ACTIONS ==========================
ButtonCancel:
GuiClose:
Script_End:
Gui, Hide
Progress, Off
exitapp
