#NoTrayIcon

#cs -----------------------------------------------------------------------------------------------------------------------------------------------

	Name: myFileSysDB

	Version: 1.1

	Autor: d.schomburg@gmx.net
	und Mario

#ce -----------------------------------------------------------------------------------------------------------------------------------------------


#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <GuiListView.au3>
#include <File.au3>
#include <array.au3>
#include <WinAPI.au3>
#include <EditConstants.au3>
#include <StaticConstants.au3>
#include <APIConstants.au3>
#include <WinAPIEx.au3>
#include <Constants.au3>
#include <Date.au3>

#include <MsgBoxConstants.au3>
#include <GUIConstantsEx.au3>
#include <ListViewConstants.au3>


Global $nCurCol = -1
Global $nSortDir = 1
Global $bSet = 0
Global $nCol = -1

Global $Projects[500][7] ; 0-Jahr  1-Nummer  2-Name  3-Aktiv  4-Abgebrochen  5-Fertig  6-Pfadname
Global $row_handels[500] ; GUI-ID der Tabellenzeile
Global $mnu_handels[500][5] ; GUI-ID des Kontextmenueelements

Global $path
Global $wintitle
Global $Iconfile
Global $IconID
global $nL ; numberletter


Global $modalwin = 0

Global $Form1 = -1
Global $Form2 = -1
Global $Form3 = -1
Global $btn_Neu_OK = -1
Global $btn_Neu_Abbrechen = -1
Global $btn_Auto_Nummer = -1
Global $inp_Neu_Jahr
Global $inp_Neu_Nummer
Global $inp_Neu_Name
Global $chk_Neu_Oeffnen
Global $neuJahr
Global $neuNummer
Global $neuName
Global $neudoopen

Global $reporttable[1][3] = [[0]]

; msgbox(0,"",stringtrimright(@Scriptfullpath,4) & ".ini")


; set rootfolder for projects

; standardsettings for ThHF, they get overwritten if .ini-File exists
$path = "J:\_Projekte"
$wintitle = "Projekte"
$Iconfile = ""
$IconID = -1
$nL = "P"

Local $inifilename = StringTrimRight(@ScriptFullPath, 4) & ".ini" ; get .ini-name from script- or exe-name


If Drive_and_File_Exists($inifilename) Then

	$xpath = IniRead($inifilename, "general", "rootpath", @ScriptDir)

	If PathIsRelative($xpath) Then

		$path = _WinAPI_PathAppend(@ScriptDir, $xpath)
	ElseIf PathIsRoot($xpath) Then

		$path = $xpath
	Else
		$path = @ScriptDir
	EndIf

	$wintitle = IniRead($inifilename, "general", "title", "myFileSysDB")

	$xiconfile = IniRead($inifilename, "general", "iconfile", "")
	If PathIsRelative($xiconfile) Then
		$Iconfile = _WinAPI_PathAppend(@ScriptDir, $xiconfile)
	ElseIf PathIsRoot($xiconfile) Then
		$Iconfile = $xiconfile
	ElseIf StringInStr($xiconfile, "\") = 0 Then
		$Iconfile = @ScriptDir & "\" & $xiconfile
	Else
		$Iconfile = ""
	EndIf
	If Not FileExists($Iconfile) Then
		$Iconfile = ""
	EndIf

	$IconID = IniRead($inifilename, "general", "iconID", -1)

	$nL = IniRead($inifilename, "general", "numberletter", "P")


	IF Not Drive_and_File_Exists($path) Then
		MsgBox(0, "Fehler", "Der Wurzelordner "& $path & " exisitier nicht!", Default)

		$sFileSelectFolder = FileSelectFolder("Bitte wählen Sie einene Wurzelordner!", "")

		If @error Then
			Exit
		Else
			$path=$sFileSelectFolder

			$answer=MsgBox(3,"Frage","Soll der eben gewählte ordner in die ini.-Datei "& $inifilename &" eingetragen werden?")
			if $answer=$IDYES Then
				IniWrite ( $inifilename, "general", "rootpath",$path)
			elseif $answer=$IDCANCEL Then
				Exit
			endif

		EndIf

	endif

else

	IF Not Drive_and_File_Exists($path) Then
		MsgBox(0, "Fehler", "Die .ini-Datei  "& $inifilename &"  fehlt oder der Wurzelordner  "& $path & "  exisitier nicht!", Default)

		$sFileSelectFolder = FileSelectFolder("Bitte wählen Sie einene Wurzelordner!", "")

		If @error Then
			Exit
		Else
			$path=$sFileSelectFolder

			$answer=MsgBox(3,"Frage","Soll der eben gewählte Ordner in eine neue ini.-Datei "& $inifilename &" eingetragen werden?")
			if $answer=$IDYES Then
				IniWrite ( $inifilename, "general", "rootpath", $path )
			elseif $answer=$IDCANCEL Then
				Exit
			endif

		EndIf

	endif

endif


$path = _WinAPI_PathRemoveBackslash($path)


; --------------------- fill tables with folder and shortcutsinfo -------------------------

fillprojectlist($path)

marc_properties($path & "\_aktiv", 3, True) ; Activ
marc_properties($path & "\_abgebrochen", 4, True) ; Abgebrochen
marc_properties($path & "\_fertig", 5, False) ; Fertig

; <-------------GUI--------------------------GUI--------------------------GUI---------------------------GUI------------------------------------------->

Local $listview, $button, $msg

$Form1 = GUICreate($wintitle, 600, 600, -1, -1, BitOR($WS_SIZEBOX, $WS_MINIMIZEBOX))
$WinPos = WinGetPos($Form1)
$hListView = GUICtrlCreateListView("#|Jahr|Nummer|Name|Aktiv|Abgebrochen|Fertig", 0, 0, $WinPos[2] - 16, $WinPos[3] - 58)
GUICtrlRegisterListViewSort(-1, "LVSort") ; Register the function "SortLV" for the sorting callback

GUICtrlSetResizing($hListView, $GUI_DOCKBORDERS)

If $Iconfile <> "" Then
	GUISetIcon($Iconfile, $IconID)
EndIf

$i = 1
While $i <= $Projects[0][0]
	If ($Projects[$i][0] <> "") Then

		$row_handels[$i] = GUICtrlCreateListViewItem($i & "|" & $Projects[$i][0] & "|" & $Projects[$i][1] & "|" _
				 & $Projects[$i][2] & "|" & $Projects[$i][3] & "|" & $Projects[$i][4] & "|" & $Projects[$i][5], $hListView)

		Local $LVcontextmenu = GUICtrlCreateContextMenu($row_handels[$i])
		$mnu_handels[$i][0] = GUICtrlCreateMenuItem("Löschen", $LVcontextmenu)
		$mnu_handels[$i][1] = GUICtrlCreateMenuItem("Umbenennen", $LVcontextmenu)
		$mnu_handels[$i][2] = GUICtrlCreateMenuItem("aktiv", $LVcontextmenu)
		If StringStripWS($Projects[$i][3], 3) <> "" Then GUICtrlSetState(-1, $GUI_CHECKED)
		$mnu_handels[$i][3] = GUICtrlCreateMenuItem("abgebrochen", $LVcontextmenu)
		If StringStripWS($Projects[$i][4], 3) <> "" Then GUICtrlSetState(-1, $GUI_CHECKED)
		$mnu_handels[$i][4] = GUICtrlCreateMenuItem("fertig", $LVcontextmenu)
		If StringStripWS($Projects[$i][5], 3) <> "" Then GUICtrlSetState(-1, $GUI_CHECKED)
	EndIf

	$i += 1
WEnd
$row_handels[0] = $Projects[0][0]

_GUICtrlListView_SetColumnWidth($hListView, 0, 80)
_GUICtrlListView_SetColumnWidth($hListView, 1, 100)
_GUICtrlListView_SetColumnWidth($hListView, 2, 200)
_GUICtrlListView_SetColumnWidth($hListView, 3, 200)


; <----------------------------------------------------------Menu Bar Names---------------------------------------------------------------------------->


$Menue_Main_1 = GUICtrlCreateMenu("Datei")

$Menue_Main_1_1 = GUICtrlCreateMenuItem("Neues Projekt", $Menue_Main_1)
GUICtrlCreateMenuItem("", $Menue_Main_1)
$Menue_Main_1_2 = GUICtrlCreateMenuItem("Projekteliste exportieren", $Menue_Main_1)
$Menue_Main_1_3 = GUICtrlCreateMenuItem("Öffne Wurzelordner", $Menue_Main_1)
$Menue_Main_1_4 = GUICtrlCreateMenuItem("Zeige .ini-Datei", $Menue_Main_1)
GUICtrlCreateMenuItem("", $Menue_Main_1)
$Menue_Main_1_5 = GUICtrlCreateMenuItem("Beenden", $Menue_Main_1)

$Menue_Main_2 = GUICtrlCreateMenu("Testen")

$Menue_Main_2_1 = GUICtrlCreateMenuItem("Überprüfen Namen in _Name", $Menue_Main_2)
$Menue_Main_2_2 = GUICtrlCreateMenuItem("Überprüfen Übereinstimmung des Jahrs", $Menue_Main_2)
$Menue_Main_2_3 = GUICtrlCreateMenuItem("Überprüfen Links in _Name", $Menue_Main_2)
$Menue_Main_2_4 = GUICtrlCreateMenuItem("Überprüfen Links in _Abgeborchen", $Menue_Main_2)
$Menue_Main_2_5 = GUICtrlCreateMenuItem("Überprüfen Links in _Fertig", $Menue_Main_2)
$Menue_Main_2_6 = GUICtrlCreateMenuItem("Überprüfen Links in _Aktiv", $Menue_Main_2)

$Menue_Main_3 = GUICtrlCreateMenu("Werkzeuge")

$Menue_Main_3_1 = GUICtrlCreateMenuItem("Verschieben von Ordnerbäumen", $Menue_Main_3)

$Menue_Main_4 = GUICtrlCreateMenu("Hilfe")

$Menue_Main_4_1 = GUICtrlCreateMenuItem("Über myFileSysDB ...", $Menue_Main_4)

GUISetState(@SW_SHOW)
_GUICtrlListView_SetColumnWidth($hListView, 0, 30)
_GUICtrlListView_SetColumnWidth($hListView, 1, 50)
_GUICtrlListView_SetColumnWidth($hListView, 2, 70)
_GUICtrlListView_SetColumnWidth($hListView, 3, 250)
_GUICtrlListView_SetColumnWidth($hListView, 4, 60)
_GUICtrlListView_SetColumnWidth($hListView, 5, 60)
_GUICtrlListView_SetColumnWidth($hListView, 5, 60)
; <------------------------------------------------------------Menu Bar Content-------------------------------------------------------------------------->

; _WinAPI_FlashWindowEx($hwnd)

While 1
	$msg = GUIGetMsg(1)

	Switch $msg[1]
		Case $Form1

			If $modalwin <> 0 Then
				If $msg[0] <> $GUI_EVENT_MOUSEMOVE Then
					_WinAPI_FlashWindowEx($modalwin, 3, 5, 100)
				EndIf
			Else

				For $i = 1 To $row_handels[0]
					If $Projects[$i][0] <> "" Then
						Switch $msg[0]
							Case $row_handels[$i]
								ShellExecute($Projects[$i][6])

							Case $mnu_handels[$i][0] ; Löschen

								If (test_if_folder_empty($Projects[$i][6]) <> 0) Then
									MsgBox(0, "Fehler", "Der Projektordner ist nicht leer!", Default, $Form1)
								Else
									If MsgBox(4, "Löschen", "Wirklich löschen?", Default, $Form1) = 6 Then
										loeschen($i)
									EndIf
								EndIf

							Case $mnu_handels[$i][1] ; Umbenennen
								$neuerName = InputBox("Umbenennen", "neuer Name?", $Projects[$i][2], " M", Default, Default, Default, Default, Default, $Form1)
								If $neuerName <> "" Then
									umbenennen($i, $neuerName)
								EndIf

							Case $mnu_handels[$i][2] ;aktiv
								If BitAND(GUICtrlRead($mnu_handels[$i][2]), $GUI_CHECKED) = $GUI_CHECKED Then
									GUICtrlSetState($mnu_handels[$i][2], $GUI_UNCHECKED)

									$Projects[$i][3] = " "
								Else
									GUICtrlSetState($mnu_handels[$i][2], $GUI_CHECKED)

									$Projects[$i][3] = "x"
								EndIf

								do_shortcut_mark($path, "_aktiv", 3, $i, "x")

								GUICtrlSetData($row_handels[$i], $i & "|" & $Projects[$i][0] & "|" & $Projects[$i][1] & "|" _
										 & $Projects[$i][2] & "|" & $Projects[$i][3] & "|" & $Projects[$i][4] & "|" & $Projects[$i][5])

							Case $mnu_handels[$i][3] ;abgebrochen
								If BitAND(GUICtrlRead($mnu_handels[$i][3]), $GUI_CHECKED) = $GUI_CHECKED Then
									GUICtrlSetState($mnu_handels[$i][3], $GUI_UNCHECKED)

									$Projects[$i][4] = " "
								Else
									GUICtrlSetState($mnu_handels[$i][3], $GUI_CHECKED)

									$Projects[$i][4] = "x"
								EndIf

								do_shortcut_mark($path, "_abgebrochen", 4, $i, "x")

								GUICtrlSetData($row_handels[$i], $i & "|" & $Projects[$i][0] & "|" & $Projects[$i][1] & "|" _
										 & $Projects[$i][2] & "|" & $Projects[$i][3] & "|" & $Projects[$i][4] & "|" & $Projects[$i][5])
							Case $mnu_handels[$i][4] ;fertig
								If BitAND(GUICtrlRead($mnu_handels[$i][4]), $GUI_CHECKED) = $GUI_CHECKED Then
									GUICtrlSetState($mnu_handels[$i][4], $GUI_UNCHECKED)

									$fertigjahr = $Projects[$i][5]
									$Projects[$i][5] = " "

									do_shortcut_mark($path, "_fertig", 5, $i, $fertigjahr)

									GUICtrlSetData($row_handels[$i], $i & "|" & $Projects[$i][0] & "|" & $Projects[$i][1] & "|" _
											 & $Projects[$i][2] & "|" & $Projects[$i][3] & "|" & $Projects[$i][4] & "|" & $Projects[$i][5])
								Else
									GUICtrlSetState($mnu_handels[$i][4], $GUI_CHECKED)

									$fertigjahr = InputBox("Fertigeintragung", "Wann?", @YEAR, Default, Default, Default, Default, Default, Default, $Form1)
									If $fertigjahr = "" Then
										GUICtrlSetState($mnu_handels[$i][4], $GUI_UNCHECKED)
										$Projects[$i][5] = " "
									Else
										$Projects[$i][5] = $fertigjahr

										do_shortcut_mark($path, "_fertig", 5, $i, $fertigjahr)

										GUICtrlSetData($row_handels[$i], $i & "|" & $Projects[$i][0] & "|" & $Projects[$i][1] & "|" _
												 & $Projects[$i][2] & "|" & $Projects[$i][3] & "|" & $Projects[$i][4] & "|" & $Projects[$i][5])

									EndIf
								EndIf

								do_shortcut_mark($path, "_fertig", 5, $i, $fertigjahr)

						EndSwitch
					EndIf
				Next

				Switch $msg[0]
					Case $GUI_EVENT_CLOSE, $Menue_Main_1_5
						Exit
					Case $Menue_Main_1_1
						neues_Projekt($path)
					Case $Menue_Main_1_2
						saveall($path)
					Case $Menue_Main_1_3
						ShellExecute($path)
					Case $Menue_Main_1_4
						ShellExecute($inifilename)

					Case $Menue_Main_2_1
						checkfolder($path)
					Case $Menue_Main_2_2
						checknameyear($path)
					Case $Menue_Main_2_3
						checkname($path)
					Case $Menue_Main_2_4
						checkab($path)
					Case $Menue_Main_2_5
						checkfertig($path)
					Case $Menue_Main_2_6
						checkaktiv($path)

					Case $Menue_Main_3_1
						foldermove($path)

					Case $Menue_Main_4_1
						MsgBox(0, "Info", "myFileSysDB 1.1" & @CRLF & @CRLF & "programmiert von" & @CRLF & @CRLF & "Dirk Schomburg" & @CRLF & "d.schomburg@gmx.net" & @CRLF & @CRLF & "und" & @CRLF & "Mario", Default, $Form1)

					Case $hListView
						$bSet = 0
						$nCurCol = $nCol
						GUICtrlSendMsg($hListView, $LVM_SETSELECTEDCOLUMN, GUICtrlGetState($hListView), 0)
						DllCall("user32.dll", "int", "InvalidateRect", "hwnd", GUICtrlGetHandle($hListView), "int", 0, "int", 1)
				EndSwitch
			EndIf

		Case $Form3
			Switch $msg[0]
				Case $btn_Neu_Abbrechen
					GUIDelete($Form3)
					$modalwin = 0
				Case $btn_Auto_Nummer

					$neuJahr = GUICtrlRead($inp_Neu_Jahr)

					If StringStripWS($neuJahr, 3) = "" Then
						$neuJahr = @YEAR
						GUICtrlSetData($inp_Neu_Jahr, $neuJahr)
					EndIf

					$maxnumber = 0

					For $i = 1 To $Projects[0][0]
						If $Projects[$i][0] = $neuJahr Then
							If StringRight($Projects[$i][1], 4) > $maxnumber Then
								$maxnumber = StringRight($Projects[$i][1], 4)
							EndIf
						EndIf
					Next

					GUICtrlSetData($inp_Neu_Nummer, StringRight($neuJahr, 2) & $nL & StringRight(10001 + $maxnumber, 4)) ; normaly: $nl = "P"

				Case $btn_Neu_OK

					$neuJahr = GUICtrlRead($inp_Neu_Jahr)
					$neuNummer = GUICtrlRead($inp_Neu_Nummer)
					$neuName = GUICtrlRead($inp_Neu_Name)
					$neudoopen = GUICtrlRead($chk_Neu_Oeffnen)

					If (StringStripWS($neuJahr, 3) = "") Or (StringStripWS($neuNummer, 3) = "") Or (StringStripWS($neuName, 3) = "") Then
						MsgBox(0, "Fehler", "Ohne Eingabe geht es nicht!", Default, $Form3)
					ElseIf Number(StringRight(StringStripWS($neuJahr, 3), 2)) <> Number(StringLeft(StringStripWS($neuNummer, 3), 2)) Then
						MsgBox(0, "Fehler", "Projektnummer passt nicht zum Jahr!", Default, $Form3)
					Else
						$error = 0
						For $i = 1 To $Projects[0][0]
							If $Projects[$i][1] = $neuNummer Then
								$error = 1
								ExitLoop
							EndIf
						Next
						If $error = 1 Then
							MsgBox(0, "Fehler", "Projektnummer ist schon vorhanden!", Default, $Form3)
						Else
							GUIDelete($Form3)
							$modalwin = 0
							create_new_project($path)
						EndIf
					EndIf

				Case $GUI_EVENT_CLOSE
					GUIDelete($Form3)
					$modalwin = 0
			EndSwitch

	EndSwitch

WEnd

; Our sorting callback funtion
Func LVSort($hWnd, $nItem1, $nItem2, $nColumn)
	Local $val1, $val2, $nResult

	; Switch the sorting direction
	If $nColumn = $nCurCol Then
		If Not $bSet Then
			$nSortDir = $nSortDir * - 1
			$bSet = 1
		EndIf
	Else
		$nSortDir = 1
	EndIf
	$nCol = $nColumn

	$val1 = GetSubItemText($hWnd, $nItem1, $nColumn)
	$val2 = GetSubItemText($hWnd, $nItem2, $nColumn)


	If $nColumn = 0 Then
		$val1 = Number($val1)
		$val2 = Number($val2)
	EndIf
	if $nColumn = 1 Then
		if $val1 = "20xx" then $val1="1999"
		if $val1 = "200x" then $val1="1999"

		if $val2 = "20xx" then $val2="1999"
		if $val2 = "200x" then $val2="1999"
	endif

	$nResult = 0 ; No change of item1 and item2 positions

	If $val1 < $val2 Then
		$nResult = -1 ; Put item2 before item1
	ElseIf $val1 > $val2 Then
		$nResult = 1 ; Put item2 behind item1
	EndIf

	$nResult = $nResult * $nSortDir

	Return $nResult
EndFunc   ;==>LVSort


; Retrieve the text of a listview item in a specified column
Func GetSubItemText($nCtrlID, $nItemID, $nColumn)
	Local $stLvfi = DllStructCreate("uint;ptr;int;int[2];int")
	Local $nIndex, $stBuffer, $stLvi, $sItemText

	DllStructSetData($stLvfi, 1, $LVFI_PARAM)
	DllStructSetData($stLvfi, 3, $nItemID)

	$stBuffer = DllStructCreate("char[260]")

	$nIndex = GUICtrlSendMsg($nCtrlID, $LVM_FINDITEM, -1, DllStructGetPtr($stLvfi));

	$stLvi = DllStructCreate("uint;int;int;uint;uint;ptr;int;int;int;int")

	DllStructSetData($stLvi, 1, $LVIF_TEXT)
	DllStructSetData($stLvi, 2, $nIndex)
	DllStructSetData($stLvi, 3, $nColumn)
	DllStructSetData($stLvi, 6, DllStructGetPtr($stBuffer))
	DllStructSetData($stLvi, 7, 260)

	GUICtrlSendMsg($nCtrlID, $LVM_GETITEMA, 0, DllStructGetPtr($stLvi));

	$sItemText = DllStructGetData($stBuffer, 1)

	$stLvi = 0
	$stLvfi = 0
	$stBuffer = 0

	Return $sItemText
EndFunc   ;==>GetSubItemText

; --- End of main ------------------------

Func loeschen($i)

	; change the name from the folder

	$oldPathName = $path & "\" & $Projects[$i][0] & "\" & $Projects[$i][1] & " " & $Projects[$i][2]

	;MsgBox(0, "x", $oldPathName)
	DirRemove($oldPathName)

	; change shortcut in the _Namen - Folder

	$oldshortcutname = $path & "\_Namen\" & $Projects[$i][2] & " (" & $Projects[$i][1] & ")" & ".lnk"
	FileDelete($oldshortcutname)

	; change shortcut in the _fertig - Folder

	$oldshortcutname = $Projects[$i][1] & " " & $Projects[$i][2] & ".lnk"
	If FileExists($path & "\_fertig\" & $oldshortcutname) Then
		FileDelete($path & "\_fertig\" & $oldshortcutname)
	EndIf

	If test_if_yearsubfolder($path & "\_fertig") Then
		$search = FileFindFirstFile($path & "\_fertig\*")
		If $search <> -1 Then
			While 1
				$file = FileFindNextFile($search)
				If @error Then ExitLoop
				If test_if_year_folder($file) Then
					If FileExists($path & "\_fertig\" & $file & $oldshortcutname) Then
						FileDelete($path & "\_fertig\" & $file & $oldshortcutname)
					EndIf
				EndIf
			WEnd
			FileClose($search)
		EndIf
	EndIf

	; change shortcut in the _aktiv - Folder

	$oldshortcutname = $Projects[$i][1] & " " & $Projects[$i][2] & ".lnk"
	If FileExists($path & "\_aktiv\" & $oldshortcutname) Then
		FileDelete($path & "\_aktiv\" & $oldshortcutname)
	EndIf
	;	msgbox(0,"a",$path & "\_aktiv\" & $Projects[$i][0] & "\" & $oldshortcutname)
	If FileExists($path & "\_aktiv\" & $Projects[$i][0]) Then
		;	msgbox(0,"b",$path & "\_aktiv\" & $Projects[$i][0] & "\" & $oldshortcutname)
		If FileExists($path & "\_aktiv\" & $Projects[$i][0] & "\" & $oldshortcutname) Then
			;		msgbox(0,"c",$path & "\_aktiv\" & $Projects[$i][0] & "\" & $oldshortcutname)
			FileDelete($path & "\_aktiv\" & $Projects[$i][0] & "\" & $oldshortcutname)
		EndIf
	EndIf

	; change Shortcut in the _abgebrochen - Folder

	$oldshortcutname = $Projects[$i][1] & " " & $Projects[$i][2] & ".lnk"

	If FileExists($path & "\_abgebrochen\" & $oldshortcutname) Then
		FileDelete($path & "\_abgebrochen\" & $oldshortcutname)
	EndIf
	;msgbox(0,"a", $path & "\_abgebrochen\" & $Projects[$i][0] & "\" & $oldshortcutname)

	If FileExists($path & "\_abgebrochen\" & $Projects[$i][0]) Then
		;	msgbox(0,"b", $path & "\_abgebrochen\" & $Projects[$i][0] & "\" & $oldshortcutname)
		If FileExists($path & "\_abgebrochen\" & $Projects[$i][0] & "\" & $oldshortcutname) Then
			;		msgbox(0,"c", $path & "\_abgebrochen\" & $Projects[$i][0] & "\" & $oldshortcutname)
			FileDelete($path & "\_abgebrochen\" & $Projects[$i][0] & "\" & $oldshortcutname)
		EndIf
	EndIf

	; change the name in the area in the scrip

	For $j = 0 To 6
		$Projects[$i][$j] = ""
	Next

	; change the graphical display

	GUICtrlDelete($row_handels[$i])

	; change the name in the area in the scrip



EndFunc   ;==>loeschen

Func umbenennen($i, $NewName)

	; change the name from the folder
	if stringinstr(StringReplace($Projects[$i][6],$path & "\" ,""),"\") then ; if yearfolder
		$NewPathName = $path & "\" & $Projects[$i][0] & "\" & $Projects[$i][1] & " " & $NewName
		$oldPathName = $path & "\" & $Projects[$i][0] & "\" & $Projects[$i][1] & " " & $Projects[$i][2]
	Else
		$NewPathName = $path & "\" & $Projects[$i][1] & " " & $NewName
		$oldPathName = $path & "\" & $Projects[$i][1] & " " & $Projects[$i][2]
	endif


	;MsgBox(0, $oldPathName, $NewPathName)
	DirMove($oldPathName, $NewPathName)

	; change shortcut in the _Namen - Folder

	$oldshortcutname = $path & "\_Namen\" & $Projects[$i][2] & " (" & $Projects[$i][1] & ")" & ".lnk"
	$newshortcutname = $path & "\_Namen\" & $NewName & " (" & $Projects[$i][1] & ")" & ".lnk"
	FileDelete($oldshortcutname)
	FileCreateShortcut($NewPathName, $newshortcutname)

	; change shortcut in the _fertig - Folder

	$oldshortcutname = $Projects[$i][1] & " " & $Projects[$i][2] & ".lnk"
	$newshortcutname = $Projects[$i][1] & " " & $NewName & ".lnk"
	;	msgbox(0,"a",$path & "\_fertig\" & $oldshortcutname)
	If FileExists($path & "\_fertig\" & $oldshortcutname) Then
		FileDelete($path & "\_fertig\" & $oldshortcutname)
		FileCreateShortcut($NewPathName, $path & "\_fertig\" & $newshortcutname)
	EndIf

	If test_if_yearsubfolder($path & "\_fertig") Then
		$search = FileFindFirstFile($path & "\_fertig\*")
		If $search <> -1 Then
			While 1
				$file = FileFindNextFile($search)
				If @error Then ExitLoop
				;		msgbox(0,"c",$path & "\_fertig\" & $file & "\" & $oldshortcutname)
				If test_if_year_folder($file) Then
					If FileExists($path & "\_fertig\" & $file & "\" & $oldshortcutname) Then
						FileDelete($path & "\_fertig\" & $file & "\" & $oldshortcutname)
						FileCreateShortcut($NewPathName, $path & "\_fertig\" & $file & "\" & $newshortcutname)
					EndIf
				EndIf
			WEnd
			FileClose($search)
		EndIf
	EndIf

	; change shortcut in the _aktiv - Folder

	$oldshortcutname = $Projects[$i][1] & " " & $Projects[$i][2] & ".lnk"
	$newshortcutname = $Projects[$i][1] & " " & $NewName & ".lnk"
	If FileExists($path & "\_aktiv\" & $oldshortcutname) Then
		FileDelete($path & "\_aktiv\" & $oldshortcutname)
		FileCreateShortcut($NewPathName, $path & "\_aktiv\" & $newshortcutname)
	EndIf
	;	msgbox(0,"a",$path & "\_aktiv\" & $Projects[$i][0] & "\" & $oldshortcutname)
	If FileExists($path & "\_aktiv\" & $Projects[$i][0]) Then
		;	msgbox(0,"b",$path & "\_aktiv\" & $Projects[$i][0] & "\" & $oldshortcutname)
		If FileExists($path & "\_aktiv\" & $Projects[$i][0] & "\" & $oldshortcutname) Then
			;		msgbox(0,"c",$path & "\_aktiv\" & $Projects[$i][0] & "\" & $oldshortcutname)
			FileDelete($path & "\_aktiv\" & $Projects[$i][0] & "\" & $oldshortcutname)
			FileCreateShortcut($NewPathName, $path & "\_aktiv\" & $Projects[$i][0] & "\" & $newshortcutname)
		EndIf
	EndIf

	; change Shortcut in the _abgebrochen - Folder

	$oldshortcutname = $Projects[$i][1] & " " & $Projects[$i][2] & ".lnk"
	$newshortcutname = $Projects[$i][1] & " " & $NewName & ".lnk"

	If FileExists($path & "\_abgebrochen\" & $oldshortcutname) Then
		FileDelete($path & "\_abgebrochen\" & $oldshortcutname)
		FileCreateShortcut($NewPathName, $path & "\_abgebrochen\" & $newshortcutname)
	EndIf
	;msgbox(0,"a", $path & "\_abgebrochen\" & $Projects[$i][0] & "\" & $oldshortcutname)

	If FileExists($path & "\_abgebrochen\" & $Projects[$i][0]) Then
		;	msgbox(0,"b", $path & "\_abgebrochen\" & $Projects[$i][0] & "\" & $oldshortcutname)
		If FileExists($path & "\_abgebrochen\" & $Projects[$i][0] & "\" & $oldshortcutname) Then
			;		msgbox(0,"c", $path & "\_abgebrochen\" & $Projects[$i][0] & "\" & $oldshortcutname)
			FileDelete($path & "\_abgebrochen\" & $Projects[$i][0] & "\" & $oldshortcutname)
			FileCreateShortcut($NewPathName, $path & "\_aktiv\" & $Projects[$i][0] & "\" & $newshortcutname)
		EndIf
	EndIf

	; change the name in the area in the scrip

	$Projects[$i][2] = $NewName
	$Projects[$i][6] = $NewPathName

	; change the graphical display

	GUICtrlSetData($row_handels[$i], $i & "|" & $Projects[$i][0] & "|" & $Projects[$i][1] & "|" _
			 & $Projects[$i][2] & "|" & $Projects[$i][3] & "|" & $Projects[$i][4] & "|" & $Projects[$i][5])


EndFunc   ;==>umbenennen


; fill tables with folder and shortcutsinfo

Func fillprojectlist($path)

	Local $file
	Local $search = FileFindFirstFile($path & "\*")
	Local $search2
	Local $i = 1

	; Check if the search was successful

	If $search = -1 Then
		Exit
	EndIf

	While 1

		$file = FileFindNextFile($search)

		If @error Then ExitLoop

		If test_if_year_folder($file) Then

			$search2 = FileFindFirstFile($path & "\" & $file & "\*")

			If $search2 <> -1 Then

				While 1
					$file2 = FileFindNextFile($search2)
					If @error Then ExitLoop

					$Projects[$i][0] = $file
					$Projects[$i][1] = StringStripWS(StringLeft($file2, 7), 3)
					$Projects[$i][2] = StringStripWS(StringMid($file2, 8), 3)
					$Projects[$i][6] = $path & "\" & $file & "\" & $file2
					$Projects[0][0] = $i
					$i += 1

				WEnd

			EndIf

			FileClose($search2)

		ElseIf test_if_proj_folder($file) Then

			$Projects[$i][0] = "x"
			$Projects[$i][1] = StringStripWS(StringLeft($file, 7), 3)
			$Projects[$i][2] = StringStripWS(StringMid($file, 8), 3)
			$Projects[$i][6] = $path & "\" & $file
			$Projects[0][0] = $i
			$i += 1

		EndIf


	WEnd

	FileClose($search)

EndFunc   ;==>fillprojectlist

; <-----------------------------------------------------------Columns Active,Abgebrochen,Fertig---------------------------------------------------------->

Func marc_properties($path, $Proj_col, $nur_x)

	; $Proj_col = 3 -> Active
	; $Proj_col = 4 -> Abgebrochen
	; $Proj_col = 5 -> Fertig

	Local $file
	Local $search = FileFindFirstFile($path & "\*")
	Local $search2
	Local $g = 1


	If $search = -1 Then

		Return False

	Else

		While 1

			$file = FileFindNextFile($search)

			If @error Then ExitLoop

			If test_if_year_folder($file) Then

				$search2 = FileFindFirstFile($path & "\" & $file & "\*")

				If $search2 <> -1 Then

					While 1
						$file2 = FileFindNextFile($search2)
						If @error Then ExitLoop

						$g = 1
						While $g <= $Projects[0][0]
							If $Projects[$g][1] = StringLeft($file2, 7) Then
								If $nur_x = True Then
									$Projects[$g][$Proj_col] = "x"
								Else
									$Projects[$g][$Proj_col] = $file
								EndIf
							EndIf
							$g += 1
						WEnd


					WEnd
				EndIf

				FileClose($search2)

			ElseIf test_if_proj_folder($file) Then

				$g = 1
				While $g <= $Projects[0][0]
					If $Projects[$g][1] = StringLeft($file, 7) Then
						$Projects[$g][$Proj_col] = "x"
					EndIf
					$g += 1
				WEnd

			EndIf

		WEnd

		FileClose($search)

		Return True

	EndIf

EndFunc   ;==>marc_properties




Func do_shortcut_mark($path, $markfolder, $Proj_col, $pn, $subfolder)

	If StringStripWS($Projects[$pn][$Proj_col], 3) = "" Then ; löschen
		If $subfolder = "x" Then
			If test_if_yearsubfolder($path & "\" & $markfolder) Then
				;msgbox(0,"",$path & "\" & $markfolder & "\" & $Projects[$pn][0] & "\" & $Projects[$pn][1] & " " & $Projects[$pn][2] & ".lnk")
				FileDelete($path & "\" & $markfolder & "\" & $Projects[$pn][0] & "\" & $Projects[$pn][1] & " " & $Projects[$pn][2] & ".lnk")
			Else
				;msgbox(0,"",$path & "\" & $markfolder & "\" & $Projects[$pn][1] & " " & $Projects[$pn][2] & ".lnk")
				FileDelete($path & "\" & $markfolder & "\" & $Projects[$pn][1] & " " & $Projects[$pn][2] & ".lnk")
			EndIf
		Else
			FileDelete($path & "\" & $markfolder & "\" & $subfolder & "\" & $Projects[$pn][1] & " " & $Projects[$pn][2] & ".lnk")
		EndIf
	Else ; setzen
		If $subfolder = "x" Then
			If test_if_yearsubfolder($path & "\" & $markfolder) Then
				;msgbox(0,"",$path & "\" & $markfolder & "\" & $Projects[$pn][0] & "\" & $Projects[$pn][1] & " " & $Projects[$pn][2] & ".lnk")
				FileCreateShortcut($Projects[$pn][6], $path & "\" & $markfolder & "\" & $Projects[$pn][0] & "\" & $Projects[$pn][1] & " " & $Projects[$pn][2] & ".lnk")
			Else
				;msgbox(0,"",$path & "\" & $markfolder & "\" & $Projects[$pn][1] & " " & $Projects[$pn][2] & ".lnk")
				FileCreateShortcut($Projects[$pn][6], $path & "\" & $markfolder & "\" & $Projects[$pn][1] & " " & $Projects[$pn][2] & ".lnk")
			EndIf
		Else
			If Not FileExists($path & "\" & $markfolder & "\" & $subfolder) Then
				DirCreate($path & "\" & $markfolder & "\" & $subfolder)
			EndIf

			FileCreateShortcut($Projects[$pn][6], $path & "\" & $markfolder & "\" & $subfolder & "\" & $Projects[$pn][1] & " " & $Projects[$pn][2] & ".lnk")
		EndIf
	EndIf

EndFunc   ;==>do_shortcut_mark


; <-----------------------------------------------------------Checks all subfolders in the _Namen folder to find errors---------------------------------->


Func checkfolder($path)

	Local $szDrive, $szDir, $szFName, $szExt

	Local $file
	Local $search = FileFindFirstFile($path & "\*")
	Local $search2
	Local $reporttable[1] = [0]
	local $error = 0

	If $search = -1 Then ; test if first search has errors
		$error = -1
	Else

		While 1 ; loop on first level
			$file = FileFindNextFile($search)

			If @error Then ExitLoop ; no more files

			If @extended Then ; if Folder

				If test_if_year_folder($file) Then
					;msgbox(0,"Year",$file)
					$search2 = FileFindFirstFile($path & "\" & $file & "\*")
					If $search2 = -1 Then ; test if secound loop has errors
						if @error <> 1 then ; @error=1 means that there are no files or folder in the folder, that is no error
							$error = -1
							exitloop
						endif

					Else
						While 1 ; loop on secound level
							$file2 = FileFindNextFile($search2)
							If @error=1 Then ExitLoop ; no more files
							;msgbox(0,$file,$file2)
							If test_if_proj_folder($file2) Then ; if Project-Folder
								$Linkname = StringMid($file2, 9) & " (" & StringLeft($file2, 7) & ")"
								If Not FileExists($path & "\_Namen\" & $Linkname & ".lnk") Then ; if Name-Link ist missing
									;msgbox(0,"x",$path & "\_Namen\" & $Linkname)
									extendarray1D($reporttable)
									$reporttable[$reporttable[0] + 1] = $Linkname
									$reporttable[0] = $reporttable[0] + 1
								EndIf
							EndIf
						WEnd
						FileClose($search2)
					EndIf
				ElseIf test_if_proj_folder($file) Then ; if Project-Folder on level 1
					$Linkname = StringMid($file, 9) & " (" & StringLeft($file, 7) & ")"
					If Not FileExists($path & "\_Namen\" & $Linkname & ".lnk") Then ; if Name-Link ist missing
						extendarray1D($reporttable)
						$reporttable[$reporttable[0] + 1] = $Linkname
						$reporttable[0] = $reporttable[0] + 1
					EndIf
				EndIf
			EndIf

		WEnd
		FileClose($search)
	EndIf

	If $error <> 0 Then
		MsgBox(0, "Fehler", "Ein Fehler passierte bei der Analyse!", Default, $Form1)
	ElseIf $reporttable[0] = 0 Then
		MsgBox(0, "Ergebnis", "Alle Elemente sind Ok!", Default, $Form1)
	Else
		$Liste = "Namen" & @CRLF & "=====" & @CRLF
		For $i = 1 To $reporttable[0]
			$Liste = $Liste & $reporttable[$i] & @CRLF
		next
		MsgBox(0, "Elemente ohne Nameneintrag:", $Liste, Default, $Form1)
	endif
EndFunc   ;==>checkfolder

Func checknameyear($path)

	Local $szDrive, $szDir, $szFName, $szExt

	Local $file
	Local $search = FileFindFirstFile($path & "\*")
	Local $search2
	Local $reporttable[1] = [0]
	local $error = 0

	If $search = -1 Then ; test if first search has errors
		$error = -1
	Else

		While 1 ; loop on first level
			$file = FileFindNextFile($search)

			If @error Then ExitLoop ; no more files

			If @extended Then ; if Folder

				If test_if_year_folder($file) Then
					;msgbox(0,"Year",$file)
					$search2 = FileFindFirstFile($path & "\" & $file & "\*")
					If $search2 = -1 Then ; test if secound loop has errors
						if @error <> 1 then ; @error=1 means that there are no files or folder in the folder, that is no error
							$error = -1
							exitloop
						endif

					Else
						While 1 ; loop on secound level
							$file2 = FileFindNextFile($search2)
							If @error=1 Then ExitLoop ; no more files
							;msgbox(0,$file,$file2)
							If test_if_proj_folder($file2) Then ; if Project-Folder
								$Yearname = StringLeft($file2, 2)
								If StringLeft($file2, 2) <> stringright($file,2) Then ; if Year-Folder did not match Year in Number
									if $file<>"20xx" Then
										extendarray1D($reporttable)
										$reporttable[$reporttable[0] + 1] = $file & ", " & $file2
										$reporttable[0] = $reporttable[0] + 1
									endif
								EndIf
							EndIf
						WEnd
						FileClose($search2)
					EndIf
				endif
			EndIf

		WEnd
		FileClose($search)
	EndIf

	If $error <> 0 Then
		MsgBox(0, "Fehler", "Ein Fehler passierte bei der Analyse!", Default, $Form1)
	ElseIf $reporttable[0] = 0 Then
		MsgBox(0, "Ergebnis", "Alle Elemente sind Ok!", Default, $Form1)
	Else
		$Liste = "Jahr, Name" & @CRLF & "==========" & @CRLF
		For $i = 1 To $reporttable[0]
			$Liste = $Liste & $reporttable[$i] & @CRLF
		next
		MsgBox(0, "Elemente ohne Übereinstimmung:", $Liste, Default, $Form1)
	endif
EndFunc   ;==>checkfolder

Func checkname($path)
	Local $error
	Local $Liste

	$reporttable[0][0] = 0
	$error = ckeckshortcutlinks($path & "\_Namen", $reporttable)

	If $error <> 0 Then
		MsgBox(0, "Fehler", "Ein Fehler passierte bei der Analyse!", Default, $Form1)
	ElseIf $reporttable[0][0] = 0 Then
		MsgBox(0, "Ergebnis", "Alle Links sind Ok!", Default, $Form1)
	Else
		$Liste = "Art,  Linkdatei,  Ziel" & @CRLF & "======================" & @CRLF
		For $i = 1 To $reporttable[0][0]
			Switch $reporttable[$i][1]
				Case 1
					$Liste = $Liste & "Target,  "
				Case 2
					$Liste = $Liste & "WorkDir,  "
				Case 3
					$Liste = $Liste & "Icon,  "
			EndSwitch
			$Liste = $Liste & $reporttable[$i][0] & ",  "
			$Liste = $Liste & $reporttable[$i][2] & @CRLF
		Next
		MsgBox(0, "Fehlerhafte Shortcuts:", $Liste, Default, $Form1)

	EndIf

EndFunc   ;==>checkname

Func checkab($path)
	Local $error
	Local $Liste

	$reporttable[0][0] = 0
	$error = ckeckshortcutlinks($path & "\_abgebrochen", $reporttable)

	If $error <> 0 Then
		MsgBox(0, "Fehler", "Ein Fehler passierte bei der Analyse!", Default, $Form1)
	ElseIf $reporttable[0][0] = 0 Then
		MsgBox(0, "Ergebnis", "Alle Links sind Ok!", Default, $Form1)
	Else
		$Liste = "Art,  Linkdatei,  Ziel" & @CRLF & "======================" & @CRLF
		For $i = 1 To $reporttable[0][0]
			Switch $reporttable[$i][1]
				Case 1
					$Liste = $Liste & "Target,  "
				Case 2
					$Liste = $Liste & "WorkDir,  "
				Case 3
					$Liste = $Liste & "Icon,  "
			EndSwitch
			$Liste = $Liste & $reporttable[$i][0] & ",  "
			$Liste = $Liste & $reporttable[$i][2] & @CRLF
		Next
		MsgBox(0, "Fehlerhafte Shortcuts:", $Liste, Default, $Form1)

	EndIf

EndFunc   ;==>checkab

Func checkfertig($path)
	Local $error
	Local $Liste

	$reporttable[0][0] = 0
	$error = ckeckshortcutlinks($path & "\_fertig", $reporttable)

	If $error <> 0 Then
		MsgBox(0, "Fehler", "Ein Fehler passierte bei der Analyse!", Default, $Form1)
	ElseIf $reporttable[0][0] = 0 Then
		MsgBox(0, "Ergebnis", "Alle Links sind Ok!", Default, $Form1)
	Else
		$Liste = "Art,  Linkdatei,  Ziel" & @CRLF & "======================" & @CRLF
		For $i = 1 To $reporttable[0][0]
			Switch $reporttable[$i][1]
				Case 1
					$Liste = $Liste & "Target,  "
				Case 2
					$Liste = $Liste & "WorkDir,  "
				Case 3
					$Liste = $Liste & "Icon,  "
			EndSwitch
			$Liste = $Liste & $reporttable[$i][0] & ",  "
			$Liste = $Liste & $reporttable[$i][2] & @CRLF
		Next
		MsgBox(0, "Fehlerhafte Shortcuts:", $Liste, Default, $Form1)

	EndIf

EndFunc   ;==>checkfertig

Func checkaktiv($path)
	Local $error
	Local $Liste

	$reporttable[0][0] = 0
	$error = ckeckshortcutlinks($path & "\_aktiv", $reporttable)

	If $error <> 0 Then
		MsgBox(0, "Fehler", "Ein Fehler passierte bei der Analyse!", Default, $Form1)
	ElseIf $reporttable[0][0] = 0 Then
		MsgBox(0, "Ergebnis", "Alle Links sind Ok!", Default, $Form1)
	Else
		$Liste = "Art,  Linkdatei,  Ziel" & @CRLF & "======================" & @CRLF
		For $i = 1 To $reporttable[0][0]
			Switch $reporttable[$i][1]
				Case 1
					$Liste = $Liste & "Target,  "
				Case 2
					$Liste = $Liste & "WorkDir,  "
				Case 3
					$Liste = $Liste & "Icon,  "
			EndSwitch
			$Liste = $Liste & $reporttable[$i][0] & ",  "
			$Liste = $Liste & $reporttable[$i][2] & @CRLF
		Next
		MsgBox(0, "Fehlerhafte Shortcuts:", $Liste, Default, $Form1)

	EndIf

EndFunc   ;==>checkaktiv




; <-----------------------------------------------------------Gives the user the possibility to move folders and subfolders------------------------------>

Func foldermove($path)

	If 2 = MsgBox(1, "Foldermove", "Mit dieser Funktion werden auch die Ziele der Shortcuts bearbeitet, wenn ein Ordnerbaum verschoben wird." & @CRLF & _
			"Die Funktion benötigt drei Ordner. " & @CRLF & _
			"1. den Wurzelordner, dessen Inhalt verschoben werden soll," & @CRLF & _
			"2. das Ziel, wo er hingeschoben werden soll und" & @CRLF & _
			"3. einen Wurzelordner von dem Bereich, in dem alle Shortcuts angepasst werden sollen.") Then
		Return
	endif

	Local $folder_to_move = FileSelectFolder("Ordner, dessen Inhalt der verschoben werden soll", "")
	If $folder_to_move = "" Then Return
	;MsgBox(0,"Ordner, dessen Inhalt verschoben werden soll",$folder_to_move)

	Local $target = FileSelectFolder("Zielordner in den der Inhalt hin soll.", "", 1)
	If $target = "" Then Return
	;MsgBox(0,"Zielordner in den der Inhalt hin soll.",$target)


	; Local $linkadjusting = FileSelectFolder("Bereich, in dem die Links angepasst werden.", "")
	; If $linkadjusting = "" Then Return
	; MsgBox(0,"Bereich, in dem die Links angepasst werden.",$linkadjusting)

	; return 0

	; start execution

	DirCopy($folder_to_move, $target, 1)

	fm_checkshortcuts1($target, $folder_to_move, $target)

	; fm_checkshortcuts2($linkadjusting, $folder_to_move, $target)

	; DirRemove($folder_to_move, 1)

EndFunc   ;==>foldermove


Func fm_checkshortcuts1($target, $folder_to_move_root, $target_root)


	Local $file
	Local $search = FileFindFirstFile($target & "\*")

	; Check if the search was successful
	If $search = -1 Then
		Exit
	EndIf

	MsgBox(0,"1","2")

	While 1
		$file = FileFindNextFile($search)
		MsgBox(0,$file,"w")

		If @error Then ExitLoop
		If $file="" Then ExitLoop

		$Type = FileGetAttrib($target & "\" & $file)
        MsgBox(0,$target & "\" & $file,$Type)


		If 0 <> StringInStr($Type, "D") Then

			fm_checkshortcuts1($target & "\" & $file, $folder_to_move_root, $target_root)

		Else

			If StringRight($file, 4) = ".lnk" Then

				MsgBox(1, "File:", $target & "\" & $file & " Type " & $Type)

				$aDetails = FileGetShortcut($target & "\" & $file)
				If Not @error Then
					MsgBox(0,$folder_to_move_root,$aDetails[0])

					If StringLeft($aDetails[0], StringLen($folder_to_move_root)) = $folder_to_move_root Then

						MsgBox(0, $target & "\" & $file, "Path: " & $aDetails[0])
						MsgBox(0, "needs to change !", "")

						FileDelete($target & "\" & $file)

						$workdir = $aDetails[1]
						If StringLeft($aDetails[1], StringLen($folder_to_move_root)) = $folder_to_move_root Then
							$workdir = $target_root & StringMid($aDetails[1], StringLen($folder_to_move_root) + 1)
						EndIf

						$icon = $aDetails[4]
						If StringLeft($aDetails[4], StringLen($folder_to_move_root)) = $folder_to_move_root Then
							$icon = $target_root & StringMid($aDetails[1], StringLen($folder_to_move_root) + 1)
						EndIf

						MsgBox(0, "old target", $aDetails[0])
						MsgBox(0, "target", $target_root & StringMid($aDetails[0], StringLen($folder_to_move_root) + 1))
						MsgBox(0, "linkname", $target & "\" & $file)


						FileCreateShortcut($target_root & StringMid($aDetails[0], StringLen($folder_to_move_root) + 1), _
								$target & "\" & $file, _
								$workdir, $aDetails[2], $aDetails[3], _
								$icon, $aDetails[5], $aDetails[6])

					EndIf
				EndIf
			EndIf

		EndIf

	WEnd

	FileClose($search)

EndFunc   ;==>fm_checkshortcuts1

Func fm_checkshortcuts2($linkadjusting, $folder_to_move_root, $target_root)

	Local $file
	Local $search = FileFindFirstFile($linkadjusting & "\*")

	; Check if the search was successful
	If $search = -1 Then
		; MsgBox(0, "Error", "No files/directories matched the search pattern")
		Exit
	EndIf

	While 1
		$file = FileFindNextFile($search)

		If @error Then ExitLoop

		$Type = FileGetAttrib($linkadjusting & "\" & $file)

		; MsgBox(1, "File:", $file & " Type " & $Type)

		If 0 <> StringInStr($Type, "D") Then

			fm_checkshortcuts2($linkadjusting & "\" & $file, $folder_to_move_root, $target_root)

		Else


			If StringRight($file, 4) = ".lnk" Then

				; Msgbox(1,"Link", $file)

				$aDetails = FileGetShortcut($linkadjusting & "\" & $file)

				If Not @error Then

					If StringLeft($aDetails[0], StringLen($folder_to_move_root)) = $folder_to_move_root Then

						MsgBox(0, $linkadjusting & "\" & $file, "Path: " & $aDetails[0])
						MsgBox(0, "needs to change !", "")

						FileDelete($linkadjusting & "\" & $file)

						$workdir = $aDetails[1]
						If StringLeft($aDetails[1], StringLen($folder_to_move_root)) = $folder_to_move_root Then
							$workdir = $target_root & StringMid($aDetails[1], StringLen($folder_to_move_root) + 1)
						EndIf

						$icon = $aDetails[4]
						If StringLeft($aDetails[4], StringLen($folder_to_move_root)) = $folder_to_move_root Then
							$icon = $target_root & StringMid($aDetails[1], StringLen($folder_to_move_root) + 1)
						EndIf

						MsgBox(0, "old target", $aDetails[0])
						MsgBox(0, "target", $target_root & StringMid($aDetails[0], StringLen($folder_to_move_root) + 1))
						MsgBox(0, "linkname", $linkadjusting & "\" & $file)


						FileCreateShortcut($target_root & StringMid($aDetails[0], StringLen($folder_to_move_root) + 1), _
								$linkadjusting & "\" & $file, _
								$workdir, $aDetails[2], $aDetails[3], _
								$icon, $aDetails[5], $aDetails[6])

					EndIf
				EndIf
			EndIf

		EndIf

	WEnd

	FileClose($search)


EndFunc   ;==>fm_checkshortcuts2


; create new Projekt-Folder


Func neues_projekt($path) ; wird vom menü gestartet, wenn der Benutzer den Menüpunkt auswählt

	Local $neuJahr = @YEAR
	$maxnumber = 0
	$i = 1

	While $i <= $Projects[0][0]
		If $Projects[$i][0] = $neuJahr Then
			If StringRight($Projects[$i][1], 4) > $maxnumber Then
				$maxnumber = StringRight($Projects[$i][1], 4)
			EndIf
		EndIf
		$i += 1
	WEnd

	$Form3 = GUICreate('Neues Projekt', 400, 200, -1, -1, $WS_SYSMENU, -1, $Form1)
	If $Iconfile <> "" Then
		GUISetIcon($Iconfile, $IconID)
	EndIf

	$Label1 = GUICtrlCreateLabel("Jahr", 8, 8, 24, 17)
	$inp_Neu_Jahr = GUICtrlCreateInput($neuJahr, 8, 24, 49, 21)
	$Label2 = GUICtrlCreateLabel("Nummer", 72, 8, 43, 17)
	$inp_Neu_Nummer = GUICtrlCreateInput(StringRight($neuJahr, 2) & $nL & StringRight(10001 + $maxnumber, 4), 72, 24, 81, 21) ; $nL is usualy "P"

	$btn_Auto_Nummer = GUICtrlCreateButton("generiere Projekt-Nummer", 170, 21, 150)

	$Label3 = GUICtrlCreateLabel("Name", 8, 56, 32, 17)
	$inp_Neu_Name = GUICtrlCreateInput("", 8, 72, 369, 21)
	$chk_Neu_Oeffnen = GUICtrlCreateCheckbox("Soll gleich der Ordner geöffnet werden?", 8, 104, 225, 17)

	$btn_Neu_OK = GUICtrlCreateButton("Ok", 200, 140, 80)
	$btn_Neu_Abbrechen = GUICtrlCreateButton("Abbrechen", 300, 140, 80)

	GUISetState(@SW_SHOW, $Form3)
	$modalwin = $Form3

EndFunc   ;==>neues_projekt

Func create_new_project($path) ; Erzeugt die Dateien, Verlinkungen und Feldeinträge für ein neues Projekt
	Local $i

	If Not FileExists($path & "\" & $neuJahr) Then
		DirCreate($path & "\" & $neuJahr)
	EndIf


	DirCreate($path & "\" & $neuJahr & "\" & $neuNummer & " " & $neuName)


	FileCreateShortcut($path & "\" & $neuJahr & "\" & $neuNummer & " " & $neuName, _
			$path & "\_Namen\" & $neuName & " (" & $neuNummer & ")")

	$i = $Projects[0][0] + 1
	$Projects[0][0] = $i
	$Projects[$i][0] = $neuJahr
	$Projects[$i][1] = $neuNummer
	$Projects[$i][2] = $neuName
	$Projects[$i][6] = $path & "\" & $neuJahr & "\" & $neuNummer & " " & $neuName


	$row_handels[$i] = GUICtrlCreateListViewItem($i & "|" & $Projects[$i][0] & "|" & $Projects[$i][1] & "|" & $Projects[$i][2] & "|" & $Projects[$i][3] & "|" & $Projects[$i][4] & "|" & $Projects[$i][5], $hListView)
	$row_handels[0] = $Projects[0][0]

	Local $LVcontextmenu = GUICtrlCreateContextMenu($row_handels[$i])
	$mnu_handels[$i][0] = GUICtrlCreateMenuItem("Löschen", $LVcontextmenu)
	$mnu_handels[$i][1] = GUICtrlCreateMenuItem("Umbenennen", $LVcontextmenu)
	$mnu_handels[$i][2] = GUICtrlCreateMenuItem("aktiv", $LVcontextmenu)
	If StringStripWS($Projects[$i][3], 3) <> "" Then GUICtrlSetState(-1, $GUI_CHECKED)
	$mnu_handels[$i][3] = GUICtrlCreateMenuItem("abgebrochen", $LVcontextmenu)
	If StringStripWS($Projects[$i][4], 3) <> "" Then GUICtrlSetState(-1, $GUI_CHECKED)
	$mnu_handels[$i][4] = GUICtrlCreateMenuItem("fertig", $LVcontextmenu)
	If StringStripWS($Projects[$i][5], 3) <> "" Then GUICtrlSetState(-1, $GUI_CHECKED)

	GUISetState(@SW_SHOW, $Form1)

	; if checked thann open in explorer
	If $neudoopen = $GUI_CHECKED Then
		ShellExecute($path & "\" & $neuJahr & "\" & $neuNummer & " " & $neuName)
	EndIf


EndFunc   ;==>create_new_project


; <-----------------------------------------------------------Gives the option to save the result as a report-------------------------------------------->


Func saveall($path)

	Local $g

	If $Projects[0][0] = 0 Then
		MsgBox(0, "Error", "keine Projekte")
		Return False
	Else

		$aFile = _WinAPI_GetSaveFileName("Speichern unter", "*.csv - Datei (*.csv)", ".", "", "csv", 2, 0, 0, $Form1)

		If $aFile[1] = "" Then ;Abbrechen
			Return False
		Else
			MsgBox(0, "", "x")
			$answer = $IDYES
			If FileExists(_WinAPI_PathAppend($aFile[1], $aFile[2])) Then
				$answer = MsgBox($MB_YESNO, "Datei existiert schon", "Die Datei " & _WinAPI_PathAppend($aFile[1], $aFile[2]) & " existiert schon" & @CRLF & @CRLF & "Soll die Datei überschrieben werden?")
			EndIf
			If $answer <> $IDYES Then
				Return False
			Else

				$handle = FileOpen(_WinAPI_PathAppend($aFile[1], $aFile[2]), 2)


				FileWriteLine($handle, "LfNr;Jahr;Nummer;Name;Aktiv;Abgebrochen;Fertig;Ordner")
				$g = 1
				While $g <= $Projects[0][0]
					FileWriteLine($handle, $g & ";" & $Projects[$g][0] & ";" & $Projects[$g][1] & ";" & $Projects[$g][2] & ";" & $Projects[$g][3] & ";" & $Projects[$g][4] & ";" & $Projects[$g][5] & ";" & $Projects[$g][6])
					$g += 1
				WEnd

				FileClose($handle)

				Return True
			EndIf
		EndIf

	EndIf

EndFunc   ;==>saveall






;-------------------------------------------------------- help-functions ------------------------------


Func test_if_folder_empty($folder)

	Local $search = FileFindFirstFile($folder & "\*.*")
	If ($search <> -1) Then
		FileClose($search)
		Return 1 ; folder is not empty
	Else
		If (@error = 1) Then
			Return 0 ; folder is empty
		Else
			Return 2 ; Folder kann nicht geöffnet werden
		EndIf
	EndIf
EndFunc   ;==>test_if_folder_empty

Func test_if_proj_folder($folder)

	; if first char is a number and secound is a number or "_" and
	; fourth to seventh are numbers, than it is a projectfolder

	; true Examples: 99P0001 abc         0_P0003_test       9_P0023 qwr

	If StringLen($folder) >= 7 Then

		If (char_is_number(StringMid($folder, 1, 1)) = True) Then
			If (char_is_number(StringMid($folder, 2, 1)) = True) Or (StringMid(2, 1) = "_") Then
				If (char_is_number(StringMid($folder, 4, 1)) = True) And (char_is_number(StringMid($folder, 5, 1)) = True) Then
					If (char_is_number(StringMid($folder, 6, 1)) = True) And (char_is_number(StringMid($folder, 7, 1)) = True) Then
						Return True
					EndIf
				EndIf
			EndIf
		EndIf
	EndIf
	Return False
EndFunc   ;==>test_if_proj_folder

Func test_if_year_folder($folder)



	; all four-digit-integers return true
	; first two digits are integers and than "xx" returns true
	; other returns false

	; true Examples = 1999    2000   20xx

	If StringLen($folder) = 4 Then
		If (char_is_number(StringMid($folder, 1, 1)) = True) And (char_is_number(StringMid($folder, 2, 1)) = True) Then
			If (char_is_number(StringMid($folder, 3, 1)) = True) And (char_is_number(StringMid($folder, 4, 1)) = True) Then
				return True
			ElseIf StringMid($folder, 3, 2) = "xx" Then
				return True
			Else
				return False
			EndIf
		EndIf
	EndIf
	return false
EndFunc   ;==>test_if_year_folder

Func test_if_yearsubfolder($folder)

	;test if the folder contains subfolders for the years.

	Local $file
	Local $search = FileFindFirstFile($folder & "\*")

	Local $flag = False

	If $search <> -1 Then

		While 1
			$file = FileFindNextFile($search)

			If @error Then ExitLoop

			If test_if_year_folder($file) Then
				$flag = True
				ExitLoop
			EndIf
		WEnd

	EndIf

	FileClose($search)

	Return $flag

EndFunc   ;==>test_if_yearsubfolder

Func char_is_number($char)

	; all numbers from 0 to 9 return true other gives false

	If (Asc($char) >= 48) And (Asc($char) <= 57) Then
		Return True
	Else
		Return False
	EndIf
EndFunc   ;==>char_is_number

Func ckeckshortcutlinks($path, ByRef $reporttable)

	; $reporttabele[$row][[$col]
	; $reporttabell has three columns:
	; - Name from the Shortcut
	; - Problem: 1 - Target, 2 - Workdir, 3 - Icon
	; - Entry which is broken

	Local $aDetails

	Local $search = FileFindFirstFile($path & "\*")
	Local $error = 0

	If $search = -1 Then
		If @error <> 1 Then ; if @error = 1 then remains $error = 0		Folder is empty, but no error
			$error = -1 ; some unkonwn Error
		EndIf
	Else
		While 1
			$file = FileFindNextFile($search)
			If @error Then
				ExitLoop ; no next file
			EndIf
			If @extended Then ; $file is directory
				$error = $error Or ckeckshortcutlinks($path & "\" & $file, $reporttable)
			Else
				$aDetails = FileGetShortcut($path & "\" & $file)

				If StringStripWS($aDetails[0], 3) = "" Then ; check if shortcuttargeg is not empty
					extendarray2D($reporttable)
					$reporttable[$reporttable[0][0] + 1][0] = ""
					$reporttable[$reporttable[0][0] + 1][1] = 1
					$reporttable[$reporttable[0][0] + 1][2] = $aDetails[0]
					$reporttable[0][0] = $reporttable[0][0] + 1
				Else
					If Not FileExists($aDetails[0]) Then ; chek shortcuttarget
						extendarray2D($reporttable)
						$reporttable[$reporttable[0][0] + 1][0] = $path & "\" & $file
						$reporttable[$reporttable[0][0] + 1][1] = 1
						$reporttable[$reporttable[0][0] + 1][2] = $aDetails[0]
						$reporttable[0][0] = $reporttable[0][0] + 1
					EndIf
				EndIf

				If StringStripWS($aDetails[1], 3) <> "" Then
					If Not FileExists($aDetails[1]) Then ; check shortcutworkdir
						extendarray2D($reporttable)
						$reporttable[$reporttable[0][0] + 1][0] = $path & "\" & $file
						$reporttable[$reporttable[0][0] + 1][1] = 2
						$reporttable[$reporttable[0][0] + 1][2] = $aDetails[1]
						$reporttable[0][0] = $reporttable[0][0] + 1
					EndIf
				EndIf

				If StringStripWS($aDetails[4], 3) <> "" Then
					If Not FileExists($aDetails[4]) Then ; check shortcuticonpath
						extendarray2D($reporttable)
						$reporttable[$reporttable[0][0] + 1][0] = $path & "\" & $file
						$reporttable[$reporttable[0][0] + 1][1] = 3
						$reporttable[$reporttable[0][0] + 1][2] = $aDetails[2]
						$reporttable[0][0] = $reporttable[0][0] + 1
					EndIf
				EndIf
			EndIf
		WEnd
		FileClose($search)
	EndIf

	Return $error
EndFunc   ;==>ckeckshortcutlinks


; $arrary[0] (or [0][0] or [0][0][0])  contanis the number of used places/$rows/$layer in the array without $array[0](..).
; empty -> $array[0](..) = 0
; first element/row/layer -> $array[1]				$array[1][i]			$array[1][i][j]
; last element/row/layer -> $array[$array[0]]		$array[$array[0]][i]	$array[$array[0]][i][j]

; the following funcions make sure, that there is at least on free place/$row/$layer at the end/bottom of the array
; after calling this function you allways can write to $array[$array[0]+1]
; if done do not forget to set $array[0] = $array[0] + 1

Func extendarray1D(ByRef $array)
	If UBound($array) = $array[0] + 1 Then
		ReDim $array[UBound($array) * 2]
	EndIf
EndFunc   ;==>extendarray1D

Func extendarray2D(ByRef $array)
	If UBound($array, 1) = $array[0][0] + 1 Then
		ReDim $array[UBound($array, 1) * 2][UBound($array, 2)]
	EndIf
EndFunc   ;==>extendarray2D

Func extendarray3D(ByRef $array)
	If UBound($array, 1) = $array[0][0][0] + 1 Then
		ReDim $array[UBound($array, 1) * 2][UBound($array, 2)][UBound($array, 3)]
	EndIf
EndFunc   ;==>extendarray3D

; ----

func Drive_and_File_Exists($name)
	If Not ("READY"==DriveStatus(StringLeft($name,2))) Then
		return 0
	endif
	return FileExists($name)
endFunc

Func PathIsRelative($xpath)
	IF ("..\"=StringLeft($xpath,3)) Then
		return 1
	Else
		return 0
	EndIf
EndFunc

Func PathIsRoot($xpath)
	If (":\"=StringMid($xpath,2,2)) Then
		return 1
	Else
		return 0
	EndIF
EndFunc


