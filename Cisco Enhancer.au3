#include <File.au3>
#include <Array.au3>
#include <WinAPIEx.au3>
#include <TrayConstants.au3>
#include <GUIConstantsEx.au3>
#include <ButtonConstants.au3>
#include <MsgBoxConstants.au3>
#include <StringConstants.au3>
#include <WindowsConstants.au3>

#include ".\FinesseUDF\FinesseUser.au3"
#include ".\FinesseUDF\FinesseDialog.au3"

Opt("TrayMenuMode", 3) ; Disable Default Tray menu

Global Const $CTRL_ALL = 0
Global Const $CTRL_CREATED = 1
Global Const $ABM_GETTASKBARPOS = 0x5

Main()

Func Main()

	Local $TrayOpts = TrayCreateItem("Settings")
	TrayCreateItem("")
	Local $TrayExit = TrayCreateItem("Exit"    )

	Local $hGUI = GUICreate("Settings", 280, 80, @DesktopWidth - 300, @DesktopHeight - 150, BitXOR($GUI_SS_DEFAULT_GUI, $WS_MINIMIZEBOX), $WS_EX_TOOLWINDOW + $WS_EX_TOPMOST)

	Local $hWindowLess = GUICtrlCreateCheckbox("Close Non-Work Windows while Talking"    , 10, 00, 260, 20, $BS_RIGHTBUTTON)
	Local $hCTIToolkit = GUICtrlCreateCheckbox("Show Reminders for Non-Ready Status"     , 10, 20, 260, 20, $BS_RIGHTBUTTON)
	Local $hUseCTILess = GUICtrlCreateCheckbox("Use 1 Minute Wrap-Up Reminders"          , 10, 40, 260, 20, $BS_RIGHTBUTTON)
	Local $honStartup  = GUICtrlCreateCheckbox("Start with Windows"                      , 10, 60, 260, 20, $BS_RIGHTBUTTON)
	GUICtrlSetTip($hCTIToolkit, "Display 2 minute reminders for Wrap-Up and" & @CRLF & "15 minute, 30 minute, & 1 hour reminders for Not Ready")
	GUICtrlSetTip($hUseCTILess, "Use 1 minute reminders for Wrap-Up to reduce AHT compared to 2 minutes")

	$aSettings = _LoadSettings()
	GUICtrlSetState($hCTIToolkit, Execute($aSettings[0]))
	GUICtrlSetState($hUseCTILess, Execute($aSettings[1]))
	GUICtrlSetState($hOnStartup , Execute($aSettings[2]))

	Local $aFinesse = Execute($aSettings[3])
	Local $sAPI = $aSettings[4]
	Local $iUser = $aSettings[5]
	Local $iPass = $aSettings[6]


	Local $iNRC = 0
	Local $iPoll = 500
	Local $aCtrls = Null
	Local $bCLock = False
	Local $bLLock = False
	Local $bTimer = False
	Local $hTimer = TimerInit()
	Local $hMsgBox = Null
	Local $hTaskBar = Null
	Local $hNRTimer = TimerInit()
	Local $hLastPoll = TimerInit()
	Local $bReserved = False


	While 1

		$hGMsg = GUIGetMsg()
		$hTMsg = TrayGetMsg()

		Select

			Case $hTMsg = $TrayOpts
				GUISetState(@SW_SHOW, $hGUI)
				$hTaskBar = _GetTaskBarPos()
				#cs
				Select
					Case $hTaskBar[1] = 0 And $hTaskBar[2] = 0
						If $hTaskBar[3] > $hTaskBar[4] Then ; Top
							WinMove($hGUI, "", $hTaskBar[4], @DesktopWidth - 280)
						ElseIf $hTaskBar[3] < $hTaskBar[4] Then ; Left
							WinMove($hGUI, "", @DesktopHeight - 60, $hTaskBar[3], -1, -1, 50)
						Else ; Throw Error
							MsgBox($MB_ICONQUESTION, "Error", "Unexpected Taskbar Positioning, Please Contact Developer" & @CRLF & _
								"Positioning: " & $hTaskBar[1] & "," & $hTaskBar[2] & ", " & $hTaskBar[3] & ", " & $hTaskBar[4])
						EndIf
					Case $hTaskBar[1] = 0 And Not $hTaskBar[2] = 0 ; Bottom
				#ce
						WinMove($hGUI, "", @DesktopWidth - 300, $hTaskBar[2] - 120)
				#cs
					Case Not $hTaskBar[1] = 0 And $hTaskBar[2] = 0 ; Right
						WinMove($hGUI, "", @DesktopHeight - 60, $hTaskBar[1] - 280)
					Case Not $hTaskBar[1] = 0 And Not $hTaskBar[2] = 0
						MsgBox($MB_ICONQUESTION, "Error", "Unexpected Taskbar Positioning, Please Contact Developer" & @CRLF & _
							"Positioning: " & $hTaskBar[1] & "," & $hTaskBar[2] & ", " & $hTaskBar[3] & ", " & $hTaskBar[4])
						WinMove($hGUI, "", -1, -1)
				EndSelect
				#ce

			Case $hGMsg = $GUI_EVENT_CLOSE
				GUISetState(@SW_HIDE, $hGUI)

			Case $hTMsg = $TrayExit
				Exit

			Case $hGMsg = $hCTIToolkit Or $hGMsg = $hUseCTILess Or $hGMsg = $honStartup
				If _IsChecked($hOnStartup) And Not FileExists(@StartupDir & "\ITSD Desktop Tools.lnk") Then
					FileCreateShortcut(@AutoItExe, @StartupDir & "\CTI Enhancer.lnk")
				ElseIf Not _IsChecked($honStartup) And FileExists(@StartupDir & "\Cisco Enhancer.lnk") Then
					FileDelete(@StartupDir & "\Cisco Enhancer.lnk")
				EndIf
				_SaveSettings($hCTIToolkit, $hUseCTILess, $hOnStartup)

			Case WinExists("CTI Toolkit Agent Desktop") Or $aFinesse

				If Not _IsChecked($hCTIToolkit) Then
					TraySetToolTip("Not Running - Disabled in Settings")
					ContinueLoop
				EndIf

				TraySetToolTip("Running...")

					If WinExists("CTI Toolkit Agent Desktop") Then
						$sStatus = StatusbarGetText("CTI Toolkit Agent Desktop", "", "4")
						$sStatus = StringStripWS($sStatus, $STR_STRIPLEADING + $STR_STRIPTRAILING)
						$sStatus = StringReplace($sStatus, "Agent Status: ", "")
					ElseIf WinExists("Cisco Finesse") Then
						If TimerDiff($hLastPoll) >= $iPoll Then ; Rate Limiting is important
							$iPoll = 500
							$hLastPoll = TimerInit()
							$sStatus = _FinesseGetState($sAPI, $iUser, $iPass)
							If StringInStr($sStatus, "SERVER_OUT_OF_SERVICE") Then $iPoll = 30000
							If @error Then
								ConsoleWrite($sStatus & @CRLF & _
								 @error & @CRLF & _
								 @extended & @CRLF)
							EndIf
						EndIf
						Else
					EndIf

					$bCLock = _GetDesktopLock()

					Switch $sStatus

						Case "Not Ready", "NOT_READY"
							If Not $bTimer Then
								$bTimer = True
								$hNRTimer = TimerInit()
							ElseIf TimerDiff($hNRTimer) >= 900000 Then
								$iNRC += 1
								If $iNRC = 3 Or $iNRC = 6 Or $iNRC = 7 Then
									;;;
								Else
									WinMinimizeAll()
									$hMsgBox = MsgBox($MB_YESNO + $MB_ICONWARNING + $MB_TOPMOST, "Reminder", "You've been in Not Ready Status for over " & $iNRC * 15 & " minutes. Would you like to go back to Ready Status?", 15)
									If $hMsgBox = $IDYES Then
										If WinExists("CTI Toolkit Agent Desktop") Then
											WinMove("CTI Toolkit Agent Desktop", "", Default, Default, 940, 360)
											$aCtrls = _WinGetHandleListFromPos("CTI Toolkit Agent Desktop", "", 83, 15, $CTRL_CREATED)
											Opt("WinTitleMatchMode", 2)
											ControlSend("CTI Toolkit Agent Desktop", "", $aCtrls[2], "{ENTER}")
										ElseIf WinExists("Cisco Finesse") Then
											_FinesseSetState($sAPI, "READY", "", $iUser, $iPass)
										EndIf
									EndIf
									WinMinimizeAllUndo()
									$hMsgBox = Null
								EndIf

								$hNRTimer = TimerInit()
							EndIf

							If $bCLock = True Then ; If Desktop is locked
								$bLLock = True
								$CLock = Null
							ElseIf $bCLock = False And $bLLock = True Then ; If Desktop is unlocked but WAS locked
								$bLLock = False
								$bCLock = True
								$hMsgBox = MsgBox($MB_YESNO + $MB_ICONWARNING + $MB_TOPMOST, "Reminder", "You've just logged back in while Not Ready. You've been Not Ready for " & ($iNRC * 15) + Floor(TimerDiff($hNRTimer) / 60000) & " minutes. Would you like to go back to Ready Status?", 30)
								If $hMsgBox = $IDYES Then
									If WinExists("CTI Toolkit Agent Desktop") Then
										WinMove("CTI Toolkit Agent Desktop", "", Default, Default, 940, 360)
										$aCtrls = _WinGetHandleListFromPos("CTI Toolkit Agent Desktop", "", 83, 15, $CTRL_CREATED)
										Opt("WinTitleMatchMode", 2)
										ControlSend("CTI Toolkit Agent Desktop", "", $aCtrls[2], "{ENTER}")
									ElseIf WinExists("Cisco Finesse") Then
										_FinesseSetState($sAPI, "READY", "", $iUser, $iPass)
									EndIf
								EndIf
								$hMsgBox = Null
							Else
								$bLLock = $bCLock
								$CLock = Null
							EndIf

						Case "Ready", "READY"
							If $bTimer And $bReserved Then
								FileWrite(@ScriptDir & "\IdleTime.log", "RONA/Hangup, " & Round(TimerDiff($hNRTimer) / 1000, 2) & "s" & @CRLF)
;								MsgBox($MB_OK + $MB_ICONWARNING + $MB_TOPMOST, "ALERT", "Ready/Reserved Issue Detected.", 30)
								$hBTimer = TimerInit()
								$bReserved = False
							EndIf
							$iNRC = 0
							$bTimer = False
							$hNRTimer = TimerInit()

						Case "Reserved"
							If Not $bReserved Then
								FileWrite(@ScriptDir & "\IdleTime.log", Round(TimerDiff($hBTimer) / 1000, 2) & "s, ")
							EndIf
							$iNRC = 0
							$bTimer = True
							If Not $bReserved Then
								$hNRTimer = TimerInit()
								$bReserved = True
							EndIf

						Case "Talking"
							If $bReserved Then
								FileWrite(@ScriptDir & "\IdleTime.log", "Answered, " & Round(TimerDiff($hNRTimer) / 1000, 2) & "s" & @CRLF)
							EndIf
							$iNRC = 0
							$bTimer = False
							$hBTimer = TimerInit()
							$hNRTimer = TimerInit()
							$bReserved = False
							If _IsChecked($hWindowLess) Then
								$hList = FileOpen(@ScriptDir & "\Blacklist.def", $FO_READ + $FO_CREATEPATH)
								If $hList = -1 Then
									FileClose($hList)
								Else
									$iLines = _FileCountLines(@ScriptDir & "\Blacklist.def")
									For $iLine = 1 to $iLines Step 1
										$sLine = FileReadLine($hList, $iLine)
										$aTask = StringSplit($sLine, ",")
										If $aTask[0] < 3 Then ContinueLoop
										If $aTask[0] > 3 Then
											For $i = 4 To $aTask[0]
												$aTask[3] &= $aTask[$i]
											Next
										EndIf
										Opt("WinTitleMatchMode", -2)
										If WinActive($aTask[1], $aTask[2]) Then
											Send($aTask[3])
											Sleep(500)
										EndIf
									Next
									Opt("WinTitleMatchMode", 2)
									FileClose($hList)
								EndIf
							EndIf

						Case "WrapUp", "WORK_READY"
							If Not $bTimer Then
								$bTimer = True
								$hTimer = TimerInit()
							ElseIf _IsChecked($hUseCTILess) And TimerDiff($hNRTimer) >= 60000 Then
								$iNRC += 1
								$hMsgBox = MsgBox($MB_YESNO + $MB_ICONWARNING + $MB_TOPMOST, "Reminder", "You've been in WrapUp Status for over " & 1 * $iNRC & " minute(s). Would you like to go back to Ready Status?", 15)
								If $hMsgBox = $IDYES Then
									If WinExists("CTI Toolkit Agent Desktop") Then
										WinMove("CTI Toolkit Agent Desktop", "", Default, Default, 940, 360)
										$aCtrls = _WinGetHandleListFromPos("CTI Toolkit Agent Desktop", "", 83, 15, $CTRL_CREATED)
										Opt("WinTitleMatchMode", 2)
										ControlSend("CTI Toolkit Agent Desktop", "", $aCtrls[2], "{ENTER}")
									ElseIf WinExists("Cisco Finesse") Then
										_FinesseSetState($sAPI, "READY", "", $iUser, $iPass)
									EndIf
								EndIf
								$hMsgBox = Null
								$hNRTimer = TimerInit()
							ElseIf TimerDiff($hNRTimer) >= 120000 Then
								$iNRC += 1
								$hMsgBox = MsgBox($MB_YESNO + $MB_ICONWARNING + $MB_TOPMOST, "Reminder", "You've been in WrapUp Status for over " & 2 * $iNRC & "minutes. Would you like to go back to Ready Status?", 15)
								If $hMsgBox = $IDYES Then
									If WinExists("CTI Toolkit Agent Desktop") Then
										WinMove("CTI Toolkit Agent Desktop", "", Default, Default, 940, 360)
										$aCtrls = _WinGetHandleListFromPos("CTI Toolkit Agent Desktop", "", 83, 15, $CTRL_CREATED)
										Opt("WinTitleMatchMode", 2)
										ControlSend("CTI Toolkit Agent Desktop", "", $aCtrls[2], "{ENTER}")
									ElseIf WinExists("Cisco Finesse") Then
										_FinesseSetState($sAPI, "READY", "", $iUser, $iPass)
									EndIf
								EndIf
								$hMsgBox = Null
								$hNRTimer = TimerInit()
							EndIf

						Case "LOGOUT"

						Case "HOLD"

						Case Else
							ConsoleWrite($sStatus & @CRLF)

					EndSwitch

			Case Else

				TraySetToolTip("Not Running - Unable to find CTI Window to Interface")

		EndSelect

	WEnd

EndFunc   ;==>Main

; Additional Functions Called

Func _IsChecked($idControlID)
	Return BitAND(GUICtrlRead($idControlID), $GUI_CHECKED) = $GUI_CHECKED
EndFunc   ;==>_IsChecked

Func _GetDesktopLock()
    Local $fIsLocked = False
    Local Const $hDesktop = _WinAPI_OpenDesktop('Default', $DESKTOP_SWITCHDESKTOP)
    If @error = 0 Then
        $fIsLocked = Not _WinAPI_SwitchDesktop($hDesktop)
        _WinAPI_CloseDesktop($hDesktop)
    EndIf
    Return $fIsLocked
EndFunc   ;==>_GetDesktopLock

Func _GetTaskBarPos()
    $h_taskbar = WinGetHandle("","Start")
    $AppBarData = DllStructCreate("dword;int;uint;uint;int;int;int;int;int")
;~ DWORD cbSize;
;~   HWND hWnd;
;~   UINT uCallbackMessage;
;~   UINT uEdge;
;~   RECT rc;
;~   LPARAM lParam;
    DllStructSetData($AppBarData,1,DllStructGetSize($AppBarData))
    DllStructSetData($AppBarData,2,$h_taskbar)
    $lResult = DllCall("shell32.dll","int","SHAppBarMessage","int",$ABM_GETTASKBARPOS,"ptr",DllStructGetPtr($AppBarData))
    If Not @error Then
        If $lResult[0] Then
            Return StringSplit(DllStructGetData($AppBarData,5) & "|" & _
                DllStructGetData($AppBarData,6) & "|"   & DllStructGetData($AppBarData,7) & "|" & _
                DllStructGetData($AppBarData,8),"|")
        EndIf
    EndIf
    SetError(1)
    Return 0
EndFunc   ;==>_GetTaskBarPos

Func _LoadSettings()
	Local $aSettings[7]
	$aSettings[0] = IniRead(".\CiscoE.ini", "Cisco"  , "Monitor Status"  , False)
	$aSettings[1] = IniRead(".\CiscoE.ini", "Cisco"  , "1 Min Reminder"  , False)
	$aSettings[2] = IniRead(".\CiscoE.ini", "CiscoE" , "Start w. Windows", False)
	$aSettings[3] = IniRead(".\CiscoE.ini", "CiscoE" , "Use Fineese APIs", False)
	$aSettings[4] = IniRead(".\CiscoE.ini", "Fineese", "API URL"         , False)
	$aSettings[5] = IniRead(".\CiscoE.ini", "Fineese", "User ID"         , False)
	$aSettings[6] = IniRead(".\CiscoE.ini", "Fineese", "Password"        , False)
	Return $aSettings
EndFunc   ;==>_LoadSettings

Func _SaveSettings($bToolkit, $bRemind, $bStartup)
	$bToolkit   = _IsChecked($bToolkit)
	$bRemind    = _IsChecked($bRemind )
	$bStartup   = _IsChecked($bStartup)
	IniWrite(".\CiscoE", "Cisco"  , "Monitor Status"  , $bToolkit)
	IniWrite(".\CiscoE", "Cisco"  , "1 Min Reminder"  , $bRemind )
	IniWrite(".\CiscoE", "CiscoE" , "Start w. Windows", $bStartup)
EndFunc   ;==>_SaveSettings

Func _WinGetHandleListFromPos($sTitle, $sText, $iX, $iY, $dFlags = $CTRL_ALL)

	If $dFlags > $CTRL_CREATED Then
		$dFlags = $CTRL_CREATED
	EndIf

	Local $hWnd

	$hWnd = WinGetHandle($sTitle, $sText)
	If @error Then SetError(1, @error, 0)

	Local $sClasses

	$sClasses = WinGetClassList($hWnd)
	If @error Then SetError(2, @error, 0)

	Local $aClasses

	$aClasses = StringSplit($sClasses, @LF) ; For some WEIRD reason, WinGetClassList returns a @LF seperated list instead of an array
	If @error Then
		SetError(3, 0, 0)
	ElseIf $aClasses[0] = "0" Then
		SetError(3, 1, 0)
	EndIf

	ReDim $aClasses[UBound($aClasses) - 1] ; Remove Trailing Blank Class due to Extra @LF
	$aClasses[0] = UBound($aClasses) - 1 ; Set Array Size Value Based on New Size

	Local $iInstance = 0

	For $iLoop = 1 To $aClasses[0] Step 1
		$iInstance = 0
		Do
			$iInstance += 1 ; Increment Instance count for a Class ID
			_ArraySearch($aClasses, ControlGetHandle($hWnd, "", "[CLASS:" & $aClasses[$iLoop] & "; INSTANCE: " & $iInstance & "]")) ; See a handle has already been found for that ID + Instance
		Until @error ; If so, Exit, Otherwise Repeat
		$aClasses[$iLoop] = ControlGetHandle($hWnd, "", "[CLASS:" & $aClasses[$iLoop] & "; INSTANCE: " & $iInstance & "]")
	Next

	Local $aCoords, $iOut = 0, $aOut[1]

	For $iLoop = 1 To $aClasses[0] Step 1
		$aCoords = ControlGetPos($hWnd, "", $aClasses[$iLoop])
		If $dFlags = $CTRL_ALL Then
			If $aCoords[0] <= $iX And $aCoords[0] + $aCoords[2] >= $iX And $aCoords[1] <= $iY And $aCoords[1] + $aCoords[3] >= $iY Then
				$iOut += 1
				$aOut[0] = $iOut
				ReDim $aOut[UBound($aOut) + 1]
				$aOut[$iOut] = $aClasses[$iLoop]
			EndIf
		ElseIf $dFlags = $CTRL_CREATED Then
			If $aCoords[0] = $iX And $aCoords[1] = $iY Then
				$iOut += 1
				$aOut[0] = $iOut
				ReDim $aOut[UBound($aOut) + 1]
				$aOut[$iOut] = $aClasses[$iLoop]
			EndIf
		EndIf
	Next

	If $aOut[0] = 0 Then SetError(4, 0, 0)
	Return $aOut

EndFunc   ;==>_WinGetHandleListFromPos