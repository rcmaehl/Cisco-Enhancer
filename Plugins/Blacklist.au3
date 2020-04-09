#include <File.au3>
#include <FileConstants.au3>

Main()

Func Main()

	Local $sIn, $sStatus

	While True

		$sIn = ConsoleRead()
		If @extended = 0 Then
			; Uncomment the following line if you only want the code ran on Agent Status Changes
			; ContinueLoop
		Else
			$sStatus = $sIn
		EndIf

		Switch $sStatus

			Case "Not Ready", "NOT_READY"

			Case "Ready", "READY"
				_ProcessBlacklist()

			Case "Reserved"
				_ProcessBlacklist()

			Case "Talking"
				_ProcessBlacklist()

			Case "WrapUp", "WORK_READY"
				_ProcessBlacklist()

			Case "LOGOUT"

			Case "HOLD"
				_ProcessBlacklist()

			Case Else
				ConsoleWrite("Plugin Caught unhandled agent state: " & $sStatus)

		EndSwitch

	WEnd
EndFunc

Func _ProcessBlacklist()
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
EndFunc
