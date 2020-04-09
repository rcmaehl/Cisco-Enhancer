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

			Case "Reserved"

			Case "Talking"

			Case "WrapUp", "WORK_READY"

			Case "LOGOUT"

			Case "HOLD"

			Case Else
				ConsoleWrite("Plugin Caught unhandled agent state: " & $sStatus)

		EndSwitch

	WEnd
EndFunc
