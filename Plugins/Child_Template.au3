Main()

Func Main()

	Local $sStatus

	While True

		$sStatus = ConsoleRead()
		If @extended = 0 Then ContinueLoop

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
