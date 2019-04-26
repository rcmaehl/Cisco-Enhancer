#include "./FinesseConstants.au3"

; #FUNCTION# ====================================================================================================================
; Name ..........: _FinesseGetCall
; Description ...: This UDF allows a user to get a copy of a Dialog (call) object.
; Syntax ........: _FinesseGetCall($sAPI, $iID, $sUser, $sPassword)
; Parameters ....: $sAPI                - Finesse API URL
;                  $iID                 - Dialog (Call) Id
;                  $sUser               - The ID of the user
;                  $sPassword           - The Password of the user
; Return values .: Success - Returns True
;                  Failure - Sets @error:
;                  |1   - WinHTTP Open Error, see Return Value for @error and @extended for @extended
;                  |2   - WinHTTP Send Error, see Return Value for @error and @extended for @extended
;                  |400 - Bad Request, See Returned Data (Finesse)
;                  |401 - Unauthorized, See Returned Data (Finesse)
;                  |404 - Not Found (Finesse)
;                  |500 - Internal Server Error (Finesse)
;                  |503 - Service Unavailable (Finesse)
; Author ........: Robert Maehl
; Modified ......: 7/11/2018 11:00 EST
; Remarks .......: Agents and administrators can use this API. Agents can only get their own Dialog object. Administrators can
;                  get any Dialog object.
; Related .......:
; Link ..........: https://solutionpartner.cisco.com/media/finesseDevGuide2/CFIN_RF_D6A2EEC4_00_dialog-get-dialog.html
; Example .......: No
; ===============================================================================================================================
Func _FinesseGetCall($sAPI, $iID, $sUser, $sPassword)

	Local $oHTTP = ObjCreate("WinHttp.WinHttpRequest.5.1")

	$oHTTP.Open("GET", $sAPI & "Dialog/" & $iID, False)
	If @error Then Return SetError(1, @extended, @error)

	$oHTTP.SetCredentials($sUser, $sPassword, 0)

	$oHTTP.SetRequestHeader("Content-Type", "application/xml")

	$oHTTP.Send()
	If @error Then Return SetError(2, 0, 0)

	If $oHTTP.Status <> $FINESSE_STATUS_SUCCESS Then Return SetError($oHTTP.Status, 0, $oHTTP.ResponseText)

	Return SetError(0, 0, True)

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _FinesseMakeCall
; Description ...: This UDF allows a user to make a call. To make a call, a new Dialog object is created that specifies the
;                  the caller's extension and the destination target. The new Dialog object is posted to the Dialog (Call)
;                  collection for that user. This UDF supports the use of any ASCII character in the destination target.
;                  Finesse does not convert any entered letters into numbers, nor does it remove non-numeric characters
;                  (including parentheses and hyphens) from the toAddress.
; Syntax ........: _FinesseMakeCall($sAPI, $iFrom, $iTo, $sUser, $sPassword)
; Parameters ....: $sAPI                - Finesse API URL
;                  $iFrom               - The extension with which the user is currently signed in
;                  %sTo                 - The destination for the call
;                  $sUser               - The ID of the user
;                  $sPassword           - The Password of the user
; Return values .: Success - Returns True
;                  Failure - Sets @error:
;                  |1   - WinHTTP Open Error, see Return Value for @error and @extended for @extended
;                  |2   - WinHTTP Send Error, see Return Value for @error and @extended for @extended
;                  |400 - Bad Request, See Returned Data (Finesse)
;                  |401 - Unauthorized, See Returned Data (Finesse)
;                  |404 - Not Found (Finesse)
;                  |500 - Internal Server Error (Finesse)
;                  |503 - Service Unavailable (Finesse)
; Author ........: Robert Maehl
; Modified ......: 7/11/2018 11:00 EST
; Remarks .......: All users can use this API. Users can only create dialogs using a fromAddress to which they are currently
;                  signed in.
; Related .......:
; Link ..........: https://solutionpartner.cisco.com/media/finesseDevGuide2/CFIN_RF_DC4120C3_00_dialog-create-a-new-dialog.html
; Example .......: No
; ===============================================================================================================================
Func _FinesseMakeCall($sAPI, $iFrom, $sTo, $sUser, $sPassword)

	Local $oHTTP = ObjCreate("WinHttp.WinHttpRequest.5.1")

	$oHTTP.Open("POST", $sAPI & "User/" & $sUser & "/Dialogs", False)
	If @error Then Return SetError(1, @extended, @error)

	$oHTTP.SetCredentials($sUser, $sPassword, 0)

	$oHTTP.SetRequestHeader("Content-Type", "application/xml")

	$oHTTP.Send("<Dialog><requestedAction>MAKE_CALL</requestedAction><fromAddress>" & $iFrom & "</fromAddress><toAddress>" & $sTo & "</toAddress></Dialog>")
	If @error Then Return SetError(2, 0, 0)

	If $oHTTP.Status <> $FINESSE_STATUS_ACCEPTED Then Return SetError($oHTTP.Status, 0, $oHTTP.ResponseText)

	Return SetError(0, 0, True)

EndFunc

