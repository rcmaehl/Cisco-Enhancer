#include "./FinesseConstants.au3"

; #FUNCTION# ====================================================================================================================
; Name ..........: _FinesseGetQueue
; Description ...: This UDF allows a user to get a Queue object. Use this API to access statistics for a queue that is assigned
;                  to agents or supervisors. If you use this UDF to get a queue that is not assigned to any users, the response
;                  contains a value of -1 for numeric statistics and is empty for string statistics.
; Syntax ........: _FinesseGetQueue($sAPI, $iFrom, $iTo, $sUser, $sPassword)
; Parameters ....: $sAPI                - Finesse API URL
;                  $iID                 - Queue ID
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
; Remarks .......: Any user can use this UDF to retrieve information about a specific queue. The user does not need to belong
;                  to that queue.
; Related .......:
; Link ..........: https://solutionpartner.cisco.com/media/finesseDevGuide2/CFIN_RF_Q5AF8C66_00_queue-get-queue.html
; Example .......: No
; ===============================================================================================================================
Func _FinesseGetQueue($sAPI, $iID, $sUser, $sPassword)

	Local $oHTTP = ObjCreate("WinHttp.WinHttpRequest.5.1")

	$oHTTP.Open("GET", $sAPI & "Queue/" & $iID, False)
	If @error Then Return SetError(1, @extended, @error)

	$oHTTP.SetCredentials($sUser, $sPassword, 0)

	$oHTTP.SetRequestHeader("Content-Type", "application/xml")

	$oHTTP.Send()
	If @error Then Return SetError(2, 0, 0)

	If $oHTTP.Status <> $FINESSE_STATUS_SUCCESS Then Return SetError($oHTTP.Status, 0, $oHTTP.ResponseText)

	Return SetError(0, 0, True)

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _FinesseGetUserQueues
; Description ...: This UDF allows a user to get a list of all queues associated with that user.
; Syntax ........: _FinesseGetUserQueues($sAPI, $iFrom, $iTo, $sUser, $sPassword)
; Parameters ....: $sAPI                - Finesse API URL
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
; Remarks .......: All users can use this UDF to retrieve a list of queues for any user.
; Related .......:
; Link ..........: https://solutionpartner.cisco.com/media/finesseDevGuide2/CFIN_RF_QEB7CD10_00_queue-get-list.html
; Example .......: No
; ===============================================================================================================================
Func _FinesseGetUserQueues($sAPI, $sUser, $sPassword)

	Local $oHTTP = ObjCreate("WinHttp.WinHttpRequest.5.1")

	$oHTTP.Open("GET", $sAPI & "User/" & $sUser & "/Queues", False)
	If @error Then Return SetError(1, @extended, @error)

	$oHTTP.SetCredentials($sUser, $sPassword, 0)

	$oHTTP.SetRequestHeader("Content-Type", "application/xml")

	$oHTTP.Send()
	If @error Then Return SetError(2, 0, 0)

	If $oHTTP.Status <> $FINESSE_STATUS_SUCCESS Then Return SetError($oHTTP.Status, 0, $oHTTP.ResponseText)

	Return SetError(0, 0, True)

EndFunc