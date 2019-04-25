#include <String.au3>
#include "./FinesseConstants.au3"

; #FUNCTION# ====================================================================================================================
; Name ..........: _FinesseGetCalls
; Description ...: This API allows an agent or administrator to get a list of dialogs (calls) associated with a particular user.
; Syntax ........: _FinesseGetCalls($sAPI, $sUser, $sPassword)
; Parameters ....: $sAPI                - Finesse API URL
;                  $sUser               - The ID of the user
;                  $sPassword           - The Password of the user
; Return values .: Success - Returns Response from Finesse
;                  Failure - Sets @error:
;                  |1   - WinHTTP Open Error, see Return Value for @error and @extended for @extended
;                  |2   - WinHTTP Send Error, see Return Value for @error and @extended for @extended
;                  |400 - Bad Request (Finesse)
;                  |401 - Unauthorized, See Returned Data (Finesse)
;                  |404 - Not Found (Finesse)
;                  |500 - Internal Server Error (Finesse)
;                  |503 - Service Unavailable (Finesse)
; Author ........: Robert Maehl
; Modified ......: 6/22/2018 9:00 EST
; Remarks .......: Agents can only get a list of their own dialogs. Administrators can get a list of dialogs associated with
;                  any user.
; Related .......:
; Link ..........: https://solutionpartner.cisco.com/media/finesseDevGuide2/CFIN_RF_G7DC228B_00_get-list-of-dialogs-for.html
; Example .......: No
; ===============================================================================================================================
Func _FinesseGetCalls($sAPI, $sUser, $sPassword)

	Local $oHTTP = ObjCreate("WinHttp.WinHttpRequest.5.1")

	$oHTTP.Open("GET", $sAPI & "User/" & $sUser & "/Dialogs", False)
	If @error Then Return SetError(1, @extended, @error)

	$oHTTP.SetCredentials($sUser, $sPassword, 0)

	$oHTTP.Send()
	If @error Then Return SetError(2, 0, 0)

	If $oHTTP.Status <> $FINESSE_STATUS_SUCCESS Then Return SetError($oHTTP.Status, 0, $oHTTP.ResponseText)

	Return SetError(0, 0, $oHTTP.ResponseText)

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _FinesseGetReason
; Description ...: This UDF allows an agent or supervisor to get an individual Not Ready or Sign Out reason code, which is
;                  already defined and stored in the Finesse database (and that is applicable to the agent or supervisor). Users
;                  can select the reason code to display on their desktops when they change their state to NOT_READY or LOGOUT.
; Syntax ........: _FinesseGetReason($sAPI, $sState, $sUser, $sPassword)
; Parameters ....: $sAPI                - Finesse API URL
;                  $iID                 - ID of the Reason
;                  $sUser               - The ID of the user
;                  $sPassword           - The Password of the user
; Return values .: Success - Returns Response from Finesse
;                  Failure - Sets @error:
;                  |1   - WinHTTP Open Error, see Return Value for @error and @extended for @extended
;                  |2   - WinHTTP Send Error, see Return Value for @error and @extended for @extended
;                  |400 - Bad Request (Finesse)
;                  |401 - Unauthorized, See Returned Data (Finesse)
;                  |404 - Not Found (Finesse)
;                  |500 - Internal Server Error (Finesse)
;                  |503 - Service Unavailable (Finesse)
; Author ........: Robert Maehl
; Modified ......: 6/22/2018 9:00 EST
; Remarks .......: Administrators, agents, and supervisors can use this API. The reason code must be global (forAll parameter
;                  set to true) or be assigned to a team to which the user belongs. Only an administrator can get another user's
;                  reason codes.
; Related .......:
; Link ..........: https://solutionpartner.cisco.com/media/finesseDevGuide2/CFIN_RF_GF71D968_00_get-reason-code-user.html
; Example .......: No
; ===============================================================================================================================
Func _FinesseGetReason($sAPI, $iID, $sUser, $sPassword)

	Local $oHTTP = ObjCreate("WinHttp.WinHttpRequest.5.1")

	$oHTTP.Open("GET", $sAPI & "User/" & $sUser & "/ReasonCode/" & $iID, False)
	If @error Then Return SetError(1, @extended, @error)

	$oHTTP.SetCredentials($sUser, $sPassword, 0)

	$oHTTP.Send()
	If @error Then Return SetError(2, 0, 0)

	If $oHTTP.Status <> $FINESSE_STATUS_SUCCESS Then Return SetError($oHTTP.Status, 0, $oHTTP.ResponseText)

	Return SetError(0, 0, $oHTTP.ResponseText)

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _FinesseGetReasonList
; Description ...: This UDF allows an agent or supervisor to get a list of Not Ready or Sign Out reason codes (that are
;                  applicable to that agent or supervisor), which are defined and stored in the Finesse database. Users can
;                  assign one of the reason codes on the desktop when they change their state to NOT_READY or LOGOUT.
; Syntax ........: _FinesseGetReasonList($sAPI, $sState, $sUser, $sPassword)
; Parameters ....: $sAPI                - Finesse API URL
;                  $sState              - State to get Reason List for (NOT_READY, LOGOFF)
;                  $sUser               - The ID of the user
;                  $sPassword           - The Password of the user
; Return values .: Success - Returns Response from Finesse
;                  Failure - Sets @error:
;                  |1   - WinHTTP Open Error, see Return Value for @error and @extended for @extended
;                  |2   - WinHTTP Send Error, see Return Value for @error and @extended for @extended
;                  |400 - Bad Request (Finesse)
;                  |401 - Unauthorized, See Returned Data (Finesse)
;                  |404 - Not Found (Finesse)
;                  |500 - Internal Server Error (Finesse)
;                  |503 - Service Unavailable (Finesse)
; Author ........: Robert Maehl
; Modified ......: 6/22/2018 9:00 EST
; Remarks .......: Administrators, agents and supervisors can use this API. Only an administrator can get another user's list
;                  of reason codes.
; Related .......:
; Link ..........: https://solutionpartner.cisco.com/media/finesseDevGuide2/CFIN_RF_G0C330E7_00_get-reason-code-list-user.html
; Example .......: No
; ===============================================================================================================================
Func _FinesseGetReasonList($sAPI, $sState, $sUser, $sPassword)

	Local $oHTTP = ObjCreate("WinHttp.WinHttpRequest.5.1")

	$oHTTP.Open("GET", $sAPI & "User/" & $sUser & "/ReasonCodes?category=" & $sState, False)
	If @error Then Return SetError(1, @extended, @error)

	$oHTTP.SetCredentials($sUser, $sPassword, 0)

	$oHTTP.Send()
	If @error Then Return SetError(2, 0, 0)

	If $oHTTP.Status <> $FINESSE_STATUS_SUCCESS Then Return SetError($oHTTP.Status, 0, $oHTTP.ResponseText)

	Return SetError(0, 0, $oHTTP.ResponseText)

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _FinesseGetState
; Description ...: This UDF allows a user to get a copy of the User object's State
; Syntax ........: _FinesseGetState($sAPI, $sUser, $sPassword)
; Parameters ....: $sAPI                - Finesse API URL
;                  $sUser               - The ID of the user
;                  $sPassword           - The Password of the user
; Return values .: Success - Returns Response from Finesse
;                  Failure - Sets @error:
;                  |1   - WinHTTP Open Error, see Return Value for @error and @extended for @extended
;                  |2   - WinHTTP Send Error, see Return Value for @error and @extended for @extended
;                  |400 - Bad Request (Finesse)
;                  |401 - Unauthorized, See Returned Data (Finesse)
;                  |404 - Not Found (Finesse)
;                  |500 - Internal Server Error (Finesse)
;                  |503 - Service Unavailable (Finesse)
; Author ........: Robert Maehl
; Modified ......: 6/22/2018 9:00 EST
; Remarks .......: Agents can only get their own User object. Administrators can get any User object.
; Related .......:
; Link ..........: https://solutionpartner.cisco.com/media/finesseDevGuide2/CFIN_RF_G80CB92B_00_get-user.html
; Example .......: No
; ===============================================================================================================================
Func _FinesseGetState($sAPI, $sUser, $sPassword)

	Local $oHTTP = ObjCreate("WinHttp.WinHttpRequest.5.1")

	$oHTTP.Open("GET", $sAPI & "User/" & $sUser, False)
	If @error Then Return SetError(1, @extended, @error)

	$oHTTP.SetCredentials($sUser, $sPassword, 0)

	$oHTTP.Send()
	If @error Then Return SetError(2, 0, 0)

	If $oHTTP.Status <> $FINESSE_STATUS_SUCCESS Then Return SetError($oHTTP.Status, 0, $oHTTP.ResponseText)

	$sState = $oHTTP.ResponseText

	$sState = StringReplace($sState, @CR, "") ; Trim Carriage Returns
	$sState = StringReplace($sState, @LF, "") ; Trim Line Feeds
	$sState = _StringBetween($sState, "<state>", "</state>")[0]

	Return SetError(0, 0, $sState)

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _FinesseGetUser
; Description ...: This UDF allows a user to get a copy of the User object. For a mobile agent, this operation returns the full
;                  User object, including the mobile agent node.
; Syntax ........: _FinesseGetUser($sAPI, $sUser, $sPassword)
; Parameters ....: $sAPI                - Finesse API URL
;                  $sUser               - The ID of the user
;                  $sPassword           - The Password of the user
; Return values .: Success - Returns Response from Finesse
;                  Failure - Sets @error:
;                  |1   - WinHTTP Open Error, see Return Value for @error and @extended for @extended
;                  |2   - WinHTTP Send Error, see Return Value for @error and @extended for @extended
;                  |400 - Bad Request (Finesse)
;                  |401 - Unauthorized, See Returned Data (Finesse)
;                  |404 - Not Found (Finesse)
;                  |500 - Internal Server Error (Finesse)
;                  |503 - Service Unavailable (Finesse)
; Author ........: Robert Maehl
; Modified ......: 6/22/2018 9:00 EST
; Remarks .......: Agents can only get their own User object. Administrators can get any User object.
; Related .......:
; Link ..........: https://solutionpartner.cisco.com/media/finesseDevGuide2/CFIN_RF_G80CB92B_00_get-user.html
; Example .......: No
; ===============================================================================================================================
Func _FinesseGetUser($sAPI, $sUser, $sPassword)

	Local $oHTTP = ObjCreate("WinHttp.WinHttpRequest.5.1")

	$oHTTP.Open("GET", $sAPI & "User/" & $sUser, False)
	If @error Then Return SetError(1, @extended, @error)

	$oHTTP.SetCredentials($sUser, $sPassword, 0)

	$oHTTP.Send()
	If @error Then Return SetError(2, 0, 0)

	If $oHTTP.Status <> $FINESSE_STATUS_SUCCESS Then Return SetError($oHTTP.Status, 0, $oHTTP.ResponseText)

	Return SetError(0, 0, $oHTTP.ResponseText)

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _FinesseGetUsers
; Description ...: This API allows an administrator to get a list of users.
; Syntax ........: _FinesseGetUsers($sAPI, $sUser, $sPassword)
; Parameters ....: $sAPI                - Finesse API URL
;                  $sUser               - The ID of the user
;                  $sPassword           - The Password of the user
; Return values .: Success - Returns Response from Finesse
;                  Failure - Sets @error:
;                  |1   - WinHTTP Open Error, see Return Value for @error and @extended for @extended
;                  |2   - WinHTTP Send Error, see Return Value for @error and @extended for @extended
;                  |400 - Bad Request (Finesse)
;                  |401 - Unauthorized, See Returned Data (Finesse)
;                  |404 - Not Found (Finesse)
;                  |500 - Internal Server Error (Finesse)
;                  |503 - Service Unavailable (Finesse)
; Author ........: Robert Maehl
; Modified ......: 6/22/2018 9:00 EST
; Remarks .......: Only administrators can get a list of users.
; Related .......:
; Link ..........: https://solutionpartner.cisco.com/media/finesseDevGuide2/CFIN_RF_G94B8B76_00_get-list.html
; Example .......: No
; ===============================================================================================================================
Func _FinesseGetUsers($sAPI, $sUser, $sPassword)

	Local $oHTTP = ObjCreate("WinHttp.WinHttpRequest.5.1")

	$oHTTP.Open("GET", $sAPI & "Users", False)
	If @error Then Return SetError(1, @extended, @error)

	$oHTTP.SetCredentials($sUser, $sPassword, 0)

	$oHTTP.Send()
	If @error Then Return SetError(2, 0, 0)

	If $oHTTP.Status <> $FINESSE_STATUS_SUCCESS Then Return SetError($oHTTP.Status, 0, $oHTTP.ResponseText)

	Return SetError(0, 0, $oHTTP.ResponseText)

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _FinesseGetWrapUpReason
; Description ...: This UDF allows a user to get a WrapUpReason object.
; Syntax ........: _FinesseGetWrapUpReason($sAPI, $iID, $sUser, $sPassword)
; Parameters ....: $sAPI                - Finesse API URL
;                  $iID                 - ID of Wrap of Reason
;                  $sUser               - The ID of the user
;                  $sPassword           - The Password of the user
; Return values .: Success - Returns Response from Finesse
;                  Failure - Sets @error:
;                  |1   - WinHTTP Open Error, see Return Value for @error and @extended for @extended
;                  |2   - WinHTTP Send Error, see Return Value for @error and @extended for @extended
;                  |400 - Bad Request (Finesse)
;                  |401 - Unauthorized, See Returned Data (Finesse)
;                  |404 - Not Found (Finesse)
;                  |500 - Internal Server Error (Finesse)
;                  |503 - Service Unavailable (Finesse)
; Author ........: Robert Maehl
; Modified ......: 6/22/2018 9:00 EST
; Remarks .......: Administrators, agents, and supervisors can use this API. Only an administrator can get another user's
;                  wrap-up reasons.
; Related .......:
; Link ..........: https://solutionpartner.cisco.com/media/finesseDevGuide2/CFIN_RF_G6FB32C5_00_get-wrap-up-reason-user.html
; Example .......: No
; ===============================================================================================================================
Func _FinesseGetWrapUpReason($sAPI, $iID, $sUser, $sPassword)

	Local $oHTTP = ObjCreate("WinHttp.WinHttpRequest.5.1")

	$oHTTP.Open("GET", $sAPI & "User/" & $sUser & "/WrapUpReason/" & $iID, False)
	If @error Then Return SetError(1, @extended, @error)

	$oHTTP.SetCredentials($sUser, $sPassword, 0)

	$oHTTP.Send()
	If @error Then Return SetError(2, 0, 0)

	If $oHTTP.Status <> $FINESSE_STATUS_SUCCESS Then Return SetError($oHTTP.Status, 0, $oHTTP.ResponseText)

	Return SetError(0, 0, $oHTTP.ResponseText)

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _FinesseGetWrapUpReasonList
; Description ...: This UDF allows a user to get a list of all wrap-up reasons applicable for that user.
; Syntax ........: _FinesseGetWrapUpReasonList($sAPI, $sUser, $sPassword)
; Parameters ....: $sAPI                - Finesse API URL
;                  $sUser               - The ID of the user
;                  $sPassword           - The Password of the user
; Return values .: Success - Returns Response from Finesse
;                  Failure - Sets @error:
;                  |1   - WinHTTP Open Error, see Return Value for @error and @extended for @extended
;                  |2   - WinHTTP Send Error, see Return Value for @error and @extended for @extended
;                  |400 - Bad Request (Finesse)
;                  |401 - Unauthorized, See Returned Data (Finesse)
;                  |404 - Not Found (Finesse)
;                  |500 - Internal Server Error (Finesse)
;                  |503 - Service Unavailable (Finesse)
; Author ........: Robert Maehl
; Modified ......: 6/22/2018 9:00 EST
; Remarks .......: Administrators, agents and supervisors can use this API. Only an administrator can get another user's list of
;                  wrap-up reasons.
; Related .......:
; Link ..........: https://solutionpartner.cisco.com/media/finesseDevGuide2/CFIN_RF_G3DA2423_00_get-wrap-up-reason-list.html
; Example .......: No
; ===============================================================================================================================
Func _FinesseGetWrapUpReasonList($sAPI, $sUser, $sPassword)

	Local $oHTTP = ObjCreate("WinHttp.WinHttpRequest.5.1")

	$oHTTP.Open("GET", $sAPI & "User/" & $sUser & "/WrapUpReasons", False)
	If @error Then Return SetError(1, @extended, @error)

	$oHTTP.SetCredentials($sUser, $sPassword, 0)

	$oHTTP.Send()
	If @error Then Return SetError(2, 0, 0)

	If $oHTTP.Status <> $FINESSE_STATUS_SUCCESS Then Return SetError($oHTTP.Status, 0, $oHTTP.ResponseText)

	Return SetError(0, 0, $oHTTP.ResponseText)

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _FinesseLogin
; Description ...: This UDF allows a user to sign in to the CTI server. If the response is successful, the user is signed in
;                  to Finesse and is automatically placed in NOT_READY state.
; Syntax ........: _FinesseLogin($sAPI, $sUser, $sPassword, $iExtension)
; Parameters ....: $sAPI                - Finesse API URL
;                  $sUser               - The ID of the user
;                  $sPassword           - The Password of the user
;                  $iExtension          - The extension with which the user wants to sign in
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
; Modified ......: 6/22/2018 9:00 EST
; Remarks .......: Users can only act on their own User objects. If five consecutive sign-ins fail due to an incorrect password,
;                  Finesse blocks access to the user account for a period of 5 minutes.
; Related .......:
; Link ..........: https://solutionpartner.cisco.com/media/finesseDevGuide2/CFIN_RF_S21A5A70_00_sign-in.html
; Example .......: No
; ===============================================================================================================================
Func _FinesseLogin($sAPI, $sUser, $sPassword, $iExtension)

	Local $oHTTP = ObjCreate("WinHttp.WinHttpRequest.5.1")

	$oHTTP.Open("PUT", $sAPI & "User/" & $sUser, False)
	If @error Then Return SetError(1, @extended, @error)

	$oHTTP.SetCredentials($sUser, $sPassword, 0)

	$oHTTP.SetRequestHeader("Content-Type", "application/xml")

	$oHTTP.Send("<User><state>LOGIN</state><extension>" & $iExtension & "</extension></User>")
	If @error Then Return SetError(2, 0, 0)

	If $oHTTP.Status <> $FINESSE_STATUS_ACCEPTED Then Return SetError($oHTTP.Status, 0, $oHTTP.ResponseText)

	Return SetError(0, 0, True)

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _FinesseLogout
; Description ...: This UDF allows a user to sign out of Finesse.
; Syntax ........: _FinesseLogout($sAPI, $sUser, $sPassword)
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
; Modified ......: 6/22/2018 9:00 EST
; Remarks .......: Agents and supervisors can use this API. Users can only act on their own User objects.
; Related .......:
; Link ..........: https://solutionpartner.cisco.com/media/finesseDevGuide2/CFIN_RF_SC0D86B5_00_sign-out-of-finesse.html
; Example .......: No
; ===============================================================================================================================
Func _FinesseLogout($sAPI, $sUser, $sPassword)

	Local $oHTTP = ObjCreate("WinHttp.WinHttpRequest.5.1")

	$oHTTP.Open("PUT", $sAPI & "User/" & $sUser, False)
	If @error Then Return SetError(1, @extended, @error)

	$oHTTP.SetCredentials($sUser, $sPassword, 0)

	$oHTTP.SetRequestHeader("Content-Type", "application/xml")

	$oHTTP.Send("<User><state>LOGOUT</state></User>")
	If @error Then Return SetError(2, 0, 0)

	If $oHTTP.Status <> $FINESSE_STATUS_ACCEPTED Then Return SetError($oHTTP.Status, 0, $oHTTP.ResponseText)

	Return SetError(0, 0, True)

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _FinesseMobileLogin
; Description ...: This UDF allows a user to sign in to the CTI server as a mobile agent. This API uses the existing User
;                  object with a LOGIN state only. The user must be authenticated to use this API successfully.
; Syntax ........: _FinesseMobileLogin($sAPI, $sUser, $sPassword, $iExtension, $sMode, $iForward)
; Parameters ....: $sAPI                - Finesse API URL
;                  $sUser               - The ID of the user
;                  $sPassword           - The Password of the user
;                  $iExtension          - The extension with which to sign in the user
;                  $iMode               - The connection mode for the call
;                  $iForward            - The phone number that the system calls to connect with the mobile agent
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
; Modified ......: 6/22/2018 9:00 EST
; Remarks .......: Users can only act on their own User objects. If five consecutive sign-ins fail due to an incorrect password,
;                  Finesse blocks access to the user account for a period of 5 minutes.
; Related .......:
; Link ..........: https://solutionpartner.cisco.com/media/finesseDevGuide2/CFIN_RF_S2C0ED2A_00_sign-in-as-mobile-agent.html
; Example .......: No
; ===============================================================================================================================
Func _FinesseMobileLogin($sAPI, $sUser, $sPassword, $iExtension, $sMode, $iForward)

	Local $oHTTP = ObjCreate("WinHttp.WinHttpRequest.5.1")

	$oHTTP.Open("PUT", $sAPI & "User/" & $sUser, False)
	If @error Then Return SetError(1, @extended, @error)

	$oHTTP.SetCredentials($sUser, $sPassword, 0)

	$oHTTP.SetRequestHeader("Content-Type", "application/xml")

	$oHTTP.Send("<User><state>LOGIN</state><extension>" & $iExtension & "</extension><mobileAgent><mode>" & $sMode & "</mobileAgent></User>")
	If @error Then Return SetError(2, 0, 0)

	If $oHTTP.Status <> $FINESSE_STATUS_ACCEPTED Then Return SetError($oHTTP.Status, 0, $oHTTP.ResponseText)

	Return SetError(0, 0, True)

EndFunc

; #FUNCTION# ====================================================================================================================
; Name ..........: _FinesseSetState
; Description ...: This UDF allows a user to change the state of an agent on the CTI server. Agents can change their own states.
;                  Additionally, when changing state to NOT_READY or LOGOUT, this UDF allows a user to change the agent state
;                  in the CTI server and pass along the code value of a corresponding reason code.
; Syntax ........: _FinesseSetState($sAPI, $sState, $iReason, $sUser, $sPassword)
; Parameters ....: $sAPI                - Finesse API URL
;                  $sState              - The new state the user wants to be in (NOT_READY, LOGOUT)
;                  $iReason             - The database ID for the reason code
;                  $sUser               - The ID of the user
;                  $sPassword           - The Password of the user
; Return values .: Success - Returns True
;                  Failure - Sets @error:
;                  |1   - WinHTTP Open Error, see Return Value for @error and @extended for @extended
;                  |2   - WinHTTP Send Error, see Return Value for @error and @extended for @extended
;                  |400 - Bad Request (Finesse)
;                  |401 - Unauthorized, See Returned Data (Finesse)
;                  |404 - Not Found (Finesse)
;                  |500 - Internal Server Error (Finesse)
;                  |503 - Service Unavailable (Finesse)
; Author ........: Robert Maehl
; Modified ......: 6/22/2018 16:00 EST
; Remarks .......: Agents can only act on their own User objects. Supervisors can act on the User objects of agents who belong
;                  to their team.
; Related .......:
; Link ..........: https://solutionpartner.cisco.com/media/finesseDevGuide2/CFIN_RF_C1D2CCD7_00_change-agent-state.html
;                  https://solutionpartner.cisco.com/media/finesseDevGuide2/CFIN_RF_C331CBFD_00_change-agent-state-with-reason.html
; Example .......: No
; ===============================================================================================================================
Func _FinesseSetState($sAPI, $sState, $iReason, $sUser, $sPassword)

	Local $oHTTP = ObjCreate("WinHttp.WinHttpRequest.5.1")

	$oHTTP.Open("PUT", $sAPI & "User/" & $sUser, False)
	If @error Then Return SetError(1, @extended, @error)

	$oHTTP.SetCredentials($sUser, $sPassword, 0)

	$oHTTP.SetRequestHeader("Content-Type", "application/xml")

	If $iReason = "" Then
		$oHTTP.Send("<User><state>" & $sState & "</state></User>")
	Else
		$oHTTP.Send("<User><state>" & $sState & "</state><reasonCodeId>" & $iReason & "</reasonCodeId></User>")
	EndIf
	If @error Then Return SetError(2, 0, 0)

	If $oHTTP.Status = $FINESSE_STATUS_SUCCESS Then
		;;;
	ElseIf $oHTTP.Status = $FINESSE_STATUS_ACCEPTED Then
		;;;
	Else
		Return SetError($oHTTP.Status, 0, $oHTTP.ResponseText)
	EndIf

	Return SetError(0, 0, True)

EndFunc