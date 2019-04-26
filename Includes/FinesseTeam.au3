#include "./FinesseConstants.au3"

; #FUNCTION# ====================================================================================================================
; Name ..........: _FinesseGetTeam
; Description ...: This UDF allows a user to get a copy of the Team object. The Team object contains the configuration
;                  information for a specific team, which includes the URI, the team ID, the team name, and a list of agents
;                  who are members of that team.
; Syntax ........: _FinesseGetTeam($sAPI, $iFrom, $iTo, $sUser, $sPassword)
; Parameters ....: $sAPI                - Finesse API URL
;                  $iID                 - Team ID
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
; Link ..........: https://solutionpartner.cisco.com/media/finesseDevGuide2/CFIN_RF_T846AB1E_00_team-get-team.html
; Example .......: No
; ===============================================================================================================================
Func _FinesseGetTeam($sAPI, $iID, $sUser, $sPassword)

	Local $oHTTP = ObjCreate("WinHttp.WinHttpRequest.5.1")

	$oHTTP.Open("GET", $sAPI & "Team/" & $iID, False)
	If @error Then Return SetError(1, @extended, @error)

	$oHTTP.SetCredentials($sUser, $sPassword, 0)

	$oHTTP.SetRequestHeader("Content-Type", "application/xml")

	$oHTTP.Send()
	If @error Then Return SetError(2, 0, 0)

	If $oHTTP.Status <> $FINESSE_STATUS_SUCCESS Then Return SetError($oHTTP.Status, 0, $oHTTP.ResponseText)

	Return SetError(0, 0, True)

EndFunc