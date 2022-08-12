' 첫번째 arg : mgcscript
' 두번째 arg : xxx.vbs 스크립트 파일
' ===== 위까지는 스크립트 구동을 위한 arg ====
' 세번째 arg : <path>\xxx.pcb   <= 실행할 .pcb 파일
' 네번째 arg : parameter
' 다섯~ 여섯~~ 등등  나머지는 parameter


Dim args, pcbApp, pcbDoc, jobName
Set args = ScriptHelper.Arguments

' xpedition app 호출
Set pcbApp = GetObject(,"MGCPCB.ExpeditionPCBApplication")
Set pcbDoc = GetLicensedDoc(pcbApp)


' 명령 : mgcscript [xxx.vbs 스크립트 파일경로] [xxx.pcb 파일경로] <= item(3)는 .pcb를 의미
' 유저가 쓴 arg가 표시됨 (3번째인자)
Arg_jobName = args.item(3)
Arg_jobName = LCase(Arg_jobName)
msgbox "UserGivenName: " & Arg_jobName



' jobName은 System에서 불러온 xxx.pcb의 Full 절대경로가 나옴 (필요하면 split으로 잘라서 쓰면 됨)
jobName = pcbDoc.FullName
jobName = LCase(jobName)
msgbox "System Call :" & jobName


if Arg_jobName = jobName then
    msgbox "You correctly called"
Else
    msgbox "That .pcb file is not existed"
End If


Public Function GetLicensedDoc(appObj)
    On Error Resume Next
    Dim key, licenseServer, licenseToken, docObj
    Set GetLicensedDoc = Nothing
    ' collect the active document
    Set docObj = appObj.ActiveDocument
    If Err Then
        Call appObj.Gui.StatusBarText("No active document: " & Err.Description, epcbStatusFieldError)
        Exit Function
    End If
    ' Ask Expedition�s document for the key
    key = docObj.Validate(0)
    ' Get token from license server
    Set licenseServer = CreateObject("MGCPCBAutomationLicensing.Application." & sComVersion)
    licenseToken = licenseServer.GetToken(key)
    Set licenseServer = Nothing
    ' Ask the document to validate the license token
    Err.Clear
    Call docObj.Validate(licenseToken)
    If Err Then
        Call appObj.Gui.StatusBarText("No active document license: " & Err.Description, epcbStatusFieldError)
        Exit Function
    End If
    ' everything is OK, return document
    Set GetLicensedDoc = docObj
End Function