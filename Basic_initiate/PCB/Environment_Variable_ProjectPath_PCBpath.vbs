
option Explicit
Dim pcbApp
Set pcbApp = Application
' Set pcbApp = GetObject (,"MGCPCB.Application")
Dim pcbDoc
Set pcbDoc = pcbApp.ActiveDocument

If (ValidateServer(pcbDoc) = 1) Then
   Call Scripting.AddTypeLibrary("MGCPCB.ExpeditionPCBApplication")
   Call Scripting.AddTypeLibrary("Scripting.FileSystemObject")
    ' 여기에 코드를 작성 시작!

	' 작업환경변수 가져옴
	Dim WDIR_env
	WDIR_env = Scripting.GetEnvVariable ("WDIR") ' Returns WDIR string

    ' 현재 실행되고 있는 PCB의 경로 가져옴 ( .pcb 파일 경로)
    Dim ExpDocLoc
    ExpDocLoc = pcbdoc.Path  

    ' 이 PCB가 참조하고 있는 .prj (포로젝트 파일 DxDesigner) 경로 가져옴
    Dim DxDesProj 
    DxDesProj = pcbDoc.Environment.ProjectFileFullName




    ' 에러가 나도 무시하고 진행
    ' On Error Resume Next
    ' 디버깅용 결과창 출력 또는 유저에게 알림
    ' AppendOutput [탭이름], [Line 한줄 쓸 text]
    AppendOutput "DEBUG", "------------------------------------------------"
    AppendOutput "DEBUG", WDIR_env
    AppendOutput "DEBUG", ExpDocLoc
    AppendOutput "DEBUG", DxDesProj
    ' AppendOutput "DEBUG", "Num of Plane Shapes: " & pcbDoc.PlaneShapes.Count
    AppendOutput "DEBUG", "------------------------------------------------"
    ' On Error GoTo 0


Else
    Msgbox("Could not validate the server. Exiting program.")
End If


'=========================================================================
' Message Window Output
'=========================================================================
Function AppendOutput(sOutputTab, str)
    Dim mnu, OutputControl, objTab
    Set mnu = Gui.CommandBars("Document Menu Bar").Controls("&View").Controls("Message Window")
    If mnu.Checked = False Then
        Call Gui.ProcessCommand(33125)
    End If
    Set OutputControl = Addins.Item("Message Window").Control
    Set objTab = OutputControl.AddTab(sOutputTab)
    Call objTab.Activate
    Addins("Message Window").Control.AddTab(sOutputTab).AppendText (str & vbCrLf)
End Function
Function ClearOutputWindow(sOutputTab)
    Addins("Message Window").Control.AddTab(sOutputTab).Clear
End Function

'---------------------------------------
' Begin Validate Server Function
'---------------------------------------
Private Function ValidateServer(doc)
    
    Dim key, licenseServer, licenseToken
    ' Ask Expedition document for the key
    key = doc.Validate(0)
    ' Get license server
    Set licenseServer = CreateObject("MGCPCBAutomationLicensing.Application")
    ' Ask the license server for the license token
    licenseToken = licenseServer.GetToken(key)
    ' Release license server
    Set licenseServer = nothing
    ' Turn off error messages.  Validate may fail if the token is incorrect
    On Error Resume Next
    Err.Clear
    ' Ask the document to validate the license token
    doc.Validate(licenseToken)
    If Err Then
        ValidateServer = 0    
    Else
        ValidateServer = 1
    End If
End Function
'---------------------------------------
' End Validate Server Function
