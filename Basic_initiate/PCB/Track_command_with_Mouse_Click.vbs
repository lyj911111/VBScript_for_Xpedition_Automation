

' 마우스 버튼으로 명령어 눌렀을 때 해당 명령 id 와 이름이 출력되도록...

Option Explicit

Dim pcbApp
Set pcbApp = Application

Dim pcbDoc
Set pcbDoc = pcbApp.ActiveDocument

If (ValidateServer(pcbDoc) = 1) Then
    Scripting.AddTypeLibrary("MGCPCB.ExpeditionPCBApplication")
    Scripting.AddTypeLibrary("Scripting.FileSystemObject")
	
    Dim cmdListener: Set cmdListener = pcbApp.Gui.CommandListener
    Call Scripting.AttachEvents(cmdListener, "myCommands")

    Scripting.DontExit = TRUE ' 스크립팅이 무한으로돌면서 이벤트 확인

	
Else
    Msgbox("Could not validate the server. Exiting program.")
End If

' 마우스 클릭으로 커맨드 확인 할 수 있는 함수 두개 post, pre
Sub myCommands_PostOnCommand(sCommandName, CommandID)
    Call AppendOutput("test", "PostCmd: " & sCommandName & " " & CommandID)
End Sub

Sub myCommands_PreOnCommand(sCommandName, commandID)
    Call AppendOutput("test", "PreCmd: " & sCommandName & " " & CommandID)
End Sub


' 아래는 마우스 이벤트 함수
'Sub myCommands_PreOnMouseClk(eButton , eFlags ,dX , dY) 
'    Call AppendOutput("test", "PreCmd: " & cmdListener.name & " " & cmdListener.id)
'
'MsgBox cmdListener.name & " | " & cmdListener.id
'End Sub



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

    ' Ask Expedition뭩 document for the key
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