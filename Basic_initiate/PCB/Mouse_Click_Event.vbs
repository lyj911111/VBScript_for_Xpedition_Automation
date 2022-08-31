' 마우스 버튼으로 명령어 눌렀을 때 해당 이벤트 발생여부, 좌표 표시

Option Explicit

Dim pcbApp
Set pcbApp = Application

Dim pcbDoc
Set pcbDoc = pcbApp.ActiveDocument

' 마우스 이벤트의 값을 활용하려면 전역변수 선언이 필요!!
Dim xCoord, yCoord, pointCoord

If (ValidateServer(pcbDoc) = 1) Then
    Scripting.AddTypeLibrary("MGCPCB.ExpeditionPCBApplication")
    Scripting.AddTypeLibrary("Scripting.FileSystemObject")
	
    Dim cmdListener
    Set cmdListener = pcbApp.Gui.CommandListener
    ' sub함수 이름은 => " " 이름이어야 함!
    Call Scripting.AttachEvents(cmdListener, "myCommands")

    ' 마우스 이벤트의 값을 활용하려면 전역변수의 값을 써야 함!
    MsgBox xCoord

    Scripting.DontExit = TRUE ' 스크립팅이 무한으로돌면서 이벤트 확인, False면 반복안함

	
Else
    Msgbox("Could not validate the server. Exiting program.")
End If


' PostOnMouseClk : 마우스버튼을 땔때 이벤트 발생
' 아래 myCommands 는 위에서 " "사이에 있는 text임을 인지
' 보통 마우스이벤트의 경우 xxx_PreOnMouseClk() 이런식으로 언더바를 사이에 두고 sub함수를 만듦
Sub myCommands_PostOnMouseClk(eButton , eFlags ,dX , dY) 
    Call AppendOutput("test", "PostOnMouseclk : " & eButton & " , " & eFlags & " , " & dX  & " , " & dY )

    ' x좌표, y좌표를 각각 담음
    xCoord = dX
    yCoord = dY
    ' 2째자리까지 반올림하기
    dX = Round(dX, 2)
    dY = Round(dY, 2)
    ' 두좌표를 배열에 담음, 출력 방법
    pointCoord = array(dX,dY)
    MsgBox pointCoord(0) & "," & pointCoord(1)


End Sub

' PreOnMouseClk : 마우스버튼을 누를때 이벤트 발생
Sub myCommands_PreOnMouseClk(eButton , eFlags ,dX , dY) 
    Call AppendOutput("test", "PreOnMouseclk : " & eButton & " , " & eFlags & " , " & dX  & " , " & dY )
End Sub


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
