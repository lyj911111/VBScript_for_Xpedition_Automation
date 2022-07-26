Dim pcbApp
Set pcbApp = Application

Dim pcbDoc
Set pcbDoc = pcbApp.ActiveDocument

If (ValidateServer(pcbDoc) = 1) Then
    Scripting.AddTypeLibrary("MGCPCB.ExpeditionPCBApplication")
    Scripting.AddTypeLibrary("Scripting.FileSystemObject")


    '여기에 코드를 작성
    Set fso = CreateObject("Scripting.FileSystemObject")
    ' 현재경로에 파일 생성
    Set txtFile = fso.CreateTextFile(".\BOMList.csv", true)

    Set CompsColl = pcbDoc.Components
    CompCount = CompsColl.count

    ' component 갯수 확인
    ' msgbox CompCount

    ' * Object에는 직접 값 대입이 불가능 함으로, 임시변수를 선언했음
    dim sValueProperty

    txtFile.Writeline("Part Name" & "," & "Value" & "," & "CompName" & "," & "CompLocX" & "," & "CompLocY" & "," & "CompSide")
    For Each Comp in CompsColl
        PartName = Comp.PartName
        CompName = Comp.Name
        Set Valueatt = Comp.FindProperty("Value")  'Object이므로 writeline에는 Valueatt.Name으로 써야 함.

    ' * Object에 값이 있다면 임시로 만든 변수 sValueProperty 를 그대로 사용하고, 그렇지 않다면 공백으로 넣음.
        If Valueatt Is Nothing Then
            sValueProperty = " "
        Else
            sValueProperty = Valueatt.value
        End If
        CompLocX = Comp.CenterX
        CompLocY = Comp.CenterY
        CompSide = Comp.side
        'Set MakerPN = Comp.FindProperty("Part Name")

        ' 값 쓰기
        txtFile.Writeline(PartName & "," & sValueProperty & "," & CompName & "," & CompLocX & "," & CompLocY & "," & CompSide)

    Next
    txtFile.close

    MsgBox("Work Well !")


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


