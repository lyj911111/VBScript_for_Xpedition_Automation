'
' PCB상의 모든 Part에 대해서 Bom List를 만들어줌.
' 속성은 아래에서 결정
'


Dim pcbApp
Set pcbApp = Application

Dim pcbDoc
Set pcbDoc = pcbApp.ActiveDocument

If (ValidateServer(pcbDoc) = 1) Then
    Scripting.AddTypeLibrary("MGCPCB.ExpeditionPCBApplication")
    Scripting.AddTypeLibrary("Scripting.FileSystemObject")


    '===여기에 코드를 작성===

    Set fso = CreateObject("Scripting.FileSystemObject")
    ' 현재경로에 파일 생성
    Set txtFile = fso.CreateTextFile(".\BOMList.csv", true)

    Set CompsColl = pcbDoc.Components
    CompCount = CompsColl.count

    ' component 갯수 확인
    ' msgbox CompCount

    ' 임시 변수 선언 (object를 string출력을 위해)
    dim sValueProperty
    dim sCellNameProperty


    ' => 넣고자 하는 값들 사용
    txtFile.Writeline("Part Name" & "," & "Value" & "," & "CompName" & "," & "CompLocX" & "," & "CompLocY" & "," & "CompSide" & "," & "Cell Name")
    For Each Comp in CompsColl
        PartName = Comp.PartName
        CompName = Comp.Name

        Set oValueatt = Comp.FindProperty("Value")  'Object이므로 writeline에는 oValueatt.Name으로 써야 함.
        ' 값이 없으면 공백을 삽입
        If oValueatt Is Nothing Then
            sValueProperty = " "
            'AppendOutput "debug", oValueatt.value
        Else
            sValueProperty = oValueatt.value
        End If

        CompLocX = Comp.CenterX
        CompLocY = Comp.CenterY
        CompSide = Comp.side

        Set oCellName = Comp.FindProperty("Cell Name")
        ' 값이 없으면 공백을 삽입
        If oCellName Is Nothing Then
            sCellNameProperty = " "
            'AppendOutput "debug", oValueatt.value
        Else
            sCellNameProperty = oCellName.value
        End If
 
        ' 값 쓰기
        txtFile.Writeline(PartName & "," & sValueProperty & "," & CompName & "," & CompLocX & "," & CompLocY & "," & CompSide & "," & sCellNameProperty)

    Next
    txtFile.close

    MsgBox("Work Well !")

    '===여기까지 작성===

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


