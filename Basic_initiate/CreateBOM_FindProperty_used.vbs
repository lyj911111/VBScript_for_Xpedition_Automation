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

    txtFile.Writeline("Part Name" & "," & "Value" & "," & "CompName" & "," & "CompLocX" & "," & "CompLocY" & "," & "CompSide")
    For Each Comp in CompsColl
        PartName = Comp.PartName
        CompName = Comp.Name
        Set Valueatt = Comp.FindProperty("Value")  'Object이므로 writeline에는 Valueatt.Name으로 써야 함.
        CompLocX = Comp.CenterX
        CompLocY = Comp.CenterY
        CompSide = Comp.side
        'Set MakerPN = Comp.FindProperty("Part Name")

        ' 값 쓰기
        ' Valueatt는 Object Type으로 값이 없으면 Nothing (ojbect 타입) 이 출력되어서 Msgbox 또는 파일로 출력 불가
        ' 형태를 아래와 같이 String으로 바꿔야 출력이 가능. Value에 값이 있으면 String으로 출력됨
        If Not Valueatt Is Nothing Then
            txtFile.Writeline(PartName & "," & Valueatt.value & "," & CompName & "," & CompLocX & "," & CompLocY & "," & CompSide)
        Else
            txtFile.Writeline(PartName & "," & " " & "," & CompName & "," & CompLocX & "," & CompLocY & "," & CompSide)
        End If


        'txtFile.Writeline(MakerPN.value)

        ' txtFile.Writeline(CompName)
        ' msgbox(CompName)
    Next
    txtFile.close

    MsgBox("Work Well !")


Else
    Msgbox("Could not validate the server. Exiting program.")
End If
 

Call pcbApp.Gui.ProcessCommand("Undo", True)
'Call pcbApp.Gui.ProcessCommand(57643, True)  

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


