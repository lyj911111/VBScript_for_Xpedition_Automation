'
' xPCB 내의 모든 Component를 하나씩 msgbox로 찍어봄 (for문없이 노가다)
'

Dim pcbApp
Set pcbApp = Application

Dim pcbDoc
Set pcbDoc = pcbApp.ActiveDocument

If (ValidateServer(pcbDoc) = 1) Then
    Scripting.AddTypeLibrary("MGCPCB.ExpeditionPCBApplication")
    Scripting.AddTypeLibrary("Scripting.FileSystemObject")


    ' 여기에 코드를 작성

    ' 출력물을 저장할 파일 생성 (현재경로)
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set txtFile = fso.CreateTextFile(".\comp_prop.txt", true)

    ' component 속성을 불러올 때 쓰이는 변수 선언
    Set CompsColl = pcbDoc.Components
    dim compObj, propObj, propName

    ' C1 의 Attached Properties (Select Mode로 Part를 클릭하면 나오는 속성)
    propName = "Value"
    propName = "Part Number"
    propName = "SSheet"
    propName = "Type"
    propName = "Cell Name"
    propName = "SPath"
    propName = "$$Internal_UUID"

    ' 1개의 component 선택 Item(1)
    set compObj = compsColl.Item(1)

    set propObj0 = compObj.FindProperty("Value")
    set propObj1 = compObj.FindProperty("Part Number")
    set propObj2 = compObj.FindProperty("SSheet")
    set propObj3 = compObj.FindProperty("Type")
    set propObj4 = compObj.FindProperty("Cell Name")
    set propObj5 = compObj.FindProperty("SPath")
    set propObj6 = compObj.FindProperty("$$Internal_UUID")


    ' 선택한 comp에 대해서 라인 쓰기
    txtFile.Writeline("Selected Component RefDes : " & compObj.refdes)
    txtFile.Writeline()
    txtFile.Writeline(propObj0.Name + "/" + propObj0.Value)
    txtFile.Writeline(propObj1.Name + "/" + propObj1.Value)
    txtFile.Writeline(propObj2.Name + "/" + propObj2.Value)
    txtFile.Writeline(propObj3.Name + "/" + propObj3.Value)
    txtFile.Writeline(propObj4.Name + "/" + propObj4.Value)
    txtFile.Writeline(propObj5.Name + "/" + propObj5.Value)
    txtFile.Writeline(propObj6.Name + "/" + propObj6.Value)

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


