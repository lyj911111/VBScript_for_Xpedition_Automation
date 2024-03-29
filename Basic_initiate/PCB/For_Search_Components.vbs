'
' xPCB 내의 모든 Component를 하나씩 msgbox로 찍어봄 (for문)
'

Dim pcbApp
Set pcbApp = Application

Dim pcbDoc
Set pcbDoc = pcbApp.ActiveDocument

If (ValidateServer(pcbDoc) = 1) Then
    Scripting.AddTypeLibrary("MGCPCB.ExpeditionPCBApplication")
    Scripting.AddTypeLibrary("Scripting.FileSystemObject")


    ' 여기에 코드를 작성
    ' 하나씩 넘어갈때마다 다음 refdes가 찍힘
    Set CompsColl = pcbDoc.Components
    dim compObj, propObj, propName
    propName = "MyProperty"

    ' msgbox에 compObj.refdes 가 정석이지만, 그냥 compObj 로 둬도 refdes가 들어감
    For i=1 To 10
        set compObj = compsColl.Item(i)
        msgbox compObj.refdes
    Next

    For i=11 To 20
        set compObj = compsColl.Item(i)
        msgbox compObj
    Next

    
    ' set compObj = compsColl.Item(1)
    ' msgbox compObj
    ' set compObj = compsColl.Item(2)
    ' msgbox compObj
    ' set compObj = compsColl.Item(3)
    ' msgbox compObj
    ' set compObj = compsColl.Item(4)
    ' msgbox compObj
    ' set compObj = compsColl.Item(5)
    ' msgbox compObj


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

