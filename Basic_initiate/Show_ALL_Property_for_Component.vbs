'
' Component 하나가 갖고 있는 모든 Property 항목을 출력해봄
'
' Naming Convention
' o 로시작 : object
' col로시작 : collection
' s로시작 : String
'
' 이 코드를 실행하면 comp_prop.txt 파일을 만들고 아래 항목이 모두 출력
' 1개의 component에 대해 사용가능한 모든 property를 출력해줌
'
' Part Number 113-RES
' $$_RFParametric_RF_REGION 
' $$Internal_UUID 396-2-51
' $$NotCommonProperty_RFParametric COMP=Fixed;RF_REGION=
' SSheet CORPORATE(PCI_Connection)
' $$_RFParametric_COMP Fixed
' $$NotCommonProperty_PARTS 1
' $$Internal_Resolution 00001mm
' $$NotCommonProperty_Power 250
' Type Resistor
' Part Label 113-RES
' $$NotCommonProperty_Primary Supplier ROHM INC.
' Part Name 113-RES
' Tolerance 1%
' Description RES, TF, SMD, 0805, 698K OHMS, 1%
' RFParametric COMP=Fixed;RF_REGION=
' $$NotCommonProperty_PKG_TYPE CC0805
' $$Internal_LastIdUsed 14
' $$Internal___BlkDate 14:42:53_01-13-12
' Value 
' DXDB_LIBNAME Resistor
' CELL_LAST_MODIFIED_BY vwnek@CAS-VWNEK-LT
' Height 508000
' Underside Space 70000
' Cost 0.10
' Cell Name CC0805
' SPath CORPORATE
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
    Set colComps = pcbDoc.Components
    dim oCompObj, propObj, propName

    ' Collection에서 1개의 component 선택 Item(1)
    set oCompObj = colComps.Item(1)

    ' 특정 Refdes에 몇개의 property가 있는지 확인 (갯수 확인해서 msgbox로 알려줌)
    dim colProp
    Set colProp = oCompObj.Properties
    MsgBox "There are " + CStr(colProp.Count) + " properties on " + oCompObj.Name

    ' 담고 있는 모든 Property를 Text파일로 출력해봄
    dim oProp
    For Each oProp In colProp
        txtFile.Writeline(oProp.name & " : " & oProp.value)
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