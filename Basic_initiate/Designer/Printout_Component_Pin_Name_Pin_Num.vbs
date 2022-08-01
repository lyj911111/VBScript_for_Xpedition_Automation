Option Explicit

' 기본
Dim vdapp, vdview, vddoc
Set vdapp = Application
Set vdview = vdapp.ActiveView
Set vddoc = vdapp.ActiveDocument

' 내장된 라이브러리 가져옴 (각종 예약어들)
Scripting.AddTypeLibrary("ViewDraw.Application")

Dim objComp, objComps, objSymBlk, iSymType
Dim Conn, ConnPin, ConnPinLabel

' 각종 object를 가져옴. 유저가 선택한 항목에 대해서
For Each objComp In vdview.query(VDM_COMP, VD_SELECTED)

	' Symbol관련 Block object를 가져옴
    Set objSymBlk = objComp.SymbolBlock

	' Symbol의 타입이 뭔지? 
' 	VDB_ANNOTATE : 

    iSymType = objSymBlk.SymbolType
	msgbox iSymType
	
	' 결과창에 출력
    Call vdapp.AppendOutput("test", "Ref : " & objComp.RefDes)

	' 선택항목중 Pin이 아니고, Annotate (페이지간 연결 포트 심볼) 이 아닐 경우에만.
    ' 
    If ((iSymType <> VDB_ANNOTATE) And (iSymType <> VDB_PIN)) Then
        ' 선택한 항목의 핀 이름, 맵핑된 핀 번호 Output창에 출력
        For Each Conn In objComp.GetConnections
            Set ConnPin = Conn.CompPin
            Set ConnPinLabel = ConnPin.Pin.Label
            Call vdapp.AppendOutput("test", "Pin Name : " & ConnPinLabel.TextString & ", Pin Number : " & ConnPin.Number)
        Next
    End If
Next
