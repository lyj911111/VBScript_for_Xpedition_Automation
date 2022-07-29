Option Explicit 
 
' **** 열려있는 도면에 대한 object access 사용 **** 
' windows OS가 찾아서 줌
' Set vdapp = GetObject (,"ViewDraw.Application") 
' 현재 실행하는 app이 parent로 잡힘 (권장) - 복수의 process가 있을 때  위 방법은 랜덤으로 잡음
Dim vdapp
Set vdapp = Application

' 라이브러리 import용 예약어 getobject()사용시 필요, Application을 호출하면 자체적으로 라이브러리는 호출됨
' Scripting.AddTypeLibrary("ViewDraw.Application") 

Dim vddoc,vdview 
Set vddoc = vdapp.ActiveDocument 
Set vdview = vdapp.ActiveView 

' 값이 다음인 시스템 환경 변수 값을 가져옵니다.
' 데이터시트가 포함된 디렉토리를 찾아 사용자에게 알립니다.
' attributes 선택 파트에 attribute 얻기
Dim DataSheet, DataSheetName, RefDes, RefDesValue, Comp 

' 환경변수에서 값 가져오기 (경로를 불러옴)
Dim DataSheetDir
DataSheetDir = Scripting.GetEnvVariable ("WDIR_EEVX_2_11") 

' 잘 출력되는지 확인
MsgBox "DataSheetDir is " & DataSheetDir,,"Datasheet Directory" 
 

' Query Method를 이용하여 회로도에서 모든 component Collection 만들기
' For Each Comp in vdview.Query(VDM_COMP, VD_ALL)
dim colComps
set colComps = vdview.Query(VDM_COMP, VD_ALL) 


' colComps 컴포넌트 object collection이 잡혀있나 확인차
if colComps Is NOTHING then
	msgbox "Comps are emptied"
Else
	msgbox "Comps are selected"
End if


Dim oAttr, oComp
For Each oComp In colComps
	' 속성(Attribute)를 담는 object
	Set oAttr = oComp.FindAttribute("Part Name")  ' 아무 속성값 넣어서 확인
	' 속성을 출력해보면서 메세지박스로 yes or no 박스로 확인
	' Id 를 출력할 땐 oComp 를 써야 함
	If Not oAttr Is Nothing Then
		If msgbox(oComp.uid & " - " & oAttr.name & " : " & oAttr.Value, vbYesNo, "" ) = vbNo Then
			Exit For
		End If
	End If
Next
	