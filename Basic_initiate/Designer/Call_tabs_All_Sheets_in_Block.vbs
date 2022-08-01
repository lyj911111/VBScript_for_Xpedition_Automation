Option Explicit

Dim vdapp
Set vdapp = Application

Scripting.AddTypeLibrary("ViewDraw.Application")

Dim vdview, vddoc, vdblock, sVdBlock
Set vdview = vdapp.ActiveView	
Set vddoc = vdapp.ActiveDocument
sVdBlock = vdview.GetName(Short_Name)

' activeView, activeDoc, Block 까지 기본적으로 불러오기

msgbox sVdBlock


' 특정 하나의 Sheet를 띄울때는 아래 코드 2줄 사용.
' ' vdapp.pushpath [block이름], ,[Sheet이름]
' vdapp.pushpath sVdBlock,"", "microprocessor"



' 모든 Sheet를 다 띄우기
' 현재 Block(Schematic)에 있는 모든 Sheet를 collection 형태로 가져옴
Dim colSheets
Set colSheets = vdapp.SchematicSheetDocuments.GetAvailableSheets(sVdBlock)

' Collection에 있는 Sheet 하나씩 띄우기
Dim oEachSheet
For Each oEachSheet in colSheets
	vdapp.pushpath sVdBlock, "", oEachSheet
	' 결과를 보여주는 탭 생성
	vdapp.AppendOutput "Poped up Sheet Lists", oEachSheet
Next
