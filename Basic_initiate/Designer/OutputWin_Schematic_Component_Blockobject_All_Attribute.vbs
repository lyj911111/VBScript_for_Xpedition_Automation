Option Explicit

' application object는 DxDesigner 그 자체이며 시작점.
' 새로운 인스턴스 생성 or 기존의 인스턴스에 접속
Dim vdapp
Set vdapp = Application

' 라이브러리 active view 소환
Scripting.AddTypeLibrary("ViewDraw.Application")

' 파일 출력을 위한
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

' object 변수 생성
' ActiveView : 현재 열린창의 instance에 대한 object를 반환 화면 확대 축소 이동 등 조작을 위함
'              View object를 반환하여 
' ActiveDoc : application object를 반환하며, 이 Schematic을 close, save 등 전체 통제를 위한 object 반환
' ActiveBlock : 현재 Block(Schematic)의 object를 가져옴
Dim vdview, vddoc, vdblock
Set vdview = vdapp.ActiveView
Set vddoc = vdapp.ActiveDocument
Set vdblock = ActiveView.Block

' Block이 잘 불러왔는지 확인
if vdblock Is NOTHING then
	msgbox "no"
Else
    msgbox "block name is : " & vdblock.GetName(SHORT_NAME)
End if

Dim objProjectData, striCDBDesign, striCDBRootBlock
Set objProjectData = vdapp.GetProjectData

striCDBDesign = objProjectData.GetiCDBDesigns.GetItem(1)
striCDBRootBlock = objProjectData.GetiCDBDesignRootBlock(striCDBDesign)

' board1 출력
MsgBox "Current Board name : " & striCDBDesign
' schematic1 출력
msgbox "Current Schematic Name : " & striCDBRootBlock

' block에 대한 object얻고 이름을 출력
Dim sDesignName, sDesignAlias
sDesignName = vdblock.GetName(SHORT_NAME)
msgbox "Block Name : " & sDesignName

' For Each Comp in vdview.Query(VDM_COMP, VD_ALL)
dim colComps
set colComps = vdview.Query(VDM_COMP, VD_ALL) 



' 하나의 component에 대해서 Block object로 추출할 수 있는 것들...
Dim oComp
' Block object관련 변수들
Dim oBlockRefdes, oBlockPartName, oBlockPartNumber, oBlockCellName, oBlockPartLabel

' Script Output이라는 탭을 만들어서 결과를 출력
Call vdapp.AppendOutput("Script Output", "*** Start to RUN CODE ***")

For Each oComp In colComps
	' 속성(Attribute)를 담는 object
    ' ** Block object이므로 Block Value Refdes가 출력됨 **
	Set oBlockRefdes = oComp.FindAttribute("Ref Designator")  ' 아무 속성값 넣어서 확인
    Set oBlockPartName = oComp.FindAttribute("Part Name")
    Set oBlockPartNumber = oComp.FindAttribute("Part Number")
    Set oBlockCellName = oComp.FindAttribute("Cell Name")
    Set oBlockPartLabel = oComp.FindAttribute("Part Label")

    
    ' 속성을 출력해보면서 메세지박스로 yes or no 박스로 확인
	' Id 를 출력할 땐 oComp 를 써야 함
	If Not oBlockRefdes Is Nothing Then

		' Script Output이라는 탭을 만들어서 결과를 출력
		Call vdapp.AppendOutput("Script Output", "unique id : " & oComp.uid)
		Call vdapp.AppendOutput("Script Output", oBlockRefdes.name & " : " & oBlockRefdes.Value)
		Call vdapp.AppendOutput("Script Output", oBlockPartName.name & " : " & oBlockPartName.Value)
		Call vdapp.AppendOutput("Script Output", oBlockPartNumber.name & " : " & oBlockPartNumber.Value)
		Call vdapp.AppendOutput("Script Output", oBlockCellName.name & " : " & oBlockCellName.Value)
		Call vdapp.AppendOutput("Script Output", oBlockPartLabel.name & " : " & oBlockPartLabel.Value)
		Call vdapp.AppendOutput("Script Output", chr(13))

	End If
Next
Call vdapp.AppendOutput("Script Output", "*** END CODE ***")
