'==============================================================================================================================
' 현장 테스트용
' decryt가 얼마나 걸리는지....
' 사용법 : .lmc 파일이 있는 해당 위치에서 이 .vbs 스크립트를 실행시키면 됌.
' 결과 : _Test 폴더에 decrypt된 파일들이 생성됨. (시간이 얼마나 오래걸리는지 보기 - 용량체크)
'        _TimeLog.txt 파일에 시작시간, 끝시간, 얼마나 걸렸는지 표시됌
'==============================================================================================================================
' Last Update: 2022/7/29
' Created by WJ

Option Explicit

' 폴더 생성을 위함
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

' 윈도우 Shell에 접근하기 위함
Dim Win
Set Win = CreateObject("WScript.shell")

' 실행되는 PCB와 동일한 directory에 .hkp decrypt를 한 파일을 저장할 폴더 생성
On Error Resume Next
Dim folderHandler
Set folderHandler = fso.CreateFolder(".\_Test")
On Error GoTo 0

' Decript 파일 경로 변수 정의
Dim sToolVersion, sUtilityDir, sCentralLibDir, sPartDBDir, sCellDBDir
sToolVersion = "EEVX.2.11"
sUtilityDir = "C:\MentorGraphics\" & sToolVersion & "\SDD_HOME\common\win64\bin\"

' Central Library의 경로를 넣음 (.lmc 파일이 위치한 곳)
sCentralLibDir = ".\"
sPartDBDir = sCentralLibDir & "PartsDBLibs\"
sCellDBDir = sCentralLibDir & "CellDBLibs\"

' 파일들의 경로 (현재경로)
Dim savedPath, oFolder, sFileName, oFile, sTrimedName, txtFile
savedPath = ".\_Test\"

' 현재경로에 임시 폴더 생성, 이미 있는경우 경고 발생, 에러 무시 코드 넣기
On Error Resume Next
Set txtFile = fso.CreateTextFile(savedPath & "_TimeLog1.txt", true)
On Error GoTo 0


' ==== 코드 시작 시간 ===
Dim dtmStartTime, dtmEndTime, strElapsedTime
dtmStartTime = Date & " " & Time
txtFile.Writeline("Start At : " & dtmStartTime)

' ' decript exe파일을 실행, PDB 데이터 파씽
' Win.run sUtilityDir & "PartsDB2HKP.exe -i " & sPartDBDir & "Discrete.pdb" & " -o " & savedPath & "decryptPDB.hkp -a"

' .hkp decrypt를 한 파일을 저장할 폴더 생성
If (fso.FolderExists(savedPath)) Then
    Set oFolder = fso.GetFolder(sPartDBDir)

    '_Test 폴더에 .PDB파일을 모두 decrypt 하여 저장
    For Each oFile in oFolder.Files
        sFileName = oFile.name
        ' hkp파일로 사용하기 위해 뒤에 .pdb를 지우고 공백까지
        sTrimedName = Rtrim(Replace(sFileName,".pdb",""))
        ' decript exe파일을 실행, PDB 데이터 파씽
        Win.run sUtilityDir & "PartsDB2HKP.exe -i " & sPartDBDir & sFileName & " -o " & savedPath & sTrimedName & ".hkp -a"
    Next

    Set oFolder = fso.GetFolder(sCellDBDir)
    '_Test 폴더에 .Cel파일을 모두 decrypt 하여 저장
    For Each oFile in oFolder.Files
        sFileName = oFile.name
        ' hkp파일로 사용하기 위해 뒤에 .pdb를 지우고 공백까지
        sTrimedName = Rtrim(Replace(sFileName,".cel",""))
        ' decript exe파일을 실행, PDB 데이터 파씽
        Win.run sUtilityDir & "CellDB2HKP.exe -i " & sCellDBDir & sFileName & " -o " & savedPath & sTrimedName & ".hkp -a"
    Next

End If

' === 코드 끝 시간 ===
dtmEndTime = Date & " " & Time
strElapsedTime = DateDiff("s", dtmStartTime, dtmEndTime)
txtFile.Writeline("End At : " & dtmEndTime )
txtFile.Writeline("Taking Time (sec) : " & strElapsedTime)

txtFile.close
MsgBox("Complete !")
