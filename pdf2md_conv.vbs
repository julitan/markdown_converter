Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' VBS 파일이 위치한 폴더 = 변환 결과 저장 위치
strVbsDir = objFSO.GetParentFolderName(WScript.ScriptFullName)
strAppDir = "g:\내 드라이브\RPA\RAG\pdf_to_markdown"
strCmd = "cmd /k ""cd /d """ & strAppDir & """ && C:\Users\USER\anaconda3\python.exe main.py --output-dir """ & strVbsDir & """"""

objShell.Run strCmd, 1, False
