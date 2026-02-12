Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' VBS 파일이 위치한 폴더 = 변환 결과 저장 위치
strVbsDir = objFSO.GetParentFolderName(WScript.ScriptFullName)

' Google Drive 경로 자동 감지 (한글/영문)
strAppDir = ""
arrPaths = Array( _
    "g:\내 드라이브\RPA\RAG\pdf_to_markdown", _
    "g:\My Drive\RPA\RAG\pdf_to_markdown" _
)

For Each strPath In arrPaths
    If objFSO.FolderExists(strPath) Then
        strAppDir = strPath
        Exit For
    End If
Next

If strAppDir = "" Then
    MsgBox "pdf_to_markdown 폴더를 찾을 수 없습니다.", vbCritical, "오류"
    WScript.Quit
End If

' Python 경로 자동 감지
strPython = ""
strUserProfile = objShell.ExpandEnvironmentStrings("%USERPROFILE%")
arrPythonPaths = Array( _
    strUserProfile & "\anaconda3\python.exe", _
    strUserProfile & "\miniconda3\python.exe", _
    "D:\Anaconda3\python.exe", _
    strUserProfile & "\AppData\Local\Programs\Python\Python311\python.exe", _
    strUserProfile & "\AppData\Local\Programs\Python\Python310\python.exe", _
    "C:\ProgramData\anaconda3\python.exe", _
    "C:\anaconda3\python.exe" _
)

For Each strPath In arrPythonPaths
    If objFSO.FileExists(strPath) Then
        strPython = strPath
        Exit For
    End If
Next

If strPython = "" Then
    MsgBox "Python을 찾을 수 없습니다. Anaconda 또는 Python이 설치되어 있는지 확인하세요.", vbCritical, "오류"
    WScript.Quit
End If

strCmd = "cmd /k ""cd /d """ & strAppDir & """ && """ & strPython & """ main.py --output-dir """ & strVbsDir & """"""

objShell.Run strCmd, 1, False
