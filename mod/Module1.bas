Attribute VB_Name = "ForMod"
'Public FFmpegPath$
'Dim FFmpegExist As Integer

Public SelectLanguage$, ConfigPath$
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Enum CheckFFmpegResault
    NotInPath = 0
    InPath = 1
    FileExists = 2
    NotFileExists = 3
    InvalidPath = 4
    NoResault = -1
End Enum
'дini�ļ�

Public Sub WriteIniKey(strSection As String, strKey As String, strValue As String, strFileName As String)
    WritePrivateProfileString strSection, strKey, strValue, strFileName
End Sub

'��ȡini�ļ�

Public Function GetIniKey(strSection As String, strKey As String, strFileName As String) As String
    Dim strTmp As String
    Dim lngRet As Long
    strTmp = String$(1024, Chr(32))
    lngRet = GetPrivateProfileString(strSection, strKey, "", strTmp, Len(strTmp), strFileName)
    GetIniKey = Left$(strTmp, lngRet)
End Function

'��ȡ�����ļ�

Public Function GetTranslation(strSection As String, strKey As String)
    
    GetTranslation = GetIniKey(strSection, strKey, App.path & "\Language\" & SelectLanguage & ".ini")
End Function

Sub Main()
    
    ConfigPath = App.path & "\Config\config.ini"
    SelectLanguage = GetIniKey("MainScreen", "language", App.path & "\Config\config.ini")
    '************************************************************************************************
    If GetIniKey("BasicOption", "checkFFmpeg", ConfigPath) = "yes" Then
        InformCheckFFmpegResault (True)
    Else
        Mainform.Show
    End If
    

End Sub
Private Sub InformCheckFFmpegResault(ByVal OP As Boolean)
    '���ؼ����
    If OP = True Then
        Select Case CheckFFmpeg
        Case NotInPath
           
            If MsgBox(GetTranslation("Fore", "notInPath"), vbYesNo + vbExclamation, GetTranslation("Title", "checkOnEach")) = vbYes Then
                Mainform.Show
                BasicOptions.Show vbModal
            Else
                Mainform.Show
            End If
        Case InPath
            Mainform.Show
        Case InvalidPath
            If MsgBox(GetTranslation("Fore", "invalidPath"), vbYesNo + vbExclamation, GetTranslation("Title", "checkOnEach")) = vbYes Then
                Mainform.Show
                BasicOptions.Show vbModal
            Else
                Mainform.Show
            End If
        Case FileExists
                Mainform.Show
        Case NotFileExists
            If MsgBox(GetTranslation("Fore", "notFileExists"), vbYesNo + vbExclamation, GetTranslation("Title", "checkOnEach")) = vbYes Then
                Mainform.Show
                BasicOptions.Show vbModal
            Else
                Mainform.Show
            End If
        End Select
        
    End If
End Sub
Private Function CheckFFmpeg() As CheckFFmpegResault
    Dim Quot$, path$
    path = Quot & GetIniKey("BasicOption", "ffmpeg", ConfigPath) & Quot
    Quot = Chr(34)
    If GetIniKey("BasicOption", "ffmpeg", ConfigPath) = "path" Then
        If FFmpegExistInPath Then
            CheckFFmpeg = InPath
        Else
            CheckFFmpeg = NotInPath
        End If
    ElseIf IsFFmpegPath(GetIniKey("BasicOption", "ffmpeg", ConfigPath)) = False Then
        CheckFFmpeg = InvalidPath
    ElseIf Dir(path) = "" Then
        CheckFFmpeg = NotFileExists
    ElseIf Dir(path) <> "" And InStr(1, GetIniKey("BasicOption", "ffmpeg", ConfigPath), Dir(path), vbTextCompare) <> 0 Then
        CheckFFmpeg = FileExists
    Else
        CheckFFmpeg = NoResault
    End If
        
End Function
Public Function IsFFmpegPath(addr As String) As Boolean
    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")
    
    With reg
        .Pattern = ".*ffmpeg\.exe"
        .IgnoreCase = True
        .Global = True
        IsFFmpegPath = .Test(addr)
    End With

End Function
Public Function FFmpegPath$()

    Dim PathVar As String
    Dim PathArray() As String
    Dim i As Integer
    Dim path As String
    
    ' ��ȡ��������PATH
    PathVar = Environ("PATH")
    
    ' ʹ��Split������PATH���������ָ��һ������
    PathArray = Split(PathVar, ";")
    
    ' ���������е�ÿһ��·��
    For i = 0 To UBound(PathArray)
        ' ʹ��Dir��������Ƿ������Ϊ"ffmpeg.exe"���ļ�
        If Dir(PathArray(i) & "\ffmpeg.exe") <> "" Then
            path = PathArray(i)
            Exit For
        End If
    Next i
      FFmpegPath = path
    
'    If FFmpegPath <> "" Then
'        'MsgBox "ffmpeg.exeλ������Ŀ¼��" & ffmpegPath
'        FFmpegExist = 1
'    Else
'        'MsgBox "ffmpeg.exeû���ҵ�": End
'        FFmpegExist = 0
'    End If

End Function
Public Function FFmpegExistInPath() As Boolean
    If FFmpegPath <> "" Then
        FFmpegExistInPath = True
    Else
        FFmpegExistInPath = False
    End If
End Function
Private Sub FirstStart()
    If FFmpegPath = "" Then Call SetFFmpeg
    If Dir(App.path & "\Config\Config.ini") = "" Then Call SettingGuide
End Sub
Private Sub SetFFmpeg()

End Sub
Private Sub SettingGuide()
'    MsgBox "test"

End Sub
