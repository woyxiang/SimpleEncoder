Attribute VB_Name = "ForMod1"
'Public FFmpegPath$
'Dim FFmpegExist As Integer
Public SelectLanguage$
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'写ini文件

Public Sub WriteIniKey(strSection As String, strKey As String, strValue As String, strFileName As String)
    WritePrivateProfileString strSection, strKey, strValue, strFileName
End Sub

'读取ini文件

Public Function GetIniKey(strSection As String, strKey As String, strFileName As String) As String
    Dim strTmp As String
    Dim lngRet As Long
    strTmp = String$(1024, Chr(32))
    lngRet = GetPrivateProfileString(strSection, strKey, "", strTmp, Len(strTmp), strFileName)
    GetIniKey = Left$(strTmp, lngRet)
End Function

'读取翻译文件

Public Function GetTranslation(strSection As String, strKey As String)
    
    GetTranslation = GetIniKey(strSection, strKey, App.Path & "\Language\" & SelectLanguage & ".ini")
End Function

Sub Main()
    SelectLanguage = GetIniKey("MainScreen", "language", App.Path & "\Config\config.ini")
    Mainform.Show
End Sub
Public Function FFmpegPath$()

    Dim PathVar As String
    Dim PathArray() As String
    Dim i As Integer
'    Dim FFmpegPath As String
    
    ' 获取环境变量PATH
    PathVar = Environ("PATH")
    
    ' 使用Split函数将PATH环境变量分割成一个数组
    PathArray = Split(PathVar, ";")
    
    ' 遍历数组中的每一个路径
    For i = 0 To UBound(PathArray)
        ' 使用Dir函数检查是否存在名为"ffmpeg.exe"的文件
        If Dir(PathArray(i) & "\ffmpeg.exe") <> "" Then
            FFmpegPath = PathArray(i)
            Exit For
        End If
    Next i
    FFmpegPath = FFmpegPath
    
'    If FFmpegPath <> "" Then
'        'MsgBox "ffmpeg.exe位于以下目录：" & ffmpegPath
'        FFmpegExist = 1
'    Else
'        'MsgBox "ffmpeg.exe没有找到": End
'        FFmpegExist = 0
'    End If

End Function
Public Function FFmpegExist() As Boolean
    Call FFmpegExist
    If FFmpegPath <> "" Then
        FFmpegExist = True
    Else
        FFmpegExist = False
    End If
End Function
Private Sub FirstStart()
    Call FFmpegExist
    If FFmpegPath = "" Then Call SetFFmpeg
    If Dir(App.Path & "\Config\Config.ini") = "" Then Call SettingGuide
End Sub
Private Sub SetFFmpeg()

End Sub
Private Sub SettingGuide()

End Sub
