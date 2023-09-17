Attribute VB_Name = "ForMod1"
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

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
    GetTranslation = GetIniKey(strSection, strKey, App.Path & "\Config\Language.ini")
End Function

Sub Main()
'    Mainform.Show
End Sub
Public Sub FFmpegExist()

    Dim PathVar As String
    Dim PathArray() As String
    Dim i As Integer
    Dim ffmpegPath As String
    
    ' ��ȡ��������PATH
    PathVar = Environ("PATH")
    
    ' ʹ��Split������PATH���������ָ��һ������
    PathArray = Split(PathVar, ";")
    
    ' ���������е�ÿһ��·��
    For i = 0 To UBound(PathArray)
        ' ʹ��Dir��������Ƿ������Ϊ"ffmpeg.exe"���ļ�
        If Dir(PathArray(i) & "\ffmpeg.exe") <> "" Then
            ffmpegPath = PathArray(i)
            Exit For
        End If
    Next i
    
    If ffmpegPath <> "" Then
        'MsgBox "ffmpeg.exeλ������Ŀ¼��" & ffmpegPath
    Else
        MsgBox "ffmpeg.exeû���ҵ�": End
    End If

End Sub
Private Sub FirstStart()
    
End Sub

