Attribute VB_Name = "Module1"
Sub Main()
    Call FFmpegExist
    Mainform.Show
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
