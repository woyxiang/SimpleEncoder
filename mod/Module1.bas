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

' 获取环境变量PATH
PathVar = Environ("PATH")

' 使用Split函数将PATH环境变量分割成一个数组
PathArray = Split(PathVar, ";")

' 遍历数组中的每一个路径
For i = 0 To UBound(PathArray)
    ' 使用Dir函数检查是否存在名为"ffmpeg.exe"的文件
    If Dir(PathArray(i) & "\ffmpeg.exe") <> "" Then
        ffmpegPath = PathArray(i)
        Exit For
    End If
Next i

If ffmpegPath <> "" Then
    'MsgBox "ffmpeg.exe位于以下目录：" & ffmpegPath
Else
    MsgBox "ffmpeg.exe没有找到": End
End If

End Sub
