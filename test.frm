VERSION 5.00
Begin VB.Form test 
   Caption         =   "Form1"
   ClientHeight    =   4416
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11352
   LinkTopic       =   "Form1"
   ScaleHeight     =   4416
   ScaleWidth      =   11352
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Left            =   1200
      Top             =   2760
   End
   Begin VB.TextBox TextBox1 
      Height          =   3372
      Left            =   3960
      TabIndex        =   1
      Top             =   480
      Width           =   6612
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   612
      Left            =   1560
      TabIndex        =   0
      Top             =   1320
      Width           =   972
   End
End
Attribute VB_Name = "test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Sub Command1_Click()
'    ' 创建一个新的 output.txt 文件。
'    Dim intFileNumber As Integer
'    intFileNumber = FreeFile
'    Open "output.txt" For Output As #intFileNumber
'    Close #intFileNumber
    ' 使用 Shell 函数调用 FFMPEG，并将其输出重定向到一个文本文件。
    Shell "ffmpeg -i input.mp4 output.avi > output.txt"
    
    ' 启动定时器，每秒读取一次输出。
    Timer1.Interval = 1000
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    ' 定期读取输出，并显示在 TextBox1 上。
    Dim strOutput As String
    Dim intFileNumber As Integer
    ' 检查 output.txt 文件是否存在。
    Do
        ' 如果文件不存在，等待一段时间，然后再次检查。
        Sleep 1000
        If Dir("output.txt") <> "" Then
            
            Exit Do
        End If
    Loop

    
    intFileNumber = FreeFile
    Open "output.txt" For Input As #intFileNumber
    strOutput = Input$(LOF(intFileNumber), #intFileNumber)
    Close #intFileNumber
    
    TextBox1.Text = strOutput
End Sub

