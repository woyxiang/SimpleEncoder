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
   StartUpPosition =   3  '����ȱʡ
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
'    ' ����һ���µ� output.txt �ļ���
'    Dim intFileNumber As Integer
'    intFileNumber = FreeFile
'    Open "output.txt" For Output As #intFileNumber
'    Close #intFileNumber
    ' ʹ�� Shell �������� FFMPEG������������ض���һ���ı��ļ���
    Shell "ffmpeg -i input.mp4 output.avi > output.txt"
    
    ' ������ʱ����ÿ���ȡһ�������
    Timer1.Interval = 1000
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    ' ���ڶ�ȡ���������ʾ�� TextBox1 �ϡ�
    Dim strOutput As String
    Dim intFileNumber As Integer
    ' ��� output.txt �ļ��Ƿ���ڡ�
    Do
        ' ����ļ������ڣ��ȴ�һ��ʱ�䣬Ȼ���ٴμ�顣
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

