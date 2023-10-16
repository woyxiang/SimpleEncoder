VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form BasicOptions 
   Caption         =   "BasicOptions"
   ClientHeight    =   4116
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6648
   LinkTopic       =   "Form1"
   ScaleHeight     =   4116
   ScaleWidth      =   6648
   StartUpPosition =   3  '窗口缺省
   Begin VB.CheckBox Check3 
      Caption         =   "Check3"
      Height          =   372
      Left            =   4080
      TabIndex        =   10
      Top             =   3000
      Value           =   1  'Checked
      Width           =   1452
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   372
      Left            =   4080
      TabIndex        =   9
      Top             =   2400
      Value           =   1  'Checked
      Width           =   2052
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   372
      Left            =   720
      TabIndex        =   8
      Top             =   1320
      Width           =   2052
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   600
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.OptionButton Option3 
      Caption         =   "CheckAtEach"
      Height          =   252
      Left            =   1800
      TabIndex        =   7
      Top             =   2520
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1692
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Check"
      Height          =   372
      Left            =   1920
      TabIndex        =   6
      Top             =   3000
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.CommandButton CMDCancel 
      Caption         =   "Cancel"
      Height          =   492
      Left            =   3360
      TabIndex        =   5
      Top             =   3480
      Width           =   1452
   End
   Begin VB.CommandButton CMDApply 
      Caption         =   "Apply"
      Height          =   492
      Left            =   4920
      TabIndex        =   4
      Top             =   3480
      Width           =   1452
   End
   Begin VB.OptionButton Option1 
      Caption         =   "envPath"
      Height          =   372
      Left            =   600
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
      Width           =   2052
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   372
      Left            =   5160
      TabIndex        =   2
      Top             =   480
      Width           =   972
   End
   Begin VB.TextBox Text1 
      Height          =   372
      Left            =   1200
      TabIndex        =   1
      Top             =   480
      Width           =   3612
   End
   Begin VB.Label Label1 
      Caption         =   "path"
      Height          =   372
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   972
   End
End
Attribute VB_Name = "BasicOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ConfirmTip$, shouldCancel As Boolean
Private Sub Translate()
    ConfirmTip = GetTranslation("BasicOptions", "confirmTip")
    Check1.Caption = GetTranslation("BasicOptions", "envpath")
    Check2.Caption = GetTranslation("BasicOptions", "checkAtEach")
    Check3.Caption = GetTranslation("BasicOptions", "check")
    Label1.Caption = GetTranslation("BasicOptions", "path")
    CMDCancel.Caption = GetTranslation("BasicOptions", "Cancel")
    CMDApply.Caption = GetTranslation("BasicOptions", "Apply")
End Sub

Private Sub Check1_Click()
    If Check1.Value = 1 Then
        WriteIniKey "Fore", "ffmpeg", "path", ConfigPath
    End If
End Sub

Private Sub CMDApply_Click()

    If Check3.Value = 1 Then
        If Check1.Value = 1 Then
            CheckApply
        ElseIf Not IsFFmpegPath(Text1.Text) Then
            MsgBox GetTranslation("BasicOptions", "queryPath"), vbQuestion
            shouldCancel = False
        Else
            shouldCancel = True
            Unload BasicOptions
        End If
    Else
        shouldCancel = True
        Unload BasicOptions
    End If
End Sub
Private Function IsFFmpegPath(addr As String) As Boolean
    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")
    
    With reg
        .Pattern = ".*ffmpeg\.exe"
        .IgnoreCase = True
        .Global = True
        IsFFmpegPath = .Test(addr)
    End With

End Function

Private Sub CheckApply()
    If FFmpegExist Then
        shouldCancel = True
        Unload BasicOptions
    Else
        MsgBox GetTranslation("BasicOptions", "wrongPathVar"), vbCritical, GetTranslation("Title", "Err")
        Check1.Value = 0
    End If
End Sub

Private Sub CMDCancel_Click()
    shouldCancel = True
    Unload BasicOptions
End Sub

Private Sub Command1_Click()
'选择输入文件
    CommonDialog1.Filter = "所有文件"
    CommonDialog1.ShowOpen
    Text1.Text = CommonDialog1.FileName
End Sub

Private Sub Form_Load()
    Translate
    Check1.ToolTipText = ConfirmTip
    If GetIniKey("Fore", "ffmpeg", ConfigPath) = "path" Then Check1.Value = 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        If MsgBox(GetTranslation("BasicOptions", "querySave"), vbYesNo + vbQuestion) = vbYes Then
           CMDApply_Click
            
        Else
            CMDCancel_Click
            
        End If
        
    End If
    If shouldCancel Then
        Cancel = 0
    Else
        Cancel = 1
    End If
    
End Sub
