VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Mainform 
   Caption         =   "SimpleEncoder"
   ClientHeight    =   6012
   ClientLeft      =   108
   ClientTop       =   756
   ClientWidth     =   12876
   BeginProperty Font 
      Name            =   "黑体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6012
   ScaleWidth      =   12876
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2172
      Left            =   480
      TabIndex        =   6
      Top             =   1200
      Width           =   4212
      Begin VB.Label Label4 
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   3360
         TabIndex        =   8
         Top             =   0
         Width           =   612
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   240
         TabIndex        =   7
         Top             =   0
         Width           =   612
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   12120
      TabIndex        =   3
      Top             =   120
      Width           =   612
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   5160
      TabIndex        =   2
      Top             =   120
      Width           =   612
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   7440
      TabIndex        =   1
      Top             =   120
      Width           =   4212
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   4212
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5880
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "output"
      BeginProperty Font 
         Name            =   "思源黑体 CN Bold"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   6360
      TabIndex        =   5
      Top             =   120
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "input"
      BeginProperty Font 
         Name            =   "思源黑体 CN Bold"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   612
   End
   Begin VB.Menu MenuFormat 
      Caption         =   "Format"
      Visible         =   0   'False
      Begin VB.Menu MenuMP4 
         Caption         =   "MP4"
      End
      Begin VB.Menu MenuMKV 
         Caption         =   "MKV"
      End
      Begin VB.Menu MenuOthers 
         Caption         =   "Others"
      End
   End
   Begin VB.Menu MenuEncoder 
      Caption         =   "Encoder"
      Visible         =   0   'False
      Begin VB.Menu Menu_libx265 
         Caption         =   "x265"
      End
      Begin VB.Menu Menu_libx264 
         Caption         =   "x264"
      End
      Begin VB.Menu MenuAV1 
         Caption         =   "AV1"
         Begin VB.Menu Menu_libsvtav1 
            Caption         =   "libsvtav1"
         End
         Begin VB.Menu Menu_librav1e 
            Caption         =   "librav1e"
         End
         Begin VB.Menu Menu_libaom_av1 
            Caption         =   "libaom-av1"
         End
      End
      Begin VB.Menu MenuNvEnc 
         Caption         =   "NvEnc(Nvidia)"
         Begin VB.Menu Menu_hevc_nvenc 
            Caption         =   "H.265"
         End
         Begin VB.Menu Menu_h264_nvenc 
            Caption         =   "H.264"
         End
         Begin VB.Menu Menu_av1_nvenc 
            Caption         =   "AV1"
         End
      End
      Begin VB.Menu MenuVcEnc 
         Caption         =   "VcEnc(AMD)"
         Begin VB.Menu Menu_hevc_amf 
            Caption         =   "H.265"
         End
         Begin VB.Menu Menu_h264_amf 
            Caption         =   "H.264"
         End
         Begin VB.Menu Menu_av1_amf 
            Caption         =   "AV1"
         End
      End
      Begin VB.Menu MenuQsvEnc 
         Caption         =   "QsvEnc(Intel)"
         Begin VB.Menu Menu_hevc_qsv 
            Caption         =   "H.265"
         End
         Begin VB.Menu Menu_h264_qsv 
            Caption         =   "H.264"
         End
         Begin VB.Menu Menu_av1_qsv 
            Caption         =   "AV1"
         End
         Begin VB.Menu Menu_mjpeg_qsv 
            Caption         =   "mjpeg"
         End
         Begin VB.Menu Menu_mpeg2_qsv 
            Caption         =   "MPEG-2"
         End
         Begin VB.Menu Menu_vp9_qsv 
            Caption         =   "VP9"
         End
      End
      Begin VB.Menu MenuCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu MenuCMDLine 
         Caption         =   "CommandLine"
      End
   End
   Begin VB.Menu MenuFile 
      Caption         =   "File(&F)"
      Begin VB.Menu MenuInput 
         Caption         =   "input"
      End
      Begin VB.Menu MenuOutput 
         Caption         =   "output"
      End
      Begin VB.Menu MenuFileSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu MenuExportBatch 
         Caption         =   "ExportBatch"
      End
      Begin VB.Menu MenuQuit 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu MenuOption 
      Caption         =   "Option(&O)"
      Begin VB.Menu MenuLanguage 
         Caption         =   "Language"
      End
   End
   Begin VB.Menu MenuAbout 
      Caption         =   "About(&A)"
   End
End
Attribute VB_Name = "Mainform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'选择输入文件
    CommonDialog1.Filter = "所有文件"
    CommonDialog1.ShowOpen
    Text1.Text = CommonDialog1.FileName

End Sub

Private Sub Command2_Click()
'保存输出文件
    CommonDialog1.ShowSave
    Text2.Text = CommonDialog1.FileName
End Sub

Private Sub Form_Load()
    Translate
End Sub

Private Sub Label3_Click()
    PopupMenu MenuEncoder
End Sub

Private Sub Translate()
    Label1.Caption = GetTranslation("MainScreen", "Source")
    Label2.Caption = GetTranslation("MainScreen", "Target")
    MenuAbout.Caption = GetTranslation("Menu", "About")
    MenuFile.Caption = GetTranslation("Menu", "File")
    MenuInput.Caption = GetTranslation("Menu", "input")
    MenuOutput.Caption = GetTranslation("Menu", "output")
    MenuExportBatch.Caption = GetTranslation("Menu", "ExportBatch")
    MenuQuit.Caption = GetTranslation("Menu", "Quit")
    MenuOption.Caption = GetTranslation("Menu", "Option")
    MenuLanguage.Caption = GetTranslation("Menu", "Language")
End Sub

Private Sub Label4_Click()
    PopupMenu MenuFormat
End Sub

Private Sub MenuAbout_Click()
    frmAbout.Show
End Sub

Private Sub MenuInput_Click()
    Command1_Click
End Sub

Private Sub MenuOutput_Click()
    Command2_Click
End Sub

Private Sub MenuQuit_Click()
    End
End Sub
