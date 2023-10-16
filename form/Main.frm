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
   Begin VB.CommandButton run 
      Caption         =   "run"
      Height          =   372
      Left            =   11880
      TabIndex        =   12
      Top             =   5520
      Width           =   852
   End
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
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   4212
      Begin VB.ComboBox Combo1 
         Height          =   276
         Left            =   1800
         TabIndex        =   15
         Text            =   "Combo1"
         Top             =   840
         Width           =   1572
      End
      Begin VB.TextBox Text3 
         Height          =   264
         Left            =   1800
         TabIndex        =   10
         Text            =   "Text3"
         Top             =   360
         Width           =   1572
      End
      Begin VB.TextBox TextCMD 
         Height          =   372
         Left            =   240
         TabIndex        =   14
         Text            =   "CommandLine"
         Top             =   1200
         Visible         =   0   'False
         Width           =   3132
      End
      Begin VB.Label Preset 
         Caption         =   "Preset"
         Height          =   252
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   732
      End
      Begin VB.Label EncoderOptions 
         Caption         =   "Options"
         ForeColor       =   &H8000000D&
         Height          =   252
         Left            =   240
         TabIndex        =   13
         Top             =   1800
         Width           =   732
      End
      Begin VB.Label Quality 
         Caption         =   "Quality"
         Height          =   252
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   732
      End
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
         ForeColor       =   &H8000000D&
         Height          =   252
         Left            =   3360
         TabIndex        =   8
         Top             =   0
         Width           =   612
         WordWrap        =   -1  'True
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
         ForeColor       =   &H8000000D&
         Height          =   252
         Left            =   240
         TabIndex        =   7
         Top             =   0
         Width           =   852
         WordWrap        =   -1  'True
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
         Begin VB.Menu Menu_zh_cn 
            Caption         =   "简体中文"
         End
         Begin VB.Menu Menu_en_us 
            Caption         =   "English(US)"
         End
      End
      Begin VB.Menu MenuFFmpegPath 
         Caption         =   "FFmpegPath"
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
Option Explicit
Dim SelectEncoder$, SelectFormat$, BitrateControlMode$

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



Private Sub EncoderOptions_Click()
    Select Case SelectEncoder
    
    Case "libx265"

    Case "libx264"
    
    Case "libsvtav1"
    
    Case "librav1e"
    
    Case "libaom_av1"
    
    Case "hevc_nvenc"
    
    Case "h264_nvenc"
    
    Case "av1_nvenc"
    
    Case "hevc_amf"
    
    Case "h264_amf"
    
    Case "av1_amf"
    
    Case "hevc_qsv"
    
    Case "h264_qsv"
    
    Case "av1_qsv"
    
    Case "mjpeg_qsv"
    
    Case "mpeg2_qsv"
    
    Case "vp9_qsv"
    
    Case "Copy"
    
    
    End Select
    
    
End Sub

Private Sub Form_Load()
    Translate
    SelectEncoder = "libx265"
    SelectFormat = MenuMP4.Caption
    Label3.Caption = "x265"
    Label4.Caption = SelectFormat
    Text3.Text = ""
'    MsgBox Menu_libx264.Caption
'    MsgBox Menu_libx265.Caption
    
End Sub

Private Sub Label3_Change()

    If Label3.Caption = "Command" Then
        TextCMD.Visible = True
        Quality.Visible = False
        Preset.Visible = False
        Text3.Visible = False
        Combo1.Visible = False
    ElseIf Label3.Caption = "Copy" Then
        TextCMD.Visible = Not True
        Quality.Visible = False
        Preset.Visible = False
        Text3.Visible = False
        Combo1.Visible = False
    Else
        TextCMD.Visible = Not True
        Quality.Visible = Not False
        Preset.Visible = Not False
        Text3.Visible = Not False
        Combo1.Visible = Not False
    End If
End Sub

Private Sub Label3_Click()
    PopupMenu MenuEncoder
    MouseLeave
End Sub

Private Sub Translate()
'**********************MainScreen*********************************************
    Label1.Caption = GetTranslation("MainScreen", "Source")
    Label2.Caption = GetTranslation("MainScreen", "Target")
    Quality.Caption = GetTranslation("MainScreen", "quality")
    Preset.Caption = GetTranslation("MainScreen", "preset")
'**********************TopMenu*****************************************
    MenuAbout.Caption = GetTranslation("Menu", "About")
    MenuFile.Caption = GetTranslation("Menu", "File")
    MenuInput.Caption = GetTranslation("Menu", "input")
    MenuOutput.Caption = GetTranslation("Menu", "output")
    MenuExportBatch.Caption = GetTranslation("Menu", "ExportBatch")
    MenuQuit.Caption = GetTranslation("Menu", "Quit")
    MenuOption.Caption = GetTranslation("Menu", "Option")
    MenuLanguage.Caption = GetTranslation("Menu", "Language")
    MenuFFmpegPath.Caption = GetTranslation("Menu", "ffmpegPath")
'************************Buttom*******************************************
    run.Caption = GetTranslation("Bottom", "run")
    
End Sub

Private Sub Label4_Click()
    PopupMenu MenuFormat
    MouseLeave
End Sub

Private Sub Menu_av1_amf_Click()
    SelectEncoder = "av1_amf"
    Label3.Caption = SelectEncoder
End Sub

Private Sub Menu_av1_nvenc_Click()
    SelectEncoder = "av1_nvenc"
    Label3.Caption = SelectEncoder
End Sub

Private Sub Menu_av1_qsv_Click()
    SelectEncoder = "av1_qsv"
    Label3.Caption = SelectEncoder
End Sub

Private Sub Menu_en_us_Click()
    WriteIniKey "MainScreen", "Language", "en_us", App.Path & "\Config\config.ini"
    MsgBox "软件重启后生效", vbInformation, "需要重启"
End Sub



Private Sub Menu_h264_amf_Click()
    SelectEncoder = "h264_amf"
    Label3.Caption = SelectEncoder
End Sub

Private Sub Menu_h264_nvenc_Click()
    SelectEncoder = "h264_nvenc"
    Label3.Caption = SelectEncoder
End Sub

Private Sub Menu_h264_qsv_Click()
    SelectEncoder = "h264_qsv"
    Label3.Caption = SelectEncoder
End Sub

Private Sub Menu_hevc_amf_Click()
    SelectEncoder = "hevc_amf"
    Label3.Caption = SelectEncoder
End Sub

Private Sub Menu_hevc_nvenc_Click()
    SelectEncoder = "hevc_nvenc"
    Label3.Caption = SelectEncoder
End Sub

Private Sub Menu_hevc_qsv_Click()
    SelectEncoder = "hevc_qsv"
    Label3.Caption = SelectEncoder
End Sub

Private Sub Menu_libaom_av1_Click()
    SelectEncoder = "libaom_av1"
    Label3.Caption = SelectEncoder
End Sub

Private Sub Menu_librav1e_Click()
    SelectEncoder = "librav1e"
    Label3.Caption = SelectEncoder
End Sub

Private Sub Menu_libsvtav1_Click()
    SelectEncoder = "libsvtav1"
    Label3.Caption = SelectEncoder
End Sub

Private Sub Menu_libx264_Click()
    SelectEncoder = "libx264"
    Label3.Caption = "x264"
End Sub


Private Sub Menu_libx265_Click()
    SelectEncoder = "libx265"
    Label3.Caption = "x265"
End Sub

Private Sub Menu_mjpeg_qsv_Click()
    SelectEncoder = "mjpeg_qsv"
    Label3.Caption = SelectEncoder
End Sub

Private Sub Menu_mpeg2_qsv_Click()
    SelectEncoder = "mpeg2_qsv"
    Label3.Caption = SelectEncoder
End Sub

Private Sub Menu_vp9_qsv_Click()
    SelectEncoder = "vp9_qsv"
    Label3.Caption = SelectEncoder
End Sub

Private Sub Menu_zh_cn_Click()
    Dim k As VbMsgBoxResult
    WriteIniKey "MainScreen", "Language", "zh_cn", App.Path & "\Config\config.ini"
    MsgBox "软件重启后生效", vbInformation, "需要重启"
'    k = MsgBox("是否重启软件？", vbYesNo + vbQuestion, "需要重启")
'    If k = vbYes Then AppRestart
End Sub
'Private Sub AppRestart()
'
'End Sub
Private Sub MenuAbout_Click()
    frmAbout.Show
End Sub

Private Sub MenuCMDLine_Click()
    Label3.Caption = "Command"
End Sub

Private Sub MenuCopy_Click()
    SelectEncoder = "Copy"
    Label3.Caption = SelectEncoder
End Sub

Private Sub MenuFFmpegPath_Click()
    BasicOptions.Show vbModal
End Sub

Private Sub MenuInput_Click()
    Command1_Click
End Sub

Private Sub MenuMKV_Click()
    SelectFormat = MenuMKV.Caption
    Label4.Caption = SelectFormat
End Sub

Private Sub MenuMP4_Click()
    SelectFormat = MenuMP4.Caption
    Label4.Caption = SelectFormat
End Sub

Private Sub MenuOutput_Click()
    Command2_Click
End Sub

Private Sub MenuQuit_Click()
    End
End Sub


Private Sub Options_Click()

End Sub

Private Sub run_Click()
    On Error GoTo runErr
    Dim cmdstr$, SourceFile$, TargetFile$, Quot$, Spac$
    Quot = Chr(34): Spac = " "                          '双引号，空格
    SourceFile = Quot & Text1.Text & Quot
    TargetFile = Quot & Text2.Text & Quot
    cmdstr = "CMD /K" & "ffmpeg" & " -i " & SourceFile & Spac & TargetFile
    Shell cmdstr, vbNormalFocus
    Exit Sub
runErr:
    MsgBox "未能执行命令", vbCritical, "错误"
End Sub
Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' 高亮显示标签
    Label3.FontBold = True
    Label3.FontItalic = True
End Sub
Private Sub EncoderOptions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' 高亮显示标签
    EncoderOptions.FontBold = True
    EncoderOptions.FontItalic = True
End Sub
Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' 高亮显示标签
    Label4.FontBold = True
    Label4.FontItalic = True
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseLeave
End Sub
Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseLeave
End Sub
Private Sub MouseLeave()
    ' 移除标签的高亮显示
    Label3.FontBold = False
    Label3.FontItalic = False
    Label4.FontBold = False
    Label4.FontItalic = False
    EncoderOptions.FontBold = False
    EncoderOptions.FontItalic = False
End Sub

Private Sub TextCMD_GotFocus()
    If TextCMD.Text = "CommandLine" Then
        TextCMD.SelStart = 0
        TextCMD.SelLength = Len(TextCMD.Text)
    End If
End Sub
