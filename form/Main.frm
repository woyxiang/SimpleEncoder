VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Mainform 
   Caption         =   "SimpleEncoder"
   ClientHeight    =   6024
   ClientLeft      =   108
   ClientTop       =   756
   ClientWidth     =   13632
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
   ScaleHeight     =   6024
   ScaleWidth      =   13632
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton CommandPlay 
      Caption         =   "Play"
      Height          =   372
      Left            =   3120
      TabIndex        =   37
      Top             =   5520
      Width           =   852
   End
   Begin VB.CommandButton CommandLog 
      Caption         =   "log"
      Height          =   372
      Left            =   1200
      TabIndex        =   36
      Top             =   5520
      Width           =   852
   End
   Begin VB.CommandButton CommandAbout 
      Caption         =   "About"
      Height          =   372
      Left            =   240
      TabIndex        =   35
      Top             =   5520
      Width           =   852
   End
   Begin VB.CommandButton Commandwindow 
      Caption         =   "window"
      Height          =   372
      Left            =   2160
      TabIndex        =   34
      Top             =   5520
      Width           =   852
   End
   Begin VB.CommandButton CommandSave 
      Caption         =   "Batch"
      Height          =   372
      Left            =   9600
      TabIndex        =   33
      Top             =   5520
      Width           =   852
   End
   Begin VB.CommandButton CommandT 
      Caption         =   "Termina"
      Height          =   372
      Left            =   10560
      TabIndex        =   32
      Top             =   5520
      Width           =   852
   End
   Begin VB.CommandButton CommandCode 
      Caption         =   "code"
      Height          =   372
      Left            =   11520
      Picture         =   "Main.frx":3C3A
      TabIndex        =   31
      Top             =   5520
      Width           =   852
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   6480
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Height          =   2172
      Left            =   9120
      TabIndex        =   25
      Top             =   1920
      Width           =   4212
      Begin VB.TextBox Text7 
         Height          =   264
         Left            =   2280
         TabIndex        =   30
         Text            =   "Text7"
         Top             =   840
         Width           =   1332
      End
      Begin VB.ComboBox Combo2 
         Height          =   276
         ItemData        =   "Main.frx":3D13
         Left            =   2280
         List            =   "Main.frx":3D1D
         TabIndex        =   28
         Text            =   "Combo2"
         Top             =   360
         Width           =   1332
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Label11"
         Height          =   180
         Left            =   360
         TabIndex        =   29
         Top             =   840
         Width           =   672
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Label10"
         Height          =   252
         Left            =   360
         TabIndex        =   27
         Top             =   360
         Width           =   732
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Audio"
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   360
         TabIndex        =   26
         Top             =   0
         Width           =   600
      End
   End
   Begin VB.CommandButton run 
      Caption         =   "run"
      Height          =   372
      Left            =   12480
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
      Top             =   1920
      Width           =   4212
      Begin VB.ComboBox Combo1 
         Height          =   276
         ItemData        =   "Main.frx":3D2C
         Left            =   1800
         List            =   "Main.frx":3D2E
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   840
         Width           =   1572
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
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
         Top             =   1320
         Visible         =   0   'False
         Width           =   3132
      End
      Begin VB.TextBox Text6 
         Height          =   264
         Left            =   1800
         TabIndex        =   24
         Text            =   "Text6"
         Top             =   360
         Width           =   1572
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Bitrate"
         Height          =   180
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Width           =   672
      End
      Begin VB.Label Preset 
         AutoSize        =   -1  'True
         Caption         =   "Preset"
         Height          =   180
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   576
      End
      Begin VB.Label EncoderOptions 
         AutoSize        =   -1  'True
         Caption         =   "Options"
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   240
         TabIndex        =   13
         Top             =   1800
         Visible         =   0   'False
         Width           =   672
      End
      Begin VB.Label Quality 
         AutoSize        =   -1  'True
         Caption         =   "Quality"
         Height          =   180
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   672
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         Height          =   180
         Left            =   3360
         TabIndex        =   8
         Top             =   0
         Width           =   252
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
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
         Height          =   180
         Left            =   360
         TabIndex        =   7
         Top             =   0
         Width           =   372
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
      Left            =   12840
      TabIndex        =   3
      Top             =   600
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
      Top             =   600
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
      Left            =   8520
      TabIndex        =   1
      Top             =   600
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
      Top             =   600
      Width           =   4212
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5880
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Height          =   2172
      Left            =   4680
      TabIndex        =   16
      Top             =   1920
      Width           =   4212
      Begin VB.CheckBox Check1 
         Caption         =   "ReCheck1"
         Height          =   252
         Left            =   240
         TabIndex        =   22
         Top             =   1200
         Width           =   1092
      End
      Begin VB.TextBox Text5 
         Height          =   264
         Left            =   2400
         TabIndex        =   21
         Text            =   "Text5"
         Top             =   840
         Width           =   1452
      End
      Begin VB.TextBox Text4 
         Height          =   264
         Left            =   2400
         TabIndex        =   18
         Text            =   "Text4"
         Top             =   360
         Width           =   1452
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "HLabel7"
         Height          =   180
         Left            =   240
         TabIndex        =   20
         Top             =   840
         Width           =   672
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "WLabel6"
         Height          =   180
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   672
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Size"
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
         Height          =   180
         Left            =   360
         TabIndex        =   17
         Top             =   0
         Width           =   384
      End
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
      Left            =   7320
      TabIndex        =   5
      Top             =   600
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
      Top             =   600
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
         Visible         =   0   'False
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
         Visible         =   0   'False
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
   Begin VB.Menu MenuSize 
      Caption         =   "Size"
      Visible         =   0   'False
      Begin VB.Menu Menu3840Size 
         Caption         =   "3840×auto"
      End
      Begin VB.Menu Menu3200Size 
         Caption         =   "3200×auto"
      End
      Begin VB.Menu Menu2560Size 
         Caption         =   "2560×auto"
      End
      Begin VB.Menu MenuSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu Menu1920Size 
         Caption         =   "1920×auto"
      End
      Begin VB.Menu Menu1280Size 
         Caption         =   "1280×auto"
      End
      Begin VB.Menu Menu960Size 
         Caption         =   "960×auto"
      End
      Begin VB.Menu MenuSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu Menu720Size 
         Caption         =   "720×auto"
      End
      Begin VB.Menu Menu640Size 
         Caption         =   "640×auto"
      End
      Begin VB.Menu Menu480Size 
         Caption         =   "480×auto"
      End
      Begin VB.Menu Menu320Size 
         Caption         =   "320×auto"
      End
   End
End
Attribute VB_Name = "Mainform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SelectEncoder$, SelectFormat$, BitrateControlMode$, VideoSize$, AuEncode$, myPreset$, RateMode$, TerminalCancel As Boolean

Private Sub Check1_Click()
    If Check1.Value = 1 Then
        Text4.Enabled = True
        Text5.Enabled = True
    ElseIf Check1.Value = 0 Then
        Text4.Enabled = False
        Text5.Enabled = False
    End If
        
End Sub



Private Sub Combo1_Change()
    If Label3.Caption = "x265" Or Label3.Caption = "x264" Then
        Select Case Combo1.ListIndex
        Case 0
            myPreset = "veryfast"
        Case 1
            myPreset = "medium"
        Case 2
            myPreset = "veryslow"
        End Select
    Else
        
    End If
End Sub

Private Sub Combo2_Click()
    If Combo2.ListIndex <> 0 Then
        Label11.Visible = True
        Text7.Visible = True
    Else
        Label11.Visible = False
        Text7.Visible = False
    End If
    
    Select Case Combo2.ListIndex
    Case 0
        AuEncode = "copy"
    Case 1
        AuEncode = "aac"
    End Select
        
End Sub



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



Private Sub CommandAbout_Click()
    MenuAbout_Click
End Sub

Private Sub CommandCode_Click()
    Dim cmdstr$, SourceFile$, TargetFile$, Quot$, Spac$, CMDOPEN$, crf$, Bitrate$, VideoWidth$, VideoHeight$
    Dim cmdstr1, cmdstr2, cmdstr3, cmdstr4, cmdstr5, cmdstr6
    Dim ffmpeg$
    If GetIniKey("BasicOption", "ffmpeg", ConfigPath) = "path" Then
        ffmpeg = "ffmpeg"
    Else
        ffmpeg = GetIniKey("BasicOption", "ffmpeg", ConfigPath)
    End If
    
    Quot = Chr(34): Spac = " "                          '双引号，空格
    SourceFile = Quot & Text1.Text & Quot
    TargetFile = Quot & Text2.Text & Quot
    '开头
    cmdstr1 = ffmpeg & " -y "
    '输入文件
    cmdstr2 = " -i " & SourceFile
    '编码器、preset、码率
    Bitrate = Text6.Text
    Select Case Combo1.ListIndex
        Case 0
            myPreset = "veryfast"
        Case 1
            myPreset = "medium"
        Case 2
            myPreset = "veryslow"
    End Select
    If RateMode = "crf" And SelectEncoder <> "copy" Then
        cmdstr3 = " -c:v " & SelectEncoder & " -preset " & myPreset & " -crf " & crf
    ElseIf RateMode = "VBR" And SelectEncoder <> "copy" Then
        cmdstr3 = " -c:v " & SelectEncoder & " -b:v " & Bitrate
    ElseIf SelectEncoder = "copy" Then
        cmdstr3 = " -c:v " & SelectEncoder
    End If

    '分辨率
    If Text4.Text = "Auto" Then
        VideoWidth = "-2"
    Else
        VideoWidth = Text4.Text
    End If
    If Text5.Text = "Auto" Then
        VideoHeight = "-2"
    Else
        VideoHeight = Text5.Text
    End If
    If Check1.Value = 0 And SelectEncoder <> "copy" Then
        cmdstr4 = " -vf ""scale=trunc(iw/2)*2:trunc(ih/2)*2"" "
    ElseIf Check1.Value = 1 And SelectEncoder <> "copy" Then
        cmdstr4 = " -vf ""scale=" & VideoWidth & ":" & VideoHeight
    ElseIf SelectEncoder <> "copy" Then
        cmdstr4 = ""
    End If
    '音频编码
    If Combo2.ListIndex = 0 Then
        cmdstr5 = " -c:a copy "
    ElseIf Combo2.ListIndex = 1 Then
        cmdstr5 = " -c:a aac " & " -b:a " & Text7.Text
    End If
    '元数据映射
    cmdstr6 = " -map_chapters 0 -map_metadata 0 "
    '输出文件
    cmdstr6 = " -f " & SelectFormat & Spac & TargetFile
    cmdstr = cmdstr1 & cmdstr2 & cmdstr3 & cmdstr4 & cmdstr5 & cmdstr6
    MsgBox cmdstr
End Sub

Private Sub CommandLog_Click()
    Dim logPath
    logPath = App.path & "\logs"
    Shell "explorer.exe " & Chr(34) & logPath & Chr(34), vbNormalFocus
End Sub

Private Sub CommandPlay_Click()
    On Error GoTo playErr
    Dim fs As New FileSystemObject
    If fs.FileExists(Text2.Text) = False Then GoTo playErr
    Shell "explorer.exe " & Chr(34) & Text2.Text & Chr(34), vbNormalFocus
    Exit Sub
playErr:
    MsgBox GetTranslation("MsgBox", "noVideo"), vbExclamation
End Sub

Private Sub CommandSave_Click()
    MenuExportBatch_Click
End Sub

Private Sub CommandT_Click()
    Shell "cmd", vbNormalFocus
End Sub

Private Sub Commandwindow_Click()
    Static a As Byte
     a = a + 1
     If (a Mod 2 = 1) Then
        Commandwindow.Caption = Replace(Commandwindow.Caption, "×", "√")
        TerminalCancel = True
     Else
        Commandwindow.Caption = Replace(Commandwindow.Caption, "√", "×")
        TerminalCancel = False
    End If
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
    MouseLeave
    
End Sub

Private Sub Form_Load()
    Translate
    SelectEncoder = "libx265"
    SelectFormat = "mp4"
    Label3.Caption = "x265"
    Label4.Caption = SelectFormat
    Text3.Text = "23"
    Text6.Text = "3500k"
    Text4.Text = ""
    Text5.Text = ""
    Text7.Text = "128k"
    Check1_Click
    Combo1.ListIndex = 1
    Combo2.ListIndex = 0
    Combo2_Click
    Commandwindow.Caption = Commandwindow.Caption & "×"
    
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseLeave
End Sub

Private Sub Label3_Change()

    If Label3.Caption = "Command" Then
        TextCMD.Visible = True
        Quality.Visible = False
        Preset.Visible = False
        Text3.Visible = False
        Combo1.Visible = False
        Label4.Visible = True
    ElseIf Label3.Caption = "Copy" Then
        TextCMD.Visible = Not True
        Quality.Visible = False
        Preset.Visible = False
        Text3.Visible = False
        Combo1.Visible = False
        Label4.Visible = False
        
    Else
        TextCMD.Visible = Not True
        Quality.Visible = Not False
        Preset.Visible = Not False
        Text3.Visible = Not False
        Combo1.Visible = Not False
        Label4.Visible = True
    End If
    If Label3.Caption = "x265" Or Label3.Caption = "x264" Then
        If InStr(1, Quality.Caption, "crf", vbTextCompare) = 0 Then
            Quality.Caption = Quality.Caption & "(crf)"
        End If
        RateMode = "crf"
        Quality.Visible = True
        Label8.Visible = False
        Text6.Visible = False
        Text3.Visible = True
        Preset.Visible = True
        Combo1.Visible = True
        Text3.Text = "23"
    ElseIf Label3.Caption = "Command" Then
        Quality.Visible = False
        Label8.Visible = False
        Text6.Visible = False
        Text3.Visible = False
        Preset.Visible = False
        Combo1.Visible = False
    ElseIf Label3.Caption = "Copy" Then
        Quality.Visible = False
        Label8.Visible = False
        Text6.Visible = False
        Text3.Visible = False
        Preset.Visible = False
        Combo1.Visible = False
        TextCMD.Visible = False
    Else
        RateMode = "VBR"
        Label8.Visible = True
        Text3.Visible = False
        Quality.Visible = False
        Text6.Visible = True
        Preset.Visible = False
        Combo1.Visible = False
    End If
    MouseLeave
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
    Check1.Caption = GetTranslation("MainScreen", "reSize")
    Label6.Caption = GetTranslation("MainScreen", "width")
    Label7.Caption = GetTranslation("MainScreen", "heith")
    Label8.Caption = GetTranslation("MainScreen", "bitrate")
    Label10.Caption = GetTranslation("MainScreen", "auEncode")
    Label11.Caption = GetTranslation("MainScreen", "auBitrate")
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
    CommandCode.Caption = GetTranslation("Bottom", "code")
    CommandT.Caption = GetTranslation("Bottom", "terminal")
    CommandSave.Caption = GetTranslation("Bottom", "batch")
    CommandPlay.Caption = GetTranslation("Bottom", "play")
    Commandwindow.Caption = GetTranslation("Bottom", "window")
    CommandLog.Caption = GetTranslation("Bottom", "log")
    CommandAbout.Caption = GetTranslation("Bottom", "about")
    run.ToolTipText = GetTranslation("ToolTip", "run")
    CommandCode.ToolTipText = GetTranslation("ToolTip", "code")
    CommandT.ToolTipText = GetTranslation("ToolTip", "terminal")
    CommandSave.ToolTipText = GetTranslation("ToolTip", "batch")
    CommandPlay.ToolTipText = GetTranslation("ToolTip", "play")
    Commandwindow.ToolTipText = GetTranslation("ToolTip", "window")
    CommandLog.ToolTipText = GetTranslation("ToolTip", "log")
    CommandAbout.ToolTipText = GetTranslation("ToolTip", "about")
'************************ComBox******************************************
    Combo1.AddItem GetTranslation("ComBox", "speed"), 0
    Combo1.AddItem GetTranslation("ComBox", "medium"), 1
    Combo1.AddItem GetTranslation("ComBox", "quality"), 2
End Sub

Private Sub Label4_Click()
    PopupMenu MenuFormat
    MouseLeave
End Sub

Private Sub Label5_Change()
    Check1.Value = 1
End Sub

Private Sub Label5_Click()
    PopupMenu MenuSize
    MouseLeave
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label5.FontBold = True
    Label5.FontItalic = True
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
    WriteIniKey "MainScreen", "Language", "en_us", App.path & "\Config\config.ini"
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
    SelectEncoder = "libaom-av1"
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
    WriteIniKey "MainScreen", "Language", "zh_cn", App.path & "\Config\config.ini"
    MsgBox "软件重启后生效", vbInformation, "需要重启"
'    k = MsgBox("是否重启软件？", vbYesNo + vbQuestion, "需要重启")
'    If k = vbYes Then AppRestart
End Sub

Private Sub Menu1280Size_Click()
    VideoSize = "1280:-2"
    Text4.Text = "1280"
    Text5.Text = "Auto"
    Check1.Value = 1
End Sub

Private Sub Menu1920Size_Click()
    VideoSize = "1920:-2"
    Text4.Text = "1920"
    Text5.Text = "Auto"
    Check1.Value = 1
End Sub

Private Sub Menu2560Size_Click()
    VideoSize = "2560:-2"
    Text4.Text = "2560"
    Text5.Text = "Auto"
    Check1.Value = 1
End Sub

Private Sub Menu3200Size_Click()
    VideoSize = "3200:-2"
    Text4.Text = "3200"
    Text5.Text = "Auto"
    Check1.Value = 1
End Sub

Private Sub Menu320Size_Click()
    VideoSize = "320:-2"
    Text4.Text = "320"
    Text5.Text = "Auto"
    Check1.Value = 1
End Sub

Private Sub Menu3840Size_Click()
    VideoSize = "3840:-2"
    Text4.Text = "3840"
    Text5.Text = "Auto"
    Check1.Value = 1
End Sub

Private Sub Menu480Size_Click()
    VideoSize = "480:-2"
    Text4.Text = "480"
    Text5.Text = "Auto"
    Check1.Value = 1
End Sub

Private Sub Menu640Size_Click()
    VideoSize = "640:-2"
    Text4.Text = "640"
    Text5.Text = "Auto"
    Check1.Value = 1
End Sub

Private Sub Menu720Size_Click()
    VideoSize = "720:-2"
    Text4.Text = "720"
    Text5.Text = "Auto"
    Check1.Value = 1
End Sub

Private Sub Menu960Size_Click()
    VideoSize = "960:-2"
    Text4.Text = "960"
    Text5.Text = "Auto"
    Check1.Value = 1
End Sub


Private Sub MenuAbout_Click()
    frmAbout.Show
End Sub

Private Sub MenuCMDLine_Click()
    Label3.Caption = "Command"
End Sub

Private Sub MenuCopy_Click()
    SelectEncoder = "copy"
    Label3.Caption = "Copy"
End Sub

Private Sub MenuExportBatch_Click()
    Dim cmdstr$, SourceFile$, TargetFile$, Quot$, Spac$, CMDOPEN$, crf$, Bitrate$, VideoWidth$, VideoHeight$
    Dim cmdstr1, cmdstr2, cmdstr3, cmdstr4, cmdstr5, cmdstr6
        
       
    
    Quot = Chr(34): Spac = " "                          '双引号，空格
    SourceFile = Quot & Text1.Text & Quot
    TargetFile = Quot & Text2.Text & Quot
    '开头
    cmdstr1 = "ffmpeg -y "
    '输入文件
    cmdstr2 = " -i " & SourceFile
    '编码器、preset、码率
    Bitrate = Text6.Text
    Select Case Combo1.ListIndex
        Case 0
            myPreset = "veryfast"
        Case 1
            myPreset = "medium"
        Case 2
            myPreset = "veryslow"
    End Select
    If RateMode = "crf" And SelectEncoder <> "copy" Then
        cmdstr3 = " -c:v " & SelectEncoder & " -preset " & myPreset & " -crf " & crf
    ElseIf RateMode = "VBR" And SelectEncoder <> "copy" Then
        cmdstr3 = " -c:v " & SelectEncoder & " -b:v " & Bitrate
    ElseIf SelectEncoder = "copy" Then
        cmdstr3 = " -c:v " & SelectEncoder
    End If

    '分辨率
    If Text4.Text = "Auto" Then
        VideoWidth = "-2"
    Else
        VideoWidth = Text4.Text
    End If
    If Text5.Text = "Auto" Then
        VideoHeight = "-2"
    Else
        VideoHeight = Text5.Text
    End If
    If Check1.Value = 0 And SelectEncoder <> "copy" Then
        cmdstr4 = " -vf ""scale=trunc(iw/2)*2:trunc(ih/2)*2"" "
    ElseIf Check1.Value = 1 And SelectEncoder <> "copy" Then
        cmdstr4 = " -vf ""scale=" & VideoWidth & ":" & VideoHeight
    ElseIf SelectEncoder <> "copy" Then
        cmdstr4 = ""
    End If
    '音频编码
    If Combo2.ListIndex = 0 Then
        cmdstr5 = " -c:a copy "
    ElseIf Combo2.ListIndex = 1 Then
        cmdstr5 = " -c:a aac " & " -b:a " & Text7.Text
    End If
    '元数据映射
    cmdstr6 = " -map_chapters 0 -map_metadata 0 "
    '输出文件
    cmdstr6 = " -f " & SelectFormat & Spac & TargetFile
    cmdstr = cmdstr1 & cmdstr2 & cmdstr3 & cmdstr4 & cmdstr5 & cmdstr6
    
    
    
    Dim FileName As String
    CommonDialog2.Filter = "*.bat"
    CommonDialog2.ShowSave
    FileName = CommonDialog2.FileName & ".bat"
    If CommonDialog2.FileName = "" Then Exit Sub
   
    ' 创建新的批处理脚本并写入FFmpeg命令
    Open FileName For Output As #1
    Print #1, cmdstr
    Print #1, "pause"
    Close #1


End Sub

Private Sub MenuFFmpegPath_Click()
    BasicOptions.Show vbModal
End Sub

Private Sub MenuInput_Click()
    Command1_Click
End Sub

Private Sub MenuMKV_Click()
    SelectFormat = "matroska"
    Label4.Caption = SelectFormat
End Sub

Private Sub MenuMP4_Click()
    SelectFormat = "mp4"
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
    Dim SourceFile$, TargetFile$, Quot$, Spac$, CMDOPEN$, crf$, Bitrate$, VideoWidth$, VideoHeight$
    Dim cmdstr1, cmdstr2, cmdstr3, cmdstr4, cmdstr5, cmdstr6, cmdstr7, CMDlog
    Dim cmdstr$
    Dim ffmpeg$, logPath$
    logPath = Replace(Replace(App.path & "\logs", "\", "\\"), ":", "\:")
    CMDlog = "set FFREPORT=file=" & logPath & "\\%p-%t.log" & ":level=32 && "
    If GetIniKey("BasicOption", "ffmpeg", ConfigPath) = "path" Then
        ffmpeg = "ffmpeg"
    Else
        ffmpeg = """" & GetIniKey("BasicOption", "ffmpeg", ConfigPath) & """"
    End If
    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")
    With reg
        .Pattern = "^\d+(k|M)$" '匹配码率单位
        .IgnoreCase = False
        .Global = True
    End With
    
    '运行前检查
    If SizeNumberTest = False Then MsgBox GetTranslation("MsgBox", "wrongSizeNum"), vbCritical: Exit Sub
    If Text1.Text = "" Or Text2.Text = "" Then GoTo runErr
    If Text6.Visible = True Then
        If reg.Test(Text6.Text) = False Then MsgBox GetTranslation("MsgBox", "checkBitRate"), vbCritical: Exit Sub
    End If
    If Text7.Visible = True Then
        If reg.Test(Text7.Text) = False Then MsgBox GetTranslation("MsgBox", "checkBitRate"), vbCritical: Exit Sub
    End If
    If Text3.Text = "" Then
        crf = "23"
    Else
        crf = Text3.Text
    End If
    If SelectEncoder = "copy" And Check1.Value = 1 Then MsgBox GetTranslation("MsgBox", "cantResize"), vbCritical: Exit Sub
    
    
    Quot = Chr(34): Spac = " "                          '双引号，空格
    SourceFile = Quot & Text1.Text & Quot
    TargetFile = Quot & Text2.Text & Quot
    '开头
    If TerminalCancel Then
        cmdstr1 = "CMD /K  " & CMDlog & ffmpeg & " -y "
    Else
        cmdstr1 = "CMD /C  " & CMDlog & ffmpeg & " -y "
    End If
    '输入文件
    cmdstr2 = " -i " & SourceFile
    '编码器、preset、码率
    Bitrate = Text6.Text
    Select Case Combo1.ListIndex
        Case 0
            myPreset = "veryfast"
        Case 1
            myPreset = "medium"
        Case 2
            myPreset = "veryslow"
    End Select
    If RateMode = "crf" And SelectEncoder <> "copy" Then
        cmdstr3 = " -c:v " & SelectEncoder & " -preset " & myPreset & " -crf " & crf
    ElseIf RateMode = "VBR" And SelectEncoder <> "copy" Then
        cmdstr3 = " -c:v " & SelectEncoder & " -b:v " & Bitrate
    ElseIf SelectEncoder = "copy" Then
        cmdstr3 = " -c:v " & SelectEncoder
    End If

    '分辨率
    If Text4.Text = "Auto" Then
        VideoWidth = "-2"
    Else
        VideoWidth = Text4.Text
    End If
    If Text5.Text = "Auto" Then
        VideoHeight = "-2"
    Else
        VideoHeight = Text5.Text
    End If
    If Check1.Value = 0 And SelectEncoder <> "copy" Then
        cmdstr4 = " -vf ""scale=trunc(iw/2)*2:trunc(ih/2)*2"" "
    ElseIf Check1.Value = 1 And SelectEncoder <> "copy" Then
        cmdstr4 = " -vf ""scale=" & VideoWidth & ":" & VideoHeight & """"
    ElseIf SelectEncoder <> "copy" Then
        cmdstr4 = ""
    End If
    '音频编码
    If Combo2.ListIndex = 0 Then
        cmdstr5 = " -c:a copy "
    ElseIf Combo2.ListIndex = 1 Then
        cmdstr5 = " -c:a aac " & " -b:a " & Text7.Text
    End If
    '元数据映射
    cmdstr6 = " -map_chapters 0 -map_metadata 0 "
    '输出文件
    If SelectEncoder = "copy" Then
        cmdstr6 = Spac & TargetFile
    Else
        cmdstr6 = " -f " & SelectFormat & Spac & TargetFile
    End If
        '日志部分
    cmdstr7 = " -report " & Quot & App.path & "\logs\" & Format(Date, "yyyy-MM-dd") & "_" & Format(Time, "hh-mm-ss") & ".log" & Quot
    
    cmdstr = cmdstr1 & cmdstr2 & cmdstr3 & cmdstr4 & cmdstr5 & cmdstr6

    
   
    

    Shell cmdstr, vbNormalFocus
'    MsgBox cmdstr
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
    Label5.FontBold = False
    Label5.FontItalic = False
    EncoderOptions.FontBold = False
    EncoderOptions.FontItalic = False
End Sub

Private Function SizeNumberTest() As Boolean
    Dim regEx
    Set regEx = CreateObject("VBScript.RegExp")
    If Check1.Value = 1 Then
        With regEx
            .Global = True
            .Pattern = "^(10000|[1-9][0-9]{0,3}|Auto)$"
            .IgnoreCase = True
            If .Test(Text4.Text) = False Or .Test(Text5.Text) = False Then
'                MsgBox GetTranslation("MsgBox", "wrongSizeNum"), vbCritical
                SizeNumberTest = False
            ElseIf Not (Text4.Text = "Auto" And Text5.Text = "Auto") Then
                SizeNumberTest = True
            End If
        End With
    Else
        SizeNumberTest = True
    End If
    

End Function

Private Sub TextCMD_GotFocus()
    If TextCMD.Text = "CommandLine" Then
        TextCMD.SelStart = 0
        TextCMD.SelLength = Len(TextCMD.Text)
    End If
End Sub

Private Function GenerateCommandString() As String
    Dim SourceFile$, TargetFile$, Quot$, Spac$, CMDOPEN$, crf$, Bitrate$, VideoWidth$, VideoHeight$
    Dim cmdstr1, cmdstr2, cmdstr3, cmdstr4, cmdstr5, cmdstr6, cmdstr7, CMDlog
    Dim cmdstr$
    Dim ffmpeg$, logPath$
    logPath = Replace(Replace(App.path & "\logs", "\", "\\"), ":", "\:")
    CMDlog = "set FFREPORT=file=" & logPath & "\\%p-%t.log" & ":level=32 && "
    If GetIniKey("BasicOption", "ffmpeg", ConfigPath) = "path" Then
        ffmpeg = "ffmpeg"
    Else
        ffmpeg = """" & GetIniKey("BasicOption", "ffmpeg", ConfigPath) & """"
    End If
    Dim reg As Object
    Set reg = CreateObject("VBScript.RegExp")
    With reg
        .Pattern = "^\d+(k|M)$" '匹配码率单位
        .IgnoreCase = False
        .Global = True
    End With
    
    '运行前检查
    If SizeNumberTest = False Then MsgBox GetTranslation("MsgBox", "wrongSizeNum"), vbCritical: Exit Function
'    If Text1.Text = "" Or Text2.Text = "" Then GoTo runErr
    If Text6.Visible = True Then
        If reg.Test(Text6.Text) = False Then MsgBox GetTranslation("MsgBox", "checkBitRate"), vbCritical: Exit Function
    End If
    If Text7.Visible = True Then
        If reg.Test(Text7.Text) = False Then MsgBox GetTranslation("MsgBox", "checkBitRate"), vbCritical: Exit Function
    End If
    If Text3.Text = "" Then
        crf = "23"
    Else
        crf = Text3.Text
    End If
    If SelectEncoder = "copy" And Check1.Value = 1 Then MsgBox GetTranslation("MsgBox", "cantResize"), vbCritical: Exit Function
    
    
    Quot = Chr(34): Spac = " "                          '双引号，空格
    SourceFile = Quot & Text1.Text & Quot
    TargetFile = Quot & Text2.Text & Quot
    '开头
    If TerminalCancel Then
        cmdstr1 = "CMD /K  " & CMDlog & ffmpeg & " -y "
    Else
        cmdstr1 = "CMD /C  " & CMDlog & ffmpeg & " -y "
    End If
    '输入文件
    cmdstr2 = " -i " & SourceFile
    '编码器、preset、码率
    Bitrate = Text6.Text
    Select Case Combo1.ListIndex
        Case 0
            myPreset = "veryfast"
        Case 1
            myPreset = "medium"
        Case 2
            myPreset = "veryslow"
    End Select
    If RateMode = "crf" And SelectEncoder <> "copy" Then
        cmdstr3 = " -c:v " & SelectEncoder & " -preset " & myPreset & " -crf " & crf
    ElseIf RateMode = "VBR" And SelectEncoder <> "copy" Then
        cmdstr3 = " -c:v " & SelectEncoder & " -b:v " & Bitrate
    ElseIf SelectEncoder = "copy" Then
        cmdstr3 = " -c:v " & SelectEncoder
    End If

    '分辨率
    If Text4.Text = "Auto" Then
        VideoWidth = "-2"
    Else
        VideoWidth = Text4.Text
    End If
    If Text5.Text = "Auto" Then
        VideoHeight = "-2"
    Else
        VideoHeight = Text5.Text
    End If
    If Check1.Value = 0 And SelectEncoder <> "copy" Then
        cmdstr4 = " -vf ""scale=trunc(iw/2)*2:trunc(ih/2)*2"" "
    ElseIf Check1.Value = 1 And SelectEncoder <> "copy" Then
        cmdstr4 = " -vf ""scale=" & VideoWidth & ":" & VideoHeight & """"
    ElseIf SelectEncoder <> "copy" Then
        cmdstr4 = ""
    End If
    '音频编码
    If Combo2.ListIndex = 0 Then
        cmdstr5 = " -c:a copy "
    ElseIf Combo2.ListIndex = 1 Then
        cmdstr5 = " -c:a aac " & " -b:a " & Text7.Text
    End If
    '元数据映射
    cmdstr6 = " -map_chapters 0 -map_metadata 0 "
    '输出文件
    If SelectEncoder = "copy" Then
        cmdstr6 = Spac & TargetFile
    Else
        cmdstr6 = " -f " & SelectFormat & Spac & TargetFile
    End If
        '日志部分
    cmdstr7 = " -report " & Quot & App.path & "\logs\" & Format(Date, "yyyy-MM-dd") & "_" & Format(Time, "hh-mm-ss") & ".log" & Quot
    
    cmdstr = cmdstr1 & cmdstr2 & cmdstr3 & cmdstr4 & cmdstr5 & cmdstr6

    
   GenerateCommandString = cmdstr
    

'    Shell cmdstr, vbNormalFocus
'    MsgBox cmdstr
    Exit Function

End Function
