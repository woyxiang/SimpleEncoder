VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Mainform 
   Caption         =   "SimpleEncoder"
   ClientHeight    =   4836
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6876
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4836
   ScaleWidth      =   6876
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   372
      Left            =   5160
      TabIndex        =   3
      Top             =   1440
      Width           =   612
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   372
      Left            =   5160
      TabIndex        =   2
      Top             =   600
      Width           =   612
   End
   Begin VB.TextBox Text2 
      Height          =   372
      Left            =   840
      TabIndex        =   1
      Top             =   1440
      Width           =   4212
   End
   Begin VB.TextBox Text1 
      Height          =   372
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   4212
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3240
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "output"
      Height          =   372
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   612
   End
   Begin VB.Label Label1 
      Caption         =   "input"
      Height          =   372
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   612
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

