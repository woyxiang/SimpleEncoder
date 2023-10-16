VERSION 5.00
Begin VB.Form BasicOptions 
   Caption         =   "BasicOptions"
   ClientHeight    =   4116
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6648
   LinkTopic       =   "Form1"
   ScaleHeight     =   4116
   ScaleWidth      =   6648
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.OptionButton Option2 
      Caption         =   "Check"
      Height          =   372
      Left            =   3240
      TabIndex        =   6
      Top             =   2880
      Value           =   -1  'True
      Width           =   2892
   End
   Begin VB.CommandButton CMDCancel 
      Caption         =   "Cancel"
      Height          =   492
      Left            =   3240
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
      Left            =   960
      TabIndex        =   3
      Top             =   1200
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
Dim ConfirmTip$
Private Sub Translate()
    ConfirmTip = GetTranslation("BasicOptions", "confirmTip")
    Option1.Caption = GetTranslation("BasicOptions", "envpath")
    Option2.Caption = GetTranslation("BasicOptions", "check")
    Label1.Caption = GetTranslation("BasicOptions", "path")
    CMDCancel.Caption = GetTranslation("BasicOptions", "Cancel")
    CMDApply.Caption = GetTranslation("BasicOptions", "Apply")
End Sub

Private Sub Form_Load()
    Translate
    Option1.ToolTipText = ConfirmTip
End Sub
