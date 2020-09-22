VERSION 5.00
Begin VB.Form frmConfig 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "AutoBrowse Configuration"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdOkay 
      Caption         =   "Okay"
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Topics"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   4455
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Text            =   "Games"
         Top             =   1320
         Width           =   4215
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Text            =   "FBI"
         Top             =   960
         Width           =   4215
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Text            =   "Bin Laden"
         Top             =   600
         Width           =   4215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Text            =   "Visual Basic"
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AutoBrowse"
      BeginProperty Font 
         Name            =   "Vibrocentric"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   465
      Left            =   200
      TabIndex        =   8
      Top             =   200
      Width           =   4275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "AutoBrowse"
      BeginProperty Font 
         Name            =   "Vibrocentric"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   465
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   4275
   End
   Begin VB.Shape Shape1 
      Height          =   615
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
End
End Sub

Private Sub cmdOkay_Click()
SaveSetting "AutoBrowse", "Topics", "Topic1", Text1.Text
SaveSetting "AutoBrowse", "Topics", "Topic2", Text2.Text
SaveSetting "AutoBrowse", "Topics", "Topic3", Text3.Text
SaveSetting "AutoBrowse", "Topics", "Topic4", Text4.Text
End
End Sub

Private Sub Form_Load()
Text1.Text = GetSetting("AutoBrowse", "Topics", "Topic1")
Text2.Text = GetSetting("AutoBrowse", "Topics", "Topic2")
Text3.Text = GetSetting("AutoBrowse", "Topics", "Topic3")
Text4.Text = GetSetting("AutoBrowse", "Topics", "Topic4")
End Sub
