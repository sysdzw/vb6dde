VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ScaleHeight     =   1215
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "test"
      Top             =   240
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.PictureBox picDDE 
      Height          =   375
      Left            =   1800
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    picDDE.LinkTopic = "app2|yuanyuan"
    picDDE.LinkMode = 2
    picDDE.LinkExecute Text1.Text
End Sub
