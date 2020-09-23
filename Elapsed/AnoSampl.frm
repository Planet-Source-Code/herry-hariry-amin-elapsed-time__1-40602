VERSION 5.00
Begin VB.Form AnoSamplFrm 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton ExitBtn 
      Caption         =   "Exit"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton ProcessBtn 
      Caption         =   "Process"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   1920
      TabIndex        =   1
      Text            =   "05:00:00 PM"
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "07:00:00 AM"
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label ResultLbl 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "to"
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   240
      Width           =   255
   End
End
Attribute VB_Name = "AnoSamplFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ExitBtn_Click()
    Unload Me
End Sub

Private Sub ProcessBtn_Click()
     ResultLbl.Caption = ElapsedTime(Text1.Text, Text2.Text) & " (hh:mm:ss)"
End Sub
