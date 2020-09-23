VERSION 5.00
Begin VB.Form ElapseFrm 
   Caption         =   "Elapsed Time Demo"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3240
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   3240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Another Sample"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton ExitBtn 
      Caption         =   "Exit"
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Top             =   3960
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Elapsed Time"
      Height          =   2775
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3015
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Text            =   "Text2"
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Start Time"
         Height          =   285
         Left            =   270
         TabIndex        =   8
         Top             =   360
         Width           =   1635
      End
      Begin VB.Label Label2 
         Caption         =   "Current Time"
         Height          =   285
         Left            =   270
         TabIndex        =   7
         Top             =   1140
         Width           =   1635
      End
      Begin VB.Label Label3 
         Caption         =   "Elapsed Time"
         Height          =   285
         Left            =   270
         TabIndex        =   6
         Top             =   1920
         Width           =   1635
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1350
      Top             =   2910
   End
   Begin VB.CommandButton StopBtn 
      Caption         =   "Stop"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   2880
      Width           =   1100
   End
   Begin VB.CommandButton StartBtn 
      Caption         =   "Start"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2880
      Width           =   1100
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Herry Hariry Amin : h2arr@yahoo.com"
      Height          =   375
      Left            =   173
      TabIndex        =   10
      Top             =   3390
      Width           =   2895
   End
End
Attribute VB_Name = "ElapseFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Elapsed Time Demo

Private Sub Command1_Click()
    AnoSamplFrm.Show 1
End Sub

Private Sub ExitBtn_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Text1.Text = Time()
End Sub

Private Sub StartBtn_Click()
    Timer1.Enabled = True
    Text1.Text = Time()
End Sub

Private Sub StopBtn_Click()
    Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
    Text2.Text = Time()
    Text3.Text = ElapsedTime(Text1.Text, Text2.Text)
End Sub
