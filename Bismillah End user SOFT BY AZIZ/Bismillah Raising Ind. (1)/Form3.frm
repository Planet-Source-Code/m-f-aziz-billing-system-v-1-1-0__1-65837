VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Splash"
   ClientHeight    =   2640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5715
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   1440
      Top             =   2640
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000E&
      Caption         =   "Bismillah Raising Industry"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   720
      Width           =   3375
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000009&
      Caption         =   "%"
      Height          =   255
      Left            =   5160
      TabIndex        =   4
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000009&
      Caption         =   "1"
      Height          =   255
      Left            =   4800
      TabIndex        =   3
      Top             =   2160
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000080&
      Height          =   2415
      Left            =   120
      Top             =   120
      Width           =   5415
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000009&
      Caption         =   "Click to proceed"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      Caption         =   "(C) Copyrights AZM TECHNOLOGY GROUP"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1065
      TabIndex        =   1
      Top             =   1920
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "Computerized Bills System For 32 Bit Application"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   5175
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Timer1_Timer()

'ProgressBar1.Appearance = ccFlat

Label4.Caption = Label4.Caption + 1
If Label4.Caption = 101 Then

 Label2.Caption = "101"
 Timer1.Enabled = False
 
 Form1.Show vbModal
 
 
 Unload Me
 
 End If
 End Sub
 
