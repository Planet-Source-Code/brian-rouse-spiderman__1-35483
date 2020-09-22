VERSION 5.00
Begin VB.Form frmSplashScreen 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Fighter"
   ClientHeight    =   4770
   ClientLeft      =   2055
   ClientTop       =   2175
   ClientWidth     =   4905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3840
      ScaleHeight     =   465
      ScaleWidth      =   585
      TabIndex        =   0
      Top             =   960
      Width           =   615
   End
   Begin VB.Timer tmrUnload 
      Interval        =   6000
      Left            =   3600
      Top             =   3720
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   "Spiderman"
      BeginProperty Font 
         Name            =   "Engravers MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   "BRouse Gaming"
      BeginProperty Font 
         Name            =   "Engravers MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   4320
      Width           =   3135
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000009&
      X1              =   2040
      X2              =   4320
      Y1              =   2280
      Y2              =   1440
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000009&
      X1              =   960
      X2              =   1560
      Y1              =   2520
      Y2              =   2400
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000009&
      X1              =   2040
      X2              =   4320
      Y1              =   2280
      Y2              =   1560
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000009&
      X1              =   960
      X2              =   1440
      Y1              =   2520
      Y2              =   2400
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      X1              =   3120
      X2              =   240
      Y1              =   1200
      Y2              =   360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      X1              =   3240
      X2              =   240
      Y1              =   1200
      Y2              =   480
   End
   Begin VB.Image Image1 
      Height          =   3315
      Left            =   0
      Picture         =   "frmSplashScreen.frx":0000
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   4260
   End
End
Attribute VB_Name = "frmSplashScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Load frmFighter
    Load frmControls
    Load frmCredits
End Sub

Private Sub tmrunload_Timer()
    frmFighter.Show
    Unload Me
End Sub

