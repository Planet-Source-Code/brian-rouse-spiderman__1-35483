VERSION 5.00
Begin VB.Form frmCredits 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Credits"
   ClientHeight    =   6510
   ClientLeft      =   2130
   ClientTop       =   2055
   ClientWidth     =   5460
   Icon            =   "frmCredits.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0FFC0&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Engravers MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000E&
      X1              =   3360
      X2              =   5400
      Y1              =   2160
      Y2              =   2040
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Brian A Rouse Author/Developer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   5400
      Width           =   5175
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H000000FF&
      Height          =   615
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   5280
      Width           =   5175
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000000FF&
      Height          =   615
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   4200
      Width           =   5175
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "All code originated by Brian A. Rouse"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   4320
      Width           =   5175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "All programming was done by Brian A. Rouse"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   5175
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      Height          =   615
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1440
      Width           =   5175
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      Height          =   855
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   5175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "All characters were created by Marvel Comics"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   5175
   End
   Begin VB.Image Image1 
      Height          =   2520
      Left            =   720
      Picture         =   "frmCredits.frx":0442
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   3600
   End
End
Attribute VB_Name = "frmCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()
    Unload Me
End Sub

