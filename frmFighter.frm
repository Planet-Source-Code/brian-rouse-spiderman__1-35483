VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFighter 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Brian's Spiderman   A BRouse Gaming Development"
   ClientHeight    =   6525
   ClientLeft      =   2040
   ClientTop       =   1860
   ClientWidth     =   6270
   Icon            =   "frmFighter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrRegeneration 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   5040
      Top             =   5040
   End
   Begin VB.Timer tmrPlayerRecover 
      Interval        =   500
      Left            =   5040
      Top             =   4560
   End
   Begin VB.Timer tmrComputerAI 
      Interval        =   500
      Left            =   5520
      Top             =   5040
   End
   Begin VB.PictureBox picPlayerPunch 
      Height          =   495
      Left            =   1560
      Picture         =   "frmFighter.frx":0442
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   18
      Top             =   2280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox picPlayerKick 
      Height          =   495
      Left            =   1320
      Picture         =   "frmFighter.frx":F6AC
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   17
      Top             =   2280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox picPlayerNone 
      Height          =   495
      Left            =   1200
      Picture         =   "frmFighter.frx":1F2EE
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   16
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picPlayerLoss 
      Height          =   735
      Left            =   840
      Picture         =   "frmFighter.frx":22A8A
      ScaleHeight     =   675
      ScaleWidth      =   1515
      TabIndex        =   15
      Top             =   2880
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox picPlayerHit 
      Height          =   495
      Left            =   720
      Picture         =   "frmFighter.frx":2A704
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   14
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picPlayerForward 
      Height          =   495
      Left            =   480
      Picture         =   "frmFighter.frx":2E12E
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   13
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picPlayerBack 
      Height          =   495
      Left            =   240
      Picture         =   "frmFighter.frx":35830
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   12
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picCompPunch 
      Height          =   495
      Left            =   4080
      Picture         =   "frmFighter.frx":3CF32
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   11
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picCompNone 
      Height          =   495
      Left            =   4320
      Picture         =   "frmFighter.frx":41AC0
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   10
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picCompLoss 
      Height          =   495
      Left            =   4560
      Picture         =   "frmFighter.frx":44352
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   9
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picCompKick 
      Height          =   495
      Left            =   4800
      Picture         =   "frmFighter.frx":46864
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   8
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picCompHit 
      Height          =   495
      Left            =   5040
      Picture         =   "frmFighter.frx":506BA
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   7
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picCompForward 
      Height          =   495
      Left            =   5160
      Picture         =   "frmFighter.frx":5203C
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   6
      Top             =   2280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox picCompBack 
      Height          =   495
      Left            =   5400
      Picture         =   "frmFighter.frx":549AE
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   5
      Top             =   2280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer tmrCompRecover 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5520
      Top             =   4560
   End
   Begin MSComctlLib.ProgressBar w 
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar l 
      Height          =   255
      Left            =   3960
      TabIndex        =   3
      Top             =   1440
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lblFighter 
      BackStyle       =   0  'Transparent
      Caption         =   "Spiderman"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   1920
      TabIndex        =   19
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label lblWinner 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1335
      Left            =   1080
      TabIndex        =   4
      Top             =   2280
      Width           =   4095
   End
   Begin VB.Image m 
      Height          =   1005
      Left            =   5205
      Picture         =   "frmFighter.frx":57320
      Top             =   3795
      Width           =   1320
   End
   Begin VB.Image f 
      Height          =   1470
      Left            =   2400
      Picture         =   "frmFighter.frx":58E6A
      Top             =   3795
      Width           =   1710
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      Height          =   1095
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   6015
   End
   Begin VB.Label lblCompName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Chameleon"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3960
      TabIndex        =   1
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label lblPlayerName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Spidey"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      Height          =   4215
      Left            =   120
      Top             =   2160
      Width           =   6015
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuNewGameItem 
         Caption         =   "&New Game"
         Begin VB.Menu mnuSinglePlayerItem 
            Caption         =   "&Single Player"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuTwoPlayerItem 
            Caption         =   "&Two - Player"
         End
      End
      Begin VB.Menu mnuFightItem 
         Caption         =   "&Fight!"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuDifficultyItem 
         Caption         =   "&Difficulty"
         Begin VB.Menu mnuEasyItem 
            Caption         =   "&Easy"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuMediumItem 
            Caption         =   "&Medium"
         End
         Begin VB.Menu mnuHardItem 
            Caption         =   "&Hard"
         End
      End
      Begin VB.Menu separator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuControlsItem 
         Caption         =   "&Controls"
      End
      Begin VB.Menu mnuCreditsItem 
         Caption         =   "C&redits"
      End
      Begin VB.Menu separator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExitItem 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmFighter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Spiderman fight game BRouseSoftware.Net

Option Explicit
Private a As Boolean, aa As Boolean
Private bb As Integer, b As Integer
Private i As Integer, i2 As Integer
Private aaa As String, aaaa As String

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If mnuTwoPlayerItem.Checked = True Then
    If lblFighter.ForeColor = vbGreen Then
        If KeyCode = vbKeyA Then
            m.Picture = picCompForward.Picture
    
            If m.Left - f.Left < 1000 Then
                Exit Sub
            Else
                m.ZOrder (1)
                m.Left = m.Left - 200
            End If
        ElseIf KeyCode = vbKeyS Then
                If m.Left + 300 >= 5400 Then
                    Exit Sub
                End If
            i = 1
            m.ZOrder (1)
            m.Picture = picCompBack.Picture
            m.Left = m.Left + 200
        ElseIf KeyCode = vbKeyG Then
            aa = True
            m.ZOrder (1)
            m.Picture = picCompPunch.Picture
        ElseIf KeyCode = vbKeyH Then
            m.ZOrder (1)
            aa = True
            m.Picture = picCompKick.Picture
        End If
    End If
End If

    If lblFighter.ForeColor = vbGreen Then
        If KeyCode = vbKeyRight Then
            f.Picture = picPlayerForward.Picture
    
            If m.Left - f.Left < 1000 Then
                Exit Sub
            Else
                f.ZOrder (1)
                f.Left = f.Left + 200
            End If
        ElseIf KeyCode = vbKeyLeft Then
                If f.Left - 300 < 100 Then
                    Exit Sub
                End If
            i2 = 1
            f.ZOrder (1)
            f.Picture = picPlayerBack.Picture
            f.Left = f.Left - 200
        ElseIf KeyCode = vbKeyControl Then
            a = True
            f.ZOrder (1)
            f.Picture = picPlayerPunch.Picture
        ElseIf KeyCode = vbKeyShift Then
            f.ZOrder (1)
            a = True
            f.Picture = picPlayerKick.Picture
        End If
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    i = 5
    i2 = 5
    
If mnuTwoPlayerItem.Checked = True Then
    If lblFighter.ForeColor = vbGreen Then
        If aa = True Then
            If m.Left - f.Left < 1000 Then
                bb = bb - i2
                f.Picture = picPlayerHit.Picture
                PlaySound App.Path & "\lion.wav"
                tmrCompRecover.Enabled = True
            End If
        End If
    
            If w.Value - 5 = 0 Then
                tmrPlayerRecover.Enabled = False
                w.Value = bb
                f.Picture = picPlayerLoss.Picture
                PlaySound App.Path & "\dead.wav"
                f.Top = f.Top + 500
                bb = 0
                w.Value = bb
                lblWinner.Caption = aaaa & " Wins!"
                tmrRegeneration.Enabled = False
                tmrComputerAI.Enabled = False
                lblFighter.ForeColor = vbRed
            Else
            On Error GoTo health3:
                w.Value = bb
health3:
    If Err.Number = 380 Then
            bb = 0
            w.Value = b
            tmrPlayerRecover.Enabled = False
            w.Value = bb
            f.Picture = picPlayerLoss.Picture
            PlaySound App.Path & "\dead.wav"
            f.Top = f.Top + 500
            lblWinner.Caption = aaaa & " Wins!"
            tmrRegeneration.Enabled = False
            tmrComputerAI.Enabled = False
            lblFighter.ForeColor = vbRed
        End If

            End If
        m.Picture = picCompNone.Picture
        aa = False
    End If
End If
'///////////////////////////////////////////////////////////////////
    If lblFighter.ForeColor = vbGreen Then
        If a = True Then
            If m.Left - f.Left < 1000 Then
                b = b - i
                m.Picture = picCompHit.Picture
                PlaySound App.Path & "\hit.wav"
                tmrCompRecover.Enabled = True
            End If
        End If
    
            If l.Value - 5 = 0 Then
                tmrCompRecover.Enabled = False
                l.Value = b
                m.Picture = picCompLoss.Picture
                PlaySound App.Path & "\dead.wav"
                m.Top = m.Top + 500
                b = 0
                l.Value = b
                lblWinner.Caption = aaa & " Wins!"
                tmrRegeneration.Enabled = False
                tmrComputerAI.Enabled = False
                lblFighter.ForeColor = vbRed
            Else
            On Error GoTo health2:
                l.Value = b
health2:
        If Err.Number = 380 Then
            b = 0
            l.Value = b
            tmrCompRecover.Enabled = False
            l.Value = b
            m.Picture = picCompLoss.Picture
            PlaySound App.Path & "\dead.wav"
            m.Top = m.Top + 500
            lblWinner.Caption = aaa & " Wins!"
            tmrRegeneration.Enabled = False
            tmrComputerAI.Enabled = False
            lblFighter.ForeColor = vbRed
        End If
            End If
        
        f.Picture = picPlayerNone.Picture
        a = False
    End If
End Sub

Private Sub Form_Load()
    PlaySound App.Path & "\music1.wav"
    Randomize
    
    aaa = "Spiderman"
    aaaa = "The Chameleon"
    lblPlayerName.Caption = aaa
    lblCompName.Caption = aaaa
    
    f.Left = 800
    m.Left = 4400
    f.Top = 3800
    m.Top = 3800
    
    bb = 100
    b = 100
    
    l.Value = b
    w.Value = bb
    
    m.Picture = picCompNone.Picture
    f.Picture = picPlayerNone.Picture
    
    lblWinner.Caption = ""
    
    lblFighter.ForeColor = vbRed
    
    mnuEasyItem.Checked = True
    mnuMediumItem.Checked = False
    mnuHardItem.Checked = False
    
    bb = 100
    b = 100
    
    w.Value = bb
    l.Value = b
    
    tmrComputerAI.Enabled = False
    tmrCompRecover.Enabled = False
    
    i2 = 5
    i = 5
End Sub

Private Sub mnuControlsItem_Click()
    frmControls.Show vbModal
End Sub

Private Sub mnuCreditsItem_Click()
    frmCredits.Show vbModal
End Sub

Private Sub mnuEasyItem_Click()
    mnuEasyItem.Checked = True
    mnuMediumItem.Checked = False
    mnuHardItem.Checked = False
End Sub

Private Sub mnuExitItem_Click()
    Unload frmControls
    Unload frmCredits
    Unload Me
End Sub

Private Sub mnuFightItem_Click()
    lblFighter.ForeColor = vbGreen
        
    If mnuSinglePlayerItem.Checked = True Then
        
        mnuDifficultyItem.Enabled = False
        tmrComputerAI.Enabled = True
        tmrCompRecover.Enabled = False
        
        Randomize
    
    f.Left = 800
    m.Left = 4400
    f.Top = 3700
    m.Top = 3700
    
    bb = 100
    b = 100
    
    l.Value = b
    w.Value = bb
    
    m.Picture = picCompNone.Picture
    f.Picture = picPlayerNone.Picture
    
    lblWinner.Caption = ""
    
    tmrRegeneration.Enabled = True
    
    bb = 100
    b = 100
    
    w.Value = bb
    l.Value = b
    
    tmrComputerAI.Enabled = True
    tmrCompRecover.Enabled = False
    
    i2 = 5
    i = 5
    
    mnuFightItem.Enabled = True
    mnuDifficultyItem.Enabled = True

    ElseIf mnuTwoPlayerItem.Checked = True Then
        mnuFightItem.Enabled = False
        mnuDifficultyItem.Enabled = False
        tmrComputerAI.Enabled = False
        tmrCompRecover.Enabled = False
        
    Randomize
    
    f.Left = 800
    m.Left = 4400
    f.Top = 3700
    m.Top = 3700
    
    bb = 100
    b = 100
    
    l.Value = b
    w.Value = bb
    
    m.Picture = picCompNone.Picture
    f.Picture = picPlayerNone.Picture
    
    lblWinner.Caption = ""
    
    bb = 100
    b = 100
    
    tmrRegeneration.Enabled = True
    
    w.Value = bb
    l.Value = b
    
    tmrComputerAI.Enabled = False
    tmrCompRecover.Enabled = False
    
    i2 = 5
    i = 5
    
    mnuFightItem.Enabled = True

    End If

End Sub

Private Sub mnuHardItem_Click()
    mnuEasyItem.Checked = False
    mnuMediumItem.Checked = False
    mnuHardItem.Checked = True
End Sub

Private Sub mnuMediumItem_Click()
    mnuEasyItem.Checked = False
    mnuMediumItem.Checked = True
    mnuHardItem.Checked = False
End Sub

Private Sub mnuSinglePlayerItem_Click()
    mnuSinglePlayerItem.Checked = True
    mnuTwoPlayerItem.Checked = False
    mnuDifficultyItem.Enabled = True
    
Do
    aaa = InputBox("Enter the player's name", "Fighter")
        If aaa = "" Then
            MsgBox ("Sorry, but that was an invalid name" & vbCrLf & "Please enter another name"), vbExclamation, "Spiderman BRouse Gaming!"
        End If
Loop While aaa = ""
Do
    aaaa = InputBox("Enter the computer's name", "Fighter")
        If aaaa = "" Then
            MsgBox ("Sorry, but that was an invalid name" & vbCrLf & "Please enter another name"), vbExclamation, "Spiderman BRouse Gaming!"
        End If
Loop While aaaa = ""

lblPlayerName.Caption = aaa
lblCompName.Caption = aaaa
End Sub

Private Sub mnuTwoPlayerItem_Click()
    mnuSinglePlayerItem.Checked = False
    mnuTwoPlayerItem.Checked = True
    mnuDifficultyItem.Enabled = False
    
Do
    aaa = InputBox("Enter the player's name", "Brian's Spiderman")
        If aaa = "" Then
            MsgBox ("Sorry, but that was an invalid name" & vbCrLf & "Please enter another name"), vbExclamation, "Brian's Spiderman"
        End If
Loop While aaa = ""
Do
    aaaa = InputBox("Enter the second player's name", "Fighter")
        If aaaa = "" Then
            MsgBox ("Sorry, but that was an invalid name" & vbCrLf & "Please enter another name"), vbExclamation, "WWW.BRouseSoftware.Net"
        End If
Loop While aaaa = ""

lblPlayerName.Caption = aaa
lblCompName.Caption = aaaa

End Sub

Private Sub tmrCompRecover_Timer()
    m.Picture = picCompNone.Picture
    tmrCompRecover.Enabled = False
End Sub

Private Sub tmrComputerAI_Timer()
 If lblFighter.ForeColor = vbGreen Then
'EASY/////////////////////////////////////////////////////////////////////////////////////
    If mnuEasyItem.Checked = True Then
        Dim intAction As Integer
        intAction = Int(7 * Rnd) + 1
        tmrComputerAI.Interval = 500
        i = 5
        
    'Go Back
        If intAction = 1 Or intAction = 2 Then
            i = 1
            m.Picture = picCompBack.Picture
            If m.Left + 200 > 5000 Then
                Exit Sub
            End If
            f.ZOrder (1)
            m.Left = m.Left + 200
    'Go Forward
        ElseIf intAction = 3 Or intAction = 4 Then
            m.Picture = picCompForward.Picture
            If m.Left - f.Left < 800 Then
                Exit Sub
            End If
            f.ZOrder (1)
            m.Left = m.Left - 200
        End If
    'Punch
        If intAction = 5 Or 6 Then
            If m.Left - f.Left < 900 Then
                m.ZOrder (1)
                m.Picture = picCompPunch.Picture
                bb = bb - i2
                f.Picture = picPlayerHit.Picture
                PlaySound App.Path & "\lion.wav"
                tmrPlayerRecover.Enabled = True
            End If
            
            If w.Value - 5 = 0 Then
                tmrPlayerRecover.Enabled = False
                w.Value = bb
                f.Picture = picPlayerLoss.Picture
                f.Top = f.Top + 500
                bb = 0
                w.Value = bb
                lblWinner.Caption = aaaa & " Wins!"
                tmrRegeneration.Enabled = False
                tmrComputerAI.Enabled = False
                lblFighter.ForeColor = vbRed
            Else
            On Error GoTo health:
                w.Value = bb
            End If
        End If
    'Kick
        If intAction = 7 Then
            If m.Left - f.Left < 900 Then
                m.ZOrder (1)
                m.Picture = picCompKick.Picture
                bb = bb - i2
                f.Picture = picPlayerHit.Picture
                tmrPlayerRecover.Enabled = True
            End If
            
            If w.Value - 5 = 0 Then
                tmrPlayerRecover.Enabled = False
                w.Value = bb
                f.Picture = picPlayerLoss.Picture
                f.Top = f.Top + 500
                bb = 0
                w.Value = bb
                lblWinner.Caption = aaaa & " Wins!"
                tmrRegeneration.Enabled = False
                tmrComputerAI.Enabled = False
                lblFighter.ForeColor = vbRed
            Else
            On Error GoTo health:
                w.Value = bb
            End If
        End If
    End If

'MEDIUM---------------------------------
        If mnuMediumItem.Checked = True Then
            intAction = Int(6 * Rnd) + 1
            tmrComputerAI.Interval = 100
            i = 5
     'Go Back
           
            If intAction = 1 Or intAction = 2 Then
                i = 1
                m.Picture = picCompBack.Picture
                    If m.Left + 200 > 5000 Then
                        Exit Sub
                    End If
                f.ZOrder (1)
                m.Picture = picCompForward.Picture
                m.Left = m.Left + 200
        'Go Forward
            m.Picture = picCompForward.Picture
            ElseIf intAction = 3 Or intAction = 4 Then
                If m.Left - f.Left < 800 Then
                    Exit Sub
                End If
                f.ZOrder (1)
                m.Picture = picCompForward.Picture
                m.Left = m.Left - 200
            End If
    'Punch
            If intAction = 5 Then
                If m.Left - f.Left < 900 Then
                    m.ZOrder (1)
                    m.Picture = picCompPunch.Picture
                    bb = bb - i2
                    f.Picture = picPlayerHit.Picture
                    PlaySound App.Path & "\lion.wav"
                    tmrPlayerRecover.Enabled = True
                End If
                
                If w.Value - 5 = 0 Then
                    tmrPlayerRecover.Enabled = False
                    w.Value = bb
                    f.Picture = picPlayerLoss.Picture
                    PlaySound App.Path & "\dead.wav"
                    f.Top = f.Top + 500
                    bb = 0
                    w.Value = bb
                    lblWinner.Caption = aaaa & " Wins!"
                    tmrRegeneration.Enabled = False
                    tmrComputerAI.Enabled = False
                    lblFighter.ForeColor = vbRed
                Else
                On Error GoTo health
                    w.Value = bb
                End If
            End If
    'Kick
            If intAction = 6 Then
                If m.Left - f.Left < 900 Then
                    m.ZOrder (1)
                    m.Picture = picCompKick.Picture
                    bb = bb - i2
                    f.Picture = picPlayerHit.Picture
                    tmrPlayerRecover.Enabled = True
                End If
                
                If w.Value - 5 = 0 Then
                    tmrPlayerRecover.Enabled = False
                    w.Value = bb
                    f.Picture = picPlayerLoss.Picture
                    f.Top = f.Top + 500
                    bb = 0
                    w.Value = bb
                    lblWinner.Caption = aaaa & " Wins!"
                    tmrRegeneration.Enabled = False
                    tmrComputerAI.Enabled = False
                    lblFighter.ForeColor = vbRed
                Else
                On Error GoTo health:
                    w.Value = bb
                End If
            End If
        End If
'HARD////////////////////////////////////////////////////////////////////////////////
    If mnuHardItem.Checked = True Then
            intAction = Int(8 * Rnd) + 1
            tmrComputerAI.Interval = 1
            i = 5
     'Go Back
        
        If intAction = 1 Then
            i = 1
            m.Picture = picCompBack.Picture
            If m.Left + 200 > 5000 Then
                Exit Sub
            End If
            f.ZOrder (1)
            m.Picture = picCompForward.Picture
            m.Left = m.Left + 200
    'Go Forward
        ElseIf intAction = 2 Or intAction = 3 Then
            m.Picture = picCompForward.Picture
            If m.Left - f.Left < 800 Then
                Exit Sub
            End If
            f.ZOrder (1)
            m.Left = m.Left - 200
        End If
    'Punch
        If intAction = 4 Or intAction = 5 Or intAction = 6 Then
            If m.Left - f.Left < 900 Then
                m.ZOrder (1)
                m.Picture = picCompPunch.Picture
                bb = bb - i2
                f.Picture = picPlayerHit.Picture
                PlaySound App.Path & "\lion.wav"
                tmrPlayerRecover.Enabled = True
            End If
            
            If w.Value - 5 = 0 Then
                tmrPlayerRecover.Enabled = False
                w.Value = bb
                f.Picture = picPlayerLoss.Picture
                PlaySound App.Path & "\dead.wav"
                f.Top = f.Top + 500
                bb = 0
                w.Value = bb
                lblWinner.Caption = aaaa & " Wins!"
                tmrRegeneration.Enabled = False
                tmrComputerAI.Enabled = False
                lblFighter.ForeColor = vbRed
            Else
            On Error GoTo health:
                w.Value = bb
health:
    If Err.Number = 380 Then
        bb = 0
        w.Value = bb
                tmrPlayerRecover.Enabled = False
                w.Value = bb
                f.Picture = picPlayerLoss.Picture
                PlaySound App.Path & "\dead.wav"
                f.Top = f.Top + 500
                lblWinner.Caption = aaaa & " Wins!"
                tmrRegeneration.Enabled = False
                tmrComputerAI.Enabled = False
                lblFighter.ForeColor = vbRed
        Resume Next
    End If
    
            End If
        End If
    'Kick
        If intAction = 7 Or intAction = 8 Then
            If m.Left - f.Left < 900 Then
                m.ZOrder (1)
                m.Picture = picCompKick.Picture
                bb = bb - i2
                f.Picture = picPlayerHit.Picture
                tmrPlayerRecover.Enabled = True
            End If
            
            If w.Value - 5 = 0 Then
                tmrPlayerRecover.Enabled = False
                w.Value = bb
                f.Picture = picPlayerLoss.Picture
                PlaySound App.Path & "\dead.wav"
                f.Top = f.Top + 500
                bb = 0
                w.Value = bb
                lblWinner.Caption = aaaa & " Wins!"
                tmrRegeneration.Enabled = False
                tmrComputerAI.Enabled = False
                lblFighter.ForeColor = vbRed
            Else
            
            On Error GoTo health:
                w.Value = bb
            End If
        End If
    End If
End If
End Sub

Private Sub tmrPlayerRecover_Timer()
    f.Picture = picPlayerNone.Picture
    tmrPlayerRecover.Enabled = False
End Sub

Private Sub tmrRegeneration_Timer()
    If w.Value < 100 Then
        bb = bb + 1
        w.Value = bb
    End If
    If l.Value < 100 Then
        b = b + 1
        l.Value = b
    End If
End Sub
