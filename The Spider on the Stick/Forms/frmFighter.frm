VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFighter 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "The Spider on the Stick"
   ClientHeight    =   7230
   ClientLeft      =   2040
   ClientTop       =   1860
   ClientWidth     =   9585
   Icon            =   "frmFighter.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmFighter.frx":0CCA
   ScaleHeight     =   7230
   ScaleWidth      =   9585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox picCompLoss 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   6720
      Picture         =   "frmFighter.frx":62B5
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   9
      Top             =   2280
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picCompPunch 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   6840
      Picture         =   "frmFighter.frx":6B6B
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   11
      Top             =   2280
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picCompKick 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   6960
      Picture         =   "frmFighter.frx":72F8
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   8
      Top             =   2280
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picCompHit 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   7080
      Picture         =   "frmFighter.frx":7B09
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   7
      Top             =   2280
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picCompBack 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   7200
      Picture         =   "frmFighter.frx":8444
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   5
      Top             =   2280
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerPunch 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   960
      Picture         =   "frmFighter.frx":8C42
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   18
      Top             =   2280
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerHit 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   840
      Picture         =   "frmFighter.frx":933B
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   14
      Top             =   2280
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerLoss 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   720
      Picture         =   "frmFighter.frx":9B8B
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   15
      Top             =   2280
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerKick 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   600
      Picture         =   "frmFighter.frx":A368
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   17
      Top             =   2280
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerBack 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   480
      Picture         =   "frmFighter.frx":AABF
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   12
      Top             =   2280
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Timer tmrRegeneration 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   4800
      Top             =   0
   End
   Begin VB.Timer tmrPlayerRecover 
      Interval        =   500
      Left            =   3480
      Top             =   0
   End
   Begin VB.Timer tmrComputerAI 
      Interval        =   500
      Left            =   5280
      Top             =   0
   End
   Begin VB.PictureBox picPlayerNone 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   360
      Picture         =   "frmFighter.frx":B215
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   16
      Top             =   2280
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerForward 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   240
      Picture         =   "frmFighter.frx":B952
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   13
      Top             =   2280
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picCompNone 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   7320
      Picture         =   "frmFighter.frx":C078
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   10
      Top             =   2280
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picCompForward 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   7440
      Picture         =   "frmFighter.frx":C8B7
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   6
      Top             =   2280
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Timer tmrCompRecover 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4080
      Top             =   0
   End
   Begin MSComctlLib.ProgressBar w 
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar l 
      Height          =   255
      Left            =   7320
      TabIndex        =   3
      Top             =   1200
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00000000&
      BorderColor     =   &H00004080&
      BorderWidth     =   3
      DrawMode        =   15  'Merge Pen Not
      FillColor       =   &H00004080&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   120
      Top             =   6800
      Width           =   9375
   End
   Begin VB.Label lblFighter 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "The Spider on the Stick"
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
      Height          =   555
      Left            =   2160
      TabIndex        =   19
      Top             =   120
      Width           =   5340
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
      ForeColor       =   &H0000FFFF&
      Height          =   1335
      Left            =   2880
      TabIndex        =   4
      Top             =   3600
      Width           =   4095
   End
   Begin VB.Image m 
      Height          =   1470
      Left            =   7560
      Picture         =   "frmFighter.frx":D0EB
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Image f 
      Height          =   1470
      Left            =   240
      Picture         =   "frmFighter.frx":D92A
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      Height          =   855
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   9375
   End
   Begin VB.Label lblCompName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Computer"
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
      Left            =   7320
      TabIndex        =   1
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label lblPlayerName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Player"
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
      Top             =   840
      Width           =   2055
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      Height          =   5415
      Left            =   120
      Top             =   1680
      Width           =   9375
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
'=============================================================================================================================
'
' Developed by Walter A. Narvasa
' jawoltze@edsamail.com.ph
'
' Walter A. Narvasa of
' WANCOM SYSTEMS
'
' Hey sir, Kindly rate this code, if you like it.
'
' READ THIS BEFORE USING THE CODE:
'
' You can study and view the source code for creating your
' own apps, but do not reproduce/release The Spider on the Stick fully
' or partially for any commercial and/or personal purposes. All
' rights of this product is related to it's author. Any violation
' of above conditions will be treated seriously and is punishable.
'
' I do not have full time to add complete explanation, read the help
' file (click Help->Contents) in The Spider on the Stick. Contact me for
' additional help/suggestions
'
' I recently inveted a technology for streaming audio, and is
' now looking promoters/investors to invest in a web-phone network
' project.
'
' VISIT MY WEBSITE : http://jawoltze.gq.nu/
'
'=============================================================================================================================

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
            Call sndPlaySound(App.Path & "\Sounds\Forward.wav", 1)
            If m.Left - f.Left < 1000 Then
                Exit Sub
            Else
                m.ZOrder (1)
                m.Left = m.Left - 200
            End If
        ElseIf KeyCode = vbKeyS Then
                If m.Left + 300 >= 5280 Then
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
            Call sndPlaySound(App.Path & "\Sounds\Punch.wav", 1)
        ElseIf KeyCode = vbKeyH Then
            m.ZOrder (1)
            aa = True
            m.Picture = picCompKick.Picture
            Call sndPlaySound(App.Path & "\Sounds\Kick.wav", 1)
        End If
    End If
End If

    If lblFighter.ForeColor = vbGreen Then
        If KeyCode = vbKeyRight Then
            f.Picture = picPlayerForward.Picture
            Call sndPlaySound(App.Path & "\Sounds\Forward.wav", 1)
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
            Call sndPlaySound(App.Path & "\Sounds\Punch.wav", 1)
        ElseIf KeyCode = vbKeyShift Then
            f.ZOrder (1)
            a = True
            f.Picture = picPlayerKick.Picture
            Call sndPlaySound(App.Path & "\Sounds\Kick.wav", 1)
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
                Call sndPlaySound(App.Path & "\Sounds\Hit.wav", 1)
                tmrCompRecover.Enabled = True
            End If
        End If
    
            If w.Value - 5 = 0 Then
                tmrPlayerRecover.Enabled = False
                w.Value = bb
                f.Picture = picPlayerLoss.Picture
                Call sndPlaySound(App.Path & "\Sounds\Win.wav", 1)
                'f.Top = f.Top + 500
                m.Left = f.Left + f.Width
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
            Call sndPlaySound(App.Path & "\Sounds\Win.wav", 1)
            'f.Top = f.Top + 500
            m.Left = f.Left + f.Width
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
                Call sndPlaySound(App.Path & "\Sounds\Hit.wav", 1)
                tmrCompRecover.Enabled = True
            End If
        End If
    
            If l.Value - 5 = 0 Then
                tmrCompRecover.Enabled = False
                l.Value = b
                m.Picture = picCompLoss.Picture
                Call sndPlaySound(App.Path & "\Sounds\Win.wav", 1)
                'm.Top = m.Top + 500
                f.Left = m.Left - 2000
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
            Call sndPlaySound(App.Path & "\Sounds\Win.wav", 1)
            'm.Top = m.Top + 500
            f.Left = m.Left - 2000
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
    
    Randomize
    
    aaa = "Player"
    aaaa = "Computer"
    lblPlayerName.Caption = aaa
    lblCompName.Caption = aaaa
    
    f.Left = 240
    m.Left = 7560
    f.Top = 5280
    m.Top = 5280
    
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
    
    f.Left = 240
    m.Left = 7560
    f.Top = 5280
    m.Top = 5280
    
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
    f.Top = 5280
    m.Top = 5280
    
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
            MsgBox ("Sorry, but that was an invalid name" & vbCrLf & "Please enter another name"), vbExclamation, "Fighter"
        End If
Loop While aaa = ""
Do
    aaaa = InputBox("Enter the computer's name", "Fighter")
        If aaaa = "" Then
            MsgBox ("Sorry, but that was an invalid name" & vbCrLf & "Please enter another name"), vbExclamation, "Fighter"
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
    aaa = InputBox("Enter the player's name", "Fighter")
        If aaa = "" Then
            MsgBox ("Sorry, but that was an invalid name" & vbCrLf & "Please enter another name"), vbExclamation, "Fighter"
        End If
Loop While aaa = ""
Do
    aaaa = InputBox("Enter the second player's name", "Fighter")
        If aaaa = "" Then
            MsgBox ("Sorry, but that was an invalid name" & vbCrLf & "Please enter another name"), vbExclamation, "Fighter"
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
            Call sndPlaySound(App.Path & "\Sounds\Forward.wav", 1)
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
                Call sndPlaySound(App.Path & "\Sounds\Punch.wav", 1)
                bb = bb - i2
                f.Picture = picPlayerHit.Picture
                Call sndPlaySound(App.Path & "\Sounds\Hit.wav", 1)
                tmrPlayerRecover.Enabled = True
            End If
            
            If w.Value - 5 = 0 Then
                tmrPlayerRecover.Enabled = False
                w.Value = bb
                f.Picture = picPlayerLoss.Picture
                Call sndPlaySound(App.Path & "\Sounds\Win.wav", 1)
                'f.Top = f.Top + 500
                m.Left = f.Left + f.Width
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
                Call sndPlaySound(App.Path & "\Sounds\Kick.wav", 1)
                bb = bb - i2
                f.Picture = picPlayerHit.Picture
                Call sndPlaySound(App.Path & "\Sounds\Hit.wav", 1)
                tmrPlayerRecover.Enabled = True
            End If
            
            If w.Value - 5 = 0 Then
                tmrPlayerRecover.Enabled = False
                w.Value = bb
                f.Picture = picPlayerLoss.Picture
                Call sndPlaySound(App.Path & "\Sounds\Win.wav", 1)
                'f.Top = f.Top + 500
                m.Left = f.Left + f.Width
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
                Call sndPlaySound(App.Path & "\Sounds\Forward.wav", 1)
                m.Left = m.Left + 200
        'Go Forward
            m.Picture = picCompForward.Picture
            Call sndPlaySound(App.Path & "\Sounds\Forward.wav", 1)
            ElseIf intAction = 3 Or intAction = 4 Then
                If m.Left - f.Left < 800 Then
                    Exit Sub
                End If
                f.ZOrder (1)
                m.Picture = picCompForward.Picture
                Call sndPlaySound(App.Path & "\Sounds\Forward.wav", 1)
                m.Left = m.Left - 200
            End If
    'Punch
            If intAction = 5 Then
                If m.Left - f.Left < 900 Then
                    m.ZOrder (1)
                    m.Picture = picCompPunch.Picture
                    Call sndPlaySound(App.Path & "\Sounds\Punch.wav", 1)
                    bb = bb - i2
                    f.Picture = picPlayerHit.Picture
                    Call sndPlaySound(App.Path & "\Sounds\Hit.wav", 1)
                    tmrPlayerRecover.Enabled = True
                End If
                
                If w.Value - 5 = 0 Then
                    tmrPlayerRecover.Enabled = False
                    w.Value = bb
                    f.Picture = picPlayerLoss.Picture
                    Call sndPlaySound(App.Path & "\Sounds\Win.wav", 1)
                    'f.Top = f.Top + 500
                    m.Left = f.Left + f.Width
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
                    Call sndPlaySound(App.Path & "\Sounds\Kick.wav", 1)
                    bb = bb - i2
                    f.Picture = picPlayerHit.Picture
                    Call sndPlaySound(App.Path & "\Sounds\Hit.wav", 1)
                    tmrPlayerRecover.Enabled = True
                End If
                
                If w.Value - 5 = 0 Then
                    tmrPlayerRecover.Enabled = False
                    w.Value = bb
                    f.Picture = picPlayerLoss.Picture
                    Call sndPlaySound(App.Path & "\Sounds\Win.wav", 1)
                    'f.Top = f.Top + 500
                    m.Left = f.Left + f.Width
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
            Call sndPlaySound(App.Path & "\Sounds\Forward.wav", 1)
            m.Left = m.Left + 200
    'Go Forward
        ElseIf intAction = 2 Or intAction = 3 Then
            m.Picture = picCompForward.Picture
            Call sndPlaySound(App.Path & "\Sounds\Forward.wav", 1)
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
                Call sndPlaySound(App.Path & "\Sounds\Punch.wav", 1)
                bb = bb - i2
                f.Picture = picPlayerHit.Picture
                Call sndPlaySound(App.Path & "\Sounds\Hit.wav", 1)
                tmrPlayerRecover.Enabled = True
            End If
            
            If w.Value - 5 = 0 Then
                tmrPlayerRecover.Enabled = False
                w.Value = bb
                f.Picture = picPlayerLoss.Picture
                Call sndPlaySound(App.Path & "\Sounds\Win.wav", 1)
                'f.Top = f.Top + 500
                m.Left = f.Left + f.Width
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
                Call sndPlaySound(App.Path & "\Sounds\Win.wav", 1)
                'f.Top = f.Top + 500
                m.Left = f.Left + f.Width
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
                Call sndPlaySound(App.Path & "\Sounds\Kick.wav", 1)
                bb = bb - i2
                f.Picture = picPlayerHit.Picture
                Call sndPlaySound(App.Path & "\Sounds\Hit.wav", 1)
                tmrPlayerRecover.Enabled = True
            End If
            
            If w.Value - 5 = 0 Then
                tmrPlayerRecover.Enabled = False
                w.Value = bb
                f.Picture = picPlayerLoss.Picture
                Call sndPlaySound(App.Path & "\Sounds\Win.wav", 1)
                'f.Top = f.Top + 500
                m.Left = f.Left + f.Width
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
