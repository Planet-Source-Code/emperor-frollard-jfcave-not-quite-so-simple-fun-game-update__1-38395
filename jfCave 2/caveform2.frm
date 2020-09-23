VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Frostys Cave Game"
   ClientHeight    =   9000
   ClientLeft      =   -195
   ClientTop       =   1155
   ClientWidth     =   12000
   Icon            =   "caveform2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRestart 
      Caption         =   "Restart"
      Height          =   375
      Left            =   10710
      TabIndex        =   6
      Top             =   4725
      Width           =   975
   End
   Begin VB.Timer tmrScoreUpdate 
      Interval        =   1
      Left            =   5400
      Top             =   3720
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   10710
      TabIndex        =   0
      Top             =   4305
      Width           =   975
   End
   Begin VB.Label lblGetReady 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CLICK HERE TO START"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   3263
      TabIndex        =   5
      Top             =   4283
      Width           =   5475
   End
   Begin VB.Label lblGameOver 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   144
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   8175
      Left            =   1365
      TabIndex        =   2
      Top             =   210
      Visible         =   0   'False
      Width           =   8895
   End
   Begin VB.Label lblFinalScore 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Index           =   1
      Left            =   6615
      TabIndex        =   4
      Top             =   4200
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Label lblFinalScore 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Final Score:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   390
      Index           =   0
      Left            =   5250
      TabIndex        =   3
      Top             =   4200
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Label lblScore 
      BackStyle       =   0  'Transparent
      Caption         =   "Score:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   9240
      TabIndex        =   1
      Top             =   0
      Width           =   2775
   End
   Begin VB.Shape scoreborder 
      BorderColor     =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   390
      Left            =   5145
      Top             =   4200
      Visible         =   0   'False
      Width           =   3210
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************
'*                  jfCave v Pi (3.14)                  *
'*  You have the rights to use/modify/destroy this code *
'*  or whatever, just please email me first, and        *
'*  possibly even put my name in your credits (if you   *
'*  learned anything from this code.  I take no         *
'*  responsability for blablabla you get the idea       *
'*  If you're curious about how anything was done       *
'*  (SIMPLE questions only please!) email me at         *
'*  drallorf4@hotmail.com                               *
'*                                                      *
'*  (C) 2001 By Jamie Frost a.k.a. Floppy Disk          *
'*                                                      *
'********************************************************

Option Explicit

Private Sub cmdRestart_Click()

    CleanUp
    Unload Me
    Load frmOptions
    frmOptions.Show


End Sub

Private Sub Form_Load()

    frmMain.Show
    cmdExit.Visible = False
    cmdRestart.Visible = False
    initvars   'initalize variables to starting values
    DrawLines  'draw all the lines how they should look when it starts
    GetReady   'Displays countdown from 3 and get user to position mouse (waits for click)
    MainLoop  'duh

End Sub

Private Sub GetReady()

  Dim i As Integer

    Do Until Ready
        DoEvents
    Loop

    For i = 4 To 1 Step -1
        lblGetReady.Caption = Trim(Str(i))
        Delay 250
    Next i
    lblGetReady.Caption = "GO!!!"

End Sub

Private Sub cmdExit_Click()

    CleanUp        'remove arrays
    Unload Me      'unload this form
    End            'terminate program

End Sub

'if the mouse moves, capture where its at
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    intMouseY = Y

End Sub

Private Sub lblGetReady_Click()

    Ready = True

End Sub

'keep the "framerate" of the timer reasonable
Private Sub tmrScoreUpdate_Timer()

    frmMain.lblScore.Caption = "Score:" + Str(Score)
    tmrScoreUpdate.Interval = 250

End Sub

Private Sub MainLoop()

    Do Until (LoseFlag)
        DoEvents
        lngFrameNumber = lngFrameNumber + 1
        'SpeedLimiter
        Delay intSpeed

        ChekMouse            'move according to mouse pos
        UpdateArrays         'update arrays
        ChangeEnd            'change the end value to appropriate cave size
        SetObs               'Start an obstacle on its path
        DrawLines            'draw lines
        CheckForCollisions   'check for collisions
        AddScore             'add appropriate value to score.
    Loop
    If LoseFlag Then

        tmrScoreUpdate.Enabled = False

        'GAME OVER!!!

        BlowUp         'play animation at y coord of
        'player for spectacular finish
        ShowStats      'make stats elements visible
        frmMain.lblScore.Caption = "Score:" + Str(Score)

    End If

End Sub

':) Ulli's VB Code Formatter V2.13.6 (8/27/2002 6:38:47 PM) 16 + 93 = 109 Lines
