VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "jfCave Setup"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   4740
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmControls 
      Caption         =   "Ship Control"
      Height          =   1590
      Left            =   3255
      TabIndex        =   27
      Top             =   105
      Width           =   1380
      Begin VB.OptionButton optControl 
         Caption         =   "Classic N/A"
         Enabled         =   0   'False
         Height          =   225
         Index           =   4
         Left            =   105
         TabIndex        =   32
         Top             =   1155
         Width           =   1170
      End
      Begin VB.OptionButton optControl 
         Caption         =   "Classic N/A"
         Enabled         =   0   'False
         Height          =   225
         Index           =   3
         Left            =   105
         TabIndex        =   31
         Top             =   945
         Width           =   1170
      End
      Begin VB.OptionButton optControl 
         Caption         =   "Classic N/A"
         Enabled         =   0   'False
         Height          =   225
         Index           =   2
         Left            =   105
         TabIndex        =   30
         Top             =   735
         Width           =   1170
      End
      Begin VB.OptionButton optControl 
         Caption         =   "Goofy"
         Height          =   225
         Index           =   1
         Left            =   105
         TabIndex        =   29
         Top             =   525
         Width           =   1170
      End
      Begin VB.OptionButton optControl 
         Caption         =   "Normal"
         Height          =   225
         Index           =   0
         Left            =   105
         TabIndex        =   28
         Top             =   315
         Value           =   -1  'True
         Width           =   1170
      End
   End
   Begin VB.Frame frmShipType 
      Caption         =   "Ship Type"
      Height          =   1590
      Left            =   105
      TabIndex        =   15
      Top             =   1785
      Width           =   2955
      Begin VB.OptionButton optShip 
         Caption         =   "CUSTOM"
         Height          =   225
         Index           =   5
         Left            =   105
         TabIndex        =   26
         Top             =   1260
         Width           =   1170
      End
      Begin VB.CheckBox chkLineDraw 
         Caption         =   "WIGGLE DN"
         Enabled         =   0   'False
         Height          =   225
         Index           =   4
         Left            =   1575
         TabIndex        =   25
         Top             =   1155
         Value           =   1  'Checked
         Width           =   1275
      End
      Begin VB.CheckBox chkLineDraw 
         Caption         =   "WIGGLE UP"
         Enabled         =   0   'False
         Height          =   225
         Index           =   3
         Left            =   1575
         TabIndex        =   24
         Top             =   945
         Value           =   1  'Checked
         Width           =   1275
      End
      Begin VB.CheckBox chkLineDraw 
         Caption         =   "FIN DN"
         Enabled         =   0   'False
         Height          =   225
         Index           =   2
         Left            =   1575
         TabIndex        =   23
         Top             =   735
         Value           =   1  'Checked
         Width           =   1275
      End
      Begin VB.CheckBox chkLineDraw 
         Caption         =   "FIN UP"
         Enabled         =   0   'False
         Height          =   225
         Index           =   1
         Left            =   1575
         TabIndex        =   22
         Top             =   525
         Value           =   1  'Checked
         Width           =   1275
      End
      Begin VB.CheckBox chkLineDraw 
         Caption         =   "MIDDLE"
         Enabled         =   0   'False
         Height          =   225
         Index           =   0
         Left            =   1575
         TabIndex        =   21
         Top             =   315
         Value           =   1  'Checked
         Width           =   1275
      End
      Begin VB.OptionButton optShip 
         Caption         =   "SWIMMER"
         Height          =   225
         Index           =   4
         Left            =   105
         TabIndex        =   20
         Top             =   1050
         Width           =   1170
      End
      Begin VB.OptionButton optShip 
         Caption         =   "EVERYTHING"
         Height          =   225
         Index           =   3
         Left            =   105
         TabIndex        =   19
         Top             =   840
         Value           =   -1  'True
         Width           =   1380
      End
      Begin VB.OptionButton optShip 
         Caption         =   "FINS"
         Height          =   225
         Index           =   1
         Left            =   105
         TabIndex        =   18
         Top             =   420
         Width           =   1170
      End
      Begin VB.OptionButton optShip 
         Caption         =   "SIMPLE"
         Height          =   225
         Index           =   0
         Left            =   105
         TabIndex        =   17
         Top             =   210
         Width           =   1170
      End
      Begin VB.OptionButton optShip 
         Caption         =   "FLAPPY"
         Height          =   225
         Index           =   2
         Left            =   105
         TabIndex        =   16
         Top             =   630
         Width           =   1170
      End
   End
   Begin VB.CommandButton cmdContact 
      Caption         =   "Contact Me"
      Height          =   435
      Left            =   3255
      TabIndex        =   12
      Top             =   2415
      Width           =   1380
   End
   Begin VB.CommandButton cmdInstructions 
      Caption         =   "Instructions"
      Height          =   435
      Left            =   3255
      TabIndex        =   11
      Top             =   1890
      Width           =   1380
   End
   Begin VB.Frame frmDifficulty 
      Caption         =   "Difficulty"
      Height          =   1590
      Left            =   1680
      TabIndex        =   7
      Top             =   105
      Width           =   1380
      Begin VB.OptionButton optDiff 
         Caption         =   "HARD"
         Height          =   225
         Index           =   4
         Left            =   105
         TabIndex        =   14
         Top             =   1155
         Width           =   1065
      End
      Begin VB.OptionButton optDiff 
         Caption         =   "TOUGH"
         Height          =   225
         Index           =   3
         Left            =   105
         TabIndex        =   13
         Top             =   945
         Width           =   1065
      End
      Begin VB.OptionButton optDiff 
         Caption         =   "NORMAL"
         Height          =   225
         Index           =   2
         Left            =   105
         TabIndex        =   10
         Top             =   735
         Width           =   1065
      End
      Begin VB.OptionButton optDiff 
         Caption         =   "EASY"
         Height          =   225
         Index           =   1
         Left            =   105
         TabIndex        =   9
         Top             =   525
         Value           =   -1  'True
         Width           =   1065
      End
      Begin VB.OptionButton optDiff 
         Caption         =   "A BREEZE"
         Height          =   225
         Index           =   0
         Left            =   105
         TabIndex        =   8
         Top             =   315
         Width           =   1170
      End
   End
   Begin VB.Frame frmSpeed 
      Caption         =   "Speed"
      Height          =   1590
      Left            =   105
      TabIndex        =   1
      Top             =   105
      Width           =   1380
      Begin VB.OptionButton optSpd 
         Caption         =   "CRAWLING"
         Height          =   225
         Index           =   4
         Left            =   105
         TabIndex        =   6
         Top             =   1155
         Width           =   1170
      End
      Begin VB.OptionButton optSpd 
         Caption         =   "SLOW"
         Height          =   225
         Index           =   3
         Left            =   105
         TabIndex        =   5
         Top             =   945
         Width           =   1170
      End
      Begin VB.OptionButton optSpd 
         Caption         =   "NORMAL"
         Height          =   225
         Index           =   2
         Left            =   105
         TabIndex        =   4
         Top             =   735
         Width           =   1170
      End
      Begin VB.OptionButton optSpd 
         Caption         =   "FASTER"
         Height          =   225
         Index           =   1
         Left            =   105
         TabIndex        =   3
         Top             =   525
         Value           =   -1  'True
         Width           =   1170
      End
      Begin VB.OptionButton optSpd 
         Caption         =   "FASTEST"
         Height          =   225
         Index           =   0
         Left            =   105
         TabIndex        =   2
         Top             =   315
         Width           =   1170
      End
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "GO!"
      Height          =   435
      Left            =   3255
      TabIndex        =   0
      Top             =   2940
      Width           =   1380
   End
End
Attribute VB_Name = "frmOptions"
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

Private Sub cmdStart_click()

  Dim i As Integer
  Dim intTemp As Integer

    For i = 0 To 4   'check value of speed
        If optSpd(i) Then
            intTemp = i
        End If
    Next i

    Select Case intTemp
      Case 0
        intSpeed = 10
      Case 1
        intSpeed = 12
      Case 2
        intSpeed = 15
      Case 3
        intSpeed = 20
      Case 4
        intSpeed = 30
    End Select

    For i = 0 To 4    'check difficulty
        If optDiff(i) Then
            intTemp = i
        End If
    Next i

    Select Case intTemp 'set difficulty (responsiveness)
      Case 0
        intResponsiveness = 1
        intSmallest = 150
      Case 1
        intResponsiveness = 10    'accelerate by 1/10th of the dist to mouse
        intSmallest = 125         'smallest possible cave
      Case 2
        intResponsiveness = 17    'accelerate by 1/17th of the ...blabla
        intSmallest = 100         'smallest possible cave
      Case 3
        intResponsiveness = 23
        intSmallest = 90
      Case 4
        intResponsiveness = 30    'same thing again
        intSmallest = 80          'smallest possible cave
    End Select

    'ship type
    For i = 0 To 4
        boolDrawLine(i) = CBool(chkLineDraw(i).Value)
    Next i

    'ship controls
    For i = 0 To 4
        If optControl(i) Then
            intControlMode = i
        End If
    Next i

    Unload Me     'get rid of options dialog
    Load frmMain  'show the game, and get started

End Sub

Private Sub cmdInstructions_Click()

  Dim vbOkOnly As Integer

    vbOkOnly = MsgBox("How to play:" & vbCrLf & _
               "Basically, you fly thru a cave.  You get more and more" & vbCrLf & _
               "points as the cave gets smaller and smaller.  Also, you" & vbCrLf & _
               "get WAY less points if you arent actually `moving`...." & vbCrLf & _
               "...so fly like a teenager who just got his licence, or" & vbCrLf & _
               "                     NO HISCORE FOR YOU!" & vbCrLf & vbCrLf & _
               "Enjoy", vbOkOnly, "jfCave v.3.2")

End Sub

Private Sub cmdContact_Click()

  Dim vbOkOnly As Integer

    vbOkOnly = MsgBox( _
               "There are many ways you can get a hold of me" & vbCrLf & _
               "for well, anything, but I'll only tell you one:" & vbCrLf & _
               "drallorf4@hotmail.com" & vbCrLf & _
               "That is all.", vbOkOnly, "Me")

End Sub

Private Sub optShip_Click(Index As Integer)

  Dim i As Integer

    For i = 0 To 4
        chkLineDraw(i).Value = 0
        chkLineDraw(i).Enabled = False
    Next i

    Select Case Index
      Case 0
        chkLineDraw(0).Value = 1        'simple
      Case 1
        chkLineDraw(0).Value = 1        'simple
        chkLineDraw(1).Value = 1        'fin up
        chkLineDraw(2).Value = 1        'fin down
      Case 2
        chkLineDraw(0).Value = 1        'simple
        chkLineDraw(3).Value = 1        'wiggle up
        chkLineDraw(4).Value = 1        'wiggle down
      Case 3
        For i = 0 To 4
            chkLineDraw(i).Value = 1        'everything
        Next i
      Case 4
        chkLineDraw(3).Value = 1        'sinewave
      Case 5           'let the user select their own
        For i = 0 To 4
            chkLineDraw(i).Enabled = True
            chkLineDraw(i).Value = 1
        Next i
      Case Else
        MsgBox "error"
    End Select

End Sub

':) Ulli's VB Code Formatter V2.13.6 (8/27/2002 6:38:45 PM) 17 + 130 = 147 Lines
