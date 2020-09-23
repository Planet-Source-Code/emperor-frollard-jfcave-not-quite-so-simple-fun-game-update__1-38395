Attribute VB_Name = "MainCode"
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
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long

Public Type Obstacle                 ' Obstacle parameters
    X1 As Integer
    X2 As Integer
    Y1 As Integer
    Y2 As Integer
    Active As Boolean
End Type

Public Const Pi As Double = 3.14159265358979

Public LoseFlag As Boolean           ' true if wall hit
Public Ready As Boolean              'ready to start the round
Public intPlayerPosY(15) As Integer   ' the player's line

Public intTopPos(50) As Integer      ' the stalactites' positions
Public intBotPos(50) As Integer      ' the stalagmites' positions

Public Obs(2) As Obstacle            ' the obstacle paramaters

Public intCaveHeight As Integer      ' the height of the cave

Public intVector As Integer          ' the players vertical vector
Public intAccel As Integer           ' acceleration of gravity and 'up'
Public intResponsiveness As Integer  ' responsiveness of the ship
' bigger = less responsive
Public intMouseY As Integer          ' mouse y value
Public intSpeed As Long              ' horizontal speed (lower = faster)
Public intRand As Integer            ' random chance integer
Public intShrinkSpeed As Integer     ' speed @ which cave decreases
' smaller = faster
Public intShrinkJump As Integer      ' amount to jump (pixels) per smaller
Public intMaxRand As Integer         ' maxamum random jump by cave itself
Public intSmallest As Integer        ' Smallest possible cave height
Public intTimeTilObs As Integer      ' number of frames until an obstacle
' can be released
Public NextTick As Long
Public Score As Long                 ' score
Public intPlayerShipHeight As Integer ' "DrawWidth" of player's 'ship' line
Public intMaxWiggle As Integer
Public lngFrameNumber As Long         'frame number of game

Public sngWiggleSpeed As Single       'how fast it wiggles
Public boolDrawLine(4) As Boolean     'Whether or not to draw each of the lines - set in options
Public intControlMode As Integer      'Method of controlling the ship





'**************************** Initialize Variables

Public Sub initvars()

  Dim i As Integer              ' for loop counter

    sngWiggleSpeed = 0.7        ' wigglespeed - higher = faster
    intPlayerShipHeight = 7
    intTimeTilObs = 300
    intCaveHeight = 500
    intShrinkJump = 6           ' # of pixels to jump
    intShrinkSpeed = 60         ' 1/intShrinkSpeed chance that itll get smaller/frame
    intMaxRand = 15             ' max amount that cave will jump
    intMaxWiggle = 30            ' max pixels for tail to wiggle
    
    'restart variables
    Ready = False
    LoseFlag = False
    lngFrameNumber = 0
    Score = 0
    '*******************
    
    
    
    
    For i = 0 To 15             ' set playerline to center
        intPlayerPosY(i) = 300

    Next i

    For i = 0 To 50             'make the initial cave jagged/random
        intTopPos(i) = 50
        intBotPos(i) = 550
        If Int(Rnd * 2) = 0 Then       'stutter the wall down
            intTopPos(i) = intTopPos(i) + Int(Rnd * intMaxRand / 2)
          Else                            'stutter the wall up'NOT INT(RND...
            intTopPos(i) = intTopPos(i) - Int(Rnd * intMaxRand / 2)
        End If
        If Int(Rnd * 2) = 0 Then       'stutter the wall down
            intBotPos(i) = intBotPos(i) + Int(Rnd * intMaxRand / 2)
          Else                            'stutter the wall up'NOT INT(RND...
            intBotPos(i) = intBotPos(i) - Int(Rnd * intMaxRand / 2)
        End If

    Next i

    For i = 0 To 2     'set obstacles outside right side of window
        Obs(i).X1 = 801
        Obs(i).X2 = 801
        Obs(i).Active = False
    Next i

End Sub

'*************************** ScoreKeep
Public Sub AddScore()

  Dim Slope As Single
  Dim Multiplier As Single

    'Slope = Abs((frmMain.linePlayer(15).Y2 - frmMain.linePlayer(15).Y1) / (frmMain.linePlayer(15).X2 - frmMain.linePlayer(15).X1))
    'Slope = (Slope / 2) + 0.5
    'Score = Score + Int(Slope * (560 - intCaveHeight))

    If intControlMode = 0 Then
        Multiplier = 1
      Else 'NOT INTCONTROLMODE...
        Multiplier = 2
    End If

    Slope = Abs((intPlayerPosY(15) - intPlayerPosY(14)) / 16)
    Slope = (Slope / 2) + 0.5
    Score = Score + (Int(Slope * (500 - intCaveHeight)) * Multiplier)

End Sub

'**************************** Animate explosion
Public Sub BlowUp()

  Dim i As Integer     'for loop again
  Dim intExpFrame As Integer   'frame number of explosion

End Sub

'**************************** Make end elements visible
Public Sub ShowStats()

  Dim i As Integer    'YET AGAIN...counter
  Dim tempi As Integer

    frmMain.lblGameOver.Visible = True      'set "GAME OVER" to visible
    frmMain.lblScore.Visible = False        'hide the top scorebox
    frmMain.tmrScoreUpdate.Enabled = False  'stop updating the top scorebox

    'animate game over sign
    For i = 0 To 8
        tempi = i
        If i > 4 Then
            tempi = i + 1
        End If
        frmMain.lblGameOver.Caption = Left("GAME OVER", tempi)
        'SpeedLimiter
        Delay 200
    Next i

    frmMain.scoreborder.Visible = True        'show the border around the final score
    frmMain.lblFinalScore(0).Visible = True   'show final score itself
    frmMain.cmdExit.Visible = True          'show the quit button
    frmMain.cmdRestart.Visible = True
    Delay 500
    frmMain.lblFinalScore(1).Caption = Str(Score)   'set finalscore value
    frmMain.lblFinalScore(1).Visible = True         'Show the actual Score Number

End Sub

'**************************** Try to start an obstacle
Public Sub SetObs()

  Dim i As Integer              'for loop counter
  Dim intAv As Integer          'available position
  Dim intObsLen As Integer      'length of obstacle

    intObsLen = 0.2 * intCaveHeight

    If intTimeTilObs > 0 Then
        intTimeTilObs = intTimeTilObs - 1
    End If
    intAv = 3
    For i = 2 To 0 Step -1     'find out what the first available line is
        If Obs(i).Active = False Then
            intAv = i
        End If
    Next i
    If intAv <> 3 Then 'if theres an available line to start....
        If (Int(Rnd * 10)) = 0 And intTimeTilObs = 0 Then  'then start it
            Obs(intAv).Active = True           'this obstacle is locked and loaded
            intTimeTilObs = intTimeTilObs + 15 'wait minimum of 15 frames to start another
            ' ***set its position
            Obs(intAv).Y1 = Int(Rnd * (intBotPos(50) - intTopPos(50) - intObsLen)) + intTopPos(50)
            Obs(intAv).Y2 = Obs(intAv).Y1 + intObsLen
            Obs(intAv).X1 = 800
            Obs(intAv).X2 = 800

        End If
    End If

End Sub

'**************************** Shift all the values left by 1
Public Sub UpdateArrays()

  Dim i As Integer      ' "for" loop counter

    'remove the 'go' label
    If lngFrameNumber = 30 Then

        frmMain.lblGetReady.Visible = False
    End If

    For i = 0 To 49       'shift top and bottom values over by 1
        intTopPos(i) = intTopPos(i + 1)
        intBotPos(i) = intBotPos(i + 1)
    Next i

    For i = 0 To 14
        intPlayerPosY(i) = intPlayerPosY(i + 1)  'shift playerline values over by 1
    Next i

    intPlayerPosY(15) = intPlayerPosY(15) + intAccel
    For i = 0 To 2        'move obstacle left by 16 pixels..or 1 frame
        If Obs(i).Active Then
            Obs(i).X1 = Obs(i).X1 - 16
            Obs(i).X2 = Obs(i).X2 - 16
        End If
    Next i
    For i = 0 To 2        'move obstacle to start if at end and deactivate
        If Obs(i).X1 <= 0 Then
            Obs(i).Active = False
            Obs(i).X1 = 801
            Obs(i).X2 = 801
        End If
    Next i

End Sub

'**************************** Change end value so cave gets smaller
Public Sub ChangeEnd()

  'determine cave height

    intRand = Int(Rnd * intShrinkSpeed) + 1       'random chance
    If (intRand = 1) And (intCaveHeight > intSmallest) Then
        intCaveHeight = intCaveHeight - intShrinkJump             'shrink cave by smalljump
    End If
    'end determining cave height

    'randomly move top of cave
    If Int(Rnd * 2) = 0 Then       'move down
        If intTopPos(50) < 600 - intMaxRand - intCaveHeight Then
            intTopPos(50) = intTopPos(50) + Int(Rnd * intMaxRand / 2)
        End If
      Else                            'move up'NOT INT(RND...
        If intTopPos(50) > 15 + intMaxRand Then
            intTopPos(50) = intTopPos(50) - Int(Rnd * intMaxRand / 2)
        End If
    End If

    If Int(Rnd * 2 + 1) = 1 Then      'move down
        intBotPos(50) = intTopPos(50) + intCaveHeight + intMaxRand / 4
      Else 'NOT INT(RND...
        intBotPos(50) = intTopPos(50) + intCaveHeight - intMaxRand / 4
    End If

End Sub

'**************************** Check for MouseMovement
Public Sub ChekMouse()

  Dim intPY As Integer

    intPY = intPlayerPosY(15) 'good to know where the player actually is
    'accelerate the mouse based on where the mouse is
    If intControlMode = 0 Then
        intAccel = Int((intMouseY - intPY) / intResponsiveness)
      Else 'NOT INTCONTROLMODE...
        intAccel = Int(((600 - intMouseY) - intPY) / intResponsiveness)
    End If

End Sub

'**************************** Check for collisions
Public Sub CheckForCollisions()

  Dim i As Integer   'yet another for loop counter

    'see if they hit the top
    If intPlayerPosY(15) - (intPlayerShipHeight / 2) <= intTopPos(15) Then
        LoseFlag = True   'run the lose procedure after moving lines
    End If
    'see if they hit the bottom
    If intPlayerPosY(15) + (intPlayerShipHeight / 2) >= intBotPos(15) Then
        LoseFlag = True   'run the lose procedure after moving lines
    End If

    For i = 0 To 2   'see if they hit an obstacle
        If Obs(i).X1 = 240 Then  'if theres an obstacle at player's x pos
            If (intPlayerPosY(15) >= Obs(i).Y1) And (intPlayerPosY(15) <= Obs(i).Y2) Then
                LoseFlag = True    'run the lose procedure after moving lines
            End If
        End If
    Next i

End Sub

'**************************** Draw the Lines
Public Sub DrawLines()

  Dim i As Integer      ' "for" loop counter...AGAIN
  Dim intWidth As Integer
  Dim intColour As Integer
  Dim intWiggle As Integer
  Dim intNextWiggle As Integer

    frmMain.Cls
    frmMain.DrawWidth = 1
    For i = 0 To 49
        frmMain.Line ((i * 16), intTopPos(i))-(((i + 1) * 16), intTopPos(i + 1)), RGB(180, 180, 180)
        frmMain.Line ((i * 16), intBotPos(i))-(((i + 1) * 16), intBotPos(i + 1)), RGB(180, 180, 180)
    Next i

    For i = 0 To 14

        intWiggle = (intMaxWiggle * Sin((i + lngFrameNumber) / (Pi / sngWiggleSpeed)))  'determine wiggle position
        intWiggle = intWiggle * (14 - i) * 0.1                   'modify for intensity

        intNextWiggle = (intMaxWiggle * Sin((i + 1 + lngFrameNumber) / (Pi / sngWiggleSpeed))) 'determine wiggle position + 1

        intNextWiggle = intNextWiggle * (13 - i) * 0.1
        If i = 14 Then
            intNextWiggle = 0
        End If

        'Debug.Print lngFrameNumber; i; intWiggle

        intWidth = Round((i / 14) * intPlayerShipHeight, 0)
        intColour = Round((i / 14) * 255, 0)
        If intWidth < 1 Then
            intWidth = 1
        End If
        frmMain.DrawWidth = 4
        If boolDrawLine(0) Then
            'standard playerline
            frmMain.Line ((i * 16), (intPlayerPosY(i)))-(((i + 1) * 16), (intPlayerPosY(i + 1))), RGB(intColour, intColour, intColour)
        End If
        If boolDrawLine(1) Then
            'positive "fin"
            frmMain.Line ((i * 16), (intPlayerPosY(i) + intWiggle))-(((i + 1) * 16), (intPlayerPosY(i + 1))), RGB(intColour, intColour, intColour)
        End If
        If boolDrawLine(2) Then
            'negative "fin"
            frmMain.Line ((i * 16), (intPlayerPosY(i) - intWiggle))-(((i + 1) * 16), (intPlayerPosY(i + 1))), RGB(intColour, intColour, intColour)
        End If
        If boolDrawLine(3) Then
            'wiggletail playerline
            frmMain.Line ((i * 16), (intPlayerPosY(i) + intWiggle))-(((i + 1) * 16), (intPlayerPosY(i + 1) + intNextWiggle)), RGB(intColour, intColour, intColour)
        End If
        If boolDrawLine(4) Then
            'negative wiggletail playerline
            frmMain.Line ((i * 16), (intPlayerPosY(i) - intWiggle))-(((i + 1) * 16), (intPlayerPosY(i + 1) - intNextWiggle)), RGB(intColour, intColour, intColour)
        End If
    Next i

    For i = 0 To 2
        If Obs(i).Active Then
            frmMain.DrawWidth = 5
            frmMain.Line (Obs(i).X1, Obs(i).Y1)-(Obs(i).X2, Obs(i).Y2), RGB(200, 200, 200)
        End If
    Next i

End Sub

Public Sub CleanUp()    'delete arrays
    Erase intTopPos()
    Erase intBotPos()
    Erase Obs()
    Erase boolDrawLine()
End Sub


'delay counter
Public Sub Delay(Length As Long)

    Do Until GetTickCount >= NextTick
        DoEvents
    Loop: NextTick = GetTickCount + Length

End Sub

':) Ulli's VB Code Formatter V2.13.6 (8/27/2002 6:38:47 PM) 62 + 320 = 382 Lines
