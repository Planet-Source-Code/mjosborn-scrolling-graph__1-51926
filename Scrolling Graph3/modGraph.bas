Attribute VB_Name = "modGraph"
'Setpixel api to draw dots
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long

Private Type POINTAPI
        X As Long
        Y As Long
End Type
'Vars
Private intXP As Integer ' xpoint used for line draw
Private intYP As Integer ' ypoint used for line draw
Private intGW As Integer 'graw width
Private intGH As Integer 'graph height
Private intGraphArray() As Integer 'array to hold graph values
Private mintDrawIndex As Integer 'draw style index, 0,1,2, dots, bars, line
Private mintValue As Integer 'graph value to plot
Private mintTimerInterval As Integer 'timer interval of the graph
Private mintStepX As Integer 'every x to point to grid
Private mintStepY As Integer 'every y point to grid
'objects
Private tmrGraphTimer As Timer 'timer of the graph
Private pGraphBox As PictureBox 'picture box to draw the graph on

Option Explicit

Public Sub InitGraph(tmr As Timer, picBox As PictureBox, tmrInterval As Integer, intStepX, intStepY)
'Init the graph, with graph width and height.
'Set the module global picture box and timer.
'Redim the grapharray which holds the values of the graph, to width of pgraph picture box.
'Set the timer interval here(or not), which can be changed in real time if the users wishes.
    Dim intI As Integer 'index of graph
    Set pGraphBox = picBox
    Set tmrGraphTimer = tmr
    intGW = pGraphBox.Width 'graph width max
    intGH = pGraphBox.Height  'graph height max
    Call SetXGridStep(intStepX)
    Call SetYGridStep(intStepY)
    Call GridGraph
    ReDim intGraphArray(intGW) 'resize the array
    mintTimerInterval = tmrInterval
    tmrGraphTimer.Interval = tmrInterval
    tmrGraphTimer.Enabled = True 'start the timer
End Sub

Private Sub GridGraph()
'Draw a grid on the picture box
    Dim intx As Integer
    Dim inty As Integer
    Dim Point As POINTAPI
    pGraphBox.ForeColor = &H8000000F
    'If GetXGridStep = 0 Or GetYGridStep = 0 Then Exit Sub 'break out if zero
    If GetXGridStep = 0 Then Call SetXGridStep(1)
    For intx = 0 To intGW Step GetXGridStep 'cycle x's steping each time
            'pGraphBox.Line (intx, inty)-(intx, intGH), vbRed 'draw x lines
            'pGraphBox.Line (intx, inty)-(intGW, inty), vbRed 'draw y lines
            
            Point.X = intx: Point.Y = 0 'Set the start-point's coordinates
            MoveToEx pGraphBox.hdc, intx, 0, Point 'Move the active point
            LineTo pGraphBox.hdc, intx, intGH 'Draw a line from the active point to the given point
            
    Next intx
    
    If GetYGridStep = 0 Then Call SetYGridStep(1)
    For inty = 0 To intGH Step GetYGridStep 'cycle y's steping each time
        
        Point.X = 0: Point.Y = inty 'Set the start-point's coordinates
        MoveToEx pGraphBox.hdc, 0, inty, Point 'Move the active point
        LineTo pGraphBox.hdc, intGW, inty 'Draw a line from the active point to the given point
    
    Next inty
End Sub

Private Sub FillGraphArray()
    Dim intI As Integer 'index of graph
    For intI = 0 To intGW 'Cycle the values of the graph array
        If (intI) <= (intGW - 1) Then 'index is less then graph max
            intGraphArray(intI) = intGraphArray(intI + 1) 'set each index to value next to it
        End If
        If intI = intGW Then 'index is at last item on the array
            intGraphArray(intI) = GetGraphValue 'add newest value to the array
        End If
    Next intI
End Sub

Public Sub DrawGraph()
    Dim inty As Double 'Y position
    Dim intx As Integer 'X position
    
    pGraphBox.Cls 'clear old picture
    GridGraph
    FillGraphArray 'refill graph array so we have a constant flow
    'Otherwise, the graph will only change when the graph array changes
    For intx = 0 To intGW  'cycle points to draw
        inty = (intGraphArray(intx)) / (intGH)   'get ratio of value to graph height
        inty = intGH - (intGH * inty)  'calc new y value
        If inty < 0 Then inty = 0 'if y is beyond top of graph, reset y to 0
        Call DrawType(GetDrawIndex, intx, inty)   'choose draw type, dots, bars, line
    Next intx
    intXP = 0 'reset x point for line drawtype
    intYP = 0 'reset y point for line drawtype
End Sub

Private Sub DrawType(ByVal intDS As Integer, intXpos As Integer, intYpos As Double)
    Dim lngRet As Long 'return of api setpixelv
    Dim lPoint As POINTAPI
    pGraphBox.ForeColor = &HFFFF00
    Select Case intDS
        Case 0 'dots
            lngRet = SetPixelV(pGraphBox.hdc, intXpos, intYpos, pGraphBox.ForeColor) 'plot dots
        Case 1 'bars
            lPoint.X = intXpos: lPoint.Y = intYpos 'Set the start-point's coordinates
            MoveToEx pGraphBox.hdc, intXpos, intYpos, lPoint 'Move the active point
            LineTo pGraphBox.hdc, intXpos, intGH 'Draw a line from the active point to the given point
            'pGraphBox.Line (intXpos, intYpos)-(intXpos, intGH), vbBlue 'plot bars
            lngRet = SetPixelV(pGraphBox.hdc, intXpos, intYpos, vbRed) 'plot dot at top of bar
        Case 2 'line
            If intXP = 0 Then
                intXP = intXpos 'start of line x
                intYP = intYpos 'start of line y
            End If
            lPoint.X = intXP: lPoint.Y = intYP
            MoveToEx pGraphBox.hdc, intXP, intYP, lPoint
            LineTo pGraphBox.hdc, intXpos, intYpos
            'picture box line method
            'pGraphBox.Line (intXP, intYP)-(intXpos, intYpos), vbBlue 'plot line
            intXP = intXpos 'set last x to first x
            intYP = intYpos 'set last y to first y
    End Select
End Sub

Private Function GetDrawIndex() As Integer
    GetDrawIndex = mintDrawIndex 'return draw style, 0,1,2,dots, bars, line
End Function

Public Sub SetDrawIndex(ByVal intDrawIndex As Integer)
    mintDrawIndex = intDrawIndex 'set draw style, 0,1,2,dots, bars, line
End Sub

Private Function GetGraphValue() As Integer
    GetGraphValue = mintValue 'return graph value
End Function

Public Sub SetGraphValue(ByVal intValue As Integer)
    mintValue = intValue 'set graph value
End Sub

Private Function GetTimerInterval() As Integer
    GetTimerInterval = mintTimerInterval 'return the timer interval
End Function

Public Sub SetTimerInterval(ByVal intInterval As Integer)
    mintTimerInterval = intInterval 'assign timer interval
    tmrGraphTimer.Interval = intInterval 'set timer interval
End Sub

Private Function GetXGridStep() As Integer
    GetXGridStep = mintStepX 'Return the X grid step amount
End Function

Public Sub SetXGridStep(ByVal intStepX As Integer)
    mintStepX = intStepX 'Set the x grid step amount
End Sub

Private Function GetYGridStep() As Integer
    GetYGridStep = mintStepY 'Return the Y grid step amount
End Function

Public Sub SetYGridStep(ByVal intStepY As Integer)
    mintStepY = intStepY 'Set the Y grid step amount
End Sub
Public Sub ModCleanUp()
    Set pGraphBox = Nothing 'clean up
End Sub

