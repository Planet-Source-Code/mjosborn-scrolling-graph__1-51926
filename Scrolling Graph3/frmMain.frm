VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9180
   ForeColor       =   &H00FFFF00&
   LinkTopic       =   "Form1"
   ScaleHeight     =   375
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   612
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Grid X/Y Steps"
      Height          =   1095
      Left            =   120
      TabIndex        =   9
      Top             =   3840
      Width           =   3735
      Begin VB.HScrollBar hsGridY 
         Height          =   255
         Left            =   960
         TabIndex        =   13
         Top             =   720
         Width           =   1935
      End
      Begin VB.HScrollBar hsGridX 
         Height          =   255
         Left            =   960
         TabIndex        =   10
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblXValue 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         Height          =   195
         Left            =   3000
         TabIndex        =   15
         Top             =   360
         Width           =   480
      End
      Begin VB.Label lblYValue 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         Height          =   195
         Left            =   3000
         TabIndex        =   14
         Top             =   720
         Width           =   480
      End
      Begin VB.Label lblY 
         AutoSize        =   -1  'True
         Caption         =   "Y Grid line:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   765
      End
      Begin VB.Label lblX 
         AutoSize        =   -1  'True
         Caption         =   "X Grid line:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   765
      End
   End
   Begin VB.Frame fraDrawStyle 
      Caption         =   "Draw Style"
      Height          =   615
      Left            =   3960
      TabIndex        =   7
      Top             =   3120
      Width           =   1335
      Begin VB.ComboBox cboDrawStyle 
         Height          =   315
         ItemData        =   "frmMain.frx":0000
         Left            =   120
         List            =   "frmMain.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame fraScrollTime 
      Caption         =   "Scrolling Time"
      Height          =   615
      Left            =   5400
      TabIndex        =   4
      Top             =   3120
      Width           =   3495
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Left            =   120
         Max             =   1000
         TabIndex        =   5
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label lblTime 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         Height          =   195
         Left            =   2760
         TabIndex        =   6
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Frame fraGraphScroller 
      Caption         =   "Graph Value"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Width           =   3735
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   120
         Max             =   300
         TabIndex        =   2
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label lblValue 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         Height          =   195
         Left            =   3000
         TabIndex        =   3
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Timer tmrGraph 
      Interval        =   10
      Left            =   4320
      Top             =   3840
   End
   Begin VB.PictureBox pGraph 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   120
      ScaleHeight     =   201
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   601
      TabIndex        =   0
      Top             =   0
      Width           =   9015
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Flow of program
'1) init the scrolling graph via making a call to InitGraph
'a. Variables of InitGraph:
'   The graph picturebox
'   The timer for the graph
'   The interval of the timer for the graph,
'   which can be set when calling initGraph, or using the SetGraphInterval function
'   X and Y grid size
'2) Set the graph value via making a call to SetGraphValue
'a. Variables of SetGraphValue:
'   The current value of the graph to plot

'3) Draw the graph via making a call to DrawGraph
Option Explicit

Private Sub cboDrawStyle_Click()
    modGraph.SetDrawIndex cboDrawStyle.ListIndex 'set the graph draw style, dots, bars, line
End Sub

Private Sub Form_Load()
    Call Init 'Init
End Sub

Private Sub Init()
    cboDrawStyle.ListIndex = 0 'set listindex to 0, dots
    Call InitGraph(tmrGraph, pGraph, 1000, 20, 20) 'init the graph
    HScroll1.Value = 50
    lblValue.Caption = HScroll1.Value 'display graph value
    lblTime.Caption = HScroll2.Value 'display timer interval
    HScroll2.Value = 1 'set interval to 1 ms
    hsGridX.Max = pGraph.Width
    hsGridY.Max = pGraph.Height
    lblXValue.Caption = hsGridX.Value 'x grid line step change value
    lblYValue.Caption = hsGridY.Value 'y grid line step change value
End Sub

Private Sub Form_Unload(Cancel As Integer)
'clean up
    Set pGraph = Nothing 'clean picbox from memory
    modGraph.ModCleanUp 'clean up mod pic box from memory
    Unload Me 'unload form from memory
End Sub

Private Sub HScroll1_Change()
    modGraph.SetGraphValue HScroll1.Value 'set the value of the graph
    lblValue.Caption = HScroll1.Value 'display graph value
End Sub

Private Sub HScroll1_Scroll()
    modGraph.SetGraphValue HScroll1.Value 'set the value of the graph
    lblValue.Caption = HScroll1.Value 'display graph value
End Sub

Private Sub HScroll2_Change()
    modGraph.SetTimerInterval HScroll2.Value 'set the timer interval of the graph
    lblTime.Caption = HScroll2.Value 'display timer interval
End Sub

Private Sub HScroll2_Scroll()
    modGraph.SetTimerInterval HScroll2.Value 'set the timer interval of the graph
    lblTime.Caption = HScroll2.Value 'display timer interval
End Sub

Private Sub hsGridX_Change()
    Call modGraph.SetXGridStep(hsGridX.Value)
    lblXValue.Caption = hsGridX.Value 'x grid line step change
    'hsGridY.Value = hsGridX.Value
End Sub

Private Sub hsGridX_Scroll()
    Call modGraph.SetXGridStep(hsGridX.Value)
    lblXValue.Caption = hsGridX.Value 'x grid line step change
    'hsGridY.Value = hsGridX.Value
End Sub

Private Sub hsGridY_Change()
    Call modGraph.SetYGridStep(hsGridY.Value)
    lblYValue.Caption = hsGridY.Value 'x grid line step change
End Sub

Private Sub hsGridY_Scroll()
    Call modGraph.SetYGridStep(hsGridY.Value)
    lblYValue.Caption = hsGridY.Value 'x grid line step change
End Sub

Private Sub tmrGraph_Timer()
    Call DrawGraph 'draw the graph
End Sub

