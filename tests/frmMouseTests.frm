VERSION 5.00
Begin VB.Form frmMouseTests
   Caption         =   "Mouse manager integration tests"
   ClientHeight    =   6600
   ClientWidth     =   7200
   StartUpPosition =   2
   ScaleMode       =   3
   Begin MouseTests.CheckBox TestCheck
      Height          =   400
      Left            =   150
      Top             =   2400
      Width           =   2700
      Caption         =   "First checkbox"
   End
   Begin MouseTests.CheckBox TestCheck2
      Height          =   400
      Left            =   3150
      Top             =   2400
      Width           =   2700
      Caption         =   "Second checkbox"
   End
   Begin MouseTests.CPUView TestCPU
      Height          =   2700
      Left            =   150
      Top             =   3150
      Width           =   3000
      AutoUpdate      =   0
   End
   Begin MouseTests.CPUView TestCPU2
      Height          =   2700
      Left            =   3450
      Top             =   3150
      Width           =   3000
      AutoUpdate      =   0
   End
   Begin MouseTests.Button TestButton
      Height          =   600
      Left            =   150
      Top             =   1500
      Width           =   2700
      Caption         =   "Test button"
   End
   Begin VB.PictureBox TargetA
      Height          =   1200
      Left            =   150
      Top             =   150
      Width           =   2700
      ScaleMode       =   3
      BorderStyle     =   0
   End
   Begin VB.PictureBox TargetB
      Height          =   1200
      Left            =   3150
      Top             =   150
      Width           =   2700
      ScaleMode       =   3
      BorderStyle     =   0
   End
End
Attribute VB_Name = "frmMouseTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ButtonClicks As Long
Public DisableOnClick As Boolean
Public CheckClicks As Long
Public Check2Clicks As Long
Public CheckAction As String
Public CPUClicks As Long
Public CPU2Clicks As Long
Public LastCPUIndex As Long
Public LastCPU2Index As Long
Public InvalidCPUClicks As Long
Public CPUAction As String

Private Sub TestButton_Click()
  ButtonClicks = ButtonClicks + 1
  If DisableOnClick Then TestButton.Enabled = False
End Sub

Private Sub TestCheck_Click()
  CheckClicks = CheckClicks + 1
  Select Case CheckAction
    Case "disable": TestCheck.Enabled = False
    Case "hide": TestCheck.Visible = False
  End Select
End Sub

Private Sub TestCheck2_Click()
  Check2Clicks = Check2Clicks + 1
End Sub

Private Sub TestCPU_Click(Index As Long)
  CPUClicks = CPUClicks + 1
  LastCPUIndex = Index
  If Index < 1 Or Index > TotalCores Then InvalidCPUClicks = InvalidCPUClicks + 1
  If CPUAction = "hide" Then TestCPU.Visible = False
End Sub

Private Sub TestCPU2_Click(Index As Long)
  CPU2Clicks = CPU2Clicks + 1
  LastCPU2Index = Index
  If Index < 1 Or Index > TotalCores Then InvalidCPUClicks = InvalidCPUClicks + 1
End Sub

Private Sub Form_Deactivate()
  MouseCancelAll mcDeactivated
End Sub

Private Sub Form_Unload(Cancel As Integer)
  MouseShutdown
End Sub
