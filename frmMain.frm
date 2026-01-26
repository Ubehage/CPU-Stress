VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00202020&
   Caption         =   "Form1"
   ClientHeight    =   4155
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   5400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Turn visual update on/off"
      Height          =   645
      Left            =   1425
      TabIndex        =   1
      Top             =   2820
      Width           =   1995
   End
   Begin CPU_Stress.CPUView CPUView1 
      Height          =   2265
      Left            =   210
      TabIndex        =   0
      Top             =   180
      Width           =   3915
      _ExtentX        =   529
      _ExtentY        =   873
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Friend Sub SetForm()
  WindowOnTop Me.hWnd, True
  Me.Show
End Sub

Private Sub Command1_Click()
  CPUView1.AutoUpdate = Not CPUView1.AutoUpdate
End Sub
