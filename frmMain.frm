VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00202020&
   Caption         =   "Form1"
   ClientHeight    =   4155
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   6390
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Turn visual update on/off"
      Height          =   645
      Left            =   4305
      TabIndex        =   1
      Top             =   1725
      Width           =   1995
   End
   Begin CPU_Stress.CPUView CPUView1 
      Height          =   3165
      Left            =   210
      TabIndex        =   0
      Top             =   180
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   5583
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
