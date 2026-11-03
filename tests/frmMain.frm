VERSION 5.00
Begin VB.Form frmMain
   Caption         =   "Power callback stub (never shown)"
   ClientHeight    =   600
   ClientWidth     =   600
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Lets the harness link the production hibernation hook without workers.
Public Sub HibernationTrigger(ByVal Suspended As Boolean)
  RecordPowerEvent Suspended
End Sub
