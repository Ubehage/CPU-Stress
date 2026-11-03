Attribute VB_Name = "modButtonSupport"
Option Explicit

' Drawing dependencies for the real controls. No CPU-Stress startup/workers.
Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
Public Type POINTAPI
  X As Long
  Y As Long
End Type
Global Const COLOR_CONTROLS As Long = 2763306
Global Const COLOR_BACKGROUND_DISABLED As Long = 5263440
Global Const COLOR_TEXT As Long = 14737632
Global Const COLOR_TEXT_DISABLED As Long = 7895160
Global Const COLOR_BUTTON_HOVER As Long = 3684408
Global Const COLOR_BUTTON_PRESSED As Long = 3289650
Global Const COLOR_GREEN As Long = 5023791
Global Const COLOR_GREEN_HOVER As Long = 6339651
Global Const COLOR_GREEN_PRESSED As Long = 4033061
Global Const COLOR_GREEN_DISABLED As Long = 6455130
Global Const COLOR_TEXT_ONGREEN As Long = 15463654
Global Const COLOR_TEXT_DISABLED_ONGREEN As Long = 13355947
Global Const COLOR_RED As Long = 4539862
Global Const COLOR_RED_HOVER As Long = 6513642
Global Const COLOR_RED_PRESSED As Long = 3223992
Global Const COLOR_RED_DISABLED As Long = 6776730
Global Const COLOR_TEXT_ONRED As Long = 15395579
Global Const COLOR_TEXT_DISABLED_ONRED As Long = 10592542
Global Const COLOR_OUTLINE As Long = 3815994
Global Const COLOR_OUTLINE_LIGHT As Long = 7368816
Global Const COLOR_TEXT_HOVER As Long = 15790320
Global Const COLOR_BACKGROUND As Long = 2105376
Global Const COLOR_YELLOW As Long = 4965861
Global Const COLOR_YELLOW_HOVER As Long = 6673645
Global Const COLOR_YELLOW_PRESSED As Long = 3710156
Global Const FONT_MAIN As String = "Segoe UI"
Global Const FONT_SECONDARY As String = "Consolas"
Global Const FONTSIZE_MAIN As Integer = 11
Global Const FONTSIZE_SECONDARY As Integer = 9
Global ChangedByCode As Boolean

Public Enum XorY_Enum
  xyX = 1
  xyY = 2
End Enum
