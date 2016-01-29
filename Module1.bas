Attribute VB_Name = "Module1"
Option Explicit
'Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (ICCEX As tagInitCommonControlsEx) As Boolean
'Private Type tagInitCommonControlsEx
' lngSize As Long
' lngICC As Long
'End Type
'Private Const ICC_USEREX_CLASSES = &H200

Public Sub Main()
 On Error Resume Next
 'Dim ICCEX As tagInitCommonControlsEx
 'ICCEX.lngSize = LenB(ICCEX)
 'ICCEX.lngICC = &H200
 'InitCommonControlsEx ICCEX
 Call ComCtlsInitIDEStopProtection
 Call InitVisualStyles
 Load Form1
 Form1.Show vbModeless
End Sub
