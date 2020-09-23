VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCoolMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Only maps to the mCoolMenu functions and procedures
'''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Public Function Install(wndHandle As Long, Optional HelpObj As HelpCallBack, Optional ilHandle As Object) As Boolean
Install = mCoolMenu.Install(wndHandle, HelpObj, ilHandle)
End Function

Public Function FontName(hWnd As Long, Optional Value As Variant) As String
FontName = mCoolMenu.FontName(hWnd, Value)
End Function

Public Function FontSize(hWnd As Long, Optional Value As Variant) As Long
FontSize = mCoolMenu.FontSize(hWnd, Value)
End Function

Public Function ForeColor(hWnd As Long, Optional Value As Variant) As Long
ForeColor = mCoolMenu.ForeColor(hWnd, Value)
End Function

Public Function SelectColor(hWnd As Long, Optional Value As Variant) As Long
SelectColor = mCoolMenu.SelectColor(hWnd, Value)
End Function

Public Function ComplexChecks(hWnd As Long, Optional Value As Variant) As Boolean
ComplexChecks = mCoolMenu.ComplexChecks(hWnd, Value)
End Function

Public Function ColorEmbossed(hWnd As Long, Optional Value As Variant) As Boolean
ColorEmbossed = mCoolMenu.ColorEmbossed(hWnd, Value)
End Function

Public Function RightToLeft(hWnd As Long, Optional Value As Variant) As Boolean
RightToLeft = mCoolMenu.RightToLeft(hWnd, Value)
End Function

Public Function FullSelect(hWnd As Long, Optional Value As Variant) As Boolean
FullSelect = mCoolMenu.FullSelect(hWnd, Value)
End Function

Public Function Uninstall(wndHandle As Long) As Boolean
Uninstall = mCoolMenu.Uninstall(wndHandle)
End Function

Public Sub MDIChildMenu(hWnd As Long)
Call mCoolMenu.MDIChildMenu(hWnd)
End Sub

Private Sub Class_Initialize()
Set CoolMenu = Me
End Sub

Private Sub Class_Terminate()
Set CoolMenu = Nothing
End Sub


