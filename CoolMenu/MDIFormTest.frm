VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm MDIFormTest 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   5970
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10740
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList 
      Left            =   3360
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFormTest.frx":0000
            Key             =   ""
            Object.Tag             =   "&New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIFormTest.frx":015C
            Key             =   ""
            Object.Tag             =   "&Cut"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
   End
   Begin VB.Menu mnuWin 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWinSplit 
         Caption         =   "&Split"
      End
   End
End
Attribute VB_Name = "MDIFormTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''
''  MDItest Form
''
''  Copyright Olivier Martin 2000
''
''  martin.olivier@bigfoot.com
''
''  This form tests CoolMenu's functionality
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private Sub MDIForm_Load()
    
    Set objCoolMenu = New cCoolMenu

  Call objCoolMenu.Install(Me.hWnd, , ImageList)
  
  objCoolMenu.FontName Me.hWnd, "Tahoma"
  objCoolMenu.FontSize Me.hWnd, 8&
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
  Call objCoolMenu.Uninstall(Me.hWnd)
End Sub

Private Sub mnuNew_Click()
  Dim newForm As frmChild
  Set newForm = New frmChild
  newForm.Visible = True
  objCoolMenu.MDIChildMenu newForm.hWnd
End Sub

