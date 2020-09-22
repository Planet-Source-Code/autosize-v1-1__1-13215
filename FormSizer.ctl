VERSION 5.00
Begin VB.UserControl FormSizer 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   CanGetFocus     =   0   'False
   ClientHeight    =   465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   InvisibleAtRuntime=   -1  'True
   Picture         =   "FormSizer.ctx":0000
   ScaleHeight     =   465
   ScaleWidth      =   480
   ToolboxBitmap   =   "FormSizer.ctx":0802
   Begin VB.Timer timFormSize 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   660
      Top             =   60
   End
End
Attribute VB_Name = "FormSizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'********************************************************
'*** This is just a litle handy control to automaticly
'*** remember a window's size and position from one
'*** run to the other.
'***
'*** Usage:
'***  Just drop the control anywhere on your form.
'***  Set the .ProgramName and .WindowName properties
'***  to whatever fits the window the control is on
'***  (or the control can't remember a setting for
'***  every window and program).
'***
'***  Though it's not intended to be used in other
'***  containers than windows, I think it might work.
'***
'*** Copyright 2000 - Christian Wikander
'*** christian.wikander@home.se
'***
'*** If you enhance this code in any way you must send
'*** me a copy of your work!
'***
'*** Please send me any comments or simply tell me if
'*** you find this useful and I might submit some more
'*** code in the future.
'********************************************************
Option Explicit
Option Compare Text

'*** Variables to store the program and window name
'*** for registry access.
Private m_strProgramName As String
Private m_strWindowName As String
'*** Variable to store a link to the parent object.
'*** This is needed, because when the terminate event
'*** occurs, the link to the parent is already detatched.
Private m_objParent As Object
'*** Flag to check if Show event is called.
'*** Show event is only called when in VB design mode.
Private m_bolShowCalled As Boolean

Public Property Let ProgramName(strProgramName As String)
  m_strProgramName = strProgramName
End Property

Public Property Get ProgramName() As String
  ProgramName = m_strProgramName
End Property

Public Property Let WindowName(strWindowName As String)
  m_strWindowName = strWindowName
End Property

Public Property Get WindowName() As String
  WindowName = m_strWindowName
End Property

Private Sub timFormSize_Timer()
  timFormSize.Enabled = False
  DoEvents
  '*** Only resize in runtime mode!
  If Not m_bolShowCalled Then
    Set m_objParent = UserControl.Parent
    '*** Set the parents position and size according to the registry
    m_objParent.Left = GetSetting(m_strProgramName, m_strWindowName, "Left", m_objParent.Left)
    m_objParent.Top = GetSetting(m_strProgramName, m_strWindowName, "Top", m_objParent.Top)
    m_objParent.Width = GetSetting(m_strProgramName, m_strWindowName, "Width", m_objParent.Width)
    m_objParent.Height = GetSetting(m_strProgramName, m_strWindowName, "Height", m_objParent.Height)
    On Error Resume Next  '*** Can occur because of parent not being a window
    m_objParent.WindowState = GetSetting(m_strProgramName, m_strWindowName, "WindowState", m_objParent.WindowState)
  End If
End Sub

Private Sub UserControl_InitProperties()
  m_strProgramName = "Defualt"
  m_strWindowName = "Default"
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  '*** Read the properties
  m_strProgramName = PropBag.ReadProperty("ProgramName", "Default")
  m_strWindowName = PropBag.ReadProperty("WindowName", "Default")
  '*** Set the clock
  timFormSize.Enabled = True
End Sub

Private Sub UserControl_Resize()
  UserControl.Width = Screen.TwipsPerPixelX * 35
  UserControl.Height = Screen.TwipsPerPixelY * 33
End Sub

Private Sub UserControl_Show()
  '*** A call to show = VB design environment!
  m_bolShowCalled = True
End Sub

Private Sub UserControl_Terminate()
  '*** Only save window size in runtime mode!
  If Not m_bolShowCalled Then
    On Error Resume Next  '*** Can occur because of parent not being a window
    Call SaveSetting(m_strProgramName, m_strWindowName, "WindowState", m_objParent.WindowState)
    '*** Only touch the Left, Top, Width and Height values if window not minimized/maximized
    If m_objParent.WindowState <> 0 Then Exit Sub
    Call SaveSetting(m_strProgramName, m_strWindowName, "Left", m_objParent.Left)
    Call SaveSetting(m_strProgramName, m_strWindowName, "Top", m_objParent.Top)
    Call SaveSetting(m_strProgramName, m_strWindowName, "Width", m_objParent.Width)
    Call SaveSetting(m_strProgramName, m_strWindowName, "Height", m_objParent.Height)
  End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("ProgramName", m_strProgramName, "Default")
  Call PropBag.WriteProperty("WindowName", m_strWindowName, "Default")
End Sub
