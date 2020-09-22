VERSION 5.00
Begin VB.UserControl ControlSizer 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   InvisibleAtRuntime=   -1  'True
   Picture         =   "ControlSizer.ctx":0000
   ScaleHeight     =   450
   ScaleWidth      =   480
   ToolboxBitmap   =   "ControlSizer.ctx":0802
   Begin VB.Timer timInit 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1260
      Top             =   240
   End
End
Attribute VB_Name = "ControlSizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'********************************************************
'*** This control helps you resize controls on a form
'*** accordning to rules that you set up at run-time.
'***
'*** Usage:
'***  Drop the control on a form.
'***  Use the AddRelation function to add relations
'***  between controls that ControlSizer seeks to uphold
'***  while your user resizes the window. (See the
'***  definition of the AddRelation function for details
'***  on the call)
'***  The relation properties points out what parts of
'***  the controls that should relate. (csrWidth and
'***  csrHeight makes the target controls Width and
'***  Height relate to the source control)
'***  If the source control is a parent to the target
'***  control you must set the ZeroBased property
'***  For more info, check the example in the example
'***  form.
'***
'*** The information on how to hook the WM_SIZE message
'*** and direct it to different instances of a control
'*** was found at Tim's VB5 tips and tricks,
'*** http://web.missouri.edu/~finaidtk/
'*** My code is based on the code presented there and
'*** may look the same at a glance, but is re-written
'*** (hey, Tim's didn't even work right from the box)
'*** Thanks to Tim for his exelent input on how to
'*** solve this problem though!
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

Public Enum csRelation
  csrLeft = 0
  csrRight = 1
  csrHCenter = 2  '*** Horizontal center
  csrWidth = 3    '*** Not available for source control
  csrTop = 4
  csrBottom = 5
  csrVCenter = 6  '*** Vertical center
  csrHeight = 7   '*** Not available for source control
End Enum

Private Type Relation
  objSourceObject As Object         '*** Source object
  objTargetObject As Object         '*** Target object (to be modified)
  csrSourceRelation As csRelation   '*** Source relation point
  csrTargetRelation As csRelation   '*** Target relation point
  lngValue As Long                  '*** Diference in twips
  bolZeroBased As Boolean           '*** Is the source target relation zero based?
                                    '*** ie control in a form, in a frame etc.
End Type

Private m_intInstance As Integer
'*** Flag to check if Show event is called.
'*** Show event is only called when in VB design mode.
Private m_bolShowCalled As Boolean
'*** Array to store relations in (array element 0 isn't used)
Private m_relRelations() As Relation

Friend Sub ParentResized(ByVal wMsg As Long)
  Static ParentWidth As Long, ParentHeight As Long
  Dim intI As Integer
  Dim lngNewPos As Long
  
  If wMsg = WM_CLOSE Then Call UnHookWindow(m_intInstance)
  If ParentWidth <> UserControl.Parent.Width Or ParentHeight <> UserControl.Parent.Height Then
    On Error Resume Next
    For intI = 1 To UBound(m_relRelations)
      With m_relRelations(intI)
        '*** Determin the new position of the target object
        Select Case .csrSourceRelation
          Case csrLeft
            If .bolZeroBased Then
              lngNewPos = .lngValue
            Else
              lngNewPos = .objSourceObject.Left + .lngValue
            End If
          Case csrRight
            If .bolZeroBased Then
              lngNewPos = .objSourceObject.Width + .lngValue
            Else
              lngNewPos = .objSourceObject.Left + .objSourceObject.Width + .lngValue
            End If
          Case csrHCenter
            If .bolZeroBased Then
              lngNewPos = .objSourceObject.Width \ 2 + .lngValue
            Else
              lngNewPos = (.objSourceObject.Left + .objSourceObject.Width) \ 2 + .lngValue
            End If
          Case csrTop
            If .bolZeroBased Then
              lngNewPos = .lngValue
            Else
              lngNewPos = .objSourceObject.Top + .lngValue
            End If
          Case csrBottom
            If .bolZeroBased Then
              lngNewPos = .objSourceObject.Height + .lngValue
            Else
              lngNewPos = .objSourceObject.Top + .objSourceObject.Height + .lngValue
            End If
          Case csrVCenter
            If .bolZeroBased Then
              lngNewPos = .objSourceObject.Height \ 2 + .lngValue
            Else
              lngNewPos = (.objSourceObject.Top + .objSourceObject.Height) \ 2 + .lngValue
            End If
        End Select
        '*** Set the new position of the target object
        Select Case .csrTargetRelation
          Case csrLeft
            .objTargetObject.Left = lngNewPos
          Case csrRight
            .objTargetObject.Left = lngNewPos - .objTargetObject.Width
          Case csrHCenter
            .objTargetObject.Left = lngNewPos - .objTargetObject.Width \ 2
          Case csrWidth
            .objTargetObject.Width = lngNewPos - .objTargetObject.Left
          Case csrTop
            .objTargetObject.Top = lngNewPos
          Case csrBottom
            .objTargetObject.Top = lngNewPos - .objTargetObject.Height
          Case csrVCenter
            .objTargetObject.Top = lngNewPos - .objTargetObject.Height \ 2
          Case csrHeight
            .objTargetObject.Height = lngNewPos - .objTargetObject.Top
        End Select
      End With
    Next intI
    On Error GoTo 0
  End If
  ParentWidth = UserControl.Parent.Width
  ParentHeight = UserControl.Parent.Height
End Sub

Private Sub timInit_Timer()
  Dim objParent As Object
  
  '*** Only run once
  timInit.Enabled = False
  DoEvents
  '*** Only resize in runtime mode!
  If Not m_bolShowCalled Then
    '*** Find parent window
    Set objParent = UserControl.Parent
    '*** Loop until an error occurs
    On Error GoTo timInit_Timer_WindowFound
    Do While True
      Set objParent = objParent.Parent
    Loop
timInit_Timer_WindowFound_Continue:
    '*** Remove error handling
    On Error GoTo 0
    m_intInstance = AddInstance(ObjPtr(Me), objParent.hwnd)
  End If
  GoTo timInit_Timer_End
timInit_Timer_WindowFound:
  Resume timInit_Timer_WindowFound_Continue
timInit_Timer_End:
End Sub

Private Sub UserControl_Initialize()
  m_bolShowCalled = False
  timInit.Enabled = True
End Sub

Private Sub UserControl_Resize()
  UserControl.Width = Screen.TwipsPerPixelX * 36
  UserControl.Height = Screen.TwipsPerPixelY * 33
End Sub

Private Sub UserControl_Show()
  m_bolShowCalled = True
End Sub

Public Sub ClearRelations()
  ReDim m_relRelations(0)
End Sub

'********************************************************
'*** objSource      Source object (not modified)
'*** relSource      Relation point on the source object
'*** objTarget      Target object (modified)
'*** relTarget      Relation point on the target object
'*** lngValue       Difference in twips between the
'***                related points. Possitive values
'***                means below or right of target.
'*** bolZeroBased   Set to true if the source control
'***                is the parent of the target control.
'***
'*** The function returns True if everything is ok,
'*** otherwise False.
'********************************************************
Public Function AddRelation(ByRef objSource As Object, ByVal relSource As csRelation, ByRef objTarget As Object, ByVal relTarget As csRelation, ByVal lngValue As Long, Optional ByVal bolZeroBased As Boolean = False) As Boolean
  Dim intI As Integer
  
  AddRelation = False
  '*** If an error occurs, simply jump out.
  On Error GoTo AddRelation_End
  '*** Check the input
  If (objSource Is Nothing) Or (objTarget Is Nothing) Then Exit Function
  Select Case relSource
    Case csrLeft, csrRight, csrHCenter, csrWidth   '*** No crossing relations!
      If relTarget = csrTop Or relTarget = csrBottom Or relTarget = csrVCenter Or relTarget = csrHeight Then Exit Function
    Case csrTop, csrBottom, csrVCenter, csrHeight  '*** No crossing relations!
      If relTarget = csrLeft Or relTarget = csrTop Or relTarget = csrHCenter Or relTarget = csrWidth Then Exit Function
  End Select
  On Error GoTo AddRelation_FirstRun
  intI = UBound(m_relRelations) + 1
AddRelation_FirstRun_Continue:
  On Error GoTo AddRelation_End
  ReDim Preserve m_relRelations(intI)
  With m_relRelations(intI)
    Set .objSourceObject = objSource
    .csrSourceRelation = relSource
    Set .objTargetObject = objTarget
    .csrTargetRelation = relTarget
    .lngValue = lngValue
    .bolZeroBased = bolZeroBased
  End With
  AddRelation = True
  GoTo AddRelation_End
AddRelation_FirstRun:
  intI = 1
  Resume AddRelation_FirstRun_Continue
AddRelation_End:
End Function
