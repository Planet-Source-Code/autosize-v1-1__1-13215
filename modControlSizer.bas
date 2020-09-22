Attribute VB_Name = "modControlSizer"
Option Explicit
Option Compare Text

Private Const WM_SIZE = &H5
Private Const GWL_WNDPROC = (-4&)
Private Const GWL_USERDATA = (-21&)
Public Const WM_CLOSE = &H10

'Hooking Related Declares
Private Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)
Private Declare Function CallWindowProc& Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long)
Private Declare Function SetWindowLong& Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)

Type Instances
    in_use As Boolean       'This instance is alive
    ClassAddr As Long       'Pointer to self
    hwnd As Long            'hWnd being hooked
    PrevWndProc As Long     'Stored for unhooking
End Type

Private m_Instances() As Instances

Public Function AddInstance(ByVal lngClassAddr As Long, ByVal lnghWnd As Long) As Integer
  Dim intI As Integer
  
  '*** Is there an unused instance that we can use?
  On Error GoTo AddInstance_FirstTime
  For intI = LBound(m_Instances) To UBound(m_Instances)
    If m_Instances(intI).in_use = False Then Exit For
  Next intI
  '*** Didn't find a free instance, set one up!
  If intI > UBound(m_Instances) Then ReDim Preserve m_Instances(intI)
AddInstance_FirstTime_Continue:
  '*** When arriving here intI points to the instance index
  m_Instances(intI).in_use = True
  m_Instances(intI).ClassAddr = lngClassAddr
  Call HookWindow(lnghWnd, intI)
  AddInstance = intI
  GoTo AddInstance_End
AddInstance_FirstTime:
  '*** First run, m_Instances not initialized!
  intI = 0
  ReDim m_Instances(intI)
  Resume AddInstance_FirstTime_Continue
AddInstance_End:
End Function

Public Function HandleMessage(ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim intI As Integer
  Dim cMyUC As ControlSizer
  Dim PrevWndProc As Long
  
  'Do this early as we may unhook
  PrevWndProc = IsHooked(hwnd)
  If MSG = WM_SIZE Or MSG = WM_CLOSE Then
    For intI = LBound(m_Instances) To UBound(m_Instances)
      If m_Instances(intI).hwnd = hwnd Then
        On Error Resume Next
        CopyMemory cMyUC, m_Instances(intI).ClassAddr, 4
        cMyUC.ParentResized MSG
        CopyMemory cMyUC, 0&, 4
      End If
    Next intI
  End If
  HandleMessage = CallWindowProc(PrevWndProc, hwnd, MSG, wParam, lParam)
End Function

'Hooks a window or acts as if it does if the window is
'already hooked by a previous instance of the control.
Public Sub HookWindow(ByVal hwnd As Long, ByVal instance_ndx As Integer)
  m_Instances(instance_ndx).PrevWndProc = IsHooked(hwnd)
  If m_Instances(instance_ndx).PrevWndProc = 0& Then
    m_Instances(instance_ndx).PrevWndProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf HandleMessage)
  End If
  m_Instances(instance_ndx).hwnd = hwnd
End Sub

' Unhooks only if no other instances need the hWnd
Public Sub UnHookWindow(ByVal instance_ndx As Integer)
  If TimesHooked(m_Instances(instance_ndx).hwnd) = 1 Then
    SetWindowLong m_Instances(instance_ndx).hwnd, GWL_WNDPROC, m_Instances(instance_ndx).PrevWndProc
  End If
  m_Instances(instance_ndx).hwnd = 0&
End Sub

'Determine if we have already hooked a window,
'and returns the PrevWndProc if true, 0& if false
Private Function IsHooked(ByVal hwnd As Long) As Long
  Dim intI As Integer
  
  IsHooked = 0&
  For intI = LBound(m_Instances) To UBound(m_Instances)
    If m_Instances(intI).hwnd = hwnd Then
      IsHooked = m_Instances(intI).PrevWndProc
      Exit For
    End If
  Next intI
End Function

'Returns a count of the number of times a given
'window has been hooked by instances of the control.
Private Function TimesHooked(ByVal hwnd As Long) As Long
  Dim intI As Integer
  Dim intCnt As Integer
    
  For intI = LBound(m_Instances) To UBound(m_Instances)
    If m_Instances(intI).hwnd = hwnd Then
      intCnt = intCnt + 1
    End If
  Next intI
  TimesHooked = intCnt
End Function

