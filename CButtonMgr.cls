VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CButtonMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' For subclass debugging
'Private m_wndprcNext As Long
#If DEBUGWINDOWPROC Then
Private m_SCHook As WindowProcHook
#End If

' This class manages the CaptionButton instances. All windows have their
' own manager, all buttons within one window are managed by a single manager
' It hold a reference to each of them, when all references are destroyed
' the subclass will be removed as well, when a new button wants to become
' active, the subclass will be installed again.

Private m_colButtonsPtr As Collection
Private m_oldWndProcAddress As Long
Private m_hWnd As Long
Private m_oSelf As CButtonMgr

Friend Property Get oldWndProcAddress() As Long
    oldWndProcAddress = m_oldWndProcAddress
End Property

Friend Sub RedrawAll()
    MDebug.Log Me, "Redrawing all buttons"

    ' Loop through all buttons and tell them to redraw
    Dim pCB As Long, oCB As CaptionButton, n As Long
    For n = 1 To m_colButtonsPtr.Count
        pCB = CLng(m_colButtonsPtr.Item(n))
        Set oCB = CaptionButtonFromPtr(pCB)
        oCB.drawTitleButton
    Next n
End Sub

Friend Sub RedrawWindow()
    MDebug.Log Me, "Redrawing window frame and all buttons"
    Call SendMessage(m_hWnd, WM_NCPAINT, 1&, 0&)
End Sub

Friend Sub DoMouseMove(p As POINTAPI)
    ' Loop through all buttons and tell them the mouse is moving and where it is
    Dim pCB As Long, oCB As CaptionButton, n As Long
    For n = 1 To m_colButtonsPtr.Count
        pCB = CLng(m_colButtonsPtr.Item(n))
        Set oCB = CaptionButtonFromPtr(pCB)
        oCB.DoMouseMove p
    Next n
End Sub

Friend Sub DoMouseUp(p As POINTAPI)
    ' Loop through all buttons and tell them the left mousebutton is released
    Dim pCB As Long, oCB As CaptionButton, n As Long
    For n = 1 To m_colButtonsPtr.Count
        pCB = CLng(m_colButtonsPtr.Item(n))
        Set oCB = CaptionButtonFromPtr(pCB)
        oCB.DoMouseUp p
    Next n
End Sub

Friend Function ButtonFromPoint(p As POINTAPI) As CaptionButton
    ' Go through all buttons and see if any are under the point specified
    Dim pCB As Long, oCB As CaptionButton, n As Long
    Set ButtonFromPoint = Nothing
    For n = 1 To m_colButtonsPtr.Count
        pCB = CLng(m_colButtonsPtr.Item(n))
        Set oCB = CaptionButtonFromPtr(pCB)
        If oCB.isOverButton(p) Then
            ' Return a reference to the button.
            Set ButtonFromPoint = oCB
            Exit For
        End If
    Next n
End Function

Friend Sub AddButton(cb As CaptionButton)
    If m_hWnd = 0 Then _
        m_hWnd = cb.hWnd
        
    If m_hWnd = cb.hWnd Then
        ' Add the button to the collection
        m_colButtonsPtr.Add ObjPtr(cb)
    
        ' Subclass the window if required
        If m_oldWndProcAddress = 0 Then
            SubClass
        Else
            ' Redraw the window frame
            RedrawAll
        End If
        MDebug.Log Me, "Button added, number of buttons: " & m_colButtonsPtr.Count
    Else
        MDebug.Log Me, "Some major error in CbuttonMgr.AddButton"
        MDebug.Log Me, "m_hWnd = " & CStr(m_hWnd)
        MDebug.Log Me, "cb.hWnd= " & CStr(cb.hWnd)
    End If
End Sub

Friend Sub RemoveButton(cb As CaptionButton)
    ' Remove the button from the collection
    Dim n As Long
    For n = 1 To m_colButtonsPtr.Count
        If CLng(m_colButtonsPtr.Item(n)) = ObjPtr(cb) Then
            m_colButtonsPtr.Remove n
            Exit For
        End If
    Next n
    
    MDebug.Log Me, "Button removed, number of buttons: " & m_colButtonsPtr.Count

    ' If there are no more buttons, unsubclass the window and destroy the manager
    If m_colButtonsPtr.Count = 0 Then
        UnSubClass
        Set m_oSelf = Nothing
    Else
        ' Force the window to redraw
        RedrawWindow
    End If
End Sub

Private Sub SubClass()
    ' If window isn't subclassed yet, subclass it
    If m_oldWndProcAddress = 0 Then
        ' Save a pointer to the manager object in an extra property in the window
        Call SetProp(m_hWnd, "gdCBObjPtr", ObjPtr(Me))
        ' Store the old wndProcAddress
        m_oldWndProcAddress = GetWindowLong(m_hWnd, GWL_WNDPROC)
        ' Install the new handler
        
#If DEBUGWINDOWPROC Then
        On Error Resume Next
        Set m_SCHook = CreateWindowProcHook
        If Err Then
            MsgBox Err.Description
            Err.Clear
            UnSubClass
            Exit Sub
        End If
        On Error GoTo 0
        With m_SCHook
            .SetMainProc AddressOf MCaptionButton.wndProc
            Call SetWindowLong(m_hWnd, GWL_WNDPROC, .ProcAddress)
            .SetDebugProc m_oldWndProcAddress
        End With
#Else
        Call SetWindowLong(m_hWnd, GWL_WNDPROC, AddressOf MCaptionButton.wndProc)
#End If

        MDebug.Log Me, "Subclass installed"
        
        RedrawAll
    End If
End Sub

Private Sub UnSubClass()
    ' If window is subclassed, remove the subclass
    If m_oldWndProcAddress <> 0 Then
        ' Install the old handler
        SetWindowLong m_hWnd, GWL_WNDPROC, m_oldWndProcAddress
        MDebug.Log Me, "Subclass removed"
        ' Force a repaint of the windowframe, erasing all buttons
        RedrawWindow
        ' Erase reference to old wndProc (not needed anymore)
        m_oldWndProcAddress = 0
        ' Remove the extra property we added to the frame
        Call RemoveProp(m_hWnd, "gdCBObjPtr")
    End If
End Sub

Private Sub Class_Initialize()
    MDebug.Log Me, "CButtonMgr_Initialize()"
    Set m_colButtonsPtr = New Collection
    Set m_oSelf = Me
End Sub

Private Sub Class_Terminate()
    MDebug.Log Me, "CButtonMgr_Terminate()"
    UnSubClass
End Sub

Private Function CaptionButtonFromPtr(cbPtr As Long) As CaptionButton
    Dim oCB As CaptionButton
    CopyMemory oCB, cbPtr, 4&
    Set CaptionButtonFromPtr = oCB
    CopyMemory oCB, 0&, 4&
End Function
