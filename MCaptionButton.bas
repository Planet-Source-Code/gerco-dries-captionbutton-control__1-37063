Attribute VB_Name = "MCaptionButton"
Option Explicit

' Declares
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Public Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function GetLastError Lib "kernel32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

' Types
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Type POINTAPI
        x As Long
        y As Long
End Type

' Constanten
Public Const SM_CYBORDER = 6
Public Const SM_CXSIZE = 30
Public Const SM_CYSIZE = 31
Public Const SM_CYFRAME = 33
Public Const BDR_RAISEDINNER = &H4
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Public Const BF_LEFT = &H1
Public Const BF_TOP = &H2
Public Const BF_RIGHT = &H4
Public Const BF_BOTTOM = &H8
Public Const BF_MIDDLE = &H800
Public Const BF_SOFT = &H1000
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const WM_SIZE = &H5
Public Const WM_SETTEXT = &HC
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_NCHITTEST = &H84
Public Const WM_NCPAINT = &H85
Public Const WM_NCACTIVATE = &H86
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_NCLBUTTONDBLCLK = &HA3
Public Const WM_NCLBUTTONUP = &HA2
Public Const WM_NCMBUTTONDBLCLK = &HA9
Public Const WM_NCMBUTTONDOWN = &HA7
Public Const WM_NCMBUTTONUP = &HA8
Public Const WM_NCMOUSEMOVE = &HA0
Public Const WM_NCRBUTTONDBLCLK = &HA6
Public Const WM_NCRBUTTONDOWN = &HA4
Public Const WM_NCRBUTTONUP = &HA5
Public Const GWL_WNDPROC = (-4)
Public Const HTCAPTION = 2
Public Const HTCAPTIONBUTTON = 19
Public Const SRCCOPY = &HCC0020

Public Enum cbEventConstants
    cbeMouseMove
    cbeMouseDown
    cbeMouseUp
    cbeMouseMoveOutside
End Enum

Public Function wndProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim oMgr As CButtonMgr
    Dim oCB As CaptionButton
    Dim bBypassDefWndProc As Boolean
    
    Set oMgr = CButtonMgrFromhWnd(hWnd)
    
With oMgr
    ' Messages that need to be processed before the standard window gets them
    Select Case Msg
        Case WM_NCMBUTTONDOWN, WM_NCRBUTTONDOWN
            If Not .ButtonFromPoint(MakePoint(lParam)) Is Nothing Then
                bBypassDefWndProc = True
                wndProc = 0
            End If
    End Select

    ' Throw messages to the default window proc
    If Not bBypassDefWndProc Then _
        wndProc = CallWindowProc(.oldWndProcAddress, hWnd, Msg, wParam, lParam)

    ' Messages to be processed after the standard windowproc is done
    Select Case Msg
        Case WM_NCHITTEST
            If wndProc = HTCAPTION Then
                If Not .ButtonFromPoint(MakePoint(lParam)) Is Nothing Then
                    ' Tell windows that the mouse is over one of our buttons
                    wndProc = HTCAPTIONBUTTON
                End If
            End If

        Case WM_NCLBUTTONDOWN
            Set oCB = .ButtonFromPoint(MakePoint(lParam))
            If Not oCB Is Nothing Then
                ' Tell the button the the mouse button is being pressed over it.
                oCB.DoNCMouseDown MakePoint(lParam)
                wndProc = 0
            End If
        
        Case WM_NCLBUTTONUP, WM_LBUTTONUP
            ' Tell all buttons the left mousebutton is being released
            .DoMouseUp MakePoint(lParam)
            wndProc = 0
        
        Case WM_NCMOUSEMOVE, WM_MOUSEMOVE
            ' Tell all buttons that the mouse is moving and where it is
            .DoMouseMove MakePoint(lParam)
        
        Case WM_NCACTIVATE, WM_NCPAINT, WM_SIZE, WM_SETTEXT
            ' Tell all buttons to redraw themselves
            .RedrawAll
    End Select
End With
End Function

Public Function CButtonMgrFromhWnd(hWnd As Long) As CButtonMgr
    Dim oObj As CButtonMgr, pObj As Long
        
    pObj = GetProp(hWnd, "gdCBObjPtr")
    CopyMemory oObj, pObj, 4&
    Set CButtonMgrFromhWnd = oObj
    CopyMemory oObj, 0&, 4&
End Function

Public Function MakePoint(lParam As Long) As POINTAPI
    Dim res As POINTAPI
    res.x = WordLo(lParam)
    res.y = WordHi(lParam)
    MakePoint = res
End Function

Public Function PtInRECT(RECT As RECT, pt As POINTAPI) As Boolean
    PtInRECT = pt.x > RECT.Left And pt.x < RECT.Right And pt.y > RECT.Top And pt.y < RECT.Bottom
End Function

Public Function WordHi(LongIn As Long) As Integer
Call CopyMemory(WordHi, ByVal (VarPtr(LongIn) + 2), 2)
End Function

Public Function WordLo(LongIn As Long) As Integer
Call CopyMemory(WordLo, ByVal VarPtr(LongIn), 2)
End Function

