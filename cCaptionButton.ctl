VERSION 5.00
Begin VB.UserControl CaptionButton 
   CanGetFocus     =   0   'False
   ClientHeight    =   210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   240
   Enabled         =   0   'False
   HasDC           =   0   'False
   HitBehavior     =   0  'None
   InvisibleAtRuntime=   -1  'True
   PaletteMode     =   4  'None
   Picture         =   "cCaptionButton.ctx":0000
   PropertyPages   =   "cCaptionButton.ctx":02E2
   ScaleHeight     =   210
   ScaleWidth      =   240
   ToolboxBitmap   =   "cCaptionButton.ctx":02F2
End
Attribute VB_Name = "CaptionButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

Private Const MAX_WIDTH = 240
Private Const MAX_HEIGHT = 210

Public Enum cbStateConstants
    cbRaised
    cbSunken
End Enum

' This var is for when the button is reading it's properties, one redraw is enough then
Private m_NoRedraw As Boolean

Private m_bClickInProgress As Boolean
Private m_bVisible As Boolean
Private m_bEnabled As Boolean
Private m_targethWnd As Long
Private m_RightToLeft As Boolean
Private WndRECT As RECT
Private ButtonRECT As RECT
Private cbButtonState As cbStateConstants
Private ButtonX As Long
Private ButtonY As Long
Private FrameY As Long

Private m_Picture As StdPicture
Private m_PictureXOffset As Long
Private m_PictureYOffset As Long
Private m_PicturehDC As Long

Private m_TopOffset As Long
Private m_LeftOffset As Long
Private m_ButtonWidth As Long
Private m_ButtonHeight As Long

Event Click()
Attribute Click.VB_Description = "Fires when the left mouse is clicked on the button"
Event MouseDown()
Event MouseMove()
Event MouseUp()

' Button size properties
Public Property Get TopOffset() As Long
Attribute TopOffset.VB_Description = "The offset of the top of the button from the top of the window"
Attribute TopOffset.VB_ProcData.VB_Invoke_Property = "Apperance"
    MDebug.Log Me, "Get TopOffset"
    TopOffset = m_TopOffset
End Property
Public Property Let TopOffset(o As Long)
    MDebug.Log Me, "Let TopOffset"
    m_TopOffset = o
    RedrawAll
    
    PropertyChanged "TopOffset"
End Property
Public Property Get LeftOffset() As Long
Attribute LeftOffset.VB_Description = "Offset of the left side of  the button from the right edge of the window"
Attribute LeftOffset.VB_ProcData.VB_Invoke_Property = "Apperance"
    MDebug.Log Me, "Get LeftOffset"
    LeftOffset = m_LeftOffset
End Property
Public Property Let LeftOffset(o As Long)
    MDebug.Log Me, "Let LeftOffset"
    m_LeftOffset = o
    RedrawAll
    
    PropertyChanged "LeftOffset"
End Property
Public Property Get ButtonWidth() As Long
Attribute ButtonWidth.VB_Description = "Sets the width of the CaptionButton in pixels"
Attribute ButtonWidth.VB_ProcData.VB_Invoke_Property = "Apperance"
    MDebug.Log Me, "Get ButtonWidth"
    ButtonWidth = m_ButtonWidth
End Property
Public Property Let ButtonWidth(o As Long)
    MDebug.Log Me, "Let ButtonWidth"
    m_ButtonWidth = o
    RedrawAll
    
    PropertyChanged "ButtonWidth"
End Property
Public Property Get ButtonHeight() As Long
Attribute ButtonHeight.VB_Description = "Sets the height of the CaptionButton in pixels"
Attribute ButtonHeight.VB_ProcData.VB_Invoke_Property = "Apperance"
    MDebug.Log Me, "Get ButtonHeigt"
    ButtonHeight = m_ButtonHeight
End Property
Public Property Let ButtonHeight(o As Long)
    MDebug.Log Me, "Let ButtonHeight"
    m_ButtonHeight = o
    RedrawAll
    
    PropertyChanged "ButtonHeight"
End Property
' End button size properties

' Picture properties
Public Property Get Picture() As StdPicture
Attribute Picture.VB_Description = "Sets the bitmap for the button."
Attribute Picture.VB_ProcData.VB_Invoke_Property = ";Appearance"
    MDebug.Log Me, "Get Picture"
    Set Picture = m_Picture
End Property
Public Property Set Picture(p As StdPicture)
    MDebug.Log Me, "Set Picture"
    Set m_Picture = p
    
    If m_PicturehDC <> 0 Then
        Call ReleaseDC(hWnd, m_PicturehDC)
        m_PicturehDC = 0
    End If
    
    If Not m_Picture Is Nothing Then
        Dim lDC As Long
        ' Get the Forms Device Context (DC)
        lDC = GetDC(hWnd)
        ' Create compatible DC for the image
        m_PicturehDC = CreateCompatibleDC(lDC)
        ' Load image in DC
        Call SelectObject(m_PicturehDC, m_Picture.Handle)
        ' Release the forms DC
        Call ReleaseDC(hWnd, lDC)
    End If

    Redraw
    
    PropertyChanged "Picture"
End Property
Public Property Get PictureXOffset() As Long
Attribute PictureXOffset.VB_ProcData.VB_Invoke_Property = "Apperance"
    MDebug.Log Me, "Get PictureXOffset"
    PictureXOffset = m_PictureXOffset
End Property
Public Property Let PictureXOffset(o As Long)
    MDebug.Log Me, "Set PictureXOffset"
    m_PictureXOffset = o
    Redraw
    
    PropertyChanged "PictureXOffset"
End Property
Public Property Get PictureYOffset() As Long
Attribute PictureYOffset.VB_ProcData.VB_Invoke_Property = "Apperance"
    MDebug.Log Me, "Get PictureYOffset"
    PictureYOffset = m_PictureYOffset
End Property
Public Property Let PictureYOffset(o As Long)
    MDebug.Log Me, "Set PictureYOffset"
    m_PictureYOffset = o
    Redraw
    
    PropertyChanged "PictureYOffset"
End Property
' End picture properties

' Misc properties
Public Property Get Visible() As Boolean
Attribute Visible.VB_Description = "Indicates wether the button is visible"
Attribute Visible.VB_ProcData.VB_Invoke_Property = "Apperance"
    MDebug.Log Me, "Get Visible"
    Visible = m_bVisible
End Property
Public Property Let Visible(v As Boolean)
    If m_bVisible = v Then Exit Property
    MDebug.Log Me, "Set Visible"
    m_bVisible = v
    
    MDebug.Log Me, "Getting ButtonManager for hWnd: " & hWnd
    ' Get a reference to the ButtonManager for our hWnd
    Dim oMgr As CButtonMgr
    Set oMgr = CButtonMgrFromhWnd(hWnd)
    If oMgr Is Nothing Then
        MDebug.Log Me, "No ButtonMgr found"
    Else
        MDebug.Log Me, "Got it!"
    End If
        
    If v = True Then
        If oMgr Is Nothing Then
            MDebug.Log Me, "Creating new ButtonManager"
            Set oMgr = New CButtonMgr
        End If
        oMgr.AddButton Me
    Else
        If Not oMgr Is Nothing Then
            oMgr.RemoveButton Me
        End If
    End If
    
    PropertyChanged "Visible"
End Property
Public Property Get State() As cbStateConstants
    MDebug.Log Me, "Get State"
    State = cbButtonState
End Property
Public Property Let State(s As cbStateConstants)
    If s = cbButtonState Then Exit Property
    MDebug.Log Me, "Set State"

    cbButtonState = s
    Redraw
    
    PropertyChanged "State"
End Property
Public Property Get Enabled() As Boolean
    Enabled = m_bEnabled
End Property
Public Property Let Enabled(e As Boolean)
    If e = m_bEnabled Then Exit Property
    MDebug.Log Me, "Set Enabled"
    m_bEnabled = e
    
    PropertyChanged "Enabled"
End Property
' End misc properties

' Internal properties, not public.
Friend Property Get hWnd() As Long
    If m_targethWnd = 0 Then
        m_targethWnd = Parent.hWnd
    End If
    hWnd = m_targethWnd
End Property
Private Property Get NoRedraw() As Boolean
    NoRedraw = m_NoRedraw
End Property
Private Property Let NoRedraw(b As Boolean)
    m_NoRedraw = b
    If b = True Then
        MDebug.Log Me, "Redrawing disabled"
    Else
        MDebug.Log Me, "Redrawing enabled"
    End If
End Property
' End internal properties

Public Sub Show()
    Visible = True
End Sub

Public Sub Hide()
    Visible = False
End Sub

' Init the control
Private Sub UserControl_Initialize()
    MDebug.Log Me, "Control_Initialize()"
    ButtonX = GetSystemMetrics(SM_CXSIZE)
    ButtonY = GetSystemMetrics(SM_CYSIZE)
    FrameY = GetSystemMetrics(SM_CYFRAME)
End Sub

Private Sub UserControl_InitProperties()
    NoRedraw = True

    MDebug.Log Me, "InitProperties"
    TopOffset = FrameY + 1
    LeftOffset = ((4 * ButtonX) + 2)
    ButtonWidth = ButtonX - 3
    ButtonHeight = ButtonY - 4
    PictureXOffset = 1
    PictureYOffset = 1
    Set Picture = Nothing
    State = cbRaised
    Enabled = True
    
    NoRedraw = False
    Visible = True
End Sub

Private Sub UserControl_Paint()
    If Not Picture Is Nothing Then _
        PaintPicture Picture, PictureXOffset * Screen.TwipsPerPixelX, PictureYOffset * Screen.TwipsPerPixelY
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    NoRedraw = True
    
    MDebug.Log Me, "ReadProperties"
    With PropBag
        TopOffset = .ReadProperty("TopOffset", FrameY + 1)
        LeftOffset = .ReadProperty("LeftOffset", ((4 * ButtonX) + 2))
        ButtonWidth = .ReadProperty("ButtonWidth", ButtonX - 3)
        ButtonHeight = .ReadProperty("ButtonHeight", ButtonY - 4)
        Set Picture = .ReadProperty("Picture", Nothing)
        PictureXOffset = .ReadProperty("PictureXOffset", 1)
        PictureYOffset = .ReadProperty("PictureYOffset", 1)
        State = .ReadProperty("State", cbRaised)
        Enabled = .ReadProperty("Enabled", True)
m_RightToLeft = Ambient.RightToLeft
NoRedraw = False
        Visible = .ReadProperty("Visible", True)
    End With
End Sub

' Make sure the control is never resized
Private Sub UserControl_Resize()
    MDebug.Log Me, "Control_Resize()"
    Height = MAX_HEIGHT
    Width = MAX_WIDTH
End Sub

Private Sub UserControl_Terminate()
    If Visible Then Hide
    If m_PicturehDC <> 0 Then
        Call ReleaseDC(hWnd, m_PicturehDC)
        m_PicturehDC = 0
    End If
    MDebug.Log Me, "Control_Terminate()"
End Sub

' Redraws the button, but only when the button should be visible
Public Sub Redraw()
Attribute Redraw.VB_Description = "Forces a redraw of the button"
    If NoRedraw Then Exit Sub
    
    If Visible Then
        drawTitleButton
    End If
End Sub

' Redraws all buttons
Public Sub RedrawAll()
    If NoRedraw Then Exit Sub
    MDebug.Log Me, "RedrawAll"

    ' Get the button manager
    Dim oMgr As CButtonMgr
    Set oMgr = CButtonMgrFromhWnd(hWnd)
    If Not oMgr Is Nothing Then
        ' Tell it to redraw the window and all buttons in it
        oMgr.RedrawWindow
    End If
End Sub

' Redraws the button
Friend Sub drawTitleButton()
    If NoRedraw Then Exit Sub
    MDebug.Log Me, "drawTitleButton"
    
    Dim hdc As Long
    Dim edge As Long

    Call GetWindowRect(hWnd, WndRECT)

    With ButtonRECT
        .Top = TopOffset
        If m_RightToLeft Then
            .Left = LeftOffset
        Else
            .Left = (WndRECT.Right - WndRECT.Left + 1) - LeftOffset
        End If
        .Right = .Left + ButtonWidth
        .Bottom = .Top + ButtonHeight
    End With
    
    If State = cbRaised Then
        edge = EDGE_RAISED
    Else
        edge = EDGE_SUNKEN
    End If
    
    hdc = GetWindowDC(hWnd)
    If m_PicturehDC <> 0 Then
        Call DrawEdge(hdc, ButtonRECT, edge, BF_SOFT Or BF_RECT)
        Call BitBlt(hdc, ButtonRECT.Left + PictureXOffset, ButtonRECT.Top + PictureYOffset, ButtonWidth, ButtonHeight, _
                    m_PicturehDC, 0, 0, _
                    SRCCOPY)
    Else
        Call DrawEdge(hdc, ButtonRECT, edge, BF_SOFT Or BF_RECT Or BF_MIDDLE)
    End If
    Call ReleaseDC(hWnd, hdc)
End Sub

Friend Function isOverButton(pt As POINTAPI) As Boolean
    Call GetWindowRect(hWnd, WndRECT)
    isOverButton = PtInRECT(MoveInToRECT(ButtonRECT, WndRECT), pt)
End Function

Private Function MoveInToRECT(Small As RECT, Large As RECT) As RECT
    Dim res As RECT
    res.Top = Small.Top + Large.Top
    res.Left = Small.Left + Large.Left
    res.Right = Small.Right + Large.Left
    res.Bottom = Small.Bottom + Large.Top
    MoveInToRECT = res
End Function

Friend Sub DoNCMouseDown(CursorPos As POINTAPI)
    If Not Enabled Then Exit Sub
    MDebug.Log Me, "DoNCMouseDown"
    
    ' The mouse is over the button and the left mouse button is pressed
    ' set the button state to sunken (down)
    State = cbSunken
    m_bClickInProgress = True
    RaiseEvent MouseDown
End Sub

Friend Sub DoMouseUp(CursorPos As POINTAPI)
    If Not Enabled Then Exit Sub

    MDebug.Log Me, "DoMouseUp"

    ' The mouse is over the button and the left mousebutton is released
    ' If the button was in the 'down' state, bring it up and raise a click event
    ' otherwise, just raise mouseup
    If isOverButton(CursorPos) Then
        If m_bClickInProgress Then
            State = cbRaised
            RaiseEvent MouseUp
            RaiseEvent Click
        Else
            RaiseEvent MouseUp
        End If
    End If
    m_bClickInProgress = False
End Sub

Friend Sub DoMouseMove(CursorPos As POINTAPI)
    If Not Enabled Then Exit Sub
    MDebug.Log Me, "DoMouseMove"
    
    If isOverButton(CursorPos) Then
        If m_bClickInProgress Then State = cbSunken
        RaiseEvent MouseMove
    Else
        If m_bClickInProgress Then
            State = cbRaised
        End If
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    MDebug.Log Me, "WriteProperties"
    With PropBag
        .WriteProperty "TopOffset", TopOffset, FrameY + 1
        .WriteProperty "LeftOffset", LeftOffset, ((4 * ButtonX) + 2)
        .WriteProperty "ButtonWidth", ButtonWidth, ButtonX - 3
        .WriteProperty "ButtonHeight", ButtonHeight, ButtonY - 4
        .WriteProperty "Picture", Picture, Nothing
        .WriteProperty "PictureXOffset", PictureXOffset, 1
        .WriteProperty "PictureYOffset", PictureYOffset, 1
        .WriteProperty "Visible", Visible, True
        .WriteProperty "State", State, cbRaised
        .WriteProperty "Enabled", Enabled, True
    End With
End Sub
