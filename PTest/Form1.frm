VERSION 5.00
Object = "*\A..\PCaptionButton.vbp"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3120
   ClientLeft      =   5280
   ClientTop       =   5475
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   4665
   Begin PCaptionButton.CaptionButton CaptionButton2 
      Left            =   2640
      Top             =   120
      _ExtentX        =   423
      _ExtentY        =   370
      LeftOffset      =   91
   End
   Begin PCaptionButton.CaptionButton CaptionButton1 
      Left            =   2280
      Top             =   120
      _ExtentX        =   423
      _ExtentY        =   370
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Show Button"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Enabled"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CaptionButton1_Click()
    Log "Event: Click"
End Sub

Private Sub Log(sText As String)
    List1.AddItem sText
    List1.TopIndex = List1.NewIndex
End Sub

Private Sub CaptionButton1_MouseDown()
    Log "Event: MouseDown"
End Sub

Private Sub CaptionButton1_MouseMove()
    Log "Event: MouseMove"
End Sub

Private Sub CaptionButton1_MouseUp()
    Log "Event: MouseUp"
End Sub

Private Sub CaptionButton2_Click()
    Log "Button2 Clicked!"
End Sub

Private Sub Check1_Click()
    If Check1.Value = vbChecked Then
        CaptionButton1.Show
    Else
        CaptionButton1.Hide
    End If
End Sub

Private Sub Check2_Click()
    If Check2.Value = vbChecked Then
        CaptionButton1.Enabled = True
    Else
        CaptionButton1.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    Check1.Value = Abs(CInt(CaptionButton1.Visible))
    Check2.Value = Abs(CInt(CaptionButton1.Enabled))
End Sub
