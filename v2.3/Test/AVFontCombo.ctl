VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl AVFontCombo 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6885
   ScaleHeight     =   3960
   ScaleWidth      =   6885
   ToolboxBitmap   =   "AVFontCombo.ctx":0000
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   3360
      Sorted          =   -1  'True
      TabIndex        =   6
      Text            =   "Combo2"
      Top             =   360
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3360
      Sorted          =   -1  'True
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   0
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.PictureBox picDrop 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      Height          =   330
      Left            =   2985
      ScaleHeight     =   330
      ScaleWidth      =   240
      TabIndex        =   4
      Top             =   30
      Width           =   240
      Begin MSComCtl2.UpDown updDrop 
         Height          =   735
         Left            =   0
         TabIndex        =   9
         Top             =   -420
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   1296
         _Version        =   393216
         Enabled         =   -1  'True
      End
   End
   Begin VB.PictureBox picOuter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      CausesValidation=   0   'False
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   0
      ScaleHeight     =   2385
      ScaleWidth      =   3225
      TabIndex        =   0
      Top             =   390
      Width           =   3255
      Begin VB.VScrollBar VScroll1 
         Height          =   2385
         LargeChange     =   500
         Left            =   2970
         SmallChange     =   50
         TabIndex        =   1
         Top             =   0
         Width           =   240
      End
      Begin VB.PictureBox picInner 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         Height          =   2535
         Left            =   0
         ScaleHeight     =   2535
         ScaleWidth      =   3015
         TabIndex        =   2
         Top             =   0
         Width           =   3015
         Begin VB.Label lblFont 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Label1"
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   480
         End
      End
   End
   Begin VB.TextBox txtMain 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Text            =   "Times New Roman"
      Top             =   0
      Width           =   3255
   End
   Begin VB.Label lblFontSize 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   255
      Left            =   3360
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   540
   End
End
Attribute VB_Name = "AVFontCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim booDrop As Boolean
Dim booClick As Boolean
Dim intOld As Integer
Dim strText As String
Dim intFontSize As Integer
Dim lngBackColor As Long
Dim lngForeColor As Long
Const ORIGINAL_TEXT = "Times New Roman"
Const ORIGINAL_SIZE = 8
Const ORIGINAL_BACKCOLOR = &H80000005
Const ORIGINAL_FORECOLOR = &H80000005

Public Event Click()
Public Event Change()
Public Event Scroll()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type POINTAPI
    x As Long
    Y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function Putfocus Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long

Private Sub Combo1_LostFocus()

    subLostIt
    
End Sub

Private Sub Combo2_LostFocus()

    subLostIt
    
End Sub

Private Sub lblfont_Click(Index As Integer)

    booClick = True
    txtMain.Text = lblFont(Index).Caption
    txtMain.FontName = lblFont(Index).Caption
    txtMain.FontBold = lblFont(Index).FontBold
    txtMain.FontItalic = lblFont(Index).FontItalic
    txtMain.FontSize = lblFont(Index).FontSize
    booDrop = True
    Height = Height - 2415
    booDrop = False
    RaiseEvent Click
    
End Sub

Private Sub lblfont_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    
    lblFont(Index).BackColor = vbHighlight
    lblFont(Index).ForeColor = vbHighlightText
    If intOld <> Index Then
        lblFont(intOld).BackColor = lngBackColor
        lblFont(intOld).ForeColor = lngForeColor
        intOld = Index
    End If
    VScroll1.SetFocus
End Sub

Private Sub picDrop_LostFocus()

    subLostIt
    
End Sub

Private Sub picInner_LostFocus()

    subLostIt
    
End Sub

Private Sub picOuter_LostFocus()

    subLostIt
    
End Sub

Private Sub txtMain_Change()

    strText = txtMain.Text
    PropertyChanged "Text"
    If booClick = False Then
        RaiseEvent Change
    End If
    booClick = False
    
End Sub

Private Sub txtMain_KeyDown(KeyCode As Integer, Shift As Integer)

    RaiseEvent KeyDown(KeyCode, Shift)
    
End Sub

Private Sub txtMain_KeyPress(KeyAscii As Integer)

    RaiseEvent KeyPress(KeyAscii)
    If KeyAscii = 13 Then
        subLostIt
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtMain_KeyUp(KeyCode As Integer, Shift As Integer)

    RaiseEvent KeyUp(KeyCode, Shift)
    
End Sub

Private Sub txtMain_LostFocus()

    subLostIt
    
End Sub

Private Sub upddrop_DownClick()
    
    booDrop = True
    If Height = txtMain.Height Then
        Height = Height + 2415
        VScroll1.SetFocus
    Else
        Height = txtMain.Height
        txtMain.SetFocus
    End If
    booDrop = False
    
End Sub

Private Sub UserControl_Initialize()

Dim x As Integer

    For x = 0 To Screen.FontCount - 1
        If Left(Screen.Fonts(x), 1) <> "@" Then
            Combo1.AddItem Screen.Fonts(x)
        Else
            Combo2.AddItem Screen.Fonts(x)
        End If
    Next x
    
End Sub

Private Sub UserControl_InitProperties()
    
    intFontSize = ORIGINAL_SIZE
    strText = ORIGINAL_TEXT
    lngBackColor = ORIGINAL_BACKCOLOR
    lngForeColor = ORIGINAL_BACKCOLOR
    
End Sub

Private Sub UserControl_LostFocus()

    subLostIt
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

Dim x As Integer

    On Error Resume Next
    strText = PropBag.ReadProperty("Text", ORIGINAL_TEXT)
    txtMain = strText
    intFontSize = PropBag.ReadProperty("FontSize", ORIGINAL_SIZE)
    txtMain.FontSize = intFontSize
    lblFontSize.FontSize = intFontSize
    subFillCombo
    lngBackColor = PropBag.ReadProperty("BackColor", ORIGINAL_BACKCOLOR)
    lngForeColor = PropBag.ReadProperty("ForeColor", ORIGINAL_FORECOLOR)
    txtMain.BackColor = lngBackColor
    txtMain.ForeColor = lngForeColor
    For x = 0 To lblFont.Count - 1
        lblFont(x).BackColor = lngBackColor
        lblFont(x).ForeColor = lngForeColor
    Next x

End Sub

Private Sub UserControl_Resize()

    If booDrop = False Then
        Height = lblFontSize.Height + 60
        txtMain.Height = Height
        txtMain.Width = Width
        picDrop.Height = Height - 55
        picDrop.Left = Width - 270
        updDrop.Top = -(Height - 55)
        updDrop.Height = Height * 2 - 120
        picOuter.Top = Height
        picOuter.Width = Width
        picInner.Width = Width
        VScroll1.Left = picDrop.Left
    End If
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    On Error Resume Next
    Call PropBag.WriteProperty("Text", strText, ORIGINAL_TEXT)
    Call PropBag.WriteProperty("FontSize", intFontSize, ORIGINAL_SIZE)
    Call PropBag.WriteProperty("BackColor", lngBackColor, ORIGINAL_BACKCOLOR)
    Call PropBag.WriteProperty("ForeColor", lngForeColor, ORIGINAL_FORECOLOR)
    
End Sub

Public Property Let Text(ByVal New_Text As String)

Dim x As Integer
Dim booGood As Boolean

    For x = 0 To Screen.FontCount - 1
        If UCase(New_Text) = UCase(lblFont(x).Caption) Then
            strText = lblFont(x).Caption
            booGood = True
            Exit For
        End If
    Next x
    If booGood Then
        txtMain.FontName = strText
        txtMain = strText
        PropertyChanged Text
    Else
        MsgBox "Invalid Font Name"
    End If
    
End Property

Public Property Get Text() As String

    Text = strText

End Property

Public Property Let FontSize(ByVal New_FontSize As Integer)

Dim x As Integer

    If intFontSize <> New_FontSize Then
        intFontSize = New_FontSize
        PropertyChanged FontSize
        lblFontSize.FontSize = intFontSize
        txtMain.FontSize = intFontSize
        picInner.Height = 0
        For x = 0 To lblFont.Count - 1
            lblFont(x).FontSize = intFontSize
            If x <> 0 Then lblFont(x).Top = lblFont(x - 1).Top + lblFont(x - 1).Height
            picInner.Height = picInner.Height + lblFont(x).Height
        Next x
        Height = lblFontSize.Height
        txtMain.Height = Height
    End If
    
End Property

Public Property Get FontSize() As Integer

    FontSize = intFontSize

End Property

Public Property Let BackColor(ByVal New_BackColor As Long)

    lngBackColor = New_BackColor
    txtMain.BackColor = New_BackColor
    
End Property

Public Property Get BackColor() As Long

    BackColor = lngBackColor

End Property

Public Property Let ForeColor(ByVal New_ForeColor As Long)

    lngForeColor = New_ForeColor
    txtMain.ForeColor = New_ForeColor
    
End Property

Public Property Get ForeColor() As Long

    ForeColor = lngForeColor

End Property

Private Sub VScroll1_Change()

    picInner.Top = VScroll1.Value / VScroll1.Max * (picInner.Height - picOuter.Height) * -1
    
End Sub

Private Sub VScroll1_LostFocus()

    subLostIt
    
End Sub

Public Sub subLostIt()

Dim x As Integer

    Dim Rec As RECT, Point As POINTAPI
    GetWindowRect UserControl.hwnd, Rec
    GetCursorPos Point
    If Point.x < Rec.Left Or Point.x > Rec.Right Or Point.Y < Rec.Top Or Point.Y > Rec.Bottom Then
        booDrop = True
        If Height > 2415 Then Height = Height - 2415
        booDrop = False
        For x = 0 To Screen.FontCount - 1
            If UCase(Trim(txtMain.Text)) = UCase(lblFont(x).Caption) Then
                txtMain.Text = lblFont(x).Caption
                txtMain.FontName = txtMain.Text
                txtMain.FontItalic = lblFont(x).FontItalic
                txtMain.FontBold = lblFont(x).FontBold
                Exit Sub
            End If
        Next x
        txtMain.Text = "Times New Roman"
        txtMain.FontName = "Times New Roman"
    End If
    
End Sub

Public Sub subFillCombo()
    
Dim x As Integer

    For x = 1 To lblFont.Count - 1
        Unload lblFont(x)
    Next x
    picInner.Height = 0
    For x = 0 To Screen.FontCount - 1
        If x <> 0 Then
            Load lblFont(x)
            lblFont(x).Top = lblFont(x - 1).Top + lblFont(x - 1).Height
            lblFont(x).Visible = True
        End If
        If x < Combo1.ListCount Then
            lblFont(x).FontName = Combo1.List(x)
        Else
            lblFont(x).FontName = Combo2.List(x - Combo1.ListCount)
        End If
        lblFont(x).Caption = lblFont(x).FontName
        lblFont(x).FontSize = intFontSize
        lblFont(x).FontBold = False
        lblFont(x).FontItalic = False
        lblFont(x).Width = picInner.Width + 3000
        picInner.Height = picInner.Height + lblFont(x).Height
    Next x
    
End Sub

Private Sub VScroll1_Scroll()

    RaiseEvent Scroll
    
End Sub
