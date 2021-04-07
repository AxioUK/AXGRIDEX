VERSION 5.00
Begin VB.UserControl AxGridGroup 
   BackColor       =   &H00808080&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox pic2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   3675
      Picture         =   "AxGridGroup.ctx":0000
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   195
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox pic1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   3345
      Picture         =   "AxGridGroup.ctx":04DA
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   180
      Visible         =   0   'False
      Width           =   135
   End
   Begin AxioGrid.AxGrid fG 
      Height          =   2295
      Left            =   75
      TabIndex        =   1
      Top             =   870
      Width           =   4080
      _ExtentX        =   7197
      _ExtentY        =   4048
      EnterKeyBehaviour=   0
      BackColorAlternate=   0
      GridLinesFixed  =   2
      BackColorFixed  =   -2147483626
      Cols            =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridColorFixed  =   8421504
      MouseIcon       =   "AxGridGroup.ctx":09B4
      Rows            =   10
   End
   Begin VB.PictureBox picGroup 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   0
      Left            =   225
      ScaleHeight     =   330
      ScaleWidth      =   1140
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   225
      Visible         =   0   'False
      Width           =   1140
   End
End
Attribute VB_Name = "AxGridGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'--------------------------------------------------------
' API declarations

Private Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Private Const DFC_BUTTON = 4
Private Const DFCS_BUTTONPUSH = &H10

Private Declare Function DrawFrameControl Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long


'--------------------------------------------------------
' private declarations

Private Type POINTSGL
    X As Single
    Y As Single
End Type

Private Type GROUPINFO
    ctl As PictureBox
    text As String
End Type

Private Const CLR_BTNFACE = &H8000000F
Private Const CLR_BTNSHADOW = &H80000010
Private Const CLR_BTNHILITE = &H80000014

Private Const HELPMSG = " Arrastre una Columna aquí para agrupar... "
Private Const DRAG_TOLERANCE = 100 ' Twips

'--------------------------------------------------------
' variables

' mouse control
Private m_bCapture As Boolean   ' mouse captured?
Private m_bDragging As Boolean  ' dragging control?
Private m_ptDown As POINTSGL    ' where was the click
Private m_ptControl As POINTSGL ' original coordinates

Private m_iGroups As Integer    ' how many groups do we have
Private m_GroupInfo() As GROUPINFO ' group information vector


Private Function FindColumn(s$) As Integer
    
    ' locate column based on header text
    Dim i%
    For i = 0 To fG.Cols - 1
        If fG.TextMatrix(0, i) = s Then
            FindColumn = i
            Exit Function
        End If
    Next
    
    ' this should never happen
    FindColumn = -1

End Function

Sub HideCols(fG As AxGrid, iCol As Integer, bHide As Boolean)
    Static arrColWith() As Long
    Static arrHideCol() As Boolean
    Dim i As Integer
    
ReDim arrColWith(fG.Cols - 1)
ReDim arrHideCol(fG.Cols - 1)

            For i = 0 To fG.Cols - 1
                arrHideCol(i) = False
                arrColWith(i) = fG.ColWidth(i)
            Next
             
            If bHide = False Then
                If fG.ColWidth(iCol) = 0 Then fG.ColWidth(iCol) = arrColWith(iCol) ' Restaurar el ancho
                arrHideCol(iCol) = False
            Else
                fG.ColWidth(iCol) = 0 ' ocultar
            End If
    
End Sub

Sub HideRows(fG As AxGrid, iRow As Integer, bHide As Boolean)
    Static arrRowHeight() As Long
    Static arrHideRow() As Boolean
    Dim i As Integer
    
ReDim arrRowHeight(fG.Rows - 1)
ReDim arrHideRow(fG.Rows - 1)

            For i = 0 To fG.Rows - 1
                arrHideRow(i) = False
                arrRowHeight(i) = fG.RowHeight(i)
            Next
             
            If bHide = False Then
                If fG.RowHeight(iRow) = 0 Then fG.RowHeight(iRow) = arrRowHeight(iRow) ' Restaurar el ancho
                arrHideRow(iRow) = False
            Else
                fG.RowHeight(iRow) = 0 ' ocultar
            End If
    
End Sub


Private Sub UpdateGrid()
On Error Resume Next

    ' redraw is off to speed things up
    fG.Redraw = False
    
    ' move groups to left
    Dim i%, col%
    For i = 0 To m_iGroups - 1
        col = FindColumn(m_GroupInfo(i).text)
        fG.ColPosition(col) = i
    Next
    
    ' hide groups, make sure they're all sortable
    For i = 0 To m_iGroups - 1
        Call HideCols(fG, i, True)
        If fG.SortColumnMode = 0 Then fG.SortColumnMode = flexSortGenericAscending
    Next
    
    ' show non-groups
    For i = m_iGroups To fG.Cols - 1
        Call HideCols(fG, i, False)
    Next
    
    ' sort
    'fG.Select fG.Row, 0, fG.Row, fG.Cols - 1
    'fG.Sort = flexSortUseColSort

    ' create groups
    'fG.Subtotal flexSTClear
    
    If m_iGroups > 0 Then
        
        For i = 0 To m_iGroups - 1
            'fG.Subtotal flexSTNone, i, , , CLR_BTNFACE, , , , , True
        Next
        
        ' group them
        'fG.Outline m_iGroups - 1
        'fG.OutlineCol = m_iGroups
        'fG.AutoSize m_iGroups
        
    End If
    
    ' move text to visible rows
    If m_iGroups > 0 Then
        For i = 1 To fG.Rows - 1
            'If fG.IsSubtotal(i) Then
            '    Dim s$
            '    s = fG.TextMatrix(i, 0)
            '    fG.TextMatrix(i, 0) = ""
            '    fG.TextMatrix(i, m_iGroups) = s
            'End If
        Next
    End If
    'fG.MergeCells = flexMergeSpill

    ' redraw is back on
    fG.Redraw = True
    
End Sub


Private Sub fG_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' if we clicked on a column, start dragging it
    If Button = 1 And Shift = 0 And fG.MouseRow = 0 Then
        
        ' make sure we don't group on everything
        If m_iGroups >= fG.Cols - 1 Then
            Exit Sub
        End If
        
        ' which column are we grouping on?
        Dim col%
        col = fG.MouseCol
        
        ' confirm that this is a groupable column
        Dim i%
        For i = 0 To m_iGroups - 1
            If m_GroupInfo(i).text = fG.TextMatrix(0, col) Then
                'Columna ya agrupada, salir!
                Beep
                Exit Sub
            End If
        Next
        ' UNDONE
        
        ' create entry in global array
        i = m_iGroups
        m_iGroups = m_iGroups + 1
        ReDim Preserve m_GroupInfo(i)
        
        ' create new group control
        Static newCtl%
        newCtl = newCtl + 1
        Load picGroup(newCtl)
        Set m_GroupInfo(i).ctl = picGroup(newCtl)
        m_GroupInfo(i).text = fG.TextMatrix(0, col)
        
        ' init group control
        With picGroup(newCtl)
            .Tag = i
            .Width = .TextWidth(m_GroupInfo(i).text) + 2 * fG.RowHeight(0)
            .Height = fG.RowHeight(0) * 1.1
            .Move fG.ColPos(col), fG.top
            .Font = fG.Font
            .ZOrder
        End With
        
        ' save original position (none in this case)
        m_ptControl.X = -1
        m_ptControl.Y = -1
        
        ' start dragging
        m_bCapture = True
        m_bDragging = True
        m_ptDown.X = X - picGroup(newCtl).left
        m_ptDown.Y = fG.top + Y - picGroup(newCtl).top
        picGroup_Paint newCtl
        
        ' this is really cool:
        ' flex got the mouse down, but we want the group control to handle it
        ' so we set Cancel to true and transfer the mouse to the group control
        ' using the SetCapture API.
        'Cancel = True
        With picGroup(newCtl)
            .Visible = True
            .SetFocus
            SetCapture .hwnd
        End With
    End If

End Sub

Private Sub picGroup_Click(Index As Integer)

    ' unless we were dragging, revert sort direction
    If (Not m_bDragging) And (m_ptControl.X > -1) Then
        
        ' revert sort direction
        Dim i%
        i = picGroup(Index).Tag
        If fG.SortColumnMode = flexSortGenericDescending Then
            fG.SortColumnMode = flexSortGenericAscending
        Else
            fG.SortColumnMode = flexSortGenericDescending
        End If
        
        ' show the change
        UpdateLayout True
        
    End If
End Sub

Private Sub picGroup_KeyPress(Index As Integer, KeyAscii As Integer)
    
    ' escape cancels dragging/clicking
    If (KeyAscii = 27) And (m_bCapture = True) Then
        
        ' move control back to its original position
        If m_bDragging Then
        
            ' if the group was still being created (not just dragged), delete it
            If m_ptControl.X < 0 And m_ptControl.Y < 0 Then
                DeleteGroup Index
            
            ' otherwise, move it back to where it was
            Else
                picGroup(Index).Move m_ptControl.X, m_ptControl.Y
            End If
        End If
        
        ' reset state variables
        m_bCapture = False
        m_bDragging = True
    
    End If
    
End Sub


Private Sub DeleteGroup(Index As Integer)
    
    ' remove control from the list
    Dim i%, j%
    i = picGroup(Index).Tag
    For j = i To m_iGroups - 2
        m_GroupInfo(j) = m_GroupInfo(j + 1)
    Next
    m_iGroups = m_iGroups - 1
    
    'If m_iGroups = 0 Then fG.Outline 1

    ' hide/unload the control
    picGroup(Index).Visible = False
    If Index > 0 Then Unload picGroup(Index)
    
End Sub

Private Sub picGroup_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    ' left button starts dragging
    If Button = 1 Then
    
        ' save dragging information
        m_bCapture = True
        m_bDragging = False
        m_ptDown.X = X
        m_ptDown.Y = Y
        
        ' bring control to top, save its original position
        picGroup(Index).ZOrder
        m_ptControl.X = picGroup(Index).left
        m_ptControl.Y = picGroup(Index).top
    End If

End Sub

Private Sub picGroup_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    ' drag control around
    If m_bCapture Then
        With picGroup(Index)
                        
            ' if we are not dragging yet, maybe it's time to start
            If Not m_bDragging Then
                If Abs(X - m_ptDown.X) > DRAG_TOLERANCE Then m_bDragging = True
                If Abs(Y - m_ptDown.Y) > DRAG_TOLERANCE Then m_bDragging = True
            End If
            
            ' if we're dragging, then do it
            If m_bDragging Then
            
                ' get new coordinates
                X = .left + (X - m_ptDown.X)
                Y = .top + (Y - m_ptDown.Y)
                
                ' restrict boundaries
                If X < 0 Then X = 0
                If Y < 0 Then Y = 0
                If X > UserControl.ScaleWidth - .Width Then X = UserControl.ScaleWidth - .Width
                If Y > UserControl.ScaleHeight - .Height Then Y = UserControl.ScaleHeight - .Height
                If Y > fG.top Then Y = fG.top
            
                ' move the control
                .Move X, Y
                
                ' show where we'd go if we dropped now
                ' UNDONE
                
            End If
        End With
    End If
End Sub

Private Sub picGroup_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ' if we were dragging,
    ' we may have just moved the group to a new position, or
    ' we may have dropped it back into the grid
    If m_bDragging Then
        
        fG.Redraw = False
        
        ' back into grid, different position
        Y = picGroup(Index).top + Y
        If Y > fG.top Then
            
            ' see which column it was and where the mouse is
            Dim col%, i%
            col = FindColumn(m_GroupInfo(picGroup(Index).Tag).text)
            i = fG.MouseCol
            
            ' different? move column
            If i <> col Then
                fG.ColPosition(col) = i
            
            ' same? switch sort order
            Else
                If fG.SortColumnMode = flexSortGenericAscending Then
                    fG.SortColumnMode = flexSortGenericDescending
                Else
                    fG.SortColumnMode = flexSortGenericAscending
                End If
            End If
            
            ' remove our brand-new group
            DeleteGroup Index
        
        End If
        
        ' either way, show changes
        UpdateLayout True
        
        fG.Redraw = True
    End If

    ' cancel capture no matter what
    m_bCapture = False

End Sub


Private Sub picGroup_Paint(Index As Integer)
On Error Resume Next
    Dim rc As RECT
    
    With picGroup(Index)
        
        ' draw frame
        rc.top = 0
        rc.left = 0
        rc.right = .Width / Screen.TwipsPerPixelX
        rc.bottom = .Height / Screen.TwipsPerPixelY
        DrawFrameControl .hDC, rc, DFC_BUTTON, DFCS_BUTTONPUSH
        
        ' draw text
        .CurrentX = .TextWidth(" ")
        .CurrentY = (.Height - .TextHeight(" ")) / 2.5
        picGroup(Index).Print m_GroupInfo(.Tag).text
        
        ' draw sort arrow if this is a group already
        If fG.ColWidth(.Tag) = 0 Then
            Dim X As Single, Y As Single, sz As Single
            sz = .Height * (1 / 3)
            X = .Width - sz
            
            ' pointing up
            If fG.SortColumnMode = flexSortGenericDescending Then
                Y = (.Height - sz) / 2 + sz
                picGroup(Index).Line (X, Y)-(X - sz, Y), CLR_BTNHILITE
                picGroup(Index).Line -(X - sz / 2, Y - sz), CLR_BTNSHADOW
                picGroup(Index).Line -(X, Y), CLR_BTNHILITE
            
            ' pointing down
            Else
                Y = (.Height - sz) / 2
                picGroup(Index).Line (X, Y)-(X - sz, Y), CLR_BTNSHADOW
                picGroup(Index).Line -(X - sz / 2, Y + sz), CLR_BTNSHADOW
                picGroup(Index).Line -(X, Y), CLR_BTNHILITE
            End If
        End If
    End With

End Sub


Private Sub UserControl_Initialize()
    
    ' initialize embedded FlexGrid
    fG.SelectionMode = flexSelectionByRow
    fG.AllowUserResizing = flexResizeColumns

    
    ' initialize group control based on grid data
    With picGroup(0)
        .Font = fG.Font
        .Height = fG.RowHeight(0)
        .Tag = 0
    End With

End Sub
Private Sub UpdateLayout(doGrid As Boolean)
    
    Dim swap As GROUPINFO
    Dim i%, cnt%, done%
    Dim X As Single, Y As Single, rh As Single
    Dim offsety As Single
    
    ' see how many groups are visible
    cnt = m_iGroups
    
    ' dimension and clear grouping area
    rh = fG.RowHeight(0)
    offsety = rh / 2
    Y = 2 * fG.RowHeight(0)
    If cnt > 1 Then Y = Y + (cnt - 1) * offsety
    Y = UserControl.ScaleHeight - Y
    If Y < 0 Then Y = 0
    fG.Height = Y
    UserControl.Cls
    
    ' if no groups, show helpful message
    If cnt = 0 Then
        UserControl.CurrentX = rh / 2
        UserControl.CurrentY = rh / 2
        UserControl.Print HELPMSG
    End If
    
    ' sort group vector by position (left-to-right)
    While Not done
        done = True
        For i = 0 To cnt - 2
            If m_GroupInfo(i).ctl.left > m_GroupInfo(i + 1).ctl.left Then
                done = False
                swap = m_GroupInfo(i)
                m_GroupInfo(i) = m_GroupInfo(i + 1)
                m_GroupInfo(i + 1) = swap
            End If
        Next
    Wend
    
    ' each control gets and index into the vector
    For i = 0 To cnt - 1
        m_GroupInfo(i).ctl.Tag = i
    Next
    
    ' position group controls
    Y = rh / 2
    X = Y
    For i = 0 To cnt - 1
        With m_GroupInfo(i).ctl
        
            ' move the control
            .Move X, Y
            Y = Y + offsety
            X = X + .Width + rh / 3
        
            ' draw connector
            If i < cnt - 1 Then
                UserControl.Line (X, Y + 2 / 3 * rh)-(X - rh * 2 / 3, Y + 2 / 3 * rh), 0
                UserControl.Line -(X - rh * 2 / 3, Y + rh / 2 - Screen.TwipsPerPixelY), 0
            End If
    
            ' draw placeholder
            UserControl.Line (.left, .top)-(.left + .Width - Screen.TwipsPerPixelX, .top + .Height - Screen.TwipsPerPixelY), 0, B
        
        End With
    Next
    
    ' redraw all controls at their new positions
    For i = 0 To cnt - 1
        picGroup_Paint m_GroupInfo(i).ctl.Index
    Next
    UserControl.Refresh
    
    ' update the grid
    If doGrid Then UpdateGrid
    
    ' redraw all controls at their new positions (to show sort direction)
    For i = 0 To cnt - 1
        picGroup_Paint m_GroupInfo(i).ctl.Index
    Next
    
End Sub

Private Sub UserControl_Resize()
  With fG
    .left = 0
    .top = 550
    .Height = UserControl.Height - 550
    .Width = UserControl.Width
  End With
    
    UpdateLayout False
    
End Sub



Public Property Get AxGrid() As AxGrid
    Set AxGrid = fG
End Property


Public Sub Update()
    
    UpdateLayout True
    
End Sub

