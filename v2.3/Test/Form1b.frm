VERSION 5.00
Object = "{56C658AA-DA75-4863-A247-648E5C2ACED0}#14.0#0"; "AXGRIDKM.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   11385
   ClientLeft      =   4965
   ClientTop       =   2745
   ClientWidth     =   13815
   LinkTopic       =   "Form1"
   ScaleHeight     =   11385
   ScaleWidth      =   13815
   Begin AxioGrid.AxGrid AxGrid1 
      Height          =   3525
      Left            =   105
      TabIndex        =   61
      Top             =   105
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   6218
      EnterKeyBehaviour=   0
      BackColorAlternate=   0
      GridLinesFixed  =   2
      Appearance      =   0
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
      MouseIcon       =   "Form1b.frx":0000
      Rows            =   10
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Controls Cols 1-2"
      Height          =   360
      Left            =   4785
      TabIndex        =   50
      Top             =   3810
      Width           =   1335
   End
   Begin VB.CheckBox chkAltColor 
      Caption         =   "Alternate Color"
      Height          =   210
      Left            =   270
      TabIndex        =   26
      Top             =   3960
      Width           =   1635
   End
   Begin VB.CheckBox chkEditable 
      Caption         =   "Editable"
      Height          =   210
      Left            =   255
      TabIndex        =   25
      Top             =   3705
      Width           =   1635
   End
   Begin VB.CheckBox chkMove 
      Caption         =   "On Enter Key Move Down"
      Height          =   270
      Left            =   2160
      TabIndex        =   24
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Caption         =   "AxBiGrid"
      Height          =   4785
      Left            =   15
      TabIndex        =   23
      Top             =   5175
      Width           =   11715
      Begin AxioGrid.AxBiGrid AxBiGrid1 
         Height          =   3180
         Left            =   120
         TabIndex        =   62
         Top             =   240
         Width           =   7395
         _ExtentX        =   13044
         _ExtentY        =   5609
         SplitterPos     =   2465
         EnterKeyBehaviour=   0
         ShowInfoBar     =   0   'False
         ColsLeft        =   3
         ColsRight       =   5
         GridLinesFixed  =   2
         Appearance      =   0
         BorderStyle     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Form1b.frx":001C
         RowHeightMin    =   315
         Rows            =   10
         GridLinesFixed  =   2
         Appearance      =   0
         BorderStyle     =   0
         FixedCols       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Form1b.frx":0038
         RowHeightMin    =   315
         Rows            =   10
      End
      Begin VB.TextBox txtinfoBar 
         Height          =   315
         Left            =   7650
         TabIndex        =   60
         Text            =   "Text7"
         Top             =   2850
         Width           =   1725
      End
      Begin VB.ListBox lstTextGrid 
         Height          =   840
         Left            =   7650
         TabIndex        =   59
         Top             =   1995
         Width           =   1725
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Flat Style"
         Height          =   360
         Left            =   7650
         TabIndex        =   58
         Top             =   1560
         Width           =   1710
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Fix Splitter"
         Height          =   300
         Left            =   4785
         TabIndex        =   57
         Top             =   4200
         Width           =   990
      End
      Begin VB.TextBox txtSplitPos 
         Height          =   285
         Left            =   4050
         TabIndex        =   55
         Top             =   4215
         Width           =   720
      End
      Begin VB.CommandButton Command9 
         Caption         =   "AddItem Right"
         Height          =   495
         Index           =   1
         Left            =   8505
         TabIndex        =   54
         Top             =   1035
         Width           =   855
      End
      Begin VB.CommandButton Command9 
         Caption         =   "AddItem Left"
         Height          =   495
         Index           =   0
         Left            =   7650
         TabIndex        =   53
         Top             =   1035
         Width           =   855
      End
      Begin VB.CheckBox Check4 
         Caption         =   "SelectionMode Free"
         Height          =   270
         Left            =   195
         TabIndex        =   52
         Top             =   4350
         Width           =   2415
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Controls Cols 1-2"
         Height          =   360
         Left            =   7650
         TabIndex        =   51
         Top             =   285
         Width           =   1710
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Clear"
         Height          =   360
         Left            =   7650
         TabIndex        =   49
         Top             =   645
         Width           =   1710
      End
      Begin VB.CommandButton cmdLoadDB2 
         Caption         =   "LoadfromDB"
         Height          =   345
         Index           =   2
         Left            =   9480
         TabIndex        =   48
         Top             =   1890
         Width           =   1080
      End
      Begin VB.TextBox Text6 
         Height          =   315
         Left            =   4050
         TabIndex        =   43
         Top             =   3510
         Width           =   720
      End
      Begin VB.TextBox Text5 
         Height          =   315
         Left            =   6240
         TabIndex        =   42
         Top             =   3510
         Width           =   1170
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   6240
         TabIndex        =   41
         Top             =   3840
         Width           =   1170
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   4050
         TabIndex        =   40
         Top             =   3840
         Width           =   720
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Change 2nd Row BackColor"
         Height          =   510
         Index           =   1
         Left            =   9495
         TabIndex        =   39
         Top             =   165
         Width           =   1980
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Change 3nd Row CellAlignment"
         Height          =   510
         Index           =   1
         Left            =   9480
         TabIndex        =   38
         Top             =   675
         Width           =   1980
      End
      Begin VB.CommandButton cmdAutosize 
         Caption         =   "Autosize 3 first Columns"
         Height          =   345
         Index           =   1
         Left            =   9480
         TabIndex        =   37
         Top             =   1200
         Width           =   1980
      End
      Begin VB.CommandButton cmdSum 
         Caption         =   "Sumar Matrix"
         Height          =   345
         Index           =   1
         Left            =   9495
         TabIndex        =   36
         Top             =   1545
         Width           =   1980
      End
      Begin VB.CommandButton cmdCreateTable 
         Caption         =   "CREATE TABLE"
         Height          =   345
         Index           =   1
         Left            =   9480
         TabIndex        =   35
         Top             =   2235
         Width           =   1980
      End
      Begin VB.CommandButton cmdInsert 
         Caption         =   "INSERT"
         Height          =   345
         Index           =   1
         Left            =   9480
         TabIndex        =   34
         Top             =   3285
         Width           =   1980
      End
      Begin VB.CommandButton cmdDrpTable 
         Caption         =   "DROP TABLE"
         Height          =   345
         Index           =   1
         Left            =   9480
         TabIndex        =   33
         Top             =   2580
         Width           =   1980
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "SELECT"
         Height          =   345
         Index           =   1
         Left            =   9480
         TabIndex        =   32
         Top             =   2925
         Width           =   1980
      End
      Begin VB.CommandButton cmdLoadDB2 
         Caption         =   "to Controls"
         Height          =   345
         Index           =   3
         Left            =   10575
         TabIndex        =   31
         Top             =   1890
         Width           =   885
      End
      Begin VB.CommandButton cmdInsrtSel 
         Caption         =   "INSERT SELECT"
         Height          =   345
         Index           =   1
         Left            =   9480
         TabIndex        =   30
         Top             =   3630
         Width           =   1980
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Alternate Color"
         Height          =   210
         Left            =   195
         TabIndex        =   29
         Top             =   3885
         Width           =   1635
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Editable"
         Height          =   210
         Left            =   195
         TabIndex        =   28
         Top             =   3630
         Width           =   1635
      End
      Begin VB.CheckBox Check1 
         Caption         =   "On Enter Key Move Down"
         Height          =   270
         Left            =   195
         TabIndex        =   27
         Top             =   4095
         Width           =   2415
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Splitter Position"
         Height          =   195
         Left            =   2910
         TabIndex        =   56
         Top             =   4260
         Width           =   1110
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sumar Columna"
         Height          =   195
         Left            =   2820
         TabIndex        =   47
         Top             =   3555
         Width           =   1110
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CalculateColumn"
         Height          =   195
         Left            =   4965
         TabIndex        =   46
         Top             =   3555
         Width           =   1185
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CalculateRow"
         Height          =   195
         Left            =   5160
         TabIndex        =   45
         Top             =   3930
         Width           =   990
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Multiplicar Fila"
         Height          =   195
         Left            =   2910
         TabIndex        =   44
         Top             =   3885
         Width           =   1110
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "..."
      Height          =   345
      Left            =   6240
      TabIndex        =   22
      Top             =   3840
      Width           =   1980
   End
   Begin VB.CommandButton Command6 
      Caption         =   "..."
      Height          =   345
      Index           =   0
      Left            =   6240
      TabIndex        =   21
      Top             =   4185
      Width           =   1980
   End
   Begin VB.CommandButton Command5 
      Caption         =   "..."
      Height          =   345
      Index           =   0
      Left            =   6240
      TabIndex        =   20
      Top             =   4545
      Width           =   1980
   End
   Begin VB.CommandButton cmdInsrtSel 
      Caption         =   "INSERT SELECT"
      Height          =   345
      Index           =   0
      Left            =   6240
      TabIndex        =   19
      Top             =   3495
      Width           =   1980
   End
   Begin VB.CommandButton cmdLoadDB2 
      Caption         =   "with Index"
      Height          =   345
      Index           =   1
      Left            =   7335
      TabIndex        =   18
      Top             =   1756
      Width           =   885
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "SELECT"
      Height          =   345
      Index           =   0
      Left            =   6240
      TabIndex        =   17
      Top             =   2790
      Width           =   1980
   End
   Begin VB.CommandButton cmdDrpTable 
      Caption         =   "DROP TABLE"
      Height          =   345
      Index           =   0
      Left            =   6240
      TabIndex        =   16
      Top             =   2445
      Width           =   1980
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "INSERT"
      Height          =   345
      Index           =   0
      Left            =   6240
      TabIndex        =   15
      Top             =   3150
      Width           =   1980
   End
   Begin VB.TextBox txtRow 
      Height          =   285
      Left            =   1440
      TabIndex        =   13
      Top             =   4755
      Width           =   720
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4605
      TabIndex        =   12
      Top             =   4755
      Width           =   1170
   End
   Begin VB.CommandButton cmdLoadDB2 
      Caption         =   "LoadfromDB"
      Height          =   345
      Index           =   0
      Left            =   6240
      TabIndex        =   10
      Top             =   1755
      Width           =   1080
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   4605
      TabIndex        =   8
      Top             =   4425
      Width           =   1170
   End
   Begin VB.TextBox txtCOL 
      Height          =   315
      Left            =   1440
      TabIndex        =   7
      Top             =   4425
      Width           =   720
   End
   Begin VB.CommandButton cmdCreateTable 
      Caption         =   "CREATE TABLE"
      Height          =   345
      Index           =   0
      Left            =   6240
      TabIndex        =   5
      Top             =   2100
      Width           =   1980
   End
   Begin VB.CommandButton cmdSum 
      Caption         =   "Sumar Matrix"
      Height          =   345
      Index           =   0
      Left            =   6240
      TabIndex        =   4
      Top             =   1407
      Width           =   1980
   End
   Begin VB.CommandButton cmdAutosize 
      Caption         =   "Autosize 3 first Columns"
      Height          =   345
      Index           =   0
      Left            =   6240
      TabIndex        =   3
      Top             =   1058
      Width           =   1980
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Change 3nd Row CellAlignment"
      Height          =   510
      Index           =   0
      Left            =   6240
      TabIndex        =   2
      Top             =   544
      Width           =   1980
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change 2nd Row BackColor"
      Height          =   510
      Index           =   0
      Left            =   6240
      TabIndex        =   1
      Top             =   30
      Width           =   1980
   End
   Begin VB.CheckBox chkOther 
      Caption         =   "Other???"
      Height          =   240
      Left            =   2160
      TabIndex        =   0
      Top             =   3735
      Width           =   2415
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Multiplicar Fila"
      Height          =   195
      Left            =   210
      TabIndex        =   14
      Top             =   4845
      Width           =   1110
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Resultado CalculateRow"
      Height          =   195
      Left            =   2550
      TabIndex        =   11
      Top             =   4845
      Width           =   1950
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Resultado CalculateColumn"
      Height          =   195
      Left            =   2550
      TabIndex        =   9
      Top             =   4470
      Width           =   1950
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sumar Columna"
      Height          =   195
      Left            =   210
      TabIndex        =   6
      Top             =   4470
      Width           =   1110
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Conn As String

Private Sub AxBiGrid1_ButtonClick(xGrid As AxioGrid.eSideGrid, ByVal Row As Long, ByVal Col As Long)
MsgBox xGrid & " " & Row & " " & Col
End Sub

Private Sub AxBiGrid1_CellTextChange(xGrid As AxioGrid.eSideGrid, ByVal Row As Long, ByVal Col As Long)
If AxBiGrid1.ColObject(eLeftGrid, 2) = eTextBoxColumn Then
  AxBiGrid1.SetColObject(eLeftGrid, 2, True) = eComboBoxColumn
End If
End Sub

Private Sub AxBiGrid1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
txtSplitPos.Text = AxBiGrid1.SplitterPos
End Sub

Private Sub Check1_Click()
    If Check1.Value = vbChecked Then
       AxBiGrid1.EnterKeyBehaviour = axEKMoveDown
    Else
       AxBiGrid1.EnterKeyBehaviour = axEKMoveRight
    End If

End Sub

Private Sub Check2_Click()
   If Check2.Value = vbChecked Then
      AxBiGrid1.Editable = True
   Else
      AxBiGrid1.Editable = False
   End If

End Sub

Private Sub Check3_Click()
   If Check3.Value = vbChecked Then
      AxBiGrid1.BackColorAlternate = &HFFFFC0
   Else
      AxBiGrid1.BackColorAlternate = vbWhite
   End If
End Sub

Private Sub Check4_Click()
If Check4.Value = vbChecked Then
  AxBiGrid1.SelectionMode = flexSelectionFree
Else
  AxBiGrid1.SelectionMode = flexSelectionByRow
End If
End Sub

Private Sub chkAltColor_Click()
   If chkAltColor.Value = vbChecked Then
      AxGrid1.BackColorAlternate = vbCyan
   Else
      AxGrid1.BackColorAlternate = vbWhite
   End If
End Sub

Private Sub chkEditable_Click()
   If chkEditable.Value = vbChecked Then
      AxGrid1.Editable = True
   Else
      AxGrid1.Editable = False
   End If
End Sub

Private Sub chkMove_Click()
    If chkMove.Value = vbChecked Then
       AxGrid1.EnterKeyBehaviour = axEKMoveDown
    Else
       AxGrid1.EnterKeyBehaviour = axEKMoveRight
    End If
End Sub

Private Sub cmdDrpTable_Click(Index As Integer)
AxGrid1.ADODropTable "MiTabla"
End Sub

Private Sub cmdInsert_Click(Index As Integer)
Dim sFields As String, I As Integer

For I = 1 To AxGrid1.Cols - 1
  sFields = sFields & AxGrid1.TextMatrix(0, I) & ","
Next I

If Right$(sFields, 1) = "," Then
  sFields = Mid$(sFields, 1, Len(sFields) - 1)
End If

With AxGrid1
  .ADOTable = "MiTabla"
  .ADOFields = sFields
  .ADOInsert AxGrid1.Row
End With

End Sub

Private Sub cmdSelect_Click(Index As Integer)
With AxGrid1
  .ADOTable = "Unidad"
  .ADOFields = "*"
  .ADOSelect "CANTIDAD", "7"
End With
End Sub

Private Sub cmdSum_Click(Index As Integer)
If Index = 0 Then
    Text1.Text = AxGrid1.CalculateMatrix(axSTSum, 1, 2, 8, 2)
Else
    Text5.Text = AxBiGrid1.CalculateColumn(eRightGrid, axSTSum, 2, 1, 10)
End If
End Sub

Private Sub Command1_Click(Index As Integer)
If Index = 0 Then
     AxGrid1.Cell(axcpCellBackColor, 2, 1, 2, AxGrid1.Cols - 1) = vbBlue
Else
    AxBiGrid1.Cell(axcpCellBackColor, 2, 0, 2, AxBiGrid1.ColsRight - 1, eRightGrid) = vbBlue
End If
End Sub

Private Sub Command10_Click()
AxBiGrid1.SplitterFixed = Not AxBiGrid1.SplitterFixed
End Sub

Private Sub Command11_Click()
If AxBiGrid1.Appearance = flex3D Then
  AxBiGrid1.Appearance = flexFlat
  Command11.Caption = "3D Style"
Else
  AxBiGrid1.Appearance = flex3D
  Command11.Caption = "Flat Style"
End If
End Sub

Private Sub Command2_Click(Index As Integer)
If Index = 0 Then
     AxGrid1.Cell(axcpCellAlignment, 3, 1, 3, AxGrid1.Cols - 1) = 3
Else
    AxBiGrid1.Cell(axcpCellAlignment, 3, 1, 3, AxBiGrid1.ColsRight - 1, eRightGrid) = 3
End If
End Sub

Private Sub cmdAutosize_Click(Index As Integer)
If Index = 0 Then
    AxGrid1.AutoSizeMode = axAutoSizeColWidth
    AxGrid1.AutoSizeCols 0, 2

Else
    AxBiGrid1.AutoSizeMode = axAutoSizeColWidth
    AxBiGrid1.AutoSizeCols eRightGrid, 0, 2
End If
End Sub

Private Sub cmdCreateTable_Click(Index As Integer)

With AxGrid1
  .ADOConnection = Conn
  If .ADOCreateTable("MiTabla") = True Then
    MsgBox "Tabla Creada!"
  Else
    MsgBox "Error al Crear Tabla!"
  End If
End With

End Sub

End Sub

Private Sub Command3_Click()
AxBiGrid1.SetColObject(eLeftGrid, 1, False) = eTextBoxColumn
AxBiGrid1.SetColObject(eLeftGrid, 2, False) = eButtonColumn
AxBiGrid1.SetColObject(eRightGrid, 1, False) = eComboBoxColumn
AxBiGrid1.SetColObject(eRightGrid, 2, False) = eListBoxColumn
End Sub

Private Sub Command4_Click()
With AxBiGrid1
  .ClearGrid BothGrids
  .Rows = 2
  .SetColObject(eLeftGrid, 2, False) = eTextBoxColumn
End With
End Sub

Private Sub Command8_Click()
Dim I As Integer
For I = 0 To 10
  AxGrid1.AddItemObject "Item " & I
Next I

AxGrid1.SetColObject(1) = eTextBoxColumn
AxGrid1.SetColObject(2) = eButtonColumn
AxGrid1.SetColObject(3) = eComboBoxColumn
AxGrid1.SetColObject(4) = eListBoxColumn
End Sub

Private Sub Command9_Click(Index As Integer)
Select Case Index
   Case 0
    AxBiGrid1.Row(eLeftGrid) = 1
    AxBiGrid1.AddItem eLeftGrid, vbTab & "TEST ADDITEM L"
   Case 1
    AxBiGrid1.Row(eRightGrid) = 1
    AxBiGrid1.AddItem eRightGrid, "TEST ADDITEM R"
End Select
End Sub

Private Sub Form_Load()
Dim I As Integer
With AxGrid1
  .TextMatrix(0, 1) = "Name"
  .TextMatrix(0, 2) = "Salary"
  .ColDisplayFormat(2) = "#0.00"
  .Cols = 4
  .Rows = 12
  .ColWidth(0) = 600
    
  For I = 1 To .Rows - 1
    .TextMatrix(I, 1) = "ABC_" & I
    .TextMatrix(I, 2) = I * 3.05
  Next I
    
  'Implemented only for Numeric Entry
  .ColInputMask(2) = "000.00"
  .TextMatrix(9, 0) = "Total"
  
End With
   
With AxBiGrid1
  .TextMatrix(eLeftGrid, 0, 1) = "Name"
  .TextMatrix(eLeftGrid, 0, 2) = "Salary"
  .ColDisplayFormat(2) = "#0.00"
  .Colsleft = 3
  .ColsRight = 5
  .Rows = 15
  .ColWidth(BothGrids, 0) = 500
  .ColWidth(LeftGrid, 1) = 1800
  .ColWidth(RightGrid, 1) = 1800
  
  For I = 1 To .Rows - 1
    .TextMatrix(eLeftGrid, I, 1) = "NOMBRE_" & I
    .TextMatrix(eLeftGrid, I, 2) = I * 103.5
    .TextMatrix(eRightGrid, I, 1) = "ABC_" & I
    .TextMatrix(eRightGrid, I, 2) = I * 3.05
  Next I
    
  'Implemented only for Numeric Entry
  .ColInputMask(2) = "000.00"
  .TextMatrix(eRightGrid, 9, 0) = "Total"
  
End With

With lstTextGrid
    .AddItem "BothGrids"
    .AddItem "LeftGrids"
    .AddItem "RightGrids"
    .AddItem "Custom"
End With

Conn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & App.Path & "\mibase2013.accdb;Persist Security Info=False;"
AxGrid1.ADOConnection = Conn
AxBiGrid1.ADOConnection = Conn

End Sub

Private Sub AxGrid1_BeforeEdit(Cancel As Boolean)
     If AxGrid1.Col = 1 Then
        If chkOther.Value = vbUnchecked Then
            Cancel = True
        End If
     End If
End Sub

Private Sub lstTextGrid_Click()
AxBiGrid1.SetInfoBar = lstTextGrid.ListIndex
If lstTextGrid.ListIndex = 3 Then AxBiGrid1.InfoBarText = txtinfoBar.Text
End Sub

Private Sub txtCOL_Change()
On Error Resume Next
Text1.Text = AxGrid1.CalculateColumn(axSTSum, txtCOL.Text, 1, AxGrid1.Rows - 1)
End Sub

Private Sub cmdLoadDB2_Click(Index As Integer)
Dim Consulta As String

Consulta = "SELECT * FROM Unidad"

Select Case Index
  Case 0
    'AxGrid1.LoadFromDB Consulta, Conn, False, True
    With AxGrid1
      .ADOFields = "*"
      .ADOTable = "Unidad"
      .ADOSelect False
    
    End With
  
  Case 1
    AxGrid1.LoadFromDB Consulta, Conn, True, True
    
  Case 2
    With AxBiGrid1
      .ADOTable = "Unidad"
      .ADOFields = "ID, MEDIDA"
      .ADOSelect eLeftGrid, False, False
      .ADOFields = "CANTIDAD, STRINGE"
      .ADOSelect eRightGrid, False, False
    End With
  
  Case 3
     AxBiGrid1.LoadtoRight Consulta, Conn, True
     AxBiGrid1.LoadToObject oComboBox, Consulta, "MEDIDA"
     AxBiGrid1.LoadToObject oListBox, Consulta, "STRINGE"
       
End Select
End Sub

Private Sub txtRow_Change()
On Error Resume Next
Text2.Text = AxGrid1.CalculateRow(axSTMultiply, txtRow.Text, 3, AxGrid1.Cols - 1)
End Sub

Private Sub txtSplitPos_Change()
AxBiGrid1.SplitterPos = txtSplitPos.Text
End Sub
