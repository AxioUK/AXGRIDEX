VERSION 5.00
Object = "{56C658AA-DA75-4863-A247-648E5C2ACED0}#15.0#0"; "AXGRIDKM.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{07C4EEB0-CDF8-4C06-9B2A-FCC7E35D524D}#2.0#0"; "AXSUITECTRL1.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso Orden de Compra"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10470
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   10470
   StartUpPosition =   2  'CenterScreen
   Begin AxioGrid.AxBiGrid bGrid 
      Height          =   4245
      Left            =   90
      TabIndex        =   23
      Top             =   1410
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   7488
      SplitterPos     =   3025
      EnterKeyBehaviour=   0
      ShowInfoBar     =   0   'False
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
      MouseIcon       =   "frmMain.frx":058A
      RowHeightMin    =   315
      Rows            =   2
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
      MouseIcon       =   "frmMain.frx":05A6
      RowHeightMin    =   315
      Rows            =   2
   End
   Begin ucAxSuite2014.jcbutton cmdInformes 
      Height          =   645
      Left            =   9375
      TabIndex        =   21
      Top             =   1320
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   1138
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Informes"
      Picture         =   "frmMain.frx":05C2
      PictureAlign    =   5
      UseMaskCOlor    =   -1  'True
   End
   Begin VB.TextBox txtID 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   105
      TabIndex        =   2
      Top             =   945
      Width           =   420
   End
   Begin VB.TextBox txtGuia 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   315
      Left            =   5295
      TabIndex        =   5
      Top             =   945
      Width           =   1290
   End
   Begin ucAxSuite2014.jcbutton cmdGrabar 
      Height          =   360
      Left            =   4500
      TabIndex        =   10
      Top             =   60
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   635
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Guardar OC"
      Picture         =   "frmMain.frx":0B5C
      PictureAlign    =   0
      UseMaskCOlor    =   -1  'True
   End
   Begin VB.TextBox txtNombre 
      Height          =   315
      Left            =   1860
      TabIndex        =   4
      Top             =   945
      Width           =   3405
   End
   Begin VB.TextBox txtRUT 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   540
      TabIndex        =   3
      Top             =   945
      Width           =   1290
   End
   Begin VB.TextBox txtOC 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   135
      TabIndex        =   0
      Top             =   345
      Width           =   1290
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   10470
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5775
      Width           =   10470
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   8250
         Picture         =   "frmMain.frx":10F6
         ScaleHeight     =   240
         ScaleWidth      =   870
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   405
         Width           =   870
      End
      Begin VB.PictureBox picLogo 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   120
         Picture         =   "frmMain.frx":1C38
         ScaleHeight     =   555
         ScaleWidth      =   2130
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   105
         Width           =   2130
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0077C181&
         BorderWidth     =   2
         Height          =   720
         Left            =   15
         Top             =   15
         Width           =   10455
      End
   End
   Begin MSMask.MaskEdBox mskFecha 
      Height          =   315
      Left            =   1590
      TabIndex        =   1
      Top             =   345
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd-mm-yyyy"
      Mask            =   "##-##-####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskFechaGuia 
      Height          =   315
      Left            =   6660
      TabIndex        =   6
      Top             =   945
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd-mm-yyyy"
      Mask            =   "##-##-####"
      PromptChar      =   "_"
   End
   Begin ucAxSuite2014.jcbutton cmdGrabarIn 
      Height          =   360
      Left            =   5985
      TabIndex        =   11
      Top             =   60
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   635
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Guardar Guia"
      Picture         =   "frmMain.frx":5A56
      PictureAlign    =   0
      UseMaskCOlor    =   -1  'True
   End
   Begin ucAxSuite2014.jcbutton cmdAbrir 
      Height          =   360
      Left            =   3030
      TabIndex        =   9
      Top             =   60
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   635
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Abrir OC"
      Picture         =   "frmMain.frx":5FF0
      PictureAlign    =   0
      UseMaskCOlor    =   -1  'True
   End
   Begin ucAxSuite2014.jcbutton cmdCerrar 
      Height          =   360
      Left            =   7470
      TabIndex        =   12
      Top             =   60
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   635
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Cerrar"
      Picture         =   "frmMain.frx":658A
      PictureAlign    =   0
      UseMaskCOlor    =   -1  'True
   End
   Begin ucAxSuite2014.jcbutton cmdNuevo 
      Height          =   360
      Left            =   3030
      TabIndex        =   8
      Top             =   435
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   635
      ShowFocusRect   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Nueva OC"
      Picture         =   "frmMain.frx":6B24
      PictureAlign    =   0
      UseMaskCOlor    =   -1  'True
   End
   Begin MSMask.MaskEdBox mskFechaIn 
      Height          =   315
      Left            =   7965
      TabIndex        =   7
      Top             =   945
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd-mm-yyyy"
      Mask            =   "##-##-####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Ingreso"
      Height          =   195
      Left            =   8010
      TabIndex        =   22
      Top             =   735
      Width           =   1035
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Guia/Factura"
      Height          =   195
      Left            =   5340
      TabIndex        =   20
      Top             =   735
      Width           =   1230
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Guia"
      Height          =   195
      Left            =   6705
      TabIndex        =   19
      Top             =   735
      Width           =   795
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      Height          =   195
      Left            =   1635
      TabIndex        =   18
      Top             =   135
      Width           =   435
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Proveedor"
      Height          =   195
      Left            =   300
      TabIndex        =   17
      Top             =   735
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Orden de Compra"
      Height          =   195
      Left            =   165
      TabIndex        =   16
      Top             =   135
      Width           =   1275
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'You have to have MSScripting Runtime referenced
Dim WshShell As Object


Dim Rst         As New ADODB.Recordset
Dim SQLString   As String
Dim I           As Integer

Private Sub SetProveedor()
Dim sDato As String
On Error GoTo ErrSub
If txtID <> "" And txtID <> vbNullString Then
  sDato = txtID
  SQLString = "WHERE PRID=" & sDato
  GoTo Procesar
End If
If txtRUT <> "" And txtRUT <> vbNullString Then
  sDato = txtRUT
  SQLString = "WHERE RUT='" & sDato & "'"
  GoTo Procesar
End If
If txtNombre <> "" And txtNombre <> vbNullString Then
  sDato = txtNombre
  SQLString = "WHERE NOMBRE LIKE '%" & sDato & "%'"
  GoTo Procesar
End If
  
Procesar:
  SQLString = "SELECT PRID, RUT, NOMBRE FROM PROVEEDOR " & SQLString
  Rst.Open SQLString, Cnn, adOpenDynamic, adLockOptimistic
    txtID = Rst(0)
    txtRUT = Rst(1)
    txtNombre = Rst(2)
  Rst.Close

Exit Sub
ErrSub:
'GRABAR PROVEEDOR
On Error GoTo ExitSub
If txtRUT = "" Then txtRUT.SetFocus: Exit Sub
If txtNombre = "" Then txtNombre.SetFocus: Exit Sub
If Rst.State = adStateOpen Then Rst.Close
  SQLString = "INSERT INTO PROVEEDOR (RUT, NOMBRE) VALUES ('" & txtRUT & "', '" & txtNombre & "');"
  Cnn.Execute SQLString, adCmdText
  
ExitSub:
End Sub

Private Sub bGrid_CellTextChange(xGrid As AxioGrid.eSideGrid, ByVal Row As Long, ByVal Col As Long)
With bGrid
  On Error GoTo ErrSub
  ' Columna CODIGO
  If .TextMatrix(eLeftGrid, Row, 1) <> "" And .TextMatrix(eLeftGrid, Row, 2) <> vbNullString Then
    SQLString = "SELECT PID, CODIGO, PRODUCTO FROM BASE WHERE CODIGO='" & .TextMatrix(xGrid, Row, 1) & "'"
    Rst.Open SQLString, Cnn, adOpenDynamic, adLockOptimistic
      .TextMatrix(eLeftGrid, Row, 1) = Rst(0)
      .TextMatrix(eLeftGrid, Row, 2) = Rst(1)
      .TextMatrix(eLeftGrid, Row, 3) = Rst(2)
    Rst.Close
  End If
  ' Columna PRODUCTO
  If .ColObject(eLeftGrid, 3) = eTextBoxColumn Then
    SQLString = "SELECT PID, CODIGO, PRODUCTO FROM BASE WHERE PRODUCTO LIKE '%" & .TextMatrix(xGrid, Row, 3) & "%'"
    .LoadToObject oComboBox, SQLString, "PRODUCTO"
    .SetColObject(eLeftGrid, 3, True) = eComboBoxColumn
  Else
    SQLString = "SELECT PID, CODIGO, PRODUCTO FROM BASE WHERE PRODUCTO='" & .TextMatrix(xGrid, Row, 3) & "'"
    Rst.Open SQLString, Cnn, adOpenDynamic, adLockOptimistic
      .TextMatrix(eLeftGrid, Row, 1) = Rst(0)
      .TextMatrix(eLeftGrid, Row, 2) = Rst(1)
      .TextMatrix(eLeftGrid, Row, 3) = Rst(2)
    Rst.Close
    .SetColObject(eLeftGrid, 3, False) = eTextBoxColumn
  End If
End With

ErrSub:
End Sub

Private Sub bGrid_KeyPressEdit(xGrid As AxioGrid.eSideGrid, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
With bGrid
  If .TextMatrix(eRightGrid, Row, 0) <> "" And KeyAscii = vbKeyReturn Then
    If .TextMatrix(eLeftGrid, Row, 1) <> vbNullString And .TextMatrix(eLeftGrid, Row, 2) <> vbNullString Then
      If .TextMatrix(eLeftGrid, Row, 1) <> "" And .TextMatrix(eLeftGrid, Row, 2) <> "" Then
        .TextMatrix(eLeftGrid, Row, 0) = Row
        .Rows = .Rows + 1
      End If
    End If
  End If
End With
End Sub

Private Sub cmdAbrir_Click()
'RECUPERO OC
  SQLString = "SELECT FECHAOC, PROVEEDOR, GUIA, FECHAGUIA, FECHAINGRESO FROM ORDEN WHERE OC=" & Trim$(txtOC) & ";"
  Rst.Open SQLString, Cnn, adOpenDynamic, adLockOptimistic
    mskFecha = Rst(0)
    txtID = Rst(1)
    Call SetProveedor
    txtGuia = "" & Rst(2)
    mskFechaGuia = "" & Rst(3)
    mskFechaIn = "" & Rst(4)
  Rst.Close

'RECUPERO DETALLE
  SQLString = "SELECT D1.PRODUCTO, B1.PRODUCTO, D1.CANTIDAD"
    SQLString = "SELECT BA.PID, DE.PRODUCTO, BA.PRODUCTO FROM ORDEN AS O1 INNER JOIN (BASE AS BA INNER JOIN DETALLE AS DE ON BA.CODIGO=DE.PRODUCTO) ON O1.OC=DE.OC WHERE DE.OC=" & txtOC & ";"
    bGrid.LoadtoLeft SQLString, Cnn, True, True
    
  If Trim(txtGuia) = "" Then
    SQLString = "SELECT DE.CANTIDAD FROM ORDEN AS O1 INNER JOIN DETALLE AS DE ON O1.OC = DE.OC WHERE DE.OC=" & txtOC & ";"
  Else
    SQLString = "SELECT DE.CANTIDAD, DE.RECIBIDO FROM ORDEN AS O1 INNER JOIN DETALLE AS DE ON O1.OC = DE.OC WHERE DE.OC=" & txtOC & ";"
  End If
  
  bGrid.LoadtoRight SQLString, Cnn, True

End Sub

Private Sub cmdCerrar_Click()
End
End Sub

Private Sub cmdGrabar_Click()
On Error GoTo ErrSave
 'GRABAR OC
  Cnn.BeginTrans
    SQLString = "INSERT INTO ORDEN (OC, FECHAOC, PROVEEDOR) VALUES (" & txtOC & ", '" & mskFecha & "', " & txtID & ");"
    Cnn.Execute SQLString, adCmdText
 'GRABAR DETALLE
 With bGrid
    For I = 1 To .Rows - 1
      SQLString = "INSERT INTO DETALLE (OC, PRODUCTO, CANTIDAD) VALUES (" & txtOC & ", " & .TextMatrix(eLeftGrid, I, 1) & ", " & .TextMatrix(eRightGrid, I, 0) & ");"
      Cnn.Execute SQLString, adCmdText
    Next
 End With
  Cnn.CommitTrans
 'SI GUARDAR=OK
  MsgBox "Orden de Compra N°" & txtOC & vbCrLf & "Proveedor :" & txtNombre & vbCrLf & "Guardada Correctamente!", vbInformation + vbOKOnly, "Guardar Orden de Compra"
  
Exit Sub
'SI GUARDAR=ERROR
ErrSave:
Cnn.RollbackTrans
MsgBox "Error Ingresando OC Nº" & txtOC & vbCrLf & Err.Number & " : " & Err.Description, vbOKOnly, "ERROR"
End Sub

Private Sub cmdGrabarIn_Click()
On Error GoTo ErrUpdate
 'GRABAR GUIA
  Cnn.BeginTrans
    SQLString = "UPDATE ORDEN SET GUIA=" & txtGuia & ", FECHAGUIA='" & mskFechaGuia & "', FECHAINGRESO='" & mskFechaIn & "';"
    Cnn.Execute SQLString, adCmdText
 'GRABAR DETALLE
 With bGrid
    For I = 1 To .Rows - 1
      SQLString = "UPDATE DETALLE SET RECIBIDO=" & .TextMatrix(eRightGrid, I, 1) & ");"
      Cnn.Execute SQLString, adCmdText
    Next
 End With
  Cnn.CommitTrans
 'SI GUARDAR=OK
  MsgBox "Guia N°" & txtGuia & vbCrLf & "Proveedor :" & txtNombre & vbCrLf & "Grabada Correctamente!", vbInformation + vbOKOnly, "Guardar Guia Proveedor"
Exit Sub
'SI GUARDAR=ERROR
ErrUpdate:
Cnn.RollbackTrans
MsgBox "Error Grabando Ingreso Guia Nº" & txtGuia & vbCrLf & Err.Number & " : " & Err.Description, vbOKOnly, "ERROR"
End Sub

Private Sub cmdInformes_Click()
frmInformes.Show 1
End Sub

Private Sub cmdNuevo_Click()
txtOC.Text = ""
txtRUT.Text = ""
txtNombre.Text = ""
txtGuia.Text = ""
mskFecha.Mask = ""
mskFecha.Text = ""
mskFecha.Mask = "##-##-####"
mskFechaGuia.Mask = ""
mskFechaGuia.Text = ""
mskFechaGuia.Mask = "##-##-####"
mskFechaIn.Mask = ""
mskFechaIn.Text = ""
mskFechaIn.Mask = "##-##-####"

With bGrid
  .ClearGrid (BothGrids)
  .Rows = 2
  .ColsLeft = 4
  .ColsRight = 1
  .AllowUserResizing(eLeftGrid) = flexResizeColumns
  '.SplitterFixed = True
  .SplitterPos = 6400
  .FormatString(eLeftGrid) = "|PID|CODIGO|PRODUCTO"
  .FormatString(eRightGrid) = "CANTIDAD"
  .ColWidth(LeftGrid, 1) = 450
  .ColWidth(LeftGrid, 2) = 1600
  .ColWidth(LeftGrid, 3) = 4000
  .SetColObject(eLeftGrid, 2, False) = eTextBoxColumn
  .SetColObject(eLeftGrid, 3, False) = eTextBoxColumn
  For I = 1 To .Rows - 1
    .TextMatrix(eLeftGrid, I, 0) = I
  Next
End With

End Sub

Private Sub Form_Load()
Set WshShell = CreateObject("WScript.Shell")

Connect True

mskFechaIn = Format(Now, "dd-mm-yyyy")

With bGrid
  .AllowUserResizing(eLeftGrid) = flexResizeColumns
  .Editable = True
  .Rows = 2
  .ColsLeft = 4
  .ColsRight = 1
  '.SplitterFixed = True
  .SplitterPos = 6400
  .FormatString(eLeftGrid) = "|PID|CODIGO|PRODUCTO"
  .FormatString(eRightGrid) = "CANTIDAD"
  .ColWidth(LeftGrid, 1) = 450
  .ColWidth(LeftGrid, 2) = 1600
  .ColWidth(LeftGrid, 3) = 4000
  .SetColObject(eLeftGrid, 2, False) = eTextBoxColumn
  .SetColObject(eLeftGrid, 3, False) = eTextBoxColumn
  For I = 1 To .Rows - 1
    .TextMatrix(eLeftGrid, I, 0) = I
  Next
  .ADOConnection = sqlConn
End With

End Sub

Private Sub txtGuia_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
  bGrid.ColsRight = 2
  bGrid.FormatString(eRightGrid) = "CANTIDAD|RECIBIDO"
  mskFechaGuia.SetFocus
End If
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
  Call SetProveedor
  txtRUT.SetFocus
  KeyAscii = 0
  'WshShell.SendKeys "{Tab}"
End If
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
  Call SetProveedor
  KeyAscii = 0
  WshShell.SendKeys "{Tab}"
End If
End Sub

Private Sub txtRUT_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
  If EsRUT(txtRUT) Then
    txtRUT = FormatoRUT(txtRUT)
    Call SetProveedor
    txtNombre.SetFocus
    KeyAscii = 0
    'WshShell.SendKeys "{Tab}"
  Else
    MsgBox "RUT inválido!", vbExclamation + vbOKOnly, "ERROR RUT"
    txtRUT.SelStart = 0
    txtRUT.SelLength = Len(txtRUT)
    txtRUT.SetFocus
  End If
End If
End Sub
