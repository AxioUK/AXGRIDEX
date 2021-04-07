VERSION 5.00
Object = "{79308990-DEA2-11D6-AEDC-DF8547B6407B}#6.1#0"; "KlexGrid.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9585
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   9585
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSum 
      Caption         =   "Sum of Salary"
      Height          =   675
      Left            =   6660
      TabIndex        =   8
      Top             =   3180
      Width           =   1755
   End
   Begin VB.CommandButton cmdAutosize 
      Caption         =   "Autosize Column 1 && 2"
      Height          =   705
      Left            =   6630
      TabIndex        =   7
      Top             =   1950
      Width           =   1725
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Change 3nd Row CellAlignment"
      Height          =   615
      Left            =   6630
      TabIndex        =   6
      Top             =   1050
      Width           =   1665
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change 2nd Row BackColor"
      Height          =   615
      Left            =   6630
      TabIndex        =   5
      Top             =   300
      Width           =   1665
   End
   Begin VB.CheckBox chkMove 
      Caption         =   "On Enter Key Move Down"
      Height          =   405
      Left            =   4620
      TabIndex        =   4
      Top             =   5040
      Width           =   2925
   End
   Begin VB.CheckBox chkNameCol 
      Caption         =   "Allow to Edit Name Column"
      Height          =   375
      Left            =   4620
      TabIndex        =   3
      Top             =   4470
      Width           =   2895
   End
   Begin VB.CheckBox chkEditable 
      Caption         =   "Editable"
      Height          =   345
      Left            =   390
      TabIndex        =   2
      Top             =   4470
      Width           =   1635
   End
   Begin VB.CheckBox chkAltColor 
      Caption         =   "Alternate Color"
      Height          =   345
      Left            =   390
      TabIndex        =   1
      Top             =   4950
      Width           =   1575
   End
   Begin Grid.KlexGrid KlexGrid1 
      Height          =   3315
      Left            =   390
      TabIndex        =   0
      Top             =   330
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   5847
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
      MouseIcon       =   "Form1.frx":0000
      Rows            =   10
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkAltColor_Click()
   If chkAltColor.Value = vbChecked Then
      KlexGrid1.BackColorAlternate = vbCyan
   Else
      KlexGrid1.BackColorAlternate = vbWhite
   End If
End Sub

Private Sub chkEditable_Click()
   If chkEditable.Value = vbChecked Then
      KlexGrid1.Editable = True
   Else
      KlexGrid1.Editable = False
   End If
End Sub


Private Sub chkMove_Click()
    If chkMove.Value = vbChecked Then
       KlexGrid1.EnterKeyBehaviour = klexEKMoveDown
    Else
       KlexGrid1.EnterKeyBehaviour = klexEKMoveRight
    End If
End Sub

Private Sub cmdSum_Click()
    KlexGrid1.TextMatrix(9, 2) = KlexGrid1.Aggregate(klexSTSum, 1, 2, 8, 2)
End Sub

Private Sub Command1_Click()
     KlexGrid1.Cell(klexcpCellBackColor, 2, 1, 2, KlexGrid1.Cols - 1) = vbBlue
End Sub

Private Sub Command2_Click()
     KlexGrid1.Cell(klexcpCellAlignment, 3, 1, 3, KlexGrid1.Cols - 1) = 3
End Sub

Private Sub cmdAutosize_Click()
    KlexGrid1.AutoSizeMode = klexAutoSizeColWidth
    KlexGrid1.AutoSize 1, 2
End Sub

Private Sub Form_Load()
    KlexGrid1.TextMatrix(0, 1) = "Name"
    KlexGrid1.TextMatrix(0, 2) = "Salary"
    KlexGrid1.ColDisplayFormat(2) = "#0.00"
    
    'Implemented only for Numeric Entry
    KlexGrid1.ColInputMask(2) = "000.00"
    KlexGrid1.TextMatrix(9, 0) = "Total"
End Sub

Private Sub KlexGrid1_BeforeEdit(Cancel As Boolean)
     If KlexGrid1.Col = 1 Then
        If chkNameCol.Value = vbUnchecked Then
            Cancel = True
        End If
     End If
End Sub
