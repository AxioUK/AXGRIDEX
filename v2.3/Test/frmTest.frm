VERSION 5.00
Object = "{70B0A7F7-C129-4E09-AE3A-C2568FC3CF36}#1.0#0"; "prjAVFontCombo.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AVFontCombo"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   7005
   StartUpPosition =   3  'Windows Default
   Begin FontCombo.AVFontCombo AVFontCombo1 
      Height          =   420
      Left            =   3300
      TabIndex        =   1
      Top             =   60
      Width           =   3675
      _extentx        =   6482
      _extenty        =   741
      fontsize        =   12
      forecolor       =   -2147483642
   End
   Begin RichTextLib.RichTextBox rtfTest 
      Height          =   3195
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   5636
      _Version        =   393217
      HideSelection   =   0   'False
      TextRTF         =   $"frmTest.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AVFontCombo1_Change()

    rtfTest.SelFontName = AVFontCombo1.Text
    
End Sub

Private Sub AVFontCombo1_Click()

    rtfTest.SelFontName = AVFontCombo1.Text
    
End Sub

Private Sub Form_Load()

    rtfTest.Text = "Here is some sample text" & vbCrLf & _
    "to test the AVFontCombo Control" & vbCrLf & _
    "blah blah blah blah" & vbCrLf & _
    "blah blah blah blah" & vbCrLf & _
    "blah blah blah blah" & vbCrLf & _
    "blah blah blah blah"
    
End Sub
