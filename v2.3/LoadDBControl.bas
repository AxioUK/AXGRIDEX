Attribute VB_Name = "LoadDBControl"
Option Explicit
' declaraciones api
''''''''''''''''''''''''''''''''''''''''''
Private Declare Function SendMessageByString& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                      ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String)

' función que deshabilita el repintado de una ventana en windows
Private Declare Function LockWindowUpdate& Lib "user32" (ByVal hwndLock As Long)

' variables y constantes
Private Const CB_ADDSTRING& = &H143
Private Const LB_ADDSTRING As Long = &H180



' Función que carga el campo en el combobox o listbox
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function LoadListBox(ObjCtrl As Object, Rst As ADODB.Recordset, Columna As String) As Boolean
  Dim ret                 As Long
  Dim Mensaje_SendMessage As Long
    
 ' On Error GoTo Error_Function:
  ' verifica que el recordset contenga un conjunto de registros
  If Rst.BOF And Rst.EOF Then
    MsgBox " No hay registros para agregar", vbInformation
    Call LockWindowUpdate(0&)
    ' sale
    Exit Function
  End If
    
  ' Chequea con TypeName el tipo de control enviado como parámetro
  If TypeName(ObjCtrl) = "axComboBox" Then
    Mensaje_SendMessage = LB_ADDSTRING ' mensaje para SendMessage CB_ADDSTRING&
  ElseIf TypeName(ObjCtrl) = "ListBox" Then
    Mensaje_SendMessage = LB_ADDSTRING ' mensaje para SendMessage
  End If
    
  ' deshabilita el repintado del control para que cargue los datos mas rapidamente
  Call LockWindowUpdate(ObjCtrl.hwnd)
  DoEvents
  ' Posiciona el recordset en el primer registro
  Rst.MoveFirst
  ' elimina todo el contenido del combo o listbox( opcional )
  ObjCtrl.Clear
  ' recorre las filas del recordset
  Do Until Rst.EOF
    ' chequea que el valor no sea un nulo
    If Not IsNull(Rst(Columna).Value) Then
      'Agrega el dato en el control con el mensaje CB_ADDSTRING o LB_ADDSTRING dependiendo del tipo de control
      ret = SendMessageByString(ObjCtrl.hwnd, Mensaje_SendMessage, 0, Rst(Columna).Value)
    End If
    ' siguiente registro
    Rst.MoveNext
  Loop
    
  ' selecciona el primer elemento del listado
  If ObjCtrl.ListCount > 0 Then
    ObjCtrl.ListIndex = 0
  End If
      
  ' vuelve a habilitar el repintado
  Call LockWindowUpdate(0&)
  ' retorno
  LoadListBox = True
  
  Exit Function
  ' rutina de error
Error_Function:
  MsgBox Err.Description, vbCritical
  ' En caso de error vuelve a activar el repintado
  Call LockWindowUpdate(0&)
  ObjCtrl.Refresh
End Function



