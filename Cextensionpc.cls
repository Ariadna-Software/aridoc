VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cextensionpc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Variables locales que contienen valores de propiedad
Private mvarcodext As Integer 'copia local
Private mvarcodequipo As Integer 'copia local
Private mvarpathexe As String 'copia local
Private mvarimpresion As String 'copia local
Private mvarextension As String
Private mvarArchivosModificables As Boolean


Private SQL As String  'Para la sentencia
Private Mens As String    'Para los mensajes
Private Rs As ADODB.Recordset 'Para la base de datos
Private Rc As Byte



'------------------------------------------------
'Propiedades del modulo Cextensionpc
'------------------------------------------------


Public Property Let codext(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.codext = 5
     mvarcodext = vData
End Property

Public Property Get codext() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.codext
     codext = mvarcodext
End Property


Public Property Let codequipo(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.codequipo = 5
     mvarcodequipo = vData
End Property

Public Property Get codequipo() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.codequipo
     codequipo = mvarcodequipo
End Property


Public Property Let pathexe(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.pathexe = 5
     mvarpathexe = vData
End Property

Public Property Get pathexe() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.pathexe
     pathexe = mvarpathexe
End Property


Public Property Let impresion(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.impresion = 5
     mvarimpresion = vData
End Property

Public Property Get impresion() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.impresion
     impresion = mvarimpresion
End Property


Public Property Get Extension() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.pathexe
     Extension = mvarextension
End Property


Public Property Get ArchivosModificables() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.pathexe
     ArchivosModificables = mvarArchivosModificables
End Property


'------------------------------------------------
'FUNCIONES del modulo Cextensionpc
'------------------------------------------------



'****   AGREGAR  ***
'Public Function Agregar() As Byte
'    On Error GoTo Err1
'
'
'            RS.AddNew
'            RS!codext = mvarcodext
'            RS!codequipo = mvarcodequipo
'            RS!pathexe = mvarpathexe
'            RS!impresion = mvarimpresion
'            RS.Update
'            Agregar = 0
'
'
'    Exit Function
'
'Err1:
'    Mens = "Se ha producido un error." & vbCrLf
'    Mens = Mens & "Número: " & Err.Number & vbCrLf
'    Mens = Mens & "Descripción: " & Err.Description
'    MsgBox Mens, vbExclamation
'    Agregar = 1
'End Function


''****   ELIMINAR  ***
'Public Function Eliminar() As Byte
'    On Error GoTo Err2
'
'
'
'    'eliminar
'
'
'
'    Exit Function
'
'Err2:
'    Mens = "Se ha producido un error." & vbCrLf
'    Mens = Mens & "Número: " & Err.Number & vbCrLf
'    Mens = Mens & "Descripción: " & Err.Description
'    MsgBox Mens, vbExclamation
'    Eliminar = 1
'End Function


'****   LEER  ***
Public Function Leer(vIdcodext As Integer, vIdcodequipo As Integer) As Byte
    On Error GoTo Err3
    Leer = 1
    SQL = "SELECT extensionpc.*,extension.exten,extension.modificable FROM extensionpc,extension"
    SQL = SQL & " WHERE "
    SQL = SQL & " extensionpc.codext = extension.codext AND"
    SQL = SQL & " extensionpc.codext = " & vIdcodext
    SQL = SQL & "  and codequipo = " & vIdcodequipo
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
            mvarcodext = Rs!codext
            mvarcodequipo = Rs!codequipo
            mvarpathexe = Rs!pathexe
            mvarimpresion = Rs!impresion
            mvarextension = Rs!Exten
            mvarArchivosModificables = (Rs!modificable = 1)
            Leer = 0
    Else
        MsgBox "Extension NO encontrada", vbExclamation
    End If

    Rs.Close
    

Err3:
    If Err.Number <> 0 Then
        Mens = "Se ha producido un error." & vbCrLf
        Mens = Mens & "Número: " & Err.Number & vbCrLf
        Mens = Mens & "Descripción: " & Err.Description
        MsgBox Mens, vbExclamation
    End If
    Set Rs = Nothing
End Function


''****   MODIFICAR  ***
'Public Function Modificar() As Byte

'
'Err4:
'    Mens = "Se ha producido un error." & vbCrLf
'    Mens = Mens & "Número: " & Err.Number & vbCrLf
'    Mens = Mens & "Descripción: " & Err.Description
'    MsgBox Mens, vbExclamation
'    Modificar = 1
'End Function






Private Sub Class_Terminate()
    Set Rs = Nothing
End Sub
