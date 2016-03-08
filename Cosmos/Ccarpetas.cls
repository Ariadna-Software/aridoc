VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Ccarpetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Variables locales que contienen valores de propiedad
Private mvarcodcarpeta As Integer 'copia local
Private mvarNombre As String 'copia local
Private mvarpadre As Integer 'copia local
Private mvaruserprop As Long 'copia local
Private mvaralmacen As Integer 'copia local
Private mvargroupprop As Long 'copia local
Private mvarlecturau As Long 'copia local
Private mvarlecturag As Long 'copia local
Private mvarescriturau As Long 'copia local
Private mvarescriturag As Long 'copia local


'Estas propiedades son privadas de la carpeta
Private mvarversion As Integer 'copia local
Private mvarpathreal As String 'copia local
Private mvarSRV As String 'copia local
Private mvaruser As String 'copia local
Private mvarpwd As String 'copia local



Private SQL As String  'Para la sentencia
Private Mens As String    'Para los mensajes
Private RS As ADODB.Recordset 'Para la base de datos
Private Rc As Byte



'------------------------------------------------
'Propiedades del modulo Ccarpetas
'------------------------------------------------


Public Property Let codcarpeta(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.codcarpeta = 5
     mvarcodcarpeta = vData
End Property

Public Property Get codcarpeta() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.codcarpeta
     codcarpeta = mvarcodcarpeta
End Property


Public Property Let Nombre(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.nombre = 5
     mvarNombre = vData
End Property

Public Property Get Nombre() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.nombre
     Nombre = mvarNombre
End Property


Public Property Let padre(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.padre = 5
     mvarpadre = vData
End Property

Public Property Get padre() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.padre
     padre = mvarpadre
End Property


Public Property Let userprop(ByVal vData As Long)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.userprop = 5
     mvaruserprop = vData
End Property

Public Property Get userprop() As Long
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.userprop
     userprop = mvaruserprop
End Property


Public Property Let Almacen(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.almacen = 5
     mvaralmacen = vData
End Property

Public Property Get Almacen() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.almacen
     Almacen = mvaralmacen
End Property


Public Property Let groupprop(ByVal vData As Long)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.groupprop = 5
     mvargroupprop = vData
End Property

Public Property Get groupprop() As Long
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.groupprop
     groupprop = mvargroupprop
End Property


Public Property Let lecturau(ByVal vData As Long)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.lecturau = 5
     mvarlecturau = vData
End Property

Public Property Get lecturau() As Long
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.lecturau
     lecturau = mvarlecturau
End Property


Public Property Let lecturag(ByVal vData As Long)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.lecturag = 5
     mvarlecturag = vData
End Property

Public Property Get lecturag() As Long
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.lecturag
     lecturag = mvarlecturag
End Property


Public Property Let escriturau(ByVal vData As Long)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.escriturau = 5
     mvarescriturau = vData
End Property

Public Property Get escriturau() As Long
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.escriturau
     escriturau = mvarescriturau
End Property


Public Property Let escriturag(ByVal vData As Long)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.escriturag = 5
     mvarescriturag = vData
End Property

Public Property Get escriturag() As Long
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.escriturag
     escriturag = mvarescriturag
End Property


Public Property Let version(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.escriturag = 5
     mvarversion = vData
End Property

Public Property Get version() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.escriturag
     version = mvarversion
End Property



Public Property Let pwd(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.escriturag = 5
     mvarpwd = vData
End Property

Public Property Get pwd() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.escriturag
     pwd = mvarpwd
End Property



Public Property Let user(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.escriturag = 5
     mvaruser = vData
End Property

Public Property Get user() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.escriturag
     user = mvaruser
End Property


Public Property Let SRV(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.escriturag = 5
     mvarSRV = vData
End Property

Public Property Get SRV() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.escriturag
     SRV = mvarSRV
End Property


Public Property Let pathreal(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.escriturag = 5
     mvarpathreal = vData
End Property

Public Property Get pathreal() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.escriturag
     pathreal = mvarpathreal
End Property

'------------------------------------------------
'FUNCIONES del modulo Ccarpetas
'------------------------------------------------



'****   AGREGAR  ***
Public Function Agregar() As Byte
    On Error GoTo Err1
    Agregar = 1

    'Compruebo k no existe otra carpeta al1 mismo nivel con ese nombre
'    SQL = "SELECT nombre from carpetas where padre= " & mvarpadre
'    SQL = SQL & " AND nombre ='" & Nombre & "'"
'    Set Rs = New ADODB.Recordset
'    Rs.Open SQL, ConnNuevoAridoc, adOpenForwardOnly, adLockPessimistic, adCmdText
'    If Not Rs.EOF Then
'        MsgBox "Ya existe una carpeta con ese nombre", vbExclamation
'    Else
'        SQL = ""
'    End If
'    Rs.Close
'
'
'    If SQL <> "" Then Exit Function
    
    'Siguiente
    'Siguiente
    
    '-----------------------------
    Set RS = Nothing
    
    'SQL
    SQL = "INSERT INTO carpetas (codcarpeta, nombre, padre, userprop,"
    SQL = SQL & "almacen, groupprop, lecturau, lecturag, escriturau,"
    SQL = SQL & "escriturag) VALUES ("
    
    'Asignamos valores
    SQL = SQL & mvarcodcarpeta & ",'" & mvarNombre & "',"
    SQL = SQL & mvarpadre & "," & mvaruserprop & "," & mvaralmacen & ","
    SQL = SQL & mvargroupprop & "," & mvarlecturau & "," & mvarlecturag & ","
    SQL = SQL & mvarescriturau & "," & mvarescriturag & ")"
    ConnNuevoAridoc.Execute SQL
    Agregar = 0
    
    Exit Function

Err1:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    Agregar = 1
End Function


'****   ELIMINAR  ***
Public Function Eliminar() As Byte
    On Error GoTo Err2
    Eliminar = 1
    If mvarcodcarpeta = 1 Then
        MsgBox "Carpeta raiz NO se puede eliminar", vbExclamation
        Exit Function
    End If
    ConnNuevoAridoc.Execute "Delete from Carpetas WHERE codcarpeta =" & mvarcodcarpeta

Err2:
    If Err.Number <> 0 Then
        MsgBox Err.Description
    Else
        Eliminar = 0
    End If
End Function


'****   LEER  ***
Public Function Leer(vIdcodcarpeta As Integer, EnHCO As Boolean) As Byte
    On Error GoTo Err3

    SQL = "SELECT carpetas.* ,almacen.* FROM "
    If EnHCO Then
        SQL = SQL & "carpetashco as carpetas,almacenhco as almacen"
    Else
        SQL = SQL & "carpetas,almacen"
    End If
    SQL = SQL & " WHERE carpetas.almacen = almacen.codalma AND "
    SQL = SQL & " codcarpeta = " & vIdcodcarpeta
    Set RS = New ADODB.Recordset

    RS.Open SQL, ConnNuevoAridoc, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RS.EOF Then
        Leer = 1
        'ELSE
        Else
            mvarcodcarpeta = RS!codcarpeta
            mvarNombre = RS!Nombre
            mvarpadre = RS!padre
            mvaruserprop = RS!userprop
            mvaralmacen = RS!Almacen
            mvargroupprop = RS!groupprop
            mvarlecturau = RS!lecturau
            mvarlecturag = RS!lecturag
            mvarescriturau = RS!escriturau
            mvarescriturag = RS!escriturag
            
            
            
            'Establezco el ALMACEN
            mvarversion = RS!version
            mvarpathreal = RS!pathreal
            mvarSRV = RS!SRV
            mvaruser = DBLet(RS!user, "T")
            mvarpwd = DBLet(RS!pwd, "T")
            
            
            Leer = 0
    End If

    RS.Close
    Set RS = Nothing
    Exit Function

Err3:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    Leer = 1
    Set RS = Nothing
End Function


'****   MODIFICAR  ***
Public Function Modificar() As Byte
    On Error GoTo Err4

    'MODIFICAR
    SQL = "UPDATE Carpetas set "
    
    
            
        SQL = SQL & "Nombre ='" & mvarNombre & "'"
        SQL = SQL & ",padre = " & mvarpadre
        SQL = SQL & ",userprop= " & mvaruserprop
        SQL = SQL & ",almacen = " & mvaralmacen
        SQL = SQL & ",groupprop = " & mvargroupprop
        SQL = SQL & ",lecturau = " & mvarlecturau
        SQL = SQL & ",lecturag = " & mvarlecturag
        SQL = SQL & ",escriturau = " & mvarescriturau
        SQL = SQL & ",escriturag = " & mvarescriturag
        
        SQL = SQL & " WHERE codCarpeta =" & mvarcodcarpeta
        
        ConnNuevoAridoc.Execute SQL
    Exit Function

Err4:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    Modificar = 1
End Function


'****   SIGUIENTE  ***
Public Function Siguiente() As Integer

    Siguiente = 1
    SQL = "SELECT MAX(codcarpeta) FROM carpetas"
    RS.Open SQL, ConnNuevoAridoc, , , adCmdText
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then Siguiente = RS.Fields(0) + 1
    End If
    mvarcodcarpeta = Siguiente
    RS.Close
End Function





Private Sub Class_Terminate()
    Set RS = Nothing
End Sub


