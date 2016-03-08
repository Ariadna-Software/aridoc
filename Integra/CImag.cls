VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CImag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'variables locales para almacenar los valores de las propiedades
Private mvarId As Long 'copia local
Private mvarClave1 As String 'copia local
Private mvarClave2 As String 'copia local
Private mvarClave3 As String 'copia local
Private mvarDes As String 'copia local   'TENDREMOS EL NOMBRE DEL FICHERO
Private mvarNomFich As String 'copia local
Private mvarFechaDig As String 'copia local
Private mvarFechaDoc As String 'copia local
Private mvarNomPath As String 'copia local
Private mvarExtension As Long


'Dim rs As Recordset
'Dim SQL As String
'Dim Mens As String

Public Property Let FechaDoc(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.FechaDoc = 5
    mvarFechaDoc = vData
End Property


Public Property Get FechaDoc() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.FechaDoc
    FechaDoc = mvarFechaDoc
End Property



Public Property Let FechaDig(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.FechaDig = 5
    mvarFechaDig = vData
End Property


Public Property Get FechaDig() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.FechaDig
    FechaDig = mvarFechaDig
End Property








'Public Function Agregar() As Byte
'Dim Path As String
'
'    On Error GoTo Err1
'    sql = "SELECT * FROM TImagen WHERE Id = " & mvarId
'    Set rs = Db.OpenRecordset(sql, dbOpenDynaset)
'    If rs.RecordCount > 0 Then
'        Mens = "Ya existe otro archivo en el sistema con el mismo código."
'        MsgBox Mens, vbExclamation
'        Agregar = 1
'    Else
'        rs.AddNew
'
'        rs!Clave1 = mvarClave1
'        rs!Id = mvarId
'        rs!Clave2 = mvarClave2
'        rs!Clave3 = mvarClave3
'        rs!Des = mvarDes
'        rs!FechaDig = PonFecha(mvarFechaDig)
'        rs!FechaDoc = PonFecha(mvarFechaDoc)
'        rs!NomFich = mvarNomFich
'        rs!NomPath = mvarNomPath
'        rs!Extension = mvarExtension
'        rs.Update
'        Agregar = 0
'    End If
'    rs.Close
'    Exit Function
'Err1:
'    Mens = "Error: " & Err.Number & " : " & Err.Description
'    MsgBox Mens, vbExclamation
'    rs.Close
'    Agregar = 1
'End Function
'
'
'Public Function Eliminar() As Byte
'    On Error GoTo Err1
'    sql = "SELECT * FROM TImagen WHERE Id = " & mvarId
'    Set rs = Db.OpenRecordset(sql, dbOpenDynaset)
'    If rs.RecordCount = 0 Then
'        Eliminar = 1
'    Else
'        rs.Delete
'        rs.MoveNext
'        If rs.EOF Then rs.MoveFirst
'        Eliminar = 0
'    End If
'    rs.Close
'    Exit Function
'Err1:
'    Mens = "Error: " & Err.Number & " : " & Err.Description
'    MsgBox Mens, vbExclamation
'    rs.Close
'    Eliminar = 1
'End Function
'
'Public Function Leer(ByVal vId As Long) As Byte
'    On Error GoTo Err1
'    sql = "SELECT * FROM TImagen WHERE Id = " & vId
'    Set rs = Db.OpenRecordset(sql, dbOpenDynaset)
'    If rs.RecordCount = 0 Then
'        Leer = 1
'    Else
'         mvarClave1 = DBLet(rs!Clave1)
'         mvarClave2 = DBLet(rs!Clave2)
'         mvarClave3 = DBLet(rs!Clave3)
'         mvarDes = DBLet(rs!Des)
'         mvarFechaDig = DBLetDate(rs!FechaDig)
'         mvarFechaDoc = DBLetDate(rs!FechaDoc)
'         mvarNomFich = DBLet(rs!NomFich)
'         mvarId = DBLet(rs!Id)
'         mvarNomPath = DBLet(rs!NomPath)
'         mvarExtension = rs!Extension
'         Leer = 0
'    End If
'    rs.Close
'    Exit Function
'Err1:
'    Mens = "Error: " & Err.Number & " : " & Err.Description
'    MsgBox Mens, vbExclamation
'    rs.Close
'    Leer = 1
'End Function
'
'
'
'Public Function modificar() As Byte
'    On Error GoTo Err1
'    sql = "SELECT * FROM TImagen WHERE Id = " & mvarId
'    Set rs = Db.OpenRecordset(sql, dbOpenDynaset)
'    If rs.RecordCount = 0 Then
'        modificar = 1
'    Else
'        rs.Edit
'        rs!Clave1 = mvarClave1
'        rs!Clave2 = mvarClave2
'        rs!Clave3 = mvarClave3
'        rs!Des = mvarDes
'        rs!FechaDig = PonFecha(mvarFechaDig)
'        rs!FechaDoc = PonFecha(mvarFechaDoc)
'        rs!NomPath = mvarNomPath
'        rs!NomFich = mvarNomFich
'        rs!Extension = mvarExtension
'        rs.Update
'        modificar = 0
'    End If
'    rs.Close
'    Exit Function
'Err1:
'    Mens = "Error: " & Err.Number & " : " & Err.Description
'    MsgBox Mens, vbExclamation
'    rs.Close
'    modificar = 1
'End Function
'
'Public Function Siguiente() As Long
'    '-- Devuelve el valor siguiente de código para esta tabla
'    Siguiente = 1
'    sql = "select MAX(Id) from TImagen"
'    Set rs = Db.OpenRecordset(sql)
'    If rs.RecordCount > 0 Then
'        If Not IsNull(rs.Fields(0)) Then Siguiente = rs.Fields(0) + 1
'    End If
'    mvarId = Siguiente
'    rs.Close
'End Function


Public Property Let NomFich(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.NomFich = 5
    mvarNomFich = vData
End Property


Public Property Get NomFich() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.NomFich
    NomFich = mvarNomFich
End Property


Public Property Let Des(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Des = 5
    mvarDes = vData
End Property


Public Property Get Des() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Des
    Des = mvarDes
End Property



Public Property Let Clave3(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Clave3 = 5
    mvarClave3 = vData
End Property


Public Property Get Clave3() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Clave3
     Clave3 = mvarClave3
End Property



Public Property Let Clave2(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Clave2 = 5
    mvarClave2 = vData
End Property


Public Property Get Clave2() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Clave2
    Clave2 = mvarClave2
End Property



Public Property Let Clave1(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Clave1 = 5
    mvarClave1 = vData
End Property


Public Property Get Clave1() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Clave1
    Clave1 = mvarClave1
End Property



Public Property Let Id(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Id = 5
    mvarId = vData
End Property


Public Property Get Id() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Id
    Id = mvarId
End Property


Public Property Let NomPath(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.NomPath = 5
    mvarNomPath = vData
End Property


Public Property Get NomPath() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.NomPath
    NomPath = mvarNomPath
End Property


Public Property Let Extension(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.NomPath = 5
    mvarExtension = vData
End Property


Public Property Get Extension() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.NomPath
    Extension = mvarExtension
End Property



Public Function Agregar(Cod2 As Long) As Boolean
Dim SQL As String

    On Error GoTo EI
    Agregar = False
    'Hemos quitado: campo4,fecha3,importe1y 2, observa
    SQL = "INSERT INTO tmpintegra (codusu, codigo,  campo1, campo2, campo3,  fecha1, fecha2, "
    SQL = SQL & "NombreArchivo, carpeta,correcto) VALUES ("
    SQL = SQL & Cod2 & "," & mvarId & ",'" & DevNombreSql(mvarClave1) & "','"
    SQL = SQL & DevNombreSql(mvarClave2) & "','" & DevNombreSql(mvarClave3) & "',"
    'Fechas
    SQL = SQL & "'" & Format(mvarFechaDoc, "yyyy-mm-dd") & "',"
    SQL = SQL & "'" & Format(mvarFechaDig, "yyyy-mm-dd") & "',"
    SQL = SQL & "'" & mvarNomFich & "','" & DevNombreSql(mvarNomPath) & "',0)"
    Conn.Execute SQL
    Agregar = True
    Exit Function
EI:
    MsgBox Err.Description, vbExclamation
End Function
