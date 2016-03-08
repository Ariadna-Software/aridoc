VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HcoRevisiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"





'FALTA###

'Faltan muchas cosas del hco revisiones.

'   1.-  Si lleva revisiones en la modificacion de archivo
'   2.-  Verificar todoas las opciones de cambios. Es decir cuando muevo o copio tb tendria
'        que modificarse las opciones
'   3.-  Cuando creo uno nuevo que meta el valor tb
'   5.-   Si se elimina un documento, que hacemos con las revisiones







'variables locales para almacenar los valores de las propiedades
Private mvarLlevaHcoRevision As Boolean 'copia local
Private mvarGuardoLasLecturas As Boolean



Dim SQL As String

Public Property Let LlevaHcoRevision(ByVal vData As Boolean)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.LlevaHcoRevision = 5
    mvarLlevaHcoRevision = vData
End Property


Public Property Get LlevaHcoRevision() As Boolean
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.LlevaHcoRevision
    LlevaHcoRevision = mvarLlevaHcoRevision
End Property


Public Property Let GuardoLasLecturas(ByVal vData As Boolean)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.LlevaHcoRevision = 5
    mvarGuardoLasLecturas = vData
End Property


Public Property Get GuardoLasLecturas() As Boolean
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.LlevaHcoRevision
    GuardoLasLecturas = mvarGuardoLasLecturas
End Property



Private Sub Class_Initialize()
    'Al inicializar el objeto, no está cargado y por eso mvarCargado se inicializa a False
    mvarLlevaHcoRevision = TieneRevisionDocumental()
End Sub






'--------------------------------------------------------------------------------
Private Function TieneRevisionDocumental() As Boolean
On Error GoTo ET
    TieneRevisionDocumental = False
    Conn.Execute "Select * from revision  where id=-1"
    TieneRevisionDocumental = True
    Exit Function
ET:
    Err.Clear
End Function



'---------------------------------------------------------------------------------
' 0.- Primera INSERCCION
' 1.- Modificar valores de los text
' 2.- Eliminar
' 3.- Modificar documento
' 4.- Lectura documento
' 5.- Cambiar propietario / grupo propietario
' 6.- Mover en carpetas
' 7.- Realizar copia del documento
Public Function InsertaRevision(CodImg As Long, accion As Byte, ByRef ElUsuario As Cusuarios, ByRef Cambios As String) As Boolean


    On Error GoTo EInsertaRevision
    
    
    'Si no estaamos en modotrabajao normal no hace nada
    If ModoTrabajo <> vbNorm Then Exit Function
    
    If Not mvarLlevaHcoRevision Then
        Exit Function
    End If
    
    'Si no guardo las lecturas
    If accion = 4 Then
        If Not mvarGuardoLasLecturas Then
            InsertaRevision = True
            Exit Function
        End If
    End If
    SQL = "INSERT INTO revision (id, fecha, usuario, pc, accion, revision, cambios) VALUES ("
    SQL = SQL & CodImg & ",now()," & ElUsuario.codusu & "," & ElUsuario.PC & ","
    SQL = SQL & accion & ","
    Select Case accion
    Case 0
        'Nuevo documento
        SQL = SQL & "NULL,NULL"
    Case 1
        'Modificar claves
        SQL = SQL & "NULL,'" & Cambios & "'"
        
    Case 2
        'Eliminar
        SQL = SQL & "NULL,NULL"
    Case 3
        'Modificar fichero
        SQL = SQL & "NULL,NULL"
        
    Case 4
        'Lectura
        SQL = SQL & "NULL,NULL"
    Case 5
        'Cambio propietario
        SQL = SQL & "NULL,'" & Cambios & "'"
    Case 6
        'Mover documentos
        SQL = SQL & "NULL,'Carpeta anterior: " & DevNombreSql(Cambios) & "'"
    
    Case 7
        'Copiado
        SQL = SQL & "NULL,'Carpeta anterior: " & DevNombreSql(Cambios) & "'"
        
    End Select
    SQL = SQL & ")"
    Conn.Execute SQL
    Exit Function
EInsertaRevision:
    Err.Clear
End Function


Public Function DevuelveTextoRev(accion As Integer) As String
    Select Case accion
    Case 0
        'Nuevo documento
        SQL = "NUEVO"
    Case 1
        'Modificar claves
        SQL = "CAMBIO"
        
    Case 2
        'Eliminar
        SQL = "ELIM"
    Case 3
        'Modificar fichero
        SQL = "MODIF"
        
    Case 4
        'Lectura
        SQL = "LEER"
    Case 5
        'Cambio propietario
        SQL = "PROP"
    Case 6
        'Mover documentos
        SQL = "MOVER"
    
    Case 7
        'Copiado
        SQL = "COPIA"
        
    End Select
    DevuelveTextoRev = SQL
End Function



'//  Elimimanos la referencia de hco
Public Function EliminarReferencia(CodImg As Long)

    On Error Resume Next
        SQL = "DELETE FROM  revision where id = " & CodImg
        Conn.Execute SQL
        If Err.Number <> 0 Then Err.Clear

End Function

