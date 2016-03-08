VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cusuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Variables locales que contienen valores de propiedad
Private mvarcodusu As Integer 'copia local
Private mvarNombre As String 'copia local
Private mvarlogin As String 'copia local
Private mvarpassword As String 'copia local


'-----------------------------
'E-MAIL

Private mvare_dir As String
Private mvare_server As String
Private mvare_login As String
Private mvare_pwd As String

'Private mvarPreferenciasC As String 'Preferencias correo





'Datos k asignara automaticamente.
Private mvarPC As Integer
Private mvarVelocidad As Single
Private mvarNomPC As String
Private mvarCargaIconosExtensiones As Boolean
Private mvarNivel As Byte   ' de 0(administrador) hasta 6(consulta)

'GRUPOS TODOS
Private mvarGrupo As Long
Private mvarGrupoPpal As Integer


Private mvarPreferencias As CPreferenciaPersonal



Private SQL As String  'Para la sentencia
Private Mens As String    'Para los mensajes
Private Rs As ADODB.Recordset 'Para la base de datos
Private Rc As Byte



'------------------------------------------------
'Propiedades del modulo Cusuarios
'------------------------------------------------
Public Property Let Nivel(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.codusu = 5
     mvarNivel = vData
End Property

Public Property Get Nivel() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.codusu
     
     Nivel = mvarNivel
End Property




Public Property Let codusu(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.codusu = 5
     mvarcodusu = vData
End Property

Public Property Get codusu() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.codusu
     codusu = mvarcodusu
End Property


Public Property Let Nombre(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.Nombre = 5
     mvarNombre = vData
End Property

Public Property Get Nombre() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.Nombre
     Nombre = mvarNombre
End Property


Public Property Let login(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.login = 5
     mvarlogin = vData
End Property

Public Property Get login() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.login
     login = mvarlogin
End Property


Public Property Let Password(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.password = 5
     mvarpassword = vData
End Property

Public Property Get Password() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.password
     Password = mvarpassword
End Property

Public Property Let e_dir(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.password = 5
     mvare_dir = vData
End Property

Public Property Get e_dir() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.password
     e_dir = mvare_dir
End Property



Public Property Let e_server(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.password = 5
     mvare_server = vData
End Property

Public Property Get e_server() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.password
     e_server = mvare_server
End Property

Public Property Let e_login(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.password = 5
     mvare_login = vData
End Property

Public Property Get e_login() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.password
     e_login = mvare_login
End Property

Public Property Let e_pwd(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.password = 5
     mvare_pwd = vData
End Property

Public Property Get e_pwd() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.password
     e_pwd = mvare_pwd
End Property



Public Property Let preferencias(ByVal vData As CPreferenciaPersonal)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.preferencias = 5
     Set mvarPreferencias = vData
End Property

Public Property Get preferencias() As CPreferenciaPersonal
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.preferencias
     Set preferencias = mvarPreferencias
End Property


Public Property Get PC() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.codusu
     PC = mvarPC
End Property


Public Property Let PC(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.codusu
     mvarPC = vData
End Property



Public Property Get NomPC() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.preferencias
     NomPC = mvarNomPC
End Property

Public Property Get Velocidad() As Single
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.codusu
     Velocidad = mvarVelocidad
End Property


Public Property Get CargaIconosExtensiones() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.codusu
     CargaIconosExtensiones = mvarCargaIconosExtensiones
End Property


Public Property Let CargaIconosExtensiones(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.codusu
     mvarCargaIconosExtensiones = vData
End Property



Public Property Get Grupo() As Long
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.codusu
     Grupo = mvarGrupo
End Property


Public Property Get GrupoPpal() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.codusu
     GrupoPpal = mvarGrupoPpal
End Property


'****   LEER  ***
'0.- TODO BIEN
'1.- MAL USUARIO
'2.- MAL PC
Public Function Leer(vIdcodusu As Integer) As Byte
    On Error GoTo Err3

    SQL = "SELECT * FROM usuarios"
    SQL = SQL & " WHERE "
    SQL = SQL & " codusu = " & vIdcodusu
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Rs.EOF Then
        Leer = 1
        Rs.Close
        'ELSE
        Else
            mvarcodusu = Rs!codusu
            mvarNombre = Rs!Nombre
            mvarlogin = Rs!login
            mvarpassword = Rs!Password
            
            
            mvare_dir = DBLet(Rs!maildir, "T")
            mvare_server = DBLet(Rs!mailserver, "T")
            mvare_login = DBLet(Rs!mailuser, "T")
            mvare_pwd = DBLet(Rs!mailpwd, "T")
                
           
            
            'mvarPreferencias = Rs!preferencias
            mvarPC = 0
            Rs.Close
            
            
            'Leemos datos del PC
            If LeerDatosPc Then
                Leer = 0
            
                'Vemos a k grupos pertenece
                FijarGrupo

            Else
                Leer = 2
            End If
            
            
            'Leemos las preferencias
            mvarPreferencias.Leer (mvarcodusu)
            
    End If


Err3:
    If Err.Number <> 0 Then
        Mens = "Se ha producido un error." & vbCrLf
        Mens = Mens & "Número: " & Err.Number & vbCrLf
        Mens = Mens & "Descripción: " & Err.Description
        MsgBox Mens, vbExclamation
        Leer = 1
    End If
    Set Rs = Nothing
End Function



Private Function LeerDatosPc() As Boolean
    SQL = ComputerName
    mvarNomPC = SQL
    mvarVelocidad = 0
    mvarPC = 0
    SQL = "SELECT * FROM equipos where descripcion='" & UCase(SQL) & "'"
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
        mvarCargaIconosExtensiones = (Rs!cargaIconsExt <> 0)
        mvarPC = Rs.Fields(0)
        LeerDatosPc = True
    Else
        'MAL
        'Lo primero k hago es
        LeerDatosPc = False
    End If
    Rs.Close
            
        
End Function


Private Sub FijarGrupo()
Dim i As Integer

    mvarNivel = 10
    Mens = ""
    For i = 1 To vbMaxGrupos
        Mens = Mens & "0"
    Next i
    SQL = "Select * from usuariosgrupos,grupos where usuariosgrupos.codgrupo=grupos.codgrupo AND codusu = " & mvarcodusu
    SQL = SQL & " ORDER BY orden"
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic
    mvarGrupoPpal = -1
    While Not Rs.EOF
        
        i = vbMaxGrupos - Rs.Fields(1) + 1
        If mvarGrupoPpal < 0 Then mvarGrupoPpal = Rs.Fields(1)
        SQL = Mid(Mens, 1, i - 1) & "1" & Mid(Mens, i + 1)
        Mens = SQL
        
        
        'Nivel
        i = CInt(Rs!Nivel)
        If i < mvarNivel Then mvarNivel = i
        
        
        Rs.MoveNext
    Wend
    Rs.Close
    
    mvarGrupo = BiarioLONG(Mens)
    
End Sub

Public Function Modificar() As Byte

    SQL = "UPDATE Usuarios SET nombre='" & mvarNombre
    SQL = SQL & "', login ='" & mvarlogin
    SQL = SQL & "', password ='" & mvarpassword
    
    
    SQL = SQL & "', maildir ='" & mvare_dir
    
    SQL = SQL & "', mailserver ='" & mvare_server
    
    SQL = SQL & "', mailuser ='" & mvare_login
    
    SQL = SQL & "', mailpwd ='" & mvare_pwd & "'"
    
    
    
    
    
    SQL = SQL & " WHERE CODUSU = " & mvarcodusu
    On Error Resume Next
    Conn.Execute SQL
    If Err.Number <> 0 Then
        MuestraError Err.Number, , Err.Description
        Modificar = 1
    Else
        Modificar = 0
    End If
End Function



Public Function ModificarDatosMail() As Byte

    SQL = "UPDATE Usuarios SET "
    
    SQL = SQL & " maildir ='" & mvare_dir
    
    SQL = SQL & "', mailserver ='" & mvare_server
    
    SQL = SQL & "', mailuser ='" & mvare_login
    
    SQL = SQL & "', mailpwd ='" & mvare_pwd & "'"
    
    
    
    
    
    SQL = SQL & " WHERE CODUSU = " & mvarcodusu
    On Error Resume Next
    Conn.Execute SQL
    If Err.Number <> 0 Then
        MuestraError Err.Number, , Err.Description
        ModificarDatosMail = 1
    Else
        ModificarDatosMail = 0
        Me.preferencias.Modificar mvarcodusu, True
    End If
End Function



Public Function Agregar() As Byte
   On Error GoTo EAgr
    'Obtenemos el siguiente codigo
    Siguiente
        
    SQL = "INSERT INTO usuarios (codusu, Nombre, login, password, maildir, mailserver, mailuser, mailpwd,preferencias) VALUES ("
    SQL = SQL & mvarcodusu
    SQL = SQL & ",'" & mvarNombre
    SQL = SQL & "','" & mvarlogin
    SQL = SQL & "','" & mvarpassword
    'email

    SQL = SQL & "','" & mvare_dir
    SQL = SQL & "','" & mvare_server
    SQL = SQL & "','" & mvare_login
    SQL = SQL & "','" & mvare_pwd
    'preferencias
    'mvarpreferencias
    SQL = SQL & "',0)"
    
    
    Conn.Execute SQL
    
    
    'Agregamos su preferencia personal por defecto
    SQL = "INSERT INTO preferenciapersonal (codusu, c1, c2, c3, c4, f1, f2, f3, imp1, imp2, obs, tamayo) VALUES ("
    SQL = SQL & mvarcodusu & ", 100, 100, 0, 0, 100, 0, 0, 0, 100, 0, 0)"
    Conn.Execute SQL
    
    Agregar = 0
    Exit Function
EAgr:
    
        MuestraError Err.Number, , Err.Description
        Agregar = 1
        
End Function


Private Function Siguiente()
    mvarcodusu = 1
    SQL = "Select max(codusu) from Usuarios"
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
        mvarcodusu = DBLet(Rs.Fields(0), "N") + 1
    End If
    Rs.Close
    Set Rs = Nothing
End Function

Public Function Eliminar() As Byte

    On Error GoTo EElim
    
    Mens = "Elimnar grupos asociados"
    SQL = "DELETE FROM usuariosgrupos WHERE codusu =" & mvarcodusu
    Conn.Execute SQL
    
    
    Mens = "Preferencia personal"
    SQL = "DELETE FROM PreferenciaPersonal WHERE codusu =" & mvarcodusu
    Conn.Execute SQL
    
    
    Mens = "Eliminar DATOS usuarios"
    SQL = "DELETE FROM usuarios where codusu =" & mvarcodusu
    Conn.Execute SQL
    
    Eliminar = 0
    Exit Function
EElim:
    MuestraError Err.Number, Mens
    MsgBox "Error inesperado. Consulte soporte técnico", vbExclamation
End Function

Private Sub Class_Initialize()
    Set mvarPreferencias = New CPreferenciaPersonal
End Sub

Private Sub Class_Terminate()
    Set mvarPreferencias = Nothing
    Set Rs = Nothing
End Sub
