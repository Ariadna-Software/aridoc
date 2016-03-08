VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPreferenciaPersonal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Variables locales que contienen valores de propiedad
Private mvarUsuario As Long 'copia local
Private mvarc1 As Integer 'copia local
Private mvarc2 As Integer 'copia local
Private mvarc3 As Integer 'copia local
Private mvarc4 As Integer 'copia local
Private mvarf1 As Integer 'copia local
Private mvarf2 As Integer 'copia local
Private mvarf3 As Integer 'copia local
Private mvarimp1 As Integer 'copia local
Private mvarimp2 As Integer 'copia local
Private mvarobs As Integer 'copia local
Private mvartamayo As Integer 'copia local
Private mvarVista As Byte
Private mvarSelect As String
Private mvarAncho As Integer  'Ancho
Private mvarORDERBY As String


'------------------------------------
'Configuracion para el mail
Private mvarmailInicio As Boolean
    'Si cuando arranca aridoc  comprueba mensajes pendientes
Private mvarmailFiltro As Byte
    'Si filtra los mensjes al mostrar
    
Private mvarmailPasarHCO As Byte
    'Pasar autmoaticamente mensajes reciidos o enviados a HCO
    '   0.- NO
    '   N.- Numero de meses.  Entre 0 y 12
    
    
Private SQL As String  'Para la sentencia
Private Mens As String    'Para los mensajes
Private Rs As ADODB.Recordset 'Para la base de datos
Private Rc As Byte



'------------------------------------------------
'Propiedades del modulo CPreferenciaPersonal
'------------------------------------------------


Public Property Let Usuario(ByVal vData As Long)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.Usuario = 5
     mvarUsuario = vData
End Property

Public Property Get Usuario() As Long
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.Usuario
     Usuario = mvarUsuario
End Property




Public Property Let Vista(ByVal vData As Byte)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.Usuario = 5
     mvarVista = vData
End Property

Public Property Get Vista() As Byte
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.Usuario
     Vista = mvarVista
End Property

Public Property Let C1(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.c1 = 5
     mvarc1 = vData
End Property

Public Property Get C1() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.c1
     C1 = mvarc1
End Property


Public Property Let C2(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.c2 = 5
     mvarc2 = vData
End Property

Public Property Get C2() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.c2
     C2 = mvarc2
End Property


Public Property Let c3(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.c3 = 5
     mvarc3 = vData
End Property

Public Property Get c3() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.c3
     c3 = mvarc3
End Property


Public Property Let c4(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.c4 = 5
     mvarc4 = vData
End Property

Public Property Get c4() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.c4
     c4 = mvarc4
End Property


Public Property Let f1(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.f1 = 5
     mvarf1 = vData
End Property

Public Property Get f1() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.f1
     f1 = mvarf1
End Property


Public Property Let f2(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.f2 = 5
     mvarf2 = vData
End Property

Public Property Get f2() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.f2
     f2 = mvarf2
End Property


Public Property Let f3(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.f3 = 5
     mvarf3 = vData
End Property

Public Property Get f3() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.f3
     f3 = mvarf3
End Property


Public Property Let imp1(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.imp1 = 5
     mvarimp1 = vData
End Property

Public Property Get imp1() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.imp1
     imp1 = mvarimp1
End Property


Public Property Let imp2(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.imp2 = 5
     mvarimp2 = vData
End Property

Public Property Get imp2() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.imp2
     imp2 = mvarimp2
End Property


Public Property Let obs(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.obs = 5
     mvarobs = vData
End Property

Public Property Get obs() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.obs
     obs = mvarobs
End Property


Public Property Let tamayo(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.tamayo = 5
     mvartamayo = vData
End Property

Public Property Get tamayo() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.tamayo
     tamayo = mvartamayo
End Property

Public Property Let Ancho(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.tamayo = 5
     mvarAncho = vData
End Property

Public Property Get Ancho() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.tamayo
     Ancho = mvarAncho
End Property


Public Property Let ORDERBY(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.tamayo = 5
     mvarORDERBY = vData
End Property

Public Property Get ORDERBY() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.tamayo
     ORDERBY = mvarORDERBY
End Property


Public Property Get vSelect() As String
    vSelect = mvarSelect
End Property



Public Property Let mailInicio(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.tamayo = 5
     mvarmailInicio = vData
End Property

Public Property Get mailInicio() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.tamayo
     mailInicio = mvarmailInicio
End Property


Public Property Let mailFiltro(ByVal vData As Byte)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.tamayo = 5
     mvarmailFiltro = vData
End Property

Public Property Get mailFiltro() As Byte
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.tamayo
     mailFiltro = mvarmailFiltro
End Property

Public Property Let mailPasarHCO(ByVal vData As Byte)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.tamayo = 5
     mvarmailPasarHCO = vData
End Property

Public Property Get mailPasarHCO() As Byte
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.tamayo
     mailPasarHCO = mvarmailPasarHCO
End Property






'****   LEER  ***
Public Function Leer(vIdUsuario As Long) As Byte
    On Error GoTo Err3

    Set Rs = New ADODB.Recordset
    'Ponemos el recordset en modo:
    Rs.CursorType = adOpenKeyset
    Rs.LockType = adLockOptimistic


    SQL = "SELECT * FROM PreferenciaPersonal"
    SQL = SQL & " WHERE "
    SQL = SQL & " codusu = " & vIdUsuario
    Rs.Open SQL, Conn, , , adCmdText
    If Rs.EOF Then
        Leer = 1
        
            'Pero tengo k poner unos valores maximos
            mvarSelect = "campo1"
            mvarVista = 1
            mvarAncho = 50
            mvarORDERBY = "" 'DBLet(Rs!ORDERBY, "T")
        
        'ELSE
        Else
            mvarUsuario = Rs!codusu
            mvarSelect = ""
            
            
            mvarc1 = Rs!C1
            If mvarc1 > 0 Then mvarSelect = mvarSelect & ",campo1"
            mvarc2 = Rs!C2
            If mvarc2 > 0 Then mvarSelect = mvarSelect & ",campo2"
            mvarc3 = Rs!c3
            If mvarc3 > 0 Then mvarSelect = mvarSelect & ",campo3"
            mvarc4 = Rs!c4
            If mvarc4 > 0 Then mvarSelect = mvarSelect & ",campo4"
            mvarf1 = Rs!f1
            If mvarf1 > 0 Then mvarSelect = mvarSelect & ",fecha1"
            mvarf2 = Rs!f2
            If mvarf2 > 0 Then mvarSelect = mvarSelect & ",fecha2"
            mvarf3 = Rs!f3
            If mvarf3 > 0 Then mvarSelect = mvarSelect & ",fecha3"
            mvarimp1 = Rs!imp1
            If mvarimp1 > 0 Then mvarSelect = mvarSelect & ",importe1"
            mvarimp2 = Rs!imp2
            If mvarimp2 > 0 Then mvarSelect = mvarSelect & ",importe2"
            mvarobs = Rs!obs
            If mvarobs > 0 Then mvarSelect = mvarSelect & ",observa"
            mvartamayo = Rs!tamayo
            If mvartamayo > 0 Then mvarSelect = mvarSelect & ",tamnyo"
            
            mvarSelect = Mid(mvarSelect, 2) 'Quitamos la primera coma
            mvarVista = Rs!Vista
            mvarAncho = Rs!Ancho
            mvarORDERBY = DBLet(Rs!ORDERBY, "T")
            
            
            
            'NUevo. Los valores para los e-mails
            '-----------------------------------
            mvarmailInicio = Rs!mailInicio
            mvarmailFiltro = Rs!mailFiltro
            mvarmailPasarHCO = Rs!mailPasarHCO
            '`mailInicio` mailFiltro mailPasarHCO
            
            Leer = 0
    End If

    Rs.Close


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


'****   MODIFICAR  ***
Public Function Modificar(Usu As Integer, email As Boolean) As Byte
    On Error GoTo Err4

    SQL = "UPDATE preferenciapersonal SET "
    If email Then
            SQL = SQL & " mailInicio =" & Abs(mvarmailInicio)
            SQL = SQL & ", mailFiltro = " & mvarmailFiltro
            SQL = SQL & ", mailPasarHCO  = " & mvarmailPasarHCO
     
            SQL = SQL & " WHERE codusu =" & Usu
    
    
    Else
            SQL = SQL & " C1 = " & mvarc1
            SQL = SQL & ", C2 = " & mvarc2
            SQL = SQL & ", c3 = " & mvarc3
            SQL = SQL & ", c4 = " & mvarc4
            SQL = SQL & ", f1 = " & mvarf1
            SQL = SQL & ", f2 = " & mvarf2
            SQL = SQL & ", f3 = " & mvarf3
            SQL = SQL & ", imp1 = " & mvarimp1
            SQL = SQL & ", imp2 = " & mvarimp2
            SQL = SQL & ", obs = " & mvarobs
            SQL = SQL & ", tamayo = " & mvartamayo
            SQL = SQL & ", vista =" & mvarVista
            SQL = SQL & ", ancho = " & mvarAncho
            SQL = SQL & ", ORDERBY = '" & mvarORDERBY & "'"
            
            SQL = SQL & " WHERE codusu =" & Usu

    End If

    Conn.Execute SQL
    Modificar = 0
    Exit Function

Err4:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    Modificar = 1
End Function



Private Sub Class_Terminate()
    Set Rs = Nothing
End Sub


