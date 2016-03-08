VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CConfiguracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Variables locales que contienen valores de propiedad
Private mvarcodigo As Long 'copia local
Private mvarc1 As String 'copia local
Private mvarc2 As String 'copia local
Private mvarc3 As String 'copia local
Private mvarc4 As String 'copia local
Private mvarf1 As String 'copia local
Private mvarf2 As String 'copia local
Private mvarf3 As String 'copia local
Private mvarimp1 As String 'copia local
Private mvarimp2 As String 'copia local
Private mvarobs As String 'copia local
Private mvarRevisaTareasAPI As Boolean
Private mvarLeyProtDatos1 As String
Private mvarLeyProtDatos2 As String


Private SQL As String  'Para la sentencia
Private Mens As String    'Para los mensajes
Private Rs As ADODB.Recordset 'Para la base de datos
Private Rc As Byte



'------------------------------------------------
'Propiedades del modulo CConfiguracion
'------------------------------------------------


Public Property Let codigo(ByVal vData As Long)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.codigo = 5
     mvarcodigo = vData
End Property

Public Property Get codigo() As Long
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.codigo
     codigo = mvarcodigo
End Property


Public Property Let C1(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.c1 = 5
     mvarc1 = vData
End Property

Public Property Get C1() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.c1
     C1 = mvarc1
End Property


Public Property Let C2(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.c2 = 5
     mvarc2 = vData
End Property

Public Property Get C2() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.c2
     C2 = mvarc2
End Property


Public Property Let c3(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.c3 = 5
     mvarc3 = vData
End Property

Public Property Get c3() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.c3
     c3 = mvarc3
End Property


Public Property Let c4(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.c4 = 5
     mvarc4 = vData
End Property

Public Property Get c4() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.c4
     c4 = mvarc4
End Property


Public Property Let f1(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.f1 = 5
     mvarf1 = vData
End Property

Public Property Get f1() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.f1
     f1 = mvarf1
End Property


Public Property Let f2(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.f2 = 5
     mvarf2 = vData
End Property

Public Property Get f2() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.f2
     f2 = mvarf2
End Property


Public Property Let f3(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.f3 = 5
     mvarf3 = vData
End Property

Public Property Get f3() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.f3
     f3 = mvarf3
End Property


Public Property Let imp1(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.imp1 = 5
     mvarimp1 = vData
End Property

Public Property Get imp1() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.imp1
     imp1 = mvarimp1
End Property


Public Property Let imp2(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.imp2 = 5
     mvarimp2 = vData
End Property

Public Property Get imp2() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.imp2
     imp2 = mvarimp2
End Property


Public Property Let obs(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.obs = 5
     mvarobs = vData
End Property

Public Property Get obs() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.obs
     obs = mvarobs
End Property





Public Property Let LeyProtDatos1(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.obs = 5
     mvarLeyProtDatos1 = vData
End Property

Public Property Get LeyProtDatos1() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.obs
     LeyProtDatos1 = mvarLeyProtDatos1
End Property





Public Property Let LeyProtDatos2(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.obs = 5
     mvarLeyProtDatos2 = vData
End Property

Public Property Get LeyProtDatos2() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.obs
     LeyProtDatos2 = mvarLeyProtDatos2
End Property


Public Property Let RevisaTareasAPI(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.obs = 5
     mvarRevisaTareasAPI = vData
End Property

Public Property Get RevisaTareasAPI() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.obs
     RevisaTareasAPI = mvarRevisaTareasAPI
End Property

'------------------------------------------------
'FUNCIONES del modulo CConfiguracion
'------------------------------------------------




'****   LEER  ***
Public Function Leer(vIdcodigo As Long) As Byte
    On Error GoTo Err3

    SQL = "SELECT * FROM Configuracion"
    SQL = SQL & " WHERE "
    SQL = SQL & " codigo = " & vIdcodigo
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Rs.EOF Then
        Leer = 1
        'ELSE
        Else
            mvarcodigo = Rs!codigo
            mvarc1 = Rs!C1
            mvarc2 = Rs!C2
            mvarc3 = Rs!c3
            mvarc4 = Rs!c4
            mvarf1 = Rs!f1
            mvarf2 = Rs!f2
            mvarf3 = Rs!f3
            mvarimp1 = Rs!imp1
            mvarimp2 = Rs!imp2
            mvarobs = Rs!obs
            mvarRevisaTareasAPI = (Rs!RevisaTareasAPI = 1)
            mvarLeyProtDatos1 = DBLet(Rs!LeyPD1, "T")
            mvarLeyProtDatos2 = DBLet(Rs!LeyPD2, "T")
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
Public Function Modificar() As Byte
    On Error Resume Next
    SQL = "UPDATE  configuracion SET "
    SQL = SQL & " c1= '" & mvarc1
    SQL = SQL & "', c2= '" & mvarc2
    SQL = SQL & "', c3= '" & mvarc3
    SQL = SQL & "', c4= '" & mvarc4
    SQL = SQL & "', f1= '" & mvarf1
    SQL = SQL & "', f2= '" & mvarf2
    SQL = SQL & "', f3= '" & mvarf3
    SQL = SQL & "', imp1= '" & mvarimp1
    SQL = SQL & "', imp2= '" & mvarimp2
    SQL = SQL & "', obs= '" & mvarobs
    SQL = SQL & "'"
    SQL = SQL & ",LeyPD1 = "
    If mvarLeyProtDatos1 = "" Then
        SQL = SQL & "NULL"
    Else
        SQL = SQL & "'" & DevNombreSql(mvarLeyProtDatos1) & "'"
    End If
    SQL = SQL & ",LeyPD2 = "
    If mvarLeyProtDatos2 = "" Then
        SQL = SQL & "NULL"
    Else
        SQL = SQL & "'" & DevNombreSql(mvarLeyProtDatos2) & "'"
    End If
    
    SQL = SQL & " WHERE   codigo=  " & mvarcodigo
    Conn.Execute SQL
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Modifica configuracion"
        Modificar = 1
    Else
        Modificar = 0
    End If
End Function
