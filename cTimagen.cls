VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTimagen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Variables locales que contienen valores de propiedad
Private mvarcodigo As Long 'copia local
Private mvarcodext As Integer 'copia local
Private mvarcodcarpeta As Integer 'copia local
Private mvarcampo1 As String 'copia local
Private mvarcampo2 As String 'copia local
Private mvarcampo3 As String 'copia local
Private mvarcampo4 As String 'copia local
Private mvarfecha1 As Date 'copia local
Private mvarfecha2 As Date 'copia local
Private mvarfecha3 As Date 'copia local
Private mvarimporte1 As Currency 'copia local
Private mvarimporte2 As Currency 'copia local
Private mvarobserva As String 'copia local
Private mvartamnyo As Currency 'copia local
Private mvaruserprop As Long 'copia local
Private mvargroupprop As Long 'copia local
Private mvarlecturau As Long 'copia local
Private mvarlecturag As Long 'copia local
Private mvarescriturau As Long 'copia local
Private mvarescriturag As Long 'copia local



Private mvarLlevaHco As Boolean

'Si la aplicacion lleva tabla de modificaciones
'----------------------------------------------
Private mvarIniciales As Collection







Private SQL As String  'Para la sentencia
Private Mens As String    'Para los mensajes
Private Rs As ADODB.Recordset 'Para la base de datos
Private Rc As Byte



'------------------------------------------------
'Propiedades del modulo cTimagen
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


Public Property Let campo1(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.campo1 = 5
     mvarcampo1 = vData
End Property

Public Property Get campo1() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.campo1
     campo1 = mvarcampo1
End Property


Public Property Let campo2(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.campo2 = 5
     mvarcampo2 = vData
End Property

Public Property Get campo2() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.campo2
     campo2 = mvarcampo2
End Property


Public Property Let campo3(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.campo3 = 5
     mvarcampo3 = vData
End Property

Public Property Get campo3() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.campo3
     campo3 = mvarcampo3
End Property


Public Property Let campo4(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.campo4 = 5
     mvarcampo4 = vData
End Property

Public Property Get campo4() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.campo4
     campo4 = mvarcampo4
End Property


Public Property Let fecha1(ByVal vData As Date)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.fecha1 = 5
     mvarfecha1 = vData
End Property

Public Property Get fecha1() As Date
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.fecha1
     fecha1 = mvarfecha1
End Property


Public Property Let fecha2(ByVal vData As Date)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.fecha2 = 5
     mvarfecha2 = vData
End Property

Public Property Get fecha2() As Date
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.fecha2
     fecha2 = mvarfecha2
End Property


Public Property Let fecha3(ByVal vData As Date)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.fecha3 = 5
     mvarfecha3 = vData
End Property

Public Property Get fecha3() As Date
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.fecha3
     fecha3 = mvarfecha3
End Property


Public Property Let importe1(ByVal vData As Currency)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.importe1 = 5
     mvarimporte1 = vData
End Property

Public Property Get importe1() As Currency
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.importe1
     importe1 = mvarimporte1
End Property


Public Property Let importe2(ByVal vData As Currency)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.importe2 = 5
     mvarimporte2 = vData
End Property

Public Property Get importe2() As Currency
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.importe2
     importe2 = mvarimporte2
End Property


Public Property Let observa(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.observa = 5
     mvarobserva = vData
End Property

Public Property Get observa() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.observa
     observa = mvarobserva
End Property


Public Property Let tamnyo(ByVal vData As Currency)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.tamnyo = 5
     mvartamnyo = vData
End Property

Public Property Get tamnyo() As Currency
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.tamnyo
     tamnyo = mvartamnyo
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




'------------------------------------------------
'FUNCIONES del modulo cTimagen
'------------------------------------------------



'****   AGREGAR  ***
Public Function Agregar(LlevarHcoRevisiones As Boolean, Optional Hueco As Boolean) As Byte
    On Error GoTo Err1

    
    Agregar = 1
    
    

    'Obtengo siguiente
    Set Rs = New ADODB.Recordset
    
    If Hueco Then
        'Esto lo utlizaremos para cuando inserta una sola imagen,
        'k buscque el hueco
        BuscaHueco
        
        
    Else
        Siguiente
    End If
    
    
    SQL = "INSERT INTO timagen (codigo, codext, codcarpeta, campo1, campo2, campo3, campo4,"
    SQL = SQL & "fecha1, fecha2, fecha3, importe1, importe2, observa,tamnyo,"
    SQL = SQL & "  userprop, groupprop, lecturau, lecturag, escriturau, escriturag) VALUES ("
    'Valores
    SQL = SQL & mvarcodigo & "," & mvarcodext & "," & mvarcodcarpeta
    SQL = SQL & ",'" & mvarcampo1 & "',"
    SQL = SQL & CampoANulo(mvarcampo2) & ","
    SQL = SQL & CampoANulo(mvarcampo3) & ","        '
    SQL = SQL & CampoANulo(mvarcampo4) & ",'"
    
    SQL = SQL & Format(mvarfecha1, FormatoFecha) & "',"
    SQL = SQL & CampoANulo(mvarfecha2, "F") & ","
    SQL = SQL & CampoANulo(mvarfecha3, "F") & ","
    SQL = SQL & CampoANulo(mvarimporte1, "N") & ","
    SQL = SQL & CampoANulo(mvarimporte2, "N") & ","
    SQL = SQL & CampoANulo(mvarobserva) & ","
    SQL = SQL & TransformaComasPuntos(CStr(mvartamnyo))
    
    SQL = SQL & "," & mvaruserprop
    SQL = SQL & "," & mvargroupprop
    SQL = SQL & "," & mvarlecturau
    SQL = SQL & "," & mvarlecturag
    SQL = SQL & "," & mvarescriturau
    SQL = SQL & "," & mvarescriturag
    SQL = SQL & ")"
    Conn.Execute SQL
    
    mvarLlevaHco = LlevarHcoRevisiones
    If mvarLlevaHco Then InsertaLineaRegistroCambios 0
    
    Agregar = 0
    Exit Function
Err1:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    
End Function


'****   ELIMINAR  ***
Public Function Eliminar() As Byte
    On Error GoTo Err2

    SQL = "DELETE FROM timagen"
    SQL = SQL & " WHERE "
    SQL = SQL & " codigo = " & mvarcodigo
    Conn.Execute SQL
    
    Exit Function

Err2:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    Eliminar = 1
End Function


'****   LEER  ***
Public Function Leer(vIdcodigo As Long, LlevaRegistroCambios As Boolean) As Byte
    On Error GoTo Err3

    SQL = "SELECT * FROM timagen"
    If ModoTrabajo = vbHistAnt Or ModoTrabajo = vbHistNue Then SQL = SQL & "hco"
    SQL = SQL & " WHERE "
    SQL = SQL & " codigo = " & vIdcodigo
    If ModoTrabajo = vbHistAnt Or ModoTrabajo = vbHistNue Then SQL = SQL & " and codequipo =" & vUsu.PC
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Rs.EOF Then
        Leer = 1
        'ELSE
        Else
            mvarcodigo = Rs!codigo
            mvarcodext = Rs!codext
            mvarcodcarpeta = Rs!codcarpeta
            mvarcampo1 = Rs!campo1
            mvarcampo2 = DBLet(Rs!campo2)
            mvarcampo3 = DBLet(Rs!campo3)
            mvarcampo4 = DBLet(Rs!campo4)
            mvarfecha1 = Rs!fecha1
            mvarfecha2 = DBLet(Rs!fecha2, "F")
            mvarfecha3 = DBLet(Rs!fecha3, "F")
            mvarimporte1 = DBLet(Rs!importe1, "N")
            mvarimporte2 = DBLet(Rs!importe2, "N")
            
            mvartamnyo = Rs!tamnyo
            mvaruserprop = Rs!userprop
            mvargroupprop = Rs!groupprop
            mvarlecturau = Rs!lecturau
            mvarlecturag = Rs!lecturag
            mvarescriturau = Rs!escriturau
            mvarescriturag = Rs!escriturag
            
            'Campo MEMO observaciones
            'mvarobserva = DBLet(Rs!observa)
            AsignarCampoMemo mvarobserva, "observa", Rs
            
            
            mvarLlevaHco = LlevaRegistroCambios
            If mvarLlevaHco Then
                'ASignamos los campos
                InsertarEnVariableDatosIniciales
            End If
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

            Modificar = 1
            On Error GoTo EModifi
            SQL = "UPDATE Timagen SET"
            SQL = SQL & " campo1 = '" & DevNombreSql(mvarcampo1)
            SQL = SQL & "',campo2 = " & ParaBD(mvarcampo2, "T", True)
            SQL = SQL & ", campo3 = " & ParaBD(mvarcampo3, "T", True)
            SQL = SQL & ", campo4 = " & ParaBD(mvarcampo4, "T", True)
            SQL = SQL & ", fecha1 = " & ParaBD(mvarfecha1, "F", False)
            SQL = SQL & ", fecha2 = " & ParaBD(mvarfecha2, "F", True)
            SQL = SQL & ", fecha3 = " & ParaBD(mvarfecha3, "F", True)
            SQL = SQL & ", importe1 = " & TransformaComasPuntos(ParaBD(mvarimporte1, "N", True))
            SQL = SQL & ", importe2 = " & TransformaComasPuntos(ParaBD(mvarimporte2, "N", True))
            SQL = SQL & ", observa = " & ParaBD(mvarobserva, "T", True)
            SQL = SQL & ", tamnyo = " & TransformaComasPuntos(CStr(mvartamnyo))
            
            
            'Permisos y demas
            SQL = SQL & ",userprop = " & mvaruserprop
            SQL = SQL & ",groupprop = " & mvargroupprop
            SQL = SQL & ",lecturau = " & mvarlecturau
            SQL = SQL & ",lecturag = " & mvarlecturag
            SQL = SQL & ",escriturau = " & mvarescriturau
            SQL = SQL & ",escriturag = " & mvarescriturag
            
            SQL = SQL & " WHERE codigo = " & mvarcodigo
            Conn.Execute SQL
            
            
            If mvarLlevaHco Then ComprobarCambioValores
            
            Modificar = 0
    Exit Function
EModifi:
    MuestraError Err.Number, "Objeto IMAGEN: Modifica"
End Function


'****   SIGUIENTE  ***
Public Function Siguiente() As Long

    Siguiente = 1
    SQL = "SELECT MAX(codigo) FROM timagen"
    Rs.Open SQL, Conn, , , adCmdText
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then Siguiente = Rs.Fields(0) + 1
    End If
    mvarcodigo = Siguiente
    Rs.Close
End Function


Public Function BuscaHueco() As Long
Dim C1 As Long
Dim Fin As Boolean

    SQL = "SELECT codigo FROM timagen ORDER BY codigo"
    C1 = 1
    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Fin = Rs.EOF
    While Not Fin
        If (Rs!codigo - C1) > 1 Then
            'Salto
            Fin = True
        Else
            C1 = Rs!codigo + 1
            Rs.MoveNext
            Fin = Rs.EOF
        End If
    Wend
    Rs.Close
    mvarcodigo = C1
End Function

Private Sub Class_Terminate()
    Set Rs = Nothing
    Set mvarIniciales = Nothing
End Sub




Private Sub InsertarEnVariableDatosIniciales()
    On Error GoTo EIn

            Set mvarIniciales = New Collection
            
            mvarIniciales.Add CStr(mvarcodcarpeta)
            mvarIniciales.Add CStr(mvarcampo1)
            mvarIniciales.Add CStr(mvarcampo2)
            mvarIniciales.Add CStr(mvarcampo3)
            mvarIniciales.Add CStr(mvarcampo4)
            mvarIniciales.Add CStr(mvarfecha1)
            mvarIniciales.Add CStr(mvarfecha2)
            mvarIniciales.Add CStr(mvarfecha3)
            mvarIniciales.Add CStr(mvarimporte1)
            mvarIniciales.Add CStr(mvarimporte2)
            
            'mvarDatosInciales = mvarDatosInciales & mvartamnyo & "|"
            mvarIniciales.Add CStr(mvaruserprop)
            mvarIniciales.Add CStr(mvargroupprop)
            mvarIniciales.Add CStr(mvarlecturau)
            mvarIniciales.Add CStr(mvarlecturag)
            mvarIniciales.Add CStr(mvarescriturau)
            mvarIniciales.Add CStr(mvarescriturag)
            Exit Sub
EIn:
    Err.Clear
    mvarLlevaHco = False   'Para que luego no intente guardar cambios
End Sub



Private Sub ComprobarCambioValores()
Dim Cambios As String
Dim I As Integer
            Cambios = ""
            If mvarIniciales(1) <> CStr(mvarcodcarpeta) Then Cambios = Cambios & "1|"
            If mvarIniciales(2) <> CStr(mvarcampo1) Then Cambios = Cambios & "2|"
            If mvarIniciales(3) <> CStr(mvarcampo2) Then Cambios = Cambios & "3|"
            If mvarIniciales(4) <> CStr(mvarcampo3) Then Cambios = Cambios & "4|"
            If mvarIniciales(5) <> CStr(mvarcampo4) Then Cambios = Cambios & "5|"
            If mvarIniciales(6) <> CStr(mvarfecha1) Then Cambios = Cambios & "6|"
            If mvarIniciales(7) <> CStr(mvarfecha2) Then Cambios = Cambios & "7|"
            If mvarIniciales(8) <> CStr(mvarfecha3) Then Cambios = Cambios & "8|"
            If mvarIniciales(9) <> CStr(mvarimporte1) Then Cambios = Cambios & "9|"
            If mvarIniciales(10) <> CStr(mvarimporte2) Then Cambios = Cambios & "10|"

            'mvarDatosInciales = mvarDatosInciales & mvartamnyo & "|"
            If mvarIniciales(11) <> CStr(mvaruserprop) Then Cambios = Cambios & "11|"
            If mvarIniciales(12) <> CStr(mvargroupprop) Then Cambios = Cambios & "12|"
            If mvarIniciales(13) <> CStr(mvarlecturau) Then Cambios = Cambios & "13|"
            If mvarIniciales(14) <> CStr(mvarlecturag) Then Cambios = Cambios & "14|"
            If mvarIniciales(15) <> CStr(mvarescriturau) Then Cambios = Cambios & "15|"
            If mvarIniciales(16) <> CStr(mvarescriturag) Then Cambios = Cambios & "16|"

    
            If Cambios <> "" Then
                SQL = ""
                While Cambios <> ""
                    I = InStr(1, Cambios, "|")
                    If I = 0 Then
                        Cambios = ""
                    Else
                        Mens = Mid(Cambios, 1, I - 1)
                        Cambios = Mid(Cambios, I + 1)
                        Mens = "CampoBD " & Mens & " :" & DevNombreSql(mvarIniciales(CInt(Mens)))
                        SQL = SQL & Mens
                    End If
                Wend
                
                InsertaLineaRegistroCambios 1, SQL
                SQL = ""
            End If
End Sub



'  2.- Modificado
'  1.- Eliminar
'  0.- Nuevo
Private Sub InsertaLineaRegistroCambios(Opcion As Byte, Optional ByRef CADENA As String)
    'ESTE TROZO ES DEMASIADO PERSONALIZADO.
    'Utiliza un objeto que tiene que previamente haber sido declarado
    'objRevision
    
    
    On Error GoTo EInsertaLineaRegistroCambios
    
    objRevision.InsertaRevision mvarcodigo, Opcion, vUsu, CADENA
    
    
    
    Exit Sub
EInsertaLineaRegistroCambios:
    Err.Clear
    mvarLlevaHco = False
End Sub



