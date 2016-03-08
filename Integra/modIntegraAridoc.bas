Attribute VB_Name = "modIntegraAridoc"
Option Explicit

Public vConfig As Configuracion
Public RevisarPendientes As Boolean


Public Const vbPermisoTotal = 2147483647
Public ErrorLlevando As Boolean

Public Conn As Connection
Public miRsAux As ADODB.Recordset
Public CodPC As Integer
Public CarpetaErroresCreada As Boolean

Public Sub Main()
Dim miNombre As String

    If App.PrevInstance Then Exit Sub
    
    Set vConfig = New Configuracion
    
    If vConfig.Leer = 1 Then
        MsgBox "Error en la configuracion del Integrador de ARIDOC", vbExclamation
        vConfig.Grabar
        Exit Sub
    End If
    
    If Dir(vConfig.PathError, vbDirectory) = "" Then
        MsgBox "Error leyendo carpeta errores", vbExclamation
        Exit Sub
    End If
    
    'Abri conexion
    If Not AbrirConexion Then
        MsgBox "Error abriendo conexion BD Aridoc", vbExclamation
        Set Conn = Nothing
        Exit Sub
    End If
    
    If Not LeerDatosPc Then
        Set Conn = Nothing
        Exit Sub
    End If
    
    If HayArchivosParaIntegrar Then
               
               
        CarpetaErroresCreada = False
        frmIntegraciones.Show vbModal
        If RevisarPendientes Then
            'If QuedanArchivosSueltos Then MoverArchivosSueltos
            miNombre = CrearCarpetaErrores("")
                If miNombre = "" Then
                    MsgBox "Se han producido errores. IMPOSIBLE GENERAR CARPETA ERROR. No realice mas integraciones. CONSULTE SOPORTE TECNICO", vbCritical
                    End
                Else
                    MoverTodosLosArchivos miNombre
                    MsgBox "Se han producido errores. La aplicación finalizara.", vbCritical
                End If
        End If
    End If
    
    
    'cerrar conexion
    Conn.Close
    Set Conn = Nothing
    
End Sub


Private Sub MoverTodosLosArchivos(Ruta As String)
Dim miNombre As String


On Error GoTo EMOv
    miNombre = Dir(vConfig.PathArchivos & "\", vbArchive)     ' Recupera la primera entrada.
    Do While miNombre <> ""   ' Inicia el bucle.
       FileCopy vConfig.PathArchivos & "\" & miNombre, Ruta & "\" & miNombre
       Kill vConfig.PathArchivos & "\" & miNombre
       miNombre = Dir   ' Obtiene siguiente entrada.
    Loop
    
    MsgBox "Errores se han llevado a " & Ruta, vbExclamation
    Exit Sub
EMOv:
     MsgBox "Se han producido errores. MOVIENDO A CARPETA ERROR. No realice mas integraciones. CONSULTE SOPORTE TECNICO" & vbCrLf & Err.Description, vbCritical
End Sub



Public Function HayArchivosParaIntegrar() As Boolean
    On Error Resume Next
    HayArchivosParaIntegrar = False
    If Dir(vConfig.PathArchivos & "\*." & vConfig.extensionGuia) <> "" Then HayArchivosParaIntegrar = True
    If Err.Number <> 0 Then
        MsgBox "Error leyendo datos en la carpeta: " & "", vbExclamation
    End If
End Function






Public Function AbrirConexion() As Boolean
Dim Cad As String
On Error GoTo EAbrirConexion

    
    AbrirConexion = False
    Set Conn = Nothing
    Set Conn = New Connection
    'Conn.CursorLocation = adUseClient   'Si ponemos este hay opciones k no van ej select con rs!campo
    Conn.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente
                        
    
    
    
    'cadenaconexion
    Cad = "DSN=Aridoc;DESC= DSN;DATABASE=aridoc;;;PORT=;OPTION=;STMT=;"
    Cad = "DSN=Aridoc;"" "
    'cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=" & vUsu.CadenaConexion & ";SERVER=" & vConfig.SERVER & ";"
    'cad = cad & ";UID=" & vConfig.User
    'cad = cad & ";PWD=" & vConfig.password
    
    
    Conn.ConnectionString = Cad
    Conn.Open
    Conn.Execute "Set AUTOCOMMIT = 1"
    AbrirConexion = True
    Exit Function
EAbrirConexion:
    MsgBox "Abrir conexión." & Err.Description, vbExclamation
End Function


Private Function LeerDatosPc() As Boolean
Dim SQL As String

    SQL = ComputerName
    SQL = "SELECT * FROM equipos where descripcion='" & UCase(SQL) & "'"
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        CodPC = miRsAux.Fields(0)
        LeerDatosPc = True
    Else
        LeerDatosPc = False
        MsgBox "PC no ha sido dado de alta en ARIDOC.", vbExclamation
    End If
    miRsAux.Close
    Set miRsAux = Nothing
End Function

Public Function DevNombreSql(Cad As String) As String
Dim I As Integer
Dim J As Integer

    'Buscamos las '
    J = 1
    Do
        I = InStr(J, Cad, "'")
        If I > 0 Then
            Cad = Mid(Cad, 1, I) & "'" & Mid(Cad, I + 1)
            J = I + 2
        End If
    Loop Until I = 0

    'Buscamos los \
    J = 1
    Do
        I = InStr(J, Cad, "\")
        If I > 0 Then
            Cad = Mid(Cad, 1, I) & "\" & Mid(Cad, I + 1)
            J = I + 2
        End If
    Loop Until I = 0
    DevNombreSql = Cad
    
    

End Function





Public Function DBLet(vData As Variant, Optional Tipo As String) As Variant
    If IsNull(vData) Then
        Select Case Tipo
            Case "N"
                DBLet = 0
            Case "F"
                DBLet = "0:00:00"
            Case Else
                DBLet = ""
        End Select
    Else
        DBLet = vData
    End If
End Function


'Para cuando insertamos en BD, si el texto es "" y pondremos NULL
Public Function CampoANulo(ByRef T As Variant, Optional Tipo As String) As String
    Select Case Tipo
    Case "N"
        If T = 0 Then
            CampoANulo = "NULL"
        Else
            CampoANulo = TransformaComasPuntos(CStr(T))
        End If
    
    Case "F"
        If Val(T) = 0 Then
            CampoANulo = "NULL"
        Else
            CampoANulo = "'" & Format(T, "yyyy-mm-dd") & "'"
        End If
    Case Else
        If T = "" Then
            CampoANulo = "NULL"
        Else
            CampoANulo = "'" & DevNombreSql(CStr(T)) & "'"
        End If
    End Select
End Function


'por comas
Public Function TransformaComasPuntos(CADENA As String) As String
    Dim I As Integer
    Do
        I = InStr(1, CADENA, ",")
        If I > 0 Then
            CADENA = Mid(CADENA, 1, I - 1) & "." & Mid(CADENA, I + 1)
        End If
        Loop Until I = 0
    TransformaComasPuntos = CADENA
End Function


Public Function espera(Segundos As Single)
    Dim T1
    T1 = Timer
    Do
    Loop Until Timer - T1 > Segundos
End Function



Public Function CrearCarpetaErrores(vGlobal As String) As String
Dim I As Integer
Dim Cad As String
    
    CrearCarpetaErrores = ""
    I = -1
    Do
       I = I + 1
       If vGlobal = "" Then
            Cad = Format(Now, "yymmdd") & Format(I, "000")
       Else
            Cad = "z" & vGlobal & Format(I, "000")
       End If
       Cad = vConfig.PathError & "\" & Cad
       If Dir(Cad, vbDirectory) = "" Then
            MkDir Cad
            I = 1000
       End If
     Loop Until I > 999
     If I > 1000 Then
        MsgBox "1000. Error"
    Else
        CrearCarpetaErrores = Cad
        CarpetaErroresCreada = True
    End If
   
End Function



Public Function EliminaArchivo(Orig As String)
    On Error Resume Next
    Kill Orig
    If Err.Number <> 0 Then
        MsgBox "Error eliminando: " & Orig & vbCrLf & "Eliminelo a mano.", vbExclamation
        Err.Clear
    End If
End Function
