Attribute VB_Name = "BaseDatos"
Option Explicit

Private Const NoValidosMSDOS = "><|.?,*/"""


Public Function AbrirConexion(Normal As Boolean) As Boolean
Dim cad As String
On Error GoTo EAbrirConexion

    
    AbrirConexion = False
    Set Conn = Nothing
    Set Conn = New Connection
    'Conn.CursorLocation = adUseClient   'Si ponemos este hay opciones k no van ej select con rs!campo
    Conn.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente
                        
    
    
    
    'cadenaconexion
    cad = "DSN=Aridoc;DESC= DSN;DATABASE="
    If Normal Then
        cad = cad & "Aridoc"
    Else
        'RECUPERAMOS BACKUP
        cad = cad & "backAridoc"
    End If
    cad = cad & ";PORT=;OPTION=3;STMT=;"
    'Cad = "DSN=Aridoc;"" "
    'cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=" & vUsu.CadenaConexion & ";SERVER=" & vConfig.SERVER & ";"
    'cad = cad & ";UID=" & vConfig.User
    'cad = cad & ";PWD=" & vConfig.password
    
    
    Conn.ConnectionString = cad
    Conn.Open
    Conn.Execute "Set AUTOCOMMIT = 1"
    AbrirConexion = True
    Exit Function
EAbrirConexion:
    MsgBox "Abrir conexión." & Err.Description, vbExclamation
End Function






Public Function SeparaCampoBusqueda(Tipo As String, Campo As String, CADENA As String, ByRef DevSQL As String) As Byte
Dim cad As String
Dim Aux As String
Dim Ch As String
Dim Fin As Boolean
Dim i, J As String

On Error GoTo ErrSepara
SeparaCampoBusqueda = 1
DevSQL = ""
cad = ""
Select Case Tipo
Case "N"
    '----------------  NUMERICO  ---------------------
    i = CararacteresCorrectos(CADENA, "N")
    If i > 0 Then Exit Function  'Ha habido un error y salimos
    'Comprobamos si hay intervalo ':'
    i = InStr(1, CADENA, ":")
    If i > 0 Then
        'Intervalo numerico
        cad = Mid(CADENA, 1, i - 1)
        Aux = Mid(CADENA, i + 1)
        If Not IsNumeric(cad) Or Not IsNumeric(Aux) Then Exit Function  'No son numeros
        'Intervalo correcto
        'Construimos la cadena
        DevSQL = Campo & " >= " & cad & " AND " & Campo & " <= " & Aux
        '----
        'ELSE
        Else
            'Prueba
            'Comprobamos que no es el mayor
            If CADENA = ">>" Or CADENA = "<<" Then
                DevSQL = "1=1"
             Else
                    Fin = False
                    i = 1
                    cad = ""
                    Aux = "NO ES NUMERO"
                    While Not Fin
                        Ch = Mid(CADENA, i, 1)
                        If Ch = ">" Or Ch = "<" Or Ch = "=" Then
                            cad = cad & Ch
                            Else
                                Aux = Mid(CADENA, i)
                                Fin = True
                        End If
                        i = i + 1
                        If i > Len(CADENA) Then Fin = True
                    Wend
                    'En aux debemos tener el numero
                    If Not IsNumeric(Aux) Then Exit Function
                    'Si que es numero. Entonces, si Cad="" entronces le ponemos =
                    If cad = "" Then cad = " = "
                    DevSQL = Campo & " " & cad & " " & Aux
            End If
        End If
Case "F"
     '---------------- FECHAS ------------------
    i = CararacteresCorrectos(CADENA, "F")
    If i = 1 Then Exit Function
    'Comprobamos si hay intervalo ':'
    i = InStr(1, CADENA, ":")
    If i > 0 Then
        'Intervalo de fechas
        cad = Mid(CADENA, 1, i - 1)
        Aux = Mid(CADENA, i + 1)
        If Not EsFechaOKString(cad) Or Not EsFechaOKString(Aux) Then Exit Function  'Fechas incorrectas
        'Intervalo correcto
        'Construimos la cadena
        cad = Format(cad, FormatoFecha)
        Aux = Format(Aux, FormatoFecha)
        'En my sql es la ' no el #
        'DevSQL = Campo & " >=#" & Cad & "# AND " & Campo & " <= #" & AUX & "#"
        DevSQL = Campo & " >='" & cad & "' AND " & Campo & " <= '" & Aux & "'"
        '----
        'ELSE
        Else
            'Comprobamos que no es el mayor
            If CADENA = ">>" Or CADENA = "<<" Then
                  DevSQL = "1=1"
            Else
                Fin = False
                i = 1
                cad = ""
                Aux = "NO ES FECHA"
                While Not Fin
                    Ch = Mid(CADENA, i, 1)
                    If Ch = ">" Or Ch = "<" Or Ch = "=" Then
                        cad = cad & Ch
                        Else
                            Aux = Mid(CADENA, i)
                            Fin = True
                    End If
                    i = i + 1
                    If i > Len(CADENA) Then Fin = True
                Wend
                'En aux debemos tener el numero
                If Not EsFechaOKString(Aux) Then Exit Function
                'Si que es numero. Entonces, si Cad="" entronces le ponemos =
                Aux = "'" & Format(Aux, FormatoFecha) & "'"
                If cad = "" Then cad = " = "
                DevSQL = Campo & " " & cad & " " & Aux
            End If
        End If
    
    
    
    
Case "T"
    '---------------- TEXTO ------------------
    i = CararacteresCorrectos(CADENA, "T")
    If i = 1 Then Exit Function
    
    'Comprobamos que no es el mayor
     If CADENA = ">>" Or CADENA = "<<" Then
        DevSQL = "1=1"
        Exit Function
    End If
    'Cambiamos el * por % puesto que en ADO es el caraacter para like
    i = 1
    Aux = CADENA
    While i <> 0
        i = InStr(1, Aux, "*")
        If i > 0 Then Aux = Mid(Aux, 1, i - 1) & "%" & Mid(Aux, i + 1)
    Wend
    'Cambiamos el ? por la _ pue es su omonimo
    i = 1
    While i <> 0
        i = InStr(1, Aux, "?")
        If i > 0 Then Aux = Mid(Aux, 1, i - 1) & "_" & Mid(Aux, i + 1)
    Wend
    cad = Mid(CADENA, 1, 2)
    If cad = "<>" Then
        Aux = Mid(CADENA, 3)
        DevSQL = Campo & " LIKE '!" & Aux & "'"
        Else
        DevSQL = Campo & " LIKE '" & Aux & "'"
    End If
    


    
Case "B"
    'Como vienen de check box o del option box
    'los escribimos nosotros luego siempre sera correcta la
    'sintaxis
    'Los booleanos. Valores buenos son
    'Verdadero , Falso, True, False, = , <>
    'Igual o distinto
    i = InStr(1, CADENA, "<>")
    If i = 0 Then
        'IGUAL A valor
        cad = " = "
        Else
            'Distinto a valor
        cad = " <> "
    End If
    'Verdadero o falso
    i = InStr(1, CADENA, "V")
    If i > 0 Then
            Aux = "True"
            Else
            Aux = "False"
    End If
    'Ponemos la cadena
    DevSQL = Campo & " " & cad & " " & Aux
    
Case Else
    'No hacemos nada
        Exit Function
End Select
SeparaCampoBusqueda = 0
ErrSepara:
    If Err.Number <> 0 Then MuestraError Err.Number
End Function



Private Function CararacteresCorrectos(vCad As String, Tipo As String) As Byte
Dim i As Integer
Dim Ch As String
Dim Error As Boolean

CararacteresCorrectos = 1
Error = False
Select Case Tipo
Case "N"
    'Numero. Aceptamos numeros, >,< = :
    For i = 1 To Len(vCad)
        Ch = Mid(vCad, i, 1)
        Select Case Ch
            Case "0" To "9"
            Case "<", ">", ":", "=", ".", " ", "-"
            Case Else
                Error = True
                Exit For
        End Select
    Next i
Case "T"
    'Texto aceptamos numeros, letras y el interrogante y el asterisco
    For i = 1 To Len(vCad)
        Ch = Mid(vCad, i, 1)
        Select Case Ch
            Case "a" To "z"
            Case "A" To "Z"
            Case "0" To "9"
            Case "*", "%", "?", "_", "\", "/", ":", ".", " " ' estos son para un caracter sol no esta demostrado , "%", "&"
            'Esta es opcional
            Case "<", ">"
            Case "Ñ", "ñ"
            Case Else
                Error = True
                Exit For
        End Select
    Next i
Case "F"
    'Numeros , "/" ,":"
    For i = 1 To Len(vCad)
        Ch = Mid(vCad, i, 1)
        Select Case Ch
            Case "0" To "9"
            Case "<", ">", ":", "/", "="
            Case Else
                Error = True
                Exit For
        End Select
    Next i
Case "B"
    'Numeros , "/" ,":"
    For i = 1 To Len(vCad)
        Ch = Mid(vCad, i, 1)
        Select Case Ch
            Case "0" To "9"
            Case "<", ">", ":", "/", "=", " "
            Case Else
                Error = True
                Exit For
        End Select
    Next i
End Select
'Si no ha habido error cambiamos el retorno
If Not Error Then CararacteresCorrectos = 0
End Function



Public Function DevuelveDesdeBD(kCampo As String, Ktabla As String, Kcodigo As String, ValorCodigo As String, Optional Tipo As String, Optional ByRef OtroCampo As String) As String
    Dim Rs As Recordset
    Dim cad As String
    Dim Aux As String
    
    On Error GoTo EDevuelveDesdeBD
    DevuelveDesdeBD = ""
    cad = "Select " & kCampo
    If OtroCampo <> "" Then cad = cad & ", " & OtroCampo
    cad = cad & " FROM " & Ktabla
    cad = cad & " WHERE " & Kcodigo & " = "
    If Tipo = "" Then Tipo = "N"
    Select Case Tipo
    Case "N"
        'No hacemos nada
        cad = cad & ValorCodigo
    Case "T", "F"
        cad = cad & "'" & ValorCodigo & "'"
    Case Else
        MsgBox "Tipo : " & Tipo & " no definido", vbExclamation
        Exit Function
    End Select
    
    
    
    'Creamos el sql
    Set Rs = New ADODB.Recordset
    Rs.Open cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
        DevuelveDesdeBD = DBLet(Rs.Fields(0))
        If OtroCampo <> "" Then OtroCampo = DBLet(Rs.Fields(1))
    End If
    Rs.Close
    Set Rs = Nothing
    
    
    
    Exit Function
EDevuelveDesdeBD:
        MsgBox Err.Number, "Devuelve DesdeBD.", Err.Description
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
            CampoANulo = "'" & Format(T, FormatoFecha) & "'"
        End If
    Case Else
        If T = "" Then
            CampoANulo = "NULL"
        Else
            CampoANulo = "'" & T & "'"
        End If
    End Select
End Function


Public Function LetDB(T As String, Optional Tipo As String) As Variant

    Select Case Tipo
    Case "N"
        If T = "" Then
            LetDB = 0
        Else
            LetDB = TransformaPuntosComas(T)
        End If
        
    Case "F"
        If T = "" Then
            LetDB = "0:00:00"
        Else
            LetDB = T
        End If
        
        
    Case Else
        LetDB = T
    End Select
End Function

Public Function BorrarTemporal1()
    Conn.Execute "DELETE FROM tmpfich WHERE codusu = " & vUsu.codusu & " AND codequipo= " & vUsu.PC
End Function

Public Function BorrarTemporal2()
    Conn.Execute "DELETE FROM tmpBusqueda WHERE codusu = " & vUsu.codusu & " AND codequipo= " & vUsu.PC
End Function


Public Function InsertaTemporal(Img As Long) As Boolean
Dim cad As String
    On Error GoTo EInsertaTemporal
    InsertaTemporal = False
    
    cad = "INSERT INTO tmpfich (codusu, codequipo, imagen) VALUES ("
    cad = cad & vUsu.codusu & "," & vUsu.PC & "," & Img & ")"
    Conn.Execute cad
    InsertaTemporal = True
    Exit Function
EInsertaTemporal:
    Err.Clear
End Function


Public Function InsertaBusqueda(ByRef Img As Long, ByRef Carpe As Integer)
Dim cad As String
    cad = "INSERT INTO tmpbusqueda (codusu, codequipo, imagen, codcarpeta) VALUES ("
    cad = cad & vUsu.codusu & "," & vUsu.PC & "," & Img & "," & Carpe & ")"
    Conn.Execute cad
End Function

Public Function DevNombreSql(cad As String) As String
Dim i As Integer
Dim J As Integer

    'Buscamos las '
    J = 1
    Do
        i = InStr(J, cad, "'")
        If i > 0 Then
            cad = Mid(cad, 1, i) & "'" & Mid(cad, i + 1)
            J = i + 2
        End If
    Loop Until i = 0

    'Buscamos los \
    J = 1
    Do
        i = InStr(J, cad, "\")
        If i > 0 Then
            cad = Mid(cad, 1, i) & "\" & Mid(cad, i + 1)
            J = i + 2
        End If
    Loop Until i = 0
    DevNombreSql = cad


End Function


Public Function ParaBD(Campo As Variant, Tipo As String, PuedeNulo As Boolean) As String

    Select Case Tipo
    Case "N"
        ParaBD = CStr(Campo)
        If Campo = 0 Then
            If PuedeNulo Then ParaBD = "NULL"
        End If
                
    Case "F"
        ParaBD = "'" & Format(Campo, FormatoFecha) & "'"
        If Val(Campo) = 0 Then
            If PuedeNulo Then ParaBD = "NULL"
        End If
        
    Case Else
        ParaBD = "'" & DevNombreSql(CStr(Campo)) & "'"
        If Campo = "" Then
            If PuedeNulo Then ParaBD = "NULL"
        End If
    End Select
End Function



'Para los nombre que pueden tener ' . Para las comillas habra que hacer dentro otro INSTR
Public Sub NombreSQL(ByRef CADENA As String)
Dim J As Integer
Dim i As Integer
Dim Aux As String

    J = 1
    Do
        i = InStr(J, CADENA, "\")
        If i > 0 Then
            Aux = Mid(CADENA, 1, i - 1) & "\"
            CADENA = Aux & Mid(CADENA, i)
            J = i + 2
        End If
    Loop Until i = 0

    J = 1
    Do
        i = InStr(J, CADENA, "'")
        If i > 0 Then
            Aux = Mid(CADENA, 1, i - 1) & "\"
            CADENA = Aux & Mid(CADENA, i)
            J = i + 2
        End If
    Loop Until i = 0
End Sub



Public Function NombreMSDOS(Nombre As String) As String
Dim K As Integer
Dim C As String
Dim i As Integer

    NombreMSDOS = Nombre
    
    'Le quitamos espacios
    For K = 1 To Len(NoValidosMSDOS)
        C = Mid(NoValidosMSDOS, K, 1)
        Do
            i = InStr(1, NombreMSDOS, C)
            If i > 0 Then NombreMSDOS = Mid(NombreMSDOS, 1, i - 1) & Mid(NombreMSDOS, i + 1)
        Loop Until i = 0
    Next K
End Function
