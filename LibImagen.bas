Attribute VB_Name = "LibImagen"
Option Explicit


'  -- Modos de Trabajo
Public Const vbNorm = 0  ' modo normal
Public Const vbHistNue = 1  ' modo de recuperar historico
Public Const vbHistAnt = 2  ' modo de recuperar historico de los antiguos
Public Const vbBackup = 3   ' REcupoeracion desde un BACKUP

Public Const vbMaxGrupos = 31

Public ModoTrabajo As Byte  '---------------------
Public LlevaHco As Boolean
Public FormatoFecha As String

Public Conn As Connection
Public vUsu As Cusuarios
Public vConfig As CConfiguracion
Public objRevision As HcoRevisiones

Public miRSAux As ADODB.Recordset


Public listacod As Collection
Public listaimpresion As Collection  'Esta lista servira para cuando queramos imprimir

'Cuiado con esta varibale
Public DatosMOdificados As Boolean


'Saber si ha coipado el archivo al server
Public DatosCopiados As String

Public SeHaEjecutadoFTP As Boolean


Public Type RegistroTipoMensaje   ' Crea un tipo definido por el usuario.
   Descripcion As String
   Color As Long
   Icono As Integer
End Type




Public ArrayTipoMen() As RegistroTipoMensaje
Public TotalTipos As Integer   'Menos 1. Es decir, si hay tres tipos la var vale 2




'Usuario As String, Pass As String --> Directamente el usuario

Public Sub Main()
Dim CadenaComandos As String


'    If App.PrevInstance Then
'        MsgBox "Ya se esta ejecutando ARIDOC. Tenga paciencia", vbExclamation, "ARIDOC"
'        End
'        Exit Sub
'    End If
    FormatoFecha = "yyyy-mm-dd"
    
    CadenaComandos = Command
    ModoTrabajo = 0
    SeHaEjecutadoFTP = False
    
    

    
    
    
    'Opcion /s /b     Subir ,bajar un fichero ya estando en aridoc
    
    'CadenaComandos = "/s erdo 5 S c:\m.jpg"
    
    'CadenaComandos = "/f root ""Nueva2"" "
    'CadenaComandos = "/n erdo ""C:\Archivos de programa\Microsoft Visual Studio\Common\Graphics\Bitmaps\Gauge\DOME.bmp"" ""raiz\rama2\obra 10"" "
    'CadenaComandos = "/a"
    'CadenaComandos = Trim(CadenaComandos)
    
    'CadenaComandos = "/u erdo 24 /f1:01/01/2001 /c1:""hola caracola"" /f3:05/06/2005"
    
    'CadenaComandos = "/f root ""2001-2002"" ""raiz"""
    'CadenaComandos = "/N root ""C:\Datos\Aridoc\Raiz\2001-2002\SOCIOS\INFORMES\575.iux"" ""raiz\Raiz\2001-2002\SOCIOS\INFORMES"""
    'CadenaComandos = CadenaComandos & "/F root ""Raiz"" ""raiz"""
        
    'Nueva forma
    'CadenaComandos = "D:\ztesor.txt ""D:\s cart.txt"" d:\mod347.txt"
   ' CadenaComandos = """C:\ariges.txt"""
'    Dim NF As Integer
'    NF = FreeFile
'    Open App.Path & "\F.txt" For Output As #NF
'    Print #NF, CadenaComandos
'    Close #NF
    
    If CadenaComandos = "" Then
        frmInicio.Show
    Else
        OpcionesFlag True
        LanzarShellPedido CadenaComandos
        '   Rc = vUsu.Leer(Val(Cad))
        
        If SeHaEjecutadoFTP Then SubCerrarFTP
        Set Conn = Nothing
        OpcionesFlag False
        End   'FUERZO EL FINAL
    End If
End Sub
Private Sub SubCerrarFTP()
    On Error Resume Next
     frmMovimientoArchivo.Inet1.Cancel
     Err.Clear
End Sub
Private Sub OpcionesFlag(Poner As Boolean)
Dim N As String
Dim NF As Integer
    
    N = App.Path & "\flag.txt"
    If Poner Then
        If Dir(N, vbArchive) = "" Then
            NF = FreeFile
            Open N For Output As #NF
            Print #NF, Now
            Close #NF
        End If
    Else
        'Eliminar el flag
        If Dir(N, vbArchive) <> "" Then Kill N
    End If
End Sub


'Realmente este es para insertar
Public Sub GestionarEquipo()
Dim cad As String

    

        Set miRSAux = New ADODB.Recordset
        cad = "Select max(codequipo) from equipos"
        miRSAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        vUsu.PC = 1
        If Not miRSAux.EOF Then vUsu.PC = DBLet(miRSAux.Fields(0), "N") + 1
        miRSAux.Close
        Set miRSAux = Nothing
        
        'De momento un insert normal
        cad = "INSERT INTO equipos (codequipo, descripcion,  cargaIconsExt) VALUES (" & vUsu.PC
        cad = cad & ",'" & vUsu.NomPC & "',1)"
        Conn.Execute cad
        

        
End Sub

Public Function ComprobacionesPrevias() As Boolean
Dim MiNombre As String
Dim MiRuta As String

    On Error GoTo EComprobacionesPrevias

    'Tiene k existir la carpeta Imagenes
    'imagenes
    If Dir(App.Path & "\imagenes", vbDirectory) = "" Then MkDir App.Path & "\imagenes"
    'Tambien la temporal
    If Dir(App.Path & "\temp", vbDirectory) = "" Then MkDir App.Path & "\temp"
            
            
    'Tiene k estar vacia
    'Como algunos les habremos cambiados la extension
    'La volvemos a poner a lecturaescritura
    MiRuta = App.Path & "\temp\"
    MiNombre = Dir(MiRuta, vbDirectory)    ' Recupera la primera entrada.
    Do While MiNombre <> ""   ' Inicia el bucle.
       ' Ignora el directorio actual y el que lo abarca.
       If MiNombre <> "." And MiNombre <> ".." Then
          ' Realiza una comparación a nivel de bit para asegurarse de que MiNombre es un directorio.
          If (GetAttr(MiRuta & MiNombre) And vbDirectory) = vbDirectory Then
            
          Else
              SetAttr MiRuta & MiNombre, vbNormal
              Kill MiRuta & MiNombre
          End If   ' solamente si representa un directorio.
       End If
       MiNombre = Dir   ' Obtiene siguiente entrada.
    Loop

    
    
    
    ComprobacionesPrevias = True
    
    Exit Function
EComprobacionesPrevias:
    MiRuta = "Comprobaciones Previas" & vbCrLf & Err.Description & vbCrLf & MiRuta & MiNombre
    MiRuta = MiRuta & vbCrLf & vbCrLf & "No se han podido borrar todos los archivos. ¿Desea continuar de igual modo ?"
    If MsgBox(MiRuta, vbCritical + vbYesNoCancel) <> vbYes Then
        ComprobacionesPrevias = False
    Else
         ComprobacionesPrevias = True
    End If
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




Public Sub Mensajes1(Valor As Integer)
Dim m As String
Select Case Valor

Case 0  'Desde hacer copiar
        m = "No se puede realizar copias desde Histórico."
Case 1 'Crear diretorio
        m = "No se puede eliminar la carpeta."
Case 2 'Eliminar archivos
        m = "No se pueden eliminar archivos."
Case 3, 4
        m = "No se puede Corta&Pegar."
Case 5 'insertar
        m = "No se puede insertar nuevos archivos."
Case 6
        m = "No se puede realizar MOVER."
Case 7
        m = "No se puede importar nuevos archivos."
Case 8
        m = "No se pueden crear carpetas nuevas."
Case 9
        m = " No hay ningún archivo seleccionado para imprimir."
Case 10
        m = " No se puede verificar los documentos de la gestión documental desde el hco"
Case 11
        m = " No se puede crear CARPETAS desde historico"
Case 12
        m = " No se puede verificar CARPETAS desde historico"
Case 13
        m = " No se puede modificar archivos desde historico"
Case 14
        m = " No se puede verificar desde historico"
Case 15
        m = "Imposible realizar estos cambios en el historico"
Case 16
        m = "Opcion no disponible en modo historico"
End Select

MsgBox m, vbInformation
End Sub

'Para obtener el nompath de un archivo a partir del
'treeview1.fullpath
Public Function DamePath(CADENA As String)
'Dim aux
'Dim l As Integer
'Dim I As Integer
'
'l = Len(mConfig.Carpeta) + 2
'I = InStr(1, cadena, mConfig.Carpeta)
'If I > 0 Then
'    aux = Mid(cadena, l)
'    Else
'        MsgBox "Error calculando PATH relativo", vbExclamation
'        aux = cadena
'End If
'DamePath = aux
End Function


Public Function devuelvePATH(NomPath As String) As String
Dim i
Dim CADENA
Dim Cad2 As String
Cad2 = NomPath
Do
    i = InStr(1, Cad2, "/")
    If i > 0 Then
        CADENA = Mid(Cad2, 1, i - 1)
        Cad2 = CADENA & "\" & Mid(Cad2, i + 1)
    End If
Loop Until i = 0
devuelvePATH = Cad2
End Function


Public Sub MostrarError(NumeroError As Long, Optional Texto As String)
Dim cad
cad = "Se ha producido un error." & vbCrLf
If Texto <> "" Then cad = cad & Texto & vbCrLf
cad = cad & "Número: " & NumeroError & vbCrLf
cad = cad & "Descripción: " & Error(NumeroError) & vbCrLf
MsgBox cad, vbExclamation
End Sub





''-----------------------------------------------
''-----------------------------------------------
''-----------------------------------------------
''-----------------------------------------------
''Estas funciones estaban antes en admin, y ahora las hemos sacado
'Public Function CompruebaCarpeta(ByVal kCarpeta As String, men As String) As Byte
'' 1 ---> No tiene archivos
'' 2 ---> Si tiene archivos
'' 3 ---> Tiene SUB directorios
'Dim Aux As String
'Dim miNombre As String
'Dim TieneDir As Boolean
'Dim TieneArch As Boolean
'
'
'TieneDir = False
'TieneArch = False
'Aux = kCarpeta
'miNombre = Dir(Aux, vbDirectory)
'Do While miNombre <> ""
'   If miNombre <> "." And miNombre <> ".." Then
'        If (GetAttr(Aux & miNombre) And vbDirectory) = vbDirectory Then
'            TieneDir = True
'            Exit Do
'            Else
'                TieneArch = True
'                Exit Do
'            End If
'     End If
'   miNombre = Dir ' Obtiene siguiente entrada.
'   Loop
'
'
'If TieneDir Then
'   men = "La carpeta contiene Subcarpetas"
'   CompruebaCarpeta = 3
'   Else
'        If TieneArch Then
'            men = "La carpeta contiene archivos"
'            CompruebaCarpeta = 2
'        Else
'            men = "NO tiene"
'            CompruebaCarpeta = 1
'        End If
'    End If
'End Function




Public Function TratarCarpeta(vCarpeta As String) As Byte
'Dim valor
'Dim directorio As String
'Dim subcarpeta As String
'Dim Fin As Boolean
'Dim ruta As String
'Dim Camino   ' aqui tendremos la ruta de la carpeta que la contiene
'Dim st1 As String ' Para poder llamar a la funcion carpeta
'
'
'TratarCarpeta = 0 ' 0 correcto    1 .- error
'directorio = vCarpeta
'Fin = False
'subcarpeta = ""
'ruta = inicial & Carpeta & "\"
'Camino = ruta
'While Not Fin
'    valor = InStr(1, directorio, "\")
'    If valor = 0 Then
'        Fin = True
'        subcarpeta = directorio
'        Else
'            subcarpeta = Mid(directorio, 1, valor - 1)
'        End If
'
'    ruta = ruta & subcarpeta & "\"
'    If Dir(ruta, vbDirectory) = "" Then ' la carpeta no existe
'        If CompruebaCarpeta(Camino, st1) <> 2 Then
'            MkDir (ruta)
'            SeHanCreadoCarpetas = True
'            Else
'                TratarCarpeta = 1
'                Exit Function
'            End If
'    End If
'    Camino = Camino & subcarpeta & "\"
'    directorio = Mid(directorio, Len(subcarpeta) + 2, Len(directorio))
'Wend
End Function


Private Function CopiaArchivo(NA As String, Car_Des As String) As Byte
'On Error GoTo ErrorCopiaArchivo
'CopiaArchivo = 1
'    FileCopy mConfig.carpetaInt & "\" & NA, Car_Des & "\" & NA
'    Kill mConfig.carpetaInt & "\" & NA
'CopiaArchivo = 0
'Exit Function
'ErrorCopiaArchivo:
'    MsgBox "Se ha producido un error copiando archivo: " & vbCrLf & _
'        "    .-" & mConfig.carpetaInt & "\" & NA & vbCrLf & _
'        "Número: " & Err.Number & vbCrLf & _
'        "Descripción: " & Err.Description, vbExclamation
End Function

'-----------------------------------------------
'-----------------------------------------------
'-----------------------------------------------
'-----------------------------------------------




Public Function ProcesaLinea2(ByRef L As String) As String
Dim i, C, l2
Dim J As Byte
l2 = ""
'Para que no tenga que hacer cada vez el select, y sabiendo que casi todo son letras y numero
'Para saber si lo tenemos que modificar
'comprobaremos que el ASC es mayor 165 para saber si hay que hacer cambios, o no
'If InStr(1, l, "CAMPA") Then Stop
For i = 1 To Len(L)
    C = Mid(L, i, 1)
    J = Asc(C)
    If J > 125 Then
        'Caracteres especiales
        Select Case J
        Case 165
            C = "Ñ"
        Case 166
            C = "ª"
        Case 167
            C = "/"
        Case 179
            C = "|"
        Case 191, 192, 193, 194, 196
            C = "-"
        Case 217, 218 ' Estas son las esquinas
            C = "-"
        End Select
    End If
    l2 = l2 & C
Next i
ProcesaLinea2 = l2
End Function

Public Function espera(Segundos As Single)
    Dim T1
    T1 = Timer
    Do
    Loop Until Timer - T1 > Segundos
End Function





'DIAS:  0.-  Dentro del mes
'       1.-  Hace mas de un mes
'       2.-  hace mas de dos meses
'       3.- NUNCA se ha hecho o fichero no existe
Public Sub FicheroVerificacion(Grabar As Boolean, ByRef Dias As Byte, Optional ByRef LosErrores As String)
Dim NF As Integer
Dim d As Long

On Error GoTo EF

    NF = FreeFile
    If Grabar Then
        Open App.Path & "\carpeta.dat" For Output As #NF
        Print #NF, Format(Now, "dd/mm/yyyy")
        Print #NF, "Hora: " & Format(Now, "hh:mm")
        Print #NF, "Revisión carpetas. "
        If LosErrores = "" Then
            Print #NF, "Todo bien"
        Else
            Print #NF, LosErrores
        End If
        Close NF
    Else
        If Dir(App.Path & "\carpeta.dat") = "" Then
            Dias = 3   'Fichero no existe
        Else
            Open App.Path & "\carpeta.dat" For Input As #NF
            Line Input #NF, LosErrores
            If IsDate(LosErrores) Then
                'FECHA
                d = DateDiff("d", Now, CDate(LosErrores))
                
                If Abs(d) > 62 Then
                    Dias = 2
                Else
                    Dias = 1
                End If
            Else
                Dias = 3 'No podremos precisar
            End If
            Close NF

        End If
    End If
    Exit Sub
EF:
    MsgBox Err.Description
End Sub






Public Sub MuestraError(Numero As Long, Optional CADENA As String, Optional Desc As String)
    Dim cad As String
   
    'Con este sub pretendemos unificar el msgbox para todos los errores
    'que se produzcan
    On Error Resume Next
    cad = "Se ha producido un error: " & vbCrLf
    If CADENA <> "" Then
        cad = cad & vbCrLf & CADENA & vbCrLf & vbCrLf
    End If

    If Desc <> "" Then cad = cad & vbCrLf & Desc & vbCrLf & vbCrLf
    cad = cad & "Número: " & Numero & vbCrLf & "Descripción: " & Error(Numero)
    MsgBox cad, vbExclamation
End Sub




'///////////////////////////////////////////////////////////////
'
'   Cogemos un numero formateado: 1.256.256,98  y deevolvemos 1256256.98
'   Tiene que venir numérico
Public Function ImporteFormateado(Importe As String) As Currency
Dim i As Integer

If Importe = "" Then
    ImporteFormateado = 0
    Else
        'Primero quitamos los puntos
        Do
            i = InStr(1, Importe, ".")
            If i > 0 Then Importe = Mid(Importe, 1, i - 1) & Mid(Importe, i + 1)
        Loop Until i = 0
        ImporteFormateado = Importe
End If
End Function
'Cambia los puntos de los numeros decimales
'por comas
Public Function TransformaComasPuntos(CADENA As String) As String
    Dim i As Integer
    Do
        i = InStr(1, CADENA, ",")
        If i > 0 Then
            CADENA = Mid(CADENA, 1, i - 1) & "." & Mid(CADENA, i + 1)
        End If
        Loop Until i = 0
    TransformaComasPuntos = CADENA
End Function



Public Function TransformaPuntosComas(CADENA As String) As String
    Dim i As Integer
    Do
        i = InStr(1, CADENA, ".")
        If i > 0 Then
            CADENA = Mid(CADENA, 1, i - 1) & "," & Mid(CADENA, i + 1)
        End If
        Loop Until i = 0
    TransformaPuntosComas = CADENA
End Function



'Cambia los puntos de los numeros decimales
'por comas
Public Function TransformaPuntosHoras(CADENA As String) As String
    Dim i As Integer
    Do
        i = InStr(1, CADENA, ".")
        If i > 0 Then
            CADENA = Mid(CADENA, 1, i - 1) & ":" & Mid(CADENA, i + 1)
        End If
    Loop Until i = 0
    TransformaPuntosHoras = CADENA
End Function



Public Function EsFechaOK(ByRef T As TextBox) As Boolean
Dim cad As String

    cad = T.Text
    If InStr(1, cad, "/") = 0 Then
        If Len(T.Text) = 8 Then
            cad = Mid(cad, 1, 2) & "/" & Mid(cad, 3, 2) & "/" & Mid(cad, 5)
        Else
            If Len(T.Text) = 6 Then cad = Mid(cad, 1, 2) & "/" & Mid(cad, 3, 2) & "/" & Mid(cad, 5)
        End If
    End If

    If IsDate(cad) Then
        EsFechaOK = True
        T.Text = Format(cad, "dd/mm/yyyy")
    Else
        EsFechaOK = False
    End If
   
End Function



Public Function EsFechaOKString(ByRef T As String) As Boolean
Dim cad As String
    
    cad = T
    If InStr(1, cad, "/") = 0 Then
        If Len(T) = 8 Then
            cad = Mid(cad, 1, 2) & "/" & Mid(cad, 3, 2) & "/" & Mid(cad, 5)
        Else
            If Len(T) = 6 Then cad = Mid(cad, 1, 2) & "/" & Mid(cad, 3, 2) & "/" & Mid(cad, 5)
        End If
    End If
    If IsDate(cad) Then
        EsFechaOKString = True
        T = Format(cad, "dd/mm/yyyy")
    Else
        EsFechaOKString = False
    End If
End Function




Public Function DevuelveNombreFichero(campo1 As String, Extension As String, ByRef NombreFinalFichero As String, ParaEmail As Boolean) As Integer
Dim i As Integer
Dim cad As String

    On Error GoTo ED

    Do
        i = InStr(1, campo1, " ")
        If i > 0 Then campo1 = Mid(campo1, 1, i - 1) & Mid(campo1, i + 1)
    Loop Until i = 0
    
    'QUito comas tambien
    Do
        i = InStr(1, campo1, ",")
        If i > 0 Then campo1 = Mid(campo1, 1, i - 1) & "_" & Mid(campo1, i + 1)
    Loop Until i = 0
    
    'quito los dospuntos : por _

    Do
        i = InStr(1, campo1, ":")
        If i > 0 Then campo1 = Mid(campo1, 1, i - 1) & "_" & Mid(campo1, i + 1)
    Loop Until i = 0

    'QUito las barras
    Do
        i = InStr(1, campo1, "/")
        If i > 0 Then campo1 = Mid(campo1, 1, i - 1) & "_" & Mid(campo1, i + 1)
    Loop Until i = 0




    i = 0
    Do
        If ParaEmail Then
            cad = App.Path & "\mail\" & campo1
        Else
            cad = App.Path & "\temp\" & campo1
        End If
        If i > 0 Then cad = cad & "(" & i & ")"
        cad = cad & "." & Extension
        i = i + 1
    Loop Until Dir(cad, vbArchive) = "" Or i > 100
    NombreFinalFichero = cad
    Exit Function
ED:
    MuestraError Err.Number, "Devuelve nombre fichero: " & Err.Description
    DevuelveNombreFichero = 101
End Function

'Public Function TraerFicheroFisico(ByRef Carpeta As Ccarpetas, Destino As String, codigo As Long) As Boolean
Public Function TraerFicheroFisico(ByRef Carpeta As Ccarpetas, Destino As String, codigo) As Boolean


        TraerFicheroFisico = False
        'Llevamos el fichero
        DatosCopiados = "NO"
        Set frmMovimientoArchivo.vOrigen = Carpeta
        frmMovimientoArchivo.Opcion = 2
        frmMovimientoArchivo.Origen = codigo
        frmMovimientoArchivo.Destino = Destino
        frmMovimientoArchivo.Show vbModal
        
        'Y si se producen errores No abrimos
        If DatosCopiados = "" Then TraerFicheroFisico = True
        
End Function

Public Function TextoParaComonDialog2(SoloNuevo As Boolean, Optional CuantasHay As Integer) As String
Dim SQL As String
Dim i As Integer

    TextoParaComonDialog2 = ""
    SQL = " SELECT extensionpc.*,descripcion,exten from extensionpc,extension where "
    SQL = SQL & " extensionpc.codext = extension.codext AND codequipo=" & vUsu.PC
    If SoloNuevo Then SQL = SQL & " AND extension.nuevo=1"
    'Que este habilitada
    SQL = SQL & " AND extension.Deshabilitada =0"
    SQL = SQL & " order by descripcion" '     'Visor<>""Predeterminado"""
    Set miRSAux = New ADODB.Recordset
    miRSAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""
    i = 0
    If Not miRSAux.EOF Then

        
        While Not miRSAux.EOF
            SQL = SQL & "|" & miRSAux!Descripcion & "   (*." & miRSAux!Exten & ")|*." & miRSAux!Exten
            miRSAux.MoveNext
            i = i + 1
        Wend
        SQL = Mid(SQL, 2) 'Quito el primer |
    End If
    miRSAux.Close
    Set miRSAux = Nothing
    CuantasHay = i
    TextoParaComonDialog2 = SQL
End Function




Public Sub PonerArrayTiposMensaje()
Dim L As Long
Dim Fin As Integer
Dim i As Integer
Dim J As Integer
Dim Cortar11 As String
'Public Type RegistroTipoMensaje   ' Crea un tipo definido por el usuario.
'   Descripcion As String * 30
'   Color As Long
'End Type
'
'Public ArrayTipoMen() As RegistroTipoMensaje
    TotalTipos = 0
    Cortar11 = "Select count(*) from mailtipo"
    Set miRSAux = New ADODB.Recordset
    miRSAux.Open Cortar11, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Fin = 0
    If Not miRSAux.EOF Then Fin = DBLet(miRSAux.Fields(0), "N")
    miRSAux.Close
    
    If Fin = 0 Then Exit Sub
    
    
    ReDim ArrayTipoMen(Fin)
    TotalTipos = Fin
    Cortar11 = "Select * from mailtipo order by tipo "
    miRSAux.Open Cortar11, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    J = 0
    i = 0
    
    
    While Not miRSAux.EOF
        
        If miRSAux!Tipo - J > 1 Then
            J = J + 1
            For Fin = J To miRSAux!Tipo - 1
                ArrayTipoMen(Fin).Color = 0
                ArrayTipoMen(Fin).Descripcion = ""
                ArrayTipoMen(Fin).Icono = 0
            Next Fin
            i = miRSAux!Tipo
        End If
        
        ArrayTipoMen(i).Color = DBLet(miRSAux!Color, "N")
        ArrayTipoMen(i).Descripcion = miRSAux!Descripcion
        ArrayTipoMen(i).Icono = miRSAux!numico
        J = miRSAux!Tipo
        
        miRSAux.MoveNext
        i = i + 1
    Wend
    miRSAux.Close
    Set miRSAux = Nothing

End Sub


Public Sub CodificacionLinea(Leer As Boolean, ByRef Linea As String)
Dim i As Integer
Dim C As String
Dim C2 As String
    C = Linea
    Linea = ""
    
        
        'Escribir
        For i = 1 To Len(C)
            C2 = Mid(C, i, 1)
            If Leer Then
                C2 = Chr(Asc(C2) - 3)
            Else
                C2 = Chr(Asc(C2) + 3)
            End If
            Linea = Linea & C2
        Next i
End Sub


Public Sub AsignarCampoMemo(ByRef Campo As String, ByRef nombrecampo As String, ByRef ADO As ADODB.Recordset)
    On Error Resume Next
    Campo = ADO.Fields(nombrecampo).Value
    If Err.Number <> 0 Then
        Err.Clear
        Campo = ""
    End If
End Sub



Public Function TienePlantillasEnCarpeta2() As Boolean
Dim RT As ADODB.Recordset
    On Error GoTo ETienePlantillasEnCarpeta
    TienePlantillasEnCarpeta2 = False
    
    Set RT = New ADODB.Recordset
    RT.Open "Select * from plantillacarpetas", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    RT.Close
    TienePlantillasEnCarpeta2 = True
    
ETienePlantillasEnCarpeta:
    If Err.Number <> 0 Then Err.Clear
    Set RT = Nothing
End Function




Public Function ComprobarCarpetaBackup(Comprobar As Boolean) As Boolean
Dim cad As String
    On Error GoTo EComprobarCarpetaBackup
    ComprobarCarpetaBackup = False

    If Comprobar Then
        If Dir(App.Path & "\tmpB", vbDirectory) = "" Then
            MkDir App.Path & "\tmpB"
    
        Else
            
            cad = Dir(App.Path & "\tmpB\*.*")
            Do While cad <> ""
                Kill App.Path & "\tmpB\" & cad
                cad = Dir
            Loop
        End If
    
    Else
        'Para eliminarlo toooo
        If Dir(App.Path & "\tmpB", vbDirectory) <> "" Then
            
            
            cad = Dir(App.Path & "\tmpB\*.*")
            Do
                If cad <> "" Then Kill App.Path & "\tmpB\" & cad
                cad = Dir
            Loop Until cad = ""
        
    
            If Dir(App.Path & "\tmpB", vbDirectory) <> "" Then Kill App.Path & "\tmpB"
        End If
    End If
    
    ComprobarCarpetaBackup = True
    
    Exit Function
EComprobarCarpetaBackup:
    MuestraError Err.Number, "Comprobar Carpeta Backup"
End Function





'-------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------
'
'       Impresion de varios archivos desde una tabla temporal
'
Public Sub ImprimirDesdeTablaTemporal(ByRef ElFormulario As Form, EstaEnHco As Boolean)
Dim vE As Cextensionpc
Dim i As Byte
Dim NombreArchivo As String
Dim J As Integer
Dim Rs As ADODB.Recordset
Dim Carpe As Ccarpetas
Dim T1 As Single
            

            On Error GoTo EImprimirDesdeTablaTemporal
            
            NombreArchivo = "select extensionpc.codext,extensionpc.impresion,extension.descripcion from tmpfich,timagen"
            If EstaEnHco Then NombreArchivo = NombreArchivo & "hco as timagen "
            NombreArchivo = NombreArchivo & " ,extensionpc,extension where"
            NombreArchivo = NombreArchivo & " tmpfich.imagen = timagen.codigo and timagen.codext=extensionpc.codext and"
            NombreArchivo = NombreArchivo & " Extension.codext = timagen.codext and timagen.codext=extensionpc.codext"
            NombreArchivo = NombreArchivo & " and  extensionpc.codequipo=" & vUsu.PC & " and codusu =" & vUsu.codusu
            NombreArchivo = NombreArchivo & " and tmpfich.codequipo=" & vUsu.PC & " group by 1"
            Set Rs = New ADODB.Recordset
            Rs.Open NombreArchivo, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            J = 0
            NombreArchivo = ""
            While Not Rs.EOF
                'if extensionpc.impresion
                If DBLet(Rs.Fields(1), "T") = "" Then
                    NombreArchivo = NombreArchivo & "    - " & Rs.Fields(2) & " (" & Rs.Fields(0) & ")" & vbCrLf
                    J = J + 1
                End If
                Rs.MoveNext
            Wend
            Rs.Close
            
            If J > 0 Then
                Set miRSAux = Nothing
                If J > 1 Then
                    NombreArchivo = "Los tipos de archivo: " & vbCrLf & vbCrLf & NombreArchivo
                Else
                    NombreArchivo = "El tipo de archivo: " & vbCrLf & vbCrLf & NombreArchivo
                End If
                
                NombreArchivo = NombreArchivo & vbCrLf & vbCrLf & "No tienen opcion de impresión para este equipo"
                MsgBox NombreArchivo, vbExclamation
                Exit Sub
            End If
            
            
            NombreArchivo = "select timagen.codigo,timagen.codext,timagen.campo1,timagen.codcarpeta from tmpfich,timagen"
            If EstaEnHco Then NombreArchivo = NombreArchivo & "hco as timagen"
            NombreArchivo = NombreArchivo & " ,extensionpc,extension where"
            NombreArchivo = NombreArchivo & " tmpfich.imagen = timagen.codigo and timagen.codext=extensionpc.codext"
            NombreArchivo = NombreArchivo & " and extension.codext=extensionpc.codext and  extensionpc.codequipo=" & vUsu.PC
            NombreArchivo = NombreArchivo & " and codusu =" & vUsu.codusu & " and tmpfich.codequipo=" & vUsu.PC & " order by 2"
            
            
            Set vE = New Cextensionpc
            Set Carpe = New Ccarpetas
            Carpe.codcarpeta = -1
            vE.codext = -1
            Rs.Open NombreArchivo, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not Rs.EOF
                Screen.MousePointer = vbHourglass
                If vE.codext <> Rs.Fields(1) Then
                    i = 0
                    If vE.Leer(Rs.Fields(1), vUsu.PC) = 1 Then
                        i = 1
                    Else
                        If vE.impresion = "" Then
                            i = 1
                            MsgBox "La extension no tiene PATH asociado.", vbExclamation
                        End If
                    End If
                        
                    If i = 1 Then
                        Set vE = Nothing
                        Rs.Close
                        Set Rs = Nothing
                        Exit Sub
                    End If
                End If   'De leeer la extension
            
                
                If Carpe.codcarpeta <> Rs.Fields(3) Then
                    If Carpe.Leer(Rs.Fields(3), EstaEnHco) = 1 Then
                        Set Carpe = Nothing
                        Set vE = Nothing
                        Rs.Close
                        Set Rs = Nothing
                        Exit Sub
                    End If
                End If
                
                
                    i = DevuelveNombreFichero(Rs.Fields(2), vE.Extension, NombreArchivo, False)
                    If i > 100 Then
                        MsgBox "Error obteniendo nombre fichero", vbExclamation
                        Rs.Close
                        Set Rs = Nothing
                        Exit Sub
                    End If
                
                T1 = Timer
                Imprimir1Fichero Carpe, Rs.Fields(0), NombreArchivo, vE
                Do
                    DoEvents
                    ElFormulario.Refresh
                    espera 0.25
                Loop Until Timer - T1 > 1.5
                
                Rs.MoveNext
                
                
            Wend
            Rs.Close
            Set Rs = Nothing
            Set vE = Nothing
            espera 1
            If vConfig.RevisaTareasAPI Then VerProcesosMuertos
            
            
EImprimirDesdeTablaTemporal:
    If Err.Number <> 0 Then MuestraError Err.Number, "Imprimir Desde Temporal" & vbCrLf & Err.Description
    Set Carpe = Nothing
    Set vE = Nothing
    Set Rs = Nothing
    Screen.MousePointer = vbDefault
End Sub





Public Sub Imprimir1Fichero(ByRef Ca As Ccarpetas, Cod As Long, Destino As String, ByRef CEx As Cextensionpc)
On Error GoTo EA
Dim cad As String
Dim TamanyoOriginal As Long


Dim FS, F   'File system


        If Not TraerFicheroFisico(Ca, Destino, Cod) Then Exit Sub
       
        If Dir(Destino, vbArchive) = "" Then
            MsgBox "Se ha producido un error trayendo los datos", vbExclamation
            Exit Sub
        End If
        
        Set FS = CreateObject("Scripting.FileSystemObject")
        Set F = FS.GetFile(Destino)
        
        
        
        'Protegemos para escritura
        SetAttr Destino, vbReadOnly
                
        '----------------------------------------

        cad = CEx.impresion & " """ & F.shortpath & """"

        TamanyoOriginal = Shell(cad, vbNormalFocus)
        InsertarEnProcesosAbiertos TamanyoOriginal, Destino
        espera 1
        
        If vConfig.RevisaTareasAPI Then VerProcesosMuertos
    
EA:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Abrir/Modificar fichero."
    End If
    Set FS = Nothing
    Set F = Nothing
    Screen.MousePointer = vbDefault
End Sub






Public Sub VerProcesosMuertos()
Dim cad As String
Dim C2 As String
    'Cad = LoadTaskList(Me.Hwnd)
    cad = ExistePId
    
    
    'Ahora aqui vere los procesos lanzados
    Set miRSAux = New ADODB.Recordset
    C2 = "Select * from procesos where codusu=" & vUsu.codusu & " AND codequipo =" & vUsu.PC
    miRSAux.Open C2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRSAux.EOF
        C2 = "·" & miRSAux!Proceso & "·"
        If InStr(1, cad, C2) = 0 Then
            'Eliminamos referencia
            EliminiarReferenciaProcesoArchivo miRSAux!fichero, CLng(miRSAux!Proceso)
        End If
        miRSAux.MoveNext
    Wend
    miRSAux.Close
    Set miRSAux = Nothing
End Sub


Private Sub EliminiarReferenciaProcesoArchivo(Arch As String, Proceso As Long)

    On Error GoTo EEliminiarReferenciaProcesoArchivo
    
    SetAttr Arch, vbNormal
    Kill Arch
    
    Conn.Execute "DELETE FROM PROCESOS WHERE codusu = " & vUsu.codusu & " AND codequipo =" & vUsu.PC & " AND proceso =" & Proceso
    

    
    Exit Sub
EEliminiarReferenciaProcesoArchivo:
    'MuestraError Err.Number, Err.Description
    Err.Clear
End Sub




Public Sub InsertarEnProcesosAbiertos(Referencia As Long, ByRef Fich As String)
Dim cad As String

    cad = "INSERT INTO Procesos (codusu, codequipo, proceso,fichero) VALUES ("
    cad = cad & vUsu.codusu & "," & vUsu.PC & "," & Referencia & ",'" & DevNombreSql(Fich) & "')"
    On Error Resume Next
    Conn.Execute cad
    If Err.Number <> 0 Then Err.Clear
   
End Sub



Public Function ExtensionNFI(ByRef R As ADODB.Recordset) As Integer


    Set R = New ADODB.Recordset
    ExtensionNFI = -1
    R.Open "Select codext,exten from extension", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not R.EOF
        If UCase(CStr(R.Fields(1))) = "NFI" Then
            ExtensionNFI = R.Fields(0)
            R.Close
            Exit Function
        End If
        R.MoveNext
    Wend
    R.Close
End Function



