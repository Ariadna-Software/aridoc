Attribute VB_Name = "LibSHELL"
Option Explicit



'-----------------------------------------------------------------
' OPCIONES
'
'
'       bajar un archivo desde aridoc
'       /b usuario  codigoaridoc  "path donde descargar"
'
'
'       Subir un archivo que ya estaba en aridoc
'       /s usuario codigoaridoc "path donde cojer arciov"
'
'
'       /ayuda
'
'
'       nueva estructura
'       /e usuario pathficherointercambio
'
'
'       copiar estructura
'       /c usuario  carpetaorigen  carpetadestino
'EN MAYUSCULAS la opcion no muestra mensaje de error

Private LaOpcion As Byte
Private RestoParametros As String
Private MiUsuario As String
Private MostrarVentanaErrores As Boolean

Private Sub MostrarVentanaAyuda()
    MiUsuario = ""
    For LaOpcion = 1 To 60
        MiUsuario = MiUsuario & "*"
    Next LaOpcion
    MiUsuario = MiUsuario & vbCrLf
    
    
    RestoParametros = "Archivo ejecucion (flag): " & App.Path & "\flag.txt" & vbCrLf & vbCrLf
    RestoParametros = RestoParametros & "SUBIR: " & vbCrLf
    RestoParametros = RestoParametros & "       /s usuario idaridoc bloquear pathlocal" & vbCrLf & vbCrLf

    RestoParametros = RestoParametros & "BAJAR: " & vbCrLf
    RestoParametros = RestoParametros & "       /b usuario idaridoc letra pathlocal" & vbCrLf & vbCrLf

    RestoParametros = RestoParametros & "NUEVO ARCHIVO: " & vbCrLf
    RestoParametros = RestoParametros & "       /n usuario pathlocal patharidoc" & vbCrLf
    RestoParametros = RestoParametros & "       Intercambio: " & App.Path & "\InterShell.txt" & vbCrLf & vbCrLf & vbCrLf
    
    RestoParametros = RestoParametros & "ACTUALIZAR CLAVES: " & vbCrLf
    RestoParametros = RestoParametros & "       /u usuario idaridoc /c1: /c2: ......" & vbCrLf
    RestoParametros = RestoParametros & "       Donde: /c[1..4]: campo texto   /f[1..3]: fecha   /i[1,2]: importe   /o: observacion " & vbCrLf & vbCrLf

    RestoParametros = RestoParametros & "CREAR CARPETA: " & vbCrLf
    RestoParametros = RestoParametros & "-       /f usuario nombrecarpeta carpetadestino" & vbCrLf & vbCrLf
    
    RestoParametros = RestoParametros & vbCrLf & " ERRORES: " & App.Path & "\ErrorShell.txt" & vbCrLf & vbCrLf
    
    RestoParametros = RestoParametros & vbCrLf & "               www.ariadnasoftware.com" & vbCrLf & vbCrLf
    MiUsuario = MiUsuario & RestoParametros & MiUsuario
    MsgBox MiUsuario, vbExclamation
    
End Sub



Public Sub LanzarShellPedido(vcmd As String)
Dim i As Integer
Dim L As Long

    'Vemos los argumentos para saber si se ha llamdo bien el programa
    RestoParametros = "Error leyendo parametros"
    
    If Len(vcmd) = 1 Then
        
        'Seguro k se ha llamado mal al programa
        ParametrosEquivocados2 1, RestoParametros
        Exit Sub
    End If
    
    'Vemos k lo primero es la /
    ' SI NO ES ASI, entonces vemos si son archivos
    
    If Mid(vcmd, 1, 1) <> "/" Then
        'NO TIENE "/"
        Screen.MousePointer = vbHourglass
        If SonArchivosParaInsertar(vcmd) Then
        
            Exit Sub
        
            'TB NOS SALIMOS
        Else
            ParametrosEquivocados2 2, RestoParametros
            Exit Sub
        End If
    End If


    'Vemos k lo segundo es una letra valida
    DatosCopiados = LCase(Mid(vcmd, 2, 1))
    DatosMOdificados = False
    
    
    MostrarVentanaErrores = (DatosCopiados = UCase(DatosCopiados))
    Select Case DatosCopiados
    Case "d", "s"
        If DatosCopiados = "d" Then
            LaOpcion = 1
        Else
            LaOpcion = 2
        End If
        DatosMOdificados = True
    Case "a"
        'Ayuda
        MostrarVentanaAyuda
        Exit Sub
    Case "f"
        LaOpcion = 3
        DatosMOdificados = True
        
    Case "n"

        LaOpcion = 4
        DatosMOdificados = True
        
    Case "u"
        'actualizar claves
        LaOpcion = 5
        DatosMOdificados = True
    End Select
    
    If Not DatosMOdificados Then
        ParametrosEquivocados2 3, RestoParametros
        Exit Sub
    End If
    
    'Vemos usuario
    i = InStr(4, vcmd, " ")
    If i = 0 Then
        ParametrosEquivocados2 4, RestoParametros
        Exit Sub
    End If
    
    
    MiUsuario = Mid(vcmd, 4, i - 4)
    vcmd = Trim(Mid(vcmd, i))
    
    
    
    'Segun sea la opcion enonces comprobaremos unas cosas u otras
    Select Case LaOpcion
    Case 1, 2
    
        'Esta pidiendo subir, bajar un fichero
        'Entonces compruebo varias cosas
        'Errores: 5,6,7,8
        RestoParametros = vcmd
        vcmd = "Valores adicionales"
        i = ComprobarSubirBajarFichero(RestoParametros)
        If i <> 0 Then
            ParametrosEquivocados2 i, vcmd
            Exit Sub
        End If
    Case 3, 4
        'Primero compruebao que el fichero de intercambio no esta creado
        If Dir(App.Path & "\InterShell.txt", vbArchive) <> "" Then
            ParametrosEquivocados2 23, "Fichero intercambio no ha sido eliminado"
            Exit Sub
        End If
        
    
        'TEngo k meter en resto parametros
        If Not SeparaRestoParametros(vcmd) Then Exit Sub
        
    Case 5
        'Compruebo k lo siguiente es un numero
        i = InStr(1, vcmd, " ")
        DatosMOdificados = False
        'Actualizar claves
        If i > 0 Then
            If Val(Mid(vcmd, 1, i - 1)) > 0 Then
                
                RestoParametros = Mid(vcmd, 1, i - 1) & "|"
                vcmd = Trim(Mid(vcmd, i)) & " "
                If InStr(1, vcmd, "/") > 0 Then DatosMOdificados = True
            End If
        End If
        
        If Not DatosMOdificados Then
            ParametrosEquivocados2 23, "Error /u. Sin opciones"
            Exit Sub
        End If
        
        'Separaremos en campos de BD
        
        'Campos clave texto
        For i = 1 To 4
            If SeparaUpdatear(vcmd, "/c" & i & ":") Then
                RestoParametros = RestoParametros & DatosCopiados & "|"
            Else
                ParametrosEquivocados2 23, "Parametros erroneos (" & i & ") " & vcmd
                Exit Sub
            End If
        Next i
            
        'Fechas
        For i = 1 To 3
            If SeparaUpdatear(vcmd, "/f" & i & ":") Then
                If DatosCopiados <> "" Then
                    If Not IsDate(DatosCopiados) Then
                        ParametrosEquivocados2 23, "Fechas erroneas (" & i & ") " & vcmd
                        Exit Sub
                    End If
                End If
                RestoParametros = RestoParametros & DatosCopiados & "|"
            Else
                ParametrosEquivocados2 23, "Parametros erroneos (" & i & ") " & vcmd
                Exit Sub
            End If
        Next i
                
        For i = 1 To 2
            If SeparaUpdatear(vcmd, "/i" & i & ":") Then
                If DatosCopiados <> "" Then
                    If Not IsNumeric(DatosCopiados) Then
                        ParametrosEquivocados2 23, "Importes erroneos (" & i & ") " & vcmd
                        Exit Sub
                    End If
                End If
                RestoParametros = RestoParametros & DatosCopiados & "|"
            Else
                ParametrosEquivocados2 23, "Parametros erroneos (" & i & ") " & vcmd
                Exit Sub
            End If
        Next i
        
        
        'Observaciones
        If SeparaUpdatear(vcmd, "/o:") Then
            RestoParametros = RestoParametros & DatosCopiados & "|"
        Else
            ParametrosEquivocados2 23, "Parametros erroneos (" & i & ") " & vcmd
            Exit Sub
        End If
    End Select
    
    
    
    
    'Primero abrimos conexion y la configuracion
    If Not AbrirConexion(True) Then
        
        Exit Sub
    End If
    
        
        
    'Leemos el objeto Confiuracion
    Set vConfig = New CConfiguracion
    If vConfig.Leer(1) = 1 Then
        
        Exit Sub
    End If

    
    
    vcmd = DevuelveDesdeBD("codusu", "usuarios", "login", MiUsuario, "T")
    If vcmd = "" Then
        ParametrosEquivocados2 9, "Usuarios incorrecto"
        Exit Sub
    End If
    
    Set vUsu = New Cusuarios
    If vUsu.Leer(CInt(vcmd)) = 1 Then Exit Sub
    
    Select Case LaOpcion
    Case 1, 2
        HacerSubirBajarFichero
    Case 3
        'NUEVA CARPETA
        L = OpcionNuevaCarpeta(False)
        If L < 0 Then
            'ParametrosEquivocados2 15, "Nueva carpeta"
            Exit Sub
        End If
        InterCambioCarpeta L
        
        
    Case 4
        CopiarArchivoNuevo
        
    Case 5
        'Updatear claves
        HacerUpdateClaves
    End Select
    Set vUsu = Nothing
End Sub
        
        
'Separacioon para la opcion UPDATEAR
'-- Metemos en DATOSCOPIADOS
Private Function SeparaUpdatear(CADENA As String, Patron As String) As Boolean
Dim i As Integer
Dim J As Integer
Dim K As Integer
Dim L As Integer

        L = Len(Patron)

        SeparaUpdatear = False
        
        J = InStr(1, CADENA, Patron)
        If J > 0 Then
            If Mid(CADENA, J + L, 1) = """" Then
                i = 1
                'Es con comillas
                K = InStr(J + L + 1, CADENA, """")
                
            Else
                i = 0
                K = InStr(J + L + 1, CADENA, " ")
            End If
            If K = 0 Then Exit Function
            
            SeparaUpdatear = True
            DatosCopiados = Mid(CADENA, J + i + L, K - J - L - i)
            
        Else
            SeparaUpdatear = True
            DatosCopiados = ""
        End If
        
End Function

        
'Separacion
Private Function SeparaRestoParametros(vcmd As String) As Boolean
Dim i As Integer
Dim J As Integer

        SeparaRestoParametros = False
        RestoParametros = ""
        i = 1
        Do
            J = InStr(i, vcmd, """")
            If J > 0 Then
                RestoParametros = RestoParametros & "1"
                i = J + 1
            End If
            
        Loop Until J = 0
        i = Len(RestoParametros)
        If i = 0 Then
            'No hay comillas
            J = InStr(1, vcmd, " ")
            If J = 0 Then
                vcmd = "Parametros incorrectos "
                ParametrosEquivocados2 17, vcmd
                Exit Function
            End If
            RestoParametros = Mid(vcmd, 1, J - 1) & "|" & Mid(vcmd, J + 1) & "|"
            
        Else
            If (i Mod 2) <> 0 Then
                'Numero comillas impares
                vcmd = "Numero "" incorrecto"
                ParametrosEquivocados2 18, vcmd
                Exit Function
            End If
            If i = 2 Then
                'Solo unot tiene comillas
                If Mid(vcmd, 1, 1) = """" Then
                    J = InStr(2, vcmd, """")
                    RestoParametros = Mid(vcmd, 2, J - 2) & "|" & Trim(Mid(vcmd, J + 1)) & "|"
                Else
                    J = InStr(1, vcmd, """")
                    RestoParametros = Trim(Mid(vcmd, 1, J - 1) & "|" & Mid(vcmd, J + 1, Len(vcmd) - J - 1)) & "|"
                End If
            Else
                'Los dos tienen comillas
                J = InStr(2, vcmd, """")
                RestoParametros = Mid(vcmd, 2, J - 2) & "|"
                vcmd = Trim(Mid(vcmd, J + 1))
                RestoParametros = RestoParametros & Mid(vcmd, 2, Len(vcmd) - 2) & "|"
            End If
        End If
    SeparaRestoParametros = True
End Function

Private Sub ParametrosEquivocados2(Numero As Integer, CADENA As String)
    If MostrarVentanaErrores Then MsgBox "Errores:" & Numero & " - " & CADENA, vbCritical
    MiUsuario = App.Path & "\ErrorShell.txt"
    If Dir(MiUsuario) <> "" Then Kill MiUsuario
    CADENA = "Nº: " & Numero & vbCrLf & CADENA
    Numero = FreeFile
    Open MiUsuario For Output As #Numero
    Print #Numero, "Fecha / Hora : " & Format(Now, "dd/mm/yyyy - hh:mm")
    Print #Numero, "": Print #Numero, "": Print #Numero, ""
    Print #Numero, CADENA
    Close #Numero
End Sub


Private Function ComprobarSubirBajarFichero(ByRef CADENA As String) As Integer
Dim Valor As String
Dim i As Integer
Dim vParametros As String
    
    '23456   S             c:\asdaasd ads asd asd.exe"
    '  id   se bloequ       fichero
    '        SOLO LECTURA
    'Primero, id
    i = InStr(1, CADENA, " ")
    If i = 0 Then
        ComprobarSubirBajarFichero = 5
        Exit Function
    End If
    
    Valor = Mid(CADENA, 1, i - 1)
    If Val(Valor) = 0 Then
        ComprobarSubirBajarFichero = 6
        Exit Function
    End If
    vParametros = Valor & "|"
    
    CADENA = Mid(CADENA, i + 1)
   
   
    i = InStr(1, CADENA, " ")
    If i = 0 Then
        ComprobarSubirBajarFichero = 7
        Exit Function
    End If
    
    Valor = Mid(CADENA, 1, i - 1)
    CADENA = Mid(CADENA, i + 1)
    If Valor <> "S" Then Valor = "N"
        
    vParametros = vParametros & Valor & "|"
    
    If Len(CADENA) = 0 Then
        ComprobarSubirBajarFichero = 8
    Else
        vParametros = vParametros & Trim(CADENA) & "|"
        CADENA = vParametros
        ComprobarSubirBajarFichero = 0
    End If
End Function


Private Sub HacerSubirBajarFichero()
Dim Img As cTimagen
Dim vcmd As String
Dim vCar As Ccarpetas
Dim i As Integer

    'Compruebo k el codimagen existe
    MiUsuario = RecuperaValor(RestoParametros, 1)
    vcmd = DevuelveDesdeBD("codigo", "timagen", "codigo", MiUsuario, "N")
    If vcmd = "" Then
        ParametrosEquivocados2 10, "Cod. imagen NO existe"
        Exit Sub
    End If
    ' Compruebo que existe el fichero
    vcmd = RecuperaValor(RestoParametros, 3)
    If LaOpcion = 2 Then
        'SUBIR.
        
        If vcmd = "" Then
            MiUsuario = "Error devolviendo nombre fichero"
        Else
            If Dir(vcmd, vbArchive) = "" Then
                MiUsuario = "Fichero NO existe"
            Else
                MiUsuario = ""
            End If
        End If
        If MiUsuario <> "" Then
            ParametrosEquivocados2 11, MiUsuario
            Exit Sub
        End If
        
    Else
        'Comprobamos si el nombre de archivo es correcot
        MiUsuario = "Error en nombre fichero"
        i = InStrRev(vcmd, "\")
        If i > 0 Then
            vcmd = Mid(vcmd, 1, i)
            If Dir(vcmd, vbDirectory) = "" Then
                MiUsuario = "Carpeta no existe"
            Else
                MiUsuario = ""
            End If
        End If
        
        If MiUsuario <> "" Then
            ParametrosEquivocados2 13, MiUsuario
            Exit Sub
        End If
        
    End If
    
    'Ahora ya hacemos las acciones propiamente dichas
    Set Img = New cTimagen
    MiUsuario = RecuperaValor(RestoParametros, 1)
    If Img.Leer(CLng(MiUsuario), objRevision.LlevaHcoRevision) = 0 Then
        'Ahora vemos la carpeta
        RecordsetCarpetas True
        If CarpetaVisible(Img.codcarpeta) Then
            Set vCar = New Ccarpetas
            If vCar.Leer(Img.codcarpeta, (ModoTrabajo = 1)) = 0 Then
                DatosCopiados = "NO"
                vcmd = RecuperaValor(RestoParametros, 3)
                If LaOpcion = 2 Then
                    
                    Set frmMovimientoArchivo.vDestino = vCar
                    frmMovimientoArchivo.Opcion = 1
                    frmMovimientoArchivo.Origen = vcmd
                    frmMovimientoArchivo.Destino = CStr(Img.codigo)
                    frmMovimientoArchivo.Show vbModal

                
                Else
                    
                    If Not TraerFicheroFisico(vCar, vcmd, Img.codigo) Then DatosCopiados = "NO"
                            
                End If
                If DatosCopiados <> "" Then
                    'Error
                    
                End If
            End If
    
        Else
            ParametrosEquivocados2 12, "Carpeta inaccesible para el usuario"
        End If
        RecordsetCarpetas False
    End If
    Set Img = Nothing
    
    
    
    
    
    
End Sub

Private Sub RecordsetCarpetas(Abrir As Boolean)
    If Abrir Then
        Set miRSAux = New ADODB.Recordset
        MiUsuario = "select codcarpeta,padre,lecturag,userprop from carpetas"
        miRSAux.Open MiUsuario, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    Else
        miRSAux.Close
        Set miRSAux = Nothing
    End If
End Sub


Private Sub EliminarArchivo()
'aqui aqui
End Sub

Private Function CarpetaVisible(codcarpeta As Integer) As Boolean
Dim Fin As Boolean
    
    CarpetaVisible = False
    If miRSAux.EOF Then
       ParametrosEquivocados2 15, "Leyendo carpetas"
    Else
        Fin = False
        While Not Fin
            miRSAux.Find "codcarpeta = " & codcarpeta, , , 1
            If miRSAux.EOF Then
                Fin = True
            Else
                If miRSAux!padre = 0 Then
                    'Fin , pero si k es accesible
                    Fin = True
                    CarpetaVisible = True
                Else
                    If vUsu.codusu = 0 Then
                        codcarpeta = miRSAux!codcarpeta
                        CarpetaVisible = True
                        Fin = True
                    Else
                        If (miRSAux!userprop = vUsu.codusu) Or (miRSAux!lecturag And vUsu.Grupo) > 0 Then
                            'Tiene permiso
                            codcarpeta = miRSAux!padre
                        Else
                            Fin = True
                        End If
                    End If
                End If
            End If
        Wend
        
    End If
End Function



Private Function OpcionNuevaCarpeta(SoloComprobar As Boolean) As Integer
Dim i As Integer
Dim cad As String
Dim padre As Integer
Dim Carpeta As Ccarpetas

    OpcionNuevaCarpeta = -1
    RecordsetCarpetas True
    padre = 0
    MiUsuario = RecuperaValor(RestoParametros, 2)
    If MiUsuario = "" Then
        'Va a la carpeta RAIZ
        MiUsuario = DevuelveDesdeBD("Nombre", "carpetas", "codcarpeta", 1, "N")
    End If
    While MiUsuario <> ""
        i = InStr(1, MiUsuario, "\")
        If i = 0 Then
            cad = MiUsuario
            MiUsuario = ""
        Else
            cad = Mid(MiUsuario, 1, i - 1)
            MiUsuario = Mid(MiUsuario, i + 1)
        End If
        i = CodigoCarpeta(cad, padre)
        padre = i
        If i < 0 Then
            ParametrosEquivocados2 16, "Carpeta erronea:" & RestoParametros & " - " & MiUsuario & " -" & padre
            Exit Function
        End If
        If Not CarpetaVisible(i) Then
            ParametrosEquivocados2 12, "Carpeta inaccesible para el usuario"
            Exit Function
        End If
        
    Wend
    RecordsetCarpetas False
    
    
    If SoloComprobar Then
        OpcionNuevaCarpeta = padre
        Exit Function
    End If
    

    'Llegados aqui, creamos la nueva carpeta
    Set miRSAux = New ADODB.Recordset
    cad = "Select * from  carpetas where padre=   " & padre
    cad = cad & " and nombre = '" & RecuperaValor(RestoParametros, 1) & "'"
    miRSAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cad = ""
    If Not miRSAux.EOF Then
        If Not IsNull(miRSAux!codcarpeta) Then cad = miRSAux!codcarpeta
    End If
    miRSAux.Close
    Set miRSAux = Nothing
    If cad <> "" Then
        OpcionNuevaCarpeta = Val(cad)
        'ParametrosEquivocados2 30, "Opcion NUEVA CARPETA: " & cad
        Exit Function
    End If
    
    Set Carpeta = New Ccarpetas
    If Carpeta.Leer(padre, (ModoTrabajo = 1)) = 0 Then

    
        'Le cambiamos valores
        With Carpeta
            'le cambiamos el padre
            .Nombre = RecuperaValor(RestoParametros, 1)
            .padre = .codcarpeta
            
            .userprop = vUsu.codusu
            .groupprop = vUsu.GrupoPpal
    
    
        
            'Insertamos
            
            If .Agregar = 1 Then
                ParametrosEquivocados2 21, "Error insertando carpeta"
                Exit Function
                
            Else
                OpcionNuevaCarpeta = .codcarpeta
            End If
                
            
            
        End With
    
    Else
        ParametrosEquivocados2 20, "Error leyendo carpeta"
        Exit Function
    End If
    Set Carpeta = Nothing
    
    'OpcionNuevaCarpeta = 0  'Ok.  Cunado solo es comprobar, indicara el codcarpeta
    
End Function


Private Function CodigoCarpeta(NomCarpeta As String, padre As Integer) As Integer
Dim Rs As New ADODB.Recordset

    CodigoCarpeta = -1
    NomCarpeta = "Select codcarpeta from carpetas where nombre='" & NomCarpeta & "' and padre = " & padre
    Set Rs = New ADODB.Recordset
    Rs.Open NomCarpeta, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not Rs.EOF Then
        If Not IsNull(Rs.Fields(0)) Then CodigoCarpeta = Rs.Fields(0)
    End If
    Rs.Close
    Set Rs = Nothing
End Function


Private Sub CopiarArchivoNuevo()
Dim Img As cTimagen
Dim i As Integer
Dim Carp As Ccarpetas
Dim Extension As String
Dim Tama As Currency

    'Comprobar carpeta 2
    i = OpcionNuevaCarpeta(True)
    If i < 0 Then Exit Sub
    
    Set Carp = New Ccarpetas
    If Carp.Leer(i, (ModoTrabajo = 1)) = 1 Then
        ParametrosEquivocados2 26, "Leyendo obj.carpeta"
        Exit Sub
    End If
    
    'Compruebo:
    
    '1-  Existe el archivo
    MiUsuario = RecuperaValor(RestoParametros, 1)
    If Dir(MiUsuario, vbArchive) = "" Then
        ParametrosEquivocados2 25, "Archivo local no existe"
        Exit Sub
    End If
    
    If FileLen(MiUsuario) = 0 Then
        ParametrosEquivocados2 31, "Tamaño archivo vacio"
        Exit Sub
    End If
    Tama = Round((FileLen(MiUsuario) / 1024), 3)
    
    
    'La extension
    i = InStrRev(MiUsuario, ".")
    
    If i = 0 Then
        ParametrosEquivocados2 26, "Sin extension"
        Exit Sub
    End If
    MiUsuario = LCase(Mid(MiUsuario, i + 1))
    Extension = "codext"
    DatosCopiados = DevuelveDesdeBD("exten", "extension", "exten", MiUsuario, "T", Extension)
    If DatosCopiados = "" Then
        'Extension NO reconocida
        ParametrosEquivocados2 25, "Extension NO reconocida: " & MiUsuario
        Exit Sub
    End If
    
            
    'Obtenemos el nombre de archivo
    MiUsuario = RecuperaValor(RestoParametros, 1)
    i = InStrRev(MiUsuario, "\")
    If i = 0 Then
        ParametrosEquivocados2 25, "Obteniendo nombre archivo"
        Exit Sub
    End If
    'sobre todo tu
    MiUsuario = Mid(MiUsuario, i + 1)
    
    
    Set Img = New cTimagen
    
    Img.campo1 = MiUsuario
    Img.fecha1 = Now
    Img.codcarpeta = Carp.codcarpeta
    Img.codext = Val(Extension)
    Img.userprop = vUsu.codusu
    Img.groupprop = vUsu.GrupoPpal
    Img.lecturag = GrupoLongBD(vUsu.GrupoPpal)
    Img.escriturag = GrupoLongBD(vUsu.GrupoPpal)
    Img.tamnyo = Tama
    
    i = Img.Agregar(objRevision.LlevaHcoRevision, False)
    'Leer carpeta
    If i = 0 Then
        DatosCopiados = "NO"
        Set frmMovimientoArchivo.vDestino = Carp
        frmMovimientoArchivo.Opcion = 1
        'NOOOO le pongo las comillas
        'frmMovimientoArchivo.Origen = """" & RecuperaValor(RestoParametros, 1) & """"
        frmMovimientoArchivo.Origen = RecuperaValor(RestoParametros, 1)
        frmMovimientoArchivo.Destino = CStr(Img.codigo)
        frmMovimientoArchivo.Show vbModal
        If DatosCopiados <> "" Then
            ParametrosEquivocados2 26, "Leyendo obj.carpeta"
            Img.Eliminar
            
        Else
            'Creamos el archivo de intercambio
            i = FreeFile
            Open App.Path & "\InterShell.txt" For Output As #i
            Print #i, Img.codigo
            Close #i
        End If
    End If
    Set Img = Nothing
    Set Carp = Nothing
End Sub



Private Sub HacerUpdateClaves()
Dim Img As cTimagen


    Set Img = New cTimagen
    MiUsuario = RecuperaValor(RestoParametros, 1)
    If Img.Leer(CLng(MiUsuario), objRevision.LlevaHcoRevision) = 0 Then
        
        RecordsetCarpetas True
        MiUsuario = ""
        If Not CarpetaVisible(Img.codcarpeta) Then
            MiUsuario = "NO"
        End If
        RecordsetCarpetas False
        
        If MiUsuario <> "" Then
            ParametrosEquivocados2 12, "Carpeta inaccesible para el usuario"
            Set Img = Nothing
            Exit Sub
        End If
        
        
        '----------------------------------
    
        'Campos
        MiUsuario = RecuperaValor(RestoParametros, 2)
        If MiUsuario <> "" Then Img.campo1 = MiUsuario
    
        MiUsuario = RecuperaValor(RestoParametros, 3)
        If MiUsuario <> "" Then Img.campo2 = MiUsuario
    
        MiUsuario = RecuperaValor(RestoParametros, 4)
        If MiUsuario <> "" Then Img.campo3 = MiUsuario
    
        MiUsuario = RecuperaValor(RestoParametros, 5)
        If MiUsuario <> "" Then Img.campo4 = MiUsuario
    
    
    
    
    
        'Fechas
        MiUsuario = RecuperaValor(RestoParametros, 6)
        If MiUsuario <> "" Then Img.fecha1 = CDate(MiUsuario)
        
        MiUsuario = RecuperaValor(RestoParametros, 7)
        If MiUsuario <> "" Then Img.fecha2 = CDate(MiUsuario)
                
        MiUsuario = RecuperaValor(RestoParametros, 8)
        If MiUsuario <> "" Then Img.fecha3 = CDate(MiUsuario)
    
        'importes
        MiUsuario = RecuperaValor(RestoParametros, 9)
        If MiUsuario <> "" Then Img.importe1 = CCur(MiUsuario)
    
        MiUsuario = RecuperaValor(RestoParametros, 10)
        If MiUsuario <> "" Then Img.importe2 = CCur(MiUsuario)
        
        'Observaciones
        MiUsuario = RecuperaValor(RestoParametros, 11)
        If MiUsuario <> "" Then Img.observa = MiUsuario
        
        
        If Img.Modificar = 1 Then ParametrosEquivocados2 27, "Error obj.modificar"
        
    Else
        ParametrosEquivocados2 27, "Leyendo obj.imagen"
    End If
    
End Sub


Private Sub InterCambioCarpeta(vCodCarpeta As Long)
Dim i As Integer
    'Creamos el archivo de intercambio
            i = FreeFile
            Open App.Path & "\InterShell.txt" For Output As #i
            Print #i, vCodCarpeta
            Close #i
End Sub



Private Function SonArchivosParaInsertar(CadenaArchivos) As Boolean
Dim L As Long
Dim Inicio As Long
Dim PrimeraComilla As Boolean
Dim Ini As Boolean

    SonArchivosParaInsertar = False
    
    Set listacod = New Collection
    
    L = InStr(1, CadenaArchivos, """")
    
    If L > 0 Then
        'Retiro de la cadena los que van entre comillas
        PrimeraComilla = True
        Inicio = L
        While L > 0
            If PrimeraComilla Then
                'BUsco el cierre de la cadena
                PrimeraComilla = False
                
            Else
                'Quito de la cadena la subcadena de entrecomillada
                RestoParametros = Mid(CadenaArchivos, Inicio + 1, L - Inicio - 1)
                listacod.Add CStr(RestoParametros)
                
                CadenaArchivos = Mid(CadenaArchivos, 1, Inicio - 1) & Mid(CadenaArchivos, L + 1)
                Inicio = 0
                PrimeraComilla = True
            End If
            
            
            L = InStr(Inicio + 1, CadenaArchivos, """")
        Wend
        
    End If
    
    'Si quedan sin comillas los proceso loas agrego a la coleccion
    'LOs archivos vendran separados por espacio en blanco
    CadenaArchivos = Trim(CadenaArchivos)
    While CadenaArchivos <> ""
        L = InStr(1, CadenaArchivos, " ")
        If L > 0 Then
            RestoParametros = Mid(CadenaArchivos, 1, L - 1)
            
            If RestoParametros <> "" Then listacod.Add CStr(RestoParametros)
            CadenaArchivos = Mid(CadenaArchivos, L + 1)
        Else
            listacod.Add CStr(CadenaArchivos)
            CadenaArchivos = ""
        End If
    Wend
    
    If listacod.Count = 0 Then Exit Function
    
    
    'Primero abrimos conexion
    If Not AbrirConexion(True) Then
        
        Exit Function
    End If
    
    
    
    'Ya tenemos la lista de archivos
    'Veremos si existen

    If Not ExistenArchivos Then Exit Function
    



    
    
    'Leemos el objeto Confiuracion
    Set vConfig = New CConfiguracion
    If vConfig.Leer(1) = 1 Then
        
        Exit Function
    End If


    'VEo si llevamos revsion documental o no
    Set objRevision = New HcoRevisiones
    objRevision.GuardoLasLecturas = True 'para que guarde la linea por cada una
    
    

        
    'Llegados aqui tenemos la lista de archivos a integrar dentro de aridoc
    'Mostraremos la pantalla multifuncional de agrargar datos
    
    frmNuevoArchivoDrag.Show vbModal


    SonArchivosParaInsertar = True
    
    
    
    
    
End Function

'Dada la listacod de archivos, veremos si existen o no
Private Function ExistenArchivos() As Boolean
Dim L As Long
Dim ExtensionesTratadas As String
Dim Aux As String

    On Error GoTo EEx
    
    
    
    ExistenArchivos = False
    
    
    'Extensiones tratadas
    Set miRSAux = New ADODB.Recordset
    ExtensionesTratadas = "Select codext,exten from extension"
    miRSAux.Open ExtensionesTratadas, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    ExtensionesTratadas = ""
    DatosCopiados = "" 'Para luego tener los datos de las extensiones
    While Not miRSAux.EOF
        DatosCopiados = DatosCopiados & LCase(miRSAux!Exten) & ":" & miRSAux!codext & "|"
        ExtensionesTratadas = ExtensionesTratadas & LCase(miRSAux!Exten) & "|"
        miRSAux.MoveNext
    Wend
    miRSAux.Close
    Set miRSAux = Nothing
    If ExtensionesTratadas = "" Then
        MsgBox "Sin configurar extensiones", vbCritical
        Exit Function
    End If
    ExtensionesTratadas = "|" & ExtensionesTratadas
    DatosCopiados = "|" & DatosCopiados
    
    
    Set listaimpresion = New Collection
    RestoParametros = ""
    For L = 1 To listacod.Count
        MiUsuario = CStr(listacod.Item(L))
        If Dir(MiUsuario, vbArchive) = "" Then
            RestoParametros = RestoParametros & "  NO EXIS - " & MiUsuario & vbCrLf
        Else
            'Vemos la extension del fichero
            LaOpcion = InStrRev(MiUsuario, ".")
            If LaOpcion = 0 Then
                'NO tiene la extension
                RestoParametros = RestoParametros & "  SIN EXT - " & MiUsuario & vbCrLf
            Else
                MiUsuario = LCase(Mid(MiUsuario, LaOpcion + 1))
                If InStr(1, ExtensionesTratadas, "|" & MiUsuario & "|") = 0 Then
                    'No esta entodavia
                    RestoParametros = RestoParametros & "  SIN TRATAR - " & MiUsuario & vbCrLf
                Else
                    
                    'Ya sabemos que tratamos la extension
                    listaimpresion.Add listacod.Item(L)
                    
                End If
            End If
        End If
    Next L
    Set listacod = Nothing
    
    
    If RestoParametros <> "" Then
        
        
        If listaimpresion.Count > 0 Then
            'La montamos al reves la cadena
            RestoParametros = "Errores: " & vbCrLf & RestoParametros
            RestoParametros = "Archivos pendientes: " & listaimpresion.Count & vbCrLf & RestoParametros
            RestoParametros = "Cadena de archivo NO VALIDA. Desea continuar?" & vbCrLf & RestoParametros
            
            
            If MsgBox(RestoParametros, vbQuestion + vbYesNoCancel + vbDefaultButton3) = vbYes Then ExistenArchivos = True
        Else
            RestoParametros = "Cadena de archivo NO VALIDA." & vbCrLf & "Errores: " & vbCrLf & RestoParametros
            MsgBox RestoParametros, vbCritical
        End If
        Exit Function
    Else
        'Todo ok
        ExistenArchivos = True
    End If
    
    Exit Function
EEx:
    MsgBox Err.Description, vbExclamation
End Function
