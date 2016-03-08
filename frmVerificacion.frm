VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVerificacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Verificar ARIDOC"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8235
   Icon            =   "frmVerificacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   8235
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Index           =   1
      Left            =   6840
      TabIndex        =   3
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Verificar"
      Height          =   375
      Index           =   0
      Left            =   5400
      TabIndex        =   2
      Top             =   3600
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   5106
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Proceso"
         Object.Width           =   4763
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Estado"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Resultado"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5640
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVerificacion.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVerificacion.frx":08A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVerificacion.frx":12B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVerificacion.frx":1708
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image imgSele 
      Height          =   240
      Index           =   1
      Left            =   1800
      Picture         =   "frmVerificacion.frx":79A2
      ToolTipText     =   "Quitar seleccion"
      Top             =   120
      Width           =   240
   End
   Begin VB.Image imgSele 
      Height          =   240
      Index           =   0
      Left            =   1440
      Picture         =   "frmVerificacion.frx":7AEC
      ToolTipText     =   "Seleccionar todos"
      Top             =   120
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Procesos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmVerificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cad As String
Dim i As Long

Dim NFich As Integer
Dim Nomfich As String
Dim NErrores As Long

Private Sub Command1_Click(Index As Integer)
    If Index = 0 Then
        cad = ""
        For i = 1 To ListView1.ListItems.Count
            If ListView1.ListItems(i).Checked Then cad = "OK"
        Next i
        If cad = "" Then
            MsgBox "Seleccione una opción", vbExclamation
            Exit Sub
        End If
        
        'Abrimos el fichero para errores
        If Not AbrirFicheroErrores Then Exit Sub
        
        'Revisar procesos
        RevisarProcesos
        
        'Cerrarmos errores
        CerrarErrores
        
        
    Else
        Unload Me
    End If
    
    
End Sub

Private Sub Form_Load()
    Set ListView1.SmallIcons = Me.ImageList1
    'Los diferentes procesos de revision
    CargaProcesosRevision
    
    
End Sub



Private Sub CargaProcesosRevision()
Dim ItmX As ListItem

    Set ItmX = ListView1.ListItems.Add(, "c1", "Carpetas")
    Set ItmX = ListView1.ListItems.Add(, "c2", "Fichero BD/fisico")
    Set ItmX = ListView1.ListItems.Add(, "c3", "Ficheros en las carpetas")
    Set ItmX = ListView1.ListItems.Add(, "c4", "Fichero fisico/BD")
    Set ItmX = ListView1.ListItems.Add(, "c5", "Extensiones")
    Set ItmX = ListView1.ListItems.Add(, "c6", "Extensiones vacias")
    Set ItmX = ListView1.ListItems.Add(, "c7", "Almacen datos")

End Sub

Private Function AbrirFicheroErrores() As Boolean

On Error GoTo EAbrirFicheroErrores
    AbrirFicheroErrores = False
    NFich = FreeFile
    Nomfich = App.Path & "\E" & Format(Now, "yymmdd") & "_" & Format(Now, "hhmm") & ".txt"
    Open Nomfich For Output As #NFich
    NErrores = 0
    AbrirFicheroErrores = True
EAbrirFicheroErrores:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Abrir fichero errores"
    End If
End Function

Private Sub CerrarErrores()

    Close NFich
    If NErrores = 0 Then
        Kill Nomfich
    Else
        If Not MostrarFicheroErrores Then
            MsgBox "Se han producido errores. Consulte el fichero: " & Nomfich, vbExclamation
        End If
    End If
End Sub

Private Function MostrarFicheroErrores() As Boolean

On Error Resume Next
    
    Shell "notepad.exe " & Nomfich, vbNormalFocus
    If Err.Number <> 0 Then
        Err.Clear
        MostrarFicheroErrores = False
    Else
        MostrarFicheroErrores = True
    End If
End Function


Private Sub RevisarProcesos()
Dim i As Integer

    For i = 1 To ListView1.ListItems.Count
        Screen.MousePointer = vbHourglass
        HacerOpcion i
        Me.Refresh
    Next i
    
    Screen.MousePointer = vbDefault
    
    

    
    
    
    
End Sub


Private Sub HacerOpcion(vOpcion As Integer)
Dim bol As Boolean
   With ListView1.ListItems(vOpcion)
        If .Checked Then
            .SmallIcon = 3 'Trabajando
            .SubItems(1) = "Revisando"
            Me.Refresh
            
            'Para cada item hacemos una cosa
            Select Case vOpcion
    
            Case 1
                bol = RevisarCarpetas
            Case 2
                'bol = RevisarArchivos
                bol = RevisarArchivosNUEVA
            Case 3
                bol = RevisarArchivosCarpetas
            Case 4
                bol = CompruebaFisicosEnBD
                
            Case 5
                bol = ComprobarExtensiones
            Case 6
                bol = ExtensionesVacias
            
            Case 7
                bol = RevisarAlmacenes
                
            
            End Select
                
            If bol Then
                .SubItems(2) = "OK"
                .SmallIcon = 2
            Else
                .SubItems(2) = "Errores"
                .SmallIcon = 1
            End If
            .SubItems(1) = "Finalizado"
        Else
            'No trabajamos
            .SmallIcon = 4
        End If
    End With

End Sub



Private Function RevisarCarpetas() As Boolean
    RevisarCarpetas = False
    Set miRSAux = New ADODB.Recordset
    BorrarTemporal1
        
    'metemos todas las carpetas en temporal
    cad = "INSERT INTO tmpfich (codusu, codequipo, imagen) SELECT " & vUsu.codusu & "," & vUsu.PC & ","
    cad = cad & "codcarpeta from carpetas where padre>0"
    Conn.Execute cad
    
    'Eliminaremos de temporal, aquellas que padre existe
    cad = "select carpetas.codcarpeta,carpetas.padre,c2.codcarpeta from carpetas, carpetas as c2"
    cad = cad & " Where carpetas.padre = c2.codcarpeta "
    miRSAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cad = "DELETE from tmpfich WHERE codusu = " & vUsu.codusu & " and codequipo = " & vUsu.PC & " and imagen ="
    While Not miRSAux.EOF
        Conn.Execute cad & miRSAux.Fields(0)
        miRSAux.MoveNext
    Wend
    miRSAux.Close
    
    'luego quedan las k el padre no existen
    'Abrimos para ver si queda alguna
    cad = "select carpetas.nombre,codcarpeta from carpetas,tmpfich"
    cad = cad & " WHERE codusu = " & vUsu.codusu & " and codequipo = " & vUsu.PC & " and imagen = codcarpeta"
    miRSAux.Open cad, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    If Not miRSAux.EOF Then
        Encabezado "Carpetas"
        While Not miRSAux.EOF
            NErrores = NErrores + 1
            Print #NFich, miRSAux.Fields(1) & ": " & miRSAux.Fields(0)
            miRSAux.MoveNext
        Wend
        Pie
    Else
        'OK
        RevisarCarpetas = True
    End If
    miRSAux.Close
    
    
    
    'Este trozo se hace siempre
    cad = "Select padre from carpetas group by padre"
    miRSAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRSAux.EOF
        cad = "UPDATE carpetas set hijos =1 where codcarpeta =" & miRSAux.Fields(0)
        Conn.Execute cad
        miRSAux.MoveNext
    Wend
    miRSAux.Close
    Set miRSAux = Nothing

    
    
    
    Set miRSAux = Nothing
End Function


Private Sub Encabezado(Texto As String)
    Print #NFich, "######################################"
    Print #NFich, Texto
    Print #NFich, "######################################": Print #NFich, ""
End Sub

Private Sub Pie()
    Print #NFich, "------------------------": Print #NFich, "": Print #NFich, "": Print #NFich, "": Print #NFich, ""
End Sub


Private Function RevisarArchivosCarpetas() As Boolean

    RevisarArchivosCarpetas = False
    Set miRSAux = New ADODB.Recordset
    BorrarTemporal1
        
    'metemos todas las carpetas en temporal
    cad = "Select codcarpeta from timagen group by codcarpeta"
    miRSAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cad = "INSERT INTO tmpfich (codusu, codequipo, imagen) VALUES (" & vUsu.codusu & "," & vUsu.PC & ","
    While Not miRSAux.EOF
        Conn.Execute cad & miRSAux.Fields(0) & ")"
        miRSAux.MoveNext
    Wend
    miRSAux.Close
    
    
    'Eliminaremos de temporal, aquellas que  existen
    cad = "Select codcarpeta from carpetas "
    miRSAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cad = "DELETE from tmpfich WHERE codusu = " & vUsu.codusu & " and codequipo = " & vUsu.PC & " and imagen ="
    While Not miRSAux.EOF
        Conn.Execute cad & miRSAux.Fields(0)
        miRSAux.MoveNext
    Wend
    miRSAux.Close


    'las que quedan son referencias k no existen
    cad = "SELECT imagen from tmpfich WHERE codusu = " & vUsu.codusu & " and codequipo = " & vUsu.PC
    miRSAux.Open cad, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    If Not miRSAux.EOF Then
        Encabezado "Archivos carpetas"
        While Not miRSAux.EOF
            NErrores = NErrores + 1
            Print #NFich, "Carpetas no encontrada: " & miRSAux.Fields(0)
            miRSAux.MoveNext
        Wend
        Pie
    Else
        RevisarArchivosCarpetas = True
    End If
    miRSAux.Close

End Function
Private Function RevisarArchivos() As Boolean
Dim Almacen As Integer
Dim Rs As ADODB.Recordset
Dim Carpeta As Integer
Dim AntNerrores As Integer
Dim PrimerInsercion As Boolean

    AntNerrores = NErrores
    RevisarArchivos = False
    cad = "select codigo, timagen.codcarpeta, almacen from carpetas, timagen where timagen.codcarpeta=carpetas.codcarpeta"
    
    cad = cad & "  order by almacen,codigo"
    Set Rs = New ADODB.Recordset
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cad = "INSERT INTO tmpfich (codusu, codequipo, imagen) VALUES (" & vUsu.codusu & "," & vUsu.PC & ","
    Almacen = -1
    PrimerInsercion = True
    While Not Rs.EOF
        If Rs!Almacen <> Almacen Then
            If Almacen > 0 Then
                'Llamar a formulario FTP
                CompruebaFTP Almacen, Carpeta, PrimerInsercion
            End If
            Carpeta = Rs!codcarpeta
            Almacen = Rs!Almacen
            'Borramos tmpfich
            BorrarTemporal1
        End If
        Conn.Execute cad & Rs!codigo & ")"
        Rs.MoveNext
    Wend
    Rs.Close
    If Almacen > 0 Then
        'Llama FTP
        CompruebaFTP Almacen, Carpeta, PrimerInsercion
    End If
    If NErrores = AntNerrores Then
        RevisarArchivos = True
    Else
        Pie
    End If
End Function


'----------------------------------------
'Nueva forma de ver los archivos, si tiene referencas en BD
'Lo que vamos a hacer es borrar tmp
'Ver todas las carpetas con todos los archivos
'cargandolas en tmp. Y finalmente para cada archivo borraremos su entrada en tmp.
'Si al final queda alguna sin eliminar es k no tiene referencia
Private Function RevisarArchivosNUEVA() As Boolean
Dim AntNerrores As Long
    AntNerrores = NErrores
    RevisarArchivosNUEVA = False
    BorrarTemporal1
    
    DatosCopiados = ""
    Set frmMovimientoArchivo.vDestino = Nothing
    frmMovimientoArchivo.Opcion = 6
    frmMovimientoArchivo.Destino = ""
    frmMovimientoArchivo.Show vbModal
        
    
    If DatosCopiados = "" Then
        RevisarArchivosNUEVA = True
    Else
        i = InStr(1, DatosCopiados, "|")
        NErrores = Val(Mid(DatosCopiados, 1, i - 1))
        DatosCopiados = Mid(DatosCopiados, i + 1)
        'Guardamos los errores
        Encabezado "Archivos existentes en BD sin referencia física: "
        Print #NFich, "Total archivos: " & NErrores
        Print #NFich, ""
        Print #NFich, DatosCopiados
        
        Pie
    End If
End Function




Private Sub CompruebaFTP(Alma As Integer, C As Integer, ByRef PrimeraVezInsertamos As Boolean)
Dim vC As Ccarpetas

    
    DatosCopiados = ""
    Set vC = New Ccarpetas
    If vC.Leer(C, (ModoTrabajo = 1)) = 0 Then
        Set frmMovimientoArchivo.vDestino = vC
        frmMovimientoArchivo.Opcion = 6
        frmMovimientoArchivo.Destino = Alma
        frmMovimientoArchivo.Show vbModal
    End If
    Set vC = Nothing
    If DatosCopiados <> "" Then
        If PrimeraVezInsertamos Then
            PrimeraVezInsertamos = False
            Encabezado "Archivos existentes en BD sin referencia física: "
                
        End If
        'Algun error
        i = InStr(1, DatosCopiados, "|")
        cad = Mid(DatosCopiados, 1, i - 1)
        DatosCopiados = Mid(DatosCopiados, i + 1)
        i = Val(cad)
        NErrores = NErrores + i
        Print #NFich, DatosCopiados
    End If
End Sub

Private Function CompruebaFisicosEnBD() As Boolean
Dim vC As Ccarpetas
Dim Rs As ADODB.Recordset
Dim PrimeraVezInsertamos  As Boolean
Dim HaTenidoErrores As Boolean

    DatosCopiados = ""
    BorrarTemporal1

    CompruebaFisicosEnBD = False
    
    
    
    Set Rs = New ADODB.Recordset
    
    'Como hacer un dir a traves de FTP cuesta bastante,
    'lo que haremos sera en funcion del volumnen de datos en la bD
    ' realizar esto es:
    ' Si hay 100 archivos, un unico dir sin nada mas
    ' hay 1000 archivos. Diez bucles de dir(1*) dir(2*)
    'Si hay 10000 archivos . Cien bucles de dir(10*) dir(11*)
    cad = "Select count(*) from timagen"
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = 0
    If Not Rs.EOF Then i = DBLet(Rs.Fields(0), "N")
    Rs.Close
    
    
    Set vC = New Ccarpetas
    cad = "Select codalma,srv,pathreal from almacen where codalma >2"   'En la 3 empiezan
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    PrimeraVezInsertamos = True
    HaTenidoErrores = False
    While Not Rs.EOF
        cad = DevuelveDesdeBD("codcarpeta", "carpetas", "almacen", Rs.Fields(0))
        If cad <> "" Then
             'ok este almacen tiene alguna carpeta
             
             If vC.Leer(Val(cad), (ModoTrabajo = 1)) = 0 Then
                DatosCopiados = ""
                Set frmMovimientoArchivo.vDestino = vC
                frmMovimientoArchivo.Opcion = 7
                frmMovimientoArchivo.Destino = i
                frmMovimientoArchivo.Show vbModal
                
                
                
                'Hay errores. Imprimimos
                If DatosCopiados <> "" Then
                    HaTenidoErrores = True
                    If PrimeraVezInsertamos Then
                        PrimeraVezInsertamos = False
                        Encabezado "Archivos fisicos sin referencia en BD: "
                    End If
'                    'Algun error
'                    I = InStr(1, DatosCopiados, "|")
'                    Cad = Mid(DatosCopiados, 1, I - 1)
'                    DatosCopiados = Mid(DatosCopiados, I + 1)
'                    I = Val(Cad)
                    NErrores = NErrores + 1
                    Print #NFich, "Almacen: " & Rs.Fields(0) & " - Server: " & Rs.Fields(1)
                    Print #NFich, DatosCopiados
                End If
                
                
                
             End If
        End If
        Rs.MoveNext
    Wend
    Rs.Close
    Set vC = Nothing
    If Not HaTenidoErrores Then
        CompruebaFisicosEnBD = True
    Else
        Pie
    End If
End Function



Private Function ComprobarExtensiones() As Boolean
Dim PrimerError As Boolean
    cad = "Select codext from timagen group by codext"
    Set miRSAux = New ADODB.Recordset
    miRSAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    PrimerError = True
    
    While Not miRSAux.EOF
        cad = DevuelveDesdeBD("codext", "extension", "codext", miRSAux!codext, "N")
        If cad = "" Then
            If PrimerError Then Encabezado "Extensiones incorrectas"
            PrimerError = False
            Print #NFich, "Ext: " & miRSAux!codext
            NErrores = NErrores + 1
        End If
        miRSAux.MoveNext
    Wend
    miRSAux.Close
    
    
    
    
    'AHora veremos si las extensiones para cada PC se corresponden con las totales
    cad = "Select count(*) from extension"
    miRSAux.Open cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    i = 0
    If Not miRSAux.EOF Then i = DBLet(miRSAux.Fields(0), "N")
    miRSAux.Close
    
    If i > 0 Then
        'Hay extensiones
        cad = "select count(*) -" & i & " as c1,extensionpc.codequipo,descripcion  from extensionpc,equipos"
        cad = cad & " where extensionpc.codequipo=equipos.codequipo group by extensionpc.codequipo having c1<>0"
        miRSAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRSAux.EOF
            If PrimerError Then Encabezado "Extensiones incorrectas"
            PrimerError = False
            Print #NFich, "Equipo: " & miRSAux!Descripcion & " -> " & miRSAux!C1 + i & "  (" & i & ")"
            NErrores = NErrores + 1
        
            miRSAux.MoveNext
        Wend
        miRSAux.Close
    End If
    
    If Not PrimerError Then
        ComprobarExtensiones = False
        Pie
    Else
        ComprobarExtensiones = True
    End If
    Set miRSAux = Nothing
End Function




Private Function ExtensionesVacias() As Boolean
Dim PrimerError As Boolean
Dim Rs As ADODB.Recordset
    cad = "Select codext,descripcion from extension"
    Set miRSAux = New ADODB.Recordset
    miRSAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    PrimerError = True
    If Not miRSAux.EOF Then
        Set Rs = New ADODB.Recordset
        While Not miRSAux.EOF
            cad = "Select codext from timagen where codext=" & miRSAux!codext
            Rs.Open cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            If Rs.EOF Then
                If PrimerError Then Encabezado "Extensiones vacias"
                PrimerError = False
                Print #NFich, "Ext: " & miRSAux!codext & " - " & miRSAux!Descripcion
                NErrores = NErrores + 1
            End If
            Rs.Close
            
            miRSAux.MoveNext
        
        Wend
        Set Rs = Nothing
    End If
    miRSAux.Close
    If Not PrimerError Then
        ExtensionesVacias = False
        Pie
    Else
        ExtensionesVacias = True
    End If
    Set miRSAux = Nothing
End Function


Private Function RevisarAlmacenes() As Boolean
Dim Er As Boolean
    On Error GoTo ERevisarAlmacenes
    
    cad = "Select * from almacen where codalma > 0"
    Set miRSAux = New ADODB.Recordset
    miRSAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cad = ""
    Er = False
    While Not miRSAux.EOF
        DatosCopiados = "Almacen: " & miRSAux!codalma & " - " & miRSAux!SRV
        cad = miRSAux!SRV & "|" & miRSAux!pathreal & "|" & miRSAux!version & "|"
        cad = cad & DBLet(miRSAux!user, "T") & "|" & DBLet(miRSAux!pwd, "T") & "|"
'
'        vOrigen.SRV = RecuperaValor(Origen, 1)
'    vOrigen.pathreal = RecuperaValor(Origen, 2)
'    vOrigen.version = RecuperaValor(Origen, 3)
'    vOrigen.user = RecuperaValor(Origen, 4)
'    vOrigen.pwd = RecuperaValor(Origen, 5)
    
        frmMovimientoArchivo.Opcion = 12
        frmMovimientoArchivo.Origen = cad
        frmMovimientoArchivo.Show vbModal
    
        If DatosCopiados <> "" Then
            If Not Er Then
                'Es el primer error de cabecera
                Encabezado "Carpetas almacen"
            End If
            Print #NFich, DatosCopiados
            NErrores = NErrores + 1
            Er = True
        End If
        
        'Sig
        miRSAux.MoveNext
    Wend
    miRSAux.Close
    Set miRSAux = Nothing
    If Not Er Then
        RevisarAlmacenes = True
    Else
        Pie
    End If
    Exit Function
ERevisarAlmacenes:
    MuestraError Err.Number, "revisar "
End Function

Private Sub imgSele_Click(Index As Integer)
Dim B As Boolean
    B = Index = 0
    For i = 1 To ListView1.ListItems.Count
        ListView1.ListItems(i).Checked = B
    Next i
End Sub

