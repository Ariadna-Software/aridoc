VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmMovimientoArchivo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Traspaso datos"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   Icon            =   "frmMovimientoArchivo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   720
      Width           =   1455
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   4680
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   5295
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   5415
   End
End
Attribute VB_Name = "frmMovimientoArchivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vOrigen As Ccarpetas
Public vDestino As Ccarpetas
Public Verificar As Boolean
Public Opcion As Byte
    '0  .- Traer los iconos
    '1  .- Llevar un Archivo. Insercion normal
    
    '2 .- Traer un archivo.
    
    '3 .- llevamos un icono nuevo

    'De la temporal tmpfich
    '4.- Copiar archivos
    '5.- Mover archivos
    
    '6.- Verificar archivos en el alamacen Destino

    '7 .- Verificar destino en almacen
    
    '8.- Eliminar un archivo( o varios)


    '10.- Verificar almacen
    '11.-     ""      "    para borrar
    '12.-    "    desde la pantalla de verificacion
    
    
    '13.- Llevar PLANTILLA vacias
    
    '14.- Llevar fichero plantillas predefinidas
    
    '15.- Verificar carpetas
    
    '16.- Llevar archivos a HCO
    
    
    '17.- Volcar estructura sobre disco
    
Public Origen As String   'Cuando son caracteres
Public Destino As String

Private strDatos As String
Private PrimeraVez As Boolean
Private FicheroOK As Boolean
Private TrayendoFichero As Boolean
Private HanCancelado As Boolean
Private SePuedeSalir As Boolean

Private PorFTP As Boolean


Private Sub Command1_Click()
    Label2.Caption = "Cancelando acciones"
    HanCancelado = True
    Command1.Visible = False
    Me.Refresh
End Sub

Private Sub Form_Activate()
Dim T1 As Single
    If PrimeraVez Then
        
        PrimeraVez = False
        
        T1 = Timer
        'El caso 0 es un caso especial
        Select Case Opcion
        Case 0
            TraerIconos
            If PorFTP Then CancelaFTP
            
        Case 1
  
            'Dado un origen y la carpeta se llevara el archivo
            LlevarArchivoFisco
            
        Case 2
            TraerArchivoFisico
             
                     
        Case 3
            LlevaIconoNuevo
             
        Case 4
            Label1.Caption = "Copiando archivos"
            CopiarArchivos
             
        Case 5
            
            Label1.Caption = "Moviendo archivos"
            MoverArchivos
            
        Case 6
            'Verificar ficheros
            Label1.Caption = "Verificar archivos"
            VerificarFicheros
            
        Case 7
            'Como hacer un dir a traves de FTP cuesta bastante,
            'lo que haremos sera en funcion del volumnen de datos en la bD
            ' realizar esto es:
            ' Si hay 100 archivos, un unico dir sin nada mas
            ' hay 1000 archivos. Diez bucles de dir(1*) dir(2*)
            'Si hay 10000 archivos . Cien bucles de dir(10*) dir(11*)
            Label1.Caption = "Verificar archivos"
            VerificarFisico
            
        Case 8
            Label1.Caption = "Eliminar archivo(s)"
            EliminarArchivoRS
            
        Case 10, 11, 12
            VerificarAlmacen Opcion = 11
            
        Case 13, 14
            LlevaPlantillaNueva Opcion = 14
        
        
        Case 15
            'Verificar carpeta o carpeta y subcarpetas
            VerificacionCarpetas
        
        Case 16
            HacerMovimientoAHistorico
            
            
        Case 17
            VolcarFicheros
        Case Else
        
        End Select
        
        If PorFTP Then
            'Cerramos conexion
            CerrarConexion
        End If
        
        SePuedeSalir = True
        Unload Me
    End If
End Sub

Private Sub CancelaFTP()
  Inet1.RequestTimeout = 1
  Inet1.Cancel
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    Me.Command1.Visible = False
    PrimeraVez = True
    SePuedeSalir = False
    Label1.Caption = ""
    Label2.Caption = ""
    PorFTP = False
End Sub

Private Sub CerrarConexion()
    On Error Resume Next
    Label1.Caption = "Cerrar conexion"
    Label1.Refresh
    'Inet1.Execute , "CLOSE"
    Do Until Not Inet1.StillExecuting
        DoEvents
    Loop

    If Err.Number <> 0 Then Err.Clear
End Sub

Public Function ConexionQUIT2()
On Error GoTo EC
    Inet1.Execute Inet1.URL, "Close"
    Do Until Not Inet1.StillExecuting
        DoEvents
    Loop
    Inet1.Cancel
    Do Until Not Inet1.StillExecuting
        DoEvents
    Loop
    Exit Function
EC:
    Err.Clear
End Function

Private Function TraerArchivoFisico() As Boolean
Dim cad As String

On Error GoTo E2
    TraerArchivoFisico = False
    Select Case vOrigen.version
    Case 0
        Conectar vOrigen
        'le quito la primera barra
        cad = vOrigen.pathreal & "/" & Origen
        cad = "GET " & cad & " " & Destino
        Inet1.Execute Inet1.URL, cad
        ' Esperar a que se establezca la conexión
        Do Until Not Inet1.StillExecuting
            DoEvents
        Loop
        
        DatosCopiados = ""
    Case 1
    
        cad = vOrigen.pathreal & "\" & Origen
        FileCopy cad, Destino
        DatosCopiados = ""
    End Select
    TraerArchivoFisico = True
    Exit Function
E2:
    MuestraError Err.Number, "Traer archivo físico"

End Function


Private Function HacerMovimientoAHistorico()
Dim i As Integer
Dim C As String
Dim C2 As String
Dim J As Integer
Dim traspasado As String
 
    
    If vOrigen.version = 0 Then Conectar vOrigen
    C2 = Destino
    
    traspasado = ""
    For i = listaimpresion.Count To 1 Step -1
        J = InStr(1, listaimpresion.Item(i), "||")
        If J > 0 Then
            Origen = Mid(listaimpresion.Item(i), J + 2)
            Destino = C2 & Origen
            Label1.Caption = "Fich: " & Origen
            Label1.Refresh
            If MovimientoFichHco Then
                If Not ComprobarArchivoLlevadoHco(Destino) Then traspasado = traspasado & i & "|"
            End If
            
        Else
            traspasado = traspasado & i & "|"
            Label1.Caption = "ERROR : "
            Label1.Refresh
            espera 0.5
        End If
    Next i


    If vOrigen.version = 0 Then
        ConexionQUIT2
        CerrarConexion
    End If
    DatosCopiados = traspasado
End Function


Private Function ComprobarArchivoLlevadoHco(ByRef vPa As String) As Boolean
    On Error Resume Next
    ComprobarArchivoLlevadoHco = (Dir(vPa, vbArchive) <> "")

End Function

Private Function MovimientoFichHco() As Boolean
Dim cad As String
On Error GoTo EMovimientoFichHco
    MovimientoFichHco = False
    Select Case vOrigen.version
    Case 0
        'le quito la primera barra
        cad = vOrigen.pathreal & "/" & Origen
        cad = "GET " & cad & " " & Destino
        Inet1.Execute Inet1.URL, cad
        ' Esperar a que se establezca la conexión
        Do Until Not Inet1.StillExecuting
            DoEvents
        Loop

        DatosCopiados = ""
    Case 1
        
        cad = vOrigen.pathreal & "\" & Origen
        FileCopy cad, Destino
        DatosCopiados = ""
    End Select
    MovimientoFichHco = True
    Exit Function
EMovimientoFichHco:
    Err.Clear
End Function

Private Function LlevarArchivoFisco() As Boolean
Dim cad As String

On Error GoTo E1

    LlevarArchivoFisco = False
    Select Case vDestino.version
    Case 0
        Conectar vDestino
        Do Until Not Inet1.StillExecuting
            DoEvents
        Loop
        
        'le quito la primera barra
        cad = vDestino.pathreal & "/" & Destino
        cad = "PUT """ & Origen & """ " & cad
        Inet1.Execute Inet1.URL, cad
        ' Esperar a que se establezca la conexión
        Do Until Not Inet1.StillExecuting
            DoEvents
        Loop
        
        
        '------------------------------
        'COMPROBAR
        '--------------------------------
        DatosCopiados = ""
        If True Then
            
            If Not HacerDirFTP_1Archivo(Destino) Then DatosCopiados = "ERROR"
        End If
        
        
    Case 1
        cad = vDestino.pathreal & "\" & Destino
        FileCopy Origen, cad
        DatosCopiados = ""
    End Select
    LlevarArchivoFisco = True
    Exit Function
E1:
    MuestraError Err.Number, "Llevar archivo físico." & Err.Description
End Function




Private Sub Conectar(ByRef vCarpeta As Ccarpetas)
  ' Si el control está ocupado, no realizar otra conexión
  If Inet1.StillExecuting = True Then Exit Sub
  ' Establecer las propiedades
  Inet1.URL = vCarpeta.SRV                    ' dirección URL
  Inet1.UserName = vCarpeta.user  ' nombre de usuario
  Inet1.Password = vCarpeta.pwd        ' contraseña
  Inet1.Protocol = icFTP                    ' protocolo FTP
  Inet1.RequestTimeout = 50                 ' segundos
  PorFTP = True
  SeHaEjecutadoFTP = True
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If Not SePuedeSalir Then Cancel = 1
    If TrayendoFichero Then Cancel = 1
    Screen.MousePointer = vbDefault
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
 
  'Debug.Print State & " - "
  Select Case State
    Case icResolvingHost
        Label1.Caption = "Buscando la dirección IP " & _
                             "del servidor"
    Case icHostResolved
        Label1.Caption = "Encontrada la dirección IP " & _
                             "del servidor"
    Case icConnecting
        Label1.Caption = "Conectando con el servidor"
    Case icConnected
        Label1.Caption = "Conectado con el servidor"
    Case icRequesting
        Label1.Caption = "Enviando petición al servidor"
    Case icRequestSent
        Label1.Caption = "Petición enviada con éxito"
    Case icReceivingResponse
        Label1.Caption = "Recibiendo respuesta del servidor"
    Case icResponseReceived
        Label1.Caption = "Respuesta recibida del servidor"
    Case icDisconnecting
        Label1.Caption = "Desconectando del servidor"
    Case icDisconnected
        Label1.Caption = "Desconectado con éxito del " & _
                             "servidor"
    Case icError
        Label1.Caption = "Error en la comunicación " & _
                             "con el servidor"
                             TrayendoFichero = False
    Case icResponseCompleted
      Dim vtDatos As Variant ' variable de datos
      'Debug.Print "icString: " & icString
      ' Obtener el primer bloque
      vtDatos = Inet1.GetChunk(1024, icString)
      'Debug.Print vtDatos
      DoEvents

      Do
        strDatos = strDatos & vtDatos
        DoEvents
        ' Obtener el bloque siguiente
        vtDatos = Inet1.GetChunk(1024, icString)
      Loop While Len(vtDatos) <> 0
      Debug.Print " --> " & strDatos
      Label1.Caption = "Petición completada con éxito. " & _
                "Se recibieron todos los datos."
   
     TrayendoFichero = False
     FicheroOK = True
  End Select
  Label1.Refresh
End Sub


Private Sub TraerIconos()
Dim Rs As ADODB.Recordset

    'La carpeta leo los datos de ICONOS, k es el almacen 0
    Set Rs = New ADODB.Recordset
    Rs.Open "Select * from almacen where codalma = 0", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Destino = App.Path & "\imagenes\"
    
    If Not Rs.EOF Then
        Set vOrigen = New Ccarpetas
        vOrigen.Almacen = Rs!codalma
        vOrigen.version = Rs!version
        vOrigen.pathreal = Rs!pathreal
        vOrigen.SRV = Rs!SRV
        vOrigen.user = Rs!user
        vOrigen.pwd = Rs!pwd
    End If
    Rs.Close
    
            
    Select Case vOrigen.version
    Case 0
        FicheroOK = False
        TrayendoFichero = True
        Conectar vOrigen
        TraerFicheros
        
    Case 1
        TraerIconosDOS
        
    End Select
    Conn.Execute "UPDATE equipos SET cargaIconsExt= 0 WHERE codequipo=" & vUsu.PC
End Sub


Private Sub TraerIconosDOS()

    On Error GoTo ETraerIconosDOS
    
    Origen = Dir(vOrigen.pathreal & "\*.ico")
    Do
        If Origen <> "" Then
            FileCopy vOrigen.pathreal & "\" & Origen, Destino & "\" & Origen
        End If
        Origen = Dir()
    Loop Until Origen = ""
    
    Exit Sub
ETraerIconosDOS:
    MuestraError Err.Number, "Trayendo iconos"
End Sub


Private Sub TraerFicheros()
  
  Dim nFicheros As Long
  Dim sFicheros() As String
  Dim Arc As String
  Dim i As Long
  
  
  Select Case vOrigen.version
  Case 0
          Inet1.Execute , "cd " & vOrigen.pathreal
          
          
          ' Esperar a que se establezca la conexión
          Do Until Not Inet1.StillExecuting
              DoEvents
          Loop
          Me.Refresh
          
          
          
          strDatos = ""
        
          Inet1.Execute , "pwd"
          ' Esperar a que se establezca la conexión
          Do Until Not Inet1.StillExecuting
              DoEvents
          Loop
          Me.Refresh
          
          
          If strDatos <> vOrigen.pathreal Then
            'ERROR
            MsgBox "Error situando el FTP: " & strDatos & " ---- " & vOrigen.pathreal, vbCritical
            End
          End If
          
          Inet1.Execute , "dir *.*"
          ' Esperar a que se establezca la conexión
          Do Until Not Inet1.StillExecuting
              DoEvents
          Loop
          
          
          
        
          
          
          
          
          ' Obtener la lista de directorios y ficheros del
          ' directorio actual en una matriz de cadenas
          sFicheros = Split(strDatos, vbCrLf)
          nFicheros = UBound(sFicheros) - 1 ' el último está vacío
        
        
          ' Añadir el resto de directorios y ficheros
          For i = 2 To nFicheros - 1
            If sFicheros(i) <> "." Or sFicheros(i) <> ".." Then
              strDatos = Destino & sFicheros(i)
              If Dir(strDatos, vbArchive) <> "" Then Kill strDatos
              Arc = " GET " & vOrigen.pathreal & "/" & sFicheros(i) & " " & strDatos
              Inet1.Execute , Arc
              While TrayendoFichero
                     DoEvents
              Wend
              Do Until Not Inet1.StillExecuting
                DoEvents
              Loop

            End If
          Next i
    
    Case 1
        
    End Select
    
End Sub



Private Sub MoverArchivos()
    Set miRSAux = New ADODB.Recordset
    Origen = "select timagen.campo1 , tmpfich.imagen from timagen,tmpfich where"
    Origen = Origen & " tmpfich.imagen = timagen.codigo AND tmpfich.codusu = " & vUsu.codusu
    Origen = Origen & " AND tmpfich.imagen = timagen.codigo AND tmpfich.codequipo = " & vUsu.PC
    
    If vOrigen.Almacen = vDestino.Almacen Then
        MoverArchivosMismoAlmacen
    Else
        MoverArchivosDistintoAlmacen
    End If
    Set miRSAux = Nothing
End Sub

Private Sub MoverArchivosMismoAlmacen()
    'Lo unico k hay que hacer es un update, si y solo si estan en el mismo almacen
    
    'permisos sobre los ficheros. Ya k en teoria Mover realiza una operacion de eliminar
    'Origen = Origen & " AND (userprop =" & vUsu.codusu & " Or (escriturag And " & vUsu.Grupo & "))"
    
    
    miRSAux.Open Origen, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Origen = "UPDATE timagen SET codcarpeta ="
    While Not miRSAux.EOF
        Label2.Caption = miRSAux.Fields(0)
        Label2.Refresh
        'UPDATEAR
        Conn.Execute Origen & Destino & " WHERE codigo = " & miRSAux.Fields(1)
        
        If objRevision.LlevaHcoRevision Then objRevision.InsertaRevision CLng(miRSAux.Fields(1)), 6, vUsu, vOrigen.codcarpeta & " - " & vOrigen.Nombre
        'Siguiente
        miRSAux.MoveNext
    Wend
    miRSAux.Close

End Sub

Private Sub MoverArchivosDistintoAlmacen()
Dim OK As Byte
    'permisos sobre los ficheros. Ya k en teoria Mover realiza una operacion de eliminar
    'Origen = Origen & " AND (userprop =" & vUsu.codusu & " Or (escriturag And " & vUsu.Grupo & "))"
    
    
    miRSAux.Open Origen, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not miRSAux.EOF
        Label2.Caption = miRSAux.Fields(0)
        Label2.Refresh
        
        
        'TREMOS EL ARCHIVO
        OK = Traer_y_LlevarArchivo(miRSAux.Fields(1))
            
        
        'UPDATEAR
        If OK <> 0 Then
            Origen = "Error: " & DatosCopiados & " Numero: " & OK & vbCrLf & "Avise al soporte técnico"
            Origen = Origen & vbCrLf & vbCrLf & "¿Desea continuar igualmente?"
            If MsgBox(Origen, vbExclamation + vbYesNo) = vbNo Then
                While Not miRSAux.EOF
                    miRSAux.MoveNext
                Wend
            Else
                OK = 0   'para que siga en el while
            End If
        Else
            Conn.Execute "UPDATE timagen SET codcarpeta =" & vDestino.codcarpeta & " WHERE codigo = " & miRSAux.Fields(1)
            If objRevision.LlevaHcoRevision Then objRevision.InsertaRevision CLng(miRSAux.Fields(1)), 6, vUsu, vOrigen.codcarpeta & " - " & vOrigen.Nombre
        End If
        
        'Siguiente
        If OK = 0 Then miRSAux.MoveNext
    Wend
    miRSAux.Close

End Sub

Private Function Traer_y_LlevarArchivo(Id As Long) As Byte

    On Error GoTo eTraer_y_LlevarArchivo
    Traer_y_LlevarArchivo = 1
    
    Origen = Id
    Destino = App.Path & "\temp\" & Id
    If Dir(Destino, vbArchive) <> "" Then Kill Destino
    DatosCopiados = "Trayendo archivo"
    TraerArchivoFisico
    If DatosCopiados <> "" Then Exit Function
    
    'Borramos el origen en el oreigen
    Traer_y_LlevarArchivo = 2
    DatosCopiados = "Eliminando en almacen repositorio"
    'EliminarArchivoFisco True
    EliminarArchivoFisco vOrigen, Origen
    
    If DatosCopiados <> "" Then Exit Function
    
    
    
    'Llevamos al disco otra vez
    Traer_y_LlevarArchivo = 3
    Origen = Destino
    Destino = Id
    DatosCopiados = "Volviendo a llevar"
    LlevarArchivoFisco
    
    
    If DatosCopiados = "" Then
        Traer_y_LlevarArchivo = 0
        'Borramos temporal
        If Dir(Origen, vbArchive) <> "" Then Kill Origen
    End If
    
    Exit Function
eTraer_y_LlevarArchivo:
    MuestraError Err.Number, "Moviendo distintos almacen repositorio"
End Function



Private Sub CopiarArchivos()
Dim Contador As Long
Dim Pos As Integer
Dim RT As ADODB.Recordset

    'Abra k traer el archivo y llevarlo
    Set RT = New ADODB.Recordset
    Origen = "select timagen.campo1 , tmpfich.imagen from timagen,tmpfich where"
    Origen = Origen & " tmpfich.imagen = timagen.codigo AND tmpfich.codusu = " & vUsu.codusu
    Origen = Origen & " AND tmpfich.imagen = timagen.codigo AND tmpfich.codequipo = " & vUsu.PC
    
    'permisos sobre los ficheros. Ya k en teoria Mover realiza una operacion de eliminar
    'Origen = Origen & " AND (userprop =" & vUsu.codusu & " Or (escriturag And " & vUsu.Grupo & "))"
    
    Set RT = New ADODB.Recordset
    RT.Open Origen, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    Contador = 0
    While Not RT.EOF
        Contador = Contador + 1
        RT.MoveNext
    Wend
    
    
    Origen = "UPDATE timagen SET codcarpeta ="
    If Contador > 0 Then
        RT.MoveFirst
        Pos = 1
        While Not RT.EOF
            Label2.Caption = RT.Fields(0) & ". (" & Pos & " de " & Contador & ")"
            Me.Refresh
            Screen.MousePointer = vbHourglass
            'UPDATEAR
            HacerCopiaArchivo RT.Fields(1)
            'Siguiente
            Pos = Pos + 1
            RT.MoveNext
        Wend
        'Desconectar
        Inet1.Cancel
    End If
    RT.Close
    Set RT = Nothing
End Sub


Private Sub HacerCopiaArchivo(Imgagen As Long)
Dim IO As cTimagen
Dim Id As cTimagen
Dim AsignacionOk As Boolean
Dim OK As Boolean

    'Vamos con lo qu vamos
    Set Id = New cTimagen
    Set IO = New cTimagen
    
    'leemos imagen
    AsignacionOk = False
    If IO.Leer(Imgagen, objRevision.LlevaHcoRevision) = 0 Then
        Id.campo1 = IO.campo1
        Id.campo2 = IO.campo2
        Id.campo3 = IO.campo3
        Id.campo4 = IO.campo4
        'Extension
        Id.codext = IO.codext
        
        'Carpeta contenedora
        Id.codcarpeta = vDestino.codcarpeta
        
        'Fechas
        Id.fecha1 = IO.fecha1
        Id.fecha2 = IO.fecha2
        Id.fecha3 = IO.fecha3
        'importes
        Id.importe1 = IO.importe1
        Id.importe2 = IO.importe2
        
        Id.tamnyo = IO.tamnyo
        Id.userprop = vUsu.codusu
        Id.groupprop = vUsu.GrupoPpal
        
        'Permisos. Habra que estudiarlo ###
        '--------------------------------
        Id.lecturag = IO.lecturag
        Id.escriturag = IO.escriturag
        
        'intentamos añadir
        If Id.Agregar(objRevision.LlevaHcoRevision, False) = 0 Then AsignacionOk = True
    End If


    'Llegdos aqui, si la signacion  ha ido bien entonces
    If AsignacionOk Then
        OK = False
        Origen = Imgagen
        Destino = App.Path & "\Temp\" & Imgagen
        
        If TraerArchivoFisico Then
            
            'Llevamos
            Origen = Destino
            Destino = Id.codigo
            If LlevarArchivoFisco Then OK = True
            
            'ELiminamos el temporal
            If Dir(Origen) <> "" Then Kill Origen
            
        End If 'De traer el archivo
            
        'Cancelamos
        Inet1.Cancel
        
        'Si no ha ido bien elimino el objeto
        If Not OK Then Id.Eliminar
        
        
    End If
    
        
    Set IO = Nothing
    Set Id = Nothing
End Sub



Private Sub VerificarFicheros()
Dim Contador As Long
Dim Pos As Long
Dim Errores As Long

    'Metemos todos los archivos en la tabla temporal
    Origen = "Select * from almacen where codalma >2"
    Set miRSAux = New ADODB.Recordset
    miRSAux.Open Origen, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Errores = 0


    While Not miRSAux.EOF
        Set vDestino = Nothing
        Set vDestino = New Ccarpetas
        Label2.Caption = "Verificando almacen: " & miRSAux!codalma
        Label2.Refresh
        vDestino.SRV = miRSAux!SRV
        vDestino.pathreal = miRSAux!pathreal
        vDestino.pwd = miRSAux!pwd
        vDestino.user = miRSAux!user
        vDestino.version = miRSAux!version
        Destino = 6000
        
        VerificarFisico
        
        miRSAux.MoveNext
    Wend
    miRSAux.Close

    
    Label2.Visible = True
    Label2.Caption = "Borrando temporal2"
    Me.Refresh
    BorrarTemporal2
    
    Label2.Caption = "Leyendo tabla IMG"
    Label2.Refresh
    Origen = "INSERT INTO tmpbusqueda (codcarpeta,codusu, codequipo, imagen) "
    Origen = Origen & "Select 0," & vUsu.codusu & "," & vUsu.PC & ",codigo from timagen "
    
    Conn.Execute Origen
    
    
    Origen = "DELETE from tmpbusqueda WHERE codusu = " & vUsu.codusu & " AND codequipo = " & vUsu.PC & " AND imagen = "
    
    Destino = "from tmpfich where codusu = " & vUsu.codusu & " and codequipo = " & vUsu.PC
    miRSAux.Open "Select count(*) " & Destino, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    Contador = 0
    If Not miRSAux.EOF Then Contador = DBLet(miRSAux.Fields(0), "N")
    miRSAux.Close
    
    
    miRSAux.Open "Select imagen " & Destino, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Pos = 1
    While Not miRSAux.EOF
        
        Label2.Caption = "Comprobando fichero " & Pos & " de " & Contador
        Label2.Refresh
        Conn.Execute Origen & miRSAux!imagen
        
        miRSAux.MoveNext
        Pos = Pos + 1
    Wend
    miRSAux.Close
    Inet1.Cancel
    
    Label2.Caption = "Comprobando fichero sin correspondencia."
    Label2.Refresh
        
    
    'Finalmente comprobamos las referencias que quedan sin eliminar en
    'tmpbusqueda
    Origen = "Select * from tmpbusqueda where codusu = " & vUsu.codusu & " and codequipo = " & vUsu.PC
    Origen = Origen & " ORDER by imagen"
    Pos = 1
    miRSAux.Open Origen, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Origen = ""
    While Not miRSAux.EOF
        Origen = Origen & Right("          " & miRSAux!imagen, 10)
        If (Pos Mod 9) = 0 Then Origen = Origen & vbCrLf
        Pos = Pos + 1
        miRSAux.MoveNext
    Wend
    miRSAux.Close
    Set miRSAux = Nothing
    
    'metemos el numero de errores al princi`pio
    If Pos > 1 Then DatosCopiados = Pos & "|" & DatosCopiados & Origen
        
        
        
    Label2.Visible = False
'----------------------------------------------------------
'----------------------------------------------------------
'----------------------------------------------------------
'ANTIGUO
'----------------------------------------------------------
'----------------------------------------------------------
'----------------------------------------------------------
'    Set miRSAux = New ADODB.Recordset
'    Origen = " from tmpfich where codusu =" & vUsu.codusu & " AND "
'    Origen = Origen & " codequipo =" & vUsu.PC
'    If vDestino.version = 0 Then Conectar vDestino
'
'    miRSAux.Open "Select count(*)  " & Origen, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'    Contador = 0
'    If Not miRSAux.EOF Then
'        Contador = DBLet(miRSAux.Fields(0), "N")
'    End If
'    If Contador = 0 Then Exit Sub
'    miRSAux.Close
'
'    miRSAux.Open "Select * " & Origen, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'    Pos = 0
'    Errores = 0
'    While Not miRSAux.EOF
'        Pos = Pos + 1
'        Label2.Caption = "Alm: " & vDestino.Almacen & ". Fich: " & miRSAux!imagen & "  ( " & Pos & " de " & Contador & ")"
'        Label2.Refresh
'        If vDestino.version = 0 Then
'            If Not HacerDirFTP(CStr(miRSAux!imagen)) Then
'                Errores = Errores + 1
'                DatosCopiados = DatosCopiados & "Fich: " & miRSAux!imagen & vbCrLf
'            End If
'        Else
'             '
'        End If
'
'        miRSAux.MoveNext
'
'    Wend
'    miRSAux.Close
'    Inet1.Cancel
'    'metemos el numero de errores al princi`pio
'    If Errores > 0 Then DatosCopiados = Errores & "|" & DatosCopiados
'

        
End Sub

Private Function HacerDirFTP2(Archivo As String) As Boolean
    HacerDirFTP2 = False
    'Intento ver el tamño


    Inet1.Execute Inet1.URL, "dir " & Archivo
     
    ' Esperar a que se establezca la conexión
    Do Until Not Inet1.StillExecuting
        DoEvents
    Loop

    If InStr(1, strDatos, Archivo) > 0 Then
        HacerDirFTP2 = True
    Else
        strDatos = ""
        Label2.Caption = "ERROR: " & Archivo
        Me.Refresh
        'espera 1
    End If
End Function

Private Function HacerDirMSDOS(Archivo As String) As Boolean
    On Error Resume Next
    HacerDirMSDOS = False
    If Dir(Archivo, vbArchive) <> "" Then HacerDirMSDOS = True
    If Err.Number <> 0 Then Err.Clear
End Function

Private Function HacerDirFTPComodin(vDir As String) As Boolean
    HacerDirFTPComodin = False
    'Intento ver el tamño
   
    Inet1.Execute Inet1.URL, "dir " & vDir
    ' Esperar a que se establezca la conexión
    Do Until Not Inet1.StillExecuting
        DoEvents
    Loop


End Function


Private Function HacerDirMSDOSComodin(vDir As String) As Boolean
Dim MiNombre As String
    
    MiNombre = Dir(vDestino.pathreal & "\" & vDir, vbArchive)
    Do While MiNombre <> ""   ' Inicia el bucle.
        strDatos = strDatos & MiNombre & vbCrLf
    
       MiNombre = Dir   ' Obtiene siguiente entrada.
    Loop

End Function





Private Function CambiaDirectorioFTP2(ByVal Directorio As String) As Boolean
Dim J As Integer
Dim K As Integer
Dim OK As Boolean
Dim Aux As String
On Error GoTo ECambiaDirectorio
    
   
       CambiaDirectorioFTP2 = False
       
       If Right(Directorio, 1) <> "/" Then Directorio = Directorio & "/"
       J = 2
       Do
            K = InStr(J, Directorio, "/")
            If K > 0 Then
                Aux = Mid(Directorio, J, K - J)
                J = K + 1
                strDatos = ""
                Inet1.Execute Inet1.URL, "cd " & Aux
                Do Until Not Inet1.StillExecuting
                    DoEvents
                Loop
            
              
            
                Inet1.Execute Inet1.URL, "pwd"
                ' Esperar a que se establezca la conexión
                Do Until Not Inet1.StillExecuting
                    DoEvents
            
                  
                Loop
            Else
                OK = True
            End If
            Me.Refresh
          
          
        Loop Until OK
        
        
        CambiaDirectorioFTP2 = True
    
    
    
    Exit Function
ECambiaDirectorio:
    MuestraError Err.Number, "CambiaDirectorio" & Err.Description
    
End Function

Private Sub VerificarFisico()
Dim i As Integer
Dim C As Long
Dim Veces As Integer
Dim CadDIR As String


    
    
    'Dependiendo de la version hara unas cosas u otras
    If vDestino.version = 0 Then
        Conectar vDestino
        If Not CambiaDirectorioFTP2(vDestino.pathreal) Then
            'ERRROR
            
        End If
    
    End If
    
    
    
        C = Val(Destino)
        If C < 100 Then
            Veces = 1
        Else
            If C < 10000 Then
                Veces = 10
            Else
                If C < 100000 Then
                    Veces = 100
                Else
                    If C < 1000000 Then
                        Veces = 1000
                    Else
                        Veces = 10000
                    End If
                End If
            End If
        End If
        HanCancelado = False
        Command1.Visible = True
        
        
        
        
        
        
        
        Label2.Visible = True
        For i = 1 To (Veces - 1)
            Label2.Caption = "Archivos : " & i & " de " & Veces
            Label2.Refresh
            strDatos = ""
            If Veces = 1 Then
                    CadDIR = "*"
            Else
                C = Veces - i
                CadDIR = C & "*"
            End If
            
            If vDestino.version = 0 Then
                'Por FTP
                HacerDirFTPComodin CadDIR
            Else
                'Debug.Print CadDIR
                HacerDirMSDOSComodin CadDIR
                If strDatos <> "" Then strDatos = strDatos & vbCrLf
            End If
            
            If HanCancelado Then Exit Sub
            If Opcion = 7 Then
                CompruebaDatosDevueltosDir CadDIR
            Else
                'En la oppcion 6 tb lamamos a este procedimiento
                ProcesaVerificandoFisico
            End If
            If HanCancelado Then Exit Sub
        
        
            espera 0.5
        Next i
        Label2.Visible = False
    
    
    
End Sub



Private Sub CompruebaDatosDevueltosDir(vDir As String)

    Dim nFicheros As Long, i As Long
    Dim sFicheros() As String
    Dim Inserta As Boolean
    
    ' Obtener la lista de directorios y ficheros del
    ' directorio actual en una matriz de cadenas
    sFicheros = Split(strDatos, vbCrLf)
    nFicheros = UBound(sFicheros) - 1 ' el último está vacío
    
    
    ' Añadir el resto de directorios y ficheros
    If nFicheros > 0 Then
        i = InStr(1, vDir, "*")
        vDir = Mid(vDir, 1, i - 1) & "%"
        Origen = "Select codigo from timagen"
        If Len(vDir) > 1 Then Origen = Origen & " where codigo like '" & vDir & "'"
         
        Set miRSAux = New ADODB.Recordset
        miRSAux.Open Origen, Conn, adOpenKeyset, adLockOptimistic, adCmdText
        For i = 0 To nFicheros - 1
            '-------
            DoEvents
            If HanCancelado Then Exit Sub
            
            '------------------------------------
            If Right(sFicheros(i), 1) <> "/" Then
                If sFicheros(i) <> "" Then
                    If IsNumeric(sFicheros(i)) Then
                        miRSAux.Find "codigo = " & sFicheros(i), , , 1
                        If miRSAux.EOF Then
                            'NO encontrado
                            DatosCopiados = DatosCopiados & "Fichero: " & sFicheros(i) & vbCrLf
                        End If
                    Else
                        DatosCopiados = DatosCopiados & "Fichero no numerico: " & sFicheros(i) & vbCrLf
                    End If
                End If
            End If
        Next i
        miRSAux.Close
        Set miRSAux = Nothing
    End If
End Sub



Private Sub ProcesaVerificandoFisico()
    Dim nFicheros As Long, i As Long
    Dim sFicheros() As String
    Dim Inserta As Boolean
    
    ' Obtener la lista de directorios y ficheros del
    ' directorio actual en una matriz de cadenas
    sFicheros = Split(strDatos, vbCrLf)
    nFicheros = UBound(sFicheros) - 1 ' el último está vacío
    
   
    ' Añadir el resto de directorios y ficheros
    If nFicheros > 0 Then
        For i = 0 To nFicheros - 1
            
            '------------------------------------
            If Right(sFicheros(i), 1) <> "/" Then
                If sFicheros(i) <> "" Then
                    If IsNumeric(sFicheros(i)) Then
                         If Not InsertaTemporal(Val(sFicheros(i))) Then
                            'NO encontrado
                            DatosCopiados = DatosCopiados & "Fichero: " & sFicheros(i) & vbCrLf
                        End If
                    Else
                        DatosCopiados = DatosCopiados & "Fichero no numerico: " & sFicheros(i) & vbCrLf
                    End If
                End If
            End If
        Next i
    End If
End Sub


Private Sub LlevaIconoNuevo()

    Set miRSAux = New ADODB.Recordset
    miRSAux.Open "Select * from almacen where codalma = 0", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not miRSAux.EOF Then
        Set vDestino = New Ccarpetas
        vDestino.Almacen = miRSAux!codalma
        vDestino.version = miRSAux!version
        vDestino.pathreal = miRSAux!pathreal
        vDestino.SRV = miRSAux!SRV
        vDestino.user = miRSAux!user
        vDestino.pwd = miRSAux!pwd
    End If
    miRSAux.Close
    Set miRSAux = Nothing
    
    'Llevamos el archivo fisico
    LlevarArchivoFisco
    
    Set vDestino = Nothing

End Sub


Private Sub LlevaPlantillaNueva(Predefinida As Boolean)
Dim C As String

    Set miRSAux = New ADODB.Recordset
    C = "Select * from almacen where codalma = "
    If Predefinida Then
        C = C & "2"
    Else
        C = C & "1"
    End If
    miRSAux.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not miRSAux.EOF Then
        Set vDestino = New Ccarpetas
        vDestino.Almacen = miRSAux!codalma
        vDestino.version = miRSAux!version
        vDestino.pathreal = miRSAux!pathreal
        vDestino.SRV = miRSAux!SRV
        vDestino.user = miRSAux!user
        vDestino.pwd = miRSAux!pwd
    End If
    miRSAux.Close
    Set miRSAux = Nothing
    
    'Llevamos el archivo fisico
    LlevarArchivoFisco
    
    Set vDestino = Nothing

End Sub

Private Sub EliminarArchivoRS()
    'Abriremos el RS con los archivos k hay k eliminar
    Me.Refresh
    Destino = "Select tmpfich.imagen,campo1 from tmpfich,timagen where"
    Destino = Destino & " tmpfich.imagen=timagen.codigo AND  codusu =" & vUsu.codusu & " AND codequipo = " & vUsu.PC
    Set miRSAux = New ADODB.Recordset
    miRSAux.Open Destino, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not miRSAux.EOF
        Destino = miRSAux!imagen
        Label2.Caption = "Arch: " & miRSAux!imagen & " - " & miRSAux!campo1
        Label2.Refresh
        If EliminarArchivoFisco(vDestino, Destino) Then
            Conn.Execute "DELETE FROM timagen WHERE codigo = " & miRSAux!imagen
            If objRevision.LlevaHcoRevision Then objRevision.EliminarReferencia miRSAux!imagen
        End If
        miRSAux.MoveNext
    
    Wend
    miRSAux.Close
    Set miRSAux = Nothing
    Inet1.Cancel
End Sub

'Private Function EliminarArchivoFisco(EsMoverArchivos As Boolean) As Boolean
'Dim Cad As String
'
'On Error GoTo E1
'
'    EliminarArchivoFisco = False
'    Select Case vDestino.version
'    Case 0
'        If EsMoverArchivos Then
'            Conectar vOrigen
'            'le quito la primera barra
'            Cad = vOrigen.pathreal & "/" & Origen
'            Cad = "delete " & Cad
'
'        Else
'            Conectar vDestino
'            'le quito la primera barra
'            Cad = vDestino.pathreal & "/" & Destino
'            Cad = "delete " & Cad
'        End If
'        Inet1.Execute , Cad
'        ' Esperar a que se establezca la conexión
'        Do Until Not Inet1.StillExecuting
'            DoEvents
'        Loop
'
'        DatosCopiados = ""
'    Case 1
'
'        Cad = vOrigen.pathreal & "\" & Origen
'        Kill Cad
'    End Select
'    EliminarArchivoFisco = True
'    Exit Function
'E1:
'    MuestraError Err.Number, "Llevar archivo físico"
'End Function




Private Function EliminarArchivoFisco(ByRef C As Ccarpetas, ByRef F As String) As Boolean
Dim cad As String

On Error GoTo E1

    EliminarArchivoFisco = False
    Select Case C.version
    Case 0
            Conectar C
            'le quito la primera barra
            cad = C.pathreal & "/" & F
            cad = "delete " & cad
        
        
        Inet1.Execute , cad
        ' Esperar a que se establezca la conexión
        Do Until Not Inet1.StillExecuting
            DoEvents
        Loop
        
        DatosCopiados = ""
    Case 1
    
        cad = C.pathreal & "\" & F
        Kill cad
        DatosCopiados = ""
    End Select
    If DatosCopiados = "" Then EliminarArchivoFisco = True
    Exit Function
E1:
    MuestraError Err.Number, "Eliminar archivo físico"
End Function





Private Sub VerificarAlmacen(HacerDir As Boolean)
    Set vOrigen = New Ccarpetas
    vOrigen.SRV = RecuperaValor(Origen, 1)
    vOrigen.pathreal = RecuperaValor(Origen, 2)
    vOrigen.version = RecuperaValor(Origen, 3)
    vOrigen.user = RecuperaValor(Origen, 4)
    vOrigen.pwd = RecuperaValor(Origen, 5)
    If vOrigen.version = 0 Then
        If VerificaViaFtp(HacerDir) Then DatosCopiados = ""
        Inet1.Cancel
    Else
        If VerificaViaMSDOS Then DatosCopiados = ""
    End If
    
End Sub


Private Function VerificaViaFtp(HacerDir As Boolean) As Boolean
On Error GoTo ev
    VerificaViaFtp = False
    Label2.Caption = "Conectando : " & vOrigen.SRV
    Label2.Refresh
    Conectar vOrigen
    
    Inet1.Execute , "cd " & vOrigen.pathreal
    Do
        Label2.Caption = vOrigen.pathreal
        Label2.Refresh
        espera 0.1
        DoEvents
    Loop Until Not Inet1.StillExecuting
    
    If HacerDir Then
        Label2.Caption = "Dir: " & vOrigen.pathreal
        Label2.Refresh
        Inet1.Execute , "dir *" & vOrigen.pathreal
        Do
            espera 0.1
        Loop Until Not Inet1.StillExecuting
    
    End If
    VerificaViaFtp = True
    Exit Function
ev:
    If Opcion <> 12 Then MuestraError Err.Number, "Verificar FTP" & vbCrLf & Err.Description
End Function



Private Function VerificaViaMSDOS() As Boolean
On Error GoTo ev
    VerificaViaMSDOS = False
    Label2.Caption = "Conectando : " & vOrigen.SRV
    Label2.Refresh
    
    
    Dir vOrigen.pathreal & "\david.david"
    
    
    VerificaViaMSDOS = True
    Exit Function
ev:
    If Opcion <> 12 Then MuestraError Err.Number, "Verificar NETBIOS" & vbCrLf & Err.Description
End Function




Private Sub VerificacionCarpetas()
Dim i As Long
Dim TOT As Long
Dim VersionAnt As Integer
    

    'Borramos temporal
    BorrarTemporal2
    
    Label2.Visible = True
    Label2.Caption = "Obteniendo ficheros"
    Me.Refresh
    
    'Insertamos cada carpeta
    Do
        i = InStr(1, Origen, "|")
        If i > 0 Then
            strDatos = Mid(Origen, 1, i - 1)
            Origen = Mid(Origen, i + 1)
            
            Destino = "INSERT INTO tmpbusqueda(codusu, codequipo, imagen, codcarpeta) SELECT " & vUsu.codusu & "," & vUsu.PC
            Destino = Destino & " ,codigo,codcarpeta from timagen where codcarpeta = " & strDatos
            Conn.Execute Destino
            
            
            
        End If
    Loop Until i = 0
    
    
    
    'Ya tengo los archivos
    Set miRSAux = New ADODB.Recordset
    Destino = " from tmpbusqueda where codusu = " & vUsu.codusu & " AND codequipo = " & vUsu.PC
    miRSAux.Open "Select count(*) " & Destino, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    TOT = 0
    If Not miRSAux.EOF Then TOT = DBLet(miRSAux.Fields(0), "N")
    miRSAux.Close
    
    If TOT = 0 Then
        'CERO? mal. Carpeta vacia
        MsgBox "No existe documentos en la seleccion ", vbExclamation
        Set miRSAux = Nothing
        Exit Sub
    End If
    
    Set listacod = New Collection

    miRSAux.Open "Select * " & Destino & " ORDER BY codcarpeta", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = 1
    Destino = ""
    Origen = ""
    VersionAnt = -1
    Set vDestino = New Ccarpetas
    While Not miRSAux.EOF
        If miRSAux!codcarpeta <> Destino Then
            If vDestino.Leer(miRSAux!codcarpeta, (ModoTrabajo = 1)) = 1 Then
                miRSAux.Close
                Exit Sub
            End If
            If Destino = "" Or VersionAnt <> vDestino.version Then
                If vDestino.version = 0 Then
                    Conectar vDestino
                    Inet1.RequestTimeout = 7
                End If
                
                VersionAnt = vDestino.version
            End If
            Destino = miRSAux!codcarpeta
            
            
            
            'Ya hemos leido la carpeta
            If vDestino.pathreal <> Origen Then
                If vDestino.version = 0 Then CambiaDirectorioFTP2 vDestino.pathreal
                Origen = vDestino.pathreal
            End If
        
        End If
        
        Label2.Caption = "Verificando archivos: (" & miRSAux!imagen & ").      " & i & " de " & TOT
        Label2.Refresh
        'Comprobamos si el archivo existe
        strDatos = ""
        If vDestino.version = 0 Then
            If Not HacerDirFTP2(CStr(miRSAux!imagen)) Then
                listacod.Add CStr(miRSAux!imagen)
                Me.Refresh
            End If
        Else
            'DIR de NETBIOS
            If Not HacerDirMSDOS(Origen & "\" & miRSAux!imagen) Then
                listacod.Add CStr(miRSAux!imagen)
                Me.Refresh
            End If
        End If
        
        'Siguiente
        i = i + 1
        If (i Mod 150) = 0 Then
            Me.Refresh
            DoEvents
        End If
        miRSAux.MoveNext
    Wend
    miRSAux.Close
    
    
    
    
    
    If listacod.Count > 0 Then
    
        'Obteniedo datos del registro de errores
        Label2.Caption = "Obteniendo datos del registro de errores"
        Label2.Refresh
        
        'Meto los valores en la tabla tmpfich1
        BorrarTemporal1
        Destino = "INSERT INTO tmpfich (codusu, codequipo, imagen) VALUES ("
        Destino = Destino & vUsu.codusu & "," & vUsu.PC & ","
        
        For i = 1 To listacod.Count
            Label2.Caption = "Errores : " & i & " de " & listacod.Count
            Label2.Refresh
            Conn.Execute Destino & listacod.Item(i) & ")"
        Next i
    End If
    
End Sub



'-----------------------------------------------------
'
' VOLCAR FICHEROS A DISCO DURO




Private Sub VolcarFicheros()


    If vOrigen.version = 0 Then
        Conectar vOrigen
    End If
    
    While Not miRSAux.EOF
        
        Origen = miRSAux!codigo
        Destino = NombreMSDOS(miRSAux!campo1)
        Label2.Caption = vOrigen.Nombre & "\" & Destino
        Label2.Refresh
        
        Select Case vOrigen.version
        Case 0
            'Le pongo las comillas
            Destino = """" & vDestino.pathreal & "\" & Destino & "." & miRSAux!Exten & """"
            'le quito la primera barra
            Origen = vOrigen.pathreal & "/" & Origen
            Origen = "GET " & Origen & " " & Destino
            
            
           
            
            
        Case 1
            Destino = vDestino.pathreal & "\" & Destino & "." & miRSAux!Exten
            Origen = vOrigen.pathreal & "\" & Origen
           
        End Select
        CopiaDesdeVolcado
        espera 0.05
        miRSAux.MoveNext
    Wend
                    

        
    Label2.Caption = "fin"
    Me.Refresh
    If vOrigen.version = 0 Then CerrarConexion

End Sub


Private Sub CopiaDesdeVolcado()
On Error GoTo EC

    '''''If miRSAux!codigo = 28 Then Stop

    If vOrigen.version = 0 Then
        Inet1.Execute , Origen
        ' Esperar a que se establezca la conexión
        Do Until Not Inet1.StillExecuting
            DoEvents
        Loop
    Else
         FileCopy Origen, Destino
         
    End If
    
    Exit Sub
EC:
    listaimpresion.Add miRSAux!codigo & "-" & miRSAux!campo1 & " (" & Origen & ")   --> " & Destino
End Sub




Private Function HacerDirFTP_1Archivo(Archivo As String) As Boolean
    HacerDirFTP_1Archivo = False
    'Intento ver el tamño

    CambiaDirectorioFTP2 vDestino.pathreal
    ' Esperar a que se establezca la conexión
    Do Until Not Inet1.StillExecuting
        DoEvents
    Loop
    Inet1.Execute Inet1.URL, "dir " & vDestino.pathreal & "/" & Archivo
     
    ' Esperar a que se establezca la conexión
    Do Until Not Inet1.StillExecuting
        DoEvents
    Loop

    If InStr(1, strDatos, Archivo) > 0 Then
        HacerDirFTP_1Archivo = True
    Else
        strDatos = ""
        Label2.Caption = "ERROR: " & Archivo
        Me.Refresh
        espera 0.1
    End If
End Function

