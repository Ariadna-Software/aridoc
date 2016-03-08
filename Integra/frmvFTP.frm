VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmvFTP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Traspaso datos"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
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
Attribute VB_Name = "frmvFTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public opcion As Byte
        'Copiar archivos
Public Fich As String   'Sera el fichero ARI


'Public Usuario As String
'Public Password As String
'Public Destino As String
'Public Origen As String
'Public Servidor As String

Private Origen As String
Private CarpetaD As Ccarpetas


Private strDatos As String
Private PrimeraVez As Boolean
Private FicheroOK As Boolean
Private TrayendoFichero As Boolean
Private HanCancelado As Boolean
Private SePuedeSalir As Boolean

Private PorFTP As Boolean

Dim Im As cTimagen   'Nueva clase imagen

Dim Rs As ADODB.Recordset



'Para la copia real
Dim Orig As String, dest As String
Dim AlgunError As Boolean


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
                
        HacerAccion
        
        If PorFTP Then
            Me.Label2.Caption = "Cerrando conexion servidor"
            Me.Label2.Refresh
            
            'Cerramos conexion
            CerrarConexion
            
            Me.Refresh
            espera 2
            
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
    HanCancelado = False
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
    Inet1.Execute , "Close"
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






Private Sub Conectar()
  ' Si el control está ocupado, no realizar otra conexión
  If Inet1.StillExecuting = True Then Exit Sub
  ' Establecer las propiedades
  Inet1.URL = CarpetaD.SRV                  ' dirección URL
  Inet1.UserName = CarpetaD.user
  Inet1.Password = CarpetaD.pwd
  Inet1.Protocol = icFTP                    ' protocolo FTP
  Inet1.RequestTimeout = 50                 ' segundos
  PorFTP = True

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
      'Debug.Print " --> " & strDatos
      Label1.Caption = "Petición completada con éxito. " & _
                "Se recibieron todos los datos."
   
     TrayendoFichero = False
     FicheroOK = True
  End Select
  Label1.Refresh
End Sub





Private Sub HacerAccion()
On Error GoTo EHacerAccion


    Command1.Visible = True
    Me.Refresh
    If opcion = 1 Then
        ErrorLlevando = Not LlevarArchivos
    
    End If

    Exit Sub
EHacerAccion:
    MsgBox Err.Description
    'Habra que poner una variabale a mal
End Sub







Private Function InsertarBDImagen(ByRef Rs As ADODB.Recordset) As Boolean
Dim Cur As Currency
    Set Im = New cTimagen
    Im.campo1 = Rs!campo1
    Im.campo2 = Rs!campo2
    Im.campo3 = DBLet(Rs!campo3, "T")
    Im.campo4 = "Integracion Ariadna "
    Im.fecha1 = Rs!fecha1
    Im.fecha2 = Format(Now, "dd/mm/yyyy")
    Im.fecha3 = Format(Now, "dd/mm/yyyy")
    Im.lecturag = vbPermisoTotal
    Im.lecturau = vbPermisoTotal
    Im.codcarpeta = Rs!codcarpeta
    Im.codext = vConfig.ExtensionArchivos
    'Im.codigo = Rs!codigo
    Im.escriturag = vbPermisoTotal
    Im.escriturau = vbPermisoTotal
    Im.groupprop = 1
    Im.observa = DBLet(Rs!observa, "T")
    Cur = Rs!tamanyo
    Cur = Round(Cur / 1024, 2)
    If Cur = 0 Then Cur = 0.01
    Im.tamnyo = Cur
    Im.userprop = 0 'root
    If Im.Agregar(False) = 1 Then
        InsertarBDImagen = False
    Else
        InsertarBDImagen = True
    End If
    
End Function

Private Function LlevarArchivos() As Boolean
Dim Leida As Boolean
Dim CarpetaCreada As Boolean
Dim CarpErr As String
Dim Mal As Boolean


    'Abrimos el RS
    LlevarArchivos = False
    
    Set Rs = New ADODB.Recordset
    strDatos = "Select * from tmpintegra where codusu =" & CodPC
    Rs.Open strDatos, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Leida = False
    Set CarpetaD = New Ccarpetas
    CarpetaCreada = False
    While Not Rs.EOF
        If CarpetaD.codcarpeta <> Rs!codcarpeta Then
             If CarpetaD.Leer(Rs!codcarpeta, False) = 1 Then
                MsgBox "Error leyenod datos carpeta: " & Rs!Carpeta & " - ARIDOC: " & Rs!codcarpeta, vbExclamation
                Rs.Close
                Exit Function
            End If
            Leida = False
        End If
        If Not Leida Then
                If CarpetaD.version = 0 Then
                    Conectar
                    'Situar cd en srv
                    CambiaDirectorioFTP2 CarpetaD.pathreal
            Else
                ChDir vConfig.PathArchivos
            End If
            Leida = True
        End If
        
    
        Me.Label2.Caption = Rs!NombreArchivo
        Me.Label2.Refresh
        Orig = vConfig.PathArchivos & "\" & Rs!NombreArchivo
        
        Mal = True
        If InsertarBDImagen(Rs) Then
            dest = Im.codigo
            If Not CopiarArchivo Then
                Im.Eliminar
            Else
                Mal = False
            End If
        Else
            
        End If
        If Mal Then
            If Not CarpetaCreada Then
                'Crear carpeta
                CarpErr = CrearCarpetaErrores(Fich)
                CarpetaCreada = True
                'Lo primero k hago es copiar 1 archivo, el
                CopiarArchivoGuia CarpErr
                
            End If
            dest = CarpErr & "\" & Rs!NombreArchivo
            AccionesPorError
        End If
        DoEvents
        If HanCancelado Then Rs.MoveLast
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    If Not CarpetaCreada Then LlevarArchivos = True
    
End Function
Private Sub CopiarArchivoGuia(C As String)
    On Error Resume Next
    FileCopy vConfig.PathArchivos & "\" & Fich & ".ira", C & "\" & Fich
    If Err.Number <> 0 Then
        MsgBox "Error copiando archivo guia", vbExclamation
    End If
End Sub
Private Sub AccionesPorError()
Dim D As String
    On Error GoTo EEE
            D = "LLevando archivo : " & Orig & " --> " & dest
            FileCopy Orig, dest
            D = "Eliminando : " & Orig
            Kill Orig
    Exit Sub
EEE:
    MsgBox D & vbCrLf & Err.Description, vbExclamation
End Sub
Private Function CopiarArchivo() As Boolean
    On Error GoTo ECopiarArchivo
    
    CopiarArchivo = False
    If CarpetaD.version = 1 Then
        'MSDOS
        FileCopy Orig, CarpetaD.pathreal & "\" & dest
        EliminaArchivo Orig
    Else
        CopiaFTP
    End If
    CopiarArchivo = True
    Exit Function
ECopiarArchivo:
    Err.Clear
    Debug.Print Orig
        
End Function



        
        

Private Function CopiaFTP()
Dim Cad As String
        Do Until Not Inet1.StillExecuting
            DoEvents
        Loop
        
        'le quito la primera barra
       ' Cad = vDestino.pathreal & "/" & Destino
        Cad = "PUT """ & Orig & """ " & dest
        Inet1.Execute , Cad
        ' Esperar a que se establezca la conexión
        Do Until Not Inet1.StillExecuting
            DoEvents
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
                Inet1.Execute , "cd " & Aux
                Do Until Not Inet1.StillExecuting
                    DoEvents
                Loop
            
              
            
                Inet1.Execute , "pwd"
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
    'MuestraError Err.Number, "CambiaDirectorio" & Err.Description
    
End Function

