VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmvFTP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Traspaso datos"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   Icon            =   "frmvFTP.frx":0000
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


'Public Usuario As String
'Public Password As String
'Public Destino As String
'Public Origen As String
'Public Servidor As String

Public Origen As String
Public CarpetaD As Ccarpetas


Private strDatos As String
Private PrimeraVez As Boolean
Private FicheroOK As Boolean
Private TrayendoFichero As Boolean
Private HanCancelado As Boolean
Private SePuedeSalir As Boolean

Private PorFTP As Boolean

Dim Im As cTimagen

Dim RS As ADODB.Recordset



'Para la copia real
Dim orig As String, dest As String

Private Sub Command1_Click()
    Label2.Caption = "Cancelando acciones"
    HanCancelado = True
    SeHaCancelado = True
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
            espera 1
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
        LlevarArchivos
    End If

    Exit Sub
EHacerAccion:
    MsgBox Err.Description
    'Habra que poner una variabale a mal
End Sub







Private Function InsertarBDImagen(ByRef RS As ADODB.Recordset) As Boolean
Dim Cur As Currency
    Set Im = New cTimagen
    Im.campo1 = RS!clave1
    Im.campo2 = RS!clave2
    Im.campo3 = RS!clave3
    Im.campo4 = "Traspaso al Nuevo "
    Im.fecha1 = RS!fechadoc
    Im.fecha2 = RS!FechaDig
    Im.fecha3 = Format(Now, "dd/mm/yyyy")
    Im.lecturag = vbPermisoTotal
    Im.lecturau = vbPermisoTotal
    Im.codcarpeta = RS!codcarpeta
    Im.codext = RS.Fields(8)  'EL 8 es el codextension en timagen
    Im.codigo = RS!codigo
    Im.escriturag = vbPermisoTotal
    Im.escriturau = vbPermisoTotal
    Im.groupprop = 1
    Im.observa = RS!des
    Cur = RS!tamayo
    Cur = Round(Cur / 1024, 2)
    If Cur = 0 Then Cur = 0.01
    Im.tamnyo = Cur
    Im.userprop = 0 'root
    If Im.Agregar(False) = 1 Then
        InsertarBDImagen = False
    Else
        InsertarBDImagen = True
    End If
    Set Im = Nothing
    
End Function

Private Function LlevarArchivos()
    'Cambiamos el directorio local
    ChDir (Origen)




    'Abrimos el RS
    Set RS = New ADODB.Recordset
    strDatos = "Select timagen.*,tablaespecial.* from timagen,tablaespecial"
    strDatos = strDatos & " WHERE timagen.id=tablaespecial.codigo AND codcarpeta =" & CarpetaD.codcarpeta
    RS.Open strDatos, ConnAntiguoAridoc, adOpenForwardOnly, adLockPessimistic, adCmdText
    If CarpetaD.version = 0 Then
        Conectar
        'Situar cd en srv
        CambiaDirectorioFTP2 CarpetaD.pathreal
    Else
        
    End If
    
    While Not RS.EOF
        Me.Label2.Caption = RS!nomfich
        Me.Label2.Refresh
        orig = RS!nomfich
        dest = RS!codigo
        If CopiarArchivo Then
            InsertarBDImagen RS
        End If

        DoEvents
        If HanCancelado Then RS.MoveLast
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
End Function

Private Function CopiarArchivo() As Boolean
    On Error GoTo ECopiarArchivo
    
    CopiarArchivo = False
    If CarpetaD.version = 1 Then
        'MSDOS
        FileCopy orig, CarpetaD.pathreal & "\" & dest
    Else
        CopiaFTP
    End If
    CopiarArchivo = True
    Exit Function
ECopiarArchivo:
    Err.Clear
    
        
End Function



        
        

Private Function CopiaFTP()
Dim Cad As String
        Do Until Not Inet1.StillExecuting
            DoEvents
        Loop
        
        'le quito la primera barra
       ' Cad = vDestino.pathreal & "/" & Destino
        Cad = "PUT """ & orig & """ " & dest
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

