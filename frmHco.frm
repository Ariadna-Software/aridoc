VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmHco 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   Icon            =   "frmHco.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameAccion 
      Height          =   1815
      Left            =   840
      TabIndex        =   16
      Top             =   960
      Width           =   6015
      Begin ComCtl2.Animation Animation1 
         Height          =   615
         Left            =   360
         TabIndex        =   17
         Top             =   840
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   1085
         _Version        =   327681
         FullWidth       =   361
         FullHeight      =   41
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   615
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   5655
      End
   End
   Begin VB.Frame FrameRecuperar 
      Height          =   4215
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   7455
      Begin VB.CheckBox Check4 
         Caption         =   "Formato antiguo"
         Height          =   255
         Left            =   5280
         TabIndex        =   21
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Cerrar"
         Height          =   375
         Left            =   5640
         TabIndex        =   20
         Top             =   3600
         Width           =   1455
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3135
         Left            =   240
         TabIndex        =   15
         Top             =   960
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   5530
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   6526
         EndProperty
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   240
         TabIndex        =   14
         Text            =   "N:\Usuarios\david\PruebaHCO"
         Top             =   480
         Width           =   6495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Recuperar"
         Height          =   375
         Left            =   5640
         TabIndex        =   13
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Image Image3 
         Height          =   240
         Left            =   1680
         Picture         =   "frmHco.frx":030A
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "PATH del historico"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame FrameCrear 
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   360
         TabIndex        =   10
         Text            =   "Text2"
         Top             =   480
         Width           =   6855
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Poner propietario / grupo"
         Height          =   195
         Left            =   3240
         TabIndex        =   9
         Top             =   2040
         Width           =   2895
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   4800
         TabIndex        =   8
         Top             =   3600
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Salir"
         Height          =   375
         Index           =   0
         Left            =   6000
         TabIndex        =   7
         Top             =   3600
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Eliminar archivos en origen"
         Enabled         =   0   'False
         Height          =   195
         Left            =   360
         TabIndex        =   6
         Top             =   2040
         Width           =   4575
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   360
         TabIndex        =   5
         Text            =   "N:\Usuarios\david\PruebaHCO"
         Top             =   3000
         Width           =   6855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Habilitar permisos de lectura para todos los usuarios"
         Height          =   195
         Left            =   360
         TabIndex        =   3
         Top             =   1680
         Width           =   4575
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1200
         Width           =   6855
      End
      Begin VB.Label Label1 
         Caption         =   "Descripcion"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   1200
         Picture         =   "frmHco.frx":040C
         Top             =   2760
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   1200
         Picture         =   "frmHco.frx":050E
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "DESTINO"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   4
         Top             =   2760
         Width           =   855
      End
      Begin VB.Line Line1 
         X1              =   360
         X2              =   7200
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Label Label1 
         Caption         =   "ORIGEN"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   960
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmHco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcion As Byte
    '0.- Crear
    '1.- Recuperar
    
    
Private Conn1 As Connection




Private Sub cmdCancelar_Click(Index As Integer)
    If Index = 0 Then
        Unload Me
    End If
End Sub

Private Sub Command1_Click()
Dim EspacioCarpetas As Currency
Dim C As String
Dim i As Integer
Dim Carpeta As String
Dim Nombre As String


    If Not DatosOk Then Exit Sub


    Nombre = Format(Now, "yyddmmhhmm")
    'Comprobaciones
    'Fichero YA existe
    If Dir(Text1(1).Text & "\" & Nombre & ".hco", vbArchive) <> "" Then
        MsgBox "Ya existe un HCO: " & Nombre, vbExclamation
        Exit Sub
    End If

    
    'Comprobamos espacio en disco.....
    'Primero veremos todas las carpetas que hay que trasapasar
    EspacioCarpetas = 0
    DatosCopiados = Text1(0).Tag
    Set miRSAux = New ADODB.Recordset
    While DatosCopiados <> ""
        i = InStr(1, DatosCopiados, "|")
        If i > 0 Then
            C = Mid(DatosCopiados, 1, i - 1)
            DatosCopiados = Mid(DatosCopiados, i + 1)
            EspacioCarpetas = EspacioCarpetas + DevEspacioCarpeta(C)
        Else
            DatosCopiados = ""
        End If
    Wend
    Set miRSAux = Nothing
    
'    Dim F, Fs
'    Set Fs = CreateObject("Scripting.FileSystemObject")
'    Set F = Fs.GetFolder(Text1(1).Text)
    
    
    If EspacioCarpetas = 0 Then
        If MsgBox("Error obteniendo tamaño necesario. ¿Continuar?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    Else
        If MsgBox("Espacio aproximado requerido: " & EspacioCarpetas & ".  ¿Continuar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    
    'Procedemos al traspaso de datos
    Carpeta = Text1(1).Text & "\" & Nombre
    If crearCarpetaAlmacen(Carpeta) Then
    
        'Ya puedo poner en marcha el video, el avi
        Label2.Caption = "Comienzo proceso"
        Video True
        Me.FrameCrear.Enabled = False
        Me.FrameAccion.Visible = True
        Me.Refresh
        
        'Meto en temporal los archivos que hay que trasapasar
        OpcionesTablasHco 0
        
        
        'InsertarDatosEnTemporal
        DatosCopiados = Text1(0).Tag
        While DatosCopiados <> ""
            i = InStr(1, DatosCopiados, "|")
            If i > 0 Then
                C = Mid(DatosCopiados, 1, i - 1)
                DatosCopiados = Mid(DatosCopiados, i + 1)
                C = "INSERT INTO timagenhco SELECT " & vUsu.PC & ",TIMAGEN.* from TIMAGEN where codcarpeta = " & C
                Conn.Execute C
                
            Else
                DatosCopiados = ""
            End If
        Wend
        
        
        'AHora, si quito los permisos de lectura escritura
        
        If Check1.Value = 1 Then
            'Pongo lectura a todo el mundo
            C = "UPDATE timagenhco set lecturag =" & vbPermisoTotal & " WHERE codequipo = " & vUsu.PC
            Conn.Execute C
        End If
        
        If Check3.Value = 1 Then
            'Pongo usuario propietario a root
            C = "UPDATE timagenhco set userprop =" & vUsu.codusu & " , groupprop = " & vUsu.GrupoPpal & " WHERE codequipo = " & vUsu.PC
            Conn.Execute C
        End If
        
        
        'Llegados aqui, traspasamos los ficheros
        Label2.Caption = "Traspasando"
        Label2.Refresh
        If FicherosTraspaso(i) Then
            TraspasarFicheros i, Carpeta
            
            Close #i
            
            Label2.Caption = "Cerrando traspaso"
            Label2.Refresh
            
            'Copiamos el fichero
            FileCopy App.Path & "\vhco.dat", Text1(1).Text & "\" & Nombre & ".dat"
                    

                    
            'AHora creamos el archivo con todos los
            CrearFicheroDatos Nombre, EspacioCarpetas
            
            Label2.Caption = "Finalizando proceso"
            Label2.Refresh
            espera 1
        End If
    End If
    Me.FrameCrear.Enabled = True
    Me.FrameAccion.Visible = False
    Me.Refresh
    Screen.MousePointer = vbDefault
    
End Sub

Private Function DatosOk() As Boolean
    DatosOk = False
    Text1(0).Text = Trim(Text1(0).Text)
    Text1(1).Text = Trim(Text1(1).Text)
    Text2.Text = Trim(Text2.Text)
    If Text1(0).Text = "" Or Text1(0).Text = "" Or Text2.Text = "" Then
        MsgBox "Campos requeridos", vbExclamation
        Exit Function
    End If
    
    'la carpeta destino
    If Dir(Text1(1).Text, vbDirectory) = "" Then
        MsgBox "No existe carpeta destino", vbExclamation
        Exit Function
    End If
    DatosOk = True
    Exit Function
ED:
    MuestraError Err.Number, "DatosOk"
End Function

Private Sub Video(Encender As Boolean)
  On Error GoTo EVideo
    If Encender Then
        Me.Animation1.Open App.Path & "\Imagenes\Filemove.avi"
        Me.Animation1.Play -1
    Else
        Me.Animation1.Stop
        Me.Animation1.Close
    End If
    Exit Sub
EVideo:
    MuestraError Err.Number, "Poner animacion"
End Sub
Private Sub CrearFicheroDatos(ByRef Nom As String, EspacioAproximado As Currency)
Dim i As Integer
Dim C As String

    On Error GoTo ECre
    C = App.Path & "\cabhco.dat"
    If Dir(C, vbArchive) <> "" Then Kill C
    i = FreeFile
    Open C For Output As #i
    C = "## HCO de Aridoc: " & Text2.Text
    Print #i, C
    'AHora, el resto de lineas va codificado
    'para que no se lea del todo bien
    C = "Fecha traspaso: " & Format(Now, "dd/mm/yyyy hh:nn")
    CodificacionLinea False, C
    Print #i, C
    Print #i, ""
    
    C = "Carpetas Inicio : " & Text1(0).Text
    CodificacionLinea False, C
    Print #i, C
    Print #i, ""
    
    
    'Carpetas seleccionadas
    C = "Carpetas selecionadas : " & Text1(0).Tag
    CodificacionLinea False, C
    Print #i, C
    Print #i, ""
    
    C = "Ficheros traspasados : " & listacod.Count
    CodificacionLinea False, C
    Print #i, C
    Print #i, ""
    
    C = "Espacio aproximado : " & EspacioAproximado
    CodificacionLinea False, C
    Print #i, C
    Print #i, ""
    
        
    
    Close #1
    espera 1
    FileCopy App.Path & "\cabhco.dat", Text1(1).Text & "\" & Nom & ".hco"
    Exit Sub
ECre:
    MuestraError Err.Number, "Fichero final de configuracion de Historico"
End Sub
Private Sub TraspasarFicheros(NF As Integer, ByRef NomCarpeta As String)
Dim cad As String
Dim IZQ As String
Dim Der As String
Dim J As Integer
Dim Car As Ccarpetas

    Set miRSAux = New ADODB.Recordset
    cad = "Select * from timagenhco where codequipo = " & vUsu.PC
    cad = cad & " ORDER By codcarpeta,codigo"
    miRSAux.Open cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    
'    Cad = "DELETE FROM tmpbusqueda where codusu =" & vUsu.codusu & " AND codequipo = " & vUsu.PC
'    Conn.Execute Cad
    BorrarTemporal1
    
    'Obtengo el INSERT
    BACKUP_TablaIzquierda miRSAux, IZQ
    Set Car = New Ccarpetas
    Set listacod = New Collection
    'Para que leea a la primera
    Car.codcarpeta = -1
    
    Set listaimpresion = New Collection
    While Not miRSAux.EOF
    
        If Car.codcarpeta <> miRSAux!codcarpeta Then
            If Car.codcarpeta >= 0 Then
                If listaimpresion.Count >= 0 Then LlevarArchivo Car, NomCarpeta, IZQ, NF
            End If
            If Car.Leer(miRSAux!codcarpeta, (ModoTrabajo = 1)) = 1 Then
                'ERROR graaave
                MsgBox "Error grave. Leer=1. Carpeta:" & miRSAux!codcarpeta, vbExclamation
                
            Else
            '    Der = "INSERT INTO Carpetashco SELECT " & vUsu.PC & ",Carpetas.* From Carpetas WHERE codcarpeta =" & Car.codcarpeta
            '   Conn.Execute Der

            End If
        End If
        
        Label2.Caption = "Carpeta: " & Car.Nombre & " - Re: " & miRSAux!codigo
        Label2.Refresh
        
        
        BACKUP_Tabla miRSAux, Der
        listaimpresion.Add Der & "||" & CStr(miRSAux!codigo)
        
        'Cada 30 archvios los llevamos al destino. Para no abrir y cerrar tantas veces
        If listaimpresion.Count > 30 Then LlevarArchivo Car, NomCarpeta, IZQ, NF
            
        miRSAux.MoveNext
    Wend
    miRSAux.Close
    If listaimpresion.Count > 0 Then LlevarArchivo Car, NomCarpeta, IZQ, NF
    Set listaimpresion = Nothing
    
    Print #NF, ""
    Print #NF, ""
    
    
    'Creo la estructura de carpetas en hco
    '---------------------------------------
    DatosCopiados = Text1(0).Tag
    BorrarTemporal1
    Set miRSAux = New ADODB.Recordset
    While DatosCopiados <> ""
        J = InStr(1, DatosCopiados, "|")
        If J > 0 Then
            Der = Mid(DatosCopiados, 1, J - 1)
            DatosCopiados = Mid(DatosCopiados, J + 1)
            InsertaTemporal Val(Der)
        Else
            DatosCopiados = ""
        End If
    Wend
 
    'Insertamos la carpeta RAIZ
    InsertaTemporal 1
    
    
    
    Der = "INSERT INTO carpetashco  SELECT " & vUsu.PC & ",carpetas.* FROM carpetas,tmpFich WHERE carpetas.codcarpeta= tmpFich.imagen"
    Der = Der & " AND codusu = " & vUsu.codusu & " AND codequipo = " & vUsu.PC
    Conn.Execute Der
    
    'PRIMERa carpeta le pongo a raiz-> almacen
    'Y le pongo como padre de la primera carpeta a 1
    Der = "UPDATE carpetashco SET nombre='HCO: " & Format(Now, "dd/mm/yyyy") & "' WHERE codcarpeta =1"
    Conn.Execute Der
    
    J = InStr(1, Text1(0).Tag, "|")
    If J > 0 Then
        Der = Mid(Text1(0).Tag, 1, J - 1)
        Der = "UPDATE carpetashco set padre=1 where codcarpeta = " & Der
        Conn.Execute Der
    End If
    'AHora traspaso las carpetas
    'Primero. Actaulizo el almacen a 200
    cad = "UPDATE carpetashco SET almacen=200 where codequipo = " & vUsu.PC
    Conn.Execute cad
    
    If Check1.Value = 1 Then
        'Pongo lectura a todo el mundo
        cad = "UPDATE carpetashco set lecturag =" & vbPermisoTotal & " WHERE codequipo = " & vUsu.PC
        Conn.Execute cad
    End If
    
    If Check3.Value = 1 Then
        'Pongo usuario propietario a root
        cad = "UPDATE carpetashco set userprop =" & vUsu.codusu & " , groupprop = " & vUsu.GrupoPpal & " WHERE codequipo = " & vUsu.PC
        Conn.Execute cad
    End If
    
    
    
    
    cad = "Select * from carpetashco where codequipo = " & vUsu.PC
    miRSAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    BACKUP_TablaIzquierda miRSAux, IZQ
    While Not miRSAux.EOF
        BACKUP_Tabla miRSAux, Der
        'Le quito el primer codeuipo y se lo pongo en negativo
        PonerCodequipoNegativo True, Der
        Der = "INSERT INTO carpetashco " & IZQ & " VALUES" & Der & ";"
        CodificacionLinea False, Der
        Print #NF, Der
        miRSAux.MoveNext
    Wend
    miRSAux.Close
    
    
    
    'iNSERTO LA LINEA DE ALMACEN
    cad = "INSERT INTO almacenhco (codequipo, codalma, version, pathreal, SRV, user, pwd) VALUES ("
    cad = cad & "#N#,200,1,'','','','');"
    CodificacionLinea False, cad
    Print #NF, cad
    Set miRSAux = Nothing
End Sub



Private Sub PonerCodequipoNegativo(Escribir As Boolean, C As String)
Dim i As Integer
    If Escribir Then
        i = InStr(1, C, ",")
        C = "(#N#" & Mid(C, i)
    Else
        i = InStr(1, C, "#N#")
        C = Mid(C, 1, i - 1) & vUsu.PC & Mid(C, i + 3)
    End If
End Sub



Private Function FicherosTraspaso(ByRef NFF As Integer) As Boolean
On Error GoTo EFicherosTraspaso
    FicherosTraspaso = False
    If Dir(App.Path & "\vhco.dat", vbArchive) <> "" Then Kill App.Path & "\vhco.dat"
    NFF = FreeFile
    Open App.Path & "\vhco.dat" For Output As #NFF
    FicherosTraspaso = True
Exit Function
EFicherosTraspaso:
    MuestraError Err.Number, "Abriendo FicherosTraspaso"
End Function



Private Function LlevarArchivo(ByRef Ca As Ccarpetas, ByRef CarpetaDestino As String, IZQ As String, NFich As Integer) As Boolean
Dim R As ADODB.Recordset
Dim Der As String
Dim i As Integer
Dim J As Integer

        LlevarArchivo = False
        'Llevamos el fichero
        Set frmMovimientoArchivo.vOrigen = Ca
        frmMovimientoArchivo.Opcion = 16
        frmMovimientoArchivo.Destino = CarpetaDestino & "\"
        frmMovimientoArchivo.Show vbModal
        Me.Refresh
        Screen.MousePointer = vbHourglass
        espera 0.5
        'Una vez vuelve miramos un par de cositas
        If DatosCopiados <> "" Then
            While DatosCopiados <> ""
                i = InStr(1, DatosCopiados, "|")
                If i = 0 Then
                    DatosCopiados = ""
                Else
                    Der = Mid(DatosCopiados, 1, i - 1)
                    DatosCopiados = Mid(DatosCopiados, i + 1)
                    listaimpresion.Remove Val(Der)
                End If
            Wend
        End If
        
        For i = 1 To listaimpresion.Count
            Label2.Caption = "Llevando fichero SQL: " & i & " de " & listaimpresion.Count
            Label2.Refresh
            J = InStr(1, listaimpresion(i), "||")
            If J = 0 Then
                Me.Label2.Caption = "ERROR"
                espera 1
            Else
                Der = Mid(listaimpresion(i), 1, J - 1)
                PonerCodequipoNegativo True, Der
                Der = "INSERT INTO timagenhco " & IZQ & " VALUES" & Der & ";"
                CodificacionLinea False, Der
                Print #NFich, Der

                Der = Mid(listaimpresion(i), J + 2)
                listacod.Add Der
                Der = "DELETE FROM Timagenhco where codigo =" & Der
                Conn.Execute Der
                
                
            End If
        Next i
        Me.Refresh
        Set listaimpresion = Nothing
        Set listaimpresion = New Collection
End Function



Private Function crearCarpetaAlmacen(Ca As String) As Boolean
    On Error Resume Next
    MkDir Ca
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Creando carpeta"
        crearCarpetaAlmacen = False
    Else
        crearCarpetaAlmacen = True
    End If
End Function
Private Function DevEspacioCarpeta(ByRef codcarpeta As String) As Currency
Dim C As String
    DevEspacioCarpeta = 0
    C = "Select sum(tamnyo) from Timagen WHERE codcarpeta = " & codcarpeta
    miRSAux.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRSAux.EOF Then DevEspacioCarpeta = DBLet(miRSAux.Fields(0), "N")
    miRSAux.Close
    
    
End Function


Private Sub Command2_Click()
Dim T1 As Single
Dim TodoOk As Boolean
Dim CarRaiz As String
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    
   
    'Terminamos de leer el fichero y mostramos un msgbox
    If Check4.Value = 0 Then
        'Nueva forma de integracion
        If Not AbrirFicheroHco Then Exit Sub
    Else
        'Comprobar Datos de antiguo
        If Not CompruebaAntiguo(CarRaiz) Then Exit Sub
    End If
        
    T1 = Timer
    'Ya puedo poner en marcha el video, el avi
    Screen.MousePointer = vbHourglass
    Label2.Caption = "Comienzo recuperacion"
    
    Me.FrameRecuperar.Enabled = False
    Me.FrameAccion.Visible = True
    Video True
    Me.Refresh
    espera 1
    
    
    If Check4.Value = 0 Then
        'NUEVO
        TodoOk = RecuperarArchivo(Text3.Text & "\" & ListView1.SelectedItem.Tag)
    Else
        'ANTIGUO
        TodoOk = RecuperaAntiguo(CarRaiz)
    End If
    
    'Nuevo  19 Mayo 2006
    '---------------------
    AjustaTablaCarpetasHijos
    
    
    If TodoOk Then
        Screen.MousePointer = vbHourglass
        Label2.Caption = "Creando estructura hco: " & vUsu.PC
        Label2.Refresh
        T1 = Timer - T1
        T1 = 5 - T1
        If T1 > 0 Then espera T1
        Video False
        If Check4.Value = 1 Then
            'ES ANTIGUO
            DatosCopiados = "ANT"
        Else
            DatosCopiados = "RECUPERACION"
        End If
        Unload Me
    Else
        Video False
        Me.FrameRecuperar.Enabled = True
        Me.FrameAccion.Visible = False
        Me.Refresh
        Screen.MousePointer = vbDefault
    End If
    
End Sub

Private Function AbrirFicheroHco() As Boolean
Dim C As String
Dim NF As Integer

On Error GoTo EAbrirFicheroHco
    AbrirFicheroHco = False
    
    
    If ListView1.SelectedItem.Tag = "" Then
        MsgBox "Error en LISTVIEW", vbExclamation
        Exit Function
    End If
    
        
    C = Text3.Text & "\" & ListView1.SelectedItem.Tag & ".hco"
    If Dir(C, vbArchive) = "" Then
        MsgBox "Fichero no encontrado: " & C
        Exit Function
    End If

    C = Text3.Text & "\" & ListView1.SelectedItem.Tag & ".dat"
    If Dir(C, vbArchive) = "" Then
        MsgBox "Fichero no encontrado: " & C
        Exit Function
    End If

    'La carpeta
    C = Text3.Text & "\" & ListView1.SelectedItem.Tag
    If Dir(C, vbDirectory) = "" Then
        MsgBox "Carpeta no encontrada: " & C
        Exit Function
    End If



    'AHora abro el fichero hco y muestro el mensaje
    NF = FreeFile
    C = Text3.Text & "\" & ListView1.SelectedItem.Tag & ".hco"
    Open C For Input As #NF
    C = vbCrLf & vbCrLf & C
    C = "Va a recuperar el siguiente HCO: " & C & vbCrLf & vbCrLf
    DatosMOdificados = False
    While Not EOF(NF)
        Line Input #NF, DatosCopiados
        If DatosMOdificados Then
            CodificacionLinea True, DatosCopiados
            C = C & DatosCopiados & vbCrLf
        Else
            DatosMOdificados = True
        End If

    Wend
    Close #NF
    DatosMOdificados = False
    C = C & " ¿CONTINUAR?"
    If MsgBox(C, vbQuestion + vbYesNoCancel + vbDefaultButton2) <> vbYes Then Exit Function

    AbrirFicheroHco = True
    Exit Function
EAbrirFicheroHco:
    MuestraError Err.Number, "AbrirFicheroHco"
End Function

Private Sub Command3_Click()
    DatosCopiados = ""
    Unload Me
End Sub

Private Sub Form_Load()
    Me.FrameAccion.Visible = False
    Me.FrameCrear.Visible = False
    Me.FrameRecuperar.Visible = False
    
    If Opcion = 0 Then
        Me.FrameCrear.Visible = True
        Caption = "Generar"
    Else
        Me.FrameRecuperar.Visible = True
        ListView1.ListItems.Clear
        Caption = "Recuperar"
    End If
    Caption = Caption & " histórico"
    Limpiar Me
    DatosCopiados = ""
End Sub



Private Sub Image1_Click()
Dim i As Integer
    frmPregunta.Opcion = 20
    frmPregunta.origenDestino = 1
    DatosCopiados = ""
    frmPregunta.Show vbModal
    If DatosCopiados <> "" Then
        i = InStr(1, DatosCopiados, "·")
        If i = 0 Then
            Text1(0).Text = ""
        Else
            Text1(0).Text = Mid(DatosCopiados, 1, i - 1)
            Text1(0).Tag = Mid(DatosCopiados, i + 1)
        End If
    End If
        
End Sub

Private Sub Image2_Click()
Dim C As String
    C = GetFolder("Carpeta destino")
    If C <> "" Then Me.Text1(1).Text = C
End Sub

Private Sub OpcionesTablasHco(Opc As Byte)

    Select Case Opc
    Case 0
        'ELIMINAR
        Conn.Execute "Delete from timagenhco where codequipo = " & vUsu.PC
         
        Conn.Execute "Delete from carpetashco where codequipo = " & vUsu.PC
        
        Conn.Execute "Delete from almacenhco where codequipo = " & vUsu.PC


    End Select
End Sub

Private Sub CambiarCarpeta()
Dim C As String
Dim N As String
Dim It As ListItem

    ListView1.ListItems.Clear
    If Text3.Text = "" Then Exit Sub
    If Check4.Value = 1 Then
        '----------------------------------
        ' ANTIGUO ARIDOC
        C = Dir(Text3.Text & "\ImgDatos.txt", vbArchive)
        If C <> "" Then
            Set It = ListView1.ListItems.Add
            It.Text = C
            It.Tag = ""
        End If
        
        
   
    
    
    Else
        C = Dir(Text3.Text & "\*.hco", vbArchive)
        Do
            If C <> "" Then
               N = LeerFicheroConfig(C)
               Set It = ListView1.ListItems.Add
               It.Text = N
               It.Tag = Mid(C, 1, InStr(1, C, ".hco") - 1)
               C = Dir
            End If
            
        Loop Until C = ""
    End If
End Sub


Private Function LeerFicheroConfig(Fich As String)
Dim i As Integer
Dim C As String
    
    On Error GoTo ELeer
    i = FreeFile
    LeerFicheroConfig = Fich
    Open Text3.Text & "\" & Fich For Input As i
    Line Input #i, C
    
    Close #i
    'Primera leina
    'C = "## HCO de Aridoc: "
    LeerFicheroConfig = Mid(C, 19)
    Exit Function
ELeer:
    MuestraError Err.Number, "Leer fichero"
End Function

Private Function RecuperarArchivo(Nombre As String) As Boolean
Dim NF As Integer
Dim C As String
Dim N As Long

    OpcionesTablasHco 0
    Label2.Caption = "Recuperando datos fichero"
    
    'Abrimos el fichero en .dat y empezamos a insertar
    NF = FreeFile
    Open Nombre & ".dat" For Input As NF
    N = 0
    While Not EOF(NF)
        N = N + 1
        Label2.Caption = "Leyendo : " & N
        Label2.Refresh
        
        Line Input #NF, C
        CodificacionLinea True, C
        C = Trim(C)
        If C <> "" Then
            PonerCodequipoNegativo False, C
            Conn.Execute C
        End If
    Wend
    Close #NF
    
    
    'Actualizamos la BD de hco
    Label2.Caption = "Procesando carpetas"
    Label2.Refresh
    C = CStr(Nombre)
    NombreSQL C
    C = "UPDATE almacenhco SET pathreal = '" & C & "' WHERE codequipo = " & vUsu.PC
    Conn.Execute C
    
    Set miRSAux = New ADODB.Recordset
    C = "Select * from carpetashco where codequipo = " & vUsu.PC
    miRSAux.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRSAux.EOF Then N = DBLet(miRSAux.Fields(0), "N")
    miRSAux.Close
    
    
    
    Set miRSAux = Nothing
    
    
    If N > 0 Then
        RecuperarArchivo = True
    Else
        MsgBox "No han sido devueltos los datos de las carpetas", vbExclamation
        RecuperarArchivo = False
    End If
        
End Function

Private Sub Image3_Click()
Dim C As String
    C = GetFolder("path del historico")
    If C <> "" Then
        Screen.MousePointer = vbHourglass
        Text3.Text = C
        CambiarCarpeta
        Screen.MousePointer = vbDefault
    End If
    
End Sub

Private Function CompruebaAntiguo(ByRef CarpetaRaiz As String) As Boolean
Dim NF As Integer
Dim S As String
Dim C As String

On Error GoTo ECompruebaAntiguo
    CompruebaAntiguo = False
    
    If ListView1.SelectedItem.Tag <> "" Then
        MsgBox "Error en LISTVIEW", vbExclamation
        Exit Function
    End If
    
    If Dir(Text3.Text & "\" & ListView1.SelectedItem.Text, vbArchive) = "" Then
        MsgBox "No se ha encontrado el archivo", vbExclamation
        Exit Function
    End If


    'BD
    If Dir(Text3.Text & "\Historia.mdb", vbArchive) = "" Then
        MsgBox "No se ha encontrado el archivo BD.", vbExclamation
        Exit Function
    End If
    
    
    'Igual la carpeta no se llama RAIZ
    If Not CompruebaNombreCarpetaRaiz(CarpetaRaiz) Then
        MsgBox "Error carpeta DATOS", vbExclamation
        Exit Function
    End If
    
    
    
    'Abrimos el fichero imgdatos y lo mostramos
    C = Text3.Text & "\" & ListView1.SelectedItem.Text
    NF = FreeFile
    Open C For Input As #NF
    S = "Va a recuperar los datos: " & vbCrLf
    While Not EOF(NF)
        Line Input #NF, C
        S = S & C & vbCrLf
    Wend
    Close #NF
    S = S & vbCrLf & vbCrLf & " DESEA CONTINUAR?"
    If MsgBox(S, vbQuestion + vbYesNoCancel + vbDefaultButton3) <> vbYes Then Exit Function
    
    CompruebaAntiguo = True
    Exit Function
ECompruebaAntiguo:
    MuestraError Err.Number, "CompruebaAntiguo"

End Function

'Devolvera el mensaje de error
Private Function CompruebaNombreCarpetaRaiz(ByRef Carpeta As String) As Boolean
Dim C As String
Dim i As Integer
    
    i = 0
    CompruebaNombreCarpetaRaiz = False ' "No existen la subcarpeta para " & Text3.Text
    C = Dir(Text3.Text & "\*.", vbDirectory)
    
    Do
        
            If C <> "." And C <> ".." Then
                i = i + 1
                
                If i > 1 Then
                    If MsgBox("Hay mas de una carpeta. La carpeta raiz es: " & Carpeta & " ?", vbQuestion + vbYesNo) = vbYes Then
                        C = ""
                        CompruebaNombreCarpetaRaiz = True
                        Exit Function
                    End If
                End If
                Carpeta = C
            End If
            C = Dir()
    Loop Until C = ""
    If i = 1 Then
        If Carpeta <> "" Then CompruebaNombreCarpetaRaiz = True
    Else
        If MsgBox(" La carpeta raiz es: " & Carpeta & " ?", vbQuestion + vbYesNo) = vbYes Then CompruebaNombreCarpetaRaiz = True
    End If
End Function


Private Function RecuperaAntiguo(ByRef LaCarpetaRaiz As String) As Boolean
Dim codcarpeta As Integer
On Error GoTo eRecuperaAntiguo
    RecuperaAntiguo = False
    'MsgBox "TODAVIA NO DISPONIBLE", vbExclamation
    'Exit Function
    '---------------------------------
    Label2.Caption = "preparando datos"
    Me.Refresh
    OpcionesTablasHco 0
    
    
    
    
    'Recupero todas las entradas de los archivos sobre las tablas de hco
    If Not LeerDatosTimagen Then Exit Function
    
    'Los datos se montaran por carpetas
    'Carpeta RAIZ
    codcarpeta = 1
    If ProcesaCarpeta(codcarpeta, Text3.Text, LaCarpetaRaiz, 0) Then RecuperaAntiguo = True
    
    
    
        
    
    
    Exit Function
eRecuperaAntiguo:
    MuestraError Err.Number, "Recupera Antiguo"
End Function


Private Sub AjustaTablaCarpetasHijos()
Dim C As String

    Label2.Caption = "Ajuste tabla carphco"
    Me.Refresh
    espera 0.5
    C = "Select padre from carpetashco group by padre"
    Set miRSAux = New ADODB.Recordset
    miRSAux.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    C = "UPDATE carpetashco set hijos =1 where codequipo =" & vUsu.PC & " AND  codcarpeta ="
    While Not miRSAux.EOF
        Conn.Execute C & miRSAux.Fields(0)
        miRSAux.MoveNext
    Wend
    miRSAux.Close
    Set miRSAux = Nothing
    
End Sub


Private Function ProcesaCarpeta(ByRef codCa As Integer, ruta As String, Carpeta As String, padre As Integer) As Boolean
Dim Actual As Integer
Dim cad As String
Dim Carpetas As String
Dim J As Integer
Dim i As Long
Dim Tama As Long
Dim Aux As String

    ProcesaCarpeta = False
    Label2.Caption = "Carpeta: " & ruta & "  (" & codCa & ")"
    Label2.Refresh
    Actual = codCa
    If Not InsertaCarpetaAlmacen(codCa, Carpeta, padre, ruta) Then Exit Function

    'Si tiene archivos
    cad = Dir(ruta & "\" & Carpeta & "\*.*", vbArchive)
    Do
        If cad <> "" Then
            Tama = FileLen(ruta & "\" & Carpeta & "\" & cad)
            Aux = "UPDATE timagenhco Set tamnyo=" & TransformaComasPuntos(CStr(Round((Tama / 1024), 2)))
            Aux = Aux & ",codcarpeta = " & Actual & " WHERE codequipo = " & vUsu.PC & " AND  codigo ="
            i = InStr(1, cad, ".")
            If i = 0 Then
                Aux = Aux & cad
            Else
                Aux = Aux & Mid(cad, 1, i - 1)
            End If
            Conn.Execute Aux
            cad = Dir
         End If
         
     Loop Until cad = ""





    'Si tiene carpetas
    Label2.Caption = "Leyendo subcarpetas en : " & ruta & "\" & Carpeta & "\"
    Label2.Refresh
    cad = Dir(ruta & "\" & Carpeta & "\", vbDirectory)
    Carpetas = ""
    J = 0
    Do
        If cad <> "" Then
               If cad <> "." And cad <> ".." Then
                    If (GetAttr(ruta & "\" & Carpeta & "\" & cad) And vbDirectory) = vbDirectory Then
                        Carpetas = Carpetas & cad & "|"
                        J = J + 1
                    End If
                End If
        End If
        cad = Dir
    Loop Until cad = ""

    cad = ""
    If J = 0 Then
        ProcesaCarpeta = True
    Else
        For i = 1 To J
                cad = RecuperaValor(Carpetas, CInt(i))
                If Not ProcesaCarpeta(codCa, ruta & "\" & Carpeta, cad, Actual) Then
                    cad = ""
                    Exit For
                End If
        Next i
        If cad <> "" Then ProcesaCarpeta = True
    End If
End Function

'METEREMOS UN ALMACEN Y UNA CARPETA
Private Function InsertaCarpetaAlmacen(ByRef codcarpeta As Integer, Carpeta As String, padre As Integer, ruta As String) As Boolean
Dim C As String
    On Error GoTo EI
    InsertaCarpetaAlmacen = False
    C = "INSERT INTO carpetashco (codequipo, codcarpeta, nombre, padre, userprop, "
    C = C & "almacen, groupprop, lecturau, lecturag, escriturau, escriturag) VALUES ("
    C = C & vUsu.PC & "," & codcarpeta & ",'" & DevNombreSql(Carpeta) & "'," & padre & ",0,"
    C = C & codcarpeta & ",1," & vbPermisoTotal & "," & vbPermisoTotal & "," & vbPermisoTotal & "," & vbPermisoTotal & ")"
    Conn.Execute C
    
    C = "INSERT INTO almacenhco (codequipo, codalma, version, pathreal, SRV, user, pwd) VALUES ("
    C = C & vUsu.PC & "," & codcarpeta & ",1,'" & DevNombreSql(ruta & "\" & Carpeta) & "','',NULL,NULL)"
    
    Conn.Execute C
    codcarpeta = codcarpeta + 1
    InsertaCarpetaAlmacen = True
    Exit Function
EI:
    MuestraError Err.Number, "Insertando carpeta almacen" & vbCrLf & C
End Function



Private Function LeerDatosTimagen() As Boolean
    LeerDatosTimagen = False
    'Abro la conexion de AridocOLD
    If Not AbriBDHistoria Then Exit Function
    

    'Volcar registros
    VolcarRegistros


    'Cerramos la conexioon
    Conn1.Close
    Set Conn1 = Nothing
    LeerDatosTimagen = True
End Function


Private Function AbriBDHistoria() As Boolean

    AbriBDHistoria = False
    On Error Resume Next
    Set Conn1 = New ADODB.Connection
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Creando Objeto CONN ADODB.Connection"
        Exit Function
    End If
    Conn1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Text3.Text & "\historia.mdb;Persist Security Info=Fal"
    Conn1.CursorLocation = adUseServer
    Conn1.Open
    If Err.Number <> 0 Then
        MuestraError Err.Number, "abriendo BD Historia"
        Exit Function
    Else
        AbriBDHistoria = True
    End If
    
End Function


Private Sub VolcarRegistros()
Dim Insert As String
Dim SQL As String

    'codigo,codext,campo1, campo2,campo3,fecha1, fecha2,observa,
    Insert = "INSERT INTO timagenhco (codequipo, codcarpeta,  campo4,  fecha3, importe1, importe2, "
    Insert = Insert & "tamnyo, userprop, groupprop, lecturau, lecturag, escriturau, escriturag, bloqueo,"
    'Estos campos son los que se modifican relamente
    Insert = Insert & "codigo,codext,campo1, campo2,campo3,fecha1, fecha2,observa"
    Insert = Insert & ") VALUES (" & vUsu.PC & ",0,NULL,NULL,NULL,NULL,"
    Insert = Insert & "0," & vbPermisoTotal & "," & vbPermisoTotal & "," & vbPermisoTotal
    Insert = Insert & "," & vbPermisoTotal & "," & vbPermisoTotal & "," & vbPermisoTotal & ",0,"
    
    Set miRSAux = New ADODB.Recordset
    SQL = "Select * from Timagen"
    miRSAux.Open SQL, Conn1, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRSAux.EOF
        Label2.Caption = miRSAux!Id & " - " & miRSAux!clave1
        Label2.Refresh
    
        SQL = miRSAux!Id & "," & miRSAux!Extension & ",'" & DevNombreSql(miRSAux!clave1) & "',"
        'Campo2
        SQL = SQL & ParaBD(DBLet(miRSAux!clave2, "T"), "T", True) & ","
        SQL = SQL & ParaBD(DBLet(miRSAux!clave3, "T"), "T", True) & ",'"
        SQL = SQL & Format(miRSAux!fechadig, FormatoFecha) & "',"
        SQL = SQL & ParaBD(DBLet(miRSAux!fechadoc, "T"), "T", True) & ","
        SQL = SQL & ParaBD(DBLet(miRSAux!DES, "T"), "T", True) & ")"
        SQL = Insert & SQL
        Conn.Execute SQL
        miRSAux.MoveNext
    Wend
    miRSAux.Close
End Sub

Private Sub Text3_LostFocus()
    If Text3.Text = "" Then Me.ListView1.ListItems.Clear
End Sub
