VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmRecuperaBackup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recuperar archivos desde BAKCUP"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8445
   Icon            =   "frmRecuperaBackup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   8445
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameRecupera 
      Height          =   3735
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   8295
      Begin VB.CommandButton cmdRestaurar 
         Caption         =   "&Restaurar"
         Height          =   375
         Left            =   5160
         TabIndex        =   18
         Top             =   3240
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "C&omprobar"
         Height          =   375
         Left            =   5160
         TabIndex        =   17
         Top             =   3240
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   6720
         TabIndex        =   16
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Text            =   "Text3"
         Top             =   550
         Width           =   7935
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2175
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   3836
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "C"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre aridoc"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Ext. BK"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Ext Ari."
            Object.Width           =   1587
         EndProperty
      End
      Begin VB.Label Label5 
         Caption         =   "Label5"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   3240
         Width           =   4695
      End
      Begin VB.Label Label4 
         Caption         =   "Carpeta"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1815
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   7560
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   3735
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   8295
      Begin VB.CommandButton Command1 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   6840
         TabIndex        =   11
         Top             =   3240
         Width           =   1215
      End
      Begin VB.CommandButton cmdRegresar 
         Caption         =   "Regresar"
         Height          =   375
         Left            =   5520
         TabIndex        =   10
         Top             =   3240
         Width           =   1215
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2655
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   4683
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "PATH"
            Object.Width           =   10584
         EndProperty
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   0
         Left            =   1440
         Picture         =   "frmRecuperaBackup.frx":030A
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Asigna carpetas"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   210
         Width           =   1335
      End
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   360
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1200
      Width           =   7815
   End
   Begin VB.CommandButton cmdRecuperar 
      Caption         =   "Recuperar"
      Height          =   375
      Left            =   5280
      TabIndex        =   4
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar"
      Height          =   375
      Index           =   0
      Left            =   6840
      TabIndex        =   2
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   480
      Width           =   7815
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   2
      Left            =   1920
      Picture         =   "frmRecuperaBackup.frx":040C
      Top             =   240
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   1
      Left            =   2400
      Picture         =   "frmRecuperaBackup.frx":050E
      Top             =   960
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Carpetas archivo backup"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   6
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      TabIndex        =   3
      Top             =   1680
      Width           =   7815
   End
   Begin VB.Label Label1 
      Caption         =   "Path archivo DUMP"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmRecuperaBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcion As Byte
    ' 0.- Normal. Empezara recuperar archivos
    ' 1.- Recuperar archivos
    
Dim cad As String
Dim i As Integer

Private Car As Ccarpetas
Private Origen As String

Private Sub cmdCancelar_Click()
    Conn.Execute "DELETE FROM timagenhco"
    Set Car = Nothing
    Unload Me
End Sub



Private Sub cmdRecuperar_Click()
    If Text1.Text = "" Then
        MsgBox "Seleccione el archivo de BACKUP", vbExclamation
        Exit Sub
    End If
    
    If Text2.Text = "" Then
        cad = "No ha puesto el path desde el cual restaurar. ¿Continuar?"
        If MsgBox(cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    Else
        
    End If
    
    cad = "¿El proceso puede costar mucho tiempo. ¿Continuar?"
    If MsgBox(cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    Label3.Caption = "Comienzo proceso"
    Me.Refresh
    IniciarProceso
    Label3.Caption = ""
    Screen.MousePointer = vbDefault
End Sub



Private Sub cmdRegresar_Click()
    If ListView1.ListItems.Count = 0 Then Exit Sub
    
    For i = 1 To ListView1.ListItems.Count
        cad = ListView1.ListItems(i).SubItems(1)
        If Not CompruebaCarpeta Then Exit Sub
    Next i
        
    For i = 1 To ListView1.ListItems.Count
        cad = ListView1.ListItems(i).SubItems(1)
        cad = "UPDATE almacen Set pathreal = '" & pathToMySql(cad)
        cad = cad & "' WHERE codalma = " & ListView1.ListItems(i).Text
        Conn.Execute cad
    Next i
    DatosMOdificados = True
    Unload Me
End Sub


Private Function pathToMySql(cad As String) As String
Dim J As Integer
Dim K As Integer

    J = -1
    Do
       K = J + 2
       J = InStr(K, cad, "\")
       If J > 0 Then cad = Mid(cad, 1, J) & Mid(cad, J)
    Loop Until J = 0
    pathToMySql = cad
End Function

Private Function CompruebaCarpeta() As Boolean
    On Error Resume Next
    CompruebaCarpeta = False
    cad = Dir(cad, vbDirectory)
    If Err.Number <> 0 Then
        MuestraError Err.Number
    Else
        If cad = "" Then
            MsgBox "Carpeta NO encontrada: " & ListView1.ListItems(i).SubItems(1), vbExclamation
        Else
            CompruebaCarpeta = True
        End If
    End If
End Function

Private Sub cmdRestaurar_Click()


    cad = ""
    For i = 1 To ListView2.ListItems.Count
        If ListView2.ListItems(i).Checked Then
            cad = "OK"
            Exit For
        End If
    Next i
    If cad = "" Then
        MsgBox "Seleccione los archivos a recuperar", vbExclamation
        Exit Sub
    End If
            
    
    cad = ""
    For i = 1 To ListView2.ListItems.Count
        If ListView2.ListItems(i).Checked Then
            If Dir(Origen & "\" & Mid(ListView2.ListItems(i).Key, 2), vbArchive) = "" Then
                cad = cad & "  - " & ListView2.ListItems(i).Text & " (" & Mid(ListView2.ListItems(i).Key, 2) & ")" & vbCrLf
            End If
        End If
    Next i
    
    If cad <> "" Then
        cad = "Archivos NO encontrados en la copia de seguridad: " & vbCrLf & vbCrLf
        MsgBox cad, vbInformation
        Exit Sub
    End If
    
    
    For i = 1 To ListView2.ListItems.Count
        Label5.Caption = "Copiando : " & ListView2.ListItems(i).Text
        Me.Refresh
        
        'auqi
        DatosCopiados = "NO"
        frmMovimientoArchivo.Origen = Origen & "\" & Mid(ListView2.ListItems(i).Key, 2)
        Set frmMovimientoArchivo.vDestino = Car
        frmMovimientoArchivo.Destino = Mid(ListView2.ListItems(i).Key, 2)
        frmMovimientoArchivo.Opcion = 1
        frmMovimientoArchivo.Show vbModal
        
        If DatosCopiados = "" Then
            'HA SIDO COPIADO EL ARCHIVo
            'Ahora, elimino la entrada, y la vulevo a meter
            cad = "DELETE FROM ARIDOC.timagen where codigo = " & Mid(ListView2.ListItems(i).Key, 2)
            EjecutaSQL
            espera 0.5
            cad = "INSERT INTO Aridoc.timagen Select * from timagen where codigo =" & Mid(ListView2.ListItems(i).Key, 2)
            EjecutaSQL
        End If
    Next i
    
    
    
    Set Car = Nothing
    Unload Me
End Sub

Private Sub EjecutaSQL()
    On Error Resume Next
    Conn.Execute cad
    If Err.Number <> 0 Then
        cad = "ERROR SQL.          Avise soporte técnico. NO cierre la ventana" & vbCrLf & vbCrLf & cad & vbCrLf & vbCrLf
    End If
End Sub


Private Sub Command1_Click(Index As Integer)
    Unload Me
End Sub



Private Sub Command2_Click()


    On Error GoTo ERecu

    'COMPROBAMOS MUCHAS COSAS
    '-------------------------------------------------------------
    
    '1.- QUE LAS carpeta desde donde queremos copiar EXISTE y cuelga de
    '    la misma subcarpeta que en realidad
    
    cad = "SELECT carpetas.* ,almacen.* FROM "
    cad = cad & "aridoc.carpetas,aridoc.almacen"
    
    cad = cad & " WHERE carpetas.almacen = almacen.codalma AND "
    cad = cad & " codcarpeta = " & Mid(Text3.Tag, 2)
        
    Set miRSAux = New ADODB.Recordset
    miRSAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRSAux.EOF Then
        miRSAux.Close
        MsgBox "Carpeta NO encontrada en ARIDOC: " & Text3.Tag, vbExclamation
        Exit Sub
    End If

    cad = Mid(Text3.Text, InStrRev(Text3.Text, "\") + 1)
    If miRSAux!padre <> DatosCopiados Then
        MsgBox "Carpeta contenedora de " & cad & " distinta de la actual " & miRSAux!codcarpeta, vbExclamation
        miRSAux.Close
        Exit Sub
    End If
    
    
    If cad <> miRSAux!Nombre Then
        cad = "Nombre de la carpeta ha cambiado." & vbCrLf & vbCrLf & "Backup: " & cad & vbCrLf
        cad = cad & "Aridoc:  " & miRSAux!Nombre & vbCrLf & vbCrLf & "¿Continuar?"
        If MsgBox(cad, vbQuestion + vbYesNo) = vbNo Then
            miRSAux.Close
            Exit Sub
        End If
    End If
    
    Set Car = New Ccarpetas
    
    
    'Compruebo la carpeta origen donde se encontraran los datos
    If Car.Leer(CInt(Mid(Text3.Tag, 2)), False) = 1 Then
        miRSAux.Close
        Set Car = Nothing
    End If
    
    'Fijo el orgine para los que voy a llevar a ARIDOC
    Origen = Car.pathreal
    Set Car = Nothing
    Set Car = New Ccarpetas
    
    
    
    
    
    
    
    
    
    With Car
        .codcarpeta = miRSAux!codcarpeta
        .Nombre = miRSAux!Nombre
        .padre = miRSAux!padre
        .Almacen = miRSAux!Almacen
        .groupprop = miRSAux!groupprop
        'Establezco el ALMACEN
        .version = miRSAux!version
        .pathreal = miRSAux!pathreal
        .SRV = miRSAux!SRV
        .user = DBLet(miRSAux!user, "T")
        .pwd = DBLet(miRSAux!pwd, "T")
    End With
    miRSAux.Close

    'YA tenemos que la carpeta esta bien.
    'Ahora lo que tenemos que hacer es cargar el listivew
    ' con el resto de datos
    For i = 1 To ListView2.ListItems.Count
        cad = "Select * from aridoc.timagen where codigo =" & Mid(ListView2.ListItems(i).Key, 2)
        miRSAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not miRSAux.EOF Then
            ListView2.ListItems(i).SubItems(1) = miRSAux!campo1
            ListView2.ListItems(i).SubItems(3) = miRSAux!codext
            If miRSAux!codext = ListView2.ListItems(i).SubItems(2) Then ListView2.ListItems(i).Checked = True
        Else
            'HA sido borrado
            ListView2.ListItems(i).Checked = True
            ListView2.ListItems(i).SubItems(1) = " ***** NUEVO ***** "
            ListView2.ListItems(i).SubItems(3) = "N"
        End If
        miRSAux.Close
    Next i

    Conn.Execute "DELETE FROM timagenhco"

    Command2.Visible = False
    Me.cmdRestaurar.Visible = True
    Exit Sub
ERecu:
    MuestraError Err.Number, "", cad
    Set miRSAux = Nothing
    Set Car = Nothing
End Sub

Private Sub Form_Load()
    Limpiar Me
    Label3.Caption = ""
    Label5.Caption = ""
    If Opcion = 0 Then
        ElFrame False
    Else
        cmdRestaurar.Visible = False
        Me.Command2.Visible = True
        Set ListView2.SmallIcons = Admin.ImageList2
        MostrarArchivos
    End If
    Me.FrameRecupera.Visible = Opcion = 1
    
End Sub



Private Sub IniciarProceso()

    'preaparamos  Todo
    If Not ComprobarCarpetaBackup(True) Then Exit Sub
    
    
    'Comprobamos que no existe la BD
    If Not ComprobarBD_DelBackUp Then Exit Sub
    
    
    'Actualizmos la BD con la carpetas almacen
    ACtualizarCarpetaAlmacen
    
    'Todo OK---> cerramos y punto en boca
    ElFrame True
        
    
    
End Sub

Private Function ComprobarBD_DelBackUp() As Boolean
Dim Existe As Boolean

    
    ComprobarBD_DelBackUp = False
    On Error Resume Next
    Set miRSAux = New ADODB.Recordset
    
    
    
    'Comprobamos si existe la BD
    Existe = False
    miRSAux.Open "Select * from backAridoc.almacen", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Err.Number <> 0 Then
        'NO EXISTE LA BD
        Err.Clear
    Else
        miRSAux.Close
        Existe = True
    End If
    Set miRSAux = Nothing
    
    If Existe Then
        Conn.Execute "DROP DATABASE backAridoc"
        If Err.Number <> 0 Then
            MsgBox "Imposible eliminar BD temporal para la recuperacion de archivos", vbExclamation
            Exit Function
        End If
    End If
    
    'Creamos la BD
    Conn.Execute "CREATE DATABASE backAridoc"
    If Err.Number <> 0 Then
        MsgBox "Imposible Crear BD temporal para la recuperacion de archivos", vbExclamation
        'INTENTAMOS ELIMINAR BD
        Conn.Execute "DROP DATABASE backAridoc"
        Err.Clear
        Exit Function
    End If
    
    'Cerramos conexion
    Conn.Close
    espera 1
    If Not AbrirConexion(False) Then
        MsgBox "Error abriendo conexion BAKCUP", vbExclamation
        If Not AbrirConexion(True) Then
            MsgBox "ERROR GRAVE. Abrirconexion BD", vbCritical
            End
        End If
    End If
    
    Label3.Caption = "Restaurar BD"
    Label3.Refresh
    
    'OK. Hemos llegado aqui. Siguiente paso....  recuperar el backUP
    If Not ProcesarFicherobackUP Then Exit Function
    
    
    'Eliminamos tmp integra
     cad = "DELETE from tmpintegra"
    Conn.Execute cad
    
    
    ComprobarBD_DelBackUp = True
End Function



Private Function ProcesarFicherobackUP() As Boolean
Dim NF As Integer
Dim cad As String
Dim Fin As Boolean
Dim Seguir As Boolean
Dim SQL As String
Dim T1 As Single

    On Error GoTo EProcesarFicherobackUP
    ProcesarFicherobackUP = False
    cad = Text1.Text
    NF = FreeFile
    Open cad For Input As #NF
    
    'Buscamos la cadena de inicio del bakcup
    '-- Host: localhost    Database: aridoc
    Seguir = False
    Fin = False
    While Not Fin
        Line Input #NF, cad
        If InStr(cad, "Database: aridoc") > 0 Then
            'OK ESTA CORRECTO
            Fin = True
            Seguir = True
        Else
            If EOF(NF) Then Fin = True
        End If
    Wend
    
    If Not Seguir Then
        MsgBox "Error en el archivo de BACKUP.   'Database aridoc' no encontrado", vbExclamation
        Exit Function
    End If
    
    
    
    'OK. AHora iremos realizando los inserts correspondientes
    T1 = Timer
    Fin = False
    SQL = ""
    While Not Fin
    
        Line Input #NF, cad
        Debug.Print cad
        cad = Trim(cad)
        If cad <> "" Then
            If Mid(cad, 1, 2) = "--" Then
                'NO HACEMOS NADA
                SQL = ""
            Else
                If Mid(cad, 1, 1) = "#" Then
                    SQL = ""
                Else
                    SQL = SQL & cad
                    If Right(cad, 1) = ";" Then
                        Conn.Execute SQL
                        SQL = ""
                    End If
                End If
            End If
        End If
        If Timer - T1 > 1 Then
             
            If Len(cad) Then cad = Mid(cad, 1, 20)
            Label3.Caption = Format(Now, "hh:mm:ss") & " - " & cad & "  ........."
            Label3.Refresh
            T1 = Timer
        End If
        Fin = EOF(NF)
    Wend
    Close #NF
    
    ProcesarFicherobackUP = True
    Exit Function
EProcesarFicherobackUP:
    MuestraError Err.Number, Err.Description
    i = NF
    IntentaCerrar NF
End Function

Private Function IntentaCerrar()
    On Error Resume Next
    Close #i
    Err.Clear
End Function


Private Sub ElFrame(Habilitar As Boolean)

    Frame1.Visible = Habilitar
    If Habilitar Then
        i = 0
    Else
        i = 1
    End If
    Command1(i).Cancel = True
    Command1(i).Default = True
    
End Sub

Private Function ACtualizarCarpetaAlmacen() As Boolean
Dim J As Integer
Dim It As ListItem

    On Error GoTo EACtualizarCarpetaAlmacen
    ACtualizarCarpetaAlmacen = False
    Set miRSAux = New ADODB.Recordset
    cad = "Select * from almacen where codalma >=3"
    miRSAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRSAux.EOF Then
        miRSAux.Close
        MsgBox "Ningun codigo almacen correcto(>=3)", vbExclamation
        Exit Function
    End If
    If Right(Text2.Text, 1) <> "\" Then Text2.Text = Text2.Text & "\"
    While Not miRSAux.EOF
        If miRSAux!version = 0 Then
            J = InStrRev(miRSAux!pathreal, "/")
        Else
            J = InStrRev(miRSAux!pathreal, "\")
        End If
        If J > 0 Then
            cad = Mid(miRSAux!pathreal, J + 1)
            cad = ", pathreal = """ & DevNombreSql(Text2.Text & cad & """")
            cad = "UPDATE almacen set version =1 " & cad & " WHERE codalma = " & miRSAux!codalma
        Else
            cad = ""
        End If
        miRSAux.MoveNext
        If cad <> "" Then Conn.Execute cad
        
    Wend
    miRSAux.Close

    
    'Y ya de paso, las cargamos
    cad = "Select * from almacen where codalma >=3"
    miRSAux.Open cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not miRSAux.EOF
        Set It = ListView1.ListItems.Add()
        It.Text = miRSAux!codalma
        It.SubItems(1) = miRSAux!pathreal

        miRSAux.MoveNext
    Wend
    miRSAux.Close
    
    
    'Borramos las tablas auxiliares
    Label3.Caption = "Eliminando temporales"
    Label3.Refresh
    cad = "Delete from carpetashco"
    Conn.Execute cad
    cad = "Delete from timagenhco"
    Conn.Execute cad
    cad = "Delete from almacenhco"
    Conn.Execute cad
    
    ACtualizarCarpetaAlmacen = True
    Exit Function
EACtualizarCarpetaAlmacen:
    MuestraError Err.Number, "ACtualizar Carpeta Almacen BK"
    Set miRSAux = Nothing
End Function

Private Sub Image1_Click(Index As Integer)
    If Index < 2 Then
        If Index = 0 Then
            If ListView1.SelectedItem Is Nothing Then Exit Sub
        End If
        cad = GetFolder("Carpeta archivos BACKUP")
        If cad <> "" Then
            If Right(cad, 1) <> "\" Then cad = cad & "\"
            If Index = 1 Then
                Text2.Text = cad
            Else
                ListView1.SelectedItem.SubItems(1) = cad
            End If
        End If
    Else
        cd1.DialogTitle = "Archivo DUMP"
        cd1.Filter = "Archivos sql | *.sql"
        cd1.ShowOpen
        If cd1.FileName <> "" Then Text1.Text = cd1.FileName
    End If
End Sub

Private Sub ListView1_DblClick()
    Image1_Click 0
End Sub



'Mostramos los archivos
Private Sub MostrarArchivos()
Dim ItmX As ListItem
Dim cad As String



    Screen.MousePointer = vbHourglass
    
    Text3.Text = RecuperaValor(DatosCopiados, 1)
    Text3.Tag = RecuperaValor(DatosCopiados, 2)
    DatosCopiados = RecuperaValor(DatosCopiados, 3) 'EL PADRE
    DatosCopiados = Mid(DatosCopiados, 2)
    ListView2.ColumnHeaders(1).Text = vConfig.C1
    ListView2.ColumnHeaders(2).Text = vConfig.C1 & " en ARIDOC"
    ListView2.ListItems.Clear
    ListView2.View = lvwReport
    cad = "Select  * from timagenhco"
    cad = cad & " ORDER BY campo1"

    
    
    
    Set miRSAux = New ADODB.Recordset
    miRSAux.Open cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    
    While Not miRSAux.EOF
        Set ItmX = ListView2.ListItems.Add(, "C" & miRSAux!codigo)
        ItmX.Text = miRSAux!campo1
        ItmX.SubItems(2) = miRSAux!codext
        ItmX.SmallIcon = miRSAux!codext + 1
        miRSAux.MoveNext
    Wend
        
    
    miRSAux.Close
    Set miRSAux = Nothing
    


Screen.MousePointer = vbDefault
End Sub

