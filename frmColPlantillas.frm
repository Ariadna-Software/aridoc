VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmColPlantillas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Administrador plantillas"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6435
   Icon            =   "frmColPlantillas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameSinTabla2 
      Height          =   6495
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   6375
      Begin VB.CommandButton Command1 
         Caption         =   "CREAR"
         Height          =   375
         Left            =   3720
         TabIndex        =   9
         Top             =   4320
         Width           =   1815
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Crear estructura carpetas para las plantillas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   5775
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   6060
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   6000
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmColPlantillas.frx":030A
      Height          =   5325
      Left            =   60
      TabIndex        =   5
      Top             =   540
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   9393
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   4980
      TabIndex        =   4
      Top             =   6060
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   120
      TabIndex        =   2
      Top             =   5895
      Width           =   2865
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   2550
      End
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   375
      Left            =   6030
      Top             =   30
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Modificar Lineas"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
         Visible         =   0   'False
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver todos"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmColPlantillas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'grupos (codgrupo, nomgrupo, nive

'Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
'Public Event DatoSeleccionado(CadenaSeleccion As String)
Private WithEvents frmP As frmPregunta
Attribute frmP.VB_VarHelpID = -1

Private CadenaConsulta As String
Dim CadAncho As Boolean  'Para cuando llamemos al al form de lineas
Dim Modo As Byte

'----------------------------------------------
'----------------------------------------------
'   Deshabilitamos todos los botones menos
'   el de salir
'   Ademas mostramos aceptar y cancelar
'   Modo 0->  Normal
'   Modo 1 -> Lineas  INSERTAR
'   Modo 2 -> Lineas MODIFICAR
'   Modo 3 -> Lineas BUSCAR
'----------------------------------------------
'----------------------------------------------

Private Sub PonerModo(vModo)
Dim B As Boolean
Modo = vModo

B = (Modo = 0)

'txtAux(0).Visible = Not B
'txtAux(1).Visible = Not B
'txtAux(2).Visible = Not B
mnOpciones.Enabled = B
Toolbar1.Buttons(1).Enabled = B
Toolbar1.Buttons(2).Enabled = B
Toolbar1.Buttons(8).Enabled = B
Toolbar1.Buttons(7).Enabled = B
Toolbar1.Buttons(6).Enabled = B

'Prueba
cmdAceptar.Visible = Not B
cmdCancelar.Visible = Not B
DataGrid1.Enabled = B

'Si es regresar
'If DatosADevolverBusqueda <> "" Then
'    cmdRegresar.Visible = B
'End If
'Si estamo mod or insert
'If Modo = 2 Then
'   txtAux(0).BackColor = &H80000018
'   Else
'    txtAux(0).BackColor = &H80000005
'End If
'txtAux(0).Enabled = (Modo <> 2)
End Sub

Private Sub BotonAnyadir()

    AccionesCarpetas 0

End Sub



Private Sub BotonVerTodos()
    CargaGrid ""
End Sub

Private Sub BotonBuscar()
'    CargaGrid "carpeta = -1"
'    'Buscar
'    txtAux(0).Text = ""
'    txtAux(1).Text = ""
'    txtAux(2).Text = ""
'    LLamaLineas DataGrid1.Top + 206, 2
'    txtAux(0).SetFocus
End Sub

Private Sub BotonModificar()
'    '---------
'    'MODIFICAR
'    '----------
AccionesCarpetas 1
End Sub

'Private Sub LLamaLineas(alto As Single, xModo As Byte)
'PonerModo xModo + 1
''Fijamos el ancho
'txtAux(0).Top = alto
'txtAux(1).Top = alto
'txtAux(2).Top = alto
'
'End Sub




Private Sub BotonEliminar()

    AccionesCarpetas 2

'Dim SQL As String
'    On Error GoTo Error2
'    'Ciertas comprobaciones
'    If adodc1.Recordset.EOF Then Exit Sub
'    If adodc1.Recordset!codgrupo < 1 Then
'        MsgBox "No puede elimiminar el grupo", vbExclamation
'        Exit Sub
'    End If
'
'    If adodc1.Recordset!Nivel < vUsu.Nivel Then
'        MsgBox "No tiene permiso para eliminar el grupo", vbExclamation
'        Exit Sub
'    End If
'
'    If Not SepuedeBorrar Then Exit Sub
'    '### a mano
'    SQL = "Seguro que desea eliminar el grupo:"
'    SQL = SQL & vbCrLf & "Código: " & adodc1.Recordset.Fields(0)
'    SQL = SQL & vbCrLf & "Denominación: " & adodc1.Recordset.Fields(1)
'    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
'        'Hay que eliminar
'        'SQL = "Delete from gruo where codconce=" & adodc1.Recordset!codconce
'        'Conn.Execute SQL
'
'
'        CargaGrid ""
'        adodc1.Recordset.Cancel
'    End If
'    Exit Sub
'Error2:
'        Screen.MousePointer = vbDefault
'        MuestraError Err.Number, "Eliminando registro", Err.Description
End Sub





'Private Sub cmdAceptar_Click()
'Dim i As Integer
'Dim CadB As String
'Select Case Modo
'    Case 1
'    If DatosOk Then
'            '-----------------------------------------
'            'Hacemos insertar
'            If InsertarDesdeForm(Me) Then
'                'MsgBox "Registro insertado.", vbInformation
'                CargaGrid
'                BotonAnyadir
'            End If
'        End If
'    Case 2
'            'Modificar
'            If DatosOk Then
'                '-----------------------------------------
'                'Hacemos insertar
'                If ModificaDesdeFormulario(Me) Then
'                    i = adodc1.Recordset.Fields(0)
'                    PonerModo 0
'                    CargaGrid
'                    adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " =" & i)
'                End If
'            End If
'    Case 3
'        'HacerBusqueda
'        CadB = ObtenerBusqueda(Me)
'        If CadB <> "" Then
'            PonerModo 0
'            CargaGrid CadB
'        End If
'    End Select
'
'
'End Sub

Private Sub cmdCancelar_Click()
Select Case Modo
Case 1
    DataGrid1.AllowAddNew = False
    'CargaGrid
    If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
    
Case 3
    CargaGrid
End Select
PonerModo 0
lblIndicador.Caption = ""
DataGrid1.SetFocus
End Sub

'Private Sub cmdRegresar_Click()
'Dim cad As String
'
'If adodc1.Recordset.EOF Then
'    MsgBox "Ningún registro a devolver.", vbExclamation
'    Exit Sub
'End If
'
'If adodc1.Recordset.Fields(0) >= 900 Then
'    MsgBox "Los conceptos superiores a 900 se los reserva la aplicación.", vbExclamation
'    Exit Sub
'End If
'cad = adodc1.Recordset.Fields(0) & "|"
'cad = cad & adodc1.Recordset.Fields(1) & "|"
'cad = cad & adodc1.Recordset.Fields(2) & "|"
'RaiseEvent DatoSeleccionado(cad)
'Unload Me
'End Sub

Private Sub cmdRegresar_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Command1_Click()

    'Msgbox varios
    CadenaConsulta = "Va a generar la estructura necesaria para la creación de carpetas en las plantillas."
    CadenaConsulta = CadenaConsulta & vbCrLf & vbCrLf & vbCrLf & "¿Desea continuar?"
    If MsgBox(CadenaConsulta, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    'Si tiene carpetas



    'Crear estrucutra
    CadenaConsulta = "CREATE TABLE plantillacarpetas ("
    CadenaConsulta = CadenaConsulta & " carpeta smallint  NOT NULL default '0',"
    CadenaConsulta = CadenaConsulta & " texto varchar(255) NOT NULL default '0',"
    CadenaConsulta = CadenaConsulta & " groupprop int(10) unsigned NOT NULL default '0',"
    CadenaConsulta = CadenaConsulta & " lecturag int(10) unsigned NOT NULL default '0',"
    CadenaConsulta = CadenaConsulta & " Primary Key(Carpeta)"
    CadenaConsulta = CadenaConsulta & " ) TYPE=MyISAM COMMENT='Carpetas plantillas';"
    
    Conn.Execute CadenaConsulta
    
    'CREAMOS LA PRIMERA VACIA
    CadenaConsulta = "INSERT INTO plantillacarpetas (carpeta, texto, groupprop, lecturag) VALUES (1,'PLANTILLAS',1," & vbPermisoTotal & ")"
    Conn.Execute CadenaConsulta
    
    'PASAMOS TODAS LAS PLANTILLAS A LA ESTRCUTURA NUEVA
    CadenaConsulta = "UPDATE plantilla SET Carpeta= 1"
    Conn.Execute CadenaConsulta
    
    
    
    'AHORA. Finalmente ya tenemos la estructura
    CadenaConsulta = "Select carpeta,texto,nomgrupo from plantillacarpetas,grupos where groupprop = codgrupo "
    CargaGrid
    Me.FrameSinTabla2.Visible = False
End Sub

Private Sub DataGrid1_DblClick()
'If cmdRegresar.Visible Then cmdRegresar_Click
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()

          ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = Admin.imgListComun
        .Buttons(1).Image = 1
        .Buttons(2).Image = 2
        .Buttons(6).Image = 3
        .Buttons(7).Image = 4
        .Buttons(8).Image = 5
        '.Buttons(10).Image = 10
        .Buttons(11).Image = 16
        .Buttons(12).Image = 15
        .Buttons(14).Image = 6
        .Buttons(15).Image = 7
        .Buttons(16).Image = 8
        .Buttons(17).Image = 9
    End With

    
    'Bloqueo de tabla, cursor type
'    Adodc1.UserName = vUsu.Login
'    Adodc1.password = vUsu.Passwd
    
    
   ' cmdRegresar.Visible = (DatosADevolverBusqueda <> "")
    
    DespalzamientoVisible False
    PonerModo 0
    CadAncho = TienePlantillasEnCarpeta2
    If CadAncho Then
        'TIENE CARPETAS
        Me.FrameSinTabla2.Visible = False
        
        PonerOpcionesMenu  'En funcion del usuario
    
        'Cadena consulta
        CadenaConsulta = "Select carpeta,texto,nomgrupo from plantillacarpetas,grupos where groupprop = codgrupo "
        CadAncho = False
        CargaGrid
        
    Else
    
        Me.FrameSinTabla2.Visible = True
        mnOpciones.Enabled = False
        
    End If
    lblIndicador.Caption = ""
End Sub




Private Sub frmP_DatoSeleccionado(OpcionSeleccionada As Byte)
    CadenaConsulta = OpcionSeleccionada
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnModificar_Click()
    BotonModificar
End Sub

Private Sub mnNuevo_Click()
BotonAnyadir
End Sub

Private Sub mnSalir_Click()
Screen.MousePointer = vbHourglass
Unload Me
End Sub

Private Sub mnVerTodos_Click()
BotonVerTodos
End Sub



'----------------------------------------------------------------
'----------------------------------------------------------------
'   En PONERMODO se habilitan, o no, los diverso campos del
'   formulario en funcion del modo en k vayamos a trabajar
'
'### A mano
'Esto es para que cuando pincha en siguiente le sugerimos
'Se puede comentar todo y asi no hace nada ni da error
'El SQL es propio de cada tabla
Private Function SugerirCodigoSiguiente() As String
'    Dim SQL As String
'    Dim RS As ADODB.Recordset
    
'    SQL = "Select Max(codconce) from conceptos where codconce<900"
'
'    Set RS = New ADODB.Recordset
'    RS.Open SQL, Conn, , , adCmdText
'    SQL = "1"
'    If Not RS.EOF Then
'        If Not IsNull(RS.Fields(0)) Then
'            SQL = CStr(RS.Fields(0) + 1)
'        End If
'    End If
'    RS.Close
'    SugerirCodigoSiguiente = SQL
End Function

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
        BotonBuscar
Case 2
        BotonVerTodos
Case 6
        BotonAnyadir
Case 7
        BotonModificar
Case 8
        BotonEliminar
Case 11

Case 12
        Unload Me
Case Else

End Select
End Sub


Private Sub DespalzamientoVisible(bol As Boolean)
    Dim I
    For I = 14 To 17
        Toolbar1.Buttons(I).Visible = bol
    Next I
End Sub

Private Sub CargaGrid(Optional SQL As String)
    Dim J As Integer
    Dim TotalAncho As Integer
    Dim I As Integer
    
    adodc1.ConnectionString = Conn
    If SQL <> "" Then
        SQL = CadenaConsulta & " WHERE " & SQL
        Else
        SQL = CadenaConsulta
    End If
    SQL = SQL & " ORDER BY carpeta"
    adodc1.RecordSource = SQL
    adodc1.CursorType = adOpenDynamic
    adodc1.LockType = adLockOptimistic
    adodc1.Refresh
    
    DataGrid1.AllowRowSizing = False
    DataGrid1.RowHeight = 290
    
    
    'Nombre producto
    I = 0
        DataGrid1.Columns(I).Caption = "Codigo"
        DataGrid1.Columns(I).Width = 900
        DataGrid1.Columns(I).NumberFormat = "000"
        
    
    'Leemos del vector en 2
    I = 1
        DataGrid1.Columns(I).Caption = "Denominación"
        DataGrid1.Columns(I).Width = 2500
        TotalAncho = TotalAncho + DataGrid1.Columns(I).Width
    
    I = 2
        DataGrid1.Columns(I).Caption = "Grupo propietario"
        DataGrid1.Columns(I).Width = 2300
        TotalAncho = TotalAncho + DataGrid1.Columns(I).Width

            
   ' DataGrid1.Columns(3).Visible = False
        'Fiajamos el cadancho
    If Not CadAncho Then
        'La primera vez fijamos el ancho y alto de  los txtaux
'        txtAux(0).Width = DataGrid1.Columns(0).Width - 60
'        txtAux(1).Width = DataGrid1.Columns(1).Width - 60
'        txtAux(2).Width = DataGrid1.Columns(2).Width - 60
'        txtAux(0).Left = DataGrid1.Left + 340
'        txtAux(1).Left = txtAux(0).Left + txtAux(0).Width + 45
'        txtAux(2).Left = txtAux(1).Left + txtAux(1).Width + 45
        CadAncho = True
    End If
    'Habilitamos modificar y eliminar
    If vUsu.Nivel < 2 Then
        Toolbar1.Buttons(7).Enabled = Not adodc1.Recordset.EOF
        Toolbar1.Buttons(8).Enabled = Not adodc1.Recordset.EOF
    End If
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
'With txtAux(Index)
'    .SelStart = 0
'    .SelLength = Len(.Text)
'End With
End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtAux_LostFocus(Index As Integer)

'txtAux(Index).Text = Trim(txtAux(Index).Text)
'If txtAux(Index).Text = "" Then Exit Sub
'If Modo = 3 Then Exit Sub 'Busquedas
'If Index = 0 Then
'    If Not IsNumeric(txtAux(0).Text) Then
'        MsgBox "Código concepto tiene que ser numérico", vbExclamation
'        Exit Sub
'    End If
'    txtAux(0).Text = Format(txtAux(0).Text, "000")
'End If
End Sub


'Private Function DatosOk() As Boolean
'Dim Datos As String
'Dim B As Boolean
'DatosOk = False
'B = CompForm(Me)
'If Not B Then Exit Function
'
'
'
'
'
'If Modo = 1 Then
'    'Comprobamos k el numero NO se > 32
'    If Val(txtAux(0).Text) > vbMaxGrupos Then
'        MsgBox "Numero de grupo mayor del permitido: " & vbMaxGrupos, vbExclamation
'        B = False
'
'    Else
'        'Estamos insertando
'         Datos = DevuelveDesdeBD("codgrupo", "grupos", "codgrupo", txtAux(0).Text, "N")
'         If Datos <> "" Then
'            MsgBox "Ya existe el grupo : " & txtAux(0).Text, vbExclamation
'            B = False
'        End If
'    End If
'End If
'DatosOk = B
'End Function



Private Sub PonerOpcionesMenu()
PonerOpcionesMenuGeneral Me
End Sub



Private Function SepuedeBorrar() As Boolean
Dim SQL As String
'    SepuedeBorrar = False
'    SQL = DevuelveDesdeBD("tipoamor", "samort", "condebes", adodc1.Recordset!codconce, "N")
'    If SQL <> "" Then
'        MsgBox "Esta vinculada con parametros de amortizacion", vbExclamation
'        Exit Function
'    End If
'    SQL = DevuelveDesdeBD("tipoamor", "samort", "conhaber", adodc1.Recordset!codconce, "N")
'    If SQL <> "" Then
'        MsgBox "Esta vinculada con parametros de amortizacion", vbExclamation
'        Exit Function
'    End If
    
    SepuedeBorrar = True
End Function


Private Sub DeseleccionaGrid()
    On Error GoTo EDeseleccionaGrid
        
    While DataGrid1.SelBookmarks.Count > 0
        DataGrid1.SelBookmarks.Remove 0
    Wend
    Exit Sub
EDeseleccionaGrid:
        Err.Clear
End Sub


Private Sub KEYpress(KeyAscii As Integer)
    'Caption = KeyAscii
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    Else
        If KeyAscii = 27 Then
            If Modo = 0 Then Unload Me
        End If
    End If
End Sub



'Private Sub AccionesUsuario(Index As Integer)
'Dim vUs As Cusuarios
'Dim Valor As Integer
'Dim SQL As String
'
'    If Index = 0 Then
'        'Nuevo
'        DatosMOdificados = False
'        Set frmUsuario.vU = Nothing
'        frmUsuario.Show vbModal
'        If DatosMOdificados Then BotonVerTodos
'
'    Else
'        If adodc1.Recordset.EOF Then Exit Sub
'
'
'        Valor = adodc1.Recordset!Carpeta
'
'            If Index = 2 Then
'                    'ELIMINAR
'                SQL = "Desea elimniar el usuario: " & vUs.codusu & " - " & vUs.Nombre & "?"
'                If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
'                    If vUs.Eliminar = 0 Then
'                        If vUsu.codusu = vUs.codusu Then
'                            MsgBox "El programa  finalizará", vbCritical
'                            End
'                            Exit Sub
'                        Else
'                            BotonVerTodos
'                        End If
'                    End If
'                End If
'            Else
'
'                DatosMOdificados = False
'                Set frmUsuario.vU = vUs
'                frmUsuario.Show vbModal
'                If DatosMOdificados Then
'                    If vUsu.codusu = vUs.codusu Then
'                        MsgBox "El programa  finalizará", vbCritical
'                        End
'                        Exit Sub
'                    Else
'                        BotonVerTodos
'                    End If
'                End If 'De datos modificados
'            End If 'index=2
'        End If
'    End If
'End Sub
'Private Function SoloUnaCarpeta() As Boolean
'        Set miRSAux = New ADODB.Recordset
'        miRSAux.Open "Select count(*) from plantillacarpetas", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'        SoloUnaCarpeta = False
'        If DBLet(miRSAux.Fields(0), "N") = 1 Then SoloUnaCarpeta = True
'        miRSAux.Close
'        Set miRSAux = Nothing
'End Function

Private Sub AccionesCarpetas(Opcion As Byte)
Dim cad As String
    If Opcion > 0 Then
        If adodc1.Recordset.EOF Then Exit Sub
    End If
    
    
    If Opcion = 2 Then
        cad = DevuelveDesdeBD("codigo", "plantilla", "carpeta", adodc1.Recordset!Carpeta, "N")
        If cad <> "" Then
            MsgBox "No se puede eliminar. Hay plantillas  dentro de la carpeta", vbExclamation
            Exit Sub
            
        Else
            'SoloUnaCarpeta
            If adodc1.Recordset.RecordCount = 1 Then
                MsgBox "Solo queda un carpeta. Debe eliminar la estrucutra", vbExclamation
                Exit Sub
            End If
        
        
            'Vamos a eliminar
            cad = "Desea eliminar la carpeta : " & adodc1.Recordset!Texto & " ?"
            If MsgBox(cad, vbQuestion + vbYesNo) = vbYes Then
                cad = "DELETE FROM plantillacarpetas where carpeta = " & adodc1.Recordset!Carpeta
                Conn.Execute cad
                CargaGrid
            End If
        End If
        
    Else
    
        'NUEVO O MODIFICAR
        Me.Tag = CadenaConsulta
        CadenaConsulta = ""
        Set frmP = New frmPregunta
        If Opcion = 0 Then
            frmP.origenDestino = ""
        Else
            frmP.origenDestino = adodc1.Recordset!Carpeta & "|" & adodc1.Recordset!Texto & "|"
        End If
        frmP.Opcion = 23
        frmP.Show vbModal
        If CadenaConsulta <> "" Then
            cad = CadenaConsulta
            CadenaConsulta = Me.Tag
            Me.Tag = ""
            CargaGrid
            adodc1.Recordset.Find "carpeta = " & cad
            If adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
        End If
        Set frmP = Nothing
    End If
    
    
    
    
    
    
End Sub
