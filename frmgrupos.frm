VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form frmGrupos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GRUPOS"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   Icon            =   "frmgrupos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2340
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Tag             =   "Nivel|N|N|||grupos|nivel|||"
      Top             =   5640
      Width           =   615
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   6060
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4980
      TabIndex        =   4
      Top             =   6060
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   1
      Left            =   900
      MaxLength       =   30
      TabIndex        =   1
      Tag             =   "Denominación|T|N|||grupos|nomgrupo|||"
      Text            =   "Dato2"
      Top             =   5640
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   0
      Left            =   60
      MaxLength       =   3
      TabIndex        =   0
      Tag             =   "Código grupo|N|N|0||grupos|codgrupo|000|S|"
      Text            =   "Dat"
      Top             =   5640
      Width           =   800
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmgrupos.frx":030A
      Height          =   5325
      Left            =   60
      TabIndex        =   8
      Top             =   540
      Width           =   5970
      _ExtentX        =   10530
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
      TabIndex        =   7
      Top             =   6060
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   120
      TabIndex        =   5
      Top             =   5895
      Width           =   2865
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   6
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
      TabIndex        =   9
      Top             =   0
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   1
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
Attribute VB_Name = "frmGrupos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'grupos (codgrupo, nomgrupo, nive

Public DatosADevolverBusqueda As String    'Tendra el nº de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)


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

txtAux(0).Visible = Not B
txtAux(1).Visible = Not B
Combo1.Visible = Not B
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
If DatosADevolverBusqueda <> "" Then
    cmdRegresar.Visible = B
End If
'Si estamo mod or insert
If Modo = 2 Then
   txtAux(0).BackColor = &H80000018
   Else
    txtAux(0).BackColor = &H80000005
End If
txtAux(0).Enabled = (Modo <> 2)
End Sub

Private Sub BotonAnyadir()
    Dim NumF As String
    Dim anc As Single
    
    'Obtenemos la siguiente numero de factura
    NumF = SugerirCodigoSiguiente
    lblIndicador.Caption = "INSERTANDO"
    'Situamos el grid al final
    DataGrid1.AllowAddNew = True
    If adodc1.Recordset.RecordCount > 0 Then
        DataGrid1.HoldFields
        adodc1.Recordset.MoveLast
        DataGrid1.Row = DataGrid1.Row + 1
    End If
    
    
   
    If DataGrid1.Row < 0 Then
        anc = 770
        Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + 545
    End If
    txtAux(0).Text = NumF
    txtAux(1).Text = ""
    Combo1.ListIndex = -1
    LLamaLineas anc, 0
    
    
    'Ponemos el foco
    txtAux(0).SetFocus
    
'    If FormularioHijoModificado Then
'        CargaGrid
'        BotonAnyadir
'        Else
'            'cmdCancelar.SetFocus
'            If Not Adodc1.Recordset.EOF Then _
'                Adodc1.Recordset.MoveFirst
'    End If
End Sub



Private Sub BotonVerTodos()
    CargaGrid ""
End Sub

Private Sub BotonBuscar()
    CargaGrid "codgrupo = -1"
    'Buscar
    txtAux(0).Text = ""
    txtAux(1).Text = ""
    LLamaLineas DataGrid1.Top + 206, 2
    txtAux(0).SetFocus
End Sub

Private Sub BotonModificar()
    '---------
    'MODIFICAR
    '----------
    Dim cad As String
    Dim anc As Single
    Dim i As Integer
    If adodc1.Recordset.EOF Then Exit Sub
    If adodc1.Recordset.RecordCount < 1 Then Exit Sub


    If adodc1.Recordset!codgrupo < 2 Then
        MsgBox "Este grupo no puede camiarse de nivel", vbExclamation
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Me.lblIndicador.Caption = "MODIFICAR"
    DeseleccionaGrid
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        i = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, i
        DataGrid1.Refresh
    End If
    
    If DataGrid1.Row < 0 Then
        anc = 320
        Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + 545
    End If

    'Llamamos al form
    txtAux(0).Text = DataGrid1.Columns(0).Text
    txtAux(1).Text = DataGrid1.Columns(1).Text
    For i = 0 To Combo1.ListCount - 1
        If Combo1.ItemData(i) = adodc1.Recordset!Nivel Then
            Combo1.ListIndex = i
            Exit For
        End If
    Next i
    
    LLamaLineas anc, 1
   
   'Como es modificar
   txtAux(1).SetFocus
   
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
PonerModo xModo + 1
'Fijamos el ancho
txtAux(0).Top = alto
txtAux(1).Top = alto
Combo1.Top = alto - 15
End Sub




Private Sub BotonEliminar()
Dim SQL As String
    On Error GoTo Error2
    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
    If adodc1.Recordset!codgrupo < 2 Then
        MsgBox "No puede elimiminar el grupo", vbExclamation
        Exit Sub
    End If
    
    If adodc1.Recordset!Nivel < vUsu.Nivel Then
        MsgBox "No tiene permiso para eliminar el grupo", vbExclamation
        Exit Sub
    End If
    
    If Not SepuedeBorrar Then Exit Sub
    '### a mano
    SQL = "Seguro que desea eliminar el grupo:"
    SQL = SQL & vbCrLf & "Código: " & adodc1.Recordset.Fields(0)
    SQL = SQL & vbCrLf & "Denominación: " & adodc1.Recordset.Fields(1)
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        'Hay que eliminar
        SQL = "Delete from grupos where codgrupo =" & adodc1.Recordset!codgrupo
        Conn.Execute SQL
        
        
        CargaGrid ""
        adodc1.Recordset.Cancel
    End If
    Exit Sub
Error2:
        Screen.MousePointer = vbDefault
        MuestraError Err.Number, "Eliminando registro", Err.Description
End Sub





Private Sub cmdAceptar_Click()
Dim i As Integer
Dim CadB As String
Select Case Modo
    Case 1
    If DatosOk Then
            '-----------------------------------------
            'Hacemos insertar
            If InsertarDesdeForm(Me) Then
                'MsgBox "Registro insertado.", vbInformation
                CargaGrid
                BotonAnyadir
            End If
        End If
    Case 2
            'Modificar
            If DatosOk Then
                '-----------------------------------------
                'Hacemos insertar
                If ModificaDesdeFormulario(Me) Then
                    i = adodc1.Recordset.Fields(0)
                    PonerModo 0
                    CargaGrid
                    adodc1.Recordset.Find (adodc1.Recordset.Fields(0).Name & " =" & i)
                End If
            End If
    Case 3
        'HacerBusqueda
        CadB = ObtenerBusqueda(Me)
        If CadB <> "" Then
            PonerModo 0
            CargaGrid CadB
        End If
    End Select


End Sub

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

Private Sub cmdRegresar_Click()
Dim cad As String

If adodc1.Recordset.EOF Then
    MsgBox "Ningún registro a devolver.", vbExclamation
    Exit Sub
End If

If adodc1.Recordset.Fields(0) >= 900 Then
    MsgBox "Los conceptos superiores a 900 se los reserva la aplicación.", vbExclamation
    Exit Sub
End If
cad = adodc1.Recordset.Fields(0) & "|"
cad = cad & adodc1.Recordset.Fields(1) & "|"
cad = cad & adodc1.Recordset.Fields(2) & "|"
RaiseEvent DatoSeleccionado(cad)
Unload Me
End Sub

Private Sub cmdRegresar_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub DataGrid1_DblClick()
If cmdRegresar.Visible Then cmdRegresar_Click
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
    CargaCombo
    
    cmdRegresar.Visible = (DatosADevolverBusqueda <> "")
    
    DespalzamientoVisible False
    PonerModo 0
    CadAncho = False
    PonerOpcionesMenu  'En funcion del usuario
    'Cadena consulta
    CadenaConsulta = "Select grupos.*,descripcion from grupos,descniveles"
    CadenaConsulta = CadenaConsulta & " WHERE grupos.nivel = descniveles.codnivel"
    CargaGrid
    lblIndicador.Caption = ""
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
    Dim i
    For i = 14 To 17
        Toolbar1.Buttons(i).Visible = bol
    Next i
End Sub

Private Sub CargaGrid(Optional SQL As String)
    Dim J As Integer
    Dim TotalAncho As Integer
    Dim i As Integer
    
    adodc1.ConnectionString = Conn
    If SQL <> "" Then
        SQL = CadenaConsulta & " AND " & SQL
        Else
        SQL = CadenaConsulta
    End If
    SQL = SQL & " ORDER BY codgrupo"
    adodc1.RecordSource = SQL
    adodc1.CursorType = adOpenDynamic
    adodc1.LockType = adLockOptimistic
    adodc1.Refresh
    
    DataGrid1.AllowRowSizing = False
    DataGrid1.RowHeight = 290
    
    
    'Nombre producto
    i = 0
        DataGrid1.Columns(i).Caption = "Cod."
        DataGrid1.Columns(i).Width = 700
        DataGrid1.Columns(i).NumberFormat = "000"
        
    
    'Leemos del vector en 2
    i = 1
        DataGrid1.Columns(i).Caption = "Denominación"
        DataGrid1.Columns(i).Width = 3000
        TotalAncho = TotalAncho + DataGrid1.Columns(i).Width
    
    'El importe es campo calculado
    i = 2
        DataGrid1.Columns(i).Visible = False
        
    i = 3
        DataGrid1.Columns(i).Caption = "Tipo concepto"
        DataGrid1.Columns(i).Width = 1500
            
        
        'Fiajamos el cadancho
    If Not CadAncho Then
        'La primera vez fijamos el ancho y alto de  los txtaux
        txtAux(0).Width = DataGrid1.Columns(0).Width - 60
        txtAux(1).Width = DataGrid1.Columns(1).Width - 60
        Combo1.Width = DataGrid1.Columns(3).Width
        txtAux(0).Left = DataGrid1.Left + 340
        txtAux(1).Left = txtAux(0).Left + txtAux(0).Width + 45
        Combo1.Left = txtAux(1).Left + txtAux(1).Width + 45
        CadAncho = True
    End If
    'Habilitamos modificar y eliminar
    If vUsu.Nivel < 2 Then
        Toolbar1.Buttons(7).Enabled = Not adodc1.Recordset.EOF
        Toolbar1.Buttons(8).Enabled = Not adodc1.Recordset.EOF
    End If
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
With txtAux(Index)
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtAux_LostFocus(Index As Integer)

txtAux(Index).Text = Trim(txtAux(Index).Text)
If txtAux(Index).Text = "" Then Exit Sub
If Modo = 3 Then Exit Sub 'Busquedas
If Index = 0 Then
    If Not IsNumeric(txtAux(0).Text) Then
        MsgBox "Código concepto tiene que ser numérico", vbExclamation
        Exit Sub
    End If
    txtAux(0).Text = Format(txtAux(0).Text, "000")
End If
End Sub


Private Function DatosOk() As Boolean
Dim Datos As String
Dim B As Boolean
DatosOk = False
B = CompForm(Me)
If Not B Then Exit Function



'Un USUARIO de un nivel "n" no puede poner un nivel
' n-1
Datos = Combo1.ItemData(Combo1.ListIndex)
If Val(Datos) < vUsu.Nivel Then
    MsgBox "No tiene privilegios para fijar ese nivel al grupo", vbExclamation
    Exit Function
End If


If Modo = 1 Then
    'Comprobamos k el numero NO se > 32
    If Val(txtAux(0).Text) > vbMaxGrupos Then
        MsgBox "Numero de grupo mayor del permitido: " & vbMaxGrupos, vbExclamation
        B = False
    
    Else
        'Estamos insertando
         Datos = DevuelveDesdeBD("codgrupo", "grupos", "codgrupo", txtAux(0).Text, "N")
         If Datos <> "" Then
            MsgBox "Ya existe el grupo : " & txtAux(0).Text, vbExclamation
            B = False
        End If
    End If
End If
DatosOk = B
End Function

Private Sub CargaCombo()
    Combo1.Clear
    'Conceptos
    Set miRSAux = New ADODB.Recordset
    miRSAux.Open "Select * from descniveles order by codnivel", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRSAux.EOF
        Combo1.AddItem miRSAux!descripcion
        Combo1.ItemData(Combo1.NewIndex) = miRSAux!codnivel
        miRSAux.MoveNext
        
    Wend
    miRSAux.Close
    Set miRSAux = Nothing
    
End Sub


Private Sub PonerOpcionesMenu()
PonerOpcionesMenuGeneral Me
End Sub



Private Function SepuedeBorrar() As Boolean
Dim SQL As String
    SepuedeBorrar = False
    SQL = DevuelveDesdeBD("codgrupo", "usuariosgrupos", "codgrupo", adodc1.Recordset!codgrupo, "N")
    If SQL <> "" Then
        MsgBox "Hay usuarios que pertenecen a este grupo", vbExclamation
        Exit Function
    End If
    SQL = DevuelveDesdeBD("codigo", "timagen", "groupprop", adodc1.Recordset!codgrupo, "N")
    If SQL <> "" Then
        MsgBox "Hay archivos pertenecientes al grupo", vbExclamation
        Exit Function
    End If
    SQL = DevuelveDesdeBD("codcarpeta", "carpetas", "groupprop", adodc1.Recordset!codgrupo, "N")
    If SQL <> "" Then
        MsgBox "Hay carpetas pertenecientes al grupo", vbExclamation
        Exit Function
    End If
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



