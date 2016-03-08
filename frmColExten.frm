VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmColExten 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Extensiones ARIDOC"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6285
   Icon            =   "frmColExten.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   6285
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1320
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdUsuario 
      Height          =   375
      Index           =   0
      Left            =   3240
      Picture         =   "frmColExten.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdUsuario 
      Height          =   375
      Index           =   1
      Left            =   3720
      Picture         =   "frmColExten.frx":040C
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdUsuario 
      Height          =   375
      Index           =   2
      Left            =   4200
      Picture         =   "frmColExten.frx":050E
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   120
      Width           =   375
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4455
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   6015
      Begin VB.CheckBox Check4 
         Caption         =   "Deshabilitar"
         Height          =   255
         Left            =   1080
         TabIndex        =   25
         Top             =   3840
         Width           =   2655
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Aparece en el menu"
         Height          =   255
         Left            =   1080
         TabIndex        =   24
         Top             =   3360
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   5
         Left            =   1200
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   2160
         Width           =   975
      End
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         Caption         =   "Plantilla"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   1560
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   2760
         Width           =   4215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   1200
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   1680
         Width           =   4575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   4800
         TabIndex        =   12
         Top             =   3840
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   0
         Left            =   4800
         TabIndex        =   11
         Top             =   3360
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Modificable"
         Height          =   255
         Left            =   2520
         TabIndex        =   10
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   2880
         MaxLength       =   4
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   1200
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1200
         Width           =   4575
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Imprimir"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   23
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Codigo"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   735
      End
      Begin VB.Image Image4 
         Height          =   240
         Left            =   1245
         Picture         =   "frmColExten.frx":0610
         Top             =   2760
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Left            =   840
         Picture         =   "frmColExten.frx":0712
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "EXE "
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   1680
         Width           =   540
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   3
         Height          =   4455
         Left            =   0
         Top             =   0
         Width           =   6015
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   4440
         Picture         =   "frmColExten.frx":0814
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Icono"
         Height          =   195
         Index           =   0
         Left            =   3840
         TabIndex        =   16
         Top             =   720
         Width           =   405
      End
      Begin VB.Image Image1 
         Height          =   495
         Left            =   4800
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Extension"
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   9
         Top             =   735
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Descripcion"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   1215
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   5160
      Width           =   975
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4455
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   7858
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cod."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descripcion"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Exten."
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Extensiones configuradas para ARIDOC"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   2835
   End
End
Attribute VB_Name = "frmColExten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Itm As ListItem
Dim Cad As String
Dim RecargarIconos As Boolean
Dim SeAñadenNuevas As Boolean  'Pq habra k reiniciar la apliacion
Dim NFich As String  'Para el ciono

Private Sub Check2_Click()
    Text1(4).Visible = Check2.Value = 1
End Sub

Private Sub cmdUsuario_Click(Index As Integer)
    If Index > 0 Then
        If ListView1.SelectedItem Is Nothing Then Exit Sub
    End If

    

    Select Case Index
    Case 0, 1
            If Index = 0 Then
                Limpiar Me
                PonerSiguiente
                Label2.Caption = "INSERTAR"
                Label2.ForeColor = vbBlue
                Label2.Tag = 0
                Image1.Tag = 0
                Check2.Value = 0
                Check2.Tag = 0
            Else
                If Not PonerDatosExtension Then Exit Sub
                Label2.Caption = "MODIFICAR"
                Label2.ForeColor = vbBlack
                Label2.Tag = 1
            End If
            NFich = ""
            Text1(0).Enabled = Index = 0
            Text1(2).Enabled = Index = 0
            Text1(4).Visible = Check2.Value = 1
            Text1(4).Visible = False
            PonerFrames True
    Case 2
         Cad = "Si va a realizar traspasos a historico de documetos de aridoc"
         Cad = Cad & " NO deberia eliminar ninguna extension."
         Cad = Cad & vbCrLf & vbCrLf & vbCrLf
         Cad = Cad & "¿Desea eliminar la extension " & ListView1.SelectedItem.SubItems(2) & " - "
         Cad = Cad & ListView1.SelectedItem.SubItems(1) & "?"
         If MsgBox(Cad, vbQuestion + vbYesNo) <> vbYes Then Exit Sub
             
             
         Set miRSAux = New ADODB.Recordset
         NFich = "NO"
         If EliminarExtension Then
            
            NFich = ""
         End If
         Set miRSAux = Nothing
         
         If NFich = "" Then Recargar
        
            
        
         
    End Select

End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click(Index As Integer)

    If Index = 1 Then
        PonerFrames False
    Else
        'Si es nuevo o modificar
        If Not DatosOk Then Exit Sub
        'CopiaNuevoICono
        Cad = ""
        If Me.Image1.Tag = -1 Then
            If Not CopiaNuevoIcono Then Cad = "NO"
        End If
        If Cad <> "" Then
            'No se ha podido llevar el icono"
            Exit Sub
        End If
        
        
        'Llevar plantilla
        Cad = ""
        If Check2.Value = 1 Then
            If Text1(4).Text <> "" Then
                'Llevamos plantilla
                If Not LlevaElArchivoPlantilla Then Cad = "NO"
            End If
        End If
        If Cad <> "" Then
            Exit Sub
        End If
        '---------------------
        If Label2.Tag = 0 Then
            
            Cad = "INSERT INTO extension (codext, descripcion, exten, Modificable, Nuevo, OfertaExe, OfertaPrint,Aparecemenu,Deshabilitada) VALUES ("
            Cad = Cad & Text1(0).Text & ",'" & DevNombreSql(Text1(1).Text) & "','" & Text1(2).Text & "',"
            Cad = Cad & Abs(Val(Me.Check1.Value)) & ","
            
            'Nuevo el 25 abril
            Cad = Cad & Abs(Val(Me.Check2.Value)) & ","
            If Text1(3).Text = "" Then
                Cad = Cad & "NULL"
            Else
                Cad = Cad & "'" & DevNombreSql(Text1(3).Text) & "'"
            End If
            Cad = Cad & ","
            If Text1(3).Text <> "" And Text1(5).Text <> "" Then
                Cad = Cad & "'" & DevNombreSql(Text1(3).Text) & " " & DevNombreSql(Text1(5).Text) & "'"
            Else
                Cad = Cad & "NULL"
            End If
            Cad = Cad & ","
            
            'Aparecemenu
            If Check2.Value = 1 And Check3.Value = 1 Then
                Cad = Cad & "1"
            Else
                Cad = Cad & "0"
            End If
            'Deshabilitada
            If Check4.Value Then
                Cad = Cad & ",1"
            Else
                Cad = Cad & ",0"
            End If
            
            Cad = Cad & ")"
            
            
            
            SeAñadenNuevas = True
        Else
            Cad = "UPDATE extension set descripcion='" & DevNombreSql(Text1(1).Text) & "',"
            Cad = Cad & "Modificable = " & Abs(Val(Me.Check1.Value))
            Cad = Cad & ",Nuevo = " & Abs(Val(Me.Check2.Value))
            ' ,
            Cad = Cad & ",OfertaExe = "
            If Text1(3).Text = "" Then
                Cad = Cad & "NULL"
            Else
                Cad = Cad & "'" & DevNombreSql(Text1(3).Text) & "'"
            End If
            Cad = Cad & ",OfertaPrint = "
            If Text1(3).Text <> "" And Text1(5).Text <> "" Then
                Cad = Cad & "'" & DevNombreSql(Text1(3).Text) & " " & DevNombreSql(Text1(5).Text) & "'"
            Else
                Cad = Cad & "NULL"
            End If
            Cad = Cad & ",Aparecemenu ="
            If Check2.Value = 1 And Check3.Value = 1 Then
                Cad = Cad & "1"
            Else
                Cad = Cad & "0"
            End If
            
            Cad = Cad & ",Deshabilitada ="
            If Check4.Value = 0 Then
                Cad = Cad & "0"
            Else
                Cad = Cad & "1"
            End If
            
            
            
            Cad = Cad & " WHERE codext =" & Text1(0).Text
        End If
        Conn.Execute Cad
        
        'Si es nueva, cargamos, para cada equipo, las extensiones a blanco
        If Label2.Tag = 0 Then CopiaTblaExtensionPC
        
        'Volvemos a recargar
        Recargar
        PonerFrames False
    End If
    
    


End Sub


Private Sub Recargar()
        Set Me.ListView1.SmallIcons = Nothing
        CargarListviews
        CargaDatos

End Sub

Private Function DatosOk() As Boolean
Dim i As Integer
    DatosOk = False
    For i = 0 To 2
        Text1(i).Text = Trim(Text1(i).Text)
        If Text1(i).Text = "" Then
            MsgBox "Campos requeridos", vbExclamation
            Exit Function
        Else
            If i = 0 Then
                If Not IsNumeric(Text1(0).Text) Then
                    MsgBox "Campo codigo debe ser numérico", vbExclamation
                    Exit Function
                End If
            End If
        End If
    Next i
    
    If Check2.Tag = 0 Then
        If Check2.Value = 1 And Text1(4).Text = "" Then
            MsgBox "Ha marcado la plantilla pero no ha indicado ningun archivo", vbExclamation
            Exit Function
        End If
    End If
    
    If Text1(4).Text <> "" Then
        If Dir(Text1(4).Text, vbArchive) = "" Then
            MsgBox "Archivo: " & Text1(4).Text & " NO se ha encontrado", vbExclamation
            Exit Function
        End If
    End If
    If Label2.Tag = 0 Then
        'Es nevo. Con lo cual la extension no debe existir
        Cad = DevuelveDesdeBD("codext", "extension", "codext", Text1(0).Text)
        If Cad <> "" Then
            MsgBox "Ya existe la extension : " & Cad, vbExclamation
            Exit Function
        End If
    Else
        'Si no tenia plantilla

                
                
        
    End If
    
    'Si el tag de imagen es 0 den MAL
    If Image1.Tag = 0 Then
        MsgBox "Error en el icono", vbExclamation
        Exit Function
    End If
    
    If Image1.Tag = -1 Then
        If NFich = "" Then
            MsgBox "Error en icono", vbExclamation
            Exit Function
        Else
            If Dir(NFich, vbArchive) = "" Then
                MsgBox "Error en icono", vbExclamation
                Exit Function
            End If
        End If
    End If
    DatosOk = True
    
End Function

Private Sub Form_Load()
    Recargar
    SeAñadenNuevas = False
    RecargarIconos = False
    PonerFrames False
End Sub

Private Sub PonerFrames(ModificarAnyadir As Boolean)
    Frame1.Visible = ModificarAnyadir
    Me.Command1.Enabled = Not ModificarAnyadir
    Me.cmdUsuario(0).Enabled = Not ModificarAnyadir
    Me.cmdUsuario(1).Enabled = Not ModificarAnyadir
    Me.cmdUsuario(2).Enabled = Not ModificarAnyadir
    ListView1.Enabled = Not ModificarAnyadir
    Me.Refresh
End Sub

Private Sub CargaDatos()
    ListView1.ListItems.Clear
    Set miRSAux = New ADODB.Recordset
    miRSAux.Open "Select * from Extension order by codext", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRSAux.EOF
        Set Itm = ListView1.ListItems.Add(, "C" & miRSAux!codext)
        Itm.Text = miRSAux!codext
        Itm.SubItems(1) = miRSAux!Descripcion
        Itm.SubItems(2) = miRSAux!Exten
        'El icono
        Itm.SmallIcon = miRSAux!codext + 1 '+1 pq el el primero es el defuatl
        
        If miRSAux!deshabilitada = 1 Then
            Itm.Ghosted = True
            Itm.Bold = True
        End If
        
        miRSAux.MoveNext
    Wend
    miRSAux.Close
    Set miRSAux = Nothing
End Sub


Private Sub PonerSiguiente()
Dim i As Integer

    Set miRSAux = New ADODB.Recordset
    miRSAux.Open "Select max(codext) from  extension  ", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = 0
    If Not miRSAux.EOF Then
        i = DBLet(miRSAux.Fields(0), "N")
    End If
    miRSAux.Close
    Set miRSAux = Nothing
    i = i + 1
    Text1(0).Text = i
End Sub


Private Function PonerDatosExtension() As Boolean
Dim i As Integer

    PonerDatosExtension = False
    Set miRSAux = New ADODB.Recordset
    miRSAux.Open "Select * from  extension where codext = " & ListView1.SelectedItem.Text, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If miRSAux.EOF Then
        MsgBox "Error leyendo extension: " & ListView1.SelectedItem.Text & " - " & ListView1.SelectedItem.SubItems(1), vbExclamation
    
    Else
        'Cargamos los datos
        Text1(0).Text = miRSAux!codext
        Text1(1).Text = miRSAux!Descripcion
        Text1(2).Text = miRSAux!Exten
        Text1(4).Text = ""
        If miRSAux!modificable = 0 Then
            Check1.Value = 0
        Else
            Check1.Value = 1
        End If
        
        Text1(3).Text = DBLet(miRSAux!ofertaexe, "T")
        If Text1(3).Text = "" Then
            Text1(5).Text = ""
        Else
            Cad = DBLet(miRSAux!ofertaprint, "T")
            If Cad <> "" Then
                i = InStr(1, Cad, Text1(3).Text)
                If i = 0 Then
                    Cad = ""
                Else
                    Cad = Trim(Mid(Cad, i + Len(Text1(3).Text) + 1))
                End If
            End If
            Text1(5).Text = Cad
        End If
        If miRSAux!Nuevo = 1 Then
            Check2.Value = 1
        Else
            Check2.Value = 0
        End If
        Check2.Tag = Check2.Value
        
        CargaIcono miRSAux!codext
        
        If miRSAux!deshabilitada = 1 Then
            Check4.Value = 1
        Else
            Check4.Value = 0
        End If
        
        PonerDatosExtension = True
    End If
    miRSAux.Close
    Set miRSAux = Nothing
    
End Function

Private Sub CargaIcono(Referencia As Integer)
    On Error GoTo CargaIcono
    'Vaciamos
    Me.Image1.Tag = 0
    Me.Image1.Picture = LoadPicture()
    
    'Cargamos la referencia
    Me.Image1.Picture = LoadPicture(App.Path & "\imagenes\" & Referencia & ".ico")
    Me.Image1.Tag = Referencia
    
CargaIcono:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description, "Cargando ICONO"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If RecargarIconos Then Conn.Execute " UPDATE equipos SET cargaIconsExt= 0 WHERE codequipo<>1"
    If RecargarIconos Or SeAñadenNuevas Then
        MsgBox "Debe reiniciar la aplicacion", vbCritical
        End
    End If
End Sub

Private Sub Image2_Click()

    Me.CommonDialog1.Filter = "Iconos *.ICO|*.ico|"
    Me.CommonDialog1.ShowOpen
    If Me.CommonDialog1.FileName <> "" Then
            'Compruebo que la extension es ICO
            If UCase(Right(Me.CommonDialog1.FileName, 3)) <> "ICO" Then
                MsgBox "No es un icono", vbExclamation
            Else
                'Cargo ICONO desde aqui
                CargaNuevoIcono Me.CommonDialog1.FileName
            End If
    End If
End Sub


Private Sub CargaNuevoIcono(Arch As String)
Dim OLD As Integer
    
    On Error GoTo ECA
    If Label2.Tag = 0 Then
        OLD = 0
    Else
        OLD = CInt(Text1(0).Text)
    End If
    
    Me.Image1.Picture = LoadPicture(Arch)
    NFich = Arch
    Me.Image1.Tag = -1
        
    Exit Sub
ECA:
    MuestraError Err.Number, Err.Description
    NFich = ""
    If OLD = 0 Then
        Me.Image1.Tag = 0
        Me.Image1.Picture = LoadPicture()
    Else
        CargaIcono OLD
    End If
End Sub


Private Sub CargarListviews()
Dim i As Integer
Dim J As Integer
Dim Pos As Integer
Dim SQL As String
Dim Errores As String
    Errores = False
    If Dir(App.Path & "\Defaultico.dat", vbArchive) = "" Then
        MsgBox "Archivos necesarios para la aplicacion han sido borrados(Defaultico.dat)", vbCritical
        End
    End If

    ImageList1.ListImages.Clear
    ImageList1.ImageHeight = 16
    ImageList1.ImageWidth = 16
'Cargamos IMAGELIST con los iconos de las imagenes
    SQL = "SELECT * FROM extensionpc where codequipo=" & vUsu.PC & "  order by codext"
    Set miRSAux = New ADODB.Recordset
    miRSAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Errores = ""
    Pos = 1
    
    While Not miRSAux.EOF
        J = miRSAux!codext + 1
        If J > Pos Then
            For i = Pos To J - 1
                CargaIconoF i, App.Path & "\Defaultico.dat"
            Next i
        End If
        Pos = J + 1
            
        'Cargamos el icono
        SQL = App.Path & "\imagenes\" & miRSAux!codext & ".ico"
        If Dir(SQL) <> "" Then
            
        Else
            Errores = Errores & SQL & vbCrLf
            SQL = App.Path & "\Defaultico.dat"
        End If
        CargaIconoF J, SQL
    
        'Siguiente
        miRSAux.MoveNext
    Wend
    miRSAux.Close
    Set miRSAux = Nothing
    
    'Si han habido errores
    'Proponemos carga iconos
    If Errores <> "" Then
        vUsu.CargaIconosExtensiones = True
        Conn.Execute "UPDATE equipos SET cargaIconsExt= 1 WHERE codequipo=" & vUsu.PC
        SQL = "Se han producido errores cargando iconos. Reinicie la aplicacion y si continua el problema consulte con el soporte técnico:" & vbCrLf & Errores
        SQL = SQL & vbCrLf & vbCrLf & "¿Finalizar?"
        If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then End
    Else
        If ImageList1.ListImages.Count > 0 Then ListView1.SmallIcons = ImageList1
    End If
End Sub


Private Sub CargaIconoF(Cod As Integer, vpath As String)
    ImageList1.ListImages.Add , "C" & Cod, LoadPicture(vpath)
End Sub



Private Function CopiaNuevoIcono() As Boolean

    On Error GoTo ECopiaNuevoIcono
    CopiaNuevoIcono = False
    
    'Primero lo cpopiaremos en \imagenes\codex.ico
    FileCopy NFich, App.Path & "\imagenes\" & Text1(0).Text & ".ico"
    
    'AHora lo llevaremos al SRV
    DatosCopiados = "NO"
    
    
    frmMovimientoArchivo.Opcion = 3   'Llevar icono
    frmMovimientoArchivo.Origen = App.Path & "\imagenes\" & Text1(0).Text & ".ico"
    frmMovimientoArchivo.Destino = Text1(0).Text & ".ico"
    frmMovimientoArchivo.Show vbModal
    
    If DatosCopiados = "" Then
        CopiaNuevoIcono = True
        RecargarIconos = True
    End If
        
    
    Exit Function
ECopiaNuevoIcono:
        MuestraError Err.Number, "Copiando fichero: " & NFich
End Function


'--------------------------------
' La extension para cada PC
Private Sub CopiaTblaExtensionPC()
    Set miRSAux = New ADODB.Recordset
    miRSAux.Open "Select codequipo from equipos", Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Cad = "INSERT INTO extensionpc (codext,  pathexe, impresion,codequipo) VALUES ("
    Cad = Cad & Text1(0).Text & ",'','',"
    While Not miRSAux.EOF
        Conn.Execute Cad & miRSAux.Fields(0) & ")"
        miRSAux.MoveNext
    Wend
    miRSAux.Close
    Set miRSAux = Nothing
    
    

End Sub


Private Function EliminarExtension() As Boolean
Dim i As Long
    
    On Error GoTo EEliminarExtension
    
    
    EliminarExtension = False
    
    i = 0
    Cad = "Select count(*) from TIMAGEN where codext = " & ListView1.SelectedItem.Text
    miRSAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRSAux.EOF Then
        i = DBLet(miRSAux.Fields(0), "N")
    End If
    miRSAux.Close
    
    If i > 0 Then
        MsgBox "Existen archivos con esa extension", vbExclamation
        Exit Function
    End If
    
        
        
        
     'habra k comprobar mas cosas
     'Vemos en las  plantillas si hay alguna
    Cad = "Select count(*) from plantilla where tipo = " & ListView1.SelectedItem.Text
    miRSAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    i = 0
    If Not miRSAux.EOF Then
        i = DBLet(miRSAux.Fields(0), "N")
    End If
    miRSAux.Close
         
    If i > 0 Then
        MsgBox "Existen plantillas con esa extension", vbExclamation
        Exit Function
    End If
         
         
     
     'Eliminamos
     Cad = "DELETE from extensionpc where codext =" & ListView1.SelectedItem.Text
     Conn.Execute Cad
     
     
     Cad = "DELETE from extension where codext =" & ListView1.SelectedItem.Text
     Conn.Execute Cad
     
     
     'FIN
     EliminarExtension = True
    Exit Function
EEliminarExtension:
    MuestraError Err.Number
End Function





Private Sub Image3_Click()
    Me.CommonDialog1.Filter = "Ejecutables *.EXE|*.exe|"
    Me.CommonDialog1.ShowOpen
    If Me.CommonDialog1.FileName <> "" Then
        Text1(3).Text = Me.CommonDialog1.FileName
    End If
End Sub

Private Sub Image4_Click()
    If Text1(1).Text <> "" Then
        If Text1(2).Text <> "" Then
            Me.CommonDialog1.Filter = Text1(1).Text & " *." & Text1(2).Text & "|*." & Text1(2).Text & "|"
            Me.CommonDialog1.ShowOpen
            If Me.CommonDialog1.FileName <> "" Then Text1(4).Text = Me.CommonDialog1.FileName
        End If
    End If
End Sub

Private Function LlevaElArchivoPlantilla() As Boolean
    LlevaElArchivoPlantilla = False
    frmMovimientoArchivo.Opcion = 13   'Llevar plantilla
    frmMovimientoArchivo.Origen = Text1(4).Text
    frmMovimientoArchivo.Destino = Text1(0).Text
    frmMovimientoArchivo.Show vbModal
    
    If DatosCopiados = "" Then LlevaElArchivoPlantilla = True

        
End Function
