VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmResultados2 
   Caption         =   "Resultados de la busqueda"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10395
   Icon            =   "frmResultados.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   10395
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   2400
      Width           =   9135
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   4680
         TabIndex        =   8
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdMover 
         Caption         =   "Mover"
         Height          =   375
         Left            =   5760
         TabIndex        =   7
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdRegresar 
         Caption         =   "Regresar"
         Height          =   375
         Left            =   3240
         TabIndex        =   6
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "Eliminar"
         Height          =   375
         Left            =   6720
         TabIndex        =   5
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdVer 
         Caption         =   "Ver"
         Height          =   375
         Left            =   960
         TabIndex        =   4
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Propiedades"
         Height          =   375
         Left            =   1920
         TabIndex        =   3
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "Salir"
         Default         =   -1  'True
         Height          =   375
         Left            =   8040
         TabIndex        =   2
         Top             =   120
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   0
         Left            =   120
         Picture         =   "frmResultados.frx":030A
         Top             =   240
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   1
         Left            =   480
         Picture         =   "frmResultados.frx":0454
         Top             =   240
         Width           =   240
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5655
      Left            =   600
      TabIndex        =   0
      Top             =   0
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   9975
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Archivo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Carpeta"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "CodCarpeta"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "EscrituraCarpeta"
         Object.Width           =   0
      EndProperty
   End
End
Attribute VB_Name = "frmResultados2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private PrimeraVez As Boolean
Public DesdeEmail As Boolean

Private Sub cmdEliminar_Click()
    If ModoTrabajo <> vbNorm Then
        Mensajes1 15
        Exit Sub
    End If
    MsgBox "Opcion no disponible", vbExclamation
End Sub

Private Sub cmdImprimir_Click()
Dim J As Integer
Dim i As Byte
Dim C As Ccarpetas

        BorrarTemporal1
            
        i = 0
        For J = 1 To ListView1.ListItems.Count
            If ListView1.ListItems(J).Checked Then
                InsertaTemporal (CLng(Mid(ListView1.ListItems(J).Key, 2)))
                i = 1
            End If
        Next J
                
        If i = 0 Then
            MsgBox "Seleccione algun archivo para imprimir", vbExclamation
            Exit Sub
        End If
        
        
        
        ImprimirDesdeTablaTemporal Me, (ModoTrabajo = vbHistAnt Or ModoTrabajo = vbHistNue)
        
       
End Sub

Private Sub cmdMover_Click()
Dim i As Integer
Dim C As String
Dim vD As Ccarpetas
Dim vO As Ccarpetas
Dim Rs As ADODB.Recordset
Dim J As Integer

    If ModoTrabajo <> vbNorm Then
        Mensajes1 15
        Exit Sub
    End If

    C = ""
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked Then C = C & "1"
    Next i
    If C = "" Then
        MsgBox "Selecciona algun archivo", vbExclamation
        Exit Sub
        
    Else
        If Len(C) = 1 Then
            C = "el archivo"
        Else
            C = "los archivos(" & Len(C) & ")"
        End If
        C = "Seguro que desea mover de carpeta " & C & "?"
        If MsgBox(C, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
        
        
    DatosCopiados = ""
    frmPregunta.Opcion = 20
    frmPregunta.origenDestino = ""
    frmPregunta.Show vbModal
    If DatosCopiados = "" Then Exit Sub
    
    C = "La carpeta donde se moverán los archivos sera:" & vbCrLf & vbCrLf
    C = C & "        " & RecuperaValor(DatosCopiados, 3)
    i = Val(Mid(RecuperaValor(DatosCopiados, 1), 2))
    C = C & vbCrLf & vbCrLf & "¿Desea continuar"
    If MsgBox(C, vbQuestion + vbYesNoCancel) <> vbYes Then C = ""
    DatosCopiados = ""  'Para que cuando vuelva no haga nada
    If C = "" Then Exit Sub
    
    'Borramos de tmpBusqueda los que NO estan marcados
    
    Set vD = New Ccarpetas
    Set vO = New Ccarpetas
    Set Rs = New ADODB.Recordset
    
    If vD.Leer(i, (ModoTrabajo = 1)) = 0 Then
    
        'HACEMOS MOVER
        BorrarTemporal1
        C = "DELETE FROM tmpbusqueda where codusu = " & vUsu.codusu & " AND codequipo= " & vUsu.PC
        C = C & " AND Imagen = "
        For i = 1 To ListView1.ListItems.Count
            If Not ListView1.ListItems(i).Checked Then Conn.Execute C & Mid(ListView1.ListItems(i).Key, 2)
        Next i
        
        C = "SELECT codcarpeta FROM tmpbusqueda where codusu = " & vUsu.codusu & " AND codequipo= " & vUsu.PC
        C = C & " GROUP BY codcarpeta"
        
        Rs.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Set listacod = New Collection
        Set listaimpresion = New Collection
        While Not Rs.EOF
            listacod.Add Val(Rs!codcarpeta)
            Rs.MoveNext
        Wend
        Rs.Close
        
        For i = 1 To listacod.Count
            If vO.Leer(listacod(i), (ModoTrabajo = 1)) = 0 Then
                If vO.codcarpeta <> vD.codcarpeta Then
                    BorrarTemporal1
                    C = "INSERT INTO tmpFich(codusu,codequipo,imagen) Select codusu,codequipo,imagen FROM "
                    C = C & " tmpBusqueda where codusu = " & vUsu.codusu & " AND codequipo= " & vUsu.PC
                    C = C & " AND codcarpeta = " & vO.codcarpeta
                    Conn.Execute C
                
                    Set frmMovimientoArchivo.vDestino = vD
                    Set frmMovimientoArchivo.vOrigen = vO
                    frmMovimientoArchivo.Destino = vD.codcarpeta
                    frmMovimientoArchivo.Opcion = 5
                    frmMovimientoArchivo.Show vbModal
                    If DatosCopiados <> "" Then
                        GoTo Final
                    Else
                        C = "SELECT IMAGEN FROM tmpFich where codusu = " & vUsu.codusu & " AND codequipo= " & vUsu.PC
                        Rs.Open C, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                        While Not Rs.EOF
                            listaimpresion.Add Val(Rs.Fields(0))
                            Rs.MoveNext
                        Wend
                        Rs.Close
                        
                    End If
                        
                End If  'Que sean la misma carpeta
            End If
        Next
        
        'Ahora se trata de eliminar todas las entradas que hayan sido movidas de carpeta
        'Voy a reutilizar variables
        For i = 1 To listaimpresion.Count
            C = "C" & listaimpresion.Item(i)
            For J = 1 To ListView1.ListItems.Count
                If ListView1.ListItems(J).Key = C Then
                    ListView1.ListItems.Remove J
                    Exit For
                End If
            Next J
        Next i
        
    End If
    
Final:
    Set vO = Nothing
    Set Rs = Nothing
    Set vD = Nothing
    
End Sub

Private Sub cmdRegresar_Click()
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    DatosCopiados = "C" & ListView1.SelectedItem.SubItems(2) & "|" & ListView1.SelectedItem.Key & "|"
    Unload Me
End Sub

Private Sub cmdVer_Click()
Dim mCarpetas As Ccarpetas
Dim miEXT As Cextensionpc
Dim Destini As String
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    Set mCarpetas = New Ccarpetas
    If mCarpetas.Leer(Val(ListView1.SelectedItem.SubItems(2)), (ModoTrabajo = vbHistNue Or ModoTrabajo = vbHistAnt)) = 0 Then
        Set miEXT = New Cextensionpc
            If miEXT.Leer(Val(ListView1.SelectedItem.SmallIcon - 1), vUsu.PC) = 0 Then
                If miEXT.pathexe <> "" Then
                    If DevuelveNombreFichero(ListView1.SelectedItem.Text, miEXT.Extension, Destini, False) < 101 Then
                        Admin.AbirFichero True, mCarpetas, CLng(Mid(ListView1.SelectedItem.Key, 2)), Destini, miEXT, 0, False, 0
                    Else
                        MsgBox "No se ha podido traer archivo", vbExclamation
                    End If
                Else
                    MsgBox "La extensión no tiene path asociado", vbExclamation
                End If
            End If
        Set miEXT = Nothing
    End If
    Set mCarpetas = Nothing
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
Dim mCarpetas As Ccarpetas
Dim vImg As cTimagen
Dim LecturaSolo As Boolean

    If ListView1.SelectedItem Is Nothing Then Exit Sub
   
'    Admin.VerPropiedades ListView1.SelectedItem.Key, True
    Screen.MousePointer = vbHourglass ' luego en el form.load lo ponemos a normal
    Set vImg = New cTimagen
    
    Set mCarpetas = New Ccarpetas
    
    If mCarpetas.Leer(Val(ListView1.SelectedItem.SubItems(2)), (ModoTrabajo <> vbNorm)) = 0 Then
    
    
        If vImg.Leer(CLng(Mid(ListView1.SelectedItem.Key, 2)), objRevision.LlevaHcoRevision) = 0 Then
            
   
        If vImg.userprop = vUsu.codusu Or (vImg.escriturag And vUsu.Grupo) Or vUsu.codusu = 0 Then LecturaSolo = False
 
        If ModoTrabajo <> vbNorm Then LecturaSolo = True

        If LecturaSolo Then
            frmNuevoArchivo.Opcion = 3
        Else
            frmNuevoArchivo.Opcion = 2
        End If
        Set frmNuevoArchivo.mImag = vImg
        frmNuevoArchivo.Carpeta = mCarpetas.Nombre
        frmNuevoArchivo.Show vbModal
        End If
        
    End If
    Set mCarpetas = Nothing
    Set vImg = Nothing
End Sub





Private Sub Command3_Click()
    
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        
        '---------------------------------------
        CargarDatos

        'Si es hco no dejamos que modifique ni borre
        If (ModoTrabajo <> vbNorm) Or DesdeEmail Then
            cmdMover.Enabled = False
            cmdEliminar.Enabled = False
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    Set Me.ListView1.SmallIcons = Admin.ImageList2
    ListView1.Checkboxes = Not DesdeEmail
    PrimeraVez = True
End Sub


Private Sub CargarDatos()
Dim cad As String
Dim Itm As ListItem

    If ModoTrabajo = vbNorm Then


        cad = "Select carpetas.nombre,carpetas.codcarpeta,timagen.campo1,timagen.codigo,codext,"
        cad = cad & "timagen.escriturag as permisoU,carpetas.escriturag as permisoG from"
        cad = cad & " carpetas,timagen,tmpbusqueda  WHERE "
        cad = cad & "tmpbusqueda.codusu = " & vUsu.codusu & " and tmpbusqueda.codequipo =" & vUsu.PC
        cad = cad & " and tmpbusqueda.imagen=codigo and tmpbusqueda.codcarpeta =carpetas.codcarpeta"
        
    
    Else
    
        'HISTORICO  #####################
        
        cad = "Select carpetashco.nombre,carpetashco.codcarpeta,timagenhco.campo1,timagenhco.codigo,codext,"
        cad = cad & "timagenhco.escriturag as permisoU,carpetashco.escriturag as permisoG from"
        cad = cad & " carpetashco,timagenhco,tmpbusqueda  WHERE "
        cad = cad & "tmpbusqueda.codusu = " & vUsu.codusu & " and tmpbusqueda.codequipo =" & vUsu.PC
        cad = cad & " and tmpbusqueda.imagen=codigo and tmpbusqueda.codcarpeta =carpetashco.codcarpeta"
        'Carpetas solo para usuario
        cad = cad & " and carpetashco.codequipo = " & vUsu.PC
        
    End If
    Set miRSAux = New ADODB.Recordset
    miRSAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic
    While Not miRSAux.EOF
        Set Itm = ListView1.ListItems.Add(, "C" & miRSAux!codigo)
        Itm.Text = miRSAux!campo1
        Itm.SubItems(1) = miRSAux!Nombre
        Itm.SubItems(2) = miRSAux!codcarpeta
        Itm.SubItems(3) = miRSAux!permisog
        'El tag tiene permiso escritura del fichero
        Itm.Tag = miRSAux!permisou
        
        Itm.SmallIcon = miRSAux!codext + 1
        
        miRSAux.MoveNext
    Wend
    miRSAux.Close
    Set miRSAux = Nothing
End Sub

Private Sub Form_Resize()
    If Me.Width < Frame1.Width Then
        Frame1.Left = 30
    Else
        Frame1.Left = Me.Width - Frame1.Width - 30
    End If
    If Me.Height < 500 Then
        Frame1.Top = 500
    Else
        Frame1.Top = Me.Height - Me.Frame1.Height - 390
    End If
    ListView1.Left = 60
    ListView1.Width = Me.Width - ListView1.Left - 320
    ListView1.Height = Frame1.Top - 120 - ListView1.Top
    ListView1.ColumnHeaders(1).Width = CInt(((ListView1.Width - 400) / 3) * 2)
    ListView1.ColumnHeaders(2).Width = CInt(((ListView1.Width - 400) / 3))
End Sub




Private Sub Image1_Click(Index As Integer)
    Dim i As Integer
    For i = 1 To ListView1.ListItems.Count
        ListView1.ListItems(i).Checked = Index = 0
    Next i
End Sub
