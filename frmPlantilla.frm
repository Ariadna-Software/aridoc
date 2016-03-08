VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmPlantilla 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Archivos plantilla"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8625
   Icon            =   "frmPlantilla.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   8625
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameCarpeta 
      Height          =   4815
      Left            =   120
      TabIndex        =   18
      Top             =   0
      Width           =   8415
      Begin MSComctlLib.ListView ListView2 
         Height          =   4455
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   7858
         View            =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame FramePlant 
      Height          =   2775
      Left            =   1440
      TabIndex        =   0
      Top             =   1080
      Width           =   5415
      Begin VB.Frame FrameTapa 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   495
         Left            =   240
         TabIndex        =   15
         Top             =   960
         Width           =   5055
      End
      Begin VB.CheckBox chkModificarPlantilla 
         Caption         =   "Modificar el fichero"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   1080
         Width           =   3375
      End
      Begin MSComDlg.CommonDialog cd1 
         Left            =   120
         Top             =   2040
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdPlantilla 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   4080
         TabIndex        =   6
         Top             =   2280
         Width           =   975
      End
      Begin VB.CommandButton cmdPlantilla 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   0
         Left            =   3000
         TabIndex        =   5
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1680
         Width           =   4815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   480
         Width           =   4815
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   840
         Picture         =   "frmPlantilla.frx":030A
         Top             =   1440
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Plantilla"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Descripción"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   4815
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   8415
      Begin VB.Frame Frame1 
         Height          =   615
         Left            =   120
         TabIndex        =   10
         Top             =   4200
         Width           =   2775
         Begin VB.CommandButton cmdUsuario 
            Height          =   375
            Index           =   3
            Left            =   2000
            Picture         =   "frmPlantilla.frx":040C
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Mover a carpeta"
            Top             =   180
            Width           =   375
         End
         Begin VB.CommandButton cmdUsuario 
            Height          =   375
            Index           =   2
            Left            =   1080
            Picture         =   "frmPlantilla.frx":050E
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Eliminar"
            Top             =   180
            Width           =   375
         End
         Begin VB.CommandButton cmdUsuario 
            Height          =   375
            Index           =   1
            Left            =   600
            Picture         =   "frmPlantilla.frx":0610
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Modificar"
            Top             =   180
            Width           =   375
         End
         Begin VB.CommandButton cmdUsuario 
            Height          =   375
            Index           =   0
            Left            =   120
            Picture         =   "frmPlantilla.frx":0712
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Nueva"
            Top             =   180
            Width           =   375
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Regresar"
         Height          =   375
         Index           =   0
         Left            =   6000
         TabIndex        =   9
         Top             =   4320
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   7200
         TabIndex        =   8
         Top             =   4320
         Width           =   1095
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3735
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   6588
         View            =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Descripcion"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "General"
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
         Left            =   240
         TabIndex        =   17
         Top             =   120
         Width           =   7935
      End
   End
End
Attribute VB_Name = "frmPlantilla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Opcion As Byte
    '0.- Normal. Regresar o cancelar
    '1.- Mantenimiento
    
    
    
Private WithEvents frmP As frmPregunta
Attribute frmP.VB_VarHelpID = -1

Dim cad As String
Dim TieneCarpetas As Boolean

Private Sub chkModificarPlantilla_Click()
    If Me.chkModificarPlantilla.Value = 1 Then
        Text1(1).Enabled = True
    Else
        Text1(1).Enabled = False
        Text1(1).Text = ""
    End If
End Sub

Private Sub cmdPlantilla_Click(Index As Integer)
    If Index = 0 Then
        Text1(0).Text = Trim(Text1(0).Text)
        If Text1(0).Text = "" Then
            MsgBox "Escriba la descipcion", vbExclamation
            Exit Sub
        End If
            
        If Me.FrameTapa.Visible Or Me.chkModificarPlantilla.Value = 1 Then
            Text1(1).Text = Trim(Text1(1).Text)
            If Text1(1).Text = "" Then
                MsgBox "Selecccione un archivo", vbExclamation
                Exit Sub
            End If
            If Dir(Text1(1).Text, vbArchive) = "" Then
                MsgBox "No existe el archivo", vbExclamation
                Exit Sub
            End If
            
        End If
        
            'SUBIR Y UPDATEAR
            If Not SubirPlantilla Then Exit Sub
            
            
            'Refrescaremos listview
            
            'Pondremos los frames
            
            CargaList
        
    
    End If
    PonerFrames False
End Sub



Private Sub cmdUsuario_Click(Index As Integer)
    If Index > 0 Then
        If ListView1.ListItems.Count = 0 Then Exit Sub
        If TieneCarpetas Then
            If ListView1.SelectedItem.Tag = 0 Then Exit Sub
        End If
        If ListView1.SelectedItem Is Nothing Then
            MsgBox "Seleccione una plantilla", vbExclamation
            Exit Sub
        End If
        
    End If
    If Index < 2 Then
        Text1(0).Text = ""
        Text1(1).Text = ""
        'insertar modificar
        If Index = 1 Then
            Text1(0).Text = ListView1.SelectedItem.Text
            
            Text1(1).Enabled = False
            
        Else
            Text1(1).Enabled = True
        End If
        
        FrameTapa.Visible = Index = 0
        PonerFrames True
    Else
        'ELIMINAR
        If Index = 2 Then
            
            cad = "Seguro que desea eliminar la plantilla : " & ListView1.SelectedItem.Text & "?"
            If MsgBox(cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub
            EliminarPlantilla
        Else
            MoverPlantillas
            CargaList
        End If
    End If
End Sub

Private Sub Command1_Click(Index As Integer)

    DatosCopiados = ""
    If Index = 0 Then
    
        If TieneCarpetas Then
            If ListView1.SelectedItem.Tag = 0 Then
                Me.FrameCarpeta.Visible = True
                Exit Sub
            End If
        End If
    
    
    
        If ListView1.SelectedItem Is Nothing Then Exit Sub
        DatosCopiados = ListView1.SelectedItem.Tag
    End If
    Unload Me
End Sub

Private Sub Form_Load()
        
        
    
        

    Set Me.ListView1.Icons = Admin.ImageList2
    Set Me.ListView1.SmallIcons = Admin.ImageList3
    Set Me.ListView2.SmallIcons = Admin.ImgUsersPCs


    Me.ListView1.MultiSelect = Opcion = 1

    TieneCarpetas = TienePlantillasEnCarpeta2
    FrameCarpeta.Visible = TieneCarpetas
    cmdUsuario(3).Visible = TieneCarpetas
    If TieneCarpetas Then
        cargaCarpetas
    Else
        CargaList
    End If
    
    
    If Opcion = 1 Then Me.Command1(1).Caption = "Salir"
    
    
    Me.Command1(0).Visible = Opcion = 0
    Me.Frame1.Visible = Opcion = 1
    PonerFrames False
     
    
End Sub

Private Sub cargaCarpetas()
Dim Itm As ListItem
    On Error GoTo ECargaList
 
    ListView2.ListItems.Clear
    Set miRSAux = New ADODB.Recordset
    cad = "Select * from plantillacarpetas  "
    'Cad = Cad & "where ( lectura & " & vUsu.Grupo & ")"
    miRSAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRSAux.EOF
        Set Itm = ListView2.ListItems.Add(, "C: " & miRSAux!Carpeta)
        Itm.Text = miRSAux!Texto
        Itm.SmallIcon = "v_cerrado"
        Itm.Tag = miRSAux!Carpeta
        miRSAux.MoveNext
    Wend
    miRSAux.Close
    Set miRSAux = Nothing
    Me.Command1(0).Enabled = ListView1.ListItems.Count > 0
    Exit Sub
ECargaList:
    MuestraError Err.Number, "Carga List", Err.Description

End Sub

Private Sub CargaList()
Dim Itm As ListItem
    On Error GoTo ECargaList
 
    ListView1.ListItems.Clear
    Set miRSAux = New ADODB.Recordset
    cad = "Select * from plantilla  where 1=1"
    'Cad = Cad & "where ( lectura & " & vUsu.Grupo & ")"
    
    
    If TieneCarpetas Then
        cad = cad & " AND carpeta= " & ListView2.SelectedItem.Tag
        Set Itm = ListView1.ListItems.Add(, "C:0")
        Itm.Text = ".. VOLVER"
        Itm.Tag = 0
    End If
    miRSAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRSAux.EOF
        Set Itm = ListView1.ListItems.Add(, "C: " & miRSAux!codigo)
        Itm.Text = miRSAux!Descripcion
        Itm.SmallIcon = miRSAux!Tipo + 1
        Itm.Tag = miRSAux!codigo
        miRSAux.MoveNext
    Wend
    miRSAux.Close
    Set miRSAux = Nothing
    Me.Command1(0).Enabled = ListView1.ListItems.Count > 0
    Exit Sub
ECargaList:
    MuestraError Err.Number, "Carga List", Err.Description
End Sub

Private Sub frmP_DatoSeleccionado(OpcionSeleccionada As Byte)
    cad = CStr(OpcionSeleccionada)
End Sub

Private Sub Image1_Click()
    If Text1(1).Enabled = True Then
        Me.cd1.Filter = TextoParaComonDialog2(True)
        Me.cd1.ShowOpen
        If Me.cd1.FileName <> "" Then Text1(1).Text = Me.cd1.FileName
    End If
End Sub

Private Sub ListView1_DblClick()
    Command1_Click 0
End Sub

Private Sub PonerFrames(Habilitar As Boolean)
    Me.FramePlant.Visible = Habilitar
    Frame2.Enabled = Not Habilitar
    
End Sub


Private Function SubirPlantilla() As Boolean
Dim I As Integer
Dim Fin As Boolean
Dim Ext As Integer
Dim OK As Boolean
Dim Fecha As String

    SubirPlantilla = False
    Fecha = ""
    
    'Voy a comprobar la extension
    If Text1(1).Text = "" Then
        OK = True
    
    Else
            I = InStrRev(Text1(1).Text, ".")
            If I = 0 Then
                MsgBox "Error en la extension", vbExclamation
                Exit Function
            End If
            cad = Mid(Text1(1).Text, I + 1)
            cad = DevuelveDesdeBD("codext", "extension", "exten", cad, "T")
            If cad = "" Then
                MsgBox "Extension incorrecta.", vbExclamation
                Exit Function
            End If
            Ext = Val(cad)
            
            cad = "Select codigo from plantilla"
            Set miRSAux = New ADODB.Recordset
            miRSAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            I = 1
            Fin = False
            Do
                If Not miRSAux.EOF Then
                    If miRSAux!codigo = I Then
                        I = I + 1
                        miRSAux.MoveNext
                    Else
                        Fin = True
                    End If
                Else
                    Fin = True
                End If
            
            Loop Until Fin
                miRSAux.Close
                Set miRSAux = Nothing
                DatosCopiados = "NO"
                frmMovimientoArchivo.Opcion = 14   'Llevar plantilla
                frmMovimientoArchivo.Origen = Text1(1).Text
                frmMovimientoArchivo.Destino = I
                frmMovimientoArchivo.Show vbModal
                OK = True
                Fecha = "'" & Format(Now, FormatoFecha) & "'"
    End If
    
    If OK Then
        If Me.FrameTapa.Visible Then
            If DatosCopiados = "" Then
                'INSERTAMOS EN BD
                
                cad = "INSERT INTO plantilla (codigo, Descripcion, tipo, lectura, fecha,carpeta) VALUES ("
                cad = cad & I & ",'" & DevNombreSql(Text1(0).Text) & "'," & Ext & ",0,'" & Format(Now, FormatoFecha) & "',"
                If TieneCarpetas Then
                    cad = cad & ListView2.SelectedItem.Tag & ")"
                Else
                    cad = cad & "0)"
                End If
                
                Conn.Execute cad
        
            End If
        Else
            
            'MODIFICAR
            cad = "UPDATE Plantilla set descripcion='" & DevNombreSql(Text1(0).Text) & "'"
            If Fecha <> "" Then cad = cad & ",Fecha =" & Fecha
            cad = cad & " WHERE codigo =" & ListView1.SelectedItem.Tag
            Conn.Execute cad
        End If
        SubirPlantilla = True
    End If
    

    
    
End Function


Private Function EliminarPlantilla() As Boolean

   cad = "DELETE FROM plantilla where codigo =" & ListView1.SelectedItem.Tag
   Conn.Execute cad
    CargaList
End Function

Private Sub ListView2_DblClick()
    If ListView2.SelectedItem Is Nothing Then Exit Sub
    CargaList
    Label2.Caption = ListView2.SelectedItem.Text
    Me.FrameCarpeta.Visible = False
    
End Sub


Private Sub MoverPlantillas()
Dim I As Integer
    
    Set frmP = New frmPregunta
    frmP.origenDestino = ListView2.SelectedItem.Tag
    frmP.Opcion = 24
    cad = ""
    frmP.Show vbModal
    If cad <> "" Then
        For I = 2 To ListView1.ListItems.Count
            If ListView1.SelectedItem.Selected Then _
                Conn.Execute "UPDATE plantilla set carpeta=" & cad & " WHERE codigo = " & ListView1.SelectedItem.Tag
        Next I
    End If
End Sub
