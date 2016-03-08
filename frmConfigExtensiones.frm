VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmConfigExtensiones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuracion extensiones LOCAL"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9765
   Icon            =   "frmConfigExtensiones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   9765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6480
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Left            =   720
      TabIndex        =   17
      Text            =   "Text2"
      Top             =   120
      Width           =   4455
   End
   Begin VB.Frame FrameDatos 
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   1200
      TabIndex        =   0
      Top             =   1200
      Width           =   7455
      Begin VB.CommandButton cmdModifi 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   6360
         TabIndex        =   21
         Top             =   3360
         Width           =   855
      End
      Begin VB.CommandButton cmdModifi 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   0
         Left            =   5280
         TabIndex        =   20
         Top             =   3360
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   1200
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   2640
         Width           =   6135
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   1200
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   2040
         Width           =   6015
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   1200
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1440
         Width           =   2895
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   3960
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   840
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   3
         Height          =   3975
         Left            =   0
         Top             =   0
         Width           =   7445
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   1
         Left            =   840
         Picture         =   "frmConfigExtensiones.frx":030A
         Top             =   2640
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   0
         Left            =   840
         Picture         =   "frmConfigExtensiones.frx":040C
         Top             =   2040
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "PRINT"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   10
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "EXE"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   8
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Descripcion"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   840
      End
      Begin VB.Label Label2 
         Caption         =   "Extension"
         Height          =   255
         Index           =   1
         Left            =   3120
         TabIndex        =   4
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Codigo"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   615
      End
   End
   Begin VB.Frame FrameTodo 
      Height          =   5535
      Left            =   0
      TabIndex        =   11
      Top             =   480
      Width           =   9615
      Begin VB.CommandButton Command1 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   8160
         TabIndex        =   22
         Top             =   5040
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "GUARDAR"
         Height          =   375
         Left            =   6600
         TabIndex        =   19
         Top             =   5040
         Width           =   1215
      End
      Begin VB.CommandButton cmdExten 
         Height          =   375
         Index           =   2
         Left            =   2880
         Picture         =   "frmConfigExtensiones.frx":050E
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdExten 
         Height          =   375
         Index           =   1
         Left            =   2400
         Picture         =   "frmConfigExtensiones.frx":0A98
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdExten 
         Height          =   375
         Index           =   0
         Left            =   1920
         Picture         =   "frmConfigExtensiones.frx":1022
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3855
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   6800
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cod"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ext."
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Descripcion"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Pathexe"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "impresion"
            Object.Width           =   7056
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Extensiones sistema"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Label Label3 
      Caption         =   "PC"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "frmConfigExtensiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public NuevoEquipo As Boolean

Dim HanPulsadoCerrar As Boolean

Private Sub cmdExten_Click(Index As Integer)
Dim I As Integer
    If Index = 1 Then
        If ListView1.SelectedItem Is Nothing Then Exit Sub
        'Modificar
        With ListView1.SelectedItem
            Text1(0).Text = .Text
            For I = 1 To 4
                Text1(I).Text = .SubItems(I)
            Next I
        End With
        Me.FrameTodo.Enabled = False
        Me.FrameDatos.Visible = True
    End If
End Sub

Private Sub cmdModifi_Click(Index As Integer)

    '---------------------
    If Index = 0 Then
        ListView1.SelectedItem.SubItems(3) = Text1(3).Text
        ListView1.SelectedItem.SubItems(4) = Text1(4).Text
    End If
    
    Me.FrameDatos.Visible = False
    Me.FrameTodo.Enabled = True
End Sub

Private Sub Command1_Click()
    HanPulsadoCerrar = True
    If NuevoEquipo Then
        If MsgBox("No deberia cancelar la configuracion de extensiones." & vbCrLf & _
            "Sin las extensiones el programa no se ejecutará correctamente" & vbCrLf & vbCrLf & _
            "Desea continuar igualmente?", vbQuestion + vbYesNoCancel) = vbYes Then
            'Borro el PC
            Conn.Execute "DELETE FROM equipos where codequipo =" & vUsu.PC
            End
         Else
            MsgBox "Debe reiniciar la aplicación", vbExclamation
         End If
        'Ponemos k ya no hace falta descargar las extensiones
        Conn.Execute "UPDATE equipos SET cargaIconsExt= 1 WHERE codequipo=" & vUsu.PC
        End
        Exit Sub
    End If
    Unload Me
End Sub

Private Sub Command2_Click()
Dim Cad As String
Dim I As Integer

    'Para cerrar
    'Updateamos todos los valores de la tabla
    'Por eso borramos y reinsertamos
    Cad = "DELETE FROM extensionpc WHERE codequipo =" & vUsu.PC
    Conn.Execute Cad
    
    For I = 1 To ListView1.ListItems.Count
        Cad = "INSERT INTO extensionpc (codext, codequipo, pathexe, impresion) VALUES ("
        Cad = Cad & ListView1.ListItems(I).Text & "," & vUsu.PC & ",'" & DevNombreSql(ListView1.ListItems(I).SubItems(3))
        Cad = Cad & "','" & DevNombreSql(ListView1.ListItems(I).SubItems(4)) & "')"
        Conn.Execute Cad
    Next I
    HanPulsadoCerrar = True
    MsgBox "Debe reiniciar la aplicacion", vbExclamation
    End
End Sub

Private Sub Form_Load()
    HanPulsadoCerrar = False
    'Cargamos las extensiones
    FrameDatos.Visible = False
    Text2.Text = vUsu.NomPC
    CargarExtensiones True
End Sub



Private Sub CargarExtensiones(Todo As Boolean)
Dim Rs As ADODB.Recordset
Dim Cad As String
Dim ItmX As ListItem
Dim I As Integer

    
    Cad = "Select * from extension ORDER BY codext"
    Set Rs = New ADODB.Recordset
    If Todo Then
        Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not Rs.EOF
               Set ItmX = ListView1.ListItems.Add
               ItmX.Text = Rs!codext
               ItmX.SubItems(1) = Rs!Exten
               ItmX.SubItems(2) = Rs!Descripcion
               If NuevoEquipo Then
                    ItmX.SubItems(3) = DBLet(Rs!ofertaexe, "T")
                    ItmX.SubItems(4) = DBLet(Rs!ofertaprint, "T")
                End If
               ItmX.Tag = 0 'Indicara si tiene o no
               Rs.MoveNext
        Wend
        Rs.Close
    End If
    
    'Vemos para este PC cuales tiene
    If Not NuevoEquipo Then
    
        Cad = "Select * from extensionpc where codequipo=" & vUsu.PC
        Rs.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not Rs.EOF
            Cad = ""
            For I = 1 To ListView1.ListItems.Count
                If ListView1.ListItems(I).Text = Rs!codext Then
                    Cad = I
                    Exit For
                End If
            Next I
            If Cad <> "" Then
                I = CInt(Cad)
                ListView1.ListItems(I).SubItems(3) = DBLet(Rs!pathexe)
                ListView1.ListItems(I).SubItems(4) = DBLet(Rs!impresion)
                ListView1.ListItems(I).Tag = 1
            End If
            Rs.MoveNext
        Wend
        Rs.Close
    End If
    Set Rs = Nothing
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If Not HanPulsadoCerrar Then Cancel = 1
End Sub

Private Sub Image1_Click(Index As Integer)
    CommonDialog1.Filter = "Archivos ejecutables (*.exe)|*.exe"
    CommonDialog1.CancelError = False
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then
        Text1(3 + Index) = CommonDialog1.FileName
    End If
End Sub
