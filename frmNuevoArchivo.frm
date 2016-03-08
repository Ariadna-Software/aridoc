VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmNuevoArchivo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Archivo"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8805
   Icon            =   "frmNuevoArchivo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   8805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameNuevaPlantilla 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   120
      TabIndex        =   71
      Top             =   0
      Width           =   8535
      Begin VB.Label Label1Plantilla 
         Caption         =   "Nuevo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   72
         Top             =   120
         Width           =   8295
      End
   End
   Begin VB.Frame FrameNuevoEditando 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   120
      TabIndex        =   69
      Top             =   0
      Width           =   8655
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   70
         Top             =   240
         Width           =   8295
      End
   End
   Begin VB.Frame FrCarpeta 
      Height          =   7695
      Left            =   120
      TabIndex        =   56
      Top             =   120
      Width           =   8655
      Begin VB.Frame FrameDatosCarpetas 
         Height          =   5415
         Left            =   120
         TabIndex        =   63
         Top             =   1680
         Width           =   8295
         Begin MSComctlLib.TreeView TreeView1 
            Height          =   4815
            Left            =   120
            TabIndex        =   65
            Top             =   480
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   8493
            _Version        =   393217
            LabelEdit       =   1
            Style           =   7
            Appearance      =   1
         End
         Begin MSComctlLib.ListView ListView2 
            Height          =   4815
            Left            =   4680
            TabIndex        =   64
            Top             =   480
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   8493
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
               Text            =   "Archivo"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label10 
            Caption         =   "Archivos"
            Height          =   195
            Left            =   4680
            TabIndex        =   67
            Top             =   240
            Width           =   1890
         End
         Begin VB.Label Label9 
            Caption         =   "Estructura carpeta"
            Height          =   255
            Left            =   120
            TabIndex        =   66
            Top             =   240
            Width           =   4335
         End
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   120
         TabIndex        =   61
         Text            =   "Text7"
         Top             =   480
         Width           =   8295
      End
      Begin VB.TextBox Text6 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   60
         Top             =   1320
         Width           =   8295
      End
      Begin VB.CommandButton cmdCarpeta 
         Cancel          =   -1  'True
         Caption         =   "Salir"
         Height          =   375
         Index           =   1
         Left            =   7080
         TabIndex        =   58
         Top             =   7200
         Width           =   1215
      End
      Begin VB.CommandButton cmdCarpeta 
         Caption         =   "Siguiente"
         Height          =   375
         Index           =   0
         Left            =   5760
         TabIndex        =   57
         Top             =   7200
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Cargando datos ...."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   600
         TabIndex        =   68
         Top             =   3960
         Visible         =   0   'False
         Width           =   7335
      End
      Begin VB.Label Label8 
         Caption         =   "Carpeta Destino"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   62
         Top             =   240
         Width           =   2130
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   2400
         Picture         =   "frmNuevoArchivo.frx":030A
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label8 
         Caption         =   "Carpeta a insertar en ARIDOC"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   59
         Top             =   1080
         Width           =   2130
      End
   End
   Begin VB.Frame FrameMultiple 
      Height          =   7575
      Left            =   240
      TabIndex        =   16
      Top             =   120
      Width           =   8535
      Begin VB.CommandButton cmdMultiple1 
         Height          =   375
         Index           =   0
         Left            =   3480
         Picture         =   "frmNuevoArchivo.frx":040C
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   200
         Width           =   375
      End
      Begin VB.CommandButton cmdMultiple1 
         Height          =   375
         Index           =   2
         Left            =   3960
         Picture         =   "frmNuevoArchivo.frx":050E
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   200
         Width           =   375
      End
      Begin VB.CommandButton cmdMultiple 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   7200
         TabIndex        =   20
         Top             =   7080
         Width           =   975
      End
      Begin VB.CommandButton cmdMultiple 
         Caption         =   "Siguiente"
         Height          =   375
         Index           =   0
         Left            =   6000
         TabIndex        =   19
         Top             =   7080
         Width           =   975
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   6375
         Left            =   240
         TabIndex        =   18
         Top             =   600
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   11245
         View            =   3
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
            Text            =   "Fichero"
            Object.Width           =   12524
         EndProperty
      End
      Begin VB.Label Label6 
         Caption         =   "Lista de archivos para la insercion por lotes"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.Frame FrameDatos 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   6735
      Left            =   240
      TabIndex        =   23
      Top             =   960
      Width           =   8415
      Begin VB.Frame FrRevision 
         Height          =   1335
         Left            =   5280
         TabIndex        =   73
         Top             =   5400
         Width           =   855
         Begin VB.CommandButton cmdRevision 
            Height          =   735
            Left            =   120
            Picture         =   "frmNuevoArchivo.frx":0610
            Style           =   1  'Graphical
            TabIndex        =   74
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Propietario"
         Height          =   1335
         Left            =   0
         TabIndex        =   51
         Top             =   5400
         Width           =   5175
         Begin VB.TextBox Text5 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Left            =   720
            TabIndex        =   53
            Text            =   "Text2"
            Top             =   360
            Width           =   4335
         End
         Begin VB.TextBox Text4 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   285
            Left            =   720
            TabIndex        =   52
            Text            =   "Text2"
            Top             =   840
            Width           =   4335
         End
         Begin VB.Image imgChangaProp 
            Height          =   240
            Left            =   960
            Picture         =   "frmNuevoArchivo.frx":6472
            Top             =   0
            Width           =   240
         End
         Begin VB.Label Label3 
            Caption         =   "Usuario"
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   55
            Top             =   440
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Grupo"
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   54
            Top             =   880
            Width           =   615
         End
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Left            =   0
         TabIndex        =   46
         Text            =   "Text1"
         Top             =   240
         Width           =   8295
      End
      Begin VB.Frame Frame1 
         Height          =   3735
         Left            =   0
         TabIndex        =   34
         Top             =   720
         Width           =   8295
         Begin VB.TextBox txtClaves 
            Height          =   285
            Index           =   0
            Left            =   240
            MaxLength       =   50
            TabIndex        =   1
            Text            =   "Text3"
            Top             =   480
            Width           =   3975
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   2760
            MaxLength       =   15
            TabIndex        =   50
            Text            =   "Text3"
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox txtClaves 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Index           =   10
            Left            =   6720
            TabIndex        =   10
            Text            =   "Text3"
            Top             =   1920
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox txtClaves 
            Height          =   285
            Index           =   1
            Left            =   4320
            MaxLength       =   50
            TabIndex        =   2
            Text            =   "Text3"
            Top             =   480
            Width           =   3735
         End
         Begin VB.TextBox txtClaves 
            Height          =   285
            Index           =   2
            Left            =   240
            MaxLength       =   50
            TabIndex        =   3
            Text            =   "Text3"
            Top             =   1200
            Width           =   3975
         End
         Begin VB.TextBox txtClaves 
            Height          =   285
            Index           =   3
            Left            =   4320
            MaxLength       =   50
            TabIndex        =   4
            Text            =   "Text3"
            Top             =   1200
            Width           =   3735
         End
         Begin VB.TextBox txtClaves 
            Height          =   285
            Index           =   4
            Left            =   240
            TabIndex        =   5
            Text            =   "99/99/9999"
            Top             =   1920
            Width           =   1095
         End
         Begin VB.TextBox txtClaves 
            Height          =   285
            Index           =   5
            Left            =   1440
            TabIndex        =   6
            Text            =   "Text3"
            Top             =   1920
            Width           =   1095
         End
         Begin VB.TextBox txtClaves 
            Height          =   285
            Index           =   6
            Left            =   2640
            TabIndex        =   7
            Text            =   "Text3"
            Top             =   1920
            Width           =   1095
         End
         Begin VB.TextBox txtClaves 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   7
            Left            =   3840
            TabIndex        =   8
            Text            =   "Text3"
            Top             =   1920
            Width           =   1335
         End
         Begin VB.TextBox txtClaves 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   8
            Left            =   5280
            TabIndex        =   9
            Text            =   "Text3"
            Top             =   1920
            Width           =   1335
         End
         Begin VB.TextBox txtClaves 
            Height          =   885
            Index           =   9
            Left            =   240
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   35
            Text            =   "frmNuevoArchivo.frx":6B34
            Top             =   2640
            Width           =   7815
         End
         Begin VB.Image Image3 
            Height          =   240
            Index           =   6
            Left            =   2640
            Picture         =   "frmNuevoArchivo.frx":6B3A
            Top             =   1680
            Width           =   240
         End
         Begin VB.Image Image3 
            Height          =   240
            Index           =   5
            Left            =   1440
            Picture         =   "frmNuevoArchivo.frx":6C3C
            Top             =   1680
            Width           =   240
         End
         Begin VB.Image Image3 
            Height          =   240
            Index           =   4
            Left            =   240
            Picture         =   "frmNuevoArchivo.frx":6D3E
            Top             =   1680
            Width           =   240
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Añadir final archivo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   840
            TabIndex        =   49
            Top             =   480
            Width           =   1755
         End
         Begin VB.Label Label3 
            Caption         =   "Tamaño"
            Height          =   255
            Index           =   10
            Left            =   6720
            TabIndex        =   48
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Label3"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   45
            Top             =   240
            Width           =   3255
         End
         Begin VB.Label Label3 
            Caption         =   "Label3"
            Height          =   255
            Index           =   1
            Left            =   4320
            TabIndex        =   44
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label3 
            Caption         =   "Label3"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   43
            Top             =   960
            Width           =   3255
         End
         Begin VB.Label Label3 
            Caption         =   "Label3"
            Height          =   255
            Index           =   3
            Left            =   4320
            TabIndex        =   42
            Top             =   960
            Width           =   3255
         End
         Begin VB.Label Label3 
            Caption         =   "Label3"
            Height          =   255
            Index           =   4
            Left            =   480
            TabIndex        =   41
            Top             =   1680
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "Label3"
            Height          =   255
            Index           =   5
            Left            =   1680
            TabIndex        =   40
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Label3"
            Height          =   255
            Index           =   6
            Left            =   2880
            TabIndex        =   39
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Importe"
            Height          =   255
            Index           =   7
            Left            =   3840
            TabIndex        =   38
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Importe"
            Height          =   255
            Index           =   8
            Left            =   5280
            TabIndex        =   37
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Label3"
            Height          =   255
            Index           =   9
            Left            =   240
            TabIndex        =   36
            Top             =   2400
            Width           =   3255
         End
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   4080
         TabIndex        =   29
         Top             =   4560
         Width           =   4215
         Begin VB.OptionButton optEscriutra 
            Caption         =   "Propietario"
            Height          =   195
            Index           =   2
            Left            =   2880
            TabIndex        =   32
            Top             =   360
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton optEscriutra 
            Caption         =   "Grupo"
            Height          =   195
            Index           =   1
            Left            =   1680
            TabIndex        =   31
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton optEscriutra 
            Caption         =   "Todos"
            Height          =   195
            Index           =   0
            Left            =   480
            TabIndex        =   30
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "Escritura"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   33
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   0
         TabIndex        =   24
         Top             =   4560
         Width           =   3975
         Begin VB.OptionButton OptLectura 
            Caption         =   "Propietario"
            Height          =   255
            Index           =   2
            Left            =   2520
            TabIndex        =   27
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton OptLectura 
            Caption         =   "Grupo"
            Height          =   255
            Index           =   1
            Left            =   1440
            TabIndex        =   26
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton OptLectura 
            Caption         =   "Todos"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   25
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label5 
            Caption         =   "Lectura"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   28
            Top             =   0
            Width           =   660
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   0
         Left            =   6960
         TabIndex        =   11
         Top             =   5640
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   6960
         TabIndex        =   12
         Top             =   6240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Carpeta"
         Height          =   255
         Left            =   0
         TabIndex        =   47
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.Frame FrameTapa1 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   240
      TabIndex        =   14
      Top             =   240
      Width           =   8415
      Begin VB.Label Label4 
         Caption         =   "Datos comunes para los archivos de inserción por lotes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   7815
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7680
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   480
      Width           =   8175
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   960
      Picture         =   "frmNuevoArchivo.frx":6E40
      Top             =   240
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Archivo"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmNuevoArchivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmU As frmUsuarios2
Attribute frmU.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1



Public Mc As Ccarpetas
Public Carpeta As String

Public Opcion As Integer
    '0 .- Insercion normal
    '1 .- Insercion multiple
    
    '2 .- Ver propiedades pudiendo modicar
    '3 .- Ver propiedades SIN permiso de modificar
    
    '5 .- Insercion de carpeta entera
    
    
    'NUEVO  25 Abril 2005
    '--------------------------
    
    '
    ' 100 .-
    'Tendremos en una carpeta en el servidor, los archivos vacios.
    'Si me dice nuevo, entonces mostrare la pantalla para que rellene los
    'datos, y si acepta bajare el archivo, insertare en BD, y cuando vuelva a
    'la pantalla ppal, entonces abrire el archivo en modo exclusivo
    '    La opcion sera 100 + el codigo de extension
       
    
    '----------------------------------------
    '  31 Mayo 2005
    '200 .- Nuevo desde plantilla

    
Public mImag As cTimagen

Private Nuevo As Boolean
Private Ndo As Node
Private FSS
Private FinRecursivo As Boolean



Private Sub cmdCarpeta_Click(Index As Integer)
    If Index = 1 Then
        Unload Me
        Exit Sub
    End If
    
    If TreeView1.Nodes.Count = 0 Then
        MsgBox "Seleccione una carpeta", vbExclamation
        Exit Sub
    End If
    
    
    Me.FrCarpeta.Visible = False
    
End Sub

Private Sub cmdMultiple_Click(Index As Integer)
    If Index = 1 Then
        'Salir
        Unload Me
        Exit Sub
    End If
    
    If ListView1.ListItems.Count = 0 Then
        MsgBox "Seleccione algun archivo", vbExclamation
        Exit Sub
    End If
    
    Me.txtClaves(0).Visible = False
    FrameMultiple.Visible = False
    
    
End Sub

Private Sub cmdMultiple1_Click(Index As Integer)
    If Index = 2 Then
        'Quitar uno
        If ListView1.SelectedItem Is Nothing Then Exit Sub
        If MsgBox("Quitar el archivo: " & ListView1.SelectedItem.Text & "?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        ListView1.ListItems.Remove ListView1.SelectedItem.Index
        
    Else
    
        'Añadir
        InsertarArchivos
        
    End If
    Me.Refresh
End Sub

Private Sub cmdRevision_Click()
    Screen.MousePointer = vbHourglass
    Set listacod = New Collection
    listacod.Add mImag.codigo
    frmVarios.Opcion = 7
    frmVarios.Show vbModal
    Set listacod = Nothing
End Sub

Private Sub Command1_Click(Index As Integer)
    If Index = 0 Then
        If Opcion = 1 Then
        
            'Multiple
            If Not HacerMultiple(True, Mc) Then
                Exit Sub
            Else
                Unload Me
            End If
        
        Else
        
        
            If Opcion = 5 Then
        
                If MsgBox("El proceso puede durar muchos minutos. ¿Desea continuar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
                HacerInsercionCarpetas
                Unload Me
                
                
        
        
            Else
        
                    If Not DatosOk(Mc) Then Exit Sub
                        'INSERTAMOS Y TAL Y TAL
                        '
                    If Nuevo Then
                    
                        Insertarimagen Mc
                    Else
                        If mImag.Modificar = 0 Then DatosMOdificados = True
                    End If
                    
            End If
        End If
    End If
    Unload Me
        
End Sub

'para cuando la carpeta no sea la MC, es decir, para cuando hacemos la insercion
' de archivos desde el punto de munu LOTES->Carpetas / subacarpetas
Private Sub Insertarimagen(ByRef vCar As Ccarpetas)
Dim F
    
            If mImag.Agregar(objRevision.LlevaHcoRevision, True) = 0 Then
                If Opcion >= 100 Then
                    'SIMULAREMOS TRAER EL ARCHIVO
                    DatosCopiados = mImag.codigo
                    DatosMOdificados = True
                Else
                    DatosCopiados = "NO"
                    'Llevamos el fichero
                    Set FSS = CreateObject("Scripting.FileSystemObject")
                    Set F = FSS.GetFile(Text1.Text)
                    
                    
                    Set frmMovimientoArchivo.vDestino = vCar
                    frmMovimientoArchivo.Opcion = 1
                    frmMovimientoArchivo.Origen = F.shortpath
                    frmMovimientoArchivo.Destino = CStr(mImag.codigo)
                    frmMovimientoArchivo.Show vbModal
                    
                    'Y si se producen errores entonces borramos el Imgag
                    If DatosCopiados <> "" Then
                        'Error llevando datos
                        mImag.Eliminar
                    Else
                        DatosMOdificados = True
                    End If
                    Set FSS = Nothing
                    Set F = Nothing
                End If
            End If
    
End Sub



Private Sub Form_Load()
Dim I As Integer
    Limpiar Me
    FrCarpeta.Visible = False
    FrameNuevoEditando.Visible = False
    frameNuevaPlantilla.Visible = False
    FrRevision.Visible = False
    Select Case Opcion
    Case 1
    
        Nuevo = True
        Set Me.ListView1.SmallIcons = Admin.ImageList2
    
    
    Case 2, 3
        Nuevo = False
        Me.Command1(0).Visible = Opcion = 2
        PonerValoresImagen
        
        
        
        If Opcion = 3 Then
            For I = 0 To 10
                txtClaves(I).Enabled = False
            Next I
        End If
        If ModoTrabajo = vbNorm And vUsu.Nivel < 3 And objRevision.LlevaHcoRevision Then FrRevision.Visible = True
        PonerCancel
    Case 0, 5
    
        PonerCancel
        txtClaves(4).Text = Format(Now, "dd/mm/yyyy")
        txtClaves(5).Text = Format(Now, "dd/mm/yyyy")
        txtClaves(6).Text = Format(Now, "dd/mm/yyyy")
        
        Nuevo = True
        If Opcion = 5 Then
            Me.txtClaves(0).Visible = False
            Set TreeView1.ImageList = Admin.ImgUsersPCs
            If Admin.ImageList2.ListImages.Count > 0 Then
                Set ListView2.SmallIcons = Admin.ImageList2
                Set Me.ListView1.SmallIcons = Admin.ImageList2
            End If
            Me.FrCarpeta.Visible = True
            FrameDatosCarpetas.Visible = False
            Text7.Text = Carpeta
        End If
        
        
    Case 100 To 150
        
        txtClaves(4).Text = Format(Now, "dd/mm/yyyy")
        txtClaves(5).Text = Format(Now, "dd/mm/yyyy")
        txtClaves(6).Text = Format(Now, "dd/mm/yyyy")
        FrameNuevoEditando.Visible = True
        Nuevo = True
        CarrgaComboInsertables Opcion
        PonerCancel
    Case Is > 199
    
        frameNuevaPlantilla.Visible = True
        txtClaves(4).Text = Format(Now, "dd/mm/yyyy")
        txtClaves(5).Text = Format(Now, "dd/mm/yyyy")
        txtClaves(6).Text = Format(Now, "dd/mm/yyyy")
        Nuevo = True
        PonerDatosPlantilla
    End Select
    
    PonerPropietarios
    
    Me.imgChangaProp.Visible = False
    If Not Nuevo Then
        If vUsu.codusu = 0 Then Me.imgChangaProp.Visible = True
    End If
    
    
    
    Text1.Enabled = Opcion = 0
    If Opcion > 1 And Opcion < 5 Then
        FrameDatos.Top = 90
        'Me.Width = 12
    Else
        FrameDatos.Top = 960
        'Me.Width = 8830
    End If
    Me.Height = FrameDatos.Height + FrameDatos.Top + 550
        

    Me.FrameMultiple.Visible = Opcion = 1
    Me.FrameTapa1.Visible = Opcion = 1 Or Opcion = 5
    Me.Label3(10).Visible = Opcion > 1 And Opcion < 5
    Me.txtClaves(10).Visible = Opcion > 1 And Opcion < 5
    
    Text2.Text = Carpeta
    PonerLabels
    Screen.MousePointer = vbDefault
End Sub

Private Sub PonerCancel()
    On Error Resume Next
    Me.Command1(1).Cancel = True
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If Opcion >= 100 Then GuardarLeerReferenciaCombo False
End Sub

Private Sub frmC_Selec(vFecha As Date)
    DatosCopiados = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub frmU_DatoSeleccionado(CadenaSeleccion As String)
Dim C As String
    Screen.MousePointer = vbHourglass
    C = "Select codgrupo from usuariosgrupos where codusu=" & RecuperaValor(CadenaSeleccion, 1) & " ORDER BY orden"
    Set miRSAux = New ADODB.Recordset
    miRSAux.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    C = ""
    If Not miRSAux.EOF Then
        If Not IsNull(miRSAux.Fields(0)) Then C = miRSAux.Fields(0)
    End If
    miRSAux.Close
    Set miRSAux = Nothing
    If C = "" Then
        MsgBox "Grupo PPal para el usuario: " & CadenaSeleccion & " NO encontrado", vbExclamation
        Exit Sub
    End If
    
    'Llegado aqui, ponemos
    
    mImag.userprop = Val(RecuperaValor(CadenaSeleccion, 1))
    mImag.groupprop = Val(C)
    PonerPropietarios
    Screen.MousePointer = vbDefault
End Sub

Private Sub Image1_Click()
Dim N As Integer
    Me.CommonDialog1.Filter = TextoParaComonDialog2(False, N)
    N = N \ 2
    If N > 0 Then Me.CommonDialog1.FilterIndex = N
    Me.CommonDialog1.ShowOpen
    If Me.CommonDialog1.FileName <> "" Then Text1.Text = Me.CommonDialog1.FileName
    
End Sub


Private Sub Image2_Click()
    Image2.Tag = GetFolder("Carpeta para importar")
    If Image2.Tag <> "" Then
        If Image2.Tag <> Text6.Text Then
            Screen.MousePointer = vbHourglass
            Me.FrameDatosCarpetas.Visible = False
            Label11.Visible = True
            
            Me.Refresh
        
            'Fale, ha cambiado
            Text6.Text = Image2.Tag
            CargarDatosCarpetas
            
            Me.FrameDatosCarpetas.Visible = True
            Label11.Visible = False
            Me.Refresh
            
        End If
    End If
    
End Sub

Private Sub Image3_Click(Index As Integer)
    Set frmC = New frmCal
    frmC.Fecha = Now
    If txtClaves(Index).Text <> "" Then
        If IsDate(txtClaves(Index).Text) Then frmC.Fecha = CDate(txtClaves(Index).Text)
    End If
    DatosCopiados = ""
    frmC.Show vbModal
    Set frmC = Nothing
    If DatosCopiados <> "" Then
        txtClaves(Index).Text = DatosCopiados
        DatosCopiados = ""
    End If
End Sub

Private Sub imgChangaProp_Click()
    'Para el root
    'Cambiara usuario
    If ModoTrabajo <> vbNorm Then Exit Sub
    Set frmU = New frmUsuarios2
    frmU.DatosADevolverBusqueda = "0|"
    frmU.Show vbModal
    Set frmU = Nothing
End Sub

Private Sub Text1_GotFocus()
    With Text1
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub



Private Sub Text1_LostFocus()
    Text1.Text = Trim(Text1.Text)
    If Text1.Text <> "" Then
        If Dir(Text1.Text, vbArchive) = "" Then
            MsgBox "Imposible encontrar: " & Text1.Text, vbExclamation
            Text1.Text = ""
            Text1.SetFocus
        End If
    End If
End Sub









Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
    
    Set miRSAux = New ADODB.Recordset
    miRSAux.Open "Select codext,exten from extension", Conn, adOpenKeyset, adLockOptimistic, adCmdText
    
    CargArchivosCarpetas Node.Key, True
    
    miRSAux.Close
    Set miRSAux = Nothing
End Sub


Private Sub CargArchivosCarpetas(ByRef CADENA As String, Mostrando As Boolean)
Dim F, f1, fc
Dim Exten As String
Dim I As Integer
Dim ICO As Integer
Dim ItmX As ListItem
Dim J As Integer


    Screen.MousePointer = vbHourglass
    If Mostrando Then
        ListView2.ListItems.Clear
    Else
        ListView1.ListItems.Clear
    End If
    Set FSS = CreateObject("Scripting.FileSystemObject")
    Set F = FSS.GetFolder(CADENA)
    
    J = 1
    Set fc = F.Files
    For Each f1 In fc
        I = InStrRev(f1.Name, ".")
        ICO = 1
        If I > 0 Then
            Exten = Mid(f1.Name, I + 1)
            miRSAux.Find " exten = '" & Exten & "'", , , 1
            If Not miRSAux.EOF Then ICO = miRSAux!codext + 1
        End If
    
        If Mostrando Then
            Set ItmX = ListView2.ListItems.Add(, f1.Name)
            ItmX.Text = f1.Name
            ItmX.SmallIcon = ICO
            If ICO = 1 Then ItmX.Bold = True
        
        Else
            'Cargamos en LISTVIEW1, solo si ICO>1
            If ICO > 1 Then
                Set ItmX = ListView1.ListItems.Add(, "c" & J)
                ItmX.Text = f1.Path
                ItmX.SmallIcon = ICO
                J = J + 1

            End If
        End If
    Next
    Screen.MousePointer = vbDefault
End Sub




Private Sub txtClaves_GotFocus(Index As Integer)
    With txtClaves(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtClaves_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index <> 9 Then KEYpress KeyAscii
End Sub

Private Sub PonerLabels()
    'En funcion de la configuracion pondremos los labels
    'el texto k keramos
    Me.Label3(0).Caption = vConfig.C1
    Me.Label3(1).Caption = vConfig.C2
    Me.Label3(2).Caption = vConfig.c3
    Me.Label3(3).Caption = vConfig.c4
    Me.Label3(4).Caption = vConfig.f1
    Me.Label3(5).Caption = vConfig.f2
    Me.Label3(6).Caption = vConfig.f3
    Me.Label3(7).Caption = vConfig.imp1
    Me.Label3(8).Caption = vConfig.imp2
    Me.Label3(9).Caption = vConfig.obs
  
End Sub

'Falta cuando es multiple de carpetas, la carpeta va cambiando
'luego en imag.cdocarpeta no es el mismo codigo siempre
Private Function DatosOk(ByRef ObjCarpeta As Ccarpetas) As Boolean
Dim I As Long
Dim Aux As String
Dim mEx As Cextensionpc



On Error GoTo EDatosOK
    DatosOk = False

    If Opcion <> 2 And Opcion < 100 Then
    
        If Text1.Text = "" Then
            MsgBox "Error en fichero", vbExclamation
            Text1.SetFocus
            Exit Function
        End If
    
        If Dir(Text1.Text, vbArchive) = "" Then
            MsgBox "Error en el archivo. No existe."
            Text1.SetFocus
            Exit Function
        End If
     End If
        
   
        If Me.txtClaves(0).Text = "" Then
            MsgBox "Campos " & Label3(0).Caption & " es obligatorio", vbExclamation
            Exit Function
        End If

    
    If txtClaves(4).Text = "" Then
        MsgBox "Campos " & Label3(4).Caption & " es obligatorio", vbExclamation
        Exit Function
    End If
    
    
    If Opcion > 100 Then
        If Combo1.ListIndex < 0 Then
            MsgBox "Seleccione el tipo de documento", vbExclamation
            Exit Function
        End If
    End If
    
    
     If Opcion <> 2 And Opcion < 100 Then

        I = InStrRev(Text1.Text, ".")
        If I = 0 Then
            MsgBox "Extension incorrecta", vbExclamation
            Exit Function
        End If
  
        'La extension
        Carpeta = LCase(Mid(Text1.Text, I + 1))
    
        Aux = DevuelveDesdeBD("codext", "extension", "exten", Carpeta, "T")
        If Aux = "" Then
            MsgBox "Extension no reconocida por el sistema", vbExclamation
            Exit Function
        End If
        I = Val(Aux)
    
        Set mEx = New Cextensionpc
        If mEx.Leer(CInt(I), vUsu.PC) = 1 Then
            MsgBox "Extension para el PC erronea. Mal configurado", vbExclamation
            Set mEx = Nothing
            Exit Function
        End If
    
    

        'Tamaño del archivo
        txtClaves(10).Text = Round(FileLen(Text1.Text) / 1024, 3)

   End If
    
    
    
    'Creamos una nuevo objeto imagen
    If Nuevo Or Opcion = 1 Then Set mImag = New cTimagen
    'Textos
    mImag.campo1 = txtClaves(0).Text
    mImag.campo2 = txtClaves(1).Text
    mImag.campo3 = txtClaves(2).Text
    mImag.campo4 = txtClaves(3).Text
    'Fechas
    mImag.fecha1 = txtClaves(4).Text
    mImag.fecha2 = LetDB(txtClaves(5).Text, "F")
    mImag.fecha3 = LetDB(txtClaves(6).Text, "F")
    
    'Importes
    mImag.importe1 = LetDB(txtClaves(7).Text, "N")
    mImag.importe2 = LetDB(txtClaves(8).Text, "N")
    
    'Observaciones
    mImag.observa = txtClaves(9).Text
    
    'tamanño
    If Opcion < 100 Then
        mImag.tamnyo = CCur(txtClaves(10).Text)
    Else
        mImag.tamnyo = 0
    End If
    'permisos carpetas y demas
    If Opcion <> 2 Then
        mImag.codcarpeta = ObjCarpeta.codcarpeta
        If Opcion < 100 Then
            mImag.codext = mEx.codext  'La extension que nos dara el archvio
        Else
            'mImag.codext = opcion - 100
            mImag.codext = Combo1.ItemData(Combo1.ListIndex)
        End If
        
        mImag.groupprop = vUsu.GrupoPpal
        mImag.userprop = vUsu.codusu
    End If
    Set mEx = Nothing
    
    
    ' -----------
    'PERMISOS
    'lectura
    If Me.OptLectura(0).Value Then
        I = vbPermisoTotal
    Else
        If Me.OptLectura(1).Value Then
            I = GrupoLongBD(vUsu.GrupoPpal)
        Else
            I = 0
        End If
    End If
    mImag.lecturag = I
    
    
    'escritura
     
    If Me.optEscriutra(0).Value Then
       I = vbPermisoTotal
    Else
        If Me.optEscriutra(1).Value Then
            I = GrupoLongBD(vUsu.GrupoPpal)
        Else
            I = 0
        End If
    End If
    mImag.escriturag = I
    
    DatosOk = True
    Exit Function
EDatosOK:
    MuestraError Err.Number, "DatosOK"
End Function

Private Sub txtClaves_LostFocus(Index As Integer)
     'Importes, fechas y demas
    txtClaves(Index).Text = Trim(txtClaves(Index).Text)
    If txtClaves(Index).Text = "" Then Exit Sub
    
    
    Select Case Index
    Case 4, 5, 6
        'FECHA
        If Not EsFechaOK(txtClaves(Index)) Then txtClaves(Index).SetFocus
    Case 7, 8
        'Importes
        If Not IsNumeric(txtClaves(Index).Text) Then
            MsgBox "Campo numérico", vbExclamation
            txtClaves(Index).SetFocus
        End If
    End Select
End Sub


Private Sub PonerValoresImagen()
Dim I As Integer
On Error GoTo Eponervaloresimagen
    
    With mImag
        Caption = "Archivo (" & .codigo & ")"
        txtClaves(0).Text = .campo1
        txtClaves(1).Text = .campo2
        txtClaves(2).Text = .campo3
        txtClaves(3).Text = .campo4
    'Fechas
        txtClaves(4).Text = .fecha1
        If Val(.fecha2) <> 0 Then txtClaves(5).Text = .fecha2
        If Val(.fecha3) <> 0 Then txtClaves(6).Text = .fecha3
    
    'Importes
        If .importe1 <> 0 Then txtClaves(7).Text = TransformaComasPuntos(.importe1)
        If .importe2 <> 0 Then txtClaves(8).Text = TransformaComasPuntos(.importe2)
    
    'Observaciones
        txtClaves(9).Text = .observa
        txtClaves(10).Text = .tamnyo
        
        If .lecturag = vbPermisoTotal Then
            OptLectura(0).Value = True
        Else
            If .lecturag = 0 Then
                OptLectura(2).Value = True
            Else
                OptLectura(1).Value = True
            End If
        End If
         
        If .escriturag = vbPermisoTotal Then
            optEscriutra(0).Value = True
        Else
            If .escriturag = 0 Then
                optEscriutra(2).Value = True
            Else
                optEscriutra(1).Value = True
            End If
        End If
        
    End With
        
    'Solo el propietario, y los que pertencientes al grupo puden cambiar el
    'acceso
    I = 0
    If mImag.userprop = vUsu.codusu Or mImag.groupprop = vUsu.GrupoPpal Or vUsu.codusu = 0 Then I = 1
    Frame3.Enabled = I = 1
    Frame2.Enabled = I = 1
            
    Exit Sub
Eponervaloresimagen:
    MuestraError Err.Number, "poner valores Timagen"
End Sub


Private Sub InsertarArchivos()


Dim Texto, Extension
Dim I, Fin
Dim J As Integer
Dim miCarpeta, miArchivo, CADENA
Dim Aux, Rc
Dim ItmX As ListItem

'Primero mensaje sobre la carpeta donde van

  CommonDialog1.CancelError = True
  On Error GoTo ErrHandler
  CommonDialog1.Flags = cdlOFNExplorer + cdlOFNAllowMultiselect + cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNFileMustExist
  CommonDialog1.DialogTitle = "Importar Varios Archivos"
  Texto = TextoParaComonDialog2(False, J)
  If Texto = "" Then
      Extension = "No hay ningun tipo de archivos para importar." & vbCrLf & vbCrLf
      Extension = Extension & "Si desea añadir tipos de archivos para  importarlos" & vbCrLf
      Extension = Extension & " hágalo desde el formulario de configuración":
      MsgBox Extension, vbInformation
      Exit Sub
   End If
  
  CommonDialog1.FileName = ""
  CommonDialog1.Filter = Texto
  J = J \ 2
  If J = 0 Then J = 1
  CommonDialog1.FilterIndex = J
  CommonDialog1.InitDir = "c:\" 'App.path
  CommonDialog1.MaxFileSize = 1024 * 30
  CommonDialog1.ShowOpen
'******* Cambiamos cursor
  Screen.MousePointer = vbHourglass
  CommonDialog1.MaxFileSize = 256
  J = InStr(1, CommonDialog1.FileName, Chr(0))
  CADENA = CommonDialog1.FileName
    If J = 0 Then
        'Solo hay un archivo es decir c:\..\eje.txt
        miCarpeta = Mid(CADENA, 1, InStr(4, CADENA, CommonDialog1.FileTitle) - 1)
        miArchivo = CommonDialog1.FileTitle
        CADENA = ""
    Else
        miCarpeta = Mid(CADENA, 1, J - 1)
        If InStr(Len(miCarpeta) - 1, miCarpeta, "\") <> Len(miCarpeta) Then
            miCarpeta = miCarpeta & "\"
            End If
        Aux = InStr(J + 1, CommonDialog1.FileName, Chr(0))
        miArchivo = Mid(CommonDialog1.FileName, J + 1, Aux - 1 - J)
        CADENA = Mid(CommonDialog1.FileName, Aux + 1, Len(CommonDialog1.FileName))
    End If
  Fin = False

  J = ListView1.ListItems.Count + 1
  Do
    Extension = Mid(miArchivo, Len(miArchivo) - 2, Len(miArchivo))
    
    Texto = DevuelveDesdeBD("codext", "extension", "exten", CStr(Extension), "T")
    
    If Texto <> "" Then
        Extension = miCarpeta & miArchivo
             
             
        Set ItmX = ListView1.ListItems.Add(, "c" & J)
        ItmX.Text = Extension
        ItmX.SmallIcon = Val(Texto) + 1
        J = J + 1
        
        'Calculamos el nuevo valor para miArchivo
        If CADENA = "" Then
            Fin = True
        Else
  
            
            Aux = InStr(1, CADENA, Chr(0))
            If Aux <> 0 Then
                miArchivo = Mid(CADENA, 1, Aux - 1)
                CADENA = Mid(CADENA, Aux + 1, Len(CADENA))
            Else
                miArchivo = CADENA
                CADENA = ""
            End If
        End If ' de cadena=""
       
    End If
Loop While Not Fin
'Set Fss = Nothing

Screen.MousePointer = vbDefault

'
Exit Sub
ErrHandler:
 'Han pulsado cancelar en el dialogo
Screen.MousePointer = vbDefault
CommonDialog1.MaxFileSize = 256
If Err.Number = 32755 Then _
       Exit Sub
MsgBox "Se ha producido un error." & vbCrLf & _
    "Error número :      " & Err.Number & vbCrLf & _
    "Descripción  :      " & Err.Description, vbExclamation

End Sub


Private Function HacerMultiple(HacerPregunta As Boolean, ByRef EnQueCarpeta As Ccarpetas) As Boolean
Dim I As Integer
Dim Errores As Boolean
Dim OK As Boolean
Dim J As Integer
Dim C As String
Dim T1 As Single

    On Error GoTo eHacerMultiple
    
    HacerMultiple = False
    If HacerPregunta Then
        If MsgBox("El proceso puede llevar varios minutos. ¿Desea continuar?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
    End If
    
    If txtClaves(4).Text = "" Then
        MsgBox "Campo " & Label3(4).Caption & " es obligado", vbExclamation
        Exit Function
    End If
    
    
    For I = ListView1.ListItems.Count To 1 Step -1
        Label4.Caption = ListView1.ListItems(I).Text
        Label4.Refresh
        Text1.Text = ListView1.ListItems(I).Text
        C = ListView1.ListItems(I).Text
        'quitaos el punto
        J = InStr(1, C, ".")
        C = Mid(C, 1, J - 1)
        'Quitamos la \
        J = InStrRev(C, "\")
        C = Mid(C, J + 1)
        'Ponemos , si ha puesto algo
        C = C & " " & Text3.Text
        C = Mid(C, 1, 40)
            
        txtClaves(0).Text = C
            
            
        OK = False
        If DatosOk(EnQueCarpeta) Then
            DatosMOdificados = False
            'Nueva variable
            T1 = Timer
            Insertarimagen EnQueCarpeta
            If DatosMOdificados Then OK = True
            Set mImag = Nothing
            T1 = Timer - T1
            T1 = 1.5 - T1
            Me.Refresh
            If T1 > 0 Then espera T1
        End If
    
        If Not OK Then
            Errores = True
        Else
            ListView1.ListItems.Remove ListView1.ListItems(I).Index
        End If
    Next I
    Label4.Caption = ""
    If Not Errores Then
        If HacerPregunta Then MsgBox "Proceso finalizado", vbInformation
        HacerMultiple = True
    Else
        Me.cmdMultiple(0).Visible = False
        Me.FrameMultiple.Visible = True
        Me.Refresh
    End If
    
    
    
    Exit Function
eHacerMultiple:
    MuestraError Err.Number, "Hacer insercion multiple"
End Function


Private Sub PonerPropietarios()
Dim Cad As String

    On Error GoTo EPonerPropietarios
    Set miRSAux = New ADODB.Recordset
    'Usuario
    If Nuevo Then
        Cad = vUsu.codusu
    Else
        Cad = mImag.userprop
    End If
    Cad = "Select nombre from usuarios WHERE codusu =" & Cad
    miRSAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRSAux.EOF Then Text5.Text = DBLet(miRSAux!Nombre, "T")
    miRSAux.Close
    
    'Grupo
    If Nuevo Then
        Cad = vUsu.GrupoPpal
    Else
        Cad = mImag.groupprop
    End If
    Cad = "Select nomgrupo from grupos WHERE codgrupo =" & Cad
    miRSAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRSAux.EOF Then Text4.Text = DBLet(miRSAux!nomgrupo, "T")
    miRSAux.Close
    
    
    
    
    
    
    Set miRSAux = Nothing
    Exit Sub
EPonerPropietarios:
    MuestraError Err.Number, "Poner datos propietario"
End Sub



Private Sub CargarDatosCarpetas()
Dim FS, F


'Cargo el treeview
    On Error GoTo ECargandoDatos
    
    TreeView1.Nodes.Clear
    ListView1.ListItems.Clear
    Set FSS = CreateObject("Scripting.FileSystemObject")
    Set F = FSS.GetFolder(Text6.Text)
    Set Ndo = TreeView1.Nodes.Add(, , Text6.Text)
    Ndo.Image = "cerrado"
    Ndo.ExpandedImage = "abierto"
    Ndo.Text = F.Name
    
    CargaRamaDelArbol Ndo
    
    Me.FrameDatosCarpetas.Visible = True
    
    
    Set TreeView1.SelectedItem = TreeView1.Nodes(1)
    
        
    
    'Cargamos archivos
    TreeView1_NodeClick TreeView1.SelectedItem
ECargandoDatos:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando Arbol"
    Set F = Nothing
    Set FSS = Nothing
    
End Sub


Private Sub CargaRamaDelArbol(Nod As Node)
   Dim Xnod As Node
   Dim f1, S, sf, F

    Set F = FSS.GetFolder(Nod.Key)
    Set sf = F.SubFolders
    For Each f1 In sf
        Set Xnod = TreeView1.Nodes.Add(Nod.Key, tvwChild)
        Xnod.Key = Nod.Key & "\" & f1.Name
        Xnod.Text = f1.Name
        Xnod.Image = "cerrado"
        Xnod.ExpandedImage = "abierto"
        Xnod.EnsureVisible
        espera 0.05
        CargaRamaDelArbol Xnod
        Set Xnod = Nothing
    Next
    Set sf = Nothing
    Set F = Nothing
End Sub




Private Sub HacerInsercionCarpetas()



    Set miRSAux = New ADODB.Recordset
    miRSAux.Open "Select codext,exten from extension", Conn, adOpenKeyset, adLockOptimistic, adCmdText
        
        
        
    FinRecursivo = False
        
    InsertaArchivosNODO TreeView1.Nodes(1), Mc.codcarpeta
    
    miRSAux.Close
    Set miRSAux = Nothing
End Sub



Private Sub InsertaArchivosNODO(Nod As Node, padre As Integer)
Dim I As Integer
Dim Fin As Boolean
Dim OK As Boolean
Dim MiCa As Ccarpetas
    'Insertamos archivos de nodo
    If FinRecursivo Then
        Fin = True
    Else
        Fin = False
    End If
    
    
    While Not Fin
        Set MiCa = New Ccarpetas
        'Vemos si la carpeta EXISTE o la creo
        I = ValidarCarpetas(Nod.Text, padre, MiCa)
        If I < 0 Then
            MsgBox "Se ha producido un error leyendo Carpeta: " & Nod.Key
           Exit Sub
        End If
                
        'Meto en el tag el valor del codigo de carpeta
        Nod.Tag = I
                
        CargArchivosCarpetas Nod.Key, False
        
        
        If ListView1.ListItems.Count > 0 Then
        
            OK = HacerMultiple(False, MiCa)
            If Not OK Then
                If MsgBox("Se han producido errores. Desea continuar?", vbQuestion + vbYesNo) <> vbYes Then
                    Fin = True
                    FinRecursivo = True
                End If
            End If
        End If
        Set MiCa = Nothing
        If Not Fin Then
            If Nod.Children > 0 Then
                I = Nod.Tag
                InsertaArchivosNODO Nod.Child, I
            End If
           
            If Nod.Text <> Nod.LastSibling Then
                Set Nod = Nod.Next
            Else
                Fin = True
            End If
        End If
    Wend
    Set MiCa = Nothing
    
End Sub




Private Function ValidarCarpetas(NombreCarpeta As String, ElPadre As Integer, ByRef MiC As Ccarpetas) As Integer
Dim Cad As String
Dim RT As ADODB.Recordset


    ValidarCarpetas = -1
    
    Cad = "Select codcarpeta from carpetas where nombre='" & NombreCarpeta & "' AND padre =" & ElPadre
    Set RT = New ADODB.Recordset
    RT.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RT.EOF Then
        'NUEVA
        Set MiC = New Ccarpetas
        MiC.Almacen = Mc.Almacen
        MiC.escriturag = Mc.escriturag
        MiC.lecturag = Mc.lecturag
        MiC.padre = ElPadre
        MiC.userprop = vUsu.codusu
        MiC.groupprop = vUsu.GrupoPpal
        MiC.Nombre = NombreCarpeta
        MiC.pathreal = Mc.pathreal
        MiC.version = Mc.version
        MiC.user = Mc.user
        MiC.pwd = Mc.pwd
        MiC.SRV = Mc.SRV
        If MiC.Agregar = 0 Then ValidarCarpetas = MiC.codcarpeta

    Else
        If MiC.Leer(CInt(RT!codcarpeta), (ModoTrabajo = 1)) = 0 Then ValidarCarpetas = RT!codcarpeta
    End If
    RT.Close
    Set RT = Nothing
    
End Function

Private Sub CarrgaComboInsertables(Valor As Integer)
Dim I As Integer
Dim J As Integer
    Combo1.Clear
    Set miRSAux = New ADODB.Recordset
    miRSAux.Open "Select * from extension where nuevo=1 order by descripcion", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRSAux.EOF
        Combo1.AddItem miRSAux!Descripcion
        Combo1.ItemData(Combo1.NewIndex) = miRSAux!codext
        
        miRSAux.MoveNext
    Wend
    miRSAux.Close
    
    If Combo1.ListCount > 0 Then
        I = -1
        If Opcion > 100 Then
            For J = 0 To Combo1.ListCount - 1
                If Combo1.ItemData(J) = Valor - 100 Then
                    I = J
                    Exit For
                End If
            Next J
        End If
        If I < 0 Then I = GuardarLeerReferenciaCombo(True)
        If I > Combo1.ListCount Then I = 0
        Combo1.ListIndex = I
    End If
    Set miRSAux = Nothing
End Sub



Private Function GuardarLeerReferenciaCombo(Leer As Boolean) As Integer
Dim NF As Integer
    
    On Error GoTo EGuardarLeerReferenciaCombo
    
    GuardarLeerReferenciaCombo = 0
    If Leer Then
        DatosCopiados = App.Path & "\*.cmb"
        DatosCopiados = Dir(DatosCopiados, vbArchive)
        NF = 0
        If DatosCopiados <> "" Then
            NF = InStr(1, DatosCopiados, ".")
            If NF > 0 Then
                DatosCopiados = Mid(DatosCopiados, 1, NF - 1)
                NF = Val(DatosCopiados)
            End If
        End If
        GuardarLeerReferenciaCombo = NF
        Combo1.Tag = NF
    Else
        If Combo1.ListIndex <> Combo1.Tag Then
            If Dir(App.Path & "\*.cmb") <> "" Then Kill App.Path & "\*.cmb"
            If Combo1.ListIndex > 0 Then
               NF = FreeFile
               Open App.Path & "\" & Combo1.ListIndex & ".cmb" For Output As #NF
               Print #NF, Now & Combo1.List(Combo1.ListIndex)
               Close #NF
            End If
        End If
    End If
    
    Exit Function
EGuardarLeerReferenciaCombo:
    Err.Clear
End Function


Private Sub PonerDatosPlantilla()
Dim C As String
Dim I As Integer

    C = "Select * from plantilla where codigo=" & Opcion - 200
    Set miRSAux = New ADODB.Recordset
    miRSAux.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRSAux.EOF Then
        Command1(1).Enabled = False
        MsgBox "Error leyendo datos plantilla: " & Opcion - 200, vbExclamation
        Label1Plantilla.Caption = "ERROR"
        I = -1
    Else
        Label1Plantilla = miRSAux!Descripcion & " (" & Format(miRSAux!Fecha, "dd/mm/yyyy") & ")"
        I = miRSAux!Tipo
    End If
    miRSAux.Close
    Set miRSAux = Nothing
    If I < 0 Then Exit Sub
    CarrgaComboInsertables 100 + I
End Sub



