VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPregunta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Esta es la pregunta"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10560
   Icon            =   "frmPregunta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   10560
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FrameCambioUserprop 
      Height          =   3495
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   5175
      Begin VB.TextBox Text3 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   360
         TabIndex        =   38
         Top             =   2280
         Width           =   4455
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   35
         Top             =   1560
         Width           =   4455
      End
      Begin VB.CommandButton cmdUserProp 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   3720
         TabIndex        =   34
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton cmdUserProp 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   0
         Left            =   2400
         TabIndex        =   33
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Image imgCambiaUserGroup 
         Height          =   240
         Index           =   1
         Left            =   960
         Picture         =   "frmPregunta.frx":030A
         Top             =   2040
         Width           =   240
      End
      Begin VB.Image imgCambiaUserGroup 
         Height          =   240
         Index           =   0
         Left            =   960
         Picture         =   "frmPregunta.frx":0894
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label13 
         Caption         =   "Grupo"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   39
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label13 
         Caption         =   "Usuario"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   37
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "Cambio de propietario "
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
         Index           =   1
         Left            =   960
         TabIndex        =   36
         Top             =   840
         Width           =   3735
      End
      Begin VB.Label Label12 
         Caption         =   "Cambio de propietario "
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
         Index           =   0
         Left            =   240
         TabIndex        =   32
         Top             =   480
         Width           =   4455
      End
   End
   Begin VB.Frame FrameProps 
      Height          =   4335
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   5415
      Begin VB.TextBox Text1 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   735
         Left            =   240
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   1320
         Width           =   4815
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Salir"
         Height          =   375
         Left            =   4200
         TabIndex        =   10
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Subcarpetas"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   21
         Top             =   3120
         Width           =   1590
      End
      Begin VB.Label Label8 
         Caption         =   "Label7"
         Height          =   255
         Index           =   4
         Left            =   2520
         TabIndex        =   20
         Top             =   3120
         Width           =   2655
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   240
         X2              =   5160
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Label Label8 
         Caption         =   "Label7"
         Height          =   255
         Index           =   3
         Left            =   2520
         TabIndex        =   19
         Top             =   2640
         Width           =   2415
      End
      Begin VB.Label Label7 
         Caption         =   "Total tamaño archivos"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   18
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label Label8 
         Caption         =   "Label7"
         Height          =   255
         Index           =   2
         Left            =   2520
         TabIndex        =   17
         Top             =   2280
         Width           =   2415
      End
      Begin VB.Label Label7 
         Caption         =   "Total archivos carpeta"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   16
         Top             =   2280
         Width           =   1590
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   240
         X2              =   5160
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Label Label8 
         Caption         =   "Label7"
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   15
         Top             =   1800
         Width           =   2535
      End
      Begin VB.Label Label7 
         Caption         =   "Tamaño archivos sel."
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   1800
         Width           =   1875
      End
      Begin VB.Label Label8 
         Caption         =   "Label7"
         Height          =   255
         Index           =   0
         Left            =   2520
         TabIndex        =   13
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Nº Selecccionados"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   1440
         Width           =   1365
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   240
         X2              =   5160
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label6 
         Caption         =   "Label6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   4455
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   4800
         Picture         =   "frmPregunta.frx":0E1E
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Frame FrameMoverPlantilla 
      Height          =   4095
      Left            =   0
      TabIndex        =   62
      Top             =   0
      Width           =   4455
      Begin VB.ListBox List1 
         Height          =   2790
         Left            =   240
         TabIndex        =   65
         Top             =   480
         Width           =   3975
      End
      Begin VB.CommandButton cmdSelCarpetaPlantilla 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   3120
         TabIndex        =   64
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton cmdSelCarpetaPlantilla 
         Caption         =   "Seleccionar"
         Height          =   375
         Index           =   0
         Left            =   1680
         TabIndex        =   63
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label Label17 
         Caption         =   "Carpetas plantillas"
         Height          =   255
         Left            =   240
         TabIndex        =   66
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.Frame FramePlantillaCarpeta 
      Height          =   2895
      Left            =   0
      TabIndex        =   56
      Top             =   0
      Width           =   5175
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   360
         TabIndex        =   57
         Text            =   "Text5"
         Top             =   1200
         Width           =   4455
      End
      Begin VB.CommandButton cmdPlantiCar 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   1
         Left            =   2640
         TabIndex        =   58
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton cmdPlantiCar 
         Caption         =   " Cancelar"
         Height          =   375
         Index           =   0
         Left            =   3840
         TabIndex        =   59
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label16 
         Caption         =   "Nombre"
         Height          =   255
         Left            =   360
         TabIndex        =   61
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "NUEVO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   120
         TabIndex        =   60
         Top             =   360
         Width           =   4695
      End
   End
   Begin VB.Frame FrameInteg 
      Height          =   2055
      Left            =   0
      TabIndex        =   51
      Top             =   0
      Width           =   6015
      Begin VB.CommandButton Command4 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3600
         TabIndex        =   54
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   2
         Left            =   4800
         TabIndex        =   53
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   120
         TabIndex        =   52
         Top             =   720
         Width           =   5655
      End
      Begin VB.Image Image3 
         Height          =   240
         Left            =   1560
         Picture         =   "frmPregunta.frx":1560
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label14 
         Caption         =   "PATH integracion"
         Height          =   195
         Left            =   120
         TabIndex        =   55
         Top             =   360
         Width           =   1260
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6240
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameDirMail 
      Height          =   2535
      Left            =   0
      TabIndex        =   44
      Top             =   0
      Width           =   5775
      Begin VB.CommandButton Command3 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   4320
         TabIndex        =   50
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   0
         Left            =   3240
         TabIndex        =   49
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txtmail 
         Height          =   320
         Index           =   1
         Left            =   360
         TabIndex        =   47
         Text            =   "Text4"
         Top             =   1320
         Width           =   4935
      End
      Begin VB.TextBox txtmail 
         Height          =   320
         Index           =   0
         Left            =   360
         TabIndex        =   45
         Text            =   "Text4"
         Top             =   480
         Width           =   4935
      End
      Begin VB.Label lblMail 
         Caption         =   "Direccion e-mail"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   48
         Top             =   1080
         Width           =   1125
      End
      Begin VB.Label lblMail 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   46
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame FrameSelFolder 
      Height          =   6255
      Left            =   0
      TabIndex        =   40
      Top             =   0
      Width           =   5655
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   5415
         Left            =   240
         TabIndex        =   43
         Top             =   240
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   9551
         _Version        =   393217
         LabelEdit       =   1
         Style           =   7
         Appearance      =   1
      End
      Begin VB.CommandButton cmdSelFolder 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   4200
         TabIndex        =   42
         Top             =   5760
         Width           =   1095
      End
      Begin VB.CommandButton cmdSelFolder 
         Caption         =   "Seleccionar"
         Height          =   375
         Index           =   0
         Left            =   3000
         TabIndex        =   41
         Top             =   5760
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      Begin VB.CommandButton Command1 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   6720
         TabIndex        =   8
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   0
         Left            =   5400
         TabIndex        =   7
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Mover"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "frmPregunta.frx":1662
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label5 
         Caption         =   "Desea copiar los archivos seleccionados "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   6
         Top             =   1320
         Width           =   6135
      End
      Begin VB.Label Label4 
         Caption         =   "Desea copiar los archivos seleccionados "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   5
         Top             =   720
         Width           =   6135
      End
      Begin VB.Label Label3 
         Caption         =   "Destino:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   4
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Origen:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   3
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Desea COPIAR los archivos seleccionados ?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   1
         Top             =   240
         Width           =   5535
      End
   End
   Begin VB.Frame frImportes 
      Height          =   4335
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   6135
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   4800
         TabIndex        =   30
         Top             =   3720
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   3120
         TabIndex        =   28
         Text            =   "Text2"
         Top             =   2760
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   26
         Text            =   "Text2"
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "Label11"
         Height          =   255
         Index           =   1
         Left            =   3120
         TabIndex        =   29
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "Label11"
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   27
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Label10"
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
         Left            =   240
         TabIndex        =   25
         Top             =   720
         Width           =   5535
      End
      Begin VB.Label Label9 
         Caption         =   "Total para:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   4455
      End
   End
End
Attribute VB_Name = "frmPregunta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event DatoSeleccionado(OpcionSeleccionada As Byte)

Private WithEvents frmU As frmUsuarios2
Attribute frmU.VB_VarHelpID = -1
Private WithEvents frmG As frmGrupos
Attribute frmG.VB_VarHelpID = -1

Public Opcion As Byte
    '1.- Copiar / Mover Archivos
    '2.-   "  / "  CARPETAS
    
    
    '5.- Propiedades de unos archivos
    '6.- Propiedades carpeta
    
    
    '8.- Importes archivos seleccionados
    '9.- Importes carpeta seleccionada
    '10.- Importe subcarpetas
    
    
    '11.- Cambio de propietario para los archivos
    
    '20.- Seleccionar una carpeta para mover archivos
    
    '21.- Direccion e- mail
    
    '22.- Preguna PATH integrador
    
    '23.- Nueva( o modificar) carpeta para las plantillas
    
    '24.- Seleccionar carpeta para agregar mover las plantillas
    
Public origenDestino As String   'Separados con pipes
Private AntiguoCursor As Byte
Private PrimeraVez As Boolean

Private Sub Check1_Click()
    Select Case Opcion
    Case 1
        If Check1.Value = 0 Then
            Label1.Caption = "Desea COPIAR los archivos seleccionados"
        Else
            Label1.Caption = "Desea MOVER los archivos seleccionados"
        End If
        
    Case 2
        If Check1.Value = 0 Then
            Label1.Caption = "Desea COPIAR la carpeta"
        Else
            Label1.Caption = "Desea MOVER la carpeta"
        End If
    End Select
    
End Sub

Private Sub cmdPlantiCar_Click(Index As Integer)
Dim i As Byte
    If Index = 1 Then
        Text5.Text = Trim(Text5.Text)
        If Text5.Text = "" Then
            MsgBox "Escriba el nombre de la carpeta", vbExclamation
            Exit Sub
        End If
    
        i = InsertarModificarPlantillaCarpeta()
        If i = 0 Then Exit Sub
        RaiseEvent DatoSeleccionado(i)
    
    End If
    Unload Me   'Para index=0 o 1
    
End Sub





Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdSelCarpetaPlantilla_Click(Index As Integer)
Dim i As Integer

    If Index = 0 Then
        For i = 0 To List1.ListCount - 1
            If List1.Selected(i) Then
                If MsgBox("¿Desea mover las plantillas seleccionadas a la carpeta: " & List1.List(i) & "?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
                RaiseEvent DatoSeleccionado(CByte(List1.ItemData(i)))
                Exit For
            End If
        Next i
    End If
    Unload Me
End Sub

Private Sub cmdSelFolder_Click(Index As Integer)
    If Index = 0 Then
        If TreeView1.SelectedItem Is Nothing Then Exit Sub
        
        If origenDestino = "1" Then
            'Es para el traspaso a hco. Ademas de la carpeta voy a llevar todas las subcarpetas colgantes
            origenDestino = CopiaArchivosCarpetaRecursiva(TreeView1.SelectedItem)
            DatosCopiados = TreeView1.SelectedItem.FullPath & "·" & origenDestino
        Else
            '"0"
            DatosCopiados = TreeView1.SelectedItem.Key & "|" & TreeView1.SelectedItem.Text & "|" & TreeView1.SelectedItem.FullPath & "|"
        End If
            
    End If
    Unload Me
End Sub

Private Sub cmdUserProp_Click(Index As Integer)
    If Index = 0 Then
        
        If Val(Text3(0).Tag) > 127 Then
            MsgBox "Error critico. Supera capacidad BYTE", vbCritical
        Else
            
            RaiseEvent DatoSeleccionado(CByte(Text3(0).Tag))
        End If
        
    End If
    Unload Me
End Sub

Private Sub Command1_Click(Index As Integer)
Dim Rc As Byte
    If Index = 0 Then
        If Check1.Value = 0 Then
            Rc = 1  'Copiar
        Else
            Rc = 2  'Mover
        End If
        RaiseEvent DatoSeleccionado(Rc)
    End If
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click(Index As Integer)
    'MAIL
    DatosCopiados = ""
    If Index = 0 Then
        txtmail(0).Text = Trim(txtmail(0).Text)
        txtmail(1).Text = Trim(txtmail(1).Text)
        If txtmail(1).Text = "" Or txtmail(0).Text = "" Then
            MsgBox "Campos obligatorios", vbExclamation
            Exit Sub
        End If
        
        If InStr(1, txtmail(1).Text, "@") = 0 Then
            MsgBox "Direccion e-mail incorrecta", vbExclamation
            Exit Sub
        End If
        
        If InStr(1, txtmail(1).Text, ".") = 0 Then
            MsgBox "Direccion e-mail incorrecta", vbExclamation
            Exit Sub
        End If
        
        
        'Llegados aqui, devolvemos datos
        DatosCopiados = txtmail(0).Text & "|" & txtmail(1).Text & "|"
        
    End If
    Unload Me
End Sub

Private Sub Command4_Click()
    'PATH INTEGRA
    Text4.Text = Trim(Text4.Text)
    
    If Text4.Text = "" Then
        MsgBox "Ponga algun valor", vbExclamation
        Exit Sub
    End If
    
    DatosCopiados = Text4.Text
    
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        If Opcion = 20 Then
            TreeView1.SetFocus
        End If
    End If
End Sub

Private Sub Form_Load()
Dim H As Integer
Dim W As Integer
    PrimeraVez = True
    Check1.Value = 0
    Me.Frame1.Visible = False
    Me.FrameProps.Visible = False
    Me.frImportes.Visible = False
    FrameCambioUserprop.Visible = False
    Me.FrameSelFolder.Visible = False
    Me.FrameDirMail.Visible = False
    Me.FrameInteg.Visible = False
    Me.FramePlantillaCarpeta.Visible = False
    FrameMoverPlantilla.Visible = False
    Select Case Opcion
    Case 1
        Me.Frame1.Visible = True
        Label1.Caption = "Desea COPIAR los archivos seleccionados"
        Label4.Caption = RecuperaValor(origenDestino, 1)
        Label5.Caption = RecuperaValor(origenDestino, 2)
        
    Case 2
        Me.Frame1.Visible = True
        Label1.Caption = "Desea COPIAR la carpeta"
        Label4.Caption = RecuperaValor(origenDestino, 1)
        Label5.Caption = RecuperaValor(origenDestino, 2)
    Case 5, 6
        Me.FrameProps.Visible = True
        PonerLabels
        
    Case 8, 9, 10
        Label11(0).Caption = vConfig.imp1
        Label11(1).Caption = vConfig.imp2
        If Opcion = 8 Then
            Label10.Caption = "Archivos seleccionados en " & RecuperaValor(DatosCopiados, 1) & ": " & RecuperaValor(DatosCopiados, 2)
            Text2(0).Text = RecuperaValor(DatosCopiados, 3)
            Text2(1).Text = RecuperaValor(DatosCopiados, 4)
        Else
            If Opcion = 9 Then
                Label10.Caption = "Carpeta " & RecuperaValor(DatosCopiados, 1)
            Else
                Label10.Caption = "Carpeta y subcarpetas: " & RecuperaValor(DatosCopiados, 1)
            End If
            Text2(0).Text = RecuperaValor(DatosCopiados, 2)
            Text2(1).Text = RecuperaValor(DatosCopiados, 3)
        End If
        frImportes.Visible = True
        
    Case 11
    
        'Cambio usuario prop
        FrameCambioUserprop.Visible = True
        Label12(1).Caption = Me.origenDestino & " archivo(s) seleccionado(s)"
        Text3(0).Tag = 127
        
    Case 20
        'En origen destino tendremos
        'si donde debo devolver la carpeta es para
        'los resultado o traspaso a hco ....
        '   0.- Resultados
        '   1.- Traspaso a hco
        If origenDestino = "" Then origenDestino = 0
        FrameSelFolder.Visible = True
        Me.cmdSelFolder(1).Cancel = True
        CargaElArbolDeAmin
        Caption = "Seleccione una carpeta"
        H = Me.FrameSelFolder.Height
        W = Me.FrameSelFolder.Width
            
    Case 21
        Caption = "Direccion e-mail"
        H = Me.FrameDirMail.Height
        W = Me.FrameDirMail.Width
        Me.FrameDirMail.Visible = True
        txtmail(0).Text = RecuperaValor(DatosCopiados, 1)
        txtmail(1).Text = RecuperaValor(DatosCopiados, 2)
        DatosCopiados = ""
        Me.Command3(1).Cancel = True
    Case 22
        Caption = "Path integ"
        
        H = Me.FrameInteg.Height
        W = Me.FrameInteg.Width
        Me.FrameInteg.Visible = True
        Me.Command3(2).Cancel = True
    
    Case 23
        Caption = "Carpeta para las plantillas"
        
        H = Me.FramePlantillaCarpeta.Height
        W = Me.FramePlantillaCarpeta.Width
        Me.FramePlantillaCarpeta.Visible = True
        Me.cmdPlantiCar(0).Cancel = True
    
    
        If origenDestino = "" Then
            Label15.Caption = "NUEVO"
            Text5.Text = ""
        Else
            ' EN origenDestino vienen nombre carpeta|codigo|
            Label15.Caption = "MODIFICAR"
            Text5.Text = RecuperaValor(origenDestino, 2)
            origenDestino = RecuperaValor(origenDestino, 1)
        End If
        
        
        
    Case 24
        Caption = "Mover plantillas"
        
        H = Me.FrameMoverPlantilla.Height
        W = Me.FrameMoverPlantilla.Width
        Me.FrameMoverPlantilla.Visible = True
        Me.cmdSelCarpetaPlantilla(1).Cancel = True
        CargaPlantillas
    End Select
    
    If Opcion < 3 Then
        Caption = "Pregunta"
        H = Me.Frame1.Height
        W = Me.Frame1.Width
        Me.Command1(1).Cancel = True
    Else
        If Opcion < 8 Then
            Caption = "Informacion"
            H = Me.FrameProps.Height
            W = Me.FrameProps.Width
            Me.Command2.Cancel = True
            Text1.Visible = Opcion = 6
        Else
            If Opcion < 11 Then
                Caption = "Cálculo importes"
                H = Me.frImportes.Height
                W = Me.frImportes.Width
            Else
                If Opcion < 20 Then
                    Caption = "Cambio propietario"
                    H = Me.FrameCambioUserprop.Height
                    W = Me.FrameCambioUserprop.Width
                 End If
            End If
        End If
    End If
    Me.Height = H + 420
    Me.Width = W + 120
    AntiguoCursor = Screen.MousePointer
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Cerrar
    Screen.MousePointer = AntiguoCursor
End Sub

Private Sub PonerLabels()
    Dim C As Long
    
    'vienen empipados:
    ' nombre carpeta
    ' archvios seleccionados, tamañoselecioados
    ' archivos carpetas ,  tamño total,ocultos
    '
    Label6.Caption = RecuperaValor(DatosCopiados, 1)
    Label8(0).Caption = RecuperaValor(DatosCopiados, 2)
    Label8(1).Caption = RecuperaValor(DatosCopiados, 3) & " Kb"
    
    Label8(2).Caption = RecuperaValor(DatosCopiados, 4)
    'tamaño
    Label8(3).Caption = RecuperaValor(DatosCopiados, 5) & " Kb"
    'Coultos
    C = Val(RecuperaValor(DatosCopiados, 6))
    If C > 0 Then Label8(2).Caption = Label8(2).Caption & " - Ocultos " & C
    
    
    Label8(4).Caption = RecuperaValor(DatosCopiados, 7)
    C = Val(RecuperaValor(DatosCopiados, 8))
    If C > 0 Then Label8(4).Caption = Label8(4).Caption & " - Ocultos " & C
    
    
    'Si la opcion es 6
    C = InStrRev(Label6.Caption, "\")
    Text1.Text = ""
    If C > 0 Then
        Text1.Text = Mid(Label6.Caption, 1, C - 1)
        Label6.Caption = Mid(Label6.Caption, C + 1)
    End If
    
End Sub



Private Sub frmG_DatoSeleccionado(CadenaSeleccion As String)
    Text3(1).Text = RecuperaValor(CadenaSeleccion, 2)
    Text3(1).Tag = RecuperaValor(CadenaSeleccion, 1)
End Sub

Private Sub frmU_DatoSeleccionado(CadenaSeleccion As String)
'Dim C As String
'    Screen.MousePointer = vbHourglass
'    C = "Select grupos.codgrupo,grupos.nomgrupo from usuariosgrupos,grupos where "
'    C = C & "usuariosgrupos.codgrupo =grupos.codgrupo and codusu=" & RecuperaValor(CadenaSeleccion, 1)
'    C = C & " ORDER BY orden"
'
'    Set miRSAux = New ADODB.Recordset
'    miRSAux.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    C = ""
'    If Not miRSAux.EOF Then
'        If Not IsNull(miRSAux.Fields(1)) Then C = miRSAux.Fields(1)
'    End If
'    miRSAux.Close
'    Set miRSAux = Nothing
'    If C = "" Then
'        MsgBox "Grupo PPal para el usuario: " & CadenaSeleccion & " NO encontrado", vbExclamation
'        Exit Sub
'    End If
'
'    'Llegado aqui, ponemos
'
'    'vC.userprop = Val(RecuperaValor(CadenaSeleccion, 1))
'    'vC.groupprop = Val(C)
    Text3(0).Text = RecuperaValor(CadenaSeleccion, 3)
    Text3(0).Tag = RecuperaValor(CadenaSeleccion, 1)
'    Text3(1).Text = C
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub Image3_Click()
    cd1.ShowOpen
    If cd1.FileName <> "" Then
        Text4.Text = cd1.FileName
    End If
End Sub





Private Sub CargaElArbolDeAmin()
Dim NodD As Node
Dim Nod As Node
Dim i As Integer

    Set TreeView1.ImageList = Admin.TreeView1.ImageList
    
    'El raiz
    Set Nod = Admin.TreeView1.Nodes(1)
    Set NodD = TreeView1.Nodes.Add(, , Nod.Key, Nod.Text, Nod.Image)

    'Insertamos el primero
    For i = 2 To Admin.TreeView1.Nodes.Count
        Set Nod = Admin.TreeView1.Nodes(i)
        If Nod.Parent Is Nothing Then
            Set NodD = TreeView1.Nodes.Add(, tvwChild, Nod.Key, Nod.Text, Nod.Image)
        Else
            Set NodD = TreeView1.Nodes.Add(Nod.Parent.Key, tvwChild, Nod.Key, Nod.Text, Nod.Image)
        End If
    Next i
    TreeView1.Nodes(2).EnsureVisible
End Sub





Private Function CopiaArchivosCarpetaRecursiva(No As Node) As String
Dim Nod As Node
Dim J As Integer
Dim i As Integer
Dim C As String

    'Primero copiamos la carpeta
    C = Mid(No.Key, 2) & "|"
        If No.Children > 0 Then
            J = No.Children
            Set Nod = No.Child
            For i = 1 To J
               C = C & CopiaArchivosCarpetaRecursiva(Nod)
               If i <> J Then Set Nod = Nod.Next
            Next i
        End If
    CopiaArchivosCarpetaRecursiva = C
End Function
    

Private Function InsertarModificarPlantillaCarpeta() As Byte
Dim cad As String
    On Error GoTo EInsertarModificarPlantillaCarpeta
    InsertarModificarPlantillaCarpeta = 0
    If origenDestino = "" Then
        cad = "Select max(carpeta) from plantillacarpetas"
        Set miRSAux = New ADODB.Recordset
        miRSAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        origenDestino = 1
        If Not miRSAux.EOF Then origenDestino = CStr(DBLet(miRSAux.Fields(0), "N") + 1)
        miRSAux.Close
        If Val(origenDestino) > 128 Then
            MsgBox "Error asignado numero de carpeta: 128. Soporte técnico", vbExclamation
            Exit Function
        End If
        
        cad = "INSERT INTO plantillacarpetas (carpeta, texto, groupprop, lecturag) VALUES ("
        cad = cad & origenDestino & ",'" & DevNombreSql(Text5.Text) & "',1," & vbPermisoTotal & ")"
                
    Else
        'MODIFICAR
        cad = "UPDATE plantillacarpetas SET texto='" & DevNombreSql(Text5.Text) & "'"
        cad = cad & " WHERE carpeta = " & origenDestino
    End If
    Conn.Execute cad
    InsertarModificarPlantillaCarpeta = CByte(origenDestino)
    Exit Function
EInsertarModificarPlantillaCarpeta:
    MuestraError Err.Number, Err.Description
End Function


Private Sub CargaPlantillas()
    Set miRSAux = New ADODB.Recordset
    origenDestino = "Select * from plantillacarpetas where carpeta<>" & origenDestino
    miRSAux.Open origenDestino, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRSAux.EOF
        List1.AddItem miRSAux!Texto
        List1.ItemData(List1.NewIndex) = miRSAux!Carpeta
        miRSAux.MoveNext
    Wend
    miRSAux.Close
    Set miRSAux = Nothing
End Sub



Private Sub imgCambiaUserGroup_Click(Index As Integer)
    If Index = 0 Then
        Set frmU = New frmUsuarios2
        frmU.DatosADevolverBusqueda = "0|"
        frmU.Show vbModal
        Set frmU = Nothing
    Else
        Set frmG = New frmGrupos
        frmG.DatosADevolverBusqueda = "0|"
        frmG.Show vbModal
        Set frmG = Nothing
    End If
End Sub

Private Sub List1_DblClick()
    cmdSelCarpetaPlantilla_Click 0
End Sub

Private Sub TreeView1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub
