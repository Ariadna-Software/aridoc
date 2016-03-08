VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmListaUserPcs2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Usuarios/Equipos"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4050
   Icon            =   "frmListaUserPcs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   4050
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGrupo 
      Height          =   375
      Index           =   1
      Left            =   2400
      Picture         =   "frmListaUserPcs.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "ELIMINAR EQUIPO GESTION"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdRegresar 
      Cancel          =   -1  'True
      Caption         =   "Regresar"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdGrupo 
      Height          =   375
      Index           =   0
      Left            =   1920
      Picture         =   "frmListaUserPcs.frx":040C
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Cambiar INTEGRACION"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdUsuario 
      Height          =   375
      Index           =   2
      Left            =   2880
      Picture         =   "frmListaUserPcs.frx":050E
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdUsuario 
      Height          =   375
      Index           =   1
      Left            =   2400
      Picture         =   "frmListaUserPcs.frx":0610
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdUsuario 
      Height          =   375
      Index           =   0
      Left            =   1920
      Picture         =   "frmListaUserPcs.frx":0712
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   4920
      Width           =   1095
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4215
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   7435
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label lblPC 
      Caption         =   "EQUIPOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblUser 
      Caption         =   "USUARIOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmListaUserPcs2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcion As Byte

Dim SQL As String
Dim It As ListItem



Private Sub cmdGrupo_Click(Index As Integer)
Dim EsEsteEquipo As Boolean
Dim Equi As Integer


    If ListView1.SelectedItem Is Nothing Then Exit Sub
    Equi = CInt(Mid(ListView1.SelectedItem.Key, 2))
    
    If Index = 0 Then
    
    
        If ListView1.SelectedItem.SmallIcon <> 2 Then
            'El equipo ya tiene INTEGRCION, quitar
            SQL = "El equipo " & ListView1.SelectedItem.Text & " ya tiene integracion. ¿Desea quitarsela?"
            If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
                
            SQL = "UPDATE equipos SET exeintegra=NULL where codequipo=" & Equi
            Conn.Execute SQL
            ListView1.SelectedItem.SmallIcon = 2
            
            
        Else
            'Voy a ponerle el integrador
            DatosCopiados = ""
            frmPregunta.Opcion = 22
            frmPregunta.Show vbModal
            If DatosCopiados <> "" Then
                DatosCopiados = DevNombreSql(DatosCopiados)
                SQL = "UPDATE equipos SET exeintegra='" & DatosCopiados & "' where codequipo=" & Equi
                Conn.Execute SQL
                ListView1.SelectedItem.SmallIcon = 21
            End If
        End If
    Else
        'ELIMINAR EQUIPO
        EsEsteEquipo = False
        
        If Equi = vUsu.PC Then
            EsEsteEquipo = True
            SQL = "Quiere eliminar de la gestion este equipo. ¿Esta seguro?"
         
        Else
            SQL = "¿Seguro que desea elimiminar de la gestion documental el equipo " & ListView1.SelectedItem.Text & "?"
        End If
        If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        
        If MsgBox("El proceso es irreversible. ¿Continuar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        
        
        SQL = "Delete from extensionPC where codequipo=" & Equi
        Conn.Execute SQL
        
        SQL = "DELETE from equipos where codequipo=" & Equi
        Conn.Execute SQL
        
        ListView1.ListItems.Remove ListView1.SelectedItem.Index
        
        If EsEsteEquipo Then
            MsgBox "La aplicacion finalizará", vbExclamation
            End
        End If
    End If
End Sub

Private Sub cmdRegresar_Click()
Dim i As Integer
    SQL = ""
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked Then
            SQL = "OK"
            InsertaTemporal CLng(Mid(ListView1.ListItems(i).Key, 2))
        End If
    Next i
    
    If SQL = "" Then Exit Sub
    
    DatosCopiados = "OK"
    Unload Me
End Sub

'Private Sub cmdUsuario_Click(Index As Integer)
'Dim vUs As Cusuarios
'
'    If Index = 0 Then
'        'Nuevo
'        DatosModificados = False
'        Set frmUsuario.vU = Nothing
'        frmUsuario.Show vbModal
'        If DatosModificados Then CargaUsuarios
'
'    Else
'        If ListView1.SelectedItem Is Nothing Then
'            MsgBox "Seleccione un usuario", vbExclamation
'            Exit Sub
'        End If
'        Set vUs = New Cusuarios
'        If vUs.Leer(CInt(Mid(ListView1.SelectedItem.Key, 2))) = 0 Then
'            'Leeido con exito
'            'Vemos si intenta cambiar datos del
'            'usuario actual
'            If vUs.codusu = vUsu.codusu Then
'                Sql = "Intenta modificar datos del usuario actual." & vbCrLf
'                Sql = Sql & "Al finalizar deberá reiniciar la aplicación" & vbCrLf & vbCrLf
'                Sql = Sql & "       ¿Desea continuar?"
'                If MsgBox(Sql, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
'            End If
'            If Index = 2 Then
'                    'ELIMINAR
'                Sql = "Desea elimniar el usuario: " & vUs.codusu & " - " & vUs.Nombre & "?"
'                If MsgBox(Sql, vbQuestion + vbYesNoCancel) = vbYes Then
'                    If vUs.Eliminar = 0 Then
'                        If vUsu.codusu = vUs.codusu Then
'                            MsgBox "El programa  finalizará", vbCritical
'                            End
'                            Exit Sub
'                        Else
'                            CargaUsuarios
'                        End If
'                    End If
'                End If
'            Else
'
'                DatosModificados = False
'                Set frmUsuario.vU = vUs
'                frmUsuario.Show vbModal
'                If DatosModificados Then
'                    If vUsu.codusu = vUs.codusu Then
'                        MsgBox "El programa  finalizará", vbCritical
'                        End
'                        Exit Sub
'                    Else
'                        CargaUsuarios
'                    End If
'                End If 'De datos modificados
'            End If 'index=2
'        End If
'    End If
'End Sub

Private Sub Command1_Click()
    DatosCopiados = ""
    Unload Me
End Sub

Private Sub Form_Load()
Dim B As Boolean
    
    Me.ListView1.SmallIcons = Admin.ImgUsersPCs
    B = (Opcion = 0)
    Me.ListView1.Checkboxes = B
    Me.lblPC.Visible = Not B
    Me.cmdGrupo(0).Visible = Not B
    Me.cmdGrupo(1).Visible = Not B
    Me.lblUser.Visible = B
    cmdRegresar.Visible = B
'    Me.cmdUsuario(0).Visible = B
'    Me.cmdUsuario(1).Visible = B
'    Me.cmdUsuario(2).Visible = B
    Caption = "Admnistracion de "
    If Opcion = 0 Then
        Caption = Caption & "usuarios"
        CargaUsuarios
        DatosCopiados = ""
    Else
        Caption = Caption & "equipos"
        CargaEquipos
    End If
End Sub


Private Sub CargaUsuarios()
    Screen.MousePointer = vbHourglass
    If DatosCopiados <> "" Then DatosCopiados = "|" & DatosCopiados
    ListView1.ListItems.Clear
    Set miRSAux = New ADODB.Recordset
    SQL = "Select codusu,nombre from Usuarios where codusu <>" & vUsu.codusu
    SQL = SQL & " ORDER By nombre"
    miRSAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRSAux.EOF
        Set It = ListView1.ListItems.Add(, "C" & miRSAux!codusu)
        It.Text = miRSAux!Nombre
        It.SmallIcon = 5
        
        'Veamos si va checked o no
        SQL = "|" & miRSAux!codusu & "|"
        If InStr(1, DatosCopiados, SQL) > 0 Then It.Checked = True

        
        
        
        miRSAux.MoveNext
    Wend
    miRSAux.Close
    Set miRSAux = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub CargaEquipos()

    Set miRSAux = New ADODB.Recordset
    SQL = "Select codequipo,descripcion,exeintegra from equipos "
    SQL = SQL & " ORDER By codequipo"
    miRSAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRSAux.EOF
        Set It = ListView1.ListItems.Add(, "C" & miRSAux!codequipo)
        It.Text = miRSAux!Descripcion
        
        If Not IsNull(miRSAux!exeintegra) Then
            It.SmallIcon = 21
        Else
            It.SmallIcon = 2
        End If
        miRSAux.MoveNext
    Wend
    miRSAux.Close
    Set miRSAux = Nothing
End Sub

Private Sub imgUsus_Click()

End Sub
