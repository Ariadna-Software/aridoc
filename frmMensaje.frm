VERSION 5.00
Begin VB.Form frmMensaje2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mensaje"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8790
   ForeColor       =   &H00000000&
   Icon            =   "frmMensaje.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   8790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Enviar"
      Height          =   375
      Left            =   6480
      TabIndex        =   6
      Top             =   6000
      Width           =   975
   End
   Begin VB.CommandButton cmdCerrar 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   7560
      TabIndex        =   7
      Top             =   6000
      Width           =   975
   End
   Begin VB.Frame FrameEnviar 
      Height          =   5895
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   8775
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   5520
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   1560
         Width           =   3015
      End
      Begin VB.TextBox Text2 
         Height          =   3165
         Index           =   2
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   5
         Text            =   "frmMensaje.frx":6852
         Top             =   2520
         Width           =   8295
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   3
         Text            =   "Text2"
         Top             =   2040
         Width           =   7455
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Enviar por e-mail"
         Height          =   255
         Left            =   4920
         TabIndex        =   1
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   0
         Left            =   5640
         TabIndex        =   19
         Text            =   "Text2"
         Top             =   360
         Width           =   1335
      End
      Begin VB.ListBox List1 
         Height          =   1620
         Left            =   1080
         TabIndex        =   0
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo"
         Height          =   255
         Index           =   2
         Left            =   4920
         TabIndex        =   21
         Top             =   1560
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   720
         Picture         =   "frmMensaje.frx":6858
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Asunto"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   20
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Fecha"
         Height          =   255
         Index           =   0
         Left            =   4800
         TabIndex        =   18
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Para"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame FrameRecibido 
      Height          =   5895
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   8775
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   1320
         Width           =   4455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   720
         Width           =   6735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   240
         Width           =   4335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   4005
         Index           =   2
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Text            =   "frmMensaje.frx":D0AA
         Top             =   1800
         Width           =   8535
      End
      Begin VB.CheckBox Check1 
         Caption         =   "E-mail"
         Enabled         =   0   'False
         Height          =   255
         Left            =   6840
         TabIndex        =   8
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   23
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         Height          =   255
         Index           =   2
         Left            =   5520
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "PARA"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Asunto"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmMensaje2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enviado As Boolean
Public Opcion As Byte
Public vM As Cmailc


    '0  - NUEVO
Dim i As Integer
Dim PrimeraVez  As Boolean



Private Sub Check2_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub cmdCerrar_Click()
    Set vM = Nothing
    Unload Me
End Sub


Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Command1_Click()
Dim CadenaPara As String
    If List1.ListCount = 0 Then
        MsgBox "Selecciona un destinatario", vbExclamation
        Exit Sub
    End If
        
    Text2(1).Text = Trim(Text2(1).Text)
    If Text2(1).Text = "" Then
        MsgBox "Asunto no puede estar vacio", vbExclamation
        Exit Sub
    End If
    
    If Combo1.ListIndex < 0 Then
        MsgBox "Seleccione el tipo de mensaje", vbExclamation
        Exit Sub
    End If
    
    
    'Si tiene enviar e-mail deberiamos comprobar que todos tienen
    'direccion e-mail
    BorrarTemporal1
    Set listacod = Nothing
    Set listacod = New Collection
    CadenaPara = ""
    For i = 0 To List1.ListCount - 1
        If CadenaPara <> "" Then CadenaPara = CadenaPara & ";"
        CadenaPara = CadenaPara & List1.List(i)
        InsertaTemporal List1.ItemData(i)
        listacod.Add List1.ItemData(i)
    Next i
    
    If Len(CadenaPara) > 255 Then CadenaPara = Mid(CadenaPara, 1, 251) & " ..."
        
    
    'Tomo prestado esta variable
    Set listaimpresion = Nothing
    Set listaimpresion = New Collection
    
    Set miRSAux = New ADODB.Recordset
    DatosCopiados = "Select nombre,usuarios.codusu from tmpFich,usuarios Where imagen = usuarios.codusu"
    DatosCopiados = DatosCopiados & " AND tmpFich.codusu =" & vUsu.codusu
    DatosCopiados = DatosCopiados & " AND codequipo= " & vUsu.PC & " AND (maildir ='' or (maildir is null))"
    miRSAux.Open DatosCopiados, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    DatosCopiados = ""
    While Not miRSAux.EOF
        DatosCopiados = DatosCopiados & miRSAux!Nombre & vbCrLf
       
        listaimpresion.Add CStr(miRSAux!codusu)
        miRSAux.MoveNext
    Wend
    miRSAux.Close
    Set miRSAux = Nothing
    
    If Check1.Value Then
        If DatosCopiados <> "" Then
            DatosCopiados = "Los siguientes usuarios no tienen direccion e-mail:" & vbCrLf & vbCrLf & DatosCopiados
            DatosCopiados = DatosCopiados & vbCrLf & vbCrLf & "¿Desea continuar?"
            If MsgBox(DatosCopiados, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
        
    End If
    
    'Llegados aqui, creamos el mensaje.
    
        Set vM = Nothing
        Set vM = New Cmailc
        
    vM.asunto = Text2(1).Text
    vM.Fecha = CDate(Text2(0).Text)
    vM.Origen = vUsu.codusu
    vM.Texto = Text2(2).Text
    vM.email = Abs(Check2.Value)
    vM.Tipo = Combo1.ItemData(Combo1.ListIndex)
    If vM.GenerarMensaje(listacod, CadenaPara) = 1 Then
        'Borramos
        Conn.Execute "Delete from maill where codmail =" & vM.codmail
        Conn.Execute "Delete from mailc where codmail =" & vM.codmail
        Conn.Execute "Delete from maile where codmail =" & vM.codmail
    Else
        'Ha ido todo bien
        If Check2.Value = 1 Then
            'HAY QUE ENVIAR POR MAIL, excepto los que no tienen mail
            DatosCopiados = "UPDATE mailc SET email=0 where"
            DatosCopiados = DatosCopiados & " origen = " & vUsu.codusu & " and codmail = " & vM.codmail
            DatosCopiados = DatosCopiados & " AND destino = "
            If Not listaimpresion Is Nothing Then
                For i = 1 To listaimpresion.Count
                    Conn.Execute DatosCopiados & listaimpresion(i)
                Next i
                'PARA Abriremos la pantalla de envio de mail
                i = List1.ListCount - listaimpresion.Count
                
            Else
                i = 0
            End If
            
            
            
            
            If i > 0 Then
                'HAY que enviar mensajes
                frmEMAIL.IdMail = vM.codmail
                frmEMAIL.Show vbModal
            End If
            'PONERMO A NOTINH ALGUNOS VALORES
            Set listaimpresion = Nothing
        End If
    End If
    Set listacod = Nothing
    Set vM = Nothing
    DatosMOdificados = True
    Unload Me
        
End Sub


Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        If Opcion = 0 Then
            If List1.ListCount = 0 Then
                List1.SetFocus
            Else
                Text2(2).SetFocus
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
Dim H As Integer
Dim W As Integer
    PrimeraVez = True
    Limpiar Me
    Me.FrameEnviar.Visible = False
    Me.FrameRecibido.Visible = False
    If Opcion = 0 Then
        Me.FrameEnviar.Visible = True
        H = FrameEnviar.Height
        W = FrameEnviar.Width
        
        If vM Is Nothing Then
            Text2(0).Text = Format(Now, "dd/mm/yyyy")
            Check2.Enabled = (vUsu.e_server <> "")
                 
        Else
            'Es un reenvio o respuesta
            Text2(0).Text = Format(vM.Fecha, "dd/mm/yyyy")
            Text2(1).Text = vM.asunto
            Text2(2).Text = vM.Texto
            'Añadimos el usario
            List1.Clear
            
            If vM.email Then
                Check2.Value = 1
            Else
                Check2.Value = 0
            End If
            
            PonCampos
        End If
        Set vM = Nothing
        'Text2(0).BackColor = CLng("&H80000018")
        Command1.Visible = True
    Else
        Command1.Visible = False
        FrameRecibido.Visible = True
        H = FrameRecibido.Height
        W = FrameRecibido.Width
        
        'Ahora ponemos los campos del mensaje ande corresponda
        Text1(0).Text = ""
        Text1(4).Text = ""
        Text1(1).Text = Format(vM.Fecha, "dd/mm/yyyy")
        Text1(2).Text = vM.Texto
        Text1(3).Text = vM.asunto
        PonCampos
                
        
    End If
    
    Me.Width = W + 120
    Me.Height = H + 920
    

   
    Combo1.Clear
    For H = 0 To TotalTipos
        If ArrayTipoMen(H).Descripcion <> "" Then
            Combo1.AddItem ArrayTipoMen(H).Descripcion
            Combo1.ItemData(Combo1.NewIndex) = H
        End If
    Next H
    Combo1.ListIndex = 0

End Sub


Private Sub PonCampos()
    On Error GoTo EPonCa
    If Opcion = 1 Then
        If vM.Recibido Then
            Label1(0).Caption = "DE"
            DatosCopiados = DevuelveDesdeBD("nombre", "usuarios", "codusu", CStr(vM.Origen), "N")
        Else
            Label1(0).Caption = "PARA"
            DatosCopiados = vM.Textopara
        End If
        Text1(0).Text = DatosCopiados
        DatosCopiados = ArrayTipoMen(vM.Tipo).Descripcion
        Text1(4).Text = DatosCopiados
        
    Else
        
        If vM.Destino >= 0 Then
                DatosCopiados = DevuelveDesdeBD("nombre", "usuarios", "codusu", CStr(vM.Destino), "N")
                If DatosCopiados <> "" Then
                    List1.AddItem DatosCopiados
                    List1.ItemData(List1.NewIndex) = vM.Destino
                End If
        End If
            
        For i = 0 To Combo1.ListCount - 1
            If Combo1.ItemData(i) = vM.Tipo Then
                'Es este
                Combo1.ListIndex = i
                Exit For
            End If
        Next i
    End If
    Exit Sub
EPonCa:
    MuestraError Err.Number, "Poniendo campos(2)"
End Sub

Private Sub Image1_Click()


    'Borramos temporal
    BorrarTemporal1
    
    'Insertamos
    DatosCopiados = ""
    For i = 0 To List1.ListCount - 1
        DatosCopiados = DatosCopiados & List1.ItemData(i) & "|"
    Next

    frmListaUserPcs2.Opcion = 0
    frmListaUserPcs2.Show vbModal
    If DatosCopiados <> "" Then
        List1.Clear
'        DatosCopiados = "Select nombre,codusu from tmpFich,usaurios where codusu =" & vUsu.codusu
'        DatosCopiados = DatosCopiados & " AND codpc= " & vUsu.PC
'        DatosCopiados = DatosCopiados & " AND tmpfich.imagen = usuarios.codusu"
'        DatosCopiados = DatosCopiados & " ORDER BY nomusu"
'
        DatosCopiados = "Select nombre,usuarios.codusu from tmpFich,usuarios Where imagen = usuarios.codusu"
        DatosCopiados = DatosCopiados & " AND tmpFich.codusu =" & vUsu.codusu
        DatosCopiados = DatosCopiados & " AND codequipo= " & vUsu.PC & "  ORDER BY nombre"
        
        
        Set miRSAux = New ADODB.Recordset
        miRSAux.Open DatosCopiados, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not miRSAux.EOF
            List1.AddItem miRSAux!Nombre
            List1.ItemData(List1.NewIndex) = miRSAux!codusu
            miRSAux.MoveNext
        Wend
        miRSAux.Close
        Set miRSAux = Nothing
    End If
End Sub



Private Sub List1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
        Image1_Click
    Else
        KEYpress KeyAscii
    End If
End Sub



Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index <> 2 Then KEYpress KeyAscii
End Sub
