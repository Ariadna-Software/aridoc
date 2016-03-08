VERSION 5.00
Begin VB.Form frmEMAIL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Envio de e-mail"
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   Icon            =   "frmEMAIL.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   6600
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   5775
   End
End
Attribute VB_Name = "frmEMAIL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public IdMail As Long
Public ListaDeFicheros As String

Dim PrimeraVez As Boolean
Dim cad As String
Dim Salir As Boolean

Private Sub Enviar()
Dim imageContentID, success
Dim mailman As ChilkatMailMan
Dim Cuerpo As String


    On Error GoTo GotException
    Label1.Caption = "Obteniendo datos"
    Label1.Refresh
    Set miRSAux = New ADODB.Recordset
    
    Set listacod = Nothing
    Set listacod = New Collection
    
    cad = "Select * from maildestext where codmail=" & IdMail
    miRSAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRSAux.EOF
        listacod.Add miRSAux!Nombre & "|" & miRSAux!mail & "|"
        miRSAux.MoveNext
    Wend
    miRSAux.Close
    
    cad = "select maill.*,mailc.destino,usuarios.nombre,usuarios.maildir from mailc,maill,usuarios where mailc.codmail"
    cad = cad & " = maill.codmail And usuarios.codusu = mailc.Destino"
    cad = cad & " and maill.codmail=" & IdMail
    miRSAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRSAux.EOF
        listacod.Add miRSAux!Nombre & "|" & miRSAux!maildir & "|"
        miRSAux.MoveNext
    Wend
    miRSAux.Close
    
    
    If listacod.Count = 0 Then
        'MAAAAAl
        MsgBox "Error leyendo MSG: " & IdMail & vbCrLf & "NO SE HA ENVIADO EL MENSAJE. Destinatarios=0", vbCritical
        Set miRSAux = Nothing
        Exit Sub
    End If
    
    
    'Leemos el mensaje
    cad = "select maill.* from maill Where maill.codmail = " & IdMail
    miRSAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRSAux.EOF Then
        MsgBox "Error leyendo MSG: " & IdMail & vbCrLf & "NO SE HA ENVIADO EL MENSAJE. Codmail NO encontrado", vbCritical
        Set miRSAux = Nothing
        Exit Sub
    End If
    Set mailman = New ChilkatMailMan
    
    
    Label1.Caption = "Generando mensaje email"
    Label1.Refresh
    
    'Esta cadena es constante de la lincencia comprada a CHILKAT
    mailman.UnlockComponent "1AriadnaMAIL_BOVuuRWYpC9f"
    mailman.LogMailSentFilename = ""    'App.path & "\mailSent.log"

    mailman.SmtpHost = vUsu.e_server
    mailman.SmtpUsername = vUsu.e_login
    mailman.SmtpPassword = vUsu.e_pwd
    'mailman.SmtpLoginDomain = vUsu.e_login

    'mailman.SmtpAuthMethod = "NONE"
    mailman.SmtpAuthMethod = "LOGIN"
    
    
'   mailman.SmtpLoginDomain = vUsu.e_login
    
    ' Create the email, add content, address it, and sent it.
    Dim email As ChilkatEmail
    Set email = New ChilkatEmail
    
    
    If ListaDeFicheros <> "" Then
        Label1.Caption = "Adjuntando ficheros"
        Label1.Refresh
        
        Do
            imageContentID = InStr(1, ListaDeFicheros, "|")
            If imageContentID > 0 Then
                cad = Mid(ListaDeFicheros, 1, Val(imageContentID) - 1)
                ListaDeFicheros = Mid(ListaDeFicheros, Val(imageContentID) + 1)
            
                'Agregamos fichero
                Label1.Caption = cad
                Label1.Refresh
                email.AddFileAttachment cad
            Else
                ListaDeFicheros = ""
            End If
        Loop Until ListaDeFicheros = ""
        
    End If
    
    
    
    'Fijamos el cuerpo del mensaje
    FijarCampoMemo Cuerpo
    
    email.Subject = miRSAux!asunto
    email.AddPlainTextAlternativeBody "Programa lector e-mail NO soporta HTML. " & vbCrLf & Cuerpo
    email.From = vUsu.e_dir
    
    'La imagen
    espera 0.2
   ' imageContentID = email.AddRelatedContent(App.Path & "\minilogo.dat")
    espera 0.2
    
    cad = "<!DOCTYPE HTML PUBLIC " & Chr(34) & "-//W3C//DTD HTML 4.0 Transitional//EN" & Chr(34) & ">"
    cad = cad & "<HTML><HEAD><TITLE>Mensaje</TITLE></HEAD>"
    cad = cad & "<TABLE BORDER=""0"" CELLSPACING=1 CELLPADDING=0 WIDTH=576>"
    'Cuerpo del mensaje
    cad = cad & "<TR><TD VALIGN=""TOP""><P>"
    FijarTextoMensaje Cuerpo
    cad = cad & "</P></TD></TR>"
    
    cad = cad & "<TR><TD VALIGN=""TOP""><P><hr></P>"
    'La imagen
    cad = cad & "<P ALIGN=""CENTER"">"
    '<IMG SRC=" & Chr(34) & "cid:" & imageContentID & Chr(34) & ">
    cad = cad & "</P>"
    'cad = cad & "<P ALIGN=""CENTER""><FONT SIZE=2>Mensaje creado desde el programa ARIDOC de"
    'cad = cad & "<A HREF=""http://www.ariadnasoftware.com/"">Ariadna&nbsp;"
    'cad = cad & "Software S.L.</A></P><P ALIGN=""CENTER""></P>"
    cad = cad & "<P>Este correo electrónico y sus documentos adjuntos estan dirigidos UNICA Y EXCLUSIVAMENTE a "
    cad = cad & " los destinatarios especificados. La información contenida puesde ser CONFIDENCIAL"
    cad = cad & " y/o estar LEGALMENTE PROTEGIDA.</P>"
    cad = cad & "<P>Si usted recibe este mensaje por ERROR, por favor comuníqueselo inmediatamente al"
    cad = cad & " remitente y ELIMINELO ya que usted NO ESTA AUTORIZADO al uso, revelación, distribución,"
    cad = cad & " impresión o copia de toda o alguna parte de la información en él contenida. Muchas Gracias "
    cad = cad & ".</FONT></P><P><HR ALIGN=""LEFT"" SIZE=1></TD>"
    cad = cad & "</TR></TABLE></BODY></HTML>"
    
    email.SetHtmlBody (cad)
    
    
    
    miRSAux.Close
    
'    While Not miRSAux.EOF
'
'        email.AddTo miRSAux!Nombre, miRSAux!maildir
'        miRSAux.MoveNext
'    Wend
    
    For IdMail = 1 To listacod.Count
        cad = listacod(IdMail)
        Cuerpo = RecuperaValor(cad, 1)
        cad = RecuperaValor(cad, 2)
        email.AddTo Cuerpo, cad
    Next IdMail
    
    Set miRSAux = Nothing
            
    
'    If Opcion = 0 Then
'        'ADjunatmos el PDF
'        email.AddFileAttachment App.Path & "\docum.pdf"
'    End If
'
    
    
    Label1.Caption = "Conectando servidor"
    Label1.Refresh
    
    
    'email.SendEncrypted = 1
    success = mailman.SendEmail(email)
    If (success = 1) Then

    Else
        cad = "Han ocurrido errores durante el envio.Compruebe el archivo log.xml para mas informacion"
        mailman.SaveXmlLog App.Path & "\log.xml"
        MsgBox cad, vbExclamation
        UpdatearEmail_NO
    End If
    
    
GotException:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
        UpdatearEmail_NO
    End If
    Set miRSAux = Nothing
    Set email = Nothing
    Set mailman = Nothing
    'UPDATEO LA BD para que NO marque como enviados por mail
    
    
End Sub

Private Sub UpdatearEmail_NO()
    cad = "UPDATE mailc set email=0 where codmail =" & IdMail
    Conn.Execute cad
End Sub


Private Sub FijarTextoMensaje(ByRef C1 As String)
Dim i As Integer
Dim J As Integer

    J = 1
    Do
        i = InStr(J, C1, vbCrLf)
        If i > 0 Then
              cad = cad & Mid(C1, J, i - J) & "</P><BR><P>"
        Else
            cad = cad & Mid(C1, J)
        End If
        J = i + 2
    Loop Until i = 0
End Sub

Private Sub FijarCampoMemo(ByRef CADENA As String)
    On Error Resume Next
    CADENA = miRSAux!Texto
    If Err.Number <> 0 Then
        Err.Clear
        CADENA = vbCrLf
    End If
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Screen.MousePointer = vbHourglass
        Enviar
        Salir = True
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()
    Salir = False
    PrimeraVez = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not Salir Then Cancel = 1
End Sub
