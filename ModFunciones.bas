Attribute VB_Name = "ModFunciones"
'////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////
'   En este modulo estan las funciones que recorren el form
'   usando el each for
'   Estas son
'
'   CamposSiguiente -> Nos devuelve el el text siguiente en
'           el orden del tabindex
'
'   CompForm -> Compara los valores con su tag
'
'   InsertarDesdeForm - > Crea el sql de insert e inserta
'
'   Limpiar -> Pone a "" todos los objetos text de un form
'
'   ObtenerBusqueda -> A partir de los text crea el sql a
'       partir del WHERE ( sin el).
'
'   ModifcarDesdeFormulario -> Opcion modificar. Genera el SQL
'
'   PonerDatosForma -> Pone los datos del RECORDSET en el form
'////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////
Option Explicit

Public Const ValorNulo = "Null"

Public Function CompForm(ByRef Formulario As Form) As Boolean
    Dim Control As Object
    Dim mTag As CTag
    Dim Carga As Boolean
    Dim Correcto As Boolean
       
    CompForm = False
    Set mTag = New CTag
    For Each Control In Formulario.Controls
        'TEXT BOX
        If TypeOf Control Is TextBox Then
            Carga = mTag.Cargar(Control)
            If Carga = True Then
                Correcto = mTag.Comprobar(Control)
                If Not Correcto Then Exit Function
            Else
                MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                Exit Function
            End If
        'COMBOBOX
        ElseIf TypeOf Control Is ComboBox Then
            'Comprueba que los campos estan bien puestos
            If Control.Tag <> "" Then
                Carga = mTag.Cargar(Control)
                If Carga = False Then
                    MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                    Exit Function
                    
                Else
                    If mTag.Vacio = "N" And Control.ListIndex < 0 Then
                            MsgBox "Seleccione una dato para: " & mTag.Nombre, vbExclamation
                            Exit Function
                    End If
                End If
            End If
        End If
    Next Control
    CompForm = True
End Function


Public Sub Limpiar(ByRef Formulario As Form)
    Dim Control As Object
    
    For Each Control In Formulario.Controls
        If TypeOf Control Is TextBox Then
            Control.Text = ""
        End If
    Next Control
End Sub


Public Function CampoSiguiente(ByRef Formulario As Form, Valor As Integer) As Control
Dim Fin As Boolean
Dim Control As Object

On Error GoTo ECampoSiguiente

    'Debug.Print "Llamada:  " & Valor
    'Vemos cual es el siguiente
    Do
        Valor = Valor + 1
        For Each Control In Formulario.Controls
            'Debug.Print "-> " & Control.Name & " - " & Control.TabIndex
            'Si es texto monta esta parte de sql
            If Control.TabIndex = Valor Then
                    Set CampoSiguiente = Control
                    Fin = True
                    Exit For
            End If
        Next Control
        If Not Fin Then
            Valor = -1
        End If
    Loop Until Fin
    Exit Function
ECampoSiguiente:
    Set CampoSiguiente = Nothing
    Err.Clear
End Function




Private Function ValorParaSQL(Valor, ByRef vTag As CTag) As String
Dim Dev As String
Dim d As Single
Dim I As Integer
Dim V
    Dev = ""
    If Valor <> "" Then
        Select Case vTag.TipoDato
        Case "N"
            V = Valor
            If InStr(1, Valor, ",") Then
                If InStr(1, Valor, ".") Then
                    'ABRIL 2004
                
                    'Ademas de la coma lleva puntos
                    V = ImporteFormateado(CStr(Valor))
                    Valor = V
                Else
                
                    V = CSng(Valor)
                    Valor = V
                End If
            Else
         
            End If
            Dev = TransformaComasPuntos(CStr(Valor))
            
        Case "F"
            Dev = "'" & Format(Valor, FormatoFecha) & "'"
        Case "T"
            Dev = CStr(Valor)
            Dev = "'" & Dev & "'"
        Case Else
            Dev = "'" & Valor & "'"
        End Select
        
    Else
        'Si se permiten nulos, la "" ponemos un NULL
        If vTag.Vacio = "S" Then Dev = ValorNulo
    End If
    ValorParaSQL = Dev
End Function

Public Function InsertarDesdeForm(ByRef Formulario As Form) As Boolean
    Dim Control As Object
    Dim mTag As CTag
    Dim Izda As String
    Dim Der As String
    Dim Cad As String
    
    On Error GoTo EInsertarF
    'Exit Function
    Set mTag = New CTag
    InsertarDesdeForm = False
    Der = ""
    Izda = ""
    For Each Control In Formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If mTag.Columna <> "" Then
                        If Izda <> "" Then Izda = Izda & ","
                        'Access
                        'Izda = Izda & "[" & mTag.Columna & "]"
                        Izda = Izda & "" & mTag.Columna & ""
                    
                        'Parte VALUES
                        Cad = ValorParaSQL(Control.Text, mTag)
                        If Der <> "" Then Der = Der & ","
                        Der = Der & Cad
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Izda <> "" Then Izda = Izda & ","
                'Access
                'Izda = Izda & "[" & mTag.Columna & "]"
                Izda = Izda & "" & mTag.Columna & ""
                If Control.Value = 1 Then
                    Cad = "1"
                    Else
                    Cad = "0"
                End If
                If Der <> "" Then Der = Der & ","
                If mTag.TipoDato = "N" Then Cad = Abs(CBool(Cad))
                Der = Der & Cad
            End If
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Izda <> "" Then Izda = Izda & ","
                    'Izda = Izda & "[" & mTag.Columna & "]"
                    Izda = Izda & "" & mTag.Columna & ""
                    If Control.ListIndex = -1 Then
                        Cad = ValorNulo
                        Else
                        Cad = Control.ItemData(Control.ListIndex)
                    End If
                    If Der <> "" Then Der = Der & ","
                    Der = Der & Cad
                End If
            End If
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo
    'INSERT INTO Empleados (Nombre,Apellido, Cargo) VALUES ('Carlos', 'Sesma', 'Prácticas');
    
    Cad = "INSERT INTO " & mTag.Tabla
    Cad = Cad & " (" & Izda & ") VALUES (" & Der & ");"
    
   
    Conn.Execute Cad, , adCmdText
    
    
    InsertarDesdeForm = True
Exit Function
EInsertarF:
    MuestraError Err.Number, "Inserta. "
End Function





Public Function PonerCamposForma(ByRef Formulario As Form, ByRef vData As Adodc) As Boolean
    Dim Control As Object
    Dim mTag As CTag
    Dim Cad As String
    Dim Valor As Variant
    Dim Campo As String  'Campo en la base de datos
    Dim I As Integer

    Set mTag = New CTag
    PonerCamposForma = False

    For Each Control In Formulario.Controls
        'TEXTO
        'Debug.Print Control.Tag
        If TypeOf Control Is TextBox Then
            'Comprobamos que tenga tag
            mTag.Cargar Control
            If Control.Tag <> "" Then
                If mTag.Cargado Then
                    'Columna en la BD
                    If mTag.Columna <> "" Then
                        Campo = mTag.Columna
                        If mTag.Vacio = "S" Then
                            Valor = DBLet(vData.Recordset.Fields(Campo))
                        Else
                            Valor = vData.Recordset.Fields(Campo)
                        End If
                        If mTag.Formato <> "" And CStr(Valor) <> "" Then
                            If mTag.TipoDato = "N" Then
                                'Es numerico, entonces formatearemos y sustituiremos
                                ' La coma por el punto
                                Cad = Format(Valor, mTag.Formato)
                                'Antiguo
                                'Control.Text = TransformaComasPuntos(cad)
                                'nuevo
                                Control.Text = Cad
                            Else
                                Control.Text = Format(Valor, mTag.Formato)
                            End If
                        Else
                            Control.Text = Valor
                        End If
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    'Columna en la BD
                    Campo = mTag.Columna
                    If mTag.Vacio = "S" Then
                        Valor = DBLet(vData.Recordset.Fields(Campo), mTag.TipoDato)
                    Else
                        Valor = vData.Recordset.Fields(Campo)
                    End If
                    Else
                        Valor = 0
                End If
                Control.Value = Valor
            End If
            
         'COMBOBOX
         ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    Campo = mTag.Columna
                    Valor = vData.Recordset.Fields(Campo)
                    I = 0
                    For I = 0 To Control.ListCount - 1
                        If Control.ItemData(I) = Val(Valor) Then
                            Control.ListIndex = I
                            Exit For
                        End If
                    Next I
                    If I = Control.ListCount Then Control.ListIndex = -1
                End If 'de cargado
            End If 'de <>""
        End If
    Next Control
    
    'Veremos que tal
    PonerCamposForma = True
Exit Function
EPonerCamposForma:
    MuestraError Err.Number, "Poner campos formulario. "
End Function

Private Function ObtenerMaximoMinimo(ByRef vSQL As String) As String
Dim Rs As Recordset
ObtenerMaximoMinimo = ""
Set Rs = New ADODB.Recordset
Rs.Open vSQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
If Not Rs.EOF Then
    If Not IsNull(Rs.EOF) Then
        ObtenerMaximoMinimo = CStr(Rs.Fields(0))
    End If
End If
Rs.Close
Set Rs = Nothing
End Function


Public Function ObtenerBusqueda(ByRef Formulario As Form) As String
    Dim Control As Object
    Dim Carga As Boolean
    Dim mTag As CTag
    Dim Aux As String
    Dim Cad As String
    Dim Sql As String
    Dim Tabla As String
    Dim Rc As Byte

    On Error GoTo EObtenerBusqueda

    'Exit Function
    Set mTag = New CTag
    ObtenerBusqueda = ""
    Sql = ""

    'Recorremos los text en busca de ">>" o "<<"
    For Each Control In Formulario.Controls
        If TypeOf Control Is TextBox Then
            Aux = Trim(Control.Text)
            If Aux = ">>" Or Aux = "<<" Then
                Carga = mTag.Cargar(Control)
                If Carga Then
                    If Aux = ">>" Then
                        Cad = " MAX(" & mTag.Columna & ")"
                    Else
                        Cad = " MIN(" & mTag.Columna & ")"
                    End If
                    Sql = "Select " & Cad & " from " & mTag.Tabla
                    Sql = ObtenerMaximoMinimo(Sql)
                    Select Case mTag.TipoDato
                    Case "N"
                        Sql = mTag.Tabla & "." & mTag.Columna & " = " & TransformaComasPuntos(Sql)
                    Case "F"
                        Sql = mTag.Tabla & "." & mTag.Columna & " = '" & Format(Sql, "yyyy-mm-dd") & "'"
                    Case Else
                        Sql = mTag.Tabla & "." & mTag.Columna & " = '" & Sql & "'"
                    End Select
                    Sql = "(" & Sql & ")"
                End If
            End If
        End If
    Next

    
    
    'Recorremos los text en busca del NULL
    For Each Control In Formulario.Controls
        If TypeOf Control Is TextBox Then
            Aux = Trim(Control.Text)
            If UCase(Aux) = "NULL" Then
                Carga = mTag.Cargar(Control)
                If Carga Then

                    Sql = mTag.Tabla & "." & mTag.Columna & " is NULL"
                    Sql = "(" & Sql & ")"
                    Control.Text = ""
                End If
            End If
        End If
    Next

    

    'Recorremos los textbox
    For Each Control In Formulario.Controls
        If TypeOf Control Is TextBox Then
            'Cargamos el tag
            Carga = mTag.Cargar(Control)
            If Carga Then
                Aux = Trim(Control.Text)
                If Aux <> "" Then
                    If mTag.Tabla <> "" Then
                        Tabla = mTag.Tabla & "."
                        Else
                        Tabla = ""
                    End If
                    Rc = SeparaCampoBusqueda(mTag.TipoDato, Tabla & mTag.Columna, Aux, Cad)
                    If Rc = 0 Then
                        If Sql <> "" Then Sql = Sql & " AND "
                        Sql = Sql & "(" & Cad & ")"
                    End If
                End If
            Else
                MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                Exit Function
            End If
        
        
        
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox Then
            mTag.Cargar Control
            If mTag.Cargado Then
                If Control.ListIndex > -1 Then
                    Cad = Control.ItemData(Control.ListIndex)
                    Cad = mTag.Tabla & "." & mTag.Columna & " = " & Cad
                    If Sql <> "" Then Sql = Sql & " AND "
                    Sql = Sql & "(" & Cad & ")"
                End If
            End If
        
        
        'CHECK
        ElseIf TypeOf Control Is CheckBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Control.Value = 1 Then
                        Cad = mTag.Tabla & "." & mTag.Columna & " = 1"
                        If Sql <> "" Then Sql = Sql & " AND "
                        Sql = Sql & "(" & Cad & ")"
                    End If
                End If
            End If
        End If

        
    Next Control
    ObtenerBusqueda = Sql
Exit Function
EObtenerBusqueda:
    ObtenerBusqueda = ""
    MuestraError Err.Number, "Obtener búsqueda. "
End Function




Public Function ModificaDesdeFormulario(ByRef Formulario As Form) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadWHERE As String
Dim cadUPDATE As String

On Error GoTo EModificaDesdeFormulario
    ModificaDesdeFormulario = False
    Set mTag = New CTag
    Aux = ""
    cadWHERE = ""
    For Each Control In Formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
              
                mTag.Cargar Control
                If mTag.Cargado Then
                    If mTag.Columna <> "" Then
                        'Sea para el where o para el update esto lo necesito
                        Aux = ValorParaSQL(Control.Text, mTag)
                        'Si es campo clave NO se puede modificar y se utiliza como busqueda
                        'dentro del WHERE
                        If mTag.EsClave Then
                            'Lo pondremos para el WHERE
                             If cadWHERE <> "" Then cadWHERE = cadWHERE & " AND "
                             cadWHERE = cadWHERE & "(" & mTag.Columna & " = " & Aux & ")"
                             
                        Else
                            If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                            cadUPDATE = cadUPDATE & "" & mTag.Columna & " = " & Aux
                        End If
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            'Partimos de la base que un booleano no es nunca clave primaria
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Control.Value = 1 Then
                    Aux = "TRUE"
                    Else
                    Aux = "FALSE"
                End If
                If mTag.TipoDato = "N" Then Aux = Abs(CBool(Aux))
                If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                'Esta es para access
                'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                cadUPDATE = cadUPDATE & "" & mTag.Columna & " = " & Aux
            End If
            
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Control.ListIndex = -1 Then
                        Aux = ValorNulo
                        Else
                        Aux = Control.ItemData(Control.ListIndex)
                    End If
                    If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                    'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                    cadUPDATE = cadUPDATE & "" & mTag.Columna & " = " & Aux
                End If
            End If
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo:
    'Update Pedidos
    'SET ImportePedido = ImportePedido * 1.1,
    'Cargo = Cargo * 1.03
    'WHERE PaísDestinatario = 'México';
    If cadWHERE = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
        Exit Function
    End If
    Aux = "UPDATE " & mTag.Tabla
    Aux = Aux & " SET " & cadUPDATE & " WHERE " & cadWHERE
    Conn.Execute Aux, , adCmdText






ModificaDesdeFormulario = True
Exit Function
EModificaDesdeFormulario:
    MuestraError Err.Number, "Modificar. " & Err.Description
End Function

Public Function ParaGrid(ByRef Control As Control, AnchoPorcentaje As Integer, Optional Desc As String) As String
Dim mTag As CTag
Dim Cad As String

'Montamos al final: "Cod Diag.|idDiag|N|10·"

ParaGrid = ""
Cad = ""
Set mTag = New CTag
mTag.Cargar Control
If mTag.Cargado Then
    If Control.Tag <> "" Then
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Desc <> "" Then
                Cad = Desc
            Else
                Cad = mTag.Nombre
            End If
            Cad = Cad & "|"
            Cad = Cad & mTag.Columna & "|"
            Cad = Cad & mTag.TipoDato & "|"
            Cad = Cad & AnchoPorcentaje & "·"
            
                
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            
        ElseIf TypeOf Control Is ComboBox Then
        
        
        End If 'De los elseif
    End If
Set mTag = Nothing
ParaGrid = Cad
End If



End Function

'////////////////////////////////////////////////////
' Monta a partir de una cadena devuelta por el formulario
'de busqueda el sql para situar despues el datasource
Public Function ValorDevueltoFormGrid(ByRef Control As Control, ByRef CadenaDevuelta As String, Orden As Integer) As String
Dim mTag As CTag
Dim Cad As String
Dim Aux As String
'Montamos al final: " columnatabla = valordevuelto "

ValorDevueltoFormGrid = ""
Cad = ""
Set mTag = New CTag
mTag.Cargar Control
If mTag.Cargado Then
    If Control.Tag <> "" Then
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            Aux = RecuperaValor(CadenaDevuelta, Orden)
            If Aux <> "" Then Cad = mTag.Columna & " = " & ValorParaSQL(Aux, mTag)
                
            
            
                
        'CheckBOX
       ' ElseIf TypeOf Control Is CheckBox Then
       '
       ' ElseIf TypeOf Control Is ComboBox Then
       '
       '
        End If 'De los elseif
    End If
End If
Set mTag = Nothing
ValorDevueltoFormGrid = Cad
End Function


Public Sub FormateaCampo(vTex As TextBox)
    Dim mTag As CTag
    Dim Cad As String
    On Error GoTo EFormateaCampo
    Set mTag = New CTag
    mTag.Cargar vTex
    If mTag.Cargado Then
        If vTex.Text <> "" Then
            If mTag.Formato <> "" Then
                Cad = TransformaPuntosComas(vTex.Text)
                Cad = Format(Cad, mTag.Formato)
                vTex.Text = Cad
            End If
        End If
    End If
EFormateaCampo:
    If Err.Number <> 0 Then Err.Clear
    Set mTag = Nothing
End Sub


'recupera valor desde una cadena con pipes(acabada en pipes)
'Para ello le decimos el orden  y ya ta
Public Function RecuperaValor(ByRef CADENA As String, Orden As Integer) As String
Dim I As Integer
Dim J As Integer
Dim Cont As Integer
Dim Cad As String

I = 0
Cont = 1
Cad = ""
Do
    J = I + 1
    I = InStr(J, CADENA, "|")
    If I > 0 Then
        If Cont = Orden Then
            Cad = Mid(CADENA, J, I - J)
            I = Len(CADENA) 'Para salir del bucle
            Else
                Cont = Cont + 1
        End If
    End If
Loop Until I = 0
RecuperaValor = Cad
End Function




'-----------------------------------------------------------------------
'Deshabilitar ciertas opciones del menu
'EN funcion del nivel de usuario
'Esto es a nivel general, cuando el Toolba es el mismo

'Para ello en el tag del button tendremos k poner un numero k nos diara hasta k nivel esta permitido

Public Sub PonerOpcionesMenuGeneral(ByRef Formulario As Form)
Dim I As Integer
Dim J As Integer


On Error GoTo EPonerOpcionesMenuGeneral


'Añadir, modificar y borrar deshabilitados si no nivel
With Formulario

    'LA TOOLBAR  .--> Requisito, k se llame toolbar1
    For I = 1 To .Toolbar1.Buttons.Count
        If .Toolbar1.Buttons(I).Tag <> "" Then
            J = Val(.Toolbar1.Buttons(I).Tag)
            If J < vUsu.Nivel Then
                .Toolbar1.Buttons(I).Enabled = False
            End If
        End If
    Next I
    
    'Esto es un poco salvaje. Por si acaso , no existe en este trozo pondremos los errores on resume next
    
    On Error Resume Next
    
    'Los MENUS
    'K sean mnAlgo
    J = Val(.mnNuevo.HelpContextID)
    If J < vUsu.Nivel Then .mnNuevo.Enabled = False
    
    J = Val(.mnModificar.HelpContextID)
    If J < vUsu.Nivel Then .mnModificar.Enabled = False
    
    J = Val(.mnEliminar.HelpContextID)
    If J < vUsu.Nivel Then .mnEliminar.Enabled = False
    On Error GoTo 0
End With




Exit Sub
EPonerOpcionesMenuGeneral:
    MuestraError Err.Number, "Poner opciones usuario generales"
End Sub



'Este modifica las claves prinipales y todo
'la sentenca del WHERE cod=1 and .. viene en claves
Public Function ModificaDesdeFormularioClaves(ByRef Formulario As Form, Claves As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadWHERE As String
Dim cadUPDATE As String
Dim I As Integer

On Error GoTo EModificaDesdeFormulario
    ModificaDesdeFormularioClaves = False
    Set mTag = New CTag
    Aux = ""
    cadWHERE = ""
    For Each Control In Formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    'Sea para el where o para el update esto lo necesito
                    Aux = ValorParaSQL(Control.Text, mTag)
                    'Si es campo clave NO se puede modificar y se utiliza como busqueda
                    If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                    cadUPDATE = cadUPDATE & "" & mTag.Columna & " = " & Aux
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            'Partimos de la base que un booleano no es nunca clave primaria
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Control.Value = 1 Then
                    Aux = "TRUE"
                    Else
                    Aux = "FALSE"
                End If
                If mTag.TipoDato = "N" Then Aux = Abs(CBool(Aux))
                If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                'Esta es para access
                'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                cadUPDATE = cadUPDATE & "" & mTag.Columna & " = " & Aux
            End If
            
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Control.ListIndex = -1 Then
                        Aux = ValorNulo
                        Else
                        Aux = Control.ItemData(Control.ListIndex)
                    End If
                    If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                    'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                    cadUPDATE = cadUPDATE & "" & mTag.Columna & " = " & Aux
                End If
            End If
        End If
    Next Control
    cadWHERE = Claves
    'Construimos el SQL
    If cadWHERE = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
        Exit Function
    End If
    Aux = "UPDATE " & mTag.Tabla
    Aux = Aux & " SET " & cadUPDATE & " WHERE " & cadWHERE
    Conn.Execute Aux, , adCmdText






ModificaDesdeFormularioClaves = True
Exit Function
EModificaDesdeFormulario:
    MuestraError Err.Number, "Modificar. " & Err.Description
End Function







'Public Function BLOQUEADesdeFormulario(ByRef Formulario As Form) As Boolean
'Dim Control As Object
'Dim mTag As CTag
'Dim Aux As String
'Dim cadWHERE As String
'Dim AntiguoCursor As Byte
'
'On Error GoTo EBLOQUEADesdeFormulario
'    BLOQUEADesdeFormulario = False
'    Set mTag = New CTag
'    Aux = ""
'    cadWHERE = ""
'    AntiguoCursor = Screen.MousePointer
'    Screen.MousePointer = vbHourglass
'    For Each Control In Formulario.Controls
'        'Si es texto monta esta parte de sql
'        If TypeOf Control Is TextBox Then
'            If Control.Tag <> "" Then
'
'                mTag.Cargar Control
'                If mTag.Cargado Then
'                    'Sea para el where o para el update esto lo necesito
'                    Aux = ValorParaSQL(Control.Text, mTag)
'                    'Si es campo clave NO se puede modificar y se utiliza como busqueda
'                    'dentro del WHERE
'                    If mTag.EsClave Then
'                        'Lo pondremos para el WHERE
'                         If cadWHERE <> "" Then cadWHERE = cadWHERE & " AND "
'                         cadWHERE = cadWHERE & "(" & mTag.Columna & " = " & Aux & ")"
'                    End If
'                End If
'            End If
'        End If
'    Next Control
'
'    If cadWHERE = "" Then
'        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
'
'    Else
'        Aux = "select * FROM " & mTag.Tabla
'        Aux = Aux & " WHERE " & cadWHERE & " FOR UPDATE"
'
'        'Intenteamos bloquear
'        PreparaBloquear
'        Conn.Execute Aux, , adCmdText
'        BLOQUEADesdeFormulario = True
'    End If
'EBLOQUEADesdeFormulario:
'    If Err.Number <> 0 Then
'        MuestraError Err.Number, "Bloqueo tabla"
'        TerminaBloquear
'    End If
'    Screen.MousePointer = AntiguoCursor
'End Function




'Public Function BloqueaRegistroForm(ByRef Formulario As Form) As Boolean
'Dim Control As Object
'Dim mTag As CTag
'Dim Aux As String
'Dim AuxDef As String
'Dim AntiguoCursor As Byte
'
'On Error GoTo EBLOQ
'    BloqueaRegistroForm = False
'    Set mTag = New CTag
'    Aux = ""
'    AuxDef = ""
'    AntiguoCursor = Screen.MousePointer
'    Screen.MousePointer = vbHourglass
'    For Each Control In Formulario.Controls
'        'Si es texto monta esta parte de sql
'        If TypeOf Control Is TextBox Then
'            If Control.Tag <> "" Then
'
'                mTag.Cargar Control
'                If mTag.Cargado Then
'                    'Si es campo clave NO se puede modificar y se utiliza como busqueda
'                    'dentro del WHERE
'                    If mTag.EsClave Then
'                        Aux = ValorParaSQL(Control.Text, mTag)
'                        AuxDef = AuxDef & Aux & "|"
'                    End If
'                End If
'            End If
'        End If
'    Next Control
'
'    If AuxDef = "" Then
'        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
'
'    Else
'        Aux = "Insert into zBloqueos(codusu,tabla,clave) VALUES(" & vUsu.Codigo & ",'" & mTag.Tabla
'        Aux = Aux & "',""" & AuxDef & """)"
'        Conn.Execute Aux
'        BloqueaRegistroForm = True
'    End If
'EBLOQ:
'    If Err.Number <> 0 Then
'        Aux = ""
'        If Conn.Errors.Count > 0 Then
'            If Conn.Errors(0).NativeError = 1062 Then
'                '¡Ya existe el registro, luego esta bloqueada
'                Aux = "BLOQUEO"
'            End If
'        End If
'        If Aux = "" Then
'            MuestraError Err.Number, "Bloqueo tabla"
'        Else
'            MsgBox "Registro bloqueado por otro usuario", vbExclamation
'        End If
'    End If
'    Screen.MousePointer = AntiguoCursor
'End Function
'
'
'Public Function DesBloqueaRegistroForm(ByRef TextBoxConTag As TextBox) As Boolean
'Dim mTag As CTag
'Dim SQL As String
'
''Solo me interesa la tabla
'On Error Resume Next
'    Set mTag = New CTag
'    mTag.Cargar TextBoxConTag
'    If mTag.Cargado Then
'        SQL = "DELETE from zBloqueos where codusu=" & vUsu.Codigo & " and tabla='" & mTag.Tabla & "'"
'        Conn.Execute SQL
'        If Err.Number <> 0 Then
'            Err.Clear
'        End If
'    End If
'    Set mTag = Nothing
'End Function



Public Sub KeyPress(ByRef KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub


