VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Configuracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'#Generador DAVID(Ariadna)



'-------------------------------------------------------
'-------------------------------------------------------
'Declaracion de las variables
'---- Carpeta donde se encuentran los archivos
Private mvarPathArchivos As String 'Variable local
'---- Donde se llevaran los errores
Private mvarPathError As String 'Variable local
'---- Codigo extension ARidoc del tipo de integracion
Private mvarExtensionArchivos As Integer 'Variable local
'---- Extension fichero con los datos
Private mvarextensionGuia As String 'Variable local
'---- Ruta del aridoc
Private mvarPathAridoc As String 'Variable local


'Variables de este modulo
Dim NF As Integer
Dim Mens As String


'------------------------------------------------------
'------------------------------------------------------

'      Propiedades del modulo de configuracion
'------------------------------------------------------
'------------------------------------------------------

'----
'---- Carpeta donde se encuentran los archivos
'----
Public Property Let PathArchivos(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.PathArchivos = 5
     mvarPathArchivos = vData
End Property

Public Property Get PathArchivos() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.PathArchivos
     PathArchivos = mvarPathArchivos
End Property




'----
'---- Donde se llevaran los errores
'----
Public Property Let PathError(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.PathError = 5
     mvarPathError = vData
End Property

Public Property Get PathError() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.PathError
     PathError = mvarPathError
End Property




'----
'---- Codigo extension ARidoc del tipo de integracion
'----
Public Property Let ExtensionArchivos(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.ExtensionArchivos = 5
     mvarExtensionArchivos = vData
End Property

Public Property Get ExtensionArchivos() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.ExtensionArchivos
     ExtensionArchivos = mvarExtensionArchivos
End Property




'----
'---- Extension fichero con los datos
'----
Public Property Let extensionGuia(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.extensionGuia = 5
     mvarextensionGuia = vData
End Property

Public Property Get extensionGuia() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.extensionGuia
     extensionGuia = mvarextensionGuia
End Property




'----
'---- Ruta del aridoc
'----
Public Property Let PathAridoc(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.PathAridoc = 5
     mvarPathAridoc = vData
End Property

Public Property Get PathAridoc() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.PathAridoc
     PathAridoc = mvarPathAridoc
End Property







'//////////////////////////////////////////////////////
'             LEER  Y  GUARDAR
'//////////////////////////////////////////////////////
'
'
'
'
'
'
'------------------------------------------------------
'     LEER
'
Public Function Leer() As Byte
    '--------------------------------------------------
    '   Leer:
    '   Busca un fichero de nombre Config.ini y lee de él los valores
    '   de parametrización para la aplicación. El fichero se busca
    '   en el directorio desde el que se ejecuta la aplicación.
    '--------------------------------------------------
    Dim Fichero As String ' Nombre completo del fichero a leer
    Dim LinAnt As String
    Dim Lin As String
    Dim NomParam As String ' Para leer el nombre del parámetro
    Dim ValParam As String ' Para leer el valor del parámetro
    Dim NFile As Integer
    
    On Error GoTo Err1
    Fichero = App.Path & "\ConfigIntA.ini"
    NFile = FreeFile
    Open Fichero For Input As #NFile   ' Abre el archivo.
    Do While Not EOF(NFile) ' Repite el bucle hasta el final del archivo.
        Line Input #NFile, Lin
        Select Case Lin
           Case "[PathArchivos]", "[PathError]", "[ExtensionArchivos]" _
               , "[extensionGuia]", "[PathAridoc]"
               ValParam = ""
               NomParam = Lin

           Case "<-------->"
               Select Case NomParam
                    Case "[PathArchivos]": mvarPathArchivos = ValParam
                    Case "[PathError]": mvarPathError = ValParam
                    Case "[ExtensionArchivos]": mvarExtensionArchivos = ValParam
                    Case "[extensionGuia]": mvarextensionGuia = ValParam
                    Case "[PathAridoc]": mvarPathAridoc = ValParam
                End Select
            Case Else
                If ValParam = "" Then
                    ValParam = Lin
                Else
                    ValParam = ValParam & vbCrLf & Lin
                End If
        End Select
    Loop
    Close #NFile    ' Cierra el archivo.
    Leer = 0
    Exit Function
Err1:
   Leer = 1
End Function




'------------------------------------------------------
'     GUARDAR

Public Function Grabar() As Byte
    '--------------------------------------------------
    '   GuardaConfig:
    '   Guardamos en el fichero Config.ini los valores de los
    '   parámetros de la aplicación
    '--------------------------------------------------
    Dim Fichero As String ' Nombre completo del fichero a leer
    Dim NFile As Integer

    On Error GoTo Err1
    Fichero = App.Path & "\ConfigIntA.ini"
    NFile = FreeFile
    Open Fichero For Output As #NFile   ' Abre el archivo.
    Print #NFile, "[PathArchivos]"
    Print #NFile, mvarPathArchivos
    Print #NFile, "<-------->"
    Print #NFile, "[PathError]"
    Print #NFile, mvarPathError
    Print #NFile, "<-------->"
    Print #NFile, "[ExtensionArchivos]"
    Print #NFile, mvarExtensionArchivos
    Print #NFile, "<-------->"
    Print #NFile, "[extensionGuia]"
    Print #NFile, mvarextensionGuia
    Print #NFile, "<-------->"
    Print #NFile, "[PathAridoc]"
    Print #NFile, mvarPathAridoc
    Print #NFile, "<-------->"


    Close #NFile
    Grabar = 0
    Exit Function
Err1:
    Mens = "Error " & Err.Number & " : " & Err.Description
    MsgBox Mens, vbCritical
    Grabar = 1
End Function
