VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmpleado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public IdEmpleado As Long
Public Apellido As String
Public Nombre As String
Public Cuil As String
Public Legajo As String
Public FechaIngreso As Date
Public SueldoBasico As Currency
Public Activo As Boolean

' Propiedad calculada (solo lectura)
Public Property Get NombreCompleto() As String
    NombreCompleto = Me.Apellido & ", " & Me.Nombre
End Property
