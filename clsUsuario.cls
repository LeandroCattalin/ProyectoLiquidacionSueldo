VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Atributos
Private m_IdUsuario As Long
Private m_Username As String
Private m_NombreCompleto As String
Private m_IdRol As Long
Private m_RolDescripcion As String

'Propiedades
Public Property Let IdUsuario(ByVal vData As Long)
    m_IdUsuario = vData
End Property
Public Property Get IdUsuario() As Long
    IdUsuario = m_IdUsuario
End Property

Public Property Let Username(ByVal vData As String)
    m_Username = vData
End Property
Public Property Get Username() As String
    Username = m_Username
End Property

Public Property Let NombreCompleto(ByVal vData As String)
    m_NombreCompleto = vData
End Property
Public Property Get NombreCompleto() As String
    NombreCompleto = m_NombreCompleto
End Property

Public Property Let IdRol(ByVal vData As Long)
    m_IdRol = vData
End Property
Public Property Get IdRol() As Long
    IdRol = m_IdRol
End Property

Public Property Let RolDescripcion(ByVal vData As String)
    m_RolDescripcion = vData
End Property
Public Property Get RolDescripcion() As String
    RolDescripcion = m_RolDescripcion
End Property
