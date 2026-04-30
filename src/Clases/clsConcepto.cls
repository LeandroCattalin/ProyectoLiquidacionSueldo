VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConcepto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_idConceptos As Long
Private m_descripcion As String
Private m_montoFijo As Double
Private m_porcentaje As Double
Private m_tipo As String ' "H" para Haber, "R" para Retencion
Private m_activo As Boolean

Public Property Get idConceptos() As Long
    idConceptos = m_idConceptos
End Property

Public Property Let idConceptos(ByVal vNewValue As Long)
    m_idConceptos = vNewValue
End Property

Public Property Get Descripcion() As String
    Descripcion = m_descripcion
End Property

Public Property Let Descripcion(ByVal vNewValue As String)
    m_descripcion = vNewValue
End Property

Public Property Get MontoFijo() As Double
    MontoFijo = m_montoFijo
End Property

Public Property Let MontoFijo(ByVal vNewValue As Double)
    m_montoFijo = vNewValue
End Property

Public Property Get Porcentaje() As Double
    Porcentaje = m_porcentaje
End Property

Public Property Let Porcentaje(ByVal vNewValue As Double)
    m_porcentaje = vNewValue
End Property

Public Property Get Tipo() As String
    Tipo = m_tipo
End Property

Public Property Let Tipo(ByVal vNewValue As String)
    If vNewValue = "H" Or vNewValue = "R" Then
        m_tipo = vNewValue
    Else
        Debug.Print "Error: Tipo de concepto inválido (" & vNewValue & ")"
    End If
End Property

Public Property Get Activo() As Boolean
    Activo = m_activo
End Property

Public Property Let Activo(ByVal vNewValue As Boolean)
    m_activo = vNewValue
End Property

'Constructor
Private Sub Class_Initialize()
    m_idConceptos = 0
    m_montoFijo = 0
    m_porcentaje = 0
    m_activo = True
End Sub
