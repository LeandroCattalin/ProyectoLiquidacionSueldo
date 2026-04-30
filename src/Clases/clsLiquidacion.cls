VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLiquidacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public IdLiquidacion As Long
Public IdEmpleado As Long
Public periodo As String
Public SueldoBasico As Double
Public Detalles As Collection

Private Sub Class_Initialize()
    Set Detalles = New Collection
End Sub

Public Function CalcularHaberes() As Double
    Dim oDet As clsLiquidacionDetalle
    Dim total As Double
    For Each oDet In Me.Detalles
        If oDet.Tipo = "H" Then
            total = total + oDet.Monto
        End If
    Next
    CalcularHaberes = total
End Function

Public Function CalcularRetenciones() As Double
    Dim oDet As clsLiquidacionDetalle
    Dim total As Double
    For Each oDet In Me.Detalles
        If oDet.Tipo = "R" Then
            total = total + oDet.Monto
        End If
    Next
    CalcularRetenciones = total
End Function

Public Function CalcularNeto() As Double
    CalcularNeto = CalcularHaberes() - CalcularRetenciones()
End Function
