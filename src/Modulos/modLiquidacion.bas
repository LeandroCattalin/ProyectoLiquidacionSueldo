Attribute VB_Name = "modLiquidacion"
Option Explicit

' --- MOTOR DE LIQUIDACIÓN ---
Public Sub ProcesarLiquidacion(ByVal mes As Integer, ByVal anio As Integer)
    Dim colEmpleados As Collection
    Dim oEmp As clsEmpleado
    Dim oCon As clsConcepto
    Dim oLiq As clsLiquidacion
    Dim oDet As clsLiquidacionDetalle
    Dim colConceptosEmp As Collection

    On Error GoTo ErrProceso

    ' Traemos los empleados usando tu función de modEmpleados
    Set colEmpleados = modEmpleados.GetAllEmpleadosActivos()

    For Each oEmp In colEmpleados
        ' 1. Instanciamos la cabecera
        Set oLiq = New clsLiquidacion
        oLiq.IdEmpleado = oEmp.IdEmpleado
        oLiq.periodo = mes & "-" & anio
        oLiq.SueldoBasico = oEmp.SueldoBasico
        
        ' 2. Buscamos conceptos asignados para este empleado
        ' Asegurate que esta función exista en tu modConceptos
        Set colConceptosEmp = modConceptos.GetConceptosAsignados(oEmp.IdEmpleado)
        
        For Each oCon In colConceptosEmp
            Set oDet = New clsLiquidacionDetalle
            oDet.IdConcepto = oCon.idConceptos
            oDet.Descripcion = oCon.Descripcion
            oDet.Tipo = oCon.Tipo
            oDet.Monto = CalcularMontoConcepto(oEmp.SueldoBasico, oCon)
            Debug.Print "EMP: " & oEmp.IdEmpleado & " | BASICO: " & oEmp.SueldoBasico & _
            " | CONCEPTO: " & oCon.Descripcion & " | PORC: " & oCon.Porcentaje & _
            " | MONTO: " & oDet.Monto
            oLiq.Detalles.Add oDet
        Next
        
        ' 3. Persistimos pasando mes y año por separado
        Debug.Print "Procesando Empleado: " & oLiq.IdEmpleado & " - Conceptos: " & oLiq.Detalles.count
        Call GuardarLiquidacionCompleta(oLiq, mes, anio)
    Next

    MsgBox "Proceso de liquidación finalizado.", vbInformation
    Exit Sub

ErrProceso:
    MsgBox "Error en ProcesarLiquidacion: " & Err.Description, vbCritical
End Sub

' --- PERSISTENCIA TRANSACCIONAL ---
Public Function GuardarLiquidacionCompleta(oLiq As clsLiquidacion, ByVal mes As Integer, ByVal anio As Integer) As Boolean
    On Error GoTo ErrGuardar
    Dim sql As String
    Dim oDet As clsLiquidacionDetalle
    
    ' Obtenemos los totales desde la clase
    Dim neto As Double: neto = oLiq.CalcularNeto()
    Dim haberes As Double: haberes = oLiq.CalcularHaberes()
    Dim retenciones As Double: retenciones = oLiq.CalcularRetenciones()
    
    ' Iniciamos el script SQL
    sql = "BEGIN TRANSACTION; " & _
          "DECLARE @NewID INT; "

    ' Inserción en LiquidacionCabecera
    sql = sql & "INSERT INTO LiquidacionCabecera (idEmpleado, periodoAnio, periodoMes, netoCobrar, totalHaberes, totalRetenciones, Firmado) " & _
          "VALUES (" & oLiq.IdEmpleado & ", " & anio & ", " & mes & ", " & _
          FmtSQL(neto) & ", " & FmtSQL(haberes) & ", " & FmtSQL(retenciones) & ", 0); " & _
          "SET @NewID = SCOPE_IDENTITY(); "

    ' Inserción en LiquidacionDetalle
    For Each oDet In oLiq.Detalles
        If Not oDet Is Nothing Then
            sql = sql & "INSERT INTO LiquidacionDetalle (idLiquidacion, importe, idConceptos) " & _
                  "VALUES (@NewID, " & FmtSQL(oDet.Monto) & ", " & oDet.IdConcepto & "); "
        End If
    Next

    sql = sql & "COMMIT TRANSACTION;"

    ' Ejecución única a través de tu módulo de conexión
    modConnection.EjecutarConsulta sql
    
    GuardarLiquidacionCompleta = True
    Exit Function

ErrGuardar:
    Debug.Print "Error al guardar empleado " & oLiq.IdEmpleado & ": " & Err.Description
    GuardarLiquidacionCompleta = False
End Function

' --- FUNCIONES AUXILIARES ---

Private Function CalcularMontoConcepto(ByVal basico As Double, ByVal oCon As clsConcepto) As Double
    ' Lógica de cálculo según porcentaje o monto fijo
    If oCon.Porcentaje > 0 Then
        CalcularMontoConcepto = (basico * oCon.Porcentaje) / 100
    Else
        CalcularMontoConcepto = oCon.MontoFijo
    End If
End Function

Private Function FmtSQL(ByVal valor As Double) As String
    ' Convierte dobles a string con punto decimal para SQL
    FmtSQL = Replace(CStr(valor), ",", ".")
End Function

