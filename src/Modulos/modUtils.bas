Attribute VB_Name = "modUtils"
Public Function ToDouble(ByVal valor As Variant) As Double
    If IsNumeric(valor) Then
        ToDouble = CDbl(valor)
    Else
        ToDouble = 0
    End If
End Function

Public Sub GenerarRecibosHTML(ByVal mes As Integer, ByVal anio As Integer)
    Dim rsCab As ADODB.Recordset
    Dim rsDet As ADODB.Recordset
    Dim fso As Object
    Dim html As String
    Dim pathCarpeta As String
    Dim pathArchivo As String
    Dim sql As String
    Dim count As Integer
    Dim strTipo As String ' Para guardar el texto "Haber" o "Descuento"
    
    On Error GoTo ErrRec

    Set fso = CreateObject("Scripting.FileSystemObject")
    pathCarpeta = App.Path & "\recibos"
    
    If Not fso.FolderExists(pathCarpeta) Then
        fso.CreateFolder (pathCarpeta)
    End If

    sql = "SELECT LC.*, E.nombre, E.apellido, E.legajo, E.cuil " & _
          "FROM LiquidacionCabecera LC " & _
          "INNER JOIN Empleado E ON LC.idEmpleado = E.idEmpleado " & _
          "WHERE LC.periodoMes = " & mes & " AND LC.periodoAnio = " & anio

    Set rsCab = modConnection.EjecutarConsulta(sql)

    If rsCab.EOF Then
        MsgBox "No hay liquidaciones para el periodo " & mes & "/" & anio, vbExclamation, "Sin datos"
        Exit Sub
    End If

    count = 0
    Do While Not rsCab.EOF
        ' --- HTML CON COLUMNA DE TIPO ---
        html = "<html><head><meta charset='UTF-8'><style>" & _
               "body { font-family: 'Segoe UI', sans-serif; margin: 30px; color: #333; }" & _
               ".recibo-box { border: 2px solid #2c3e50; padding: 20px; max-width: 850px; margin: auto; }" & _
               ".header { border-bottom: 2px solid #2c3e50; margin-bottom: 15px; padding-bottom: 10px; display: flex; justify-content: space-between; align-items: center; }" & _
               ".header h1 { margin: 0; color: #2c3e50; font-size: 22px; }" & _
               ".info-table { width: 100%; margin-bottom: 20px; font-size: 14px; border: 1px solid #eee; padding: 10px; background: #fcfcfc; }" & _
               ".details-table { width: 100%; border-collapse: collapse; }" & _
               ".details-table th { background: #2c3e50; color: white; padding: 10px; text-align: left; }" & _
               ".details-table td { border-bottom: 1px solid #eee; padding: 10px; font-size: 13px; }" & _
               ".col-haber { color: #27ae60; font-weight: bold; text-align: right; }" & _
               ".col-desc { color: #c0392b; font-weight: bold; text-align: right; }" & _
               ".total-row { font-weight: bold; background: #f2f2f2; }" & _
               ".neto-row { background: #2c3e50; color: white; font-size: 18px; font-weight: bold; }" & _
               ".firma { margin-top: 50px; text-align: center; border-top: 1px solid #000; width: 250px; margin-left: auto; margin-right: auto; padding-top: 5px; }" & _
               "</style></head><body>" & _
               "<div class='recibo-box'>" & _
               "<div class='header'><div><h1>SISTEMAS CATTALIN S.A.</h1>" & _
               "<p>Ruta Militar 123 - Campo de Mayo, BA | CUIT: 30-70809010-9</p></div>" & _
               "<div style='text-align:right'><b>ORIGINAL</b></div></div>" & _
               "<table class='info-table'>" & _
               "<tr><td><b>EMPLEADO:</b> " & rsCab!Apellido & ", " & rsCab!Nombre & "</td><td><b>LEGAJO:</b> " & rsCab!Legajo & "</td></tr>" & _
               "<tr><td><b>CUIL:</b> " & rsCab!Cuil & "</td><td><b>PERIODO:</b> " & mes & "/" & anio & "</td></tr>" & _
               "</table>" & _
               "<table class='details-table'><thead><tr>" & _
               "<th>Concepto</th><th>Tipo</th><th style='text-align:right;'>Haberes</th><th style='text-align:right;'>Descuentos</th></tr></thead><tbody>"

        sql = "SELECT LD.*, C.descripcion, C.tipo FROM LiquidacionDetalle LD " & _
              "INNER JOIN Conceptos C ON LD.idConceptos = C.idConceptos " & _
              "WHERE LD.idLiquidacion = " & rsCab!IdLiquidacion
        
        Set rsDet = modConnection.EjecutarConsulta(sql)
        
        Do While Not rsDet.EOF
            ' Lógica de discriminación por tipo según tu DB (H=Haber, R=Descuento/Retención)
            If UCase(Trim(rsDet!Tipo)) = "H" Then
                strTipo = "Haber"
                html = html & "<tr><td>" & rsDet!Descripcion & "</td><td>" & strTipo & "</td>" & _
                       "<td class='col-haber'>$ " & Format(rsDet!importe, "#,##0.00") & "</td><td></td></tr>"
            Else
                strTipo = "Descuento"
                html = html & "<tr><td>" & rsDet!Descripcion & "</td><td>" & strTipo & "</td>" & _
                       "<td></td><td class='col-desc'>$ " & Format(rsDet!importe, "#,##0.00") & "</td></tr>"
            End If
            rsDet.MoveNext
        Loop

        ' Totales finales
        html = html & "<tr class='total-row'><td>TOTALES</td><td></td>" & _
               "<td style='text-align:right;'>$ " & Format(rsCab!totalHaberes, "#,##0.00") & "</td>" & _
               "<td style='text-align:right;'>$ " & Format(rsCab!totalRetenciones, "#,##0.00") & "</td></tr>" & _
               "<tr class='neto-row'><td colspan='2'>NETO A COBRAR</td>" & _
               "<td colspan='2' style='text-align:right;'>$ " & Format(rsCab!netoCobrar, "#,##0.00") & "</td></tr>" & _
               "</tbody></table>" & _
               "<div class='firma'>Firma del Empleado</div></div></body></html>"

        pathArchivo = pathCarpeta & "\Recibo_" & rsCab!Legajo & "_" & mes & "_" & anio & ".html"
        
        Open pathArchivo For Output As #1
        Print #1, html
        Close #1
        
        count = count + 1
        rsCab.MoveNext
    Loop

    MsgBox "Se generaron " & count & " recibos con columnas discriminadas.", vbInformation
    Exit Sub

ErrRec:
    On Error Resume Next
    Close #1
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub

