Attribute VB_Name = "modUI"
Public Sub LlenarComboEmpleados(ByRef combo As ComboBox)
    Dim rs As ADODB.Recordset
    
    combo.Clear
    Set rs = modEmpleados.GetEmpleadosActivos()
    
    If Not rs Is Nothing Then
        Do While Not rs.EOF
            combo.AddItem rs!NombreCompleto & " (" & rs!Legajo & ")"
            combo.ItemData(combo.NewIndex) = rs!IdEmpleado
            rs.MoveNext
        Loop
        rs.Close
    End If
End Sub

Public Sub LlenarComboEmpleadosConInactivos(ByRef combo As ComboBox)
    Dim rs As ADODB.Recordset
    
    combo.Clear
    Set rs = modEmpleados.GetEmpleados()
    
    If Not rs Is Nothing Then
        Do While Not rs.EOF
            combo.AddItem rs!NombreCompleto & " (" & rs!Legajo & ")"
            combo.ItemData(combo.NewIndex) = rs!IdEmpleado
            rs.MoveNext
        Loop
        rs.Close
    End If
End Sub

Public Sub ConfigurarColumnasListView(listview)
    With listview.ColumnHeaders
        .Clear
        ' El primer parámetro es el índice, el segundo la Key, el tercero el Texto y el cuarto el ancho
        .Add , , "Concepto", 3500
        .Add , , "Tipo", 800, lvwColumnCenter
        .Add , , "Monto/Valor", 1500, lvwColumnRight
    End With
End Sub

Public Sub LlenarListViewConceptos(ByRef lvw As listview, idEmp As Long)
    Dim colTodos As Collection
    Dim colAsignados As Collection
    Dim oCon As clsConcepto
    Dim itm As ListItem
    Dim oConAsignado As clsConcepto ' Variable auxiliar para la búsqueda
    
    Set colTodos = modConceptos.GetAllConceptosActivos()
    Set colAsignados = modConceptos.GetConceptosAsignados(idEmp)
    
    lvw.ListItems.Clear
    
    For Each oCon In colTodos
        ' Agregamos el ítem a la lista
        Set itm = lvw.ListItems.Add(, "C" & oCon.idConceptos, oCon.Descripcion)
        itm.SubItems(1) = oCon.Tipo
        itm.SubItems(2) = Format(oCon.MontoFijo + oCon.Porcentaje, "#,##0.00")
        
        ' IMPORTANTE: Resetear la variable de búsqueda
        Set oConAsignado = Nothing
        
        ' Intentamos buscar el concepto actual en la bolsa de "asignados"
        On Error Resume Next
        Set oConAsignado = colAsignados("C" & oCon.idConceptos)
        On Error GoTo 0 ' Siempre volvemos al manejo de errores normal rápido
        
        ' Si oConAsignado NO es Nothing, significa que lo encontró
        If Not oConAsignado Is Nothing Then
            itm.Checked = True
        Else
            itm.Checked = False
        End If
    Next
End Sub

Public Sub CargarCombosPeriodo(cboM As ComboBox, cboA As ComboBox)
    Dim i As Integer
    For i = 1 To 12
        cboM.AddItem Format(i, "00")
    Next i
    cboM.ListIndex = Month(Date) - 1
    
    cboA.AddItem "2025"
    cboA.AddItem "2026"
    cboA.AddItem "2027"
    cboA.ListIndex = 1
End Sub
