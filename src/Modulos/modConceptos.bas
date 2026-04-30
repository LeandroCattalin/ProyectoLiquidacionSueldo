Attribute VB_Name = "modConceptos"
Option Explicit

' Guarda o actualiza un concepto (Alta/Modificación/Baja)
Public Function GuardarConcepto(oCon As clsConcepto) As Boolean
    Dim sql As String
    Dim resultado As Boolean
    
    ' Usamos Str() para que los decimales vayan con punto a SQL Server (formato universal)
    If oCon.idConceptos = 0 Then
        ' ALTA
        sql = "INSERT INTO Conceptos (descripcion, montoFijo, porcentaje, tipo, activo) VALUES (" & _
              "'" & Replace(oCon.Descripcion, "'", "''") & "', " & _
              Str(oCon.MontoFijo) & ", " & _
              Str(oCon.Porcentaje) & ", " & _
              "'" & oCon.Tipo & "', 1)"
    Else
        ' MODIFICACION Y BAJA LOGICA
        sql = "UPDATE Conceptos SET " & _
              "descripcion = '" & Replace(oCon.Descripcion, "'", "''") & "', " & _
              "montoFijo = " & Str(oCon.MontoFijo) & ", " & _
              "porcentaje = " & Str(oCon.Porcentaje) & ", " & _
              "tipo = '" & oCon.Tipo & "', " & _
              "activo = " & IIf(oCon.Activo, 1, 0) & " " & _
              "WHERE idConceptos = " & oCon.idConceptos
    End If
    
    'El problema es que EjecutarConsulta devuelve un RS, al ser un INSERT, no hay RS que devolver, solo devuelve la cantidad de filas
    'afectadas, habria que hacer una funcion de ejecutar y otra de obtener
    modConnection.EjecutarConsulta sql
    
    GuardarConcepto = True ' Si llegó acá, es que en la DB impactó
    Exit Function

ErrHandler:
    ' Debug.Print "Error: " & Err.Description ' Solo para vos
    GuardarConcepto = False
End Function

' Trae un concepto específico por ID
Public Function GetConceptoById(id As Long) As clsConcepto
    Dim rs As ADODB.Recordset
    Dim oCon As clsConcepto
    
    Set rs = modConnection.EjecutarConsulta("SELECT * FROM Conceptos WHERE idConceptos = " & id)
    
    If Not rs.EOF Then
        Set oCon = New clsConcepto
        With oCon
            .idConceptos = rs!idConceptos
            .Descripcion = rs!Descripcion & ""
            .MontoFijo = IIf(IsNull(rs!MontoFijo), 0, CDbl(rs!MontoFijo))
            .Porcentaje = IIf(IsNull(rs!Porcentaje), 0, CDbl(rs!Porcentaje))
            .Tipo = rs!Tipo & ""
            .Activo = CBool(rs!Activo)
        End With
        Set GetConceptoById = oCon
    Else
        Set GetConceptoById = Nothing
    End If
    
    rs.Close
    Set rs = Nothing
End Function

' Trae los conceptos asociados a un empleado específico
Public Function GetConceptosAsignados(ByVal IdEmpleado As Long) As Collection
    Dim rs As ADODB.Recordset
    Dim col As New Collection
    Dim oCon As clsConcepto
    Dim sql As String
    
    ' Usamos C.* porque ya filtramos por el ID del empleado en el JOIN
    sql = "SELECT C.* FROM Conceptos C " & _
          "INNER JOIN Empleado_Concepto EC ON C.idConceptos = EC.idConcepto " & _
          "WHERE EC.idEmpleado = " & IdEmpleado & " AND C.activo = 1"
          
    Set rs = modConnection.EjecutarConsulta(sql)
    
    Do While Not rs.EOF
        Set oCon = New clsConcepto
        
        ' Mapeo de propiedades indispensables para el cálculo
        oCon.idConceptos = rs!idConceptos
        oCon.Descripcion = rs!Descripcion & ""
        oCon.Tipo = rs!Tipo & ""
        
        ' --- LAS QUE FALTABAN ---
        ' Usamos IIf e IsNull para evitar que el programa explote si hay un nulo en la DB
        oCon.Porcentaje = IIf(IsNull(rs!Porcentaje), 0, rs!Porcentaje)
        oCon.MontoFijo = IIf(IsNull(rs!MontoFijo), 0, rs!MontoFijo)
        ' ------------------------
        
        ' Agregamos a la colección usando el ID como Key para evitar duplicados
        col.Add oCon, "C" & oCon.idConceptos
        
        rs.MoveNext
    Loop
    
    Set GetConceptosAsignados = col
End Function

' Persiste la asignación masiva (Borrón y cuenta nueva)
Public Function ActualizarAsignacion(ByVal IdEmpleado As Long, ByRef colConceptos As Collection) As Boolean
    On Error GoTo ErrAsignacion
    Dim oCon As clsConcepto
    Dim sql As String
    
    ' 1. Limpiamos lo anterior
    modConnection.EjecutarConsulta "DELETE FROM Empleado_Concepto WHERE idEmpleado = " & IdEmpleado
    
    ' 2. Insertamos lo nuevo
    For Each oCon In colConceptos
        sql = "INSERT INTO Empleado_Concepto (idEmpleado, idConcepto, activo) " & _
              "VALUES (" & IdEmpleado & ", " & oCon.idConceptos & ", 1)"
        modConnection.EjecutarConsulta sql
    Next
    
    ActualizarAsignacion = True
    Exit Function

ErrAsignacion:
    Debug.Print "Error en ActualizarAsignacion: " & Err.Description
    ActualizarAsignacion = False
End Function

Public Function GetAllConceptosActivos() As Collection
    Dim rs As ADODB.Recordset
    Dim col As New Collection
    Dim oCon As clsConcepto
    Dim sql As String
    
    ' Traemos solo los que están operativos
    sql = "SELECT idConceptos, descripcion, montoFijo, porcentaje, tipo, activo " & _
          "FROM Conceptos WHERE activo = 1 ORDER BY descripcion ASC"
          
    Set rs = modConnection.EjecutarConsulta(sql)
    
    Do While Not rs.EOF
        Set oCon = New clsConcepto
        
        With oCon
            .idConceptos = rs!idConceptos
            .Descripcion = rs!Descripcion & ""
            ' Usamos CDbl e IsNull para evitar el error de "Invalid use of Null"
            .MontoFijo = IIf(IsNull(rs!MontoFijo), 0, CDbl(rs!MontoFijo))
            .Porcentaje = IIf(IsNull(rs!Porcentaje), 0, CDbl(rs!Porcentaje))
            .Tipo = rs!Tipo & ""
            .Activo = CBool(rs!Activo)
        End With
        
        ' La clave de la colección es la "C" + ID para que sea única y fácil de buscar
        col.Add oCon, "C" & oCon.idConceptos
        
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    
    Set GetAllConceptosActivos = col
End Function

