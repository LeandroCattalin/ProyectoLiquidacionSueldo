Attribute VB_Name = "modEmpleados"
Public Function GetEmpleadoById(id As Long) As clsEmpleado
    Dim rs As ADODB.Recordset
    Dim oEmp As New clsEmpleado
    
    Set rs = modConnection.EjecutarConsulta("SELECT * FROM Empleado WHERE IdEmpleado = " & id)
    
    If Not rs.EOF Then
        With oEmp
            .IdEmpleado = rs!IdEmpleado
            .Apellido = rs!Apellido & ""
            .Nombre = rs!Nombre & ""
            .Cuil = rs!Cuil & ""
            .Legajo = rs!Legajo & ""
            .FechaIngreso = IIf(IsNull(rs!FechaIngreso), Now, rs!FechaIngreso)
            .SueldoBasico = IIf(IsNull(rs!SueldoBasico), 0, CDbl(rs!SueldoBasico))
            .Activo = rs!Activo & ""
        End With
        Set GetEmpleadoById = oEmp
    Else
        Set GetEmpleadoById = Nothing
    End If
End Function

Public Function GuardarEmpleado(oEmp As clsEmpleado) As Boolean
    On Error GoTo ErrGuardar
    Dim sql As String
    sql = "UPDATE Empleado SET " & _
        "Apellido = '" & Replace(oEmp.Apellido, "'", "''") & "', " & _
        "Nombre = '" & Replace(oEmp.Nombre, "'", "''") & "', " & _
        "Cuil = '" & oEmp.Cuil & "', " & _
        "Legajo = '" & oEmp.Legajo & "', " & _
        "SueldoBasico = " & Str(oEmp.SueldoBasico) & ", " & _
        "Activo = " & IIf(oEmp.Activo, 1, 0) & " " & _
        "WHERE idEmpleado = " & oEmp.IdEmpleado

    modConnection.EjecutarConsulta sql
    GuardarEmpleado = True
    Exit Function

ErrGuardar:
    MsgBox "Error al guardar en la base de datos: " & Err.Description, vbCritical
    GuardarEmpleado = False
End Function

Public Function InsertarEmpleado(oEmp As clsEmpleado) As Boolean
    On Error GoTo ErrInsert
    Dim sql As String
    
    sql = "INSERT INTO Empleado (Apellido, Nombre, Cuil, Legajo, FechaIngreso, SueldoBasico, Activo) " & _
        "VALUES (" & _
        "'" & Replace(oEmp.Apellido, "'", "''") & "', " & _
        "'" & Replace(oEmp.Nombre, "'", "''") & "', " & _
        "'" & oEmp.Cuil & "', " & _
        "'" & oEmp.Legajo & "', " & _
        "'" & Format(oEmp.FechaIngreso, "yyyymmdd") & "', " & _
        "" & Str(oEmp.SueldoBasico) & ", 1)"

    modConnection.EjecutarConsulta sql
    InsertarEmpleado = True
    Exit Function

ErrInsert:
    MsgBox "Error al crear el empleado: " & Err.Description, vbCritical
    InsertarEmpleado = False
End Function

Public Function GetEmpleadosActivos() As ADODB.Recordset
    Dim sql As String
    sql = "SELECT IdEmpleado, NombreCompleto, legajo FROM Empleado WHERE Activo = 1 ORDER BY NombreCompleto"
    Set GetEmpleadosActivos = modConnection.EjecutarConsulta(sql)
End Function

Public Function GetEmpleados() As ADODB.Recordset
    Dim sql As String
    sql = "SELECT IdEmpleado, NombreCompleto, legajo FROM Empleado ORDER BY NombreCompleto"
    Set GetEmpleados = modConnection.EjecutarConsulta(sql)
End Function

Public Function GetAllEmpleadosActivos() As Collection
    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim col As New Collection
    Dim oEmp As clsEmpleado
    
    ' Traemos todos los campos necesarios para liquidar
    sql = "SELECT idEmpleado, nombre, apellido, legajo, sueldoBasico FROM Empleado WHERE Activo = 1"
    
    Set rs = modConnection.EjecutarConsulta(sql)
    
    Do While Not rs.EOF
        Set oEmp = New clsEmpleado
        oEmp.IdEmpleado = rs!IdEmpleado
        oEmp.Nombre = rs!Nombre
        oEmp.Apellido = rs!Apellido
        oEmp.Legajo = rs!Legajo
        oEmp.SueldoBasico = rs!SueldoBasico
        
        ' Agregamos el objeto a la colección usando el ID como Key
        col.Add oEmp, "E" & oEmp.IdEmpleado
        
        rs.MoveNext
    Loop
    
    Set GetAllEmpleadosActivos = col
End Function

