Attribute VB_Name = "modConnection"
Option Explicit
Public oConexion As ADODB.Connection

Public Sub AbrirBase()
    'Variabales
    Dim strConnection As String
    Dim sServidor As String
    Dim sDB As String
    Dim sUser As String
    Dim sPass As String
    Dim strConn As String
    
    'Manejo de errores
    On Error GoTo ErrConex
    
    'Leo el config.ini
    sServidor = LeerINI("DATABASE", "Server", ".")
    sDB = LeerINI("DATABASE", "DB", "")
    sUser = LeerINI("DATABASE", "User", "")
    sPass = LeerINI("DATABASE", "Pass", "")
    
    'Creo la instancia de conexion
    Set oConexion = New ADODB.Connection
    strConnection = "Provider=MSOLEDBSQL;Data Source=" & sServidor & _
    ";Initial Catalog=" & sDB & _
    ";User ID=" & sUser & _
    ";Password=" & sPass & ";Integrated Security=SSPI;"
    oConexion.ConnectionString = strConnection
    oConexion.Open
    Exit Sub

ErrConex:
    MsgBox "Error al conectar: " & Err.Description, vbCritical
End Sub

Public Sub CerrarBase()
    On Error Resume Next
    
    If Not oConexion Is Nothing Then
        If oConexion.State = adStateOpen Then
            oConexion.Close
        End If
        Set oConexion = Nothing
    End If
    
    Debug.Print "Conexión cerrada correctamente."
End Sub

Public Function EjecutarConsulta(ByVal sql As String) As ADODB.Recordset
    On Error GoTo ErrHandler
    Dim rs As New ADODB.Recordset
    
    Call AbrirBase
    
    rs.CursorLocation = adUseClient
    rs.Open sql, oConexion, adOpenForwardOnly, adLockReadOnly
    
    Set rs.ActiveConnection = Nothing
    Set EjecutarConsulta = rs
    
    Exit Function

ErrHandler:
    MsgBox "Error al ejecutar consulta: " & Err.Description, vbCritical
    Set EjecutarConsulta = Nothing
End Function
