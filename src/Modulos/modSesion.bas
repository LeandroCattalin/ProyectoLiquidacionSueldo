Attribute VB_Name = "modSesion"
Option Explicit

Public UsuarioActual As clsUsuario

Public Sub CerrarSesion()
    Set UsuarioActual = Nothing
End Sub

Public Function RealizarLogin(ByVal sUser As String, ByVal sPass As String) As Boolean
    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    
    On Error GoTo ErrLogin
    
    Call AbrirBase
    
    Set cmd = New ADODB.Command
    With cmd
        .ActiveConnection = oConexion
        .CommandType = adCmdStoredProc
        .CommandText = "sp_ValidarLogin"
        
        .Parameters.Append .CreateParameter("@p_Username", adVarChar, adParamInput, 20, sUser)
        .Parameters.Append .CreateParameter("@p_Password", adVarChar, adParamInput, 20, sPass)
        
        Set rs = .Execute
    End With
    
    If Not rs.EOF Then
        Set UsuarioActual = New clsUsuario
        
        UsuarioActual.IdUsuario = rs!IdUsuario
        UsuarioActual.Username = rs!Username
        UsuarioActual.NombreCompleto = rs!NombreCompleto
        UsuarioActual.IdRol = rs!IdRol
        UsuarioActual.RolDescripcion = rs!RolNombre
        
        RealizarLogin = True
        Debug.Print "Login Realizado"
    Else
        RealizarLogin = False
        Debug.Print "Login fallido"
    End If

CleanUp:
    Set rs = Nothing
    Set cmd = Nothing
    Call CerrarBase
    Exit Function

ErrLogin:
    MsgBox "Error en el proceso de Login: " & Err.Description, vbCritical
    RealizarLogin = False
    Resume CleanUp
End Function
