Attribute VB_Name = "modConfig"
Option Explicit

' API de Windows para leer archivos INI
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

' Función auxiliar para simplificar la lectura
Public Function LeerINI(Seccion As String, Clave As String, Default As String) As String
    Dim sRetVal As String * 255
    Dim iLen As Long
    Dim sFilePath As String
    
    ' Ruta completa al archivo .ini (en la misma carpeta que el ejecutable)
    sFilePath = App.Path & "\config.ini"
    iLen = GetPrivateProfileString(Seccion, Clave, Default, sRetVal, Len(sRetVal), sFilePath)
    LeerINI = Left(sRetVal, iLen)
End Function

