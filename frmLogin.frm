VERSION 5.00
Begin VB.Form frmLogin 
   Caption         =   "Login Sistema de Liquidacion"
   ClientHeight    =   6840
   ClientLeft      =   5235
   ClientTop       =   2625
   ClientWidth     =   12540
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   12540
   Begin VB.CheckBox cbxMostrarContrasena 
      Caption         =   "Check1"
      Height          =   195
      Left            =   3600
      TabIndex        =   6
      Top             =   4080
      Width           =   135
   End
   Begin VB.CommandButton btnIniciarSesion 
      Caption         =   "Iniciar sesion"
      Default         =   -1  'True
      Height          =   615
      Left            =   5280
      TabIndex        =   5
      Top             =   4440
      Width           =   1335
   End
   Begin VB.TextBox txtContrasena 
      ForeColor       =   &H00808080&
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   3600
      TabIndex        =   4
      Text            =   "Ingrese su contrasena"
      Top             =   3480
      Width           =   4455
   End
   Begin VB.TextBox txtUsuario 
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   3600
      TabIndex        =   2
      Text            =   "Ingrese su nombre de usuario"
      Top             =   2280
      Width           =   4455
   End
   Begin VB.Label lblCheckBoxMostrarContrasena 
      Caption         =   "Mostrar contrasena"
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Label lblContrasena 
      Alignment       =   2  'Center
      Caption         =   "Contrasena"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label lblUsuario 
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      Caption         =   "Login Sistema de Liquidacion BAS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   1320
      TabIndex        =   0
      Top             =   840
      Width           =   9735
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Inicio de sesion
Private Sub btnIniciarSesion_Click()
    'Validaciones de las textbox
    If Trim(txtUsuario.Text) = "Ingrese su nombre de usuario" Or Trim(txtContrasena.Text) = "Ingrese su contrasena" Then
        MsgBox "Por favor, complete todos los campos.", vbExclamation
    Else
        'Llamo al servicio
        If modSesion.RealizarLogin(txtUsuario.Text, txtContrasena.Text) Then
            'Paso al panel de control con sus credenciales
            MsgBox "Bienvenido/a " & UsuarioActual.NombreCompleto, vbInformation
            Dim f As New frmPanelDeControl
            f.Modo = UsuarioActual.IdRol
            f.Show
            Unload Me
        Else
            MsgBox "Usuario o contraseña incorrectos.", vbCritical
            txtContrasena.SetFocus
        End If
    End If
End Sub

'Gestion del CheckBox para mostrar la contrasena
Private Sub cbxMostrarContrasena_Click()
    If cbxMostrarContrasena.Value = 0 And txtContrasena.ForeColor = &H80000012 Then
        txtContrasena.PasswordChar = "*"
    ElseIf cbxMostrarContrasena.Value = 1 And txtContrasena.ForeColor = &H80000012 Then
        txtContrasena.PasswordChar = ""
    End If
End Sub

Private Sub lblCheckBoxMostrarContrasena_Click()
    If cbxMostrarContrasena = 0 Then
        cbxMostrarContrasena.Value = 1
    Else
        cbxMostrarContrasena.Value = 0
    End If
End Sub

' Limpiar TextBox en caso de que el usuario las clickee
Private Sub txtUsuario_GotFocus()
    If txtUsuario.ForeColor = &H808080 Then
        txtUsuario.Text = ""
        txtUsuario.ForeColor = &H80000012
    End If
End Sub
Private Sub txtContrasena_GotFocus()
    If txtContrasena.ForeColor = &H808080 Then
        If cbxMostrarContrasena = 0 Then
            txtContrasena.PasswordChar = "*"
        End If
        txtContrasena.ForeColor = &H80000012
        txtContrasena.Text = ""
    End If
End Sub
'Volver a poner el placeholder
Private Sub txtUsuario_LostFocus()
    If txtUsuario.Text = "" Then
        txtUsuario.ForeColor = &H808080
        txtUsuario.Text = "Ingrese su nombre de usuario"
    End If
End Sub
Private Sub txtContrasena_LostFocus()
    If txtContrasena.Text = "" Then
        txtContrasena.ForeColor = &H808080
        txtContrasena.PasswordChar = ""
        txtContrasena.Text = "Ingrese su contrasena"
    End If
End Sub
