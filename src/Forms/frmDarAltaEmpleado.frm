VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDarAltaEmpleado 
   Caption         =   "Form1"
   ClientHeight    =   5835
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9930
   LinkTopic       =   "Form1"
   ScaleHeight     =   5835
   ScaleWidth      =   9930
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnDarDeAlta 
      Caption         =   "Dar de alta"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5160
      TabIndex        =   12
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton btnCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      TabIndex        =   11
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox txtNombre 
      Height          =   285
      Left            =   5160
      TabIndex        =   4
      Top             =   1680
      Width           =   4695
   End
   Begin VB.TextBox txtApellido 
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   4815
   End
   Begin VB.TextBox txtCuil 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   2655
   End
   Begin VB.TextBox txtLegajo 
      Height          =   285
      Left            =   6120
      TabIndex        =   1
      Top             =   2280
      Width           =   2655
   End
   Begin VB.TextBox txtSueldoBasico 
      Height          =   285
      Left            =   3240
      TabIndex        =   0
      Top             =   3000
      Width           =   2655
   End
   Begin MSComCtl2.DTPicker dtpFechaDeIngreso 
      Height          =   375
      Left            =   3240
      TabIndex        =   14
      Top             =   2280
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      Format          =   115933185
      CurrentDate     =   36526
   End
   Begin VB.Label lblDarDeAlta 
      Alignment       =   2  'Center
      Caption         =   "Dar alta de empleado"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   13
      Top             =   480
      Width           =   5415
   End
   Begin VB.Label lblNombre 
      Caption         =   "Nombre"
      Height          =   255
      Left            =   5160
      TabIndex        =   10
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lblApellido 
      Caption         =   "Apellido"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lblCuil 
      Caption         =   "Cuil"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblFechaIngreso 
      Caption         =   "Fecha de ingreso"
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblLegajo 
      Caption         =   "Legajo"
      Height          =   255
      Left            =   6120
      TabIndex        =   6
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label lblSueldoBasico 
      Caption         =   "Sueldo basico"
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   2760
      Width           =   1695
   End
End
Attribute VB_Name = "frmDarAltaEmpleado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancelar_Click()
    Unload Me
End Sub

Private Sub btnDarDeAlta_Click()
    On Error GoTo ErrHandler
    Dim oNuevoEmp As New clsEmpleado
    
    ' 1. Validaciones mínimas
    If txtApellido.Text = "" Or txtNombre.Text = "" Or txtLegajo.Text = "" Then
        MsgBox "Por favor, complete los campos obligatorios (Apellido, Nombre y Legajo).", vbExclamation
        Exit Sub
    End If

    ' 2. Mapeamos los datos de la interfaz al objeto
    With oNuevoEmp
        .Apellido = txtApellido.Text
        .Nombre = txtNombre.Text
        .Cuil = txtCuil.Text
        .Legajo = txtLegajo.Text
        .FechaIngreso = dtpFechaDeIngreso.Value
        .SueldoBasico = CDbl(IIf(txtSueldoBasico.Text = "", 0, txtSueldoBasico.Text))
    End With

    ' 3. Llamamos al servicio de datos
    If modEmpleados.InsertarEmpleado(oNuevoEmp) Then
        MsgBox "Empleado dado de alta con éxito.", vbInformation
        Unload Me
    End If
ErrHandler:
    MsgBox "Ocurrió un error inesperado: " & Err.Description, vbCritical, "Error de Sistema"
End Sub
End Sub
