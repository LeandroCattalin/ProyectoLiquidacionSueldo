VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDarBajaEmpleado 
   Caption         =   "Form1"
   ClientHeight    =   4665
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   9885
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox cbxDarDeBaja 
      Caption         =   "Check1"
      Height          =   195
      Left            =   5520
      TabIndex        =   6
      Top             =   2160
      Width           =   255
   End
   Begin MSComCtl2.DTPicker dtpFechaDeBaja 
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   2040
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   115933185
      CurrentDate     =   46134
   End
   Begin VB.CommandButton btnDarDeBaja 
      Caption         =   "Grabar cambios"
      Height          =   735
      Left            =   5040
      TabIndex        =   4
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton btnSalir 
      Caption         =   "Salir"
      Height          =   735
      Left            =   2400
      TabIndex        =   3
      Top             =   3000
      Width           =   1335
   End
   Begin VB.ComboBox cboEmpleados 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   9615
   End
   Begin VB.Label lblDarDeBaja 
      Alignment       =   2  'Center
      Caption         =   "Dar de baja"
      Height          =   255
      Left            =   4920
      TabIndex        =   8
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label lblFechaDeBaja 
      Alignment       =   2  'Center
      Caption         =   "Fecha de Baja"
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label lblDarDeBajaEmpleado 
      Alignment       =   2  'Center
      Caption         =   "Dar de baja empleado"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   120
      Width           =   5895
   End
   Begin VB.Label Label1 
      Caption         =   "Buscar Empleado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   8055
   End
End
Attribute VB_Name = "frmDarBajaEmpleado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_EmpleadoActual As clsEmpleado

Private Sub btnDarDeBaja_Click()
    On Error GoTo ErrHandler
    Dim mensaje As String
    Dim nuevoEstado As Boolean
    
    ' 1. Cláusula de Guarda: Validar si hay empleado
    If m_EmpleadoActual Is Nothing Then
        MsgBox "Primero debe seleccionar un empleado de la lista.", vbExclamation
        Exit Sub
    End If

    ' 2. Definir lógica según el Checkbox (Baja o Reincorporación)
    nuevoEstado = (cbxDarDeBaja.Value = vbUnchecked) ' Si no está tildado, lo queremos Activo
    
    If nuevoEstado Then
        mensaje = "¿Está seguro que desea REINCORPORAR a "
    Else
        mensaje = "¿Está seguro que desea dar de BAJA a "
    End If

    ' 3. Confirmación única
    If MsgBox(mensaje & m_EmpleadoActual.NombreCompleto & "?", _
              vbQuestion + vbYesNo, "Gestión de Estado") = vbNo Then Exit Sub

    ' 4. Ejecución
    m_EmpleadoActual.Activo = nuevoEstado
    
    If modEmpleados.GuardarEmpleado(m_EmpleadoActual) Then
        MsgBox "El estado del empleado se actualizó correctamente.", vbInformation
        
        ' 5. Refrescar UI
        Call modUI.LlenarComboEmpleadosConInactivos(cboEmpleados)
        ' Opcional: Limpiar campos o deseleccionar para evitar errores
        cboEmpleados.ListIndex = -1
    End If
ErrHandler:
    MsgBox "Ocurrió un error inesperado: " & Err.Description, vbCritical, "Error de Sistema"
End Sub

Private Sub Form_Load()
    Call modUI.LlenarComboEmpleadosConInactivos(cboEmpleados)
End Sub

Private Sub cboEmpleados_Click()
    Dim idSeleccionado As Long
    
    If cboEmpleados.ListIndex = -1 Then Exit Sub
    
    idSeleccionado = cboEmpleados.ItemData(cboEmpleados.ListIndex)
    
    Set m_EmpleadoActual = modEmpleados.GetEmpleadoById(idSeleccionado)
End Sub
