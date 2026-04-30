VERSION 5.00
Begin VB.Form frmDarAltaConcepto 
   Caption         =   "Form1"
   ClientHeight    =   6090
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   8475
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnGuardar 
      Caption         =   "Guardar"
      Height          =   735
      Left            =   4080
      TabIndex        =   12
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton btnCancelar 
      Caption         =   "Cancelar"
      Height          =   735
      Left            =   1560
      TabIndex        =   11
      Top             =   4080
      Width           =   1455
   End
   Begin VB.OptionButton optRetencion 
      Caption         =   "Option1"
      Height          =   255
      Left            =   2400
      TabIndex        =   9
      Top             =   3120
      Width           =   255
   End
   Begin VB.OptionButton optHaber 
      Caption         =   "Option1"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   3120
      Width           =   255
   End
   Begin VB.TextBox txtPorcentaje 
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   2400
      Width           =   3615
   End
   Begin VB.TextBox txtMontoFijo 
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   2400
      Width           =   3615
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   1320
      Width           =   3615
   End
   Begin VB.Label lblRetencion 
      Caption         =   "Retencion"
      Height          =   255
      Left            =   2760
      TabIndex        =   10
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label lblHaber 
      Caption         =   "Haber"
      Height          =   255
      Left            =   720
      TabIndex        =   8
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label lblPorcentaje 
      Caption         =   "Porcentaje"
      Height          =   255
      Left            =   4440
      TabIndex        =   6
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label lblMontoFijo 
      Caption         =   "Monto fijo"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Alta de Concepto"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   3735
   End
   Begin VB.Label lblDescripcion 
      Caption         =   "Descripcion"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   2055
   End
End
Attribute VB_Name = "frmDarAltaConcepto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnGuardar_Click()
    On Error GoTo ErrHandler
    Dim oCon As New clsConcepto
    
    ' Validaciones de entrada
    If Trim(txtDescripcion.Text) = "" Then
        MsgBox "La descripción es obligatoria", vbExclamation, "Faltan datos"
        txtDescripcion.SetFocus
        Exit Sub
    End If
    
    ' Validar que no mande los dos valores en 0 (opcional)
    If Val(txtMontoFijo.Text) = 0 And Val(txtPorcentaje.Text) = 0 Then
        MsgBox "Debe ingresar un monto fijo o un porcentaje", vbInformation
        Exit Sub
    End If

    ' Mapeo al objeto
    With oCon
        .Descripcion = txtDescripcion.Text
        .MontoFijo = CDbl(IIf(txtMontoFijo.Text = "", 0, txtMontoFijo.Text))
        .Porcentaje = CDbl(IIf(txtPorcentaje.Text = "", 0, txtPorcentaje.Text))
        
        If optHaber.Value Then
            .Tipo = "H"
        ElseIf optRetencion.Value Then
            .Tipo = "R"
        Else
            MsgBox "Debe seleccionar si el concepto es un Haber o una Retención", vbExclamation
            Exit Sub
        End If
        
        .Activo = True ' Por defecto para altas
    End With

    ' 3. Persistencia mediante el módulo
    If modConceptos.GuardarConcepto(oCon) Then
        MsgBox "Concepto '" & oCon.Descripcion & "' guardado con éxito", vbInformation
        Unload Me
    Else
        MsgBox "Error al intentar guardar en la base de datos", vbCritical
    End If
ErrHandler:
    MsgBox "Ocurrió un error inesperado: " & Err.Description, vbCritical, "Error de Sistema"
End Sub

Private Sub btnCancelar_Click()
    Unload Me
End Sub

Private Sub txtMontoFijo_Change()
    If Val(txtMontoFijo.Text) > 0 Then
        txtPorcentaje.Text = "0"
        txtPorcentaje.Enabled = False
    Else
        txtPorcentaje.Enabled = True
    End If
End Sub

Private Sub txtPorcentaje_Change()
    If Val(txtPorcentaje.Text) > 0 Then
        txtMontoFijo.Text = "0"
        txtMontoFijo.Enabled = False
    Else
        txtMontoFijo.Enabled = True
    End If
End Sub
