VERSION 5.00
Begin VB.Form frmModificarConcepto 
   Caption         =   "Form1"
   ClientHeight    =   5850
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   8520
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDescripcion 
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   1200
      Width           =   3615
   End
   Begin VB.TextBox txtMontoFijo 
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   2280
      Width           =   3615
   End
   Begin VB.TextBox txtPorcentaje 
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      Top             =   2280
      Width           =   3615
   End
   Begin VB.OptionButton optHaber 
      Caption         =   "Option1"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   3000
      Width           =   255
   End
   Begin VB.OptionButton optRetencion 
      Caption         =   "Option1"
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   3000
      Width           =   255
   End
   Begin VB.CommandButton btnCancelar 
      Caption         =   "Cancelar"
      Height          =   735
      Left            =   1560
      TabIndex        =   1
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton btnGuardar 
      Caption         =   "Guardar"
      Height          =   735
      Left            =   4080
      TabIndex        =   0
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label lblDescripcion 
      Caption         =   "Descripcion"
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label lblModificarConcepto 
      Caption         =   "Modificar de Concepto"
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
      TabIndex        =   11
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label lblMontoFijo 
      Caption         =   "Monto fijo"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label lblPorcentaje 
      Caption         =   "Porcentaje"
      Height          =   255
      Left            =   4440
      TabIndex        =   9
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label lblHaber 
      Caption         =   "Haber"
      Height          =   255
      Left            =   720
      TabIndex        =   8
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label lblRetencion 
      Caption         =   "Retencion"
      Height          =   255
      Left            =   2760
      TabIndex        =   7
      Top             =   3000
      Width           =   1215
   End
End
Attribute VB_Name = "frmModificarConcepto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Concepto As clsConcepto ' El objeto "ancla"

' Esta es la puerta de entrada para el objeto
Public Property Set ConceptoAEditar(ByRef o As clsConcepto)
    Set m_Concepto = o
End Property

Private Sub Form_Load()
    If m_Concepto Is Nothing Then
        MsgBox "Error: No se recibió ningún concepto para editar.", vbCritical
        Unload Me
        Exit Sub
    End If

    ' Llenamos los controles
    txtDescripcion.Text = m_Concepto.Descripcion
    txtMontoFijo.Text = Format(m_Concepto.MontoFijo, "0.00")
    txtPorcentaje.Text = Format(m_Concepto.Porcentaje, "0.00")
    
    ' Seteamos los OptionButtons según el tipo
    If m_Concepto.Tipo = "H" Then
        optHaber.Value = True
    Else
        optRetencion.Value = True
    End If
End Sub

Private Sub btnGuardar_Click()
    On Error GoTo ErrHandler
    ' 1. Validaciones (Igual que en el Alta)
    If Trim(txtDescripcion.Text) = "" Then
        MsgBox "La descripción no puede estar vacía", vbExclamation
        Exit Sub
    End If

    ' 2. Actualizamos el objeto que ya tenemos en memoria
    With m_Concepto
        .Descripcion = txtDescripcion.Text
        ' Recordá usar CDbl para que la coma no te juegue una mala pasada
        .MontoFijo = CDbl(IIf(txtMontoFijo.Text = "", 0, txtMontoFijo.Text))
        .Porcentaje = CDbl(IIf(txtPorcentaje.Text = "", 0, txtPorcentaje.Text))
        .Tipo = IIf(optHaber.Value, "H", "R")
        ' El .idConceptos no lo tocamos, así el módulo sabe que es un UPDATE
    End With

    ' 3. Persistimos
    If modConceptos.GuardarConcepto(m_Concepto) Then
        MsgBox "Cambios guardados correctamente", vbInformation
        Unload Me
    Else
        MsgBox "Error al actualizar el concepto", vbCritical
    End If
ErrHandler:
    MsgBox "Ocurrió un error inesperado: " & Err.Description, vbCritical, "Error de Sistema"
End Sub
