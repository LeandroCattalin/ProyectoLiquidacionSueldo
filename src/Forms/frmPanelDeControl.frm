VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPanelDeControl 
   Caption         =   "Form1"
   ClientHeight    =   8085
   ClientLeft      =   4065
   ClientTop       =   2625
   ClientWidth     =   14535
   LinkTopic       =   "Form1"
   ScaleHeight     =   8085
   ScaleWidth      =   14535
   Begin TabDlg.SSTab tabPanelDeControl 
      Height          =   5895
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   10398
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Dashboard"
      TabPicture(0)   =   "frmPanelDeControl.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraDashboard"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Liquidaciones"
      TabPicture(1)   =   "frmPanelDeControl.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraLiquidaciones"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Conceptos"
      TabPicture(2)   =   "frmPanelDeControl.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraConceptos"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Empleados"
      TabPicture(3)   =   "frmPanelDeControl.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraEmpleados"
      Tab(3).ControlCount=   1
      Begin VB.Frame fraDashboard 
         BorderStyle     =   0  'None
         Height          =   5295
         Left            =   -74880
         TabIndex        =   8
         Top             =   480
         Width           =   13215
         Begin VB.Timer tmrReloj 
            Interval        =   1000
            Left            =   12600
            Top             =   120
         End
         Begin VB.Label lblHora 
            Caption         =   "00:00"
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
            Left            =   4200
            TabIndex        =   13
            Top             =   480
            Width           =   3135
         End
         Begin VB.Label lblFecha 
            Caption         =   "1/1/1900"
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4200
            TabIndex        =   12
            Top             =   120
            Width           =   5895
         End
         Begin VB.Image Image1 
            Height          =   1620
            Left            =   0
            Picture         =   "frmPanelDeControl.frx":0070
            Stretch         =   -1  'True
            Top             =   0
            Width           =   3885
         End
         Begin VB.Label Label3 
            Caption         =   "Pendientes de firma: "
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   30
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   240
            TabIndex        =   11
            Top             =   4080
            Width           =   12855
         End
         Begin VB.Label Label2 
            Caption         =   "Liquidaciones del mes: "
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   30
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   240
            TabIndex        =   10
            Top             =   3240
            Width           =   12855
         End
         Begin VB.Label lblTotalEmpleados 
            Caption         =   "Total de empleados: "
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   30
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   240
            TabIndex        =   9
            Top             =   2400
            Width           =   12855
         End
      End
      Begin VB.Frame fraEmpleados 
         BorderStyle     =   0  'None
         Height          =   5415
         Left            =   -74760
         TabIndex        =   4
         Top             =   360
         Width           =   13095
         Begin MSComCtl2.DTPicker dtpFechaDeIngreso 
            Height          =   375
            Left            =   3120
            TabIndex        =   27
            Top             =   1920
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   661
            _Version        =   393216
            Format          =   115933185
            CurrentDate     =   36526
         End
         Begin VB.CommandButton btnBajaEmpleado 
            Caption         =   "Baja de Empleado"
            Height          =   735
            Left            =   11640
            TabIndex        =   26
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton btnGuardarModificaciones 
            Caption         =   "Guardar Modificaciones"
            Height          =   735
            Left            =   3720
            TabIndex        =   25
            Top             =   3360
            Width           =   1335
         End
         Begin VB.TextBox txtSueldoBasico 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   """$"" #.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
            Height          =   285
            Left            =   3240
            TabIndex        =   23
            Top             =   2640
            Width           =   2655
         End
         Begin VB.TextBox txtLegajo 
            Height          =   285
            Left            =   6000
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   1920
            Width           =   2655
         End
         Begin VB.TextBox txtCuil 
            Height          =   285
            Left            =   120
            TabIndex        =   18
            Top             =   1920
            Width           =   2655
         End
         Begin VB.TextBox txtApellido 
            Height          =   285
            Left            =   120
            TabIndex        =   15
            Top             =   1320
            Width           =   4815
         End
         Begin VB.TextBox txtNombre 
            Height          =   285
            Left            =   5040
            TabIndex        =   14
            Top             =   1320
            Width           =   4695
         End
         Begin VB.ComboBox cboEmpleados 
            Height          =   315
            Left            =   120
            TabIndex        =   6
            Top             =   480
            Width           =   9615
         End
         Begin VB.CommandButton btnAltaEmpleado 
            Caption         =   "Alta de Empleado"
            Height          =   735
            Left            =   10080
            TabIndex        =   5
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label lblSueldoBasico 
            Caption         =   "Sueldo basico"
            Height          =   255
            Left            =   3120
            TabIndex        =   24
            Top             =   2400
            Width           =   1695
         End
         Begin VB.Label lblLegajo 
            Caption         =   "Legajo"
            Height          =   255
            Left            =   6000
            TabIndex        =   22
            Top             =   1680
            Width           =   1695
         End
         Begin VB.Label lblFechaIngreso 
            Caption         =   "Fecha de ingreso"
            Height          =   255
            Left            =   3120
            TabIndex        =   20
            Top             =   1680
            Width           =   1695
         End
         Begin VB.Label lblCuil 
            Caption         =   "Cuil"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   1680
            Width           =   1695
         End
         Begin VB.Label lblApellido 
            Caption         =   "Apellido"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label lblNombre 
            Caption         =   "Nombre"
            Height          =   255
            Left            =   5040
            TabIndex        =   16
            Top             =   1080
            Width           =   1695
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
            TabIndex        =   7
            Top             =   120
            Width           =   8055
         End
      End
      Begin VB.Frame fraConceptos 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   5295
         Left            =   -74880
         TabIndex        =   3
         Top             =   480
         Width           =   13215
         Begin VB.CommandButton btnGuardarAsignacion 
            Caption         =   "Guardar asignacion"
            Height          =   735
            Left            =   9840
            TabIndex        =   34
            Top             =   840
            Width           =   1335
         End
         Begin MSComctlLib.ListView lvwConceptos 
            Height          =   4335
            Left            =   0
            TabIndex        =   33
            Top             =   720
            Width           =   9615
            _ExtentX        =   16960
            _ExtentY        =   7646
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin VB.ComboBox cboEmpleadoConceptos 
            Height          =   315
            Left            =   0
            TabIndex        =   31
            Top             =   360
            Width           =   9615
         End
         Begin VB.CommandButton btnBajaConcepto 
            Caption         =   "Baja concepto"
            Height          =   735
            Left            =   11760
            TabIndex        =   30
            Top             =   2280
            Width           =   1335
         End
         Begin VB.CommandButton btnModificarConcepto 
            Caption         =   "Modificar concepto"
            Height          =   735
            Left            =   11760
            TabIndex        =   29
            Top             =   1200
            Width           =   1335
         End
         Begin VB.CommandButton btnAltaConcepto 
            Caption         =   "Alta concepto"
            Height          =   735
            Left            =   11760
            TabIndex        =   28
            Top             =   120
            Width           =   1335
         End
         Begin VB.Label lblEmpleadoConceptos 
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
            Left            =   0
            TabIndex        =   32
            Top             =   0
            Width           =   8055
         End
      End
      Begin VB.Frame fraLiquidaciones 
         BorderStyle     =   0  'None
         Height          =   5295
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   13215
         Begin VB.CommandButton btnGenerarRecibos 
            Caption         =   "Generar Recibos"
            Height          =   855
            Left            =   4200
            TabIndex        =   42
            Top             =   2160
            Width           =   1695
         End
         Begin VB.CommandButton btnEjecutarLiquidacion 
            Caption         =   "Ejecutar Liquidacion"
            Height          =   855
            Left            =   4200
            TabIndex        =   41
            Top             =   960
            Width           =   1695
         End
         Begin VB.ComboBox cboAnio 
            Height          =   315
            Left            =   2160
            TabIndex        =   39
            Top             =   1440
            Width           =   1575
         End
         Begin VB.ComboBox cboMes 
            Height          =   315
            Left            =   120
            TabIndex        =   35
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label lblAnio 
            Caption         =   "Anio"
            Height          =   255
            Left            =   2160
            TabIndex        =   40
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label lblMes 
            Caption         =   "Mes"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label lblTituloLiquidaciones 
            Caption         =   "Calcular Liquidaciones"
            BeginProperty Font 
               Name            =   "Arial Rounded MT Bold"
               Size            =   23.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   37
            Top             =   120
            Width           =   6615
         End
         Begin VB.Label lblPeriodo 
            Caption         =   "Periodo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   36
            Top             =   840
            Width           =   1695
         End
      End
   End
   Begin VB.Label lblEstadoSesion 
      Caption         =   "Bienvenido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   10575
   End
End
Attribute VB_Name = "frmPanelDeControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_EmpleadoActual As clsEmpleado
Private m_IdEmpleadoActual As Long

'Propiedad y variable de modo para variar los controles que se muestran a cada usuario
Private m_Modo As String

Public Property Let Modo(ByVal vNuevoValor As String)
    m_Modo = vNuevoValor
End Property

Public Sub Form_Load()
    lblEstadoSesion.Caption = "Bienvenido " & UsuarioActual.NombreCompleto & " su rol actual es: " & UsuarioActual.RolDescripcion
    tabPanelDeControl.Tab = 0
    'Configurar ComboBox Periodo
    Call CargarCombosPeriodo(cboMes, cboAnio)
    'Columnas de la ListView de conceptos
    Call modUI.ConfigurarColumnasListView(lvwConceptos)
    Select Case m_Modo
        '1 para administrador, 2 para administrativo
        Case "1"
        Case "2"
    End Select
End Sub

Private Sub btnGenerarRecibos_Click()
    On Error GoTo ErrHandler
    Dim nMes As Integer
    Dim nAnio As Integer
    Dim respuesta As VbMsgBoxResult
    
    ' Validar selección
    If cboMes.ListIndex = -1 Or cboAnio.ListIndex = -1 Then
        MsgBox "Por favor, seleccione Mes y Año para generar recibos de liquidacion.", vbExclamation
        Exit Sub
    End If
    
    ' Convertierto los valores de los combos a Integer
    nMes = CInt(cboMes.Text)
    nAnio = CInt(cboAnio.Text)
    
    ' Muestro el periodo formateado solo para el mensaje
    respuesta = MsgBox("¿Desea comenzar el proceso de impresion de recibos para el período " & nMes & "/" & nAnio & "?" & vbCrLf & _
                "Esto procesará a todas las liquidaciones realizadas.", vbQuestion + vbYesNo, "Confirmar Proceso")
    
    If respuesta = vbYes Then
        Screen.MousePointer = vbHourglass
        btnGenerarRecibos.Enabled = False
        
        ' Llamo a la funcion para generar los recibos
        Call modUtils.GenerarRecibosHTML(nMes, nAnio)
        
        Screen.MousePointer = vbDefault
        btnGenerarRecibos.Enabled = True
        
        MsgBox ("Recibos Procesados con exito")
    End If
ErrHandler:
    MsgBox "Ocurrió un error inesperado: " & Err.Description, vbCritical, "Error de Sistema"
End Sub

Private Sub btnEjecutarLiquidacion_Click()
    Dim nMes As Integer
    Dim nAnio As Integer
    Dim respuesta As VbMsgBoxResult
    
    ' Validar selección
    If cboMes.ListIndex = -1 Or cboAnio.ListIndex = -1 Then
        MsgBox "Por favor, seleccione Mes y Año para liquidar.", vbExclamation
        Exit Sub
    End If
    
    ' Convertimos los valores de los combos a Integer
    nMes = CInt(cboMes.Text)
    nAnio = CInt(cboAnio.Text)
    
    ' Muestro el periodo formateado solo para el mensaje
    respuesta = MsgBox("¿Desea comenzar el proceso de liquidación para el período " & nMes & "/" & nAnio & "?" & vbCrLf & _
                "Esto procesará a todos los empleados activos.", vbQuestion + vbYesNo, "Confirmar Proceso")
    
    If respuesta = vbYes Then
        On Error GoTo ErrLiq
        
        Screen.MousePointer = vbHourglass
        btnEjecutarLiquidacion.Enabled = False
        
        ' Llamo a la funcion para procesar las liquidaciones
        Call modLiquidacion.ProcesarLiquidacion(nMes, nAnio)
        
        Screen.MousePointer = vbDefault
        btnEjecutarLiquidacion.Enabled = True
        
        ' Ya tengo los botones de exito en la funcion
    End If
    Exit Sub

ErrLiq:
    Screen.MousePointer = vbDefault
    btnEjecutarLiquidacion.Enabled = True
    MsgBox "Hubo un error crítico durante la liquidación: " & Err.Description, vbCritical
End Sub

'Boton para baja logica de conceptos
Private Sub btnBajaConcepto_Click()
    Dim oCon As clsConcepto
    Dim idSeleccionado As Long
    Dim respuesta As VbMsgBoxResult

    'Validar que haya algo seleccionado del listview
    If lvwConceptos.SelectedItem Is Nothing Then
        MsgBox "Por favor, seleccione un concepto de la lista.", vbExclamation, "Atención"
        Exit Sub
    End If

    'Pedir confirmacion
    respuesta = MsgBox("¿Está seguro que desea dar de baja el concepto: " & _
                lvwConceptos.SelectedItem.Text & "?", vbQuestion + vbYesNo, "Confirmar Baja")
    
    If respuesta = vbYes Then
        'Recuperar el ID del concepto sacando la C
        idSeleccionado = CLng(Mid(lvwConceptos.SelectedItem.Key, 2))
        
        'Traigo el objeto de clsConcepto de la db
        Set oCon = modConceptos.GetConceptoById(idSeleccionado)
        
        If Not oCon Is Nothing Then
            'Seteo el estado a negativo
            oCon.Activo = False
            
            'Guardo el cambio
            If modConceptos.GuardarConcepto(oCon) Then
                MsgBox "El concepto ha sido dado de baja correctamente.", vbInformation
                'Refresco la lista para que desaparezca
                Call modUI.LlenarListViewConceptos(lvwConceptos, m_IdEmpleadoActual)
            Else
                MsgBox "No se pudo procesar la baja.", vbCritical
            End If
        End If
    End If
End Sub

'Boton para ir al frmAltaDeConcepto
Private Sub btnAltaConcepto_Click()
    frmDarAltaConcepto.Show vbModal
End Sub

'Boton para ir al frmDarAltaEmpleado
Private Sub btnAltaEmpleado_Click()
    frmDarAltaEmpleado.Show vbModal
    cboEmpleados.Clear
    Call modUI.LlenarComboEmpleados(cboEmpleados)
End Sub

'Boton para ir al frmDarBajaEmpleado
Private Sub btnBajaEmpleado_Click()
    frmDarBajaEmpleado.Show vbModal
    cboEmpleados.Clear
    Call modUI.LlenarComboEmpleados(cboEmpleados)
End Sub

'Boton para guardar las modificaciones del Tab de empleados
Private Sub btnGuardarModificaciones_Click()
    'Validaciones
    If txtApellido.Text = "" Or txtNombre.Text = "" Then
        MsgBox "El nombre y apellido son obligatorios", vbExclamation
        Exit Sub
    End If
    
    'Paso los datos de los Textbox al objeto
    With m_EmpleadoActual
        On Error GoTo ErrGuardarModificaionesHandler
        Dim SueldoTexto As String
        SueldoTexto = txtSueldoBasico.Text
        
        SueldoTexto = Replace(SueldoTexto, "$", "")
        SueldoTexto = Trim(SueldoTexto)
        
        .Apellido = txtApellido.Text
        .Nombre = txtNombre.Text
        .Cuil = txtCuil.Text
        .Legajo = txtLegajo.Text
        .SueldoBasico = modUtils.ToDouble(SueldoTexto)
    End With
    
    'Llamo al modulo para cargar datos en la db
    If modEmpleados.GuardarEmpleado(m_EmpleadoActual) Then
        MsgBox "Cambios guardados correctamente", vbInformation
        cboEmpleados.Clear
        Call modUI.LlenarComboEmpleados(cboEmpleados)
    End If
ErrGuardarModificacionesHandler:
    MsgBox "Ocurrió un error inesperado: " & Err.Description, vbCritical, "Error de Sistema"
End Sub

'Boton para guardar la asignacion de nuevos conceptos a empleados
Private Sub btnGuardarAsignacion_Click()
    On Error GoTo ErrGuardarAsignacionHandler
    Dim colParaGuardar As New Collection
    Dim oCon As clsConcepto
    Dim i As Integer
    Dim idLimpio As Long
    
    'Valido que tenga un empleado seleccionado
    If m_IdEmpleadoActual <= 0 Then
        MsgBox "Debe seleccionar un empleado primero.", vbExclamation
        Exit Sub
    End If
    
    'Armo la coleccion de conceptos a cargar
    For i = 1 To lvwConceptos.ListItems.count
        'Solo guardo los que tengan el check en True
        If lvwConceptos.ListItems(i).Checked Then
            Set oCon = New clsConcepto
            
            'Extraigo el id sacando la C
            idLimpio = CLng(Mid(lvwConceptos.ListItems(i).Key, 2))
            
            'Guardo el objeto y lo agrego a la coleccion
            oCon.idConceptos = idLimpio
            colParaGuardar.Add oCon
        End If
    Next
    
    'Actualizo TODOS los conceptos del empleado desde 0
    If modConceptos.ActualizarAsignacion(m_IdEmpleadoActual, colParaGuardar) Then
        MsgBox "Asignación actualizada correctamente para el empleado.", vbInformation
    Else
        MsgBox "Hubo un error al guardar en la base de datos.", vbCritical
    End If
ErrGuardarAsingacionHandler:
    MsgBox "Ocurrió un error inesperado: " & Err.Description, vbCritical, "Error de Sistema"
End Sub

'Funcion cuando activo el cboEmpleados
Private Sub cboEmpleados_Click()
    On Error GoTo ErrCboEmpleadosHandler
    Dim idSeleccionado As Long
    
    ' Validamos que haya una selección válida
    If cboEmpleados.ListIndex = -1 Then
        MsgBox ("Seleccion Invalida")
    Else
        ' RECUPERAMOS EL ID REAL:
        idSeleccionado = cboEmpleados.ItemData(cboEmpleados.ListIndex)
        m_IdEmpleadoActual = cboEmpleados.ItemData(cboEmpleados.ListIndex)
        ' Pasamos el ID correcto
        Call MostrarEmpleado(idSeleccionado)
    End If
ErrCboEmpleadosHandler:
    MsgBox "Ocurrió un error inesperado al cargar el ComboBox: " & Err.Description, vbCritical, "Error de Sistema"
End Sub

Private Sub cboEmpleadoConceptos_Click()
    Dim idSeleccionado As Long
    
    ' Validamos que haya una selección válida
    If cboEmpleadoConceptos.ListIndex = -1 Then
        MsgBox ("Seleccion Invalida")
    Else
        ' RECUPERAMOS EL ID REAL:
        idSeleccionado = cboEmpleadoConceptos.ItemData(cboEmpleadoConceptos.ListIndex)
        m_IdEmpleadoActual = cboEmpleadoConceptos.ItemData(cboEmpleadoConceptos.ListIndex)
        ' Pasamos el ID correcto
        
        Call modUI.LlenarListViewConceptos(lvwConceptos, idSeleccionado)
    End If
End Sub

'Configuracion de funciones que se ejecutan al seleccionar cada Tab
Private Sub tabPanelDeControl_Click(PreviousTab As Integer)
    Select Case tabPanelDeControl.Tab
        Case 0
        Case 1
        Case 2
            Call modUI.LlenarComboEmpleados(cboEmpleadoConceptos)
        Case 3
            cboEmpleados.Clear
            Call modUI.LlenarComboEmpleados(cboEmpleados)
    End Select
End Sub

'Timer para actualizar la fecha y la hora
Private Sub tmrReloj_Timer()
    lblFecha.Caption = Format(Date, "dd/mm/yyyy")
    lblHora.Caption = Format(Time, "hh:mm:ss")
End Sub

'Carga al empleado en el formulario
Private Sub MostrarEmpleado(id As Long)
    On Error GoTo ErrMostrarEmpleadoHandler
    Dim oEmp As clsEmpleado
    Set oEmp = modEmpleados.GetEmpleadoById(id)

    If Not oEmp Is Nothing Then
        txtApellido.Text = oEmp.Apellido
        txtNombre.Text = oEmp.Nombre
        txtLegajo.Text = oEmp.Legajo
        txtCuil.Text = oEmp.Cuil
        dtpFechaDeIngreso.Value = oEmp.FechaIngreso
        txtSueldoBasico.Text = Format(Val(oEmp.SueldoBasico), "#,##0.00")
    Else
        MsgBox "No se pudo cargar la información del empleado.", vbExclamation
    End If
    Set m_EmpleadoActual = oEmp
ErrMotrarEmpleadoHandler:
    MsgBox "Ocurrió un error inesperado al mostrar el empleado: " & Err.Description, vbCritical, "Error de Sistema"
End Sub
