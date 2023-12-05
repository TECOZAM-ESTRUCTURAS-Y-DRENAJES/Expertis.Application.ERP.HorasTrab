<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CargaHorasJPSTAFF
    Inherits Solmicro.Expertis.Engine.UI.FormBase

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(CargaHorasJPSTAFF))
        Dim GridContratos_DesignTimeLayout As Janus.Windows.GridEX.GridEXLayout = New Janus.Windows.GridEX.GridEXLayout
        Dim Grid1_DesignTimeLayout As Janus.Windows.GridEX.GridEXLayout = New Janus.Windows.GridEX.GridEXLayout
        Dim Grid2_DesignTimeLayout As Janus.Windows.GridEX.GridEXLayout = New Janus.Windows.GridEX.GridEXLayout
        Dim Grid3_DesignTimeLayout As Janus.Windows.GridEX.GridEXLayout = New Janus.Windows.GridEX.GridEXLayout
        Me.bBorrarExcel = New Solmicro.Expertis.Engine.UI.Button
        Me.btnAceptar = New Solmicro.Expertis.Engine.UI.Button
        Me.cmdUbicacion = New Solmicro.Expertis.Engine.UI.Button
        Me.lblRuta = New Solmicro.Expertis.Engine.UI.Label
        Me.Label3 = New Solmicro.Expertis.Engine.UI.Label
        Me.LProgreso = New Solmicro.Expertis.Engine.UI.Label
        Me.PvProgreso = New System.Windows.Forms.ProgressBar
        Me.Label2 = New System.Windows.Forms.Label
        Me.bHorasOficina = New Solmicro.Expertis.Engine.UI.Button
        Me.bAñadirHorasPersona = New Solmicro.Expertis.Engine.UI.Button
        Me.bBorrarOperarioObraFecha = New Solmicro.Expertis.Engine.UI.Button
        Me.Frame1 = New Solmicro.Expertis.Engine.UI.Frame
        Me.Frame2 = New Solmicro.Expertis.Engine.UI.Frame
        Me.Frame3 = New Solmicro.Expertis.Engine.UI.Frame
        Me.Frame4 = New Solmicro.Expertis.Engine.UI.Frame
        Me.Frame5 = New Solmicro.Expertis.Engine.UI.Frame
        Me.bCreaHoras = New Solmicro.Expertis.Engine.UI.Button
        Me.Button1 = New Solmicro.Expertis.Engine.UI.Button
        Me.CD = New System.Windows.Forms.OpenFileDialog
        Me.Frame6 = New Solmicro.Expertis.Engine.UI.Frame
        Me.bA3 = New Solmicro.Expertis.Engine.UI.Button
        Me.bMixA3Horas = New Solmicro.Expertis.Engine.UI.Button
        Me.bIDGET = New Solmicro.Expertis.Engine.UI.Button
        Me.Frame7 = New Solmicro.Expertis.Engine.UI.Frame
        Me.bExportarHoras = New Solmicro.Expertis.Engine.UI.Button
        Me.Frame9 = New Solmicro.Expertis.Engine.UI.Frame
        Me.bCrearHorasBaja = New Solmicro.Expertis.Engine.UI.Button
        Me.Frame10 = New Solmicro.Expertis.Engine.UI.Frame
        Me.bExtras = New Solmicro.Expertis.Engine.UI.Button
        Me.bDocumentacion = New Solmicro.Expertis.Engine.UI.Button
        Me.GridContratos = New Solmicro.Expertis.Engine.UI.Grid
        Me.Panel3 = New Solmicro.Expertis.Engine.UI.Panel
        Me.baContrato = New System.Windows.Forms.Button
        Me.bgContrato = New System.Windows.Forms.Button
        Me.Label1 = New Solmicro.Expertis.Engine.UI.Label
        Me.txtIDContrato = New Solmicro.Expertis.Engine.UI.TextBox
        Me.clbFInicio = New Solmicro.Expertis.Engine.UI.CalendarBox
        Me.clbFFin = New Solmicro.Expertis.Engine.UI.CalendarBox
        Me.Label4 = New Solmicro.Expertis.Engine.UI.Label
        Me.Label7 = New Solmicro.Expertis.Engine.UI.Label
        Me.txtRenta = New Solmicro.Expertis.Engine.UI.TextBox
        Me.advEncargado = New Solmicro.Expertis.Engine.UI.AdvSearch
        Me.Label5 = New Solmicro.Expertis.Engine.UI.Label
        Me.Label6 = New Solmicro.Expertis.Engine.UI.Label
        Me.txtFianza = New Solmicro.Expertis.Engine.UI.TextBox
        Me.advObra = New Solmicro.Expertis.Engine.UI.AdvSearch
        Me.Label8 = New Solmicro.Expertis.Engine.UI.Label
        Me.Label9 = New Solmicro.Expertis.Engine.UI.Label
        Me.txtSubcontrata = New Solmicro.Expertis.Engine.UI.TextBox
        Me.bImportar = New Solmicro.Expertis.Engine.UI.Button
        Me.Grid1 = New Solmicro.Expertis.Engine.UI.Grid
        Me.Panel2 = New Solmicro.Expertis.Engine.UI.Panel
        Me.baCuadrilla = New System.Windows.Forms.Button
        Me.bgCuadrilla = New System.Windows.Forms.Button
        Me.Frame8 = New Solmicro.Expertis.Engine.UI.Frame
        Me.ulC5 = New Solmicro.Expertis.Engine.UI.UnderLineLabel
        Me.ulC4 = New Solmicro.Expertis.Engine.UI.UnderLineLabel
        Me.ulC3 = New Solmicro.Expertis.Engine.UI.UnderLineLabel
        Me.ulC2 = New Solmicro.Expertis.Engine.UI.UnderLineLabel
        Me.ulC1 = New Solmicro.Expertis.Engine.UI.UnderLineLabel
        Me.Label16 = New Solmicro.Expertis.Engine.UI.Label
        Me.Label17 = New Solmicro.Expertis.Engine.UI.Label
        Me.advCond5 = New Solmicro.Expertis.Engine.UI.AdvSearch
        Me.Label18 = New Solmicro.Expertis.Engine.UI.Label
        Me.Label21 = New Solmicro.Expertis.Engine.UI.Label
        Me.cmCFFin = New Solmicro.Expertis.Engine.UI.CalendarBox
        Me.advCond4 = New Solmicro.Expertis.Engine.UI.AdvSearch
        Me.Label19 = New Solmicro.Expertis.Engine.UI.Label
        Me.cmCFInicio = New Solmicro.Expertis.Engine.UI.CalendarBox
        Me.Label22 = New Solmicro.Expertis.Engine.UI.Label
        Me.advCond3 = New Solmicro.Expertis.Engine.UI.AdvSearch
        Me.Label20 = New Solmicro.Expertis.Engine.UI.Label
        Me.advCond2 = New Solmicro.Expertis.Engine.UI.AdvSearch
        Me.advCond1 = New Solmicro.Expertis.Engine.UI.AdvSearch
        Me.Frame11 = New Solmicro.Expertis.Engine.UI.Frame
        Me.ulObra = New Solmicro.Expertis.Engine.UI.UnderLineLabel
        Me.ulEnc = New Solmicro.Expertis.Engine.UI.UnderLineLabel
        Me.ulJP = New Solmicro.Expertis.Engine.UI.UnderLineLabel
        Me.Label11 = New Solmicro.Expertis.Engine.UI.Label
        Me.Label15 = New Solmicro.Expertis.Engine.UI.Label
        Me.Label14 = New Solmicro.Expertis.Engine.UI.Label
        Me.txtZona = New Solmicro.Expertis.Engine.UI.TextBox
        Me.Label13 = New Solmicro.Expertis.Engine.UI.Label
        Me.advEncarg = New Solmicro.Expertis.Engine.UI.AdvSearch
        Me.txtPro = New Solmicro.Expertis.Engine.UI.TextBox
        Me.advJProd = New Solmicro.Expertis.Engine.UI.AdvSearch
        Me.advObr = New Solmicro.Expertis.Engine.UI.AdvSearch
        Me.Label12 = New Solmicro.Expertis.Engine.UI.Label
        Me.txtIDCuadrilla = New Solmicro.Expertis.Engine.UI.TextBox
        Me.Button3 = New System.Windows.Forms.Button
        Me.Button4 = New System.Windows.Forms.Button
        Me.TextBox1 = New Solmicro.Expertis.Engine.UI.TextBox
        Me.Label10 = New Solmicro.Expertis.Engine.UI.Label
        Me.txtPisoContac = New Solmicro.Expertis.Engine.UI.TextBox
        Me.txtTelefcon = New Solmicro.Expertis.Engine.UI.TextBox
        Me.txtpersocont = New Solmicro.Expertis.Engine.UI.TextBox
        Me.Label23 = New Solmicro.Expertis.Engine.UI.Label
        Me.Label24 = New Solmicro.Expertis.Engine.UI.Label
        Me.Grid2 = New Solmicro.Expertis.Engine.UI.Grid
        Me.Grid3 = New Solmicro.Expertis.Engine.UI.Grid
        Me.Tab1 = New Solmicro.Expertis.Engine.UI.Tab
        Me.UiTabPage1 = New Janus.Windows.UI.Tab.UITabPage
        Me.UiTabPage2 = New Janus.Windows.UI.Tab.UITabPage
        Me.UiTabPage3 = New Janus.Windows.UI.Tab.UITabPage
        Me.Frame16 = New Solmicro.Expertis.Engine.UI.Frame
        Me.bMatriz = New Solmicro.Expertis.Engine.UI.Button
        Me.Frame15 = New Solmicro.Expertis.Engine.UI.Frame
        Me.bNO = New Solmicro.Expertis.Engine.UI.Button
        Me.bUk = New Solmicro.Expertis.Engine.UI.Button
        Me.bDCZ = New Solmicro.Expertis.Engine.UI.Button
        Me.Frame14 = New Solmicro.Expertis.Engine.UI.Frame
        Me.bRegularizarSemestral = New Solmicro.Expertis.Engine.UI.Button
        Me.Frame12 = New Solmicro.Expertis.Engine.UI.Frame
        Me.UiTabPage4 = New Janus.Windows.UI.Tab.UITabPage
        Me.Frame13 = New Solmicro.Expertis.Engine.UI.Frame
        Me.bPisarFicheroExtra = New Solmicro.Expertis.Engine.UI.Button
        Me.UiTabPage5 = New Janus.Windows.UI.Tab.UITabPage
        Me.Frame17 = New Solmicro.Expertis.Engine.UI.Frame
        Me.bDuplicados = New Solmicro.Expertis.Engine.UI.Button
        Me.Frame1.SuspendLayout()
        Me.Frame2.SuspendLayout()
        Me.Frame3.SuspendLayout()
        Me.Frame4.SuspendLayout()
        Me.Frame5.SuspendLayout()
        Me.Frame6.SuspendLayout()
        Me.Frame7.SuspendLayout()
        Me.Frame9.SuspendLayout()
        Me.Frame10.SuspendLayout()
        CType(Me.GridContratos, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel3.suspendlayout()
        CType(Me.Grid1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.suspendlayout()
        Me.Frame8.SuspendLayout()
        Me.Frame11.SuspendLayout()
        CType(Me.Grid2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Grid3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Tab1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Tab1.SuspendLayout()
        Me.UiTabPage1.SuspendLayout()
        Me.UiTabPage2.SuspendLayout()
        Me.UiTabPage3.SuspendLayout()
        Me.Frame16.SuspendLayout()
        Me.Frame15.SuspendLayout()
        Me.Frame14.SuspendLayout()
        Me.Frame12.SuspendLayout()
        Me.UiTabPage4.SuspendLayout()
        Me.Frame13.SuspendLayout()
        Me.UiTabPage5.SuspendLayout()
        Me.Frame17.SuspendLayout()
        Me.SuspendLayout()
        '
        'bBorrarExcel
        '
        Me.bBorrarExcel.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bBorrarExcel.Icon = CType(resources.GetObject("bBorrarExcel.Icon"), System.Drawing.Icon)
        Me.bBorrarExcel.Location = New System.Drawing.Point(22, 20)
        Me.bBorrarExcel.Name = "bBorrarExcel"
        Me.bBorrarExcel.Size = New System.Drawing.Size(382, 38)
        Me.bBorrarExcel.TabIndex = 16
        Me.bBorrarExcel.Text = "Borrar horas por Obra y Fecha (DescParte)"
        '
        'btnAceptar
        '
        Me.btnAceptar.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAceptar.Icon = CType(resources.GetObject("btnAceptar.Icon"), System.Drawing.Icon)
        Me.btnAceptar.Location = New System.Drawing.Point(118, 43)
        Me.btnAceptar.Name = "btnAceptar"
        Me.btnAceptar.Size = New System.Drawing.Size(94, 38)
        Me.btnAceptar.TabIndex = 15
        Me.btnAceptar.Text = "Aceptar"
        '
        'cmdUbicacion
        '
        Me.cmdUbicacion.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdUbicacion.Icon = CType(resources.GetObject("cmdUbicacion.Icon"), System.Drawing.Icon)
        Me.cmdUbicacion.Location = New System.Drawing.Point(9, 43)
        Me.cmdUbicacion.Name = "cmdUbicacion"
        Me.cmdUbicacion.Size = New System.Drawing.Size(94, 38)
        Me.cmdUbicacion.TabIndex = 14
        Me.cmdUbicacion.Text = "Buscar"
        '
        'lblRuta
        '
        Me.lblRuta.AutoSize = False
        Me.lblRuta.Location = New System.Drawing.Point(74, 170)
        Me.lblRuta.Name = "lblRuta"
        Me.lblRuta.Size = New System.Drawing.Size(440, 38)
        Me.lblRuta.TabIndex = 11
        Me.lblRuta.Text = "Ruta"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(17, 170)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(42, 13)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "Ruta: "
        '
        'LProgreso
        '
        Me.LProgreso.Location = New System.Drawing.Point(74, 140)
        Me.LProgreso.Name = "LProgreso"
        Me.LProgreso.Size = New System.Drawing.Size(97, 13)
        Me.LProgreso.TabIndex = 9
        Me.LProgreso.Text = "Progreso Actual"
        '
        'PvProgreso
        '
        Me.PvProgreso.Location = New System.Drawing.Point(72, 87)
        Me.PvProgreso.Name = "PvProgreso"
        Me.PvProgreso.Size = New System.Drawing.Size(656, 33)
        Me.PvProgreso.TabIndex = 8
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Times New Roman", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(68, 40)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(249, 22)
        Me.Label2.TabIndex = 17
        Me.Label2.Text = "Proyecto de costes Laborales"
        '
        'bHorasOficina
        '
        Me.bHorasOficina.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bHorasOficina.Icon = CType(resources.GetObject("bHorasOficina.Icon"), System.Drawing.Icon)
        Me.bHorasOficina.Location = New System.Drawing.Point(21, 43)
        Me.bHorasOficina.Name = "bHorasOficina"
        Me.bHorasOficina.Size = New System.Drawing.Size(137, 38)
        Me.bHorasOficina.TabIndex = 18
        Me.bHorasOficina.Text = "Crear horas a gente de oficina"
        '
        'bAñadirHorasPersona
        '
        Me.bAñadirHorasPersona.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bAñadirHorasPersona.Icon = CType(resources.GetObject("bAñadirHorasPersona.Icon"), System.Drawing.Icon)
        Me.bAñadirHorasPersona.Location = New System.Drawing.Point(27, 43)
        Me.bAñadirHorasPersona.Name = "bAñadirHorasPersona"
        Me.bAñadirHorasPersona.Size = New System.Drawing.Size(144, 38)
        Me.bAñadirHorasPersona.TabIndex = 19
        Me.bAñadirHorasPersona.Text = "Crear Horas Operario Obra y Fecha"
        '
        'bBorrarOperarioObraFecha
        '
        Me.bBorrarOperarioObraFecha.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bBorrarOperarioObraFecha.Icon = CType(resources.GetObject("bBorrarOperarioObraFecha.Icon"), System.Drawing.Icon)
        Me.bBorrarOperarioObraFecha.Location = New System.Drawing.Point(22, 64)
        Me.bBorrarOperarioObraFecha.Name = "bBorrarOperarioObraFecha"
        Me.bBorrarOperarioObraFecha.Size = New System.Drawing.Size(382, 38)
        Me.bBorrarOperarioObraFecha.TabIndex = 20
        Me.bBorrarOperarioObraFecha.Text = "Borrar Horas por Operario y Obra y Fecha"
        '
        'Frame1
        '
        Me.Frame1.Controls.Add(Me.cmdUbicacion)
        Me.Frame1.Controls.Add(Me.btnAceptar)
        Me.Frame1.Location = New System.Drawing.Point(241, 23)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Size = New System.Drawing.Size(236, 105)
        Me.Frame1.TabIndex = 21
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "2. Carga horas jefes de producción, técnicos de obra y staff(1, 4, 5)"
        '
        'Frame2
        '
        Me.Frame2.Controls.Add(Me.bHorasOficina)
        Me.Frame2.Location = New System.Drawing.Point(505, 23)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.Size = New System.Drawing.Size(176, 105)
        Me.Frame2.TabIndex = 22
        Me.Frame2.TabStop = False
        Me.Frame2.Text = "3. Carga horas personal  oficina(5)"
        '
        'Frame3
        '
        Me.Frame3.Controls.Add(Me.bAñadirHorasPersona)
        Me.Frame3.Location = New System.Drawing.Point(715, 23)
        Me.Frame3.Name = "Frame3"
        Me.Frame3.Size = New System.Drawing.Size(189, 105)
        Me.Frame3.TabIndex = 23
        Me.Frame3.TabStop = False
        Me.Frame3.Text = "4. Carga horas por persona"
        '
        'Frame4
        '
        Me.Frame4.Controls.Add(Me.bBorrarExcel)
        Me.Frame4.Controls.Add(Me.bBorrarOperarioObraFecha)
        Me.Frame4.Location = New System.Drawing.Point(37, 30)
        Me.Frame4.Name = "Frame4"
        Me.Frame4.Size = New System.Drawing.Size(417, 131)
        Me.Frame4.TabIndex = 24
        Me.Frame4.TabStop = False
        Me.Frame4.Text = "Borrar Horas"
        '
        'Frame5
        '
        Me.Frame5.Controls.Add(Me.bCreaHoras)
        Me.Frame5.Controls.Add(Me.Button1)
        Me.Frame5.Location = New System.Drawing.Point(490, 31)
        Me.Frame5.Name = "Frame5"
        Me.Frame5.Size = New System.Drawing.Size(281, 130)
        Me.Frame5.TabIndex = 26
        Me.Frame5.TabStop = False
        Me.Frame5.Text = "Carga Horas Otras Bases de Datos"
        '
        'bCreaHoras
        '
        Me.bCreaHoras.Icon = CType(resources.GetObject("bCreaHoras.Icon"), System.Drawing.Icon)
        Me.bCreaHoras.Location = New System.Drawing.Point(17, 64)
        Me.bCreaHoras.Name = "bCreaHoras"
        Me.bCreaHoras.Size = New System.Drawing.Size(241, 38)
        Me.bCreaHoras.TabIndex = 8
        Me.bCreaHoras.Text = "CREAR HORAS OTRA BASE DE DATOS"
        '
        'Button1
        '
        Me.Button1.Icon = CType(resources.GetObject("Button1.Icon"), System.Drawing.Icon)
        Me.Button1.Location = New System.Drawing.Point(17, 20)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(241, 38)
        Me.Button1.TabIndex = 6
        Me.Button1.Text = "Buscar"
        '
        'CD
        '
        Me.CD.FileName = "CD"
        '
        'Frame6
        '
        Me.Frame6.Controls.Add(Me.bA3)
        Me.Frame6.Location = New System.Drawing.Point(37, 28)
        Me.Frame6.Name = "Frame6"
        Me.Frame6.Size = New System.Drawing.Size(205, 103)
        Me.Frame6.TabIndex = 27
        Me.Frame6.TabStop = False
        Me.Frame6.Text = "1. Combinación de ficheros"
        '
        'bA3
        '
        Me.bA3.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bA3.Icon = CType(resources.GetObject("bA3.Icon"), System.Drawing.Icon)
        Me.bA3.Location = New System.Drawing.Point(22, 35)
        Me.bA3.Name = "bA3"
        Me.bA3.Size = New System.Drawing.Size(144, 38)
        Me.bA3.TabIndex = 19
        Me.bA3.Text = "Selecciona Ficheros A3"
        '
        'bMixA3Horas
        '
        Me.bMixA3Horas.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bMixA3Horas.Icon = CType(resources.GetObject("bMixA3Horas.Icon"), System.Drawing.Icon)
        Me.bMixA3Horas.Location = New System.Drawing.Point(29, 35)
        Me.bMixA3Horas.Name = "bMixA3Horas"
        Me.bMixA3Horas.Size = New System.Drawing.Size(144, 38)
        Me.bMixA3Horas.TabIndex = 20
        Me.bMixA3Horas.Text = "MIX A3 / Horas"
        '
        'bIDGET
        '
        Me.bIDGET.Location = New System.Drawing.Point(37, 190)
        Me.bIDGET.Name = "bIDGET"
        Me.bIDGET.Size = New System.Drawing.Size(189, 23)
        Me.bIDGET.TabIndex = 28
        Me.bIDGET.Text = "Actualiza IDGET"
        Me.bIDGET.Visible = False
        '
        'Frame7
        '
        Me.Frame7.Controls.Add(Me.bExportarHoras)
        Me.Frame7.Location = New System.Drawing.Point(36, 24)
        Me.Frame7.Name = "Frame7"
        Me.Frame7.Size = New System.Drawing.Size(189, 105)
        Me.Frame7.TabIndex = 29
        Me.Frame7.TabStop = False
        Me.Frame7.Text = "1. Exportación de todas las horas"
        '
        'bExportarHoras
        '
        Me.bExportarHoras.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bExportarHoras.Icon = CType(resources.GetObject("bExportarHoras.Icon"), System.Drawing.Icon)
        Me.bExportarHoras.Location = New System.Drawing.Point(20, 39)
        Me.bExportarHoras.Name = "bExportarHoras"
        Me.bExportarHoras.Size = New System.Drawing.Size(144, 38)
        Me.bExportarHoras.TabIndex = 19
        Me.bExportarHoras.Text = "Exportar"
        '
        'Frame9
        '
        Me.Frame9.Controls.Add(Me.bCrearHorasBaja)
        Me.Frame9.Location = New System.Drawing.Point(36, 23)
        Me.Frame9.Name = "Frame9"
        Me.Frame9.Size = New System.Drawing.Size(172, 105)
        Me.Frame9.TabIndex = 32
        Me.Frame9.TabStop = False
        Me.Frame9.Text = "1. Carga horas personal de baja España"
        '
        'bCrearHorasBaja
        '
        Me.bCrearHorasBaja.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bCrearHorasBaja.Icon = CType(resources.GetObject("bCrearHorasBaja.Icon"), System.Drawing.Icon)
        Me.bCrearHorasBaja.Location = New System.Drawing.Point(15, 43)
        Me.bCrearHorasBaja.Name = "bCrearHorasBaja"
        Me.bCrearHorasBaja.Size = New System.Drawing.Size(144, 38)
        Me.bCrearHorasBaja.TabIndex = 19
        Me.bCrearHorasBaja.Text = "Crear horas de baja"
        '
        'Frame10
        '
        Me.Frame10.Controls.Add(Me.bExtras)
        Me.Frame10.Location = New System.Drawing.Point(36, 23)
        Me.Frame10.Name = "Frame10"
        Me.Frame10.Size = New System.Drawing.Size(201, 100)
        Me.Frame10.TabIndex = 33
        Me.Frame10.TabStop = False
        Me.Frame10.Text = "1. Generacion 6 ficheros de extras"
        '
        'bExtras
        '
        Me.bExtras.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bExtras.Icon = CType(resources.GetObject("bExtras.Icon"), System.Drawing.Icon)
        Me.bExtras.Location = New System.Drawing.Point(24, 37)
        Me.bExtras.Name = "bExtras"
        Me.bExtras.Size = New System.Drawing.Size(144, 38)
        Me.bExtras.TabIndex = 19
        Me.bExtras.Text = "Generar previsión de extras"
        '
        'bDocumentacion
        '
        Me.bDocumentacion.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bDocumentacion.Icon = CType(resources.GetObject("bDocumentacion.Icon"), System.Drawing.Icon)
        Me.bDocumentacion.Location = New System.Drawing.Point(768, 88)
        Me.bDocumentacion.Name = "bDocumentacion"
        Me.bDocumentacion.Size = New System.Drawing.Size(141, 33)
        Me.bDocumentacion.TabIndex = 34
        Me.bDocumentacion.Text = "Documentación"
        '
        'GridContratos
        '
        Me.GridContratos.AllowAddNew = Janus.Windows.GridEX.InheritableBoolean.[False]
        Me.GridContratos.ColumnAutoResize = True
        GridContratos_DesignTimeLayout.LayoutString = resources.GetString("GridContratos_DesignTimeLayout.LayoutString")
        Me.GridContratos.DesignTimeLayout = GridContratos_DesignTimeLayout
        Me.GridContratos.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GridContratos.EnterKeyBehavior = Janus.Windows.GridEX.EnterKeyBehavior.NextCell
        Me.GridContratos.EntityName = "ContratosPisos"
        Me.GridContratos.Location = New System.Drawing.Point(0, 119)
        Me.GridContratos.Name = "GridContratos"
        Me.GridContratos.PrimaryDataFields = "IDPiso"
        Me.GridContratos.SecondaryDataFields = "IDPiso"
        Me.GridContratos.Size = New System.Drawing.Size(1264, 252)
        Me.GridContratos.TabIndex = 1
        Me.GridContratos.ViewName = "vFrmPisosContratos"
        '
        'Panel3
        '
        Me.Panel3.Controls.Add(Me.baContrato)
        Me.Panel3.Controls.Add(Me.bgContrato)
        Me.Panel3.Controls.Add(Me.Label1)
        Me.Panel3.Controls.Add(Me.txtIDContrato)
        Me.Panel3.Controls.Add(Me.clbFInicio)
        Me.Panel3.Controls.Add(Me.clbFFin)
        Me.Panel3.Controls.Add(Me.Label4)
        Me.Panel3.Controls.Add(Me.Label7)
        Me.Panel3.Controls.Add(Me.txtRenta)
        Me.Panel3.Controls.Add(Me.advEncargado)
        Me.Panel3.Controls.Add(Me.Label5)
        Me.Panel3.Controls.Add(Me.Label6)
        Me.Panel3.Controls.Add(Me.txtFianza)
        Me.Panel3.Controls.Add(Me.advObra)
        Me.Panel3.Controls.Add(Me.Label8)
        Me.Panel3.Controls.Add(Me.Label9)
        Me.Panel3.Controls.Add(Me.txtSubcontrata)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel3.Location = New System.Drawing.Point(0, 0)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(1264, 119)
        Me.Panel3.TabIndex = 0
        '
        'baContrato
        '
        Me.baContrato.ForeColor = System.Drawing.Color.Black
        Me.baContrato.Location = New System.Drawing.Point(1034, 59)
        Me.baContrato.Name = "baContrato"
        Me.baContrato.Size = New System.Drawing.Size(75, 23)
        Me.baContrato.TabIndex = 72
        Me.baContrato.Text = "Actualizar"
        Me.baContrato.UseVisualStyleBackColor = True
        '
        'bgContrato
        '
        Me.bgContrato.ForeColor = System.Drawing.Color.Black
        Me.bgContrato.Location = New System.Drawing.Point(910, 59)
        Me.bgContrato.Name = "bgContrato"
        Me.bgContrato.Size = New System.Drawing.Size(75, 23)
        Me.bgContrato.TabIndex = 71
        Me.bgContrato.Text = "Guardar"
        Me.bgContrato.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(21, 26)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(73, 13)
        Me.Label1.TabIndex = 55
        Me.Label1.Text = "Fecha inicio"
        '
        'txtIDContrato
        '
        Me.txtIDContrato.DisabledBackColor = System.Drawing.Color.White
        Me.txtIDContrato.Location = New System.Drawing.Point(266, 76)
        Me.txtIDContrato.Name = "txtIDContrato"
        Me.txtIDContrato.Size = New System.Drawing.Size(34, 21)
        Me.txtIDContrato.TabIndex = 70
        Me.txtIDContrato.Visible = False
        '
        'clbFInicio
        '
        Me.clbFInicio.DisabledBackColor = System.Drawing.Color.White
        Me.clbFInicio.Location = New System.Drawing.Point(100, 23)
        Me.clbFInicio.Name = "clbFInicio"
        Me.clbFInicio.Size = New System.Drawing.Size(136, 21)
        Me.clbFInicio.TabIndex = 54
        '
        'clbFFin
        '
        Me.clbFFin.DisabledBackColor = System.Drawing.Color.White
        Me.clbFFin.Location = New System.Drawing.Point(100, 59)
        Me.clbFFin.Name = "clbFFin"
        Me.clbFFin.Size = New System.Drawing.Size(136, 21)
        Me.clbFFin.TabIndex = 56
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(21, 62)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(58, 13)
        Me.Label4.TabIndex = 57
        Me.Label4.Text = "Fecha fin"
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(821, 27)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(60, 13)
        Me.Label7.TabIndex = 67
        Me.Label7.Text = "Jefe prod"
        '
        'txtRenta
        '
        Me.txtRenta.DisabledBackColor = System.Drawing.Color.White
        Me.txtRenta.Location = New System.Drawing.Point(384, 23)
        Me.txtRenta.Name = "txtRenta"
        Me.txtRenta.Size = New System.Drawing.Size(124, 21)
        Me.txtRenta.TabIndex = 58
        '
        'advEncargado
        '
        Me.advEncargado.DisabledBackColor = System.Drawing.Color.White
        Me.advEncargado.DisplayField = "IDOperario"
        Me.advEncargado.EntityName = "Operario"
        Me.advEncargado.Location = New System.Drawing.Point(910, 22)
        Me.advEncargado.Name = "advEncargado"
        Me.advEncargado.SecondaryDataFields = "IDOperario"
        Me.advEncargado.Size = New System.Drawing.Size(199, 23)
        Me.advEncargado.TabIndex = 66
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(318, 27)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(40, 13)
        Me.Label5.TabIndex = 59
        Me.Label5.Text = "Renta"
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(562, 26)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(35, 13)
        Me.Label6.TabIndex = 65
        Me.Label6.Text = "Obra"
        '
        'txtFianza
        '
        Me.txtFianza.DisabledBackColor = System.Drawing.Color.White
        Me.txtFianza.Location = New System.Drawing.Point(384, 58)
        Me.txtFianza.Name = "txtFianza"
        Me.txtFianza.Size = New System.Drawing.Size(124, 21)
        Me.txtFianza.TabIndex = 60
        '
        'advObra
        '
        Me.advObra.DisabledBackColor = System.Drawing.Color.White
        Me.advObra.DisplayField = "NObra"
        Me.advObra.EntityName = "ObraCabecera"
        Me.advObra.Location = New System.Drawing.Point(612, 21)
        Me.advObra.Name = "advObra"
        Me.advObra.SecondaryDataFields = "IDObra"
        Me.advObra.Size = New System.Drawing.Size(182, 23)
        Me.advObra.TabIndex = 63
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(318, 62)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(43, 13)
        Me.Label8.TabIndex = 61
        Me.Label8.Text = "Fianza"
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(530, 62)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(76, 13)
        Me.Label9.TabIndex = 63
        Me.Label9.Text = "Subcontrata"
        '
        'txtSubcontrata
        '
        Me.txtSubcontrata.DisabledBackColor = System.Drawing.Color.White
        Me.txtSubcontrata.Location = New System.Drawing.Point(612, 58)
        Me.txtSubcontrata.Name = "txtSubcontrata"
        Me.txtSubcontrata.Size = New System.Drawing.Size(182, 21)
        Me.txtSubcontrata.TabIndex = 64
        '
        'bImportar
        '
        Me.bImportar.Location = New System.Drawing.Point(71, 125)
        Me.bImportar.Name = "bImportar"
        Me.bImportar.Size = New System.Drawing.Size(95, 23)
        Me.bImportar.TabIndex = 18
        Me.bImportar.Text = "Importar Pisos"
        Me.bImportar.Visible = False
        '
        'Grid1
        '
        Me.Grid1.AllowAddNew = Janus.Windows.GridEX.InheritableBoolean.[False]
        Me.Grid1.AllowEdit = Janus.Windows.GridEX.InheritableBoolean.[False]
        Grid1_DesignTimeLayout.LayoutString = resources.GetString("Grid1_DesignTimeLayout.LayoutString")
        Me.Grid1.DesignTimeLayout = Grid1_DesignTimeLayout
        Me.Grid1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Grid1.EnterKeyBehavior = Janus.Windows.GridEX.EnterKeyBehavior.NextCell
        Me.Grid1.EntityName = "PisoCuadrilla"
        Me.Grid1.Location = New System.Drawing.Point(0, 256)
        Me.Grid1.Name = "Grid1"
        Me.Grid1.PrimaryDataFields = "IDPiso"
        Me.Grid1.SecondaryDataFields = "IDPiso"
        Me.Grid1.Size = New System.Drawing.Size(1264, 109)
        Me.Grid1.TabIndex = 55
        Me.Grid1.ViewName = "tbPisosCuadrilla"
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.baCuadrilla)
        Me.Panel2.Controls.Add(Me.bgCuadrilla)
        Me.Panel2.Controls.Add(Me.Frame8)
        Me.Panel2.Controls.Add(Me.Frame11)
        Me.Panel2.Controls.Add(Me.txtIDCuadrilla)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(0, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(1264, 365)
        Me.Panel2.TabIndex = 54
        '
        'baCuadrilla
        '
        Me.baCuadrilla.ForeColor = System.Drawing.Color.Black
        Me.baCuadrilla.Location = New System.Drawing.Point(993, 263)
        Me.baCuadrilla.Name = "baCuadrilla"
        Me.baCuadrilla.Size = New System.Drawing.Size(75, 23)
        Me.baCuadrilla.TabIndex = 36
        Me.baCuadrilla.Text = "Actualizar"
        Me.baCuadrilla.UseVisualStyleBackColor = True
        '
        'bgCuadrilla
        '
        Me.bgCuadrilla.ForeColor = System.Drawing.Color.Black
        Me.bgCuadrilla.Location = New System.Drawing.Point(858, 264)
        Me.bgCuadrilla.Name = "bgCuadrilla"
        Me.bgCuadrilla.Size = New System.Drawing.Size(75, 23)
        Me.bgCuadrilla.TabIndex = 35
        Me.bgCuadrilla.Text = "Guardar"
        Me.bgCuadrilla.UseVisualStyleBackColor = True
        '
        'Frame8
        '
        Me.Frame8.Controls.Add(Me.ulC5)
        Me.Frame8.Controls.Add(Me.ulC4)
        Me.Frame8.Controls.Add(Me.ulC3)
        Me.Frame8.Controls.Add(Me.ulC2)
        Me.Frame8.Controls.Add(Me.ulC1)
        Me.Frame8.Controls.Add(Me.Label16)
        Me.Frame8.Controls.Add(Me.Label17)
        Me.Frame8.Controls.Add(Me.advCond5)
        Me.Frame8.Controls.Add(Me.Label18)
        Me.Frame8.Controls.Add(Me.Label21)
        Me.Frame8.Controls.Add(Me.cmCFFin)
        Me.Frame8.Controls.Add(Me.advCond4)
        Me.Frame8.Controls.Add(Me.Label19)
        Me.Frame8.Controls.Add(Me.cmCFInicio)
        Me.Frame8.Controls.Add(Me.Label22)
        Me.Frame8.Controls.Add(Me.advCond3)
        Me.Frame8.Controls.Add(Me.Label20)
        Me.Frame8.Controls.Add(Me.advCond2)
        Me.Frame8.Controls.Add(Me.advCond1)
        Me.Frame8.Location = New System.Drawing.Point(19, 104)
        Me.Frame8.Name = "Frame8"
        Me.Frame8.Size = New System.Drawing.Size(1049, 153)
        Me.Frame8.TabIndex = 34
        Me.Frame8.TabStop = False
        Me.Frame8.Text = "Datos Trabajadores"
        '
        'ulC5
        '
        Me.ulC5.Location = New System.Drawing.Point(207, 108)
        Me.ulC5.Name = "ulC5"
        Me.ulC5.Size = New System.Drawing.Size(300, 13)
        Me.ulC5.TabIndex = 29
        '
        'ulC4
        '
        Me.ulC4.Location = New System.Drawing.Point(702, 70)
        Me.ulC4.Name = "ulC4"
        Me.ulC4.Size = New System.Drawing.Size(300, 13)
        Me.ulC4.TabIndex = 28
        '
        'ulC3
        '
        Me.ulC3.Location = New System.Drawing.Point(207, 70)
        Me.ulC3.Name = "ulC3"
        Me.ulC3.Size = New System.Drawing.Size(300, 13)
        Me.ulC3.TabIndex = 27
        '
        'ulC2
        '
        Me.ulC2.Location = New System.Drawing.Point(702, 35)
        Me.ulC2.Name = "ulC2"
        Me.ulC2.Size = New System.Drawing.Size(300, 13)
        Me.ulC2.TabIndex = 26
        '
        'ulC1
        '
        Me.ulC1.Location = New System.Drawing.Point(207, 35)
        Me.ulC1.Name = "ulC1"
        Me.ulC1.Size = New System.Drawing.Size(300, 13)
        Me.ulC1.TabIndex = 18
        '
        'Label16
        '
        Me.Label16.Location = New System.Drawing.Point(18, 35)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(64, 13)
        Me.Label16.TabIndex = 15
        Me.Label16.Text = "Persona 1"
        '
        'Label17
        '
        Me.Label17.Location = New System.Drawing.Point(513, 35)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(64, 13)
        Me.Label17.TabIndex = 16
        Me.Label17.Text = "Persona 2"
        '
        'advCond5
        '
        Me.advCond5.DisabledBackColor = System.Drawing.Color.White
        Me.advCond5.DisplayField = "IDOperario"
        Me.advCond5.EntityName = "Operario"
        Me.advCond5.Location = New System.Drawing.Point(101, 103)
        Me.advCond5.Name = "advCond5"
        Me.advCond5.SecondaryDataFields = "IDOperario"
        Me.advCond5.Size = New System.Drawing.Size(100, 23)
        Me.advCond5.TabIndex = 11
        '
        'Label18
        '
        Me.Label18.Location = New System.Drawing.Point(18, 70)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(64, 13)
        Me.Label18.TabIndex = 17
        Me.Label18.Text = "Persona 3"
        '
        'Label21
        '
        Me.Label21.Location = New System.Drawing.Point(732, 108)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(67, 13)
        Me.Label21.TabIndex = 24
        Me.Label21.Text = "Fecha fin :"
        '
        'cmCFFin
        '
        Me.cmCFFin.DisabledBackColor = System.Drawing.Color.White
        Me.cmCFFin.Location = New System.Drawing.Point(805, 104)
        Me.cmCFFin.Name = "cmCFFin"
        Me.cmCFFin.Size = New System.Drawing.Size(113, 21)
        Me.cmCFFin.TabIndex = 25
        '
        'advCond4
        '
        Me.advCond4.DisabledBackColor = System.Drawing.Color.White
        Me.advCond4.DisplayField = "IDOperario"
        Me.advCond4.EntityName = "Operario"
        Me.advCond4.Location = New System.Drawing.Point(596, 65)
        Me.advCond4.Name = "advCond4"
        Me.advCond4.SecondaryDataFields = "IDOperario"
        Me.advCond4.Size = New System.Drawing.Size(100, 23)
        Me.advCond4.TabIndex = 10
        '
        'Label19
        '
        Me.Label19.Location = New System.Drawing.Point(513, 70)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(64, 13)
        Me.Label19.TabIndex = 18
        Me.Label19.Text = "Persona 4"
        '
        'cmCFInicio
        '
        Me.cmCFInicio.DisabledBackColor = System.Drawing.Color.White
        Me.cmCFInicio.Location = New System.Drawing.Point(601, 104)
        Me.cmCFInicio.Name = "cmCFInicio"
        Me.cmCFInicio.Size = New System.Drawing.Size(104, 21)
        Me.cmCFInicio.TabIndex = 23
        '
        'Label22
        '
        Me.Label22.Location = New System.Drawing.Point(513, 108)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(82, 13)
        Me.Label22.TabIndex = 22
        Me.Label22.Text = "Fecha inicio :"
        '
        'advCond3
        '
        Me.advCond3.DisabledBackColor = System.Drawing.Color.White
        Me.advCond3.DisplayField = "IDOperario"
        Me.advCond3.EntityName = "Operario"
        Me.advCond3.Location = New System.Drawing.Point(101, 65)
        Me.advCond3.Name = "advCond3"
        Me.advCond3.SecondaryDataFields = "IDOperario"
        Me.advCond3.Size = New System.Drawing.Size(100, 23)
        Me.advCond3.TabIndex = 9
        '
        'Label20
        '
        Me.Label20.Location = New System.Drawing.Point(18, 108)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(64, 13)
        Me.Label20.TabIndex = 19
        Me.Label20.Text = "Persona 5"
        '
        'advCond2
        '
        Me.advCond2.DisabledBackColor = System.Drawing.Color.White
        Me.advCond2.DisplayField = "IDOperario"
        Me.advCond2.EntityName = "Operario"
        Me.advCond2.Location = New System.Drawing.Point(596, 30)
        Me.advCond2.Name = "advCond2"
        Me.advCond2.SecondaryDataFields = "IDOperario"
        Me.advCond2.Size = New System.Drawing.Size(100, 23)
        Me.advCond2.TabIndex = 8
        '
        'advCond1
        '
        Me.advCond1.DisabledBackColor = System.Drawing.Color.White
        Me.advCond1.DisplayField = "IDOperario"
        Me.advCond1.EntityName = "Operario"
        Me.advCond1.Location = New System.Drawing.Point(101, 30)
        Me.advCond1.Name = "advCond1"
        Me.advCond1.SecondaryDataFields = "IDOperario"
        Me.advCond1.Size = New System.Drawing.Size(100, 23)
        Me.advCond1.TabIndex = 7
        '
        'Frame11
        '
        Me.Frame11.Controls.Add(Me.ulObra)
        Me.Frame11.Controls.Add(Me.ulEnc)
        Me.Frame11.Controls.Add(Me.ulJP)
        Me.Frame11.Controls.Add(Me.Label11)
        Me.Frame11.Controls.Add(Me.Label15)
        Me.Frame11.Controls.Add(Me.Label14)
        Me.Frame11.Controls.Add(Me.txtZona)
        Me.Frame11.Controls.Add(Me.Label13)
        Me.Frame11.Controls.Add(Me.advEncarg)
        Me.Frame11.Controls.Add(Me.txtPro)
        Me.Frame11.Controls.Add(Me.advJProd)
        Me.Frame11.Controls.Add(Me.advObr)
        Me.Frame11.Controls.Add(Me.Label12)
        Me.Frame11.Location = New System.Drawing.Point(19, 3)
        Me.Frame11.Name = "Frame11"
        Me.Frame11.Size = New System.Drawing.Size(1049, 95)
        Me.Frame11.TabIndex = 33
        Me.Frame11.TabStop = False
        Me.Frame11.Text = "Datos Obra"
        '
        'ulObra
        '
        Me.ulObra.Location = New System.Drawing.Point(702, 18)
        Me.ulObra.Name = "ulObra"
        Me.ulObra.Size = New System.Drawing.Size(300, 31)
        Me.ulObra.TabIndex = 17
        '
        'ulEnc
        '
        Me.ulEnc.Location = New System.Drawing.Point(702, 61)
        Me.ulEnc.Name = "ulEnc"
        Me.ulEnc.Size = New System.Drawing.Size(300, 23)
        Me.ulEnc.TabIndex = 16
        Me.ulEnc.Visible = False
        '
        'ulJP
        '
        Me.ulJP.Location = New System.Drawing.Point(194, 61)
        Me.ulJP.Name = "ulJP"
        Me.ulJP.Size = New System.Drawing.Size(300, 23)
        Me.ulJP.TabIndex = 15
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(18, 35)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(36, 13)
        Me.Label11.TabIndex = 1
        Me.Label11.Text = "Zona"
        '
        'Label15
        '
        Me.Label15.Location = New System.Drawing.Point(523, 66)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(64, 13)
        Me.Label15.TabIndex = 14
        Me.Label15.Text = "Jefe prod."
        Me.Label15.Visible = False
        '
        'Label14
        '
        Me.Label14.Location = New System.Drawing.Point(18, 66)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(64, 13)
        Me.Label14.TabIndex = 13
        Me.Label14.Text = "Jefe Prod."
        '
        'txtZona
        '
        Me.txtZona.DisabledBackColor = System.Drawing.Color.White
        Me.txtZona.Location = New System.Drawing.Point(88, 31)
        Me.txtZona.Name = "txtZona"
        Me.txtZona.Size = New System.Drawing.Size(191, 21)
        Me.txtZona.TabIndex = 0
        '
        'Label13
        '
        Me.Label13.Location = New System.Drawing.Point(555, 35)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(35, 13)
        Me.Label13.TabIndex = 12
        Me.Label13.Text = "Obra"
        '
        'advEncarg
        '
        Me.advEncarg.DisabledBackColor = System.Drawing.Color.White
        Me.advEncarg.DisplayField = "IDOperario"
        Me.advEncarg.EntityName = "Operario"
        Me.advEncarg.Location = New System.Drawing.Point(596, 61)
        Me.advEncarg.Name = "advEncarg"
        Me.advEncarg.SecondaryDataFields = "IDOperario"
        Me.advEncarg.Size = New System.Drawing.Size(100, 23)
        Me.advEncarg.TabIndex = 6
        Me.advEncarg.Visible = False
        '
        'txtPro
        '
        Me.txtPro.DisabledBackColor = System.Drawing.Color.White
        Me.txtPro.Location = New System.Drawing.Point(350, 31)
        Me.txtPro.Name = "txtPro"
        Me.txtPro.Size = New System.Drawing.Size(152, 21)
        Me.txtPro.TabIndex = 2
        '
        'advJProd
        '
        Me.advJProd.DisabledBackColor = System.Drawing.Color.White
        Me.advJProd.DisplayField = "IDOperario"
        Me.advJProd.EntityName = "Operario"
        Me.advJProd.Location = New System.Drawing.Point(88, 61)
        Me.advJProd.Name = "advJProd"
        Me.advJProd.SecondaryDataFields = "IDOperario"
        Me.advJProd.Size = New System.Drawing.Size(100, 23)
        Me.advJProd.TabIndex = 5
        '
        'advObr
        '
        Me.advObr.DisabledBackColor = System.Drawing.Color.White
        Me.advObr.DisplayField = "NObra"
        Me.advObr.EntityName = "ObraCabecera"
        Me.advObr.Location = New System.Drawing.Point(596, 31)
        Me.advObr.Name = "advObr"
        Me.advObr.SecondaryDataFields = "IDObra"
        Me.advObr.Size = New System.Drawing.Size(100, 23)
        Me.advObr.TabIndex = 4
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(285, 35)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(59, 13)
        Me.Label12.TabIndex = 3
        Me.Label12.Text = "Provincia"
        '
        'txtIDCuadrilla
        '
        Me.txtIDCuadrilla.DisabledBackColor = System.Drawing.Color.White
        Me.txtIDCuadrilla.Location = New System.Drawing.Point(956, 171)
        Me.txtIDCuadrilla.Name = "txtIDCuadrilla"
        Me.txtIDCuadrilla.Size = New System.Drawing.Size(112, 21)
        Me.txtIDCuadrilla.TabIndex = 32
        Me.txtIDCuadrilla.Visible = False
        '
        'Button3
        '
        Me.Button3.ForeColor = System.Drawing.Color.Black
        Me.Button3.Location = New System.Drawing.Point(380, 105)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(75, 23)
        Me.Button3.TabIndex = 10
        Me.Button3.Text = "Actualizar"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.ForeColor = System.Drawing.Color.Black
        Me.Button4.Location = New System.Drawing.Point(261, 106)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(75, 23)
        Me.Button4.TabIndex = 9
        Me.Button4.Text = "Guardar"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.AcceptsReturn = True
        Me.TextBox1.DisabledBackColor = System.Drawing.Color.White
        Me.TextBox1.Location = New System.Drawing.Point(595, 20)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TextBox1.Size = New System.Drawing.Size(261, 90)
        Me.TextBox1.TabIndex = 5
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(482, 24)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(91, 13)
        Me.Label10.TabIndex = 8
        Me.Label10.Text = "Observaciones"
        '
        'txtPisoContac
        '
        Me.txtPisoContac.DisabledBackColor = System.Drawing.Color.White
        Me.txtPisoContac.Location = New System.Drawing.Point(100, 106)
        Me.txtPisoContac.Name = "txtPisoContac"
        Me.txtPisoContac.Size = New System.Drawing.Size(100, 21)
        Me.txtPisoContac.TabIndex = 7
        Me.txtPisoContac.Visible = False
        '
        'txtTelefcon
        '
        Me.txtTelefcon.DisabledBackColor = System.Drawing.Color.White
        Me.txtTelefcon.Location = New System.Drawing.Point(261, 63)
        Me.txtTelefcon.Name = "txtTelefcon"
        Me.txtTelefcon.Size = New System.Drawing.Size(194, 21)
        Me.txtTelefcon.TabIndex = 4
        '
        'txtpersocont
        '
        Me.txtpersocont.DisabledBackColor = System.Drawing.Color.White
        Me.txtpersocont.Location = New System.Drawing.Point(261, 20)
        Me.txtpersocont.Name = "txtpersocont"
        Me.txtpersocont.Size = New System.Drawing.Size(194, 21)
        Me.txtpersocont.TabIndex = 3
        '
        'Label23
        '
        Me.Label23.Location = New System.Drawing.Point(97, 71)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(133, 13)
        Me.Label23.TabIndex = 2
        Me.Label23.Text = "Telefono de Contacto:"
        '
        'Label24
        '
        Me.Label24.Location = New System.Drawing.Point(94, 24)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(131, 13)
        Me.Label24.TabIndex = 1
        Me.Label24.Text = "Persona de Contacto:"
        '
        'Grid2
        '
        Me.Grid2.AllowAddNew = Janus.Windows.GridEX.InheritableBoolean.[False]
        Me.Grid2.AllowEdit = Janus.Windows.GridEX.InheritableBoolean.[False]
        Me.Grid2.ColumnAutoResize = True
        Grid2_DesignTimeLayout.LayoutString = resources.GetString("Grid2_DesignTimeLayout.LayoutString")
        Me.Grid2.DesignTimeLayout = Grid2_DesignTimeLayout
        Me.Grid2.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Grid2.EnterKeyBehavior = Janus.Windows.GridEX.EnterKeyBehavior.NextCell
        Me.Grid2.EntityName = "PisoContacto"
        Me.Grid2.Location = New System.Drawing.Point(0, 112)
        Me.Grid2.Name = "Grid2"
        Me.Grid2.PrimaryDataFields = "IDPiso"
        Me.Grid2.SecondaryDataFields = "IDPiso"
        Me.Grid2.Size = New System.Drawing.Size(1264, 253)
        Me.Grid2.TabIndex = 0
        Me.Grid2.ViewName = "tbPisoContacto"
        '
        'Grid3
        '
        Me.Grid3.ColumnAutoResize = True
        Grid3_DesignTimeLayout.LayoutString = resources.GetString("Grid3_DesignTimeLayout.LayoutString")
        Me.Grid3.DesignTimeLayout = Grid3_DesignTimeLayout
        Me.Grid3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Grid3.EnterKeyBehavior = Janus.Windows.GridEX.EnterKeyBehavior.NextCell
        Me.Grid3.EntityName = "PisoPago"
        Me.Grid3.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Grid3.Location = New System.Drawing.Point(0, 0)
        Me.Grid3.Name = "Grid3"
        Me.Grid3.PrimaryDataFields = "IDPiso"
        Me.Grid3.SecondaryDataFields = "IDPiso"
        Me.Grid3.Size = New System.Drawing.Size(1264, 365)
        Me.Grid3.TabIndex = 0
        Me.Grid3.ViewName = "tbPisosPagos"
        '
        'Tab1
        '
        Me.Tab1.BackColor = System.Drawing.Color.FromArgb(CType(CType(238, Byte), Integer), CType(CType(238, Byte), Integer), CType(CType(238, Byte), Integer))
        Me.Tab1.Location = New System.Drawing.Point(62, 220)
        Me.Tab1.Name = "Tab1"
        Me.Tab1.Size = New System.Drawing.Size(983, 262)
        Me.Tab1.TabIndex = 35
        Me.Tab1.TabPages.AddRange(New Janus.Windows.UI.Tab.UITabPage() {Me.UiTabPage1, Me.UiTabPage2, Me.UiTabPage3, Me.UiTabPage4, Me.UiTabPage5})
        Me.Tab1.UseThemes = True
        '
        'UiTabPage1
        '
        Me.UiTabPage1.Controls.Add(Me.Frame9)
        Me.UiTabPage1.Controls.Add(Me.Frame2)
        Me.UiTabPage1.Controls.Add(Me.Frame3)
        Me.UiTabPage1.Controls.Add(Me.Frame1)
        Me.UiTabPage1.Location = New System.Drawing.Point(1, 21)
        Me.UiTabPage1.Name = "UiTabPage1"
        Me.UiTabPage1.Size = New System.Drawing.Size(981, 240)
        Me.UiTabPage1.TabStop = True
        Me.UiTabPage1.Text = "1. Creación de horas"
        '
        'UiTabPage2
        '
        Me.UiTabPage2.Controls.Add(Me.Frame7)
        Me.UiTabPage2.Location = New System.Drawing.Point(1, 21)
        Me.UiTabPage2.Name = "UiTabPage2"
        Me.UiTabPage2.Size = New System.Drawing.Size(959, 240)
        Me.UiTabPage2.TabStop = True
        Me.UiTabPage2.Text = "2. Exportar horas"
        '
        'UiTabPage3
        '
        Me.UiTabPage3.Controls.Add(Me.Frame17)
        Me.UiTabPage3.Controls.Add(Me.Frame16)
        Me.UiTabPage3.Controls.Add(Me.Frame15)
        Me.UiTabPage3.Controls.Add(Me.Frame14)
        Me.UiTabPage3.Controls.Add(Me.Frame12)
        Me.UiTabPage3.Controls.Add(Me.Frame6)
        Me.UiTabPage3.Location = New System.Drawing.Point(1, 21)
        Me.UiTabPage3.Name = "UiTabPage3"
        Me.UiTabPage3.Size = New System.Drawing.Size(981, 240)
        Me.UiTabPage3.TabStop = True
        Me.UiTabPage3.Text = "3. Combinar A3 y Mix"
        '
        'Frame16
        '
        Me.Frame16.Controls.Add(Me.bMatriz)
        Me.Frame16.Location = New System.Drawing.Point(751, 31)
        Me.Frame16.Name = "Frame16"
        Me.Frame16.Size = New System.Drawing.Size(205, 103)
        Me.Frame16.TabIndex = 32
        Me.Frame16.TabStop = False
        Me.Frame16.Text = "5. Matriz de horas"
        '
        'bMatriz
        '
        Me.bMatriz.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bMatriz.Icon = CType(resources.GetObject("bMatriz.Icon"), System.Drawing.Icon)
        Me.bMatriz.Location = New System.Drawing.Point(29, 35)
        Me.bMatriz.Name = "bMatriz"
        Me.bMatriz.Size = New System.Drawing.Size(144, 38)
        Me.bMatriz.TabIndex = 20
        Me.bMatriz.Text = "Matriz horas"
        '
        'Frame15
        '
        Me.Frame15.Controls.Add(Me.bNO)
        Me.Frame15.Controls.Add(Me.bUk)
        Me.Frame15.Controls.Add(Me.bDCZ)
        Me.Frame15.Location = New System.Drawing.Point(37, 138)
        Me.Frame15.Name = "Frame15"
        Me.Frame15.Size = New System.Drawing.Size(692, 87)
        Me.Frame15.TabIndex = 31
        Me.Frame15.TabStop = False
        Me.Frame15.Text = "4. Pasar ficheros de PDF a EXCEL"
        '
        'bNO
        '
        Me.bNO.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bNO.Icon = CType(resources.GetObject("bNO.Icon"), System.Drawing.Icon)
        Me.bNO.Location = New System.Drawing.Point(516, 34)
        Me.bNO.Name = "bNO"
        Me.bNO.Size = New System.Drawing.Size(144, 38)
        Me.bNO.TabIndex = 32
        Me.bNO.Text = "Pasar fichero NO"
        '
        'bUk
        '
        Me.bUk.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bUk.Icon = CType(resources.GetObject("bUk.Icon"), System.Drawing.Icon)
        Me.bUk.Location = New System.Drawing.Point(268, 34)
        Me.bUk.Name = "bUk"
        Me.bUk.Size = New System.Drawing.Size(144, 38)
        Me.bUk.TabIndex = 31
        Me.bUk.Text = "Pasar fichero UK"
        '
        'bDCZ
        '
        Me.bDCZ.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bDCZ.Icon = CType(resources.GetObject("bDCZ.Icon"), System.Drawing.Icon)
        Me.bDCZ.Location = New System.Drawing.Point(22, 34)
        Me.bDCZ.Name = "bDCZ"
        Me.bDCZ.Size = New System.Drawing.Size(144, 38)
        Me.bDCZ.TabIndex = 30
        Me.bDCZ.Text = "Pasar fichero DCZ"
        '
        'Frame14
        '
        Me.Frame14.Controls.Add(Me.bRegularizarSemestral)
        Me.Frame14.Location = New System.Drawing.Point(524, 28)
        Me.Frame14.Name = "Frame14"
        Me.Frame14.Size = New System.Drawing.Size(205, 103)
        Me.Frame14.TabIndex = 29
        Me.Frame14.TabStop = False
        Me.Frame14.Text = "3. Regularización semestral A3"
        '
        'bRegularizarSemestral
        '
        Me.bRegularizarSemestral.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bRegularizarSemestral.Icon = CType(resources.GetObject("bRegularizarSemestral.Icon"), System.Drawing.Icon)
        Me.bRegularizarSemestral.Location = New System.Drawing.Point(29, 35)
        Me.bRegularizarSemestral.Name = "bRegularizarSemestral"
        Me.bRegularizarSemestral.Size = New System.Drawing.Size(144, 38)
        Me.bRegularizarSemestral.TabIndex = 20
        Me.bRegularizarSemestral.Text = "Obtener fichero regularizacion"
        '
        'Frame12
        '
        Me.Frame12.Controls.Add(Me.bMixA3Horas)
        Me.Frame12.Location = New System.Drawing.Point(276, 28)
        Me.Frame12.Name = "Frame12"
        Me.Frame12.Size = New System.Drawing.Size(205, 103)
        Me.Frame12.TabIndex = 28
        Me.Frame12.TabStop = False
        Me.Frame12.Text = "2. Fiscalización datos"
        '
        'UiTabPage4
        '
        Me.UiTabPage4.Controls.Add(Me.Frame13)
        Me.UiTabPage4.Controls.Add(Me.Frame10)
        Me.UiTabPage4.Location = New System.Drawing.Point(1, 21)
        Me.UiTabPage4.Name = "UiTabPage4"
        Me.UiTabPage4.Size = New System.Drawing.Size(981, 240)
        Me.UiTabPage4.TabStop = True
        Me.UiTabPage4.Text = "4. Generación de extras"
        '
        'Frame13
        '
        Me.Frame13.Controls.Add(Me.bPisarFicheroExtra)
        Me.Frame13.Location = New System.Drawing.Point(279, 23)
        Me.Frame13.Name = "Frame13"
        Me.Frame13.Size = New System.Drawing.Size(201, 100)
        Me.Frame13.TabIndex = 34
        Me.Frame13.TabStop = False
        Me.Frame13.Text = "2. Pisar fichero extra con A3 original"
        '
        'bPisarFicheroExtra
        '
        Me.bPisarFicheroExtra.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bPisarFicheroExtra.Icon = CType(resources.GetObject("bPisarFicheroExtra.Icon"), System.Drawing.Icon)
        Me.bPisarFicheroExtra.Location = New System.Drawing.Point(24, 37)
        Me.bPisarFicheroExtra.Name = "bPisarFicheroExtra"
        Me.bPisarFicheroExtra.Size = New System.Drawing.Size(144, 38)
        Me.bPisarFicheroExtra.TabIndex = 19
        Me.bPisarFicheroExtra.Text = "Regularizar previsión extras"
        '
        'UiTabPage5
        '
        Me.UiTabPage5.Controls.Add(Me.Frame5)
        Me.UiTabPage5.Controls.Add(Me.Frame4)
        Me.UiTabPage5.Controls.Add(Me.bIDGET)
        Me.UiTabPage5.Location = New System.Drawing.Point(1, 21)
        Me.UiTabPage5.Name = "UiTabPage5"
        Me.UiTabPage5.Size = New System.Drawing.Size(959, 240)
        Me.UiTabPage5.TabStop = True
        Me.UiTabPage5.Text = "5. Otros"
        '
        'Frame17
        '
        Me.Frame17.Controls.Add(Me.bDuplicados)
        Me.Frame17.Location = New System.Drawing.Point(751, 140)
        Me.Frame17.Name = "Frame17"
        Me.Frame17.Size = New System.Drawing.Size(205, 86)
        Me.Frame17.TabIndex = 33
        Me.Frame17.TabStop = False
        Me.Frame17.Text = "6. Duplicidad de horas en distintas empresas"
        '
        'bDuplicados
        '
        Me.bDuplicados.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bDuplicados.Icon = CType(resources.GetObject("bDuplicados.Icon"), System.Drawing.Icon)
        Me.bDuplicados.Location = New System.Drawing.Point(29, 35)
        Me.bDuplicados.Name = "bDuplicados"
        Me.bDuplicados.Size = New System.Drawing.Size(144, 38)
        Me.bDuplicados.TabIndex = 20
        Me.bDuplicados.Text = "¿Duplicados?"
        '
        'CargaHorasJPSTAFF
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1102, 558)
        Me.Controls.Add(Me.Tab1)
        Me.Controls.Add(Me.bDocumentacion)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.lblRuta)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.LProgreso)
        Me.Controls.Add(Me.PvProgreso)
        Me.Name = "CargaHorasJPSTAFF"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Costes Laborales"
        Me.Frame1.ResumeLayout(False)
        Me.Frame2.ResumeLayout(False)
        Me.Frame3.ResumeLayout(False)
        Me.Frame4.ResumeLayout(False)
        Me.Frame5.ResumeLayout(False)
        Me.Frame6.ResumeLayout(False)
        Me.Frame7.ResumeLayout(False)
        Me.Frame9.ResumeLayout(False)
        Me.Frame10.ResumeLayout(False)
        CType(Me.GridContratos, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        CType(Me.Grid1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.Frame8.ResumeLayout(False)
        Me.Frame8.PerformLayout()
        Me.Frame11.ResumeLayout(False)
        Me.Frame11.PerformLayout()
        CType(Me.Grid2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Grid3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Tab1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Tab1.ResumeLayout(False)
        Me.UiTabPage1.ResumeLayout(False)
        Me.UiTabPage2.ResumeLayout(False)
        Me.UiTabPage3.ResumeLayout(False)
        Me.Frame16.ResumeLayout(False)
        Me.Frame15.ResumeLayout(False)
        Me.Frame14.ResumeLayout(False)
        Me.Frame12.ResumeLayout(False)
        Me.UiTabPage4.ResumeLayout(False)
        Me.Frame13.ResumeLayout(False)
        Me.UiTabPage5.ResumeLayout(False)
        Me.Frame17.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents bBorrarExcel As Solmicro.Expertis.Engine.UI.Button
    Friend WithEvents btnAceptar As Solmicro.Expertis.Engine.UI.Button
    Friend WithEvents cmdUbicacion As Solmicro.Expertis.Engine.UI.Button
    Friend WithEvents lblRuta As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents Label3 As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents LProgreso As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents PvProgreso As System.Windows.Forms.ProgressBar
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents bHorasOficina As Solmicro.Expertis.Engine.UI.Button
    Friend WithEvents bAñadirHorasPersona As Solmicro.Expertis.Engine.UI.Button
    Friend WithEvents bBorrarOperarioObraFecha As Solmicro.Expertis.Engine.UI.Button
    Friend WithEvents Frame1 As Solmicro.Expertis.Engine.UI.Frame
    Friend WithEvents Frame2 As Solmicro.Expertis.Engine.UI.Frame
    Friend WithEvents Frame3 As Solmicro.Expertis.Engine.UI.Frame
    Friend WithEvents Frame4 As Solmicro.Expertis.Engine.UI.Frame
    Friend WithEvents Frame5 As Solmicro.Expertis.Engine.UI.Frame
    Friend WithEvents bCreaHoras As Solmicro.Expertis.Engine.UI.Button
    Friend WithEvents Button1 As Solmicro.Expertis.Engine.UI.Button
    Friend WithEvents CD As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Frame6 As Solmicro.Expertis.Engine.UI.Frame
    Friend WithEvents bA3 As Solmicro.Expertis.Engine.UI.Button
    Friend WithEvents bIDGET As Solmicro.Expertis.Engine.UI.Button
    Friend WithEvents Frame7 As Solmicro.Expertis.Engine.UI.Frame
    Friend WithEvents bExportarHoras As Solmicro.Expertis.Engine.UI.Button
    Friend WithEvents Frame9 As Solmicro.Expertis.Engine.UI.Frame
    Friend WithEvents bCrearHorasBaja As Solmicro.Expertis.Engine.UI.Button
    Friend WithEvents bMixA3Horas As Solmicro.Expertis.Engine.UI.Button
    Friend WithEvents Frame10 As Solmicro.Expertis.Engine.UI.Frame
    Friend WithEvents bExtras As Solmicro.Expertis.Engine.UI.Button
    Friend WithEvents bDocumentacion As Solmicro.Expertis.Engine.UI.Button
    Friend WithEvents GridContratos As Solmicro.Expertis.Engine.UI.Grid
    Friend WithEvents Panel3 As Solmicro.Expertis.Engine.UI.Panel
    Friend WithEvents baContrato As System.Windows.Forms.Button
    Friend WithEvents bgContrato As System.Windows.Forms.Button
    Friend WithEvents Label1 As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents txtIDContrato As Solmicro.Expertis.Engine.UI.TextBox
    Friend WithEvents clbFInicio As Solmicro.Expertis.Engine.UI.CalendarBox
    Friend WithEvents clbFFin As Solmicro.Expertis.Engine.UI.CalendarBox
    Friend WithEvents Label4 As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents Label7 As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents txtRenta As Solmicro.Expertis.Engine.UI.TextBox
    Friend WithEvents advEncargado As Solmicro.Expertis.Engine.UI.AdvSearch
    Friend WithEvents Label5 As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents Label6 As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents txtFianza As Solmicro.Expertis.Engine.UI.TextBox
    Friend WithEvents advObra As Solmicro.Expertis.Engine.UI.AdvSearch
    Friend WithEvents Label8 As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents Label9 As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents txtSubcontrata As Solmicro.Expertis.Engine.UI.TextBox
    Friend WithEvents bImportar As Solmicro.Expertis.Engine.UI.Button
    Friend WithEvents Grid1 As Solmicro.Expertis.Engine.UI.Grid
    Friend WithEvents Panel2 As Solmicro.Expertis.Engine.UI.Panel
    Friend WithEvents baCuadrilla As System.Windows.Forms.Button
    Friend WithEvents bgCuadrilla As System.Windows.Forms.Button
    Friend WithEvents Frame8 As Solmicro.Expertis.Engine.UI.Frame
    Friend WithEvents ulC5 As Solmicro.Expertis.Engine.UI.UnderLineLabel
    Friend WithEvents ulC4 As Solmicro.Expertis.Engine.UI.UnderLineLabel
    Friend WithEvents ulC3 As Solmicro.Expertis.Engine.UI.UnderLineLabel
    Friend WithEvents ulC2 As Solmicro.Expertis.Engine.UI.UnderLineLabel
    Friend WithEvents ulC1 As Solmicro.Expertis.Engine.UI.UnderLineLabel
    Friend WithEvents Label16 As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents Label17 As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents advCond5 As Solmicro.Expertis.Engine.UI.AdvSearch
    Friend WithEvents Label18 As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents Label21 As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents cmCFFin As Solmicro.Expertis.Engine.UI.CalendarBox
    Friend WithEvents advCond4 As Solmicro.Expertis.Engine.UI.AdvSearch
    Friend WithEvents Label19 As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents cmCFInicio As Solmicro.Expertis.Engine.UI.CalendarBox
    Friend WithEvents Label22 As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents advCond3 As Solmicro.Expertis.Engine.UI.AdvSearch
    Friend WithEvents Label20 As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents advCond2 As Solmicro.Expertis.Engine.UI.AdvSearch
    Friend WithEvents advCond1 As Solmicro.Expertis.Engine.UI.AdvSearch
    Friend WithEvents Frame11 As Solmicro.Expertis.Engine.UI.Frame
    Friend WithEvents ulObra As Solmicro.Expertis.Engine.UI.UnderLineLabel
    Friend WithEvents ulEnc As Solmicro.Expertis.Engine.UI.UnderLineLabel
    Friend WithEvents ulJP As Solmicro.Expertis.Engine.UI.UnderLineLabel
    Friend WithEvents Label11 As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents Label15 As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents Label14 As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents txtZona As Solmicro.Expertis.Engine.UI.TextBox
    Friend WithEvents Label13 As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents advEncarg As Solmicro.Expertis.Engine.UI.AdvSearch
    Friend WithEvents txtPro As Solmicro.Expertis.Engine.UI.TextBox
    Friend WithEvents advJProd As Solmicro.Expertis.Engine.UI.AdvSearch
    Friend WithEvents advObr As Solmicro.Expertis.Engine.UI.AdvSearch
    Friend WithEvents Label12 As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents txtIDCuadrilla As Solmicro.Expertis.Engine.UI.TextBox
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents TextBox1 As Solmicro.Expertis.Engine.UI.TextBox
    Friend WithEvents Label10 As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents txtPisoContac As Solmicro.Expertis.Engine.UI.TextBox
    Friend WithEvents txtTelefcon As Solmicro.Expertis.Engine.UI.TextBox
    Friend WithEvents txtpersocont As Solmicro.Expertis.Engine.UI.TextBox
    Friend WithEvents Label23 As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents Label24 As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents Grid2 As Solmicro.Expertis.Engine.UI.Grid
    Friend WithEvents Grid3 As Solmicro.Expertis.Engine.UI.Grid
    Friend WithEvents Tab1 As Solmicro.Expertis.Engine.UI.Tab
    Friend WithEvents UiTabPage1 As Janus.Windows.UI.Tab.UITabPage
    Friend WithEvents UiTabPage2 As Janus.Windows.UI.Tab.UITabPage
    Friend WithEvents UiTabPage3 As Janus.Windows.UI.Tab.UITabPage
    Friend WithEvents UiTabPage4 As Janus.Windows.UI.Tab.UITabPage
    Friend WithEvents UiTabPage5 As Janus.Windows.UI.Tab.UITabPage
    Friend WithEvents Frame12 As Solmicro.Expertis.Engine.UI.Frame
    Friend WithEvents Frame13 As Solmicro.Expertis.Engine.UI.Frame
    Friend WithEvents bPisarFicheroExtra As Solmicro.Expertis.Engine.UI.Button
    Friend WithEvents Frame14 As Solmicro.Expertis.Engine.UI.Frame
    Friend WithEvents bRegularizarSemestral As Solmicro.Expertis.Engine.UI.Button
    Friend WithEvents bDCZ As Solmicro.Expertis.Engine.UI.Button
    Friend WithEvents Frame15 As Solmicro.Expertis.Engine.UI.Frame
    Friend WithEvents bUk As Solmicro.Expertis.Engine.UI.Button
    Friend WithEvents bNO As Solmicro.Expertis.Engine.UI.Button
    Friend WithEvents Frame16 As Solmicro.Expertis.Engine.UI.Frame
    Friend WithEvents bMatriz As Solmicro.Expertis.Engine.UI.Button
    Friend WithEvents Frame17 As Solmicro.Expertis.Engine.UI.Frame
    Friend WithEvents bDuplicados As Solmicro.Expertis.Engine.UI.Button
End Class
