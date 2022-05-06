<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CIMntoBusqHorasAPP
    Inherits Solmicro.Expertis.Engine.UI.CIMnto

    'Form invalida a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim Grid_DesignTimeLayout As Janus.Windows.GridEX.GridEXLayout = New Janus.Windows.GridEX.GridEXLayout
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(CIMntoBusqHorasAPP))
        Me.advIDOperario = New Solmicro.Expertis.Engine.UI.AdvSearch
        Me.lblIDOperario = New Solmicro.Expertis.Engine.UI.Label
        Me.advNObra = New Solmicro.Expertis.Engine.UI.AdvSearch
        Me.lblNObra = New Solmicro.Expertis.Engine.UI.Label
        Me.advEncargado = New Solmicro.Expertis.Engine.UI.AdvSearch
        Me.lblEncargado = New Solmicro.Expertis.Engine.UI.Label
        Me.txtmatrVehiculo = New Solmicro.Expertis.Engine.UI.TextBox
        Me.lblmatrVehiculo = New Solmicro.Expertis.Engine.UI.Label
        Me.clbFecha = New Solmicro.Expertis.Engine.UI.CalendarBox
        Me.lblFecha = New Solmicro.Expertis.Engine.UI.Label
        Me.clbFecha1 = New Solmicro.Expertis.Engine.UI.CalendarBox
        Me.Label1 = New Solmicro.Expertis.Engine.UI.Label
        Me.FilterPanel.SuspendLayout()
        Me.CIMntoGridPanel.suspendlayout()
        CType(Me.Grid, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UiCommandManager1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Toolbar, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Menubar, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.MainPanel.SuspendLayout()
        Me.MainPanelCIMntoContainer.SuspendLayout()
        Me.SuspendLayout()
        '
        'FilterPanel
        '
        Me.FilterPanel.Controls.Add(Me.clbFecha1)
        Me.FilterPanel.Controls.Add(Me.Label1)
        Me.FilterPanel.Controls.Add(Me.clbFecha)
        Me.FilterPanel.Controls.Add(Me.lblFecha)
        Me.FilterPanel.Controls.Add(Me.txtmatrVehiculo)
        Me.FilterPanel.Controls.Add(Me.lblmatrVehiculo)
        Me.FilterPanel.Controls.Add(Me.advEncargado)
        Me.FilterPanel.Controls.Add(Me.lblEncargado)
        Me.FilterPanel.Controls.Add(Me.advNObra)
        Me.FilterPanel.Controls.Add(Me.lblNObra)
        Me.FilterPanel.Controls.Add(Me.advIDOperario)
        Me.FilterPanel.Controls.Add(Me.lblIDOperario)
        Me.FilterPanel.Location = New System.Drawing.Point(0, 300)
        Me.FilterPanel.Size = New System.Drawing.Size(918, 115)
        '
        'CIMntoGridPanel
        '
        Me.CIMntoGridPanel.Size = New System.Drawing.Size(918, 300)
        '
        'Grid
        '
        Grid_DesignTimeLayout.LayoutString = resources.GetString("Grid_DesignTimeLayout.LayoutString")
        Me.Grid.DesignTimeLayout = Grid_DesignTimeLayout
        Me.Grid.Size = New System.Drawing.Size(918, 300)
        Me.Grid.ViewName = "vParteProdHoras"
        '
        'Toolbar
        '
        Me.Toolbar.Size = New System.Drawing.Size(245, 28)
        '
        'Menubar
        '
        Me.Menubar.Location = New System.Drawing.Point(0, 28)
        '
        'MainPanel
        '
        Me.MainPanel.Size = New System.Drawing.Size(918, 415)
        '
        'MainPanelCIMntoContainer
        '
        Me.MainPanelCIMntoContainer.Size = New System.Drawing.Size(918, 415)
        '
        'advIDOperario
        '
        Me.TryDataBinding(advIDOperario, New System.Windows.Forms.Binding("Text", Me, "IDOperario", True))
        Me.advIDOperario.DisabledBackColor = System.Drawing.Color.White
        Me.advIDOperario.DisplayField = "IDOperario"
        Me.advIDOperario.EntityName = "Operario"
        Me.advIDOperario.Location = New System.Drawing.Point(96, 24)
        Me.advIDOperario.Name = "advIDOperario"
        Me.advIDOperario.PrimaryDataFields = "IDOperario"
        Me.advIDOperario.Size = New System.Drawing.Size(121, 23)
        Me.advIDOperario.TabIndex = 0
        '
        'lblIDOperario
        '
        Me.lblIDOperario.Location = New System.Drawing.Point(20, 29)
        Me.lblIDOperario.Name = "lblIDOperario"
        Me.lblIDOperario.Size = New System.Drawing.Size(71, 13)
        Me.lblIDOperario.TabIndex = 1
        Me.lblIDOperario.Text = "IDOperario"
        '
        'advNObra
        '
        Me.TryDataBinding(advNObra, New System.Windows.Forms.Binding("Text", Me, "NObra", True))
        Me.advNObra.DisabledBackColor = System.Drawing.Color.White
        Me.advNObra.DisplayField = "NObra"
        Me.advNObra.EntityName = "ObraCabecera"
        Me.advNObra.Location = New System.Drawing.Point(96, 63)
        Me.advNObra.Name = "advNObra"
        Me.advNObra.PrimaryDataFields = "NObra"
        Me.advNObra.Size = New System.Drawing.Size(121, 23)
        Me.advNObra.TabIndex = 2
        '
        'lblNObra
        '
        Me.lblNObra.Location = New System.Drawing.Point(20, 68)
        Me.lblNObra.Name = "lblNObra"
        Me.lblNObra.Size = New System.Drawing.Size(43, 13)
        Me.lblNObra.TabIndex = 3
        Me.lblNObra.Text = "NObra"
        '
        'advEncargado
        '
        Me.TryDataBinding(advEncargado, New System.Windows.Forms.Binding("Text", Me, "Encargado", True))
        Me.advEncargado.DisabledBackColor = System.Drawing.Color.White
        Me.advEncargado.DisplayField = "IDOperario"
        Me.advEncargado.EntityName = "Operario"
        Me.advEncargado.Location = New System.Drawing.Point(357, 24)
        Me.advEncargado.Name = "advEncargado"
        Me.advEncargado.PrimaryDataFields = "IDOperario"
        Me.advEncargado.Size = New System.Drawing.Size(121, 23)
        Me.advEncargado.TabIndex = 4
        '
        'lblEncargado
        '
        Me.lblEncargado.Location = New System.Drawing.Point(257, 29)
        Me.lblEncargado.Name = "lblEncargado"
        Me.lblEncargado.Size = New System.Drawing.Size(67, 13)
        Me.lblEncargado.TabIndex = 5
        Me.lblEncargado.Text = "Encargado"
        '
        'txtmatrVehiculo
        '
        Me.TryDataBinding(txtmatrVehiculo, New System.Windows.Forms.Binding("Text", Me, "matrVehiculo", True))
        Me.txtmatrVehiculo.DisabledBackColor = System.Drawing.Color.White
        Me.txtmatrVehiculo.Location = New System.Drawing.Point(357, 65)
        Me.txtmatrVehiculo.Name = "txtmatrVehiculo"
        Me.txtmatrVehiculo.Size = New System.Drawing.Size(121, 21)
        Me.txtmatrVehiculo.TabIndex = 6
        '
        'lblmatrVehiculo
        '
        Me.lblmatrVehiculo.Location = New System.Drawing.Point(257, 68)
        Me.lblmatrVehiculo.Name = "lblmatrVehiculo"
        Me.lblmatrVehiculo.Size = New System.Drawing.Size(81, 13)
        Me.lblmatrVehiculo.TabIndex = 7
        Me.lblmatrVehiculo.Text = "matrVehiculo"
        '
        'clbFecha
        '
        Me.TryDataBinding(clbFecha, New System.Windows.Forms.Binding("BindableValue", Me, "Fecha", True))
        Me.clbFecha.DisabledBackColor = System.Drawing.Color.White
        Me.clbFecha.Location = New System.Drawing.Point(608, 24)
        Me.clbFecha.Name = "clbFecha"
        Me.clbFecha.Size = New System.Drawing.Size(121, 21)
        Me.clbFecha.TabIndex = 8
        '
        'lblFecha
        '
        Me.lblFecha.Location = New System.Drawing.Point(543, 29)
        Me.lblFecha.Name = "lblFecha"
        Me.lblFecha.Size = New System.Drawing.Size(62, 13)
        Me.lblFecha.TabIndex = 9
        Me.lblFecha.Text = "Fecha >="
        '
        'clbFecha1
        '
        Me.clbFecha1.DisabledBackColor = System.Drawing.Color.White
        Me.clbFecha1.Location = New System.Drawing.Point(608, 63)
        Me.clbFecha1.Name = "clbFecha1"
        Me.clbFecha1.Size = New System.Drawing.Size(121, 21)
        Me.clbFecha1.TabIndex = 10
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(543, 68)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(62, 13)
        Me.Label1.TabIndex = 11
        Me.Label1.Text = "Fecha <="
        '
        'CIMntoBusqHorasAPP
        '
        Me.ClientSize = New System.Drawing.Size(926, 503)
        Me.Name = "CIMntoBusqHorasAPP"
        Me.ViewName = "vParteProdHoras"
        Me.FilterPanel.ResumeLayout(False)
        Me.FilterPanel.PerformLayout()
        Me.CIMntoGridPanel.ResumeLayout(False)
        CType(Me.Grid, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UiCommandManager1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Toolbar, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Menubar, System.ComponentModel.ISupportInitialize).EndInit()
        Me.MainPanel.ResumeLayout(False)
        Me.MainPanelCIMntoContainer.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents advNObra As Solmicro.Expertis.Engine.UI.AdvSearch
    Friend WithEvents lblNObra As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents advIDOperario As Solmicro.Expertis.Engine.UI.AdvSearch
    Friend WithEvents lblIDOperario As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents advEncargado As Solmicro.Expertis.Engine.UI.AdvSearch
    Friend WithEvents lblEncargado As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents clbFecha1 As Solmicro.Expertis.Engine.UI.CalendarBox
    Friend WithEvents Label1 As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents clbFecha As Solmicro.Expertis.Engine.UI.CalendarBox
    Friend WithEvents lblFecha As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents txtmatrVehiculo As Solmicro.Expertis.Engine.UI.TextBox
    Friend WithEvents lblmatrVehiculo As Solmicro.Expertis.Engine.UI.Label

End Class
