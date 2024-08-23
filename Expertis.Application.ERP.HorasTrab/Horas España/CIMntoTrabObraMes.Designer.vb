<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CIMntoTrabObraMes
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
        Dim cmbmes_DesignTimeLayout As Janus.Windows.GridEX.GridEXLayout = New Janus.Windows.GridEX.GridEXLayout
        Dim cmbanio_DesignTimeLayout As Janus.Windows.GridEX.GridEXLayout = New Janus.Windows.GridEX.GridEXLayout
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(CIMntoTrabObraMes))
        Me.advNObra = New Solmicro.Expertis.Engine.UI.AdvSearch
        Me.lblNObra = New Solmicro.Expertis.Engine.UI.Label
        Me.advIDOperario = New Solmicro.Expertis.Engine.UI.AdvSearch
        Me.lblIDOperario = New Solmicro.Expertis.Engine.UI.Label
        Me.cmbmes = New Solmicro.Expertis.Engine.UI.ComboBox
        Me.lblmes = New Solmicro.Expertis.Engine.UI.Label
        Me.cmbanio = New Solmicro.Expertis.Engine.UI.ComboBox
        Me.lblanio = New Solmicro.Expertis.Engine.UI.Label
        Me.clbFecha1 = New Solmicro.Expertis.Engine.UI.CalendarBox
        Me.lblFecha1 = New Solmicro.Expertis.Engine.UI.Label
        Me.clbFecha = New Solmicro.Expertis.Engine.UI.CalendarBox
        Me.lblFecha = New Solmicro.Expertis.Engine.UI.Label
        Me.FilterPanel.SuspendLayout()
        Me.CIMntoGridPanel.suspendlayout()
        CType(Me.Grid, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UiCommandManager1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Toolbar, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Menubar, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.MainPanel.SuspendLayout()
        Me.MainPanelCIMntoContainer.SuspendLayout()
        CType(Me.cmbmes, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cmbanio, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'FilterPanel
        '
        Me.FilterPanel.Controls.Add(Me.clbFecha1)
        Me.FilterPanel.Controls.Add(Me.lblFecha1)
        Me.FilterPanel.Controls.Add(Me.clbFecha)
        Me.FilterPanel.Controls.Add(Me.lblFecha)
        Me.FilterPanel.Controls.Add(Me.cmbanio)
        Me.FilterPanel.Controls.Add(Me.lblanio)
        Me.FilterPanel.Controls.Add(Me.cmbmes)
        Me.FilterPanel.Controls.Add(Me.lblmes)
        Me.FilterPanel.Controls.Add(Me.advIDOperario)
        Me.FilterPanel.Controls.Add(Me.lblIDOperario)
        Me.FilterPanel.Controls.Add(Me.advNObra)
        Me.FilterPanel.Controls.Add(Me.lblNObra)
        Me.FilterPanel.Location = New System.Drawing.Point(0, 303)
        Me.FilterPanel.Size = New System.Drawing.Size(954, 106)
        '
        'CIMntoGridPanel
        '
        Me.CIMntoGridPanel.Size = New System.Drawing.Size(954, 303)
        '
        'Grid
        '
        Grid_DesignTimeLayout.LayoutString = resources.GetString("Grid_DesignTimeLayout.LayoutString")
        Me.Grid.DesignTimeLayout = Grid_DesignTimeLayout
        Me.Grid.Size = New System.Drawing.Size(954, 303)
        Me.Grid.ViewName = "vSistLabListadoTrabajadoresObraMes"
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
        Me.MainPanel.Size = New System.Drawing.Size(954, 409)
        '
        'MainPanelCIMntoContainer
        '
        Me.MainPanelCIMntoContainer.Size = New System.Drawing.Size(954, 409)
        '
        'advNObra
        '
        Me.TryDataBinding(advNObra, New System.Windows.Forms.Binding("Text", Me, "NObra", True))
        Me.advNObra.DisabledBackColor = System.Drawing.Color.White
        Me.advNObra.DisplayField = "NObra"
        Me.advNObra.EntityName = "ObraCabecera"
        Me.advNObra.Location = New System.Drawing.Point(98, 28)
        Me.advNObra.Name = "advNObra"
        Me.advNObra.PrimaryDataFields = "NObra"
        Me.advNObra.Size = New System.Drawing.Size(121, 23)
        Me.advNObra.TabIndex = 0
        '
        'lblNObra
        '
        Me.lblNObra.Location = New System.Drawing.Point(21, 33)
        Me.lblNObra.Name = "lblNObra"
        Me.lblNObra.Size = New System.Drawing.Size(35, 13)
        Me.lblNObra.TabIndex = 1
        Me.lblNObra.Text = "Obra"
        '
        'advIDOperario
        '
        Me.TryDataBinding(advIDOperario, New System.Windows.Forms.Binding("Text", Me, "IDOperario", True))
        Me.advIDOperario.DisabledBackColor = System.Drawing.Color.White
        Me.advIDOperario.DisplayField = "IDOperario"
        Me.advIDOperario.EntityName = "Operario"
        Me.advIDOperario.Location = New System.Drawing.Point(98, 61)
        Me.advIDOperario.Name = "advIDOperario"
        Me.advIDOperario.PrimaryDataFields = "IDOperario"
        Me.advIDOperario.Size = New System.Drawing.Size(121, 23)
        Me.advIDOperario.TabIndex = 2
        '
        'lblIDOperario
        '
        Me.lblIDOperario.Location = New System.Drawing.Point(21, 66)
        Me.lblIDOperario.Name = "lblIDOperario"
        Me.lblIDOperario.Size = New System.Drawing.Size(57, 13)
        Me.lblIDOperario.TabIndex = 3
        Me.lblIDOperario.Text = "Operario"
        '
        'cmbmes
        '
        Me.TryDataBinding(cmbmes, New System.Windows.Forms.Binding("Value", Me, "mes", True))
        cmbmes_DesignTimeLayout.LayoutString = resources.GetString("cmbmes_DesignTimeLayout.LayoutString")
        Me.cmbmes.DesignTimeLayout = cmbmes_DesignTimeLayout
        Me.cmbmes.DisabledBackColor = System.Drawing.Color.White
        Me.cmbmes.Location = New System.Drawing.Point(393, 28)
        Me.cmbmes.Name = "cmbmes"
        Me.cmbmes.SelectedIndex = -1
        Me.cmbmes.SelectedItem = Nothing
        Me.cmbmes.Size = New System.Drawing.Size(121, 21)
        Me.cmbmes.TabIndex = 4
        '
        'lblmes
        '
        Me.lblmes.Location = New System.Drawing.Point(300, 33)
        Me.lblmes.Name = "lblmes"
        Me.lblmes.Size = New System.Drawing.Size(29, 13)
        Me.lblmes.TabIndex = 5
        Me.lblmes.Text = "Mes"
        '
        'cmbanio
        '
        Me.TryDataBinding(cmbanio, New System.Windows.Forms.Binding("Value", Me, "anio", True))
        cmbanio_DesignTimeLayout.LayoutString = resources.GetString("cmbanio_DesignTimeLayout.LayoutString")
        Me.cmbanio.DesignTimeLayout = cmbanio_DesignTimeLayout
        Me.cmbanio.DisabledBackColor = System.Drawing.Color.White
        Me.cmbanio.Location = New System.Drawing.Point(393, 63)
        Me.cmbanio.Name = "cmbanio"
        Me.cmbanio.SelectedIndex = -1
        Me.cmbanio.SelectedItem = Nothing
        Me.cmbanio.Size = New System.Drawing.Size(121, 21)
        Me.cmbanio.TabIndex = 6
        '
        'lblanio
        '
        Me.lblanio.Location = New System.Drawing.Point(300, 65)
        Me.lblanio.Name = "lblanio"
        Me.lblanio.Size = New System.Drawing.Size(29, 13)
        Me.lblanio.TabIndex = 7
        Me.lblanio.Text = "Año"
        Me.lblanio.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'clbFecha1
        '
        Me.clbFecha1.DisabledBackColor = System.Drawing.Color.White
        Me.clbFecha1.Location = New System.Drawing.Point(777, 58)
        Me.clbFecha1.Name = "clbFecha1"
        Me.clbFecha1.Size = New System.Drawing.Size(121, 21)
        Me.clbFecha1.TabIndex = 32
        '
        'lblFecha1
        '
        Me.lblFecha1.Location = New System.Drawing.Point(700, 66)
        Me.lblFecha1.Name = "lblFecha1"
        Me.lblFecha1.Size = New System.Drawing.Size(62, 13)
        Me.lblFecha1.TabIndex = 33
        Me.lblFecha1.Text = "Fecha <="
        '
        'clbFecha
        '
        Me.clbFecha.DisabledBackColor = System.Drawing.Color.White
        Me.clbFecha.Location = New System.Drawing.Point(777, 25)
        Me.clbFecha.Name = "clbFecha"
        Me.clbFecha.Size = New System.Drawing.Size(121, 21)
        Me.clbFecha.TabIndex = 30
        '
        'lblFecha
        '
        Me.lblFecha.Location = New System.Drawing.Point(700, 33)
        Me.lblFecha.Name = "lblFecha"
        Me.lblFecha.Size = New System.Drawing.Size(62, 13)
        Me.lblFecha.TabIndex = 31
        Me.lblFecha.Text = "Fecha >="
        '
        'CIMntoTrabObraMes
        '
        Me.ClientSize = New System.Drawing.Size(962, 497)
        Me.Name = "CIMntoTrabObraMes"
        Me.ViewName = "vSistLabListadoTrabajadoresObraMes"
        Me.FilterPanel.ResumeLayout(False)
        Me.FilterPanel.PerformLayout()
        Me.CIMntoGridPanel.ResumeLayout(False)
        CType(Me.Grid, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UiCommandManager1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Toolbar, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Menubar, System.ComponentModel.ISupportInitialize).EndInit()
        Me.MainPanel.ResumeLayout(False)
        Me.MainPanelCIMntoContainer.ResumeLayout(False)
        CType(Me.cmbmes, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cmbanio, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents advIDOperario As Solmicro.Expertis.Engine.UI.AdvSearch
    Friend WithEvents lblIDOperario As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents advNObra As Solmicro.Expertis.Engine.UI.AdvSearch
    Friend WithEvents lblNObra As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents cmbanio As Solmicro.Expertis.Engine.UI.ComboBox
    Friend WithEvents lblanio As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents cmbmes As Solmicro.Expertis.Engine.UI.ComboBox
    Friend WithEvents lblmes As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents clbFecha1 As Solmicro.Expertis.Engine.UI.CalendarBox
    Friend WithEvents lblFecha1 As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents clbFecha As Solmicro.Expertis.Engine.UI.CalendarBox
    Friend WithEvents lblFecha As Solmicro.Expertis.Engine.UI.Label

End Class
