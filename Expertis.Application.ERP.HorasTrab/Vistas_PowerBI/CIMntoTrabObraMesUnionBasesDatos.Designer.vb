<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CIMntoTrabObraMesUnionBasesDatos
    Inherits Solmicro.Expertis.Engine.UI.CIMnto

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
        Dim Grid_DesignTimeLayout As Janus.Windows.GridEX.GridEXLayout = New Janus.Windows.GridEX.GridEXLayout
        Dim cmbanio_DesignTimeLayout As Janus.Windows.GridEX.GridEXLayout = New Janus.Windows.GridEX.GridEXLayout
        Dim cmbmes_DesignTimeLayout As Janus.Windows.GridEX.GridEXLayout = New Janus.Windows.GridEX.GridEXLayout
        Dim cbAnioProductivo_DesignTimeLayout As Janus.Windows.GridEX.GridEXLayout = New Janus.Windows.GridEX.GridEXLayout
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(CIMntoTrabObraMesUnionBasesDatos))
        Dim cbMesProductivo_DesignTimeLayout As Janus.Windows.GridEX.GridEXLayout = New Janus.Windows.GridEX.GridEXLayout
        Me.clbFecha1 = New Solmicro.Expertis.Engine.UI.CalendarBox
        Me.lblFecha1 = New Solmicro.Expertis.Engine.UI.Label
        Me.clbFecha = New Solmicro.Expertis.Engine.UI.CalendarBox
        Me.lblFecha = New Solmicro.Expertis.Engine.UI.Label
        Me.cmbanio = New Solmicro.Expertis.Engine.UI.ComboBox
        Me.lblanio = New Solmicro.Expertis.Engine.UI.Label
        Me.cmbmes = New Solmicro.Expertis.Engine.UI.ComboBox
        Me.lblmes = New Solmicro.Expertis.Engine.UI.Label
        Me.advIDOperario = New Solmicro.Expertis.Engine.UI.AdvSearch
        Me.lblIDOperario = New Solmicro.Expertis.Engine.UI.Label
        Me.advNObra = New Solmicro.Expertis.Engine.UI.AdvSearch
        Me.lblNObra = New Solmicro.Expertis.Engine.UI.Label
        Me.cbAnioProductivo = New Solmicro.Expertis.Engine.UI.ComboBox
        Me.Label1 = New Solmicro.Expertis.Engine.UI.Label
        Me.cbMesProductivo = New Solmicro.Expertis.Engine.UI.ComboBox
        Me.Label2 = New Solmicro.Expertis.Engine.UI.Label
        Me.FilterPanel.SuspendLayout()
        Me.CIMntoGridPanel.suspendlayout()
        CType(Me.Grid, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UiCommandManager1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Toolbar, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Menubar, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.MainPanel.SuspendLayout()
        Me.MainPanelCIMntoContainer.SuspendLayout()
        CType(Me.cmbanio, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cmbmes, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cbAnioProductivo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cbMesProductivo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'FilterPanel
        '
        Me.FilterPanel.Controls.Add(Me.cbAnioProductivo)
        Me.FilterPanel.Controls.Add(Me.Label1)
        Me.FilterPanel.Controls.Add(Me.cbMesProductivo)
        Me.FilterPanel.Controls.Add(Me.Label2)
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
        Me.FilterPanel.Location = New System.Drawing.Point(0, 215)
        Me.FilterPanel.Size = New System.Drawing.Size(983, 118)
        '
        'CIMntoGridPanel
        '
        Me.CIMntoGridPanel.Size = New System.Drawing.Size(983, 215)
        '
        'Grid
        '
        Grid_DesignTimeLayout.LayoutString = resources.GetString("Grid_DesignTimeLayout.LayoutString")
        Me.Grid.DesignTimeLayout = Grid_DesignTimeLayout
        Me.Grid.Size = New System.Drawing.Size(983, 215)
        Me.Grid.ViewName = "vUnionSistLabListadoTrabajadoresObraMes"
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
        Me.MainPanel.Size = New System.Drawing.Size(983, 333)
        '
        'MainPanelCIMntoContainer
        '
        Me.MainPanelCIMntoContainer.Size = New System.Drawing.Size(983, 333)
        '
        'clbFecha1
        '
        Me.clbFecha1.DisabledBackColor = System.Drawing.Color.White
        Me.clbFecha1.Location = New System.Drawing.Point(595, 60)
        Me.clbFecha1.Name = "clbFecha1"
        Me.clbFecha1.Size = New System.Drawing.Size(121, 21)
        Me.clbFecha1.TabIndex = 44
        '
        'lblFecha1
        '
        Me.lblFecha1.Location = New System.Drawing.Point(518, 68)
        Me.lblFecha1.Name = "lblFecha1"
        Me.lblFecha1.Size = New System.Drawing.Size(62, 13)
        Me.lblFecha1.TabIndex = 45
        Me.lblFecha1.Text = "Fecha <="
        '
        'clbFecha
        '
        Me.clbFecha.DisabledBackColor = System.Drawing.Color.White
        Me.clbFecha.Location = New System.Drawing.Point(595, 27)
        Me.clbFecha.Name = "clbFecha"
        Me.clbFecha.Size = New System.Drawing.Size(121, 21)
        Me.clbFecha.TabIndex = 42
        '
        'lblFecha
        '
        Me.lblFecha.Location = New System.Drawing.Point(518, 35)
        Me.lblFecha.Name = "lblFecha"
        Me.lblFecha.Size = New System.Drawing.Size(62, 13)
        Me.lblFecha.TabIndex = 43
        Me.lblFecha.Text = "Fecha >="
        '
        'cmbanio
        '
        cmbanio_DesignTimeLayout.LayoutString = resources.GetString("cmbanio_DesignTimeLayout.LayoutString")
        Me.cmbanio.DesignTimeLayout = cmbanio_DesignTimeLayout
        Me.cmbanio.DisabledBackColor = System.Drawing.Color.White
        Me.cmbanio.Location = New System.Drawing.Point(361, 66)
        Me.cmbanio.Name = "cmbanio"
        Me.cmbanio.SelectedIndex = -1
        Me.cmbanio.SelectedItem = Nothing
        Me.cmbanio.Size = New System.Drawing.Size(121, 21)
        Me.cmbanio.TabIndex = 40
        '
        'lblanio
        '
        Me.lblanio.Location = New System.Drawing.Point(255, 67)
        Me.lblanio.Name = "lblanio"
        Me.lblanio.Size = New System.Drawing.Size(74, 13)
        Me.lblanio.TabIndex = 41
        Me.lblanio.Text = "Año Natural"
        '
        'cmbmes
        '
        cmbmes_DesignTimeLayout.LayoutString = resources.GetString("cmbmes_DesignTimeLayout.LayoutString")
        Me.cmbmes.DesignTimeLayout = cmbmes_DesignTimeLayout
        Me.cmbmes.DisabledBackColor = System.Drawing.Color.White
        Me.cmbmes.Location = New System.Drawing.Point(361, 31)
        Me.cmbmes.Name = "cmbmes"
        Me.cmbmes.SelectedIndex = -1
        Me.cmbmes.SelectedItem = Nothing
        Me.cmbmes.Size = New System.Drawing.Size(121, 21)
        Me.cmbmes.TabIndex = 38
        '
        'lblmes
        '
        Me.lblmes.Location = New System.Drawing.Point(255, 35)
        Me.lblmes.Name = "lblmes"
        Me.lblmes.Size = New System.Drawing.Size(74, 13)
        Me.lblmes.TabIndex = 39
        Me.lblmes.Text = "Mes Natural"
        '
        'advIDOperario
        '
        Me.advIDOperario.DisabledBackColor = System.Drawing.Color.White
        Me.advIDOperario.DisplayField = "IDOperario"
        Me.advIDOperario.EntityName = "Operario"
        Me.advIDOperario.Location = New System.Drawing.Point(97, 63)
        Me.advIDOperario.Name = "advIDOperario"
        Me.advIDOperario.PrimaryDataFields = "IDOperario"
        Me.advIDOperario.Size = New System.Drawing.Size(121, 23)
        Me.advIDOperario.TabIndex = 36
        '
        'lblIDOperario
        '
        Me.lblIDOperario.Location = New System.Drawing.Point(20, 68)
        Me.lblIDOperario.Name = "lblIDOperario"
        Me.lblIDOperario.Size = New System.Drawing.Size(57, 13)
        Me.lblIDOperario.TabIndex = 37
        Me.lblIDOperario.Text = "Operario"
        '
        'advNObra
        '
        Me.advNObra.DisabledBackColor = System.Drawing.Color.White
        Me.advNObra.DisplayField = "NObra"
        Me.advNObra.EntityName = "ObraCabecera"
        Me.advNObra.Location = New System.Drawing.Point(97, 30)
        Me.advNObra.Name = "advNObra"
        Me.advNObra.PrimaryDataFields = "NObra"
        Me.advNObra.Size = New System.Drawing.Size(121, 23)
        Me.advNObra.TabIndex = 34
        '
        'lblNObra
        '
        Me.lblNObra.Location = New System.Drawing.Point(20, 35)
        Me.lblNObra.Name = "lblNObra"
        Me.lblNObra.Size = New System.Drawing.Size(35, 13)
        Me.lblNObra.TabIndex = 35
        Me.lblNObra.Text = "Obra"
        '
        'cbAnioProductivo
        '
        cbAnioProductivo_DesignTimeLayout.LayoutString = resources.GetString("cbAnioProductivo_DesignTimeLayout.LayoutString")
        Me.cbAnioProductivo.DesignTimeLayout = cbAnioProductivo_DesignTimeLayout
        Me.cbAnioProductivo.DisabledBackColor = System.Drawing.Color.White
        Me.cbAnioProductivo.Location = New System.Drawing.Point(841, 61)
        Me.cbAnioProductivo.Name = "cbAnioProductivo"
        Me.cbAnioProductivo.SelectedIndex = -1
        Me.cbAnioProductivo.SelectedItem = Nothing
        Me.cbAnioProductivo.Size = New System.Drawing.Size(121, 21)
        Me.cbAnioProductivo.TabIndex = 48
        Me.cbAnioProductivo.Visible = False
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(735, 62)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(93, 13)
        Me.Label1.TabIndex = 49
        Me.Label1.Text = "Año Productivo"
        Me.Label1.Visible = False
        '
        'cbMesProductivo
        '
        cbMesProductivo_DesignTimeLayout.LayoutString = resources.GetString("cbMesProductivo_DesignTimeLayout.LayoutString")
        Me.cbMesProductivo.DesignTimeLayout = cbMesProductivo_DesignTimeLayout
        Me.cbMesProductivo.DisabledBackColor = System.Drawing.Color.White
        Me.cbMesProductivo.Location = New System.Drawing.Point(841, 26)
        Me.cbMesProductivo.Name = "cbMesProductivo"
        Me.cbMesProductivo.SelectedIndex = -1
        Me.cbMesProductivo.SelectedItem = Nothing
        Me.cbMesProductivo.Size = New System.Drawing.Size(121, 21)
        Me.cbMesProductivo.TabIndex = 46
        Me.cbMesProductivo.Visible = False
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(735, 30)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(93, 13)
        Me.Label2.TabIndex = 47
        Me.Label2.Text = "Mes Productivo"
        Me.Label2.Visible = False
        '
        'CIMntoTrabObraMesUnionBasesDatos
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(991, 421)
        Me.Name = "CIMntoTrabObraMesUnionBasesDatos"
        Me.Text = "CIMntoTrabObraMesUnionBasesDatos"
        Me.ViewName = "vUnionSistLabListadoTrabajadoresObraMes"
        Me.FilterPanel.ResumeLayout(False)
        Me.FilterPanel.PerformLayout()
        Me.CIMntoGridPanel.ResumeLayout(False)
        CType(Me.Grid, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UiCommandManager1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Toolbar, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Menubar, System.ComponentModel.ISupportInitialize).EndInit()
        Me.MainPanel.ResumeLayout(False)
        Me.MainPanelCIMntoContainer.ResumeLayout(False)
        CType(Me.cmbanio, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cmbmes, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cbAnioProductivo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cbMesProductivo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents clbFecha1 As Solmicro.Expertis.Engine.UI.CalendarBox
    Friend WithEvents lblFecha1 As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents clbFecha As Solmicro.Expertis.Engine.UI.CalendarBox
    Friend WithEvents lblFecha As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents cmbanio As Solmicro.Expertis.Engine.UI.ComboBox
    Friend WithEvents lblanio As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents cmbmes As Solmicro.Expertis.Engine.UI.ComboBox
    Friend WithEvents lblmes As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents advIDOperario As Solmicro.Expertis.Engine.UI.AdvSearch
    Friend WithEvents lblIDOperario As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents advNObra As Solmicro.Expertis.Engine.UI.AdvSearch
    Friend WithEvents lblNObra As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents cbAnioProductivo As Solmicro.Expertis.Engine.UI.ComboBox
    Friend WithEvents Label1 As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents cbMesProductivo As Solmicro.Expertis.Engine.UI.ComboBox
    Friend WithEvents Label2 As Solmicro.Expertis.Engine.UI.Label
End Class
