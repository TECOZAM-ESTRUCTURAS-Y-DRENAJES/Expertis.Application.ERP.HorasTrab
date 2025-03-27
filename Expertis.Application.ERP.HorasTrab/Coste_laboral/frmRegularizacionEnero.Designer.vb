<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmRegularizacionEnero
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.bImportar = New Solmicro.Expertis.Engine.UI.Button
        Me.lblRuta = New System.Windows.Forms.Label
        Me.bGenerar = New Solmicro.Expertis.Engine.UI.Button
        Me.txtPorcentaje1 = New Solmicro.Expertis.Engine.UI.TextBox
        Me.txtPorcentaje2 = New Solmicro.Expertis.Engine.UI.TextBox
        Me.cmbAnio = New System.Windows.Forms.ComboBox
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.cmbAnio)
        Me.GroupBox1.Controls.Add(Me.txtPorcentaje2)
        Me.GroupBox1.Controls.Add(Me.txtPorcentaje1)
        Me.GroupBox1.Controls.Add(Me.bGenerar)
        Me.GroupBox1.Controls.Add(Me.lblRuta)
        Me.GroupBox1.Controls.Add(Me.bImportar)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(363, 321)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Inputs"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 36)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(34, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Año:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 76)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(185, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Porcentaje 1 ( Del día 1 al 20):"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 116)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(192, 13)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Porcentaje 2 ( Del día 21 al 31):"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 159)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(138, 13)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "Excel 01 HORAS 01 XX"
        '
        'bImportar
        '
        Me.bImportar.Location = New System.Drawing.Point(202, 153)
        Me.bImportar.Name = "bImportar"
        Me.bImportar.Size = New System.Drawing.Size(127, 23)
        Me.bImportar.TabIndex = 4
        Me.bImportar.Text = "Importar Excel"
        '
        'lblRuta
        '
        Me.lblRuta.AutoSize = True
        Me.lblRuta.Location = New System.Drawing.Point(12, 209)
        Me.lblRuta.Name = "lblRuta"
        Me.lblRuta.Size = New System.Drawing.Size(38, 13)
        Me.lblRuta.TabIndex = 5
        Me.lblRuta.Text = "Ruta:"
        '
        'bGenerar
        '
        Me.bGenerar.Location = New System.Drawing.Point(15, 260)
        Me.bGenerar.Name = "bGenerar"
        Me.bGenerar.Size = New System.Drawing.Size(314, 23)
        Me.bGenerar.TabIndex = 6
        Me.bGenerar.Text = "Generar Excel"
        '
        'txtPorcentaje1
        '
        Me.txtPorcentaje1.DisabledBackColor = System.Drawing.Color.White
        Me.txtPorcentaje1.Location = New System.Drawing.Point(202, 76)
        Me.txtPorcentaje1.Name = "txtPorcentaje1"
        Me.txtPorcentaje1.Size = New System.Drawing.Size(127, 21)
        Me.txtPorcentaje1.TabIndex = 48
        '
        'txtPorcentaje2
        '
        Me.txtPorcentaje2.DisabledBackColor = System.Drawing.Color.White
        Me.txtPorcentaje2.Location = New System.Drawing.Point(202, 112)
        Me.txtPorcentaje2.Name = "txtPorcentaje2"
        Me.txtPorcentaje2.Size = New System.Drawing.Size(127, 21)
        Me.txtPorcentaje2.TabIndex = 49
        '
        'cmbAnio
        '
        Me.cmbAnio.FormattingEnabled = True
        Me.cmbAnio.Location = New System.Drawing.Point(202, 33)
        Me.cmbAnio.Name = "cmbAnio"
        Me.cmbAnio.Size = New System.Drawing.Size(121, 21)
        Me.cmbAnio.TabIndex = 50
        '
        'frmRegularizacionEnero
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(363, 321)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "frmRegularizacionEnero"
        Me.Text = "Regularizacion Enero Horas Administrativas Cat 2-3"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents bGenerar As Solmicro.Expertis.Engine.UI.Button
    Friend WithEvents lblRuta As System.Windows.Forms.Label
    Friend WithEvents bImportar As Solmicro.Expertis.Engine.UI.Button
    Friend WithEvents txtPorcentaje2 As Solmicro.Expertis.Engine.UI.TextBox
    Friend WithEvents txtPorcentaje1 As Solmicro.Expertis.Engine.UI.TextBox
    Friend WithEvents cmbAnio As System.Windows.Forms.ComboBox
End Class
