<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmBorrarFicheroBDD
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmBorrarFicheroBDD))
        Dim cbTipo_DesignTimeLayout As Janus.Windows.GridEX.GridEXLayout = New Janus.Windows.GridEX.GridEXLayout
        Dim cbMes_DesignTimeLayout As Janus.Windows.GridEX.GridEXLayout = New Janus.Windows.GridEX.GridEXLayout
        Dim cbAnio_DesignTimeLayout As Janus.Windows.GridEX.GridEXLayout = New Janus.Windows.GridEX.GridEXLayout
        Me.btnCancelar = New Solmicro.Expertis.Engine.UI.Button
        Me.btnAceptar = New Solmicro.Expertis.Engine.UI.Button
        Me.Label3 = New Solmicro.Expertis.Engine.UI.Label
        Me.Label1 = New Solmicro.Expertis.Engine.UI.Label
        Me.cbTipo = New Solmicro.Expertis.Engine.UI.ComboBox
        Me.Label2 = New Solmicro.Expertis.Engine.UI.Label
        Me.Label4 = New Solmicro.Expertis.Engine.UI.Label
        Me.cbMes = New Solmicro.Expertis.Engine.UI.ComboBox
        Me.cbAnio = New Solmicro.Expertis.Engine.UI.ComboBox
        CType(Me.cbTipo, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cbMes, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cbAnio, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnCancelar
        '
        Me.btnCancelar.Icon = CType(resources.GetObject("btnCancelar.Icon"), System.Drawing.Icon)
        Me.btnCancelar.Location = New System.Drawing.Point(181, 211)
        Me.btnCancelar.Name = "btnCancelar"
        Me.btnCancelar.Size = New System.Drawing.Size(94, 23)
        Me.btnCancelar.TabIndex = 11
        Me.btnCancelar.Text = "Cancelar"
        '
        'btnAceptar
        '
        Me.btnAceptar.Icon = CType(resources.GetObject("btnAceptar.Icon"), System.Drawing.Icon)
        Me.btnAceptar.Location = New System.Drawing.Point(43, 211)
        Me.btnAceptar.Name = "btnAceptar"
        Me.btnAceptar.Size = New System.Drawing.Size(87, 23)
        Me.btnAceptar.TabIndex = 10
        Me.btnAceptar.Text = "Aceptar"
        '
        'Label3
        '
        Me.Label3.AutoSize = False
        Me.Label3.Location = New System.Drawing.Point(14, 21)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(303, 32)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "Introduzca los datos del fichero que desea borra de base de datos"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(14, 77)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(36, 13)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Tipo:"
        '
        'cbTipo
        '
        cbTipo_DesignTimeLayout.LayoutString = resources.GetString("cbTipo_DesignTimeLayout.LayoutString")
        Me.cbTipo.DesignTimeLayout = cbTipo_DesignTimeLayout
        Me.cbTipo.DisabledBackColor = System.Drawing.Color.White
        Me.cbTipo.Location = New System.Drawing.Point(103, 73)
        Me.cbTipo.Name = "cbTipo"
        Me.cbTipo.SelectedIndex = -1
        Me.cbTipo.SelectedItem = Nothing
        Me.cbTipo.Size = New System.Drawing.Size(202, 21)
        Me.cbTipo.TabIndex = 7
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(14, 122)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(34, 13)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "Mes:"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(14, 167)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(34, 13)
        Me.Label4.TabIndex = 13
        Me.Label4.Text = "Año:"
        '
        'cbMes
        '
        cbMes_DesignTimeLayout.LayoutString = resources.GetString("cbMes_DesignTimeLayout.LayoutString")
        Me.cbMes.DesignTimeLayout = cbMes_DesignTimeLayout
        Me.cbMes.DisabledBackColor = System.Drawing.Color.White
        Me.cbMes.Location = New System.Drawing.Point(103, 118)
        Me.cbMes.Name = "cbMes"
        Me.cbMes.SelectedIndex = -1
        Me.cbMes.SelectedItem = Nothing
        Me.cbMes.Size = New System.Drawing.Size(202, 21)
        Me.cbMes.TabIndex = 14
        '
        'cbAnio
        '
        cbAnio_DesignTimeLayout.LayoutString = resources.GetString("cbAnio_DesignTimeLayout.LayoutString")
        Me.cbAnio.DesignTimeLayout = cbAnio_DesignTimeLayout
        Me.cbAnio.DisabledBackColor = System.Drawing.Color.White
        Me.cbAnio.Location = New System.Drawing.Point(103, 163)
        Me.cbAnio.Name = "cbAnio"
        Me.cbAnio.SelectedIndex = -1
        Me.cbAnio.SelectedItem = Nothing
        Me.cbAnio.Size = New System.Drawing.Size(202, 21)
        Me.cbAnio.TabIndex = 15
        '
        'frmBorrarFicheroBDD
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(331, 261)
        Me.Controls.Add(Me.cbAnio)
        Me.Controls.Add(Me.cbMes)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.btnCancelar)
        Me.Controls.Add(Me.btnAceptar)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cbTipo)
        Me.Name = "frmBorrarFicheroBDD"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Borrar ficheros base de datos"
        CType(Me.cbTipo, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cbMes, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cbAnio, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnCancelar As Solmicro.Expertis.Engine.UI.Button
    Friend WithEvents btnAceptar As Solmicro.Expertis.Engine.UI.Button
    Friend WithEvents Label3 As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents Label1 As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents cbTipo As Solmicro.Expertis.Engine.UI.ComboBox
    Friend WithEvents Label2 As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents Label4 As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents cbMes As Solmicro.Expertis.Engine.UI.ComboBox
    Friend WithEvents cbAnio As Solmicro.Expertis.Engine.UI.ComboBox
End Class
