<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmBorraHoras
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
        Dim cbTipoHoras_DesignTimeLayout As Janus.Windows.GridEX.GridEXLayout = New Janus.Windows.GridEX.GridEXLayout
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmBorraHoras))
        Me.bCancelar = New Solmicro.Expertis.Engine.UI.Button
        Me.bAceptar = New Solmicro.Expertis.Engine.UI.Button
        Me.Frame1 = New Solmicro.Expertis.Engine.UI.Frame
        Me.cbTipoHoras = New Solmicro.Expertis.Engine.UI.ComboBox
        Me.Label5 = New Solmicro.Expertis.Engine.UI.Label
        Me.Fecha2 = New Solmicro.Expertis.Engine.UI.CalendarBox
        Me.Label4 = New Solmicro.Expertis.Engine.UI.Label
        Me.Label3 = New Solmicro.Expertis.Engine.UI.Label
        Me.Fecha1 = New Solmicro.Expertis.Engine.UI.CalendarBox
        Me.advObra = New Solmicro.Expertis.Engine.UI.AdvSearch
        Me.Label2 = New Solmicro.Expertis.Engine.UI.Label
        Me.Label1 = New Solmicro.Expertis.Engine.UI.Label
        Me.advPersona = New Solmicro.Expertis.Engine.UI.AdvSearch
        Me.Frame1.SuspendLayout()
        CType(Me.cbTipoHoras, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'bCancelar
        '
        Me.bCancelar.Location = New System.Drawing.Point(107, 371)
        Me.bCancelar.Name = "bCancelar"
        Me.bCancelar.Size = New System.Drawing.Size(85, 23)
        Me.bCancelar.TabIndex = 5
        Me.bCancelar.Text = "Cancelar"
        '
        'bAceptar
        '
        Me.bAceptar.Location = New System.Drawing.Point(294, 371)
        Me.bAceptar.Name = "bAceptar"
        Me.bAceptar.Size = New System.Drawing.Size(89, 23)
        Me.bAceptar.TabIndex = 4
        Me.bAceptar.Text = "Aceptar"
        '
        'Frame1
        '
        Me.Frame1.Controls.Add(Me.cbTipoHoras)
        Me.Frame1.Controls.Add(Me.Label5)
        Me.Frame1.Controls.Add(Me.Fecha2)
        Me.Frame1.Controls.Add(Me.Label4)
        Me.Frame1.Controls.Add(Me.Label3)
        Me.Frame1.Controls.Add(Me.Fecha1)
        Me.Frame1.Controls.Add(Me.advObra)
        Me.Frame1.Controls.Add(Me.Label2)
        Me.Frame1.Controls.Add(Me.Label1)
        Me.Frame1.Controls.Add(Me.advPersona)
        Me.Frame1.Location = New System.Drawing.Point(62, 30)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Size = New System.Drawing.Size(365, 301)
        Me.Frame1.TabIndex = 3
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Datos para borrar horas"
        '
        'cbTipoHoras
        '
        cbTipoHoras_DesignTimeLayout.LayoutString = resources.GetString("cbTipoHoras_DesignTimeLayout.LayoutString")
        Me.cbTipoHoras.DesignTimeLayout = cbTipoHoras_DesignTimeLayout
        Me.cbTipoHoras.DisabledBackColor = System.Drawing.Color.White
        Me.cbTipoHoras.Location = New System.Drawing.Point(158, 233)
        Me.cbTipoHoras.Name = "cbTipoHoras"
        Me.cbTipoHoras.SelectedIndex = -1
        Me.cbTipoHoras.SelectedItem = Nothing
        Me.cbTipoHoras.Size = New System.Drawing.Size(163, 21)
        Me.cbTipoHoras.TabIndex = 9
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(42, 237)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(68, 13)
        Me.Label5.TabIndex = 8
        Me.Label5.Text = "Tipo Horas"
        '
        'Fecha2
        '
        Me.Fecha2.DisabledBackColor = System.Drawing.Color.White
        Me.Fecha2.Location = New System.Drawing.Point(158, 183)
        Me.Fecha2.Name = "Fecha2"
        Me.Fecha2.Size = New System.Drawing.Size(163, 21)
        Me.Fecha2.TabIndex = 7
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(42, 183)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(62, 13)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Fecha <="
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(42, 133)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(62, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Fecha >="
        '
        'Fecha1
        '
        Me.Fecha1.DisabledBackColor = System.Drawing.Color.White
        Me.Fecha1.Location = New System.Drawing.Point(158, 135)
        Me.Fecha1.Name = "Fecha1"
        Me.Fecha1.Size = New System.Drawing.Size(163, 21)
        Me.Fecha1.TabIndex = 4
        '
        'advObra
        '
        Me.advObra.DisabledBackColor = System.Drawing.Color.White
        Me.advObra.DisplayField = "NObra"
        Me.advObra.EntityName = "ObraCabecera"
        Me.advObra.Location = New System.Drawing.Point(158, 82)
        Me.advObra.Name = "advObra"
        Me.advObra.PrimaryDataFields = "NObra"
        Me.advObra.Size = New System.Drawing.Size(163, 23)
        Me.advObra.TabIndex = 3
        Me.advObra.ViewName = "tbObraCabecera"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(42, 92)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(35, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Obra"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(42, 45)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(53, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Persona"
        '
        'advPersona
        '
        Me.advPersona.DisabledBackColor = System.Drawing.Color.White
        Me.advPersona.DisplayField = "IDOperario"
        Me.advPersona.EntityName = "Operario"
        Me.advPersona.Location = New System.Drawing.Point(158, 40)
        Me.advPersona.Name = "advPersona"
        Me.advPersona.PrimaryDataFields = "IDOperario"
        Me.advPersona.Size = New System.Drawing.Size(163, 23)
        Me.advPersona.TabIndex = 0
        Me.advPersona.ViewName = "tbMaestroOperario"
        '
        'frmBorraHoras
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(485, 431)
        Me.Controls.Add(Me.bCancelar)
        Me.Controls.Add(Me.bAceptar)
        Me.Controls.Add(Me.Frame1)
        Me.Name = "frmBorraHoras"
        Me.Text = "frmBorraHoras"
        Me.Frame1.ResumeLayout(False)
        Me.Frame1.PerformLayout()
        CType(Me.cbTipoHoras, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents bCancelar As Solmicro.Expertis.Engine.UI.Button
    Friend WithEvents bAceptar As Solmicro.Expertis.Engine.UI.Button
    Friend WithEvents Frame1 As Solmicro.Expertis.Engine.UI.Frame
    Friend WithEvents cbTipoHoras As Solmicro.Expertis.Engine.UI.ComboBox
    Friend WithEvents Label5 As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents Fecha2 As Solmicro.Expertis.Engine.UI.CalendarBox
    Friend WithEvents Label4 As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents Label3 As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents Fecha1 As Solmicro.Expertis.Engine.UI.CalendarBox
    Friend WithEvents advObra As Solmicro.Expertis.Engine.UI.AdvSearch
    Friend WithEvents Label2 As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents Label1 As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents advPersona As Solmicro.Expertis.Engine.UI.AdvSearch
End Class
