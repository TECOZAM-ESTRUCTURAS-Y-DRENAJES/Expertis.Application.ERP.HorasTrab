<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmCreaHorasDobleCotizacion
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
        Me.bCancelar = New Solmicro.Expertis.Engine.UI.Button
        Me.bAceptar = New Solmicro.Expertis.Engine.UI.Button
        Me.Frame1 = New Solmicro.Expertis.Engine.UI.Frame
        Me.Label5 = New Solmicro.Expertis.Engine.UI.Label
        Me.Label3 = New Solmicro.Expertis.Engine.UI.Label
        Me.Fecha1 = New Solmicro.Expertis.Engine.UI.CalendarBox
        Me.advObra = New Solmicro.Expertis.Engine.UI.AdvSearch
        Me.Label2 = New Solmicro.Expertis.Engine.UI.Label
        Me.Label1 = New Solmicro.Expertis.Engine.UI.Label
        Me.advPersona = New Solmicro.Expertis.Engine.UI.AdvSearch
        Me.txtHoras = New Solmicro.Expertis.Engine.UI.TextBox
        Me.Frame1.SuspendLayout()
        Me.SuspendLayout()
        '
        'bCancelar
        '
        Me.bCancelar.Location = New System.Drawing.Point(87, 267)
        Me.bCancelar.Name = "bCancelar"
        Me.bCancelar.Size = New System.Drawing.Size(85, 23)
        Me.bCancelar.TabIndex = 5
        Me.bCancelar.Text = "Cancelar"
        '
        'bAceptar
        '
        Me.bAceptar.Location = New System.Drawing.Point(298, 267)
        Me.bAceptar.Name = "bAceptar"
        Me.bAceptar.Size = New System.Drawing.Size(89, 23)
        Me.bAceptar.TabIndex = 4
        Me.bAceptar.Text = "Aceptar"
        '
        'Frame1
        '
        Me.Frame1.Controls.Add(Me.txtHoras)
        Me.Frame1.Controls.Add(Me.Label5)
        Me.Frame1.Controls.Add(Me.Label3)
        Me.Frame1.Controls.Add(Me.Fecha1)
        Me.Frame1.Controls.Add(Me.advObra)
        Me.Frame1.Controls.Add(Me.Label2)
        Me.Frame1.Controls.Add(Me.Label1)
        Me.Frame1.Controls.Add(Me.advPersona)
        Me.Frame1.Location = New System.Drawing.Point(51, 12)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Size = New System.Drawing.Size(365, 227)
        Me.Frame1.TabIndex = 3
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Datos de la creacion de horas"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(42, 181)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(40, 13)
        Me.Label5.TabIndex = 8
        Me.Label5.Text = "Horas"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(42, 135)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(53, 13)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Fecha ="
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
        'txtHoras
        '
        Me.txtHoras.DisabledBackColor = System.Drawing.Color.White
        Me.txtHoras.Enabled = False
        Me.txtHoras.Location = New System.Drawing.Point(158, 177)
        Me.txtHoras.Name = "txtHoras"
        Me.txtHoras.Size = New System.Drawing.Size(100, 21)
        Me.txtHoras.TabIndex = 9
        Me.txtHoras.Text = "0"
        '
        'frmCreaHorasDobleCotizacion
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(466, 324)
        Me.Controls.Add(Me.bCancelar)
        Me.Controls.Add(Me.bAceptar)
        Me.Controls.Add(Me.Frame1)
        Me.Name = "frmCreaHorasDobleCotizacion"
        Me.Text = "frmCreaHorasDobleCotizacion"
        Me.Frame1.ResumeLayout(False)
        Me.Frame1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents bCancelar As Solmicro.Expertis.Engine.UI.Button
    Friend WithEvents bAceptar As Solmicro.Expertis.Engine.UI.Button
    Friend WithEvents Frame1 As Solmicro.Expertis.Engine.UI.Frame
    Friend WithEvents Label5 As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents Label3 As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents Fecha1 As Solmicro.Expertis.Engine.UI.CalendarBox
    Friend WithEvents advObra As Solmicro.Expertis.Engine.UI.AdvSearch
    Friend WithEvents Label2 As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents Label1 As Solmicro.Expertis.Engine.UI.Label
    Friend WithEvents advPersona As Solmicro.Expertis.Engine.UI.AdvSearch
    Friend WithEvents txtHoras As Solmicro.Expertis.Engine.UI.TextBox
End Class
