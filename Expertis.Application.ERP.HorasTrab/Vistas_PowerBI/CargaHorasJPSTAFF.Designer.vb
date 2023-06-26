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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(CargaHorasJPSTAFF))
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
        Me.bNota = New Solmicro.Expertis.Engine.UI.Button
        Me.Frame5 = New Solmicro.Expertis.Engine.UI.Frame
        Me.bCreaHoras = New Solmicro.Expertis.Engine.UI.Button
        Me.Button1 = New Solmicro.Expertis.Engine.UI.Button
        Me.CD = New System.Windows.Forms.OpenFileDialog
        Me.Frame6 = New Solmicro.Expertis.Engine.UI.Frame
        Me.bA3 = New Solmicro.Expertis.Engine.UI.Button
        Me.bIDGET = New Solmicro.Expertis.Engine.UI.Button
        Me.Frame1.SuspendLayout()
        Me.Frame2.SuspendLayout()
        Me.Frame3.SuspendLayout()
        Me.Frame4.SuspendLayout()
        Me.Frame5.SuspendLayout()
        Me.Frame6.SuspendLayout()
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
        Me.btnAceptar.Location = New System.Drawing.Point(142, 43)
        Me.btnAceptar.Name = "btnAceptar"
        Me.btnAceptar.Size = New System.Drawing.Size(120, 38)
        Me.btnAceptar.TabIndex = 15
        Me.btnAceptar.Text = "Aceptar"
        '
        'cmdUbicacion
        '
        Me.cmdUbicacion.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdUbicacion.Icon = CType(resources.GetObject("cmdUbicacion.Icon"), System.Drawing.Icon)
        Me.cmdUbicacion.Location = New System.Drawing.Point(9, 43)
        Me.cmdUbicacion.Name = "cmdUbicacion"
        Me.cmdUbicacion.Size = New System.Drawing.Size(120, 38)
        Me.cmdUbicacion.TabIndex = 14
        Me.cmdUbicacion.Text = "Buscar"
        '
        'lblRuta
        '
        Me.lblRuta.AutoSize = False
        Me.lblRuta.Location = New System.Drawing.Point(69, 179)
        Me.lblRuta.Name = "lblRuta"
        Me.lblRuta.Size = New System.Drawing.Size(440, 38)
        Me.lblRuta.TabIndex = 11
        Me.lblRuta.Text = "Ruta"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(12, 179)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(42, 13)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "Ruta: "
        '
        'LProgreso
        '
        Me.LProgreso.Location = New System.Drawing.Point(69, 149)
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
        Me.Label2.Size = New System.Drawing.Size(441, 22)
        Me.Label2.TabIndex = 17
        Me.Label2.Text = "Carga de Horas Jefe de Produccion y Staff por Obra"
        '
        'bHorasOficina
        '
        Me.bHorasOficina.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bHorasOficina.Icon = CType(resources.GetObject("bHorasOficina.Icon"), System.Drawing.Icon)
        Me.bHorasOficina.Location = New System.Drawing.Point(31, 44)
        Me.bHorasOficina.Name = "bHorasOficina"
        Me.bHorasOficina.Size = New System.Drawing.Size(137, 38)
        Me.bHorasOficina.TabIndex = 18
        Me.bHorasOficina.Text = "Crear horas a gente de oficina"
        '
        'bAñadirHorasPersona
        '
        Me.bAñadirHorasPersona.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bAñadirHorasPersona.Icon = CType(resources.GetObject("bAñadirHorasPersona.Icon"), System.Drawing.Icon)
        Me.bAñadirHorasPersona.Location = New System.Drawing.Point(22, 43)
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
        Me.Frame1.Location = New System.Drawing.Point(63, 220)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Size = New System.Drawing.Size(288, 119)
        Me.Frame1.TabIndex = 21
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "Carga Jefe de producción y Staff"
        '
        'Frame2
        '
        Me.Frame2.Controls.Add(Me.bHorasOficina)
        Me.Frame2.Location = New System.Drawing.Point(357, 220)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.Size = New System.Drawing.Size(218, 119)
        Me.Frame2.TabIndex = 22
        Me.Frame2.TabStop = False
        Me.Frame2.Text = "Carga Horas Oficina"
        '
        'Frame3
        '
        Me.Frame3.Controls.Add(Me.bAñadirHorasPersona)
        Me.Frame3.Location = New System.Drawing.Point(581, 221)
        Me.Frame3.Name = "Frame3"
        Me.Frame3.Size = New System.Drawing.Size(189, 118)
        Me.Frame3.TabIndex = 23
        Me.Frame3.TabStop = False
        Me.Frame3.Text = "Carga Horas por Persona"
        '
        'Frame4
        '
        Me.Frame4.Controls.Add(Me.bBorrarExcel)
        Me.Frame4.Controls.Add(Me.bBorrarOperarioObraFecha)
        Me.Frame4.Location = New System.Drawing.Point(66, 356)
        Me.Frame4.Name = "Frame4"
        Me.Frame4.Size = New System.Drawing.Size(417, 131)
        Me.Frame4.TabIndex = 24
        Me.Frame4.TabStop = False
        Me.Frame4.Text = "Borrar Horas"
        '
        'bNota
        '
        Me.bNota.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bNota.Icon = CType(resources.GetObject("bNota.Icon"), System.Drawing.Icon)
        Me.bNota.Location = New System.Drawing.Point(763, 87)
        Me.bNota.Name = "bNota"
        Me.bNota.Size = New System.Drawing.Size(55, 33)
        Me.bNota.TabIndex = 25
        Me.bNota.Text = "Nota"
        '
        'Frame5
        '
        Me.Frame5.Controls.Add(Me.bCreaHoras)
        Me.Frame5.Controls.Add(Me.Button1)
        Me.Frame5.Location = New System.Drawing.Point(489, 357)
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
        Me.Frame6.Location = New System.Drawing.Point(793, 221)
        Me.Frame6.Name = "Frame6"
        Me.Frame6.Size = New System.Drawing.Size(189, 118)
        Me.Frame6.TabIndex = 27
        Me.Frame6.TabStop = False
        Me.Frame6.Text = "Unifica A3"
        '
        'bA3
        '
        Me.bA3.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bA3.Icon = CType(resources.GetObject("bA3.Icon"), System.Drawing.Icon)
        Me.bA3.Location = New System.Drawing.Point(22, 43)
        Me.bA3.Name = "bA3"
        Me.bA3.Size = New System.Drawing.Size(144, 38)
        Me.bA3.TabIndex = 19
        Me.bA3.Text = "Selecciona Ficheros A3"
        '
        'bIDGET
        '
        Me.bIDGET.Location = New System.Drawing.Point(66, 507)
        Me.bIDGET.Name = "bIDGET"
        Me.bIDGET.Size = New System.Drawing.Size(189, 23)
        Me.bIDGET.TabIndex = 28
        Me.bIDGET.Text = "Actualiza IDGET"
        Me.bIDGET.Visible = False
        '
        'CargaHorasJPSTAFF
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1073, 542)
        Me.Controls.Add(Me.bIDGET)
        Me.Controls.Add(Me.Frame6)
        Me.Controls.Add(Me.Frame5)
        Me.Controls.Add(Me.bNota)
        Me.Controls.Add(Me.Frame4)
        Me.Controls.Add(Me.Frame3)
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.lblRuta)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.LProgreso)
        Me.Controls.Add(Me.PvProgreso)
        Me.Name = "CargaHorasJPSTAFF"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "CargaHorasJPSTAFF"
        Me.Frame1.ResumeLayout(False)
        Me.Frame2.ResumeLayout(False)
        Me.Frame3.ResumeLayout(False)
        Me.Frame4.ResumeLayout(False)
        Me.Frame5.ResumeLayout(False)
        Me.Frame6.ResumeLayout(False)
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
    Friend WithEvents bNota As Solmicro.Expertis.Engine.UI.Button
    Friend WithEvents Frame5 As Solmicro.Expertis.Engine.UI.Frame
    Friend WithEvents bCreaHoras As Solmicro.Expertis.Engine.UI.Button
    Friend WithEvents Button1 As Solmicro.Expertis.Engine.UI.Button
    Friend WithEvents CD As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Frame6 As Solmicro.Expertis.Engine.UI.Frame
    Friend WithEvents bA3 As Solmicro.Expertis.Engine.UI.Button
    Friend WithEvents bIDGET As Solmicro.Expertis.Engine.UI.Button
End Class
