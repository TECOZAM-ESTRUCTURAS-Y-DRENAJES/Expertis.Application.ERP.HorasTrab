Imports System.Windows.Forms
Imports OfficeOpenXml

Public Class frmRegularizacionEnero

    Private Sub frmRegularizacionEnero_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        CargaCombo()
    End Sub

    Public Sub CargaCombo()
        Dim dtAnios As New DataTable
        dtAnios.Columns.Add("Año", GetType(Integer)) ' Asegurarse de que es tipo Integer

        Dim anio As Integer = Today.Year
        For i As Integer = anio To anio - 15 Step -1
            dtAnios.Rows.Add(i)
        Next
        dtAnios.AcceptChanges()

        ' Asignar el DataTable al ComboBox
        cmbAnio.DataSource = dtAnios
        cmbAnio.DisplayMember = "Año" ' Lo que se muestra en el combo
        cmbAnio.ValueMember = "Año" ' El valor real de cada opción
    End Sub


    Private Sub cmbAnio_SelectedValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbAnio.SelectedValueChanged
        Try
            Dim anio As Integer = cmbAnio.SelectedValue
            Dim mes As Integer = 1

            ' Contar los días hábiles (lunes a viernes) en cada rango
            Dim diasHabiles1_20 As Integer = ContarDiasHabiles(anio, mes, 1, 20)
            Dim diasHabiles1_31 As Integer = ContarDiasHabiles(anio, mes, 1, 31)

            ' Calcular porcentaje
            Dim Porcentaje1 As Double = (diasHabiles1_20 / diasHabiles1_31)
            Dim Porcentaje2 As Double = 1 - Porcentaje1

            txtPorcentaje1.Text = Math.Round(Porcentaje1, 4).ToString("0.0000")
            txtPorcentaje2.Text = Math.Round(Porcentaje2, 4).ToString("0.0000")
        Catch ex As Exception

        End Try

    End Sub

    Function ContarDiasHabiles(ByVal anio As Integer, ByVal mes As Integer, ByVal diaInicio As Integer, ByVal diaFin As Integer) As Integer
        Dim contador As Integer = 0
        For dia As Integer = diaInicio To diaFin
            Dim fecha As New DateTime(anio, mes, dia)
            ' Si no es sábado (6) ni domingo (0), es día hábil
            If fecha.DayOfWeek <> DayOfWeek.Saturday AndAlso fecha.DayOfWeek <> DayOfWeek.Sunday Then
                contador += 1
            End If
        Next
        Return contador
    End Function

    Private Sub bImportar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bImportar.Click
        ' Crear un cuadro de diálogo para seleccionar archivos
        Dim openFileDialog As New Windows.Forms.OpenFileDialog()

        ' Configurar las propiedades del cuadro de diálogo
        openFileDialog.Title = "Seleccionar archivo Excel"
        openFileDialog.Filter = "Archivos de Excel (*.xlsx;*.xls)|*.xlsx;*.xls|Todos los archivos (*.*)|*.*"
        openFileDialog.InitialDirectory = "C:\" ' Puedes cambiar la ruta inicial

        ' Mostrar el cuadro de diálogo y verificar si el usuario seleccionó un archivo
        If openFileDialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
            ' Mostrar la ruta del archivo en la etiqueta lblRuta
            lblRuta.Text = openFileDialog.FileName
        End If
    End Sub

    Public Function ObtenerDatosExcelCabecera(ByVal ruta As String, ByVal hoja As String, ByVal rango As String) As DataTable
        Dim MyConnection As System.Data.OleDb.OleDbConnection
        Dim DtSet As System.Data.DataSet
        Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
        MyConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & ruta & "';Extended Properties='Excel 8.0;HDR=YES'")
        MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [" & hoja & "$" & rango & "]", MyConnection)
        'MyCommand.TableMappings.Add("Table", "Net-informations.com")
        DtSet = New System.Data.DataSet
        MyCommand.Fill(DtSet)
        Dim dt As DataTable = DtSet.Tables(0)
        MyConnection.Close()

        Return dt

    End Function

    Private Sub bGenerar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bGenerar.Click
        Dim ruta As String = lblRuta.Text
        Dim hoja As String = "HORAS"
        Dim rango As String = "A1:O100000"
        Dim dtPrincipal As DataTable = ObtenerDatosExcelCabecera(ruta, hoja, rango)

        Dim dtResultado As New DataTable
        dtResultado = TratamientoOperarios(dtPrincipal)

        ExportarFichero(dtResultado)
    End Sub

    Public Function TratamientoOperarios(ByVal dtOperarios As DataTable) As DataTable

        ' Filtrar los registros donde IDCategoriaProfesionalSCCP (columna 6) sea 2 o 3
        Dim dtFiltrado As New DataTable()
        Dim dtResultado As DataTable = dtOperarios.Clone()
        dtFiltrado = dtOperarios.Clone() ' Copiar estructura

        ' 1. Filtrar categoría 2 o 3
        For Each row As DataRow In dtOperarios.Rows
            Dim categoria As String = row(5)
            Dim horas As Double = row(10)
            If categoria = "2" Or categoria = "3" Or horas > 0 Then
                dtFiltrado.ImportRow(row)
            End If
        Next

        ' 2. Agrupar por IDOperario y NObra, solo fechas del 1 al 20
        Dim grupos As New Dictionary(Of String, List(Of DataRow))()

        For Each row As DataRow In dtFiltrado.Rows
            Dim fecha As Date = CDate(row("FechaInicio"))
            If fecha.Day >= 1 AndAlso fecha.Day <= 20 Then
                Dim clave As String = row("IDOperario").ToString().Trim() & "|" & row("NObra").ToString().Trim()
                If Not grupos.ContainsKey(clave) Then
                    grupos(clave) = New List(Of DataRow)()
                End If
                grupos(clave).Add(row)
            End If
        Next

        ' 3. Recorrer cada grupo y generar la fila de resultado
        For Each kvp In grupos
            Dim listaFilas As List(Of DataRow) = kvp.Value
            Dim filaBase As DataRow = listaFilas(0)
            Dim nuevaFila As DataRow = dtResultado.NewRow()
            nuevaFila.ItemArray = filaBase.ItemArray.Clone()

            ' Reemplazar el valor en la fila nueva
            nuevaFila("FechaInicio") = "01/01/" & cmbAnio.SelectedValue
            nuevaFila("Horas") = "0"
            nuevaFila("IDHora") = "HA"
            nuevaFila("HorasAdministrativas") = getHorasPorOperarioYObra(nuevaFila("IDOperario"), nuevaFila("NObra"), dtOperarios)
            nuevaFila("HorasBaja") = 0
            nuevaFila("Turno") = 0

            dtResultado.Rows.Add(nuevaFila)
        Next

        Return dtResultado
    End Function

    Public Function getHorasPorOperarioYObra(ByVal IDOperario As String, ByVal NObra As String, ByVal dtOperarios As DataTable) As Double
        Dim sumaHoras21Al31 As Double = 0
        sumaHoras21Al31 = getHorasPorOperarioYObra2131(IDOperario, dtOperarios)
        Dim sumaHoras1Al20 As Double = 0
        sumaHoras1Al20 = getHorasPorOperarioYObra120(IDOperario, dtOperarios)

        Dim sumaHoras1Al20UnaObra As Double = 0
        sumaHoras1Al20UnaObra = getHorasPorOperarioYObra120UnaObra(IDOperario, NObra, dtOperarios)

        Dim sumaHoras1Al20Obras As Double = 0
        sumaHoras1Al20Obras = getHorasPorOperarioYObra120Obras(IDOperario, dtOperarios)

        Dim cuenta As Double = 0
        cuenta = ((sumaHoras21Al31 * (CDbl(txtPorcentaje1.Text) / CDbl(txtPorcentaje2.Text)) - sumaHoras1Al20) * (sumaHoras1Al20UnaObra / sumaHoras1Al20Obras))

        If cuenta < 0 OrElse Double.IsNaN(cuenta) OrElse Double.IsInfinity(cuenta) Then
            Return 0
        Else
            Return CInt(Math.Ceiling(cuenta))
        End If
    End Function
    Public Function getHorasPorOperarioYObra120Obras(ByVal IDOperario As String, _
                                                        ByVal dtOperarios As DataTable) As Double

        Dim sumaHoras120 As Double = 0

        ' Usando LINQ sobre el DataTable
        sumaHoras120 = dtOperarios.AsEnumerable(). _
            Where(Function(dr) dr.Field(Of String)("IDOperario") = IDOperario _
                      AndAlso dr.Field(Of Date)("FechaInicio").Day >= 1 _
                      AndAlso dr.Field(Of Date)("FechaInicio").Day <= 20). _
            Sum(Function(dr) dr.Field(Of Double)("Horas"))

        Return sumaHoras120
    End Function

    Public Function getHorasPorOperarioYObra120UnaObra(ByVal IDOperario As String, _
                                                       ByVal NObra As String, _
                                                        ByVal dtOperarios As DataTable) As Double

        Dim sumaHoras120 As Double = 0

        ' Usando LINQ sobre el DataTable
        sumaHoras120 = dtOperarios.AsEnumerable(). _
            Where(Function(dr) dr.Field(Of String)("IDOperario") = IDOperario _
                       AndAlso dr.Field(Of String)("NObra") = NObra _
                      AndAlso dr.Field(Of Date)("FechaInicio").Day >= 1 _
                      AndAlso dr.Field(Of Date)("FechaInicio").Day <= 20). _
            Sum(Function(dr) dr.Field(Of Double)("Horas"))

        Return sumaHoras120
    End Function

    Public Function getHorasPorOperarioYObra2131(ByVal IDOperario As String, _
                                         ByVal dtOperarios As DataTable) As Double

        Dim sumaHoras21Al31 As Double = 0

        ' Usando LINQ sobre el DataTable
        sumaHoras21Al31 = dtOperarios.AsEnumerable(). _
            Where(Function(dr) dr.Field(Of String)("IDOperario") = IDOperario _
                      AndAlso dr.Field(Of Date)("FechaInicio").Day >= 21 _
                      AndAlso dr.Field(Of Date)("FechaInicio").Day <= 31). _
            Sum(Function(dr) dr.Field(Of Double)("Horas"))

        Return sumaHoras21Al31
    End Function

    Public Function getHorasPorOperarioYObra120(ByVal IDOperario As String, _
                                         ByVal dtOperarios As DataTable) As Double

        Dim sumaHoras1Al20 As Double = 0

        ' Usando LINQ sobre el DataTable
        sumaHoras1Al20 = dtOperarios.AsEnumerable(). _
            Where(Function(dr) dr.Field(Of String)("IDOperario") = IDOperario _
                      AndAlso dr.Field(Of Date)("FechaInicio").Day >= 1 _
                      AndAlso dr.Field(Of Date)("FechaInicio").Day <= 20). _
            Sum(Function(dr) dr.Field(Of Double)("Horas"))

        Return sumaHoras1Al20
    End Function

    Public Sub ExportarFichero(ByVal dtFinal As DataTable)
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial

        Dim saveFileDialog1 As New SaveFileDialog()
        saveFileDialog1.Filter = "Archivos de Excel|*.xlsx|Todos los archivos|*.*"
        saveFileDialog1.Title = "Guardar archivo"

        ' Mostrar el cuadro de diálogo y verificar si el usuario hizo clic en "Guardar"
        If saveFileDialog1.ShowDialog() = DialogResult.OK Then
            ' Obtener la ruta seleccionada por el usuario
            Dim rutaArchivo As String = saveFileDialog1.FileName

            Using package As New ExcelPackage()

                Dim nombreHoja As String = ("1")

                ' Crear una hoja de cálculo y obtener una referencia a ella.
                Dim worksheet = package.Workbook.Worksheets.Add(nombreHoja)

                ' Copiar los datos de la DataTable a la hoja de cálculo.
                worksheet.Cells("A1").LoadFromDataTable(dtFinal, True)

                Dim totalFilas As Integer = dtFinal.Rows.Count + 1 ' +1 por los encabezados
                worksheet.Cells("H2:H" & totalFilas).Style.Numberformat.Format = "dd/mm/yyyy"

                ' Guardar el paquete de Excel en la ruta seleccionada
                Dim fileInfo As New IO.FileInfo(rutaArchivo)
                package.SaveAs(fileInfo)
            End Using
        End If
        MessageBox.Show("Fichero guardado correctamente", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
End Class