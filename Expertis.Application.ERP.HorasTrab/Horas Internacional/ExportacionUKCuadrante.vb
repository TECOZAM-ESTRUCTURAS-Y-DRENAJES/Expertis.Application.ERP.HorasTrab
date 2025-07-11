﻿Imports Solmicro.Expertis.Engine
Imports System.Windows.Forms
Imports OfficeOpenXml
Imports OfficeOpenXml.Style
Imports System.Drawing
Imports System.IO
Imports System.Collections.Generic
Imports System.Globalization

Public Class ExportacionUKCuadrante
    Public tablaDatos As String
    Public tipoExportacion As String

    Public Sub generaExcel()

        Dim frm As New frmInformeFecha
        frm.ShowDialog()
        Dim Fecha1 As String
        Dim Fecha2 As String

        Fecha1 = frm.fecha1
        Fecha2 = frm.fecha2

        If frm.blEstado = False Then
            Exit Sub
        End If

        '1. OBTENGO INFO BASICA DE LAS PERSONAS
        Dim dtPersonas As New DataTable
        dtPersonas = getTablaPersonas(Fecha1)


        '2. LE DOY LA 1ª FORMA A LA TABLA
        Dim dtFinal As New DataTable
        FormaTablaSalidaUK(dtFinal, Fecha1, Fecha2)
        setPrimerCambioForma(dtFinal, dtPersonas)
        setSegundoCambioForma(dtFinal, Fecha1, Fecha2)

        ExportarFichero(dtFinal, Fecha1, Fecha2)
    End Sub

    Public Function getTablaPersonas(ByVal fecha1 As String) As DataTable
        Dim strWhere As String = "(Fecha_Baja is null or Fecha_Baja>='" & fecha1 & "') order by FechaAlta asc"
        Return New BE.DataEngine().Filter("frmMntoOperario", "IDOperario, Diccionario, DescOperario,IDDepartamento As Compañia, FechaAlta, Fecha_Baja ", strWhere)
    End Function

    Public Sub FormaTablaSalidaUK(ByRef dtFinal As DataTable, ByVal fecha1 As String, ByVal fecha2 As String)
        ' Las columnas fijas
        dtFinal.Columns.Add("IDOPERARIO", GetType(String))
        dtFinal.Columns.Add("DICCIONARIO", GetType(String))
        dtFinal.Columns.Add("NOMBRE", GetType(String))
        dtFinal.Columns.Add("COMPAÑIA", GetType(String))
        dtFinal.Columns.Add("START DAY", GetType(String))
        dtFinal.Columns.Add("FINISH DAY", GetType(String))

        ' Calcular el rango de fechas
        Dim fechaInicio As DateTime = Convert.ToDateTime(fecha1)
        Dim fechaFin As DateTime = Convert.ToDateTime(fecha2)
        Dim fechaActual As DateTime = fechaInicio

        While fechaActual <= fechaFin
            Dim nombreDia As String = fechaActual.ToString("dd") ' Nombre del día en formato dd
            Dim nombreMes As String = fechaActual.ToString("MM")
            Dim nombreAño As String = fechaActual.ToString("yy")
            dtFinal.Columns.Add(nombreDia & "/" & nombreMes & "/" & nombreAño & "-PROD", GetType(String))
            dtFinal.Columns.Add(nombreDia & "/" & nombreMes & "/" & nombreAño & "-NOPROD", GetType(String))
            fechaActual = fechaActual.AddDays(1) ' Incrementar la fecha en 1 día
        End While
    End Sub

    Public Sub setPrimerCambioForma(ByRef dtFinal As DataTable, ByRef dtPersonas As DataTable)
        ' Iterar a través de cada fila en dtPersonas
        For Each dr As DataRow In dtPersonas.Rows
            ' Crear una nueva fila en dtFinal
            Dim newRow As DataRow = dtFinal.NewRow()

            ' Asignar los valores de la fila actual de dtPersonas a la nueva fila de dtFinal
            newRow("IDOPERARIO") = dr("IDOperario")
            newRow("DICCIONARIO") = dr("Diccionario")
            newRow("NOMBRE") = dr("DescOperario").ToString.ToUpper
            newRow("COMPAÑIA") = dr("Compañia")
            ' Manejar la conversión y formato de la fecha
            Dim fechaAlta As Object = dr("FechaAlta")
            If IsDBNull(fechaAlta) OrElse Not DateTime.TryParse(fechaAlta.ToString(), New DateTime()) Then
                newRow("START DAY") = DBNull.Value
            Else
                Dim fecha As DateTime = Convert.ToDateTime(fechaAlta)
                newRow("START DAY") = fecha.ToString("dd/MM/yyyy") ' Formatear solo la fecha
            End If
            Dim fechaBaja As Object = dr("Fecha_Baja")
            If IsDBNull(fechaBaja) OrElse Not DateTime.TryParse(fechaBaja.ToString(), New DateTime()) Then
                newRow("FINISH DAY") = DBNull.Value
            Else
                Dim fecha As DateTime = Convert.ToDateTime(fechaBaja)
                newRow("FINISH DAY") = fecha.ToString("dd/MM/yyyy") ' Formatear solo la fecha
            End If

            dtFinal.Rows.Add(newRow)
        Next
    End Sub

    Public Sub setSegundoCambioForma(ByRef dtFinal As DataTable, ByVal fecha1 As String, ByVal fecha2 As String)
        ' Convertir las fechas de entrada a DateTime
        Dim fechaInicio As DateTime = Convert.ToDateTime(fecha1)
        Dim fechaFin As DateTime = Convert.ToDateTime(fecha2)

        ' Iterar a través de cada fila en la tabla final
        For Each dr As DataRow In dtFinal.Rows
            ' Iterar sobre el rango de fechas
            Dim fechaActual As DateTime = fechaInicio
            While fechaActual <= fechaFin
                ' Obtener el día, el mes y el año en formato "dd/MM/yy"
                Dim dia As String = fechaActual.ToString("dd")
                Dim mes As String = fechaActual.ToString("MM")
                Dim año As String = fechaActual.ToString("yy")

                ' Verificar las columnas PROD y NOPROD para cada día con el formato "dd/MM/yy"
                Dim colProd As String = dia & "/" & mes & "/" & año & "-PROD"
                Dim colNoProd As String = dia & "/" & mes & "/" & año & "-NOPROD"

                ' Si las columnas existen en la tabla
                If dtFinal.Columns.Contains(colProd) AndAlso dtFinal.Columns.Contains(colNoProd) Then
                    ' Crear un filtro para obtener la información desde la base de datos
                    Dim dt As New DataTable
                    Dim f As New Filter
                    f.Add("IDOperario", FilterOperator.Equal, dr("IDOPERARIO"))
                    f.Add("FechaParte", FilterOperator.Equal, fechaActual) ' Usar la fecha actual completa

                    ' Obtener los datos desde la base de datos
                    dt = New BE.DataEngine().Filter(tablaDatos, f)

                    ' Asignar los valores a las columnas correspondientes
                    If dt.Rows.Count > 0 Then
                        ' Obtener el valor de "HorasProductivas", "HorasNoProductivas" e "IDCausa" desde la fila
                        Dim horasProductivas As Object = 0
                        Dim horasNoProductivas As Object = 0
                        Dim idCausa As String = dt.Rows(0)("IDCausa").ToString()

                        For Each row As DataRow In dt.Rows
                            If Not IsDBNull(row("HorasProductivas")) Then
                                horasProductivas += Convert.ToDouble(row("HorasProductivas"))
                            End If

                            If Not IsDBNull(row("HorasNoProductivas")) Then
                                horasNoProductivas += Convert.ToDouble(row("HorasNoProductivas"))
                            End If
                        Next

                        ' Verificar si horasProductivas es un número y asignar
                        If Not String.IsNullOrEmpty(horasProductivas.ToString()) Then ' Si la longitud es distinta de 0
                            dr(colProd) = horasProductivas
                        End If

                        ' Verificar si horasNoProductivas es un número y asignar
                        If Not String.IsNullOrEmpty(horasNoProductivas.ToString()) Then ' Si la longitud es distinta de 0
                            dr(colNoProd) = horasNoProductivas

                        End If

                        ' Si IDCausa no nulo, sobreescribir datos
                        If Not String.IsNullOrEmpty(idCausa) Then
                            dr(colProd) = idCausa ' Asignar el valor de IDCausa
                            dr(colNoProd) = DBNull.Value
                        End If
                    End If


                End If

                ' Incrementar la fecha actual en 1 día
                fechaActual = fechaActual.AddDays(1)
            End While
        Next
    End Sub

    Public Sub ExportarFichero(ByVal dtFinal As DataTable, ByVal fecha1 As String, ByVal fecha2 As String)
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial

        Dim saveFileDialog1 As New SaveFileDialog()
        saveFileDialog1.Filter = "Archivos de Excel|*.xlsx|Todos los archivos|*.*"
        saveFileDialog1.Title = "Guardar archivo"

        ' Mostrar el cuadro de diálogo y verificar si el usuario hizo clic en "Guardar"
        If saveFileDialog1.ShowDialog() = DialogResult.OK Then
            Dim rutaArchivo As String = saveFileDialog1.FileName
            Using package As New ExcelPackage()
                ' Crear hoja con todos los empleados
                CrearHojaEmpleados(package, dtFinal)

                ' Obtener los ID de las obras
                Dim dtNObras As DataTable = ObtenerIDObras(dtFinal, fecha1, fecha2)

                ' Crear una hoja por cada obra
                'CrearHojasPorObra(package, dtFinal, dtIDObras, fecha1, fecha2)
                CrearHojasPorObra(package, dtFinal, dtNObras, fecha1, fecha2)

                ' Guardar el archivo
                GuardarArchivoExcel(package, rutaArchivo)
            End Using
            MessageBox.Show("Fichero guardado correctamente", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    ' Método para crear la hoja de todos los empleados
    Private Sub CrearHojaEmpleados(ByVal package As ExcelPackage, ByVal dtFinal As DataTable)
        Dim worksheet = package.Workbook.Worksheets.Add("UK")
        worksheet.Cells("A1").LoadFromDataTable(dtFinal, True)
        ConvertirCeldasANumeros(worksheet)
        CalcularTotales(dtFinal, worksheet)
        GestionarEstilos(worksheet)
        AgregarComentarios(worksheet)
    End Sub
    Private Sub AgregarComentarios(ByVal hoja As ExcelWorksheet)
        Dim filaActual As Integer = 2 ' Comienza en fila 2 (fila 1 son encabezados)
        Dim columnaInicio As Integer = 7 ' Columna G = 7

        While hoja.Cells(filaActual, 1).Value IsNot Nothing
            Dim idOperario As String = hoja.Cells(filaActual, 1).Text

            Dim columnaActual As Integer = columnaInicio
            While hoja.Cells(1, columnaActual).Value IsNot Nothing
                Dim cabecera As String = hoja.Cells(1, columnaActual).Text

                If cabecera.Length >= 8 Then
                    Dim fechaTexto As String = cabecera.Substring(0, 8) ' Ej: "01/06/25"
                    Dim fechaParsed As Date

                    If Date.TryParseExact(fechaTexto, "dd/MM/yy", Globalization.CultureInfo.InvariantCulture, Globalization.DateTimeStyles.None, fechaParsed) Then
                        Dim resultadoBD As String = BuscarInfoEnBD(idOperario, fechaParsed)

                        If Not String.IsNullOrEmpty(resultadoBD) Then
                            hoja.Cells(filaActual, columnaActual).AddComment(resultadoBD)
                            columnaActual += 1
                        End If
                    End If
                End If

                columnaActual += 1
            End While

            filaActual += 1
        End While
    End Sub
    Public Function BuscarInfoEnBD(ByVal IDOperario As String, ByVal fechaParsed As String) As String
        Dim dt As New DataTable
        Dim f As New Filter
        f.Add("IDOperario", FilterOperator.Equal, IDOperario)
        f.Add("FechaParte", FilterOperator.Equal, fechaParsed)

        dt = New BE.DataEngine().Filter("frmMntoHorasInternacionalTecozam", f)
        If dt.Rows.Count > 0 Then
            If dt.Rows(0)("Comentarios").ToString.Length > 0 Then
                Return dt.Rows(0)("Comentarios").ToString
            Else
                Return ""
            End If
        Else
            Return ""
        End If

    End Function

    ' Método para obtener los IDObra únicos
    Private Function ObtenerIDObras(ByVal dtFinal As DataTable, ByVal fecha1 As DateTime, ByVal fecha2 As DateTime) As DataTable
        Dim dtIDObras As New DataTable()
        dtIDObras.Columns.Add("obra", GetType(String))

        Dim obrasList As New List(Of String)()

        For Each filaOperario As DataRow In dtFinal.Rows
            Dim dt As DataTable
            Dim f As New Filter()
            f.Add("IDOperario", FilterOperator.Equal, filaOperario("IDOperario"))
            f.Add("FechaParte", FilterOperator.GreaterThanOrEqual, fecha1)
            f.Add("FechaParte", FilterOperator.LessThanOrEqual, fecha2)

            ' Filtrar la tabla y obtener los datos
            dt = New BE.DataEngine().Filter("frmMntoHorasInternacionalTecozam", f)

            ' Iterar sobre las filas del DataTable dt para obtener los IDObra
            For Each fila As DataRow In dt.Rows
                If fila.Table.Columns.Contains("NObra") Then
                    Dim idObra As String = fila("NObra").ToString()

                    ' Si la lista no contiene la obra (evitar duplicados)
                    If Not obrasList.Contains(idObra) Then
                        obrasList.Add(idObra)
                    End If
                End If
            Next
        Next

        'convertir lista a datatable para devolver
        For Each obra As String In obrasList
            Dim nuevaFila As DataRow = dtIDObras.NewRow()
            nuevaFila("obra") = obra
            dtIDObras.Rows.Add(nuevaFila)
        Next

        Return dtIDObras
    End Function

    Private Sub CrearHojasPorObra(ByVal package As ExcelPackage, ByVal dtFinal As DataTable, ByVal dtNObras As DataTable, ByVal fecha1 As DateTime, ByVal fecha2 As DateTime)
        'evaluar la situacion de cada obra (NObra en dtIDObras
        For Each obra As DataRow In dtNObras.Rows
            Dim NObra As String = obra("obra")

            'copiar estructura de dtFinal para la hoja
            Dim dtObra As DataTable = dtFinal.Clone()

            'evaluar situacion de cada operario en NObra
            For Each operario As DataRow In dtFinal.Rows()
                'obtner tabla con todas las horas para IDOperario en NObra
                Dim f As New Filter()
                f.Add("IDOperario", FilterOperator.Equal, operario("IDOperario"))
                f.Add("NObra", FilterOperator.Equal, NObra)
                f.Add("FechaParte", FilterOperator.GreaterThanOrEqual, fecha1)
                f.Add("FechaParte", FilterOperator.LessThanOrEqual, fecha2)

                Dim dtHorasOperario As DataTable
                dtHorasOperario = New BE.DataEngine().Filter("frmMntoHorasInternacionalTecozam", f, "FechaParte, HorasProductivas, HorasNoProductivas, IDCausa")

                If dtHorasOperario.Rows.Count > 0 Then

                    'preparar fila nueva a insertar
                    dtObra.ImportRow(operario)
                    Dim nuevaFila As DataRow = dtObra.Rows(dtObra.Rows.Count - 1)
                    LimpiarFila(nuevaFila)

                    'volcar en nueva fila de dtObra la informacion obtenida en dtHorasTrabajador
                    For Each horasOperario As DataRow In dtHorasOperario.Rows
                        If (Not String.IsNullOrEmpty(horasOperario("HorasProductivas").ToString()) AndAlso Not String.IsNullOrEmpty(horasOperario("HorasNoProductivas").ToString())) Or Not IsDBNull(horasOperario("IDCausa")) Then

                            'obtener la fecha a evaluar
                            Dim fechaOperario As DateTime = Convert.ToDateTime(horasOperario("FechaParte")).Date
                            Dim fechaProd As String = fechaOperario.ToString("dd/MM/yy") & "-PROD"
                            Dim fechaNoProd As String = fechaOperario.ToString("dd/MM/yy") & "-NOPROD"

                            'buscar en dtObra la columna que coincide con la fecha y volcar datos
                            For Each col As DataColumn In nuevaFila.Table.Columns
                                If col.ColumnName.Contains(fechaProd) Then
                                    nuevaFila(col) = horasOperario("HorasProductivas")

                                    'volcar IDcausa si corresponde
                                    If Not IsDBNull(horasOperario("IDCausa")) Then
                                        nuevaFila(col) = horasOperario("IDCausa")
                                    End If
                                End If
                                If col.ColumnName.Contains(fechaNoProd) Then
                                    nuevaFila(col) = horasOperario("HorasNoProductivas")
                                End If
                            Next
                        End If
                    Next
                End If
            Next

            'crear hoja NObra y volcar dtObra
            Dim worksheetObra = package.Workbook.Worksheets.Add(NObra)
            worksheetObra.Cells("A1").LoadFromDataTable(dtObra, True)

            ConvertirCeldasANumeros(worksheetObra)
            CalcularTotales(dtObra, worksheetObra)
            GestionarEstilos(worksheetObra)
        Next
    End Sub

    ' Método para crear una hoja de Excel por cada obra
    'Private Sub CrearHojasPorObra(ByVal package As ExcelPackage, ByVal dtFinal As DataTable, ByVal dtIDObras As DataTable, ByVal fecha1 As DateTime, ByVal fecha2 As DateTime)
    '    For Each filaIDObra As DataRow In dtIDObras.Rows
    '        Dim NObra As String = filaIDObra("obra").ToString()

    '        ' Clonar la estructura de dtFinal para dtObra
    '        Dim dtObra As DataTable = dtFinal.Clone()

    '        ' Filtrar operarios por IDObra
    '        For Each filaOperario As DataRow In dtFinal.Rows
    '            Dim dt As DataTable
    '            Dim f As New Filter()
    '            f.Add("IDOperario", FilterOperator.Equal, filaOperario("IDOperario"))
    '            f.Add("NObra", FilterOperator.Equal, NObra)
    '            f.Add("FechaParte", FilterOperator.GreaterThanOrEqual, fecha1)
    '            f.Add("FechaParte", FilterOperator.LessThanOrEqual, fecha2)

    '            ' Filtrar la tabla y obtener los datos
    '            dt = New BE.DataEngine().Filter("frmMntoHorasInternacionalTecozam", f, "NObra, FechaParte, HorasProductivas, HorasNoProductivas")

    '            dtObra.ImportRow(filaOperario)
    '            'LimpiarFila(dtObra) 'limpiar datos basura

    '            ' Iterar sobre las filas del DataTable dt para llenar el DataTable dtObra
    '            For Each fila As DataRow In dt.Rows
    '                If fila.Table.Columns.Contains("NObra") AndAlso fila("NObra").ToString() = NObra Then
    '                    If Not String.IsNullOrEmpty(fila("HorasProductivas").ToString()) AndAlso Not String.IsNullOrEmpty(fila("HorasNoProductivas").ToString()) Then
    '                        Dim fecha As Date = Convert.ToDateTime(fila("FechaParte")).Date 'obtener fecha que se esta evaluando

    '                        Dim prodCol As String = fecha.ToString("dd/MM/yy") & "-PROD"
    '                        Dim noprodCol As String = fecha.ToString("dd/MM/yy") & "-NOPROD"

    '                        'recorrer dtObra buscando coincidencias
    '                        For Each row As DataRow In dtObra.Rows
    '                            If dtObra.Columns.Contains(prodCol) Then
    '                                row(prodCol) = fila("HorasProductivas")
    '                            End If

    '                            If dtObra.Columns.Contains(noprodCol) Then
    '                                row(noprodCol) = fila("HorasNoProductivas")
    '                            End If
    '                        Next
    '                    End If
    '                End If
    '            Next
    '        Next

    '        ' Obtener la cabecera de la obra
    '        Dim dtObraCab As New DataTable
    '        Dim fObraCab As New Filter
    '        fObraCab.Add("NObra", FilterOperator.Equal, NObra)
    '        dtObraCab = New BE.DataEngine().Filter("tbObraCabecera", fObraCab)

    '        ' Crear una hoja de cálculo y obtener una referencia a ella.
    '        Dim worksheetObra = package.Workbook.Worksheets.Add(dtObraCab.Rows(0)("NObra"))
    '        worksheetObra.Cells("A1").LoadFromDataTable(dtObra, True)

    '        ConvertirCeldasANumeros(worksheetObra)
    '        CalcularTotales(dtObra, worksheetObra)
    '        GestionarEstilos(worksheetObra)
    '    Next
    'End Sub

    Private Sub LimpiarFila(ByRef row As DataRow)
        Dim buscado1 As String = "PROD"
        Dim buscado2 As String = "NOPROD"

        For Each col As DataColumn In row.Table.Columns
            If col.ColumnName.Contains(buscado1) Or col.ColumnName.Contains(buscado2) Then
                row(col) = DBNull.Value
            End If
        Next
    End Sub

    ' Método para guardar el archivo de Excel
    Private Sub GuardarArchivoExcel(ByVal package As ExcelPackage, ByVal rutaArchivo As String)
        Dim fileInfo As New IO.FileInfo(rutaArchivo)
        package.SaveAs(fileInfo)
    End Sub


    Private Sub ConvertirCeldasANumeros(ByVal worksheet As ExcelWorksheet)
        ' Iterar a través de las filas y columnas para convertir cadenas numéricas en números
        Dim lastRow As Integer = worksheet.Dimension.End.Row
        Dim lastCol As Integer = worksheet.Dimension.End.Column

        For row As Integer = 2 To lastRow ' Comenzar desde la fila 2
            For col As Integer = 7 To lastCol ' Comenzar desde la columna 7
                Dim cellValue As String = worksheet.Cells(row, col).Text
                Dim numericValue As Double

                ' Intentar convertir la cadena a un número
                If Double.TryParse(cellValue, numericValue) Then
                    worksheet.Cells(row, col).Value = numericValue ' Asignar el valor numérico
                End If
            Next
        Next
    End Sub

    Public Sub CalcularTotales(ByVal dtFinal As DataTable, ByVal worksheet As ExcelWorksheet)
        Dim lastRow As Integer = dtFinal.Rows.Count + 1 ' Contamos desde 1 y sumamos el encabezado
        Dim lastCol As Integer = dtFinal.Columns.Count ' Última columna del DataTable

        ' Definir las columnas para los totales
        Dim totalProdCol As Integer = lastCol + 1
        Dim totalNoProdCol As Integer = lastCol + 2
        Dim totalGeneralCol As Integer = lastCol + 3

        ' Agregar encabezados para los totales
        worksheet.Cells(1, totalProdCol).Value = "TOTAL PROD"
        worksheet.Cells(1, totalNoProdCol).Value = "TOTAL NOPROD"
        worksheet.Cells(1, totalGeneralCol).Value = "TOTAL GENERAL"

        ' Calcular totales por fila
        For i As Integer = 2 To lastRow ' Comenzar desde 2 para omitir el encabezado
            Dim totalProd As Double = 0
            Dim totalNoProd As Double = 0

            For j As Integer = 7 To lastCol ' Comenzar desde la columna 7
                Dim columnName As String = dtFinal.Columns(j - 1).ColumnName ' j - 1 porque Cells es 1-based
                Dim cellValue As Double

                ' Comprobar si la celda no está vacía y convertir el valor a Double
                If Double.TryParse(worksheet.Cells(i, j).Text, cellValue) Then
                    If columnName.EndsWith("-PROD") Then
                        totalProd += cellValue
                    ElseIf columnName.EndsWith("-NOPROD") Then
                        totalNoProd += cellValue
                    End If
                End If
            Next

            ' Colocar los totales en las columnas definidas
            worksheet.Cells(i, totalProdCol).Value = totalProd
            worksheet.Cells(i, totalNoProdCol).Value = totalNoProd
            worksheet.Cells(i, totalGeneralCol).Value = totalProd + totalNoProd ' Suma de ambos totales
        Next
    End Sub

    Public Sub GestionarEstilos(ByVal worksheet As ExcelWorksheet)
        ' Verificar si hay datos en el worksheet
        If worksheet.Dimension Is Nothing Then
            Return ' No hay datos para formatear
        End If

        AplicarEstiloEncabezado(worksheet)
        AjustarAnchoColumnasYFiltrar(worksheet)
        AplicarEstiloColumnas(worksheet)
        AplicarBordesTabla(worksheet)
        AplicarNegritaColumnas(worksheet, 3, 4)
        worksheet.View.FreezePanes(2, 1)
        ColorearFilasPorCompania(worksheet)
        ColorearColumnasPorFecha(worksheet)
    End Sub

    Private Sub AplicarEstiloEncabezado(ByVal worksheet As ExcelWorksheet)
        ' Aplicar negrita y pintar la primera fila de azul
        Dim filaEncabezado As ExcelRange = worksheet.Cells(1, 1, 1, worksheet.Dimension.End.Column)
        filaEncabezado.Style.Font.Bold = True
    End Sub

    Private Sub AjustarAnchoColumnasYFiltrar(ByVal worksheet As ExcelWorksheet)
        ' Ajustar el ancho de las columnas
        worksheet.Column(3).Width = 33 ' Ajusta según sea necesario
        worksheet.Column(5).Width = 15
        worksheet.Column(6).Width = 15
        worksheet.Cells("A1:" & GetExcelColumnName(worksheet.Dimension.End.Column) & "1").AutoFilter = True
    End Sub

    Private Sub AplicarEstiloColumnas(ByVal worksheet As ExcelWorksheet)
        ' Aplicar estilo azul claro a las primeras 6 columnas
        Dim azulClaro As Color = Color.FromArgb(225, 243, 250) ' Light Blue
        For i As Integer = 1 To 6
            Dim rango As ExcelRange = worksheet.Cells(1, i, worksheet.Dimension.End.Row, i)
            rango.Style.Fill.PatternType = ExcelFillStyle.Solid
            rango.Style.Fill.BackgroundColor.SetColor(azulClaro)
        Next
    End Sub

    Private Sub AplicarBordesTabla(ByVal worksheet As ExcelWorksheet)
        ' Aplicar bordes finos a todo el rango de la tabla
        Dim rangoTabla As ExcelRange = worksheet.Cells(1, 1, worksheet.Dimension.End.Row, worksheet.Dimension.End.Column)
        rangoTabla.Style.Border.Top.Style = ExcelBorderStyle.Thin
        rangoTabla.Style.Border.Bottom.Style = ExcelBorderStyle.Thin
        rangoTabla.Style.Border.Left.Style = ExcelBorderStyle.Thin
        rangoTabla.Style.Border.Right.Style = ExcelBorderStyle.Thin

        ' Centrar datos desde la columna 7 y fila 2 hasta el final
        Dim rangoACentrar As ExcelRange = worksheet.Cells(2, 7, worksheet.Dimension.End.Row, worksheet.Dimension.End.Column)
        rangoACentrar.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center
        rangoACentrar.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center

    End Sub

    Private Sub AplicarNegritaColumnas(ByVal worksheet As ExcelWorksheet, ByVal inicioColumna As Integer, ByVal finColumna As Integer)
        ' Aplicar negrita a las columnas específicas
        For col As Integer = inicioColumna To finColumna
            Dim rangoColumna As ExcelRange = worksheet.Cells(1, col, worksheet.Dimension.End.Row, col)
            rangoColumna.Style.Font.Bold = True
        Next
    End Sub

    Private Sub ColorearFilasPorCompania(ByVal worksheet As ExcelWorksheet)
        ' Definir colores para las diferentes compañías
        Dim colorNaranja As Color = Color.FromArgb(255, 229, 204) ' Naranja clarito
        Dim colorAmarillo As Color = Color.FromArgb(255, 255, 204) ' Amarillo clarito
        Dim colorVerde As Color = Color.FromArgb(204, 255, 204) ' Verde clarito
        Dim colorBlanco As Color = Color.White ' Blanco

        ' Iterar a través de las filas, comenzando desde la fila 2
        For fila As Integer = 2 To worksheet.Dimension.End.Row
            ' Obtener el valor de la columna "COMPAÑIA"
            Dim valorCompañia As String = worksheet.Cells(fila, 4).Text ' Columna COMPAÑIA (índice 4)

            ' Definir el color de fondo según el valor de la columna COMPAÑIA
            Dim colorFondo As Color = colorBlanco
            Select Case valorCompañia
                Case "TCZ"
                    colorFondo = colorNaranja
                Case "DS"
                    colorFondo = colorAmarillo
                Case "SOAR"
                    colorFondo = colorVerde
            End Select

            ' Aplicar el color de fondo a las primeras 6 columnas de la fila actual
            For col As Integer = 1 To 6
                Dim rangoCelda As ExcelRange = worksheet.Cells(fila, col)
                rangoCelda.Style.Fill.PatternType = ExcelFillStyle.Solid
                rangoCelda.Style.Fill.BackgroundColor.SetColor(colorFondo)
            Next
        Next
    End Sub

    Private Sub ColorearColumnasPorFecha(ByVal worksheet As ExcelWorksheet)
        ' Recorrer las columnas desde la columna 7 en adelante
        For col As Integer = 7 To worksheet.Dimension.End.Column
            ' Obtener el valor de la celda en la fila 1 (encabezado) en formato "dd/MM/yy-PROD"
            Dim valorFecha As String = worksheet.Cells(1, col).Text
            Dim parteFecha As String = valorFecha.Split("-"c)(0) ' Obtener solo la parte antes del "-"
            Dim fecha As DateTime

            ' Intentar convertir el valor a una fecha
            If DateTime.TryParseExact(parteFecha, "dd/MM/yy", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, fecha) Then
                ' Definir los colores gris claro para sábado y gris oscuro para domingo
                Dim grisClaro As Color = Color.FromArgb(230, 230, 230) ' Light gray
                Dim grisOscuro As Color = Color.FromArgb(200, 200, 200) ' Darker gray

                ' Comprobar si la fecha es un sábado o domingo
                Dim rango As ExcelRange = worksheet.Cells(1, col, worksheet.Dimension.End.Row, col)
                If fecha.DayOfWeek = DayOfWeek.Saturday Then
                    ' Aplicar el color gris claro a la columna completa
                    rango.Style.Fill.PatternType = ExcelFillStyle.Solid
                    rango.Style.Fill.BackgroundColor.SetColor(grisClaro)
                ElseIf fecha.DayOfWeek = DayOfWeek.Sunday Then
                    ' Aplicar el color gris oscuro a la columna completa
                    rango.Style.Fill.PatternType = ExcelFillStyle.Solid
                    rango.Style.Fill.BackgroundColor.SetColor(grisOscuro)
                End If
            End If
        Next
    End Sub

    Function GetExcelColumnName(ByVal columnNumber As Integer) As String
        Dim dividend As Integer = columnNumber
        Dim columnName As String = String.Empty
        Dim modulo As Integer

        While dividend > 0
            modulo = (dividend - 1) Mod 26
            columnName = Convert.ToChar(65 + modulo) & columnName
            dividend = CInt((dividend - modulo) / 26)
        End While

        Return columnName
    End Function

End Class
