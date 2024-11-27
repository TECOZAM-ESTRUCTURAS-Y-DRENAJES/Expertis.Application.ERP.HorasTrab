Imports Solmicro.Expertis.Engine
Imports System.Windows.Forms
Imports OfficeOpenXml
Imports OfficeOpenXml.Style
Imports System.Drawing
Imports System.IO
Imports System.Collections.Generic

Public Class ExportacionNoruegaCuadrante

    Public tablaDatos As String
    Public tipoExportacion As String

    Public Sub generaExcelNoruega()
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
        dtPersonas = getTablaPersonas()

        '2. LE DOY LA 1ª FORMA A LA TABLA
        Dim dtFinal As New DataTable
        FormaTablaSalidaNoruega(dtFinal)
        setPrimerCambioForma(dtFinal, dtPersonas)

        ExportarFichero(dtFinal, Fecha1, Fecha2)
    End Sub

    Public Function getTablaPersonas() As DataTable
        Dim strWhere As String = "IDOperario !='00'"
        Return New BE.DataEngine().Filter("frmMntoOperario", "IDOperario, Nombre, Apellidos, FechaAlta, Fecha_Baja, IDOficio", strWhere)
    End Function

    Public Sub FormaTablaSalidaNoruega(ByRef dtFinal As DataTable)
        dtFinal.Columns.Add("EXP.")
        dtFinal.Columns.Add("Name:")
        dtFinal.Columns.Add("SITE")
        dtFinal.Columns.Add("Start day:")
        dtFinal.Columns.Add("Finish day:")
        dtFinal.Columns.Add("Skill")
    End Sub

    Public Sub setPrimerCambioForma(ByRef dtFinal As DataTable, ByRef dtPersonas As DataTable)
        ' Iterar a través de cada fila en dtPersonas
        For Each dr As DataRow In dtPersonas.Rows
            ' Crear una nueva fila en dtFinal
            Dim newRow As DataRow = dtFinal.NewRow()

            ' Asignar los valores de la fila actual de dtPersonas a la nueva fila de dtFinal
            newRow("EXP.") = dr("IDOperario")
            newRow("Name:") = dr("Nombre") & " " & Nz(dr("Apellidos"), "")
            newRow("SITE") = ""
            ' Manejar la conversión y formato de la fecha
            Dim fechaAlta As Object = dr("FechaAlta")
            If IsDBNull(fechaAlta) OrElse Not DateTime.TryParse(fechaAlta.ToString(), New DateTime()) Then
                newRow("Start day:") = DBNull.Value
            Else
                Dim fecha As DateTime = Convert.ToDateTime(fechaAlta)
                newRow("Start day:") = fecha.ToString("dd/MM/yyyy") ' Formatear solo la fecha
            End If
            Dim fechaBaja As Object = dr("Fecha_Baja")
            If IsDBNull(fechaBaja) OrElse Not DateTime.TryParse(fechaBaja.ToString(), New DateTime()) Then
                newRow("Finish day:") = DBNull.Value
            Else
                Dim fecha As DateTime = Convert.ToDateTime(fechaBaja)
                newRow("Finish day:") = fecha.ToString("dd/MM/yyyy") ' Formatear solo la fecha
            End If

            Dim idOficio As String = dr("IDOficio").ToString()
            If idOficio = "ADMINOBRA" OrElse idOficio = "TEC_OBRA" OrElse idOficio = "TECPROY" OrElse idOficio = "INGENIERO" Then
                newRow("Skill") = "STAFF"
            Else
                newRow("Skill") = idOficio
            End If

            ' Agregar la nueva fila a dtFinal
            dtFinal.Rows.Add(newRow)
        Next
    End Sub

    Public Sub ExportarFichero(ByVal dtFinal As DataTable, ByVal fecha1 As String, ByVal fecha2 As String)
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial

        Dim saveFileDialog1 As New SaveFileDialog()
        saveFileDialog1.Filter = "Archivos de Excel|*.xlsx|Todos los archivos|*.*"
        saveFileDialog1.Title = "Guardar archivo"

        ' Mostrar el cuadro de diálogo y verificar si el usuario hizo clic en "Guardar"
        If saveFileDialog1.ShowDialog() = DialogResult.OK Then
            ' Obtener la ruta seleccionada por el usuario
            Dim rutaArchivo As String = saveFileDialog1.FileName

            Using package As New ExcelPackage()

                ' Convertir fecha2 a DateTime y formatear el nombre de la hoja
                Dim fecha As DateTime = Convert.ToDateTime(fecha2)
                Dim nombreHoja As String = fecha.ToString("MMM yyyy").ToUpper() ' Formato JUN 2024

                ' Crear una hoja de cálculo y obtener una referencia a ella.
                Dim worksheet = package.Workbook.Worksheets.Add(nombreHoja)

                ' Copiar los datos de la DataTable a la hoja de cálculo.
                worksheet.Cells("A5").LoadFromDataTable(dtFinal, True)

                ' Aplicar estilos personalizados
                GestionarEstilos(worksheet, dtFinal, fecha1)
                GestionDatosHoras(worksheet, dtFinal, fecha1)
                GestionarFormulacion(worksheet)
                GestionarOvertime(worksheet, fecha1)


                If tipoExportacion = "ORIGINAL" Then

                Else
                    ' Crear una nueva hoja llamada TURNOS y llamar al método creaHojaTurnos
                    Dim worksheetTurnos = package.Workbook.Worksheets.Add("TURNOS OPERARIOS")
                    creaHojaTurnos(worksheetTurnos, dtFinal, fecha1)
                End If
                

                ' Crear una nueva hoja llamada PARAMETROS y llamar al método creaHojaParametros
                Dim worksheetParametros = package.Workbook.Worksheets.Add("PARAMETROS")
                creaHojaParametros(worksheetParametros)

                ' Crear una nueva hoja llamada REGISTROS TURNOS
                Dim worksheetRegistrosTurnos = package.Workbook.Worksheets.Add("REGISTROS TURNOS")
                creaHojaRegistrosTurnos(worksheetRegistrosTurnos, fecha1, fecha2)

                ' Guardar el paquete de Excel en la ruta seleccionada
                Dim fileInfo As New IO.FileInfo(rutaArchivo)
                package.SaveAs(fileInfo)
            End Using
        End If
        MessageBox.Show("Fichero guardado correctamente", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Public Sub creaHojaTurnos(ByVal worksheet As ExcelWorksheet, ByVal dtFinal As DataTable, ByVal fecha1 As String)
        FormaTablaTurnos(dtFinal)
        DesgloseValores(dtFinal, fecha1)

        worksheet.Cells("A1").LoadFromDataTable(dtFinal, True)

        ' Aplicar formato negrita a la fila 1
        Dim fila1 As ExcelRange = worksheet.Cells(1, 1, 1, worksheet.Dimension.End.Column)
        fila1.Style.Font.Bold = True

        ' Ajusta el ancho de la columna B a 20
        worksheet.Column(2).Width = 20
        worksheet.Column(4).Width = 15
        worksheet.Column(5).Width = 15

        ' Inmoviliza paneles desde la fila 2 y la columna F
        worksheet.View.FreezePanes(2, 7)

        ' Aplica bordes a toda la tabla
        Dim dataRange As ExcelRange = worksheet.Cells(1, 1, dtFinal.Rows.Count + 1, dtFinal.Columns.Count)
        dataRange.Style.Border.Top.Style = ExcelBorderStyle.Thin
        dataRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin
        dataRange.Style.Border.Left.Style = ExcelBorderStyle.Thin
        dataRange.Style.Border.Right.Style = ExcelBorderStyle.Thin

        ' Pintar las filas en función de la fecha de  baja
        For i As Integer = 0 To dtFinal.Rows.Count - 1

            Dim filaActual As Integer = i + 2 ' Ajustar para comenzar en la fila 6
            Dim idOficio As String = dtFinal.Rows(i)("Skill").ToString()

            ' Primero, pintar en función del valor de IDOficio
            If idOficio = "STAFF" OrElse idOficio = "ENCARGADO" Then
                Dim rango As ExcelRange = worksheet.Cells(filaActual, 1, filaActual, 6)
                rango.Style.Fill.PatternType = ExcelFillStyle.Solid
                rango.Style.Fill.BackgroundColor.SetColor(Color.LightGray)
            End If

            ' Luego, pintar en función de la fecha de baja
            If Not IsDBNull(dtFinal.Rows(i)("Finish day:")) Then
                Dim rango As ExcelRange = worksheet.Cells(filaActual, 1, filaActual, 6)
                rango.Style.Fill.PatternType = ExcelFillStyle.Solid
                rango.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 0, 176, 240))

                ' Tachado del contenido de la columna "Name"
                Dim nombreCelda As ExcelRange = worksheet.Cells(filaActual, 2) ' Suponiendo que la columna "Name" es la segunda columna (B)
                nombreCelda.Style.Font.Strike = True
            End If
        Next

        PintarFindes(worksheet, dtFinal, fecha1)
    End Sub

    Public Sub PintarFindes(ByVal worksheet As ExcelWorksheet, ByVal dtFinal As DataTable, ByVal fecha1 As String)
        ' Rellenar la fila desde G6 con los números del 1 al 31 repetidos tres veces
        Dim columnaInicial As Integer = 7 ' Columna G
        Dim columnaFinal As Integer = columnaInicial + 31 * 2 - 1 ' BP
        Dim numeroDia As Integer = 1
        Dim fechaInicio As Date = New DateTime(DateTime.Now.Year, DateTime.Now.Month, 1) ' Primer día del mes actual

        For columna As Integer = columnaInicial To columnaFinal
            Dim celda As ExcelRange = worksheet.Cells(1, columna) ' Fila 1
            Dim parts() As String = celda.Text.ToString.Split(" "c)

            Dim fechaActual As Date
            Try
                fechaActual = New DateTime(Year(fecha1), Month(fecha1), parts(0))
                ' Verificar el día de la semana
                Dim diaSemana As DayOfWeek = fechaActual.DayOfWeek

                Select Case diaSemana
                    Case DayOfWeek.Saturday
                        ' Cambiar el color de fondo de la columna a amarillo fosforescente
                        For fila As Integer = 1 To worksheet.Dimension.End.Row
                            worksheet.Cells(fila, columna).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid
                            worksheet.Cells(fila, columna).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow)
                        Next
                    Case DayOfWeek.Sunday
                        For fila As Integer = 1 To worksheet.Dimension.End.Row
                            worksheet.Cells(fila, columna).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid
                            worksheet.Cells(fila, columna).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow)
                        Next
                End Select
            Catch ex As Exception
            End Try

        Next
    End Sub

    Public Sub DesgloseValores(ByVal dtFinal As DataTable, ByVal fecha1 As String)
        ' Obtener todas las fechas relevantes del mes
        Dim mes As Integer = Month(fecha1)
        Dim año As Integer = Year(fecha1)
        Dim fechaInicio As New DateTime(año, mes, 1)
        Dim fechaFin As New DateTime(año, mes, DateTime.DaysInMonth(año, mes))


        ' Hacer una sola consulta para obtener todas las horas de entrada y salida de todos los operarios
        Dim dtRegistro As New DataTable
        Dim filtro As New Filter
        filtro.Add("FechaParte", FilterOperator.GreaterThanOrEqual, fechaInicio)
        filtro.Add("FechaParte", FilterOperator.LessThanOrEqual, fechaFin)

        dtRegistro = New BE.DataEngine().Filter(tablaDatos, filtro)

        ' Crear diccionarios para almacenar las horas de entrada y salida por operario y fecha
        Dim horasEntrada As New Dictionary(Of String, Dictionary(Of DateTime, String))()
        Dim horasSalida As New Dictionary(Of String, Dictionary(Of DateTime, String))()

        For Each row As DataRow In dtRegistro.Rows
            Dim idOperario As String = row("IDOperario").ToString()
            Dim fechaParte As DateTime = CType(row("FechaParte"), DateTime)
            Dim horaEntrada As String = If(row.IsNull("HoraEntrada"), String.Empty, row("HoraEntrada").ToString())
            Dim horaSalida As String = If(row.IsNull("HoraSalida"), String.Empty, row("HoraSalida").ToString())

            If Not horasEntrada.ContainsKey(idOperario) Then
                horasEntrada(idOperario) = New Dictionary(Of DateTime, String)()
            End If
            If Not horasSalida.ContainsKey(idOperario) Then
                horasSalida(idOperario) = New Dictionary(Of DateTime, String)()
            End If

            horasEntrada(idOperario)(fechaParte) = horaEntrada
            horasSalida(idOperario)(fechaParte) = horaSalida
        Next

        ' Procesar los datos en memoria
        For Each row As DataRow In dtFinal.Rows
            Dim idOperario As String = row("EXP.").ToString()
            Dim diaColumna As Integer = 1
            For i As Integer = 6 To dtFinal.Columns.Count - 1
                Dim columnName As String = dtFinal.Columns(i).ColumnName
                If columnName.EndsWith(" E") Then
                    Dim numero As Integer = Integer.Parse(columnName.Split(" "c)(0))

                    Dim fechaParte As DateTime
                    Try
                        fechaParte = New DateTime(año, mes, numero)
                    Catch ex As Exception
                        Exit For
                    End Try


                    If horasEntrada.ContainsKey(idOperario) AndAlso horasEntrada(idOperario).ContainsKey(fechaParte) Then
                        Dim turnoEntrada As String = horasEntrada(idOperario)(fechaParte)
                        If Len(turnoEntrada) > 0 Then
                            row(diaColumna & " E") = turnoEntrada
                        End If
                    End If
                ElseIf columnName.EndsWith(" S") Then
                    Dim numero As Integer = Integer.Parse(columnName.Split(" "c)(0))
                    Dim fechaParte As New DateTime(año, mes, numero)

                    If horasSalida.ContainsKey(idOperario) AndAlso horasSalida(idOperario).ContainsKey(fechaParte) Then
                        Dim turnoSalida As String = horasSalida(idOperario)(fechaParte)
                        If Len(turnoSalida) > 0 Then
                            row(diaColumna & " S") = turnoSalida
                        End If
                        diaColumna += 1
                    Else
                        diaColumna += 1
                    End If
                End If
            Next
        Next
    End Sub

    Function ProcesarTurnoEntrada(ByVal dia As Integer, ByVal fecha1 As String, ByVal IDOperario As String) As String
        Dim mes As String = Month(fecha1)
        Dim año As String = Year(fecha1)

        Dim fechaParte As New DateTime(CInt(año), CInt(mes), dia)

        Dim dtRegistro As New DataTable
        Dim filtro As New Filter
        filtro.Add("FechaParte", FilterOperator.Equal, fechaParte)
        filtro.Add("IDOperario", FilterOperator.Equal, IDOperario)
        dtRegistro = New BE.DataEngine().Filter(tablaDatos, filtro)

        If dtRegistro.Rows.Count > 0 Then
            Return dtRegistro.Rows(0)("HoraEntrada").ToString
        End If

    End Function

    Function ProcesarTurnoSalida(ByVal dia As Integer, ByVal fecha1 As String, ByVal IDOperario As String) As String
        Dim mes As String = Month(fecha1)
        Dim año As String = Year(fecha1)

        Dim fechaParte As New DateTime(CInt(año), CInt(mes), dia)

        Dim dtRegistro As New DataTable
        Dim filtro As New Filter
        filtro.Add("FechaParte", FilterOperator.Equal, fechaParte)
        filtro.Add("IDOperario", FilterOperator.Equal, IDOperario)
        dtRegistro = New BE.DataEngine().Filter(tablaDatos, filtro)

        If dtRegistro.Rows.Count > 0 Then
            Return dtRegistro.Rows(0)("HoraSalida").ToString
        End If

    End Function

    Public Sub FormaTablaTurnos(ByVal dtFinal As DataTable)
        ' Agregar nuevas columnas al DataTable dtFinal
        For i As Integer = 1 To 31
            dtFinal.Columns.Add(i.ToString() & " E", GetType(String))
            dtFinal.Columns.Add(i.ToString() & " S", GetType(String))
        Next
    End Sub
    Public Sub GestionarOvertime(ByVal worksheet As ExcelWorksheet, ByVal fecha1 As String)
        ' Define el rango origen y destino
        Dim startRow As Integer = 6
        Dim startColumnOrigen As Integer = 7 ' Columna AL
        Dim endColumnOrigen As Integer = 38 ' Columna BP

        ' Encuentra la última fila con datos en la columna AL
        Dim lastRow As Integer = worksheet.Dimension.End.Row
        Dim columnaDia As Integer
        For row As Integer = startRow To lastRow
            columnaDia = 1
            For col As Integer = startColumnOrigen To endColumnOrigen
                ' Obtén el valor de la celda en el rango origen
                Dim value As String = worksheet.Cells(row, col).Value
                If Len(value) <> 0 Then
                    AsignaColorFuentesyDatos(row, col, worksheet, value, columnaDia, fecha1)
                End If
                columnaDia += 1
            Next
        Next
    End Sub

    Public Sub AsignaColorFuentesyDatos(ByVal row As Integer, ByVal col As Integer, ByVal worksheet As ExcelWorksheet, ByVal value As String, ByVal columnaDia As Integer, ByVal fecha1 As String)

        ' Copia el color de la fuente de la celda origen a la celda destino
        Dim origenFontColor As String = worksheet.Cells(row, col).Style.Font.Color.Rgb
        Try
            worksheet.Cells(row, col + 62).Style.Font.Color.SetColor(System.Drawing.ColorTranslator.FromHtml("#" & origenFontColor))
        Catch ex As Exception
            worksheet.Cells(row, col + 62).Style.Font.Color.SetColor(System.Drawing.Color.Black)
        End Try

        AsignaDatos(row, col, worksheet, value, columnaDia, fecha1)
    End Sub

    Public Sub AsignaDatos(ByVal row As Integer, ByVal col As Integer, ByVal worksheet As ExcelWorksheet, ByVal value As String, ByVal columnaDia As Integer, ByVal fecha1 As String)
        Dim origenBgColor As String = worksheet.Cells(row, col).Style.Fill.BackgroundColor.Rgb

        Dim r As Integer = 0
        Dim g As Integer = 0
        Dim b As Integer = 0

        If origenBgColor Is Nothing Then
            r = 255
            g = 255
            b = 255
        Else
            r = Convert.ToInt32(origenBgColor.Substring(0, 2), 16)
            g = Convert.ToInt32(origenBgColor.Substring(2, 2), 16)
            b = Convert.ToInt32(origenBgColor.Substring(4, 2), 16)
        End If


        Dim IDOficio As String
        IDOficio = worksheet.Cells(row, 6).Value

        Try
            Dim fecha As New DateTime(Year(fecha1), Month(fecha1), columnaDia)

            '1. SI NO ES NUMERIC PONLO TAL CUAL
            '2. SI ES NUMERICO, CHEQUEA SI ES STAFF SE PONE UN 0
            ' SI NO ES STAFF SE COMPRUEBA EL TURNO Y DIA DE LA SEMANA
            If Not IsNumeric(value) Then
                worksheet.Cells(row, col + 62).Value = value
            Else
                If IDOficio = "STAFF" Then
                    worksheet.Cells(row, col + 62).Value = CDbl(0)
                Else
                    'SI ES DEL TURNO 1 O DE OTRO
                    ' Condición principal para verificar el turno
                    If r = 255 AndAlso g = 255 AndAlso b = 255 Then
                        If fecha.DayOfWeek >= DayOfWeek.Monday AndAlso fecha.DayOfWeek <= DayOfWeek.Friday Then
                            ' Acción para días de lunes a viernes
                            Dim horas = CDbl(worksheet.Cells(row, col).Value - 7.5)
                            If horas < 0 Then
                                worksheet.Cells(row, col + 62).Value = CDbl(0)
                            Else
                                worksheet.Cells(row, col + 62).Value = horas
                            End If
                        Else
                            ' Acción para fines de semana
                            worksheet.Cells(row, col + 62).Value = CDbl(value)
                        End If

                    ElseIf r = 255 AndAlso g = 244 AndAlso b = 180 Then
                        If fecha.DayOfWeek >= DayOfWeek.Monday AndAlso fecha.DayOfWeek <= DayOfWeek.Friday Then
                            ' Acción para días de lunes a viernes
                            Dim horas = CDbl(worksheet.Cells(row, col).Value - 6.5)
                            If horas < 0 Then
                                worksheet.Cells(row, col + 62).Value = CDbl(0)
                            Else
                                worksheet.Cells(row, col + 62).Value = horas
                            End If
                        Else
                            ' Acción para fines de semana
                            worksheet.Cells(row, col + 62).Value = CDbl(value) - 5
                        End If

                    ElseIf r = 255 AndAlso g = 120 AndAlso b = 216 Then
                        If fecha.DayOfWeek >= DayOfWeek.Monday AndAlso fecha.DayOfWeek <= DayOfWeek.Friday Then
                            ' Acción para días de lunes a viernes
                            Dim horas = CDbl(worksheet.Cells(row, col).Value - 6.5)
                            If horas < 0 Then
                                worksheet.Cells(row, col + 62).Value = CDbl(0)
                            Else
                                worksheet.Cells(row, col + 62).Value = horas
                            End If
                        Else
                            ' Acción para fines de semana
                            worksheet.Cells(row, col + 62).Value = CDbl(value) - 5
                        End If

                    ElseIf r = 255 AndAlso g = 156 AndAlso b = 196 Then
                        If fecha.DayOfWeek >= DayOfWeek.Monday AndAlso fecha.DayOfWeek <= DayOfWeek.Friday Then
                            ' Acción para días de lunes a viernes
                            Dim horas = CDbl(worksheet.Cells(row, col).Value - 7.5)
                            If horas < 0 Then
                                worksheet.Cells(row, col + 62).Value = CDbl(0)
                            Else
                                worksheet.Cells(row, col + 62).Value = horas
                            End If
                            ' Acción para fines de semana comentada en el código original
                            ' Else
                            '    worksheet.Cells(row, col + 62).Value = CDbl(value) - 5
                        End If

                    ElseIf r = 255 AndAlso g = 168 AndAlso b = 164 Then
                        If fecha.DayOfWeek >= DayOfWeek.Monday AndAlso fecha.DayOfWeek <= DayOfWeek.Friday Then
                            ' Acción para días de lunes a viernes
                            Dim horas = CDbl(worksheet.Cells(row, col).Value - 7.5)
                            If horas < 0 Then
                                worksheet.Cells(row, col + 62).Value = CDbl(0)
                            Else
                                worksheet.Cells(row, col + 62).Value = horas
                            End If
                            ' Acción para fines de semana comentada en el código original
                            ' Else
                            '    worksheet.Cells(row, col + 62).Value = CDbl(value) - 5
                        End If
                    Else
                        If fecha.DayOfWeek >= DayOfWeek.Monday AndAlso fecha.DayOfWeek <= DayOfWeek.Friday Then
                            ' Acción para días de lunes a viernes
                            Dim horas = CDbl(worksheet.Cells(row, col).Value - 7.5)
                            If horas < 0 Then
                                worksheet.Cells(row, col + 62).Value = CDbl(0)
                            Else
                                worksheet.Cells(row, col + 62).Value = horas
                            End If
                        Else
                            ' Acción para fines de semana
                            worksheet.Cells(row, col + 62).Value = CDbl(value)
                        End If
                    End If
                End If

            End If
        Catch ex As Exception
        End Try
    End Sub

    Public Sub GestionarFormulacion(ByVal worksheet As ExcelWorksheet)
        ' Define el rango desde AL9 hasta BP y hasta la última fila
        Dim startColumn As Integer = worksheet.Cells("AL9").Start.Column
        Dim endColumn As Integer = worksheet.Cells("BP9").Start.Column
        Dim startRow As Integer = 6
        Dim endRow As Integer = worksheet.Dimension.End.Row

        ' Recorre cada celda en el rango definido
        For row As Integer = startRow To endRow
            For col As Integer = startColumn To endColumn
                ' Calcula el índice de la columna base (AL = 38, AM = 39, etc.)
                Dim columnOffset As Integer = col - startColumn + 7 ' 7 es la posición de la columna G en relación a AL
                Dim sourceColumn As Integer = columnOffset ' 7 es la columna G en términos de índice 1-based

                ' Obtén la letra de la columna de origen
                Dim sourceColumnLetter As String = GetColumnLetter(sourceColumn)

                Dim celda As String
                celda = worksheet.Cells(row, sourceColumn).Text
                ' Verifica si la celda actual tiene un valor
                If Len(celda) <> 0 Then
                    ' Asigna la fórmula a la celda actual si está vacía
                    Dim formula As String = "=IF(" & sourceColumnLetter & row & "=" & ChrW(34) & "A" & ChrW(34) & ", " & ChrW(34) & "A" & ChrW(34) & _
                         ", IF(" & sourceColumnLetter & row & "=" & ChrW(34) & ChrW(34) & ", " & ChrW(34) & ChrW(34) & _
                         ", IF(" & sourceColumnLetter & row & "=" & ChrW(34) & "V" & ChrW(34) & ", " & ChrW(34) & "V" & ChrW(34) & _
                         ", IF(" & sourceColumnLetter & row & "=" & ChrW(34) & "UA" & ChrW(34) & ", " & ChrW(34) & "UA" & ChrW(34) & _
                         ", IF(" & sourceColumnLetter & row & "=" & ChrW(34) & "VIAJE" & ChrW(34) & ", " & ChrW(34) & "VIAJE" & ChrW(34) & _
                         ", IF(" & sourceColumnLetter & row & "=" & ChrW(34) & "B" & ChrW(34) & ", " & ChrW(34) & "B" & ChrW(34) & _
                         ", IF(" & sourceColumnLetter & row & "=" & ChrW(34) & "H" & ChrW(34) & ", " & ChrW(34) & "H" & ChrW(34) & _
                         ", IF(" & sourceColumnLetter & row & "=" & ChrW(34) & "D" & ChrW(34) & ", " & ChrW(34) & "D" & ChrW(34) & _
                         ", " & sourceColumnLetter & row & "-" & GetColumnLetter(col + 31) & row & "))))))))"

                    worksheet.Cells(row, col).Formula = formula

                    ' Obtén el color de la fuente de la celda fuente
                    Dim sourceCell = worksheet.Cells(row, sourceColumn)
                    Dim sourceFontColor As ExcelColor = sourceCell.Style.Font.Color

                    ' Establece el color de la fuente en la celda destino
                    If sourceFontColor IsNot Nothing AndAlso sourceFontColor.Rgb IsNot Nothing Then
                        Dim colorHex As String = sourceFontColor.Rgb
                        Dim systemColor As System.Drawing.Color = System.Drawing.ColorTranslator.FromHtml("#" & colorHex)
                        worksheet.Cells(row, col).Style.Font.Color.SetColor(systemColor)
                    End If
                End If
            Next
        Next
    End Sub

    Private Function GetColumnLetter(ByVal columnIndex As Integer) As String
        Dim columnLetter As String = String.Empty
        Dim dividend As Integer = columnIndex
        Dim modulo As Integer

        While dividend > 0
            modulo = (dividend - 1) Mod 26
            columnLetter = Chr(65 + modulo) & columnLetter
            dividend = (dividend - modulo) \ 26
        End While

        Return columnLetter
    End Function
    Private Sub creaHojaParametros(ByVal worksheet As ExcelWorksheet)
        ' Crear un DataTable con dos columnas: IDMotivo y DescCausa
        Dim dtParametros As New DataTable()
        dtParametros.Columns.Add("IDMotivo", GetType(String))
        dtParametros.Columns.Add("DescCausa", GetType(String))

        ' Añadir registros al DataTable
        dtParametros.Rows.Add("H", "PUBLIC HOLIDAY DAY IN NORWAY")
        dtParametros.Rows.Add("V", "HOLIDAYS")
        dtParametros.Rows.Add("B", "DOES NOT WORK FOR HEALTH REASONS (16 FIRST NATURAL DAYS PAID BY EMPLOYER)")
        dtParametros.Rows.Add("XB", "DOES NOT WORK FOR HEALTH REASONS (MORE THAN 16 NATURAL DAYS PAID BY NAV)")
        dtParametros.Rows.Add("A", "HOLIDAYS")
        dtParametros.Rows.Add("UA", "DOES NOT WORK FOR HEALTH REASONS (16 FIRST NATURAL DAYS PAID BY EMPLOYER)")
        dtParametros.Rows.Add("DA", "PUBLIC HOLIDAY DAY IN NORWAY")
        dtParametros.Rows.Add("VIAJE", "VIAJE")

        ' Crear un DataTable con dos columnas: IDMotivo y DescCausa
        Dim dtTurnos As New DataTable()
        dtTurnos.Columns.Add("Turno", GetType(Integer))
        dtTurnos.Columns.Add("Horario", GetType(String))

        ' Añadir registros al DataTable
        dtTurnos.Rows.Add(1, "07:00h to 16:00h")
        dtTurnos.Rows.Add(2, "07:00h to 14:00h")
        dtTurnos.Rows.Add(3, "14:00h to 21:00h")
        dtTurnos.Rows.Add(4, "10:30h to 19:00h")
        dtTurnos.Rows.Add(5, "10:00h to 19:00h")
        dtTurnos.Rows.Add(6, "- OTRO HORARIO -")

        Dim dtTurnosExplicacion As New DataTable()
        dtTurnosExplicacion.Columns.Add("Explicacion", GetType(String))


        ' Añadir registros al DataTable
        dtTurnosExplicacion.Rows.Add("* The schedule begins at the corresponding schedule color's time. The number of hours is recorded in case more than the agreed-upon hours are worked and to monitor overtime.")
        dtTurnosExplicacion.Rows.Add("* (1) The schedule from 07:00h to 16:00h includes a 0,5 hour (10h-10:30h) break for rest and 1 hour (13h-14h) for lunch. This meal is taken in a canteen and therefore does not count as working time. MONDAY to FRIDAY")
        dtTurnosExplicacion.Rows.Add("* (2) The schedule from 07:00h to 14:00h includes a 0,5 hour (10h-10:30h) break. This meal is taken in a canteen and therefore does not count as working time. MONDAY to SATURDAY")
        dtTurnosExplicacion.Rows.Add("* (3) The schedule from 14:00h to 21:00h includes a 0,5 hour (17h-17:30h) break. This meal is taken in a canteen and therefore does not count as working time. MONDAY to SATURDAY")
        dtTurnosExplicacion.Rows.Add("* (4) The schedule from 10:30h to 19:00h includes a 1 hour (14h-15h) break. This meal is taken in a canteen and therefore does not count as working time. MONDAY to FRIDAY")
        dtTurnosExplicacion.Rows.Add("* (5) The schedule from 10:00h to 19:00h includes a 0,5 hour (13h-13:30h) break for rest and 1 hour (17h-18h) for lunch. This meal is taken in a canteen and therefore does not count as working time. MONDAY to FRIDAY")

        ' Copiar los datos de la DataTable a la hoja de cálculo
        worksheet.Cells("A1").LoadFromDataTable(dtParametros, True)
        worksheet.Cells("A11").LoadFromDataTable(dtTurnos, True)
        worksheet.Cells("A19").LoadFromDataTable(dtTurnosExplicacion, True)

        ' Aplicar formato de borde a las celdas
        ApplyBorder(worksheet.Cells("A1:B9"))
        ApplyBorder(worksheet.Cells("A11:B17"))

        ApplyCellBackgroundColor(worksheet.Cells("B13"), 244, 180, 132)
        ApplyCellBackgroundColor(worksheet.Cells("B14"), 120, 216, 112)
        ApplyCellBackgroundColor(worksheet.Cells("B15"), 156, 196, 228)
        ApplyCellBackgroundColor(worksheet.Cells("B16"), 168, 164, 164)
        ApplyCellBackgroundColor(worksheet.Cells("B17"), 255, 192, 203)

        worksheet.Column(2).Width = 20
    End Sub

    Public Sub creaHojaRegistrosTurnos(ByVal worksheet As ExcelWorksheet, ByVal fecha1 As String, ByVal fecha2 As String)
        Dim dtRegistros As New DataTable()

        Dim f As New Filter
        f.Add("FechaParte", FilterOperator.GreaterThanOrEqual, fecha1)
        f.Add("FechaParte", FilterOperator.LessThanOrEqual, fecha2)
        dtRegistros = New BE.DataEngine().Filter("frmMntoHorasInternacionalTecozam", f, _
                                         "FechaParte As DateReport, Turno As Shift, IDOperario As WorkerID, " & _
                                         "DescOperario As WorkerDescription, Oficio As Job, Estructura As Structure, HoraEntrada As StartTime, HoraSalida As EndTime, " & _
                                         "DescansoEnt As StartBreak, DescansoSal As EndBreak, Descanso2Ent As StartBreak2, Descanso2Sal As EndBreak2, " & _
                                         "TotalTurno As TotalShift, IDCausa, HoraEntradaOvertime As OvertimeStartTime, HoraSalidaOvertime As OvertimeEndTime, " & _
                                         "HoraEntradaDescansoOvertime As OvertimeStartBreak, HoraSalidaDescansoOvertime As OvertimeEndBreak, " & _
                                         "HoraEntradaDescansoOvertime2 As OvertimeStartBreak2, HoraSalidaDescansoOvertime2 As OvertimeEndBreak2, " & _
                                         "TotalOvertime, Horas As TotalHours")

        ' Crear el DataTable auxiliar con todas las columnas de tipo String
        Dim dtAuxiliar As New DataTable()

        ' Crear columnas en dtAuxiliar, ajustando el tipo según el caso
        For Each column As DataColumn In dtRegistros.Columns
            If column.ColumnName = "TotalShift" OrElse column.ColumnName = "TotalHours" Then
                ' Mantener el tipo original para TotalShift y TotalHours
                dtAuxiliar.Columns.Add(column.ColumnName, column.DataType)
            Else
                ' Convertir el resto a tipo String
                dtAuxiliar.Columns.Add(column.ColumnName, GetType(String))
            End If
        Next

        ' Copiar filas de dtRegistros a dtAuxiliar, convirtiendo los valores excepto para TotalShift y TotalHours
        For Each row As DataRow In dtRegistros.Rows
            Dim newRow As DataRow = dtAuxiliar.NewRow()
            For Each column As DataColumn In dtRegistros.Columns
                If column.ColumnName = "TotalShift" OrElse column.ColumnName = "TotalHours" Then
                    ' Mantener el valor original
                    newRow(column.ColumnName) = row(column)
                Else
                    ' Convertir a String
                    newRow(column.ColumnName) = row(column).ToString()
                End If
            Next
            dtAuxiliar.Rows.Add(newRow)
        Next


        worksheet.Cells("A1").LoadFromDataTable(dtAuxiliar, True)
        worksheet.Cells("A1:" & GetExcelColumnName(worksheet.Dimension.End.Column) & "1").AutoFilter = True

        Dim fila1 As ExcelRange = worksheet.Cells(1, 1, 1, worksheet.Dimension.End.Column)
        fila1.Style.Font.Bold = True

        Dim borderStyle As ExcelBorderStyle = ExcelBorderStyle.Thin
        For row As Integer = 1 To dtAuxiliar.Rows.Count + 1
            For col As Integer = 1 To dtAuxiliar.Columns.Count
                Dim cell As ExcelRange = worksheet.Cells(row, col)

                ' Aplicar estilo de borde a cada celda
                With cell.Style.Border
                    .Top.Style = BorderStyle
                    .Bottom.Style = BorderStyle
                    .Left.Style = BorderStyle
                    .Right.Style = BorderStyle
                End With
            Next
        Next

        ' Encontrar la posición de la columna "Shift" en dtAuxiliar
        Dim shiftColumnIndex As Integer = dtAuxiliar.Columns.IndexOf("Shift") + 1  ' +1 para ajustar al índice de columna en Excel

        For row As Integer = 2 To dtAuxiliar.Rows.Count + 1  ' Inicia en 2 para omitir el encabezado
            ' Obtener el valor de la columna "Shift" en la fila actual
            Dim shiftValue As String = worksheet.Cells(row, shiftColumnIndex).Text

            ' Determinar el color de la fila según el valor de "Shift"
            Dim rowColor As System.Drawing.Color = System.Drawing.Color.White  ' Blanco por defecto (sin color)

            If shiftValue = "1" Then
            ElseIf shiftValue = "2" Then
                rowColor = System.Drawing.Color.FromArgb(244, 180, 132)  ' Color RGB personalizado (37, 150, 190)
            ElseIf shiftValue = "3" Then
                rowColor = System.Drawing.Color.FromArgb(120, 216, 112) ' Color RGB personalizado (120, 216, 112)
            ElseIf shiftValue = "4" Then
                rowColor = System.Drawing.Color.FromArgb(156, 196, 228) ' Color RGB personalizado (156, 196, 228)
            ElseIf shiftValue = "5" Then
                rowColor = System.Drawing.Color.FromArgb(168, 164, 164)  ' Color RGB personalizado (168, 164, 164)
            End If

            For col As Integer = 1 To dtAuxiliar.Columns.Count
                Dim cell As ExcelRange = worksheet.Cells(row, col)

                ' Aplicar color de fondo si corresponde
                If rowColor <> System.Drawing.Color.White Then
                    cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid
                    cell.Style.Fill.BackgroundColor.SetColor(rowColor)
                End If

                ' Aplicar estilo de borde a cada celda
                With cell.Style.Border
                    .Top.Style = borderStyle
                    .Bottom.Style = borderStyle
                    .Left.Style = borderStyle
                    .Right.Style = borderStyle
                End With
            Next
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

    Private Sub ApplyBorder(ByVal range As ExcelRange)
        range.Style.Border.Top.Style = ExcelBorderStyle.Thin
        range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin
        range.Style.Border.Left.Style = ExcelBorderStyle.Thin
        range.Style.Border.Right.Style = ExcelBorderStyle.Thin
    End Sub

    Private Sub ApplyCellBackgroundColor(ByVal cell As ExcelRange, ByVal red As Byte, ByVal green As Byte, ByVal blue As Byte)
        cell.Style.Fill.PatternType = ExcelFillStyle.Solid
        cell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(red, green, blue))
    End Sub

    Public Sub SetInformacion(ByVal worksheet As ExcelWorksheet, ByVal fecha1 As String, ByVal dtFinal As DataTable)
        ' Poner "Fecha" en la primera celda (A1) y aplicar formato negrita
        Dim fecha As DateTime = Convert.ToDateTime(fecha1)
        Dim celdaFecha As ExcelRange = worksheet.Cells(1, 1) ' A1
        celdaFecha.Value = fecha.ToString("MMM-yy")
        celdaFecha.Style.Font.Bold = True

        ' Combinar celdas de B1 a F1 y poner "MONTHLY SHEET REPORT" en negrita
        Dim rangoCombinado As ExcelRange = worksheet.Cells(1, 2, 1, 6) ' Combina de B1 a F1
        rangoCombinado.Merge = True
        rangoCombinado.Value = "MONTHLY SHEET REPORT"
        rangoCombinado.Style.Font.Bold = True
        rangoCombinado.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center ' Opcional: centrar el texto

        ' Combinar celdas de A2 a F4 y poner "Reporte Detallado" en negrita
        Dim rangoCombinado2 As ExcelRange = worksheet.Cells(2, 1, 4, 6) ' Combina de A2 a F4
        rangoCombinado2.Merge = True
        rangoCombinado2.Value = "REPORTE DETALLADO"
        rangoCombinado2.Style.Font.Bold = True
        rangoCombinado2.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center ' Opcional: centrar el texto
        rangoCombinado2.Style.VerticalAlignment = ExcelVerticalAlignment.Center ' Opcional: centrar verticalmente el texto

        ' Combinar celdas de G2 a AK3 y poner "NORMAL + OVERTIME SCHEDULE" en negrita
        Dim rangoCombinado3 As ExcelRange = worksheet.Cells(2, 7, 3, 37) ' Combina de G2 a AK3
        rangoCombinado3.Merge = True
        rangoCombinado3.Value = MonthName(fecha.Month).ToUpper() & " (NORMAL + OVERTIME SCHEDULE)"
        rangoCombinado3.Style.Font.Bold = True
        rangoCombinado3.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
        rangoCombinado3.Style.VerticalAlignment = ExcelVerticalAlignment.Center

        ' Combinar celdas de AL2 a BP3 y poner "NORMAL SCHEDULE 1" en negrita
        Dim rangoCombinado4 As ExcelRange = worksheet.Cells(2, 38, 3, 68) ' Combina de AL2 a BP3
        rangoCombinado4.Merge = True
        rangoCombinado4.Value = MonthName(fecha.Month).ToUpper() & " (NORMAL SCHEDULE)"
        rangoCombinado4.Style.Font.Bold = True
        rangoCombinado4.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center ' Centrar el texto horizontalmente
        rangoCombinado4.Style.VerticalAlignment = ExcelVerticalAlignment.Center ' Centrar el texto verticalmente

        ' Combinar celdas de BQ2 a CU3 y poner "OVER TIME" en negrita
        Dim rangoCombinado5 As ExcelRange = worksheet.Cells(2, 69, 3, 99) ' Combina de BQ2 a CU3
        rangoCombinado5.Merge = True
        rangoCombinado5.Value = MonthName(fecha.Month).ToUpper() & " (OVER TIME)"
        rangoCombinado5.Style.Font.Bold = True
        rangoCombinado5.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center ' Centrar el texto horizontalmente
        rangoCombinado5.Style.VerticalAlignment = ExcelVerticalAlignment.Center ' Centrar el texto verticalmente

        ' Combinar celdas de CV4 a CV5 y poner "TOTAL NORMAL HOURS" en negrita
        Dim rangoTotalNormal As ExcelRange = worksheet.Cells(2, 100, 5, 100) ' Combina de CV4 a CV5
        rangoTotalNormal.Merge = True
        rangoTotalNormal.Value = "TOTAL NORMAL HOURS"
        rangoTotalNormal.Style.Font.Bold = True
        rangoTotalNormal.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center ' Centrar el texto horizontalmente
        rangoTotalNormal.Style.VerticalAlignment = ExcelVerticalAlignment.Center ' Centrar el texto verticalmente
        rangoTotalNormal.Style.WrapText = True ' Ajustar texto dentro de la celda

        ' Combinar celdas de CW4 a CW5 y poner "TOTAL OVERTIME HOURS" en negrita
        Dim rangoTotalOvertime As ExcelRange = worksheet.Cells(2, 101, 5, 101) ' Combina de CW4 a CW5
        rangoTotalOvertime.Merge = True
        rangoTotalOvertime.Value = "TOTAL OVERTIME HOURS"
        rangoTotalOvertime.Style.Font.Bold = True
        rangoTotalOvertime.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center ' Centrar el texto horizontalmente
        rangoTotalOvertime.Style.VerticalAlignment = ExcelVerticalAlignment.Center ' Centrar el texto verticalmente
        rangoTotalOvertime.Style.WrapText = True ' Ajustar texto dentro de la celda

        ' Establecer la fórmula para la suma de horas normales en la columna CV desde la fila 6 en adelante
        Dim startRow As Integer = 6
        Dim endRow As Integer = worksheet.Dimension.End.Row

        For row As Integer = startRow To endRow
            ' Fórmula de suma para TOTAL NORMAL HOURS (suma desde AL hasta BP)
            Dim formulaNormal As String = "=SUM(AL" & row & ":BP" & row & ")"
            worksheet.Cells(row, 100).Formula = formulaNormal ' CV

            ' Fórmula de suma para TOTAL OVERTIME HOURS (suma desde BQ hasta CU)
            Dim formulaOvertime As String = "=SUM(BQ" & row & ":CU" & row & ")"
            worksheet.Cells(row, 101).Formula = formulaOvertime ' CW
        Next

        ' Aplicar bordes a las celdas combinadas y con fórmulas
        Dim borderStyle As ExcelBorderStyle = ExcelBorderStyle.Thin

        ' Bordes para TOTAL NORMAL HOURS (CV4 a CV5)
        With rangoTotalNormal.Style.Border
            .Top.Style = borderStyle
            .Bottom.Style = borderStyle
            .Left.Style = borderStyle
            .Right.Style = borderStyle
        End With

        ' Bordes para TOTAL OVERTIME HOURS (CW4 a CW5)
        With rangoTotalOvertime.Style.Border
            .Top.Style = borderStyle
            .Bottom.Style = borderStyle
            .Left.Style = borderStyle
            .Right.Style = borderStyle
        End With

        ' Bordes para celdas con fórmulas en columna CV
        For row As Integer = startRow To endRow
            With worksheet.Cells(row, 100).Style.Border
                .Top.Style = borderStyle
                .Bottom.Style = borderStyle
                .Left.Style = borderStyle
                .Right.Style = borderStyle
            End With
        Next

        ' Bordes para celdas con fórmulas en columna CW
        For row As Integer = startRow To endRow
            With worksheet.Cells(row, 101).Style.Border
                .Top.Style = borderStyle
                .Bottom.Style = borderStyle
                .Left.Style = borderStyle
                .Right.Style = borderStyle
            End With
        Next

        GestionSabadosYDomingos(worksheet, fecha1)

        CentrarExcel(worksheet)
    End Sub

    Public Sub CentrarExcel(ByVal worksheet As ExcelWorksheet)
        ' Asumiendo que ya has definido `worksheet` como tu objeto ExcelWorksheet

        ' Definir el rango desde la columna G (7) hasta la columna CU (100)
        Dim startColumn As Integer = 7 ' Columna G
        Dim endColumn As Integer = 100 ' Columna CU
        Dim totalRows As Integer = worksheet.Dimension.End.Row ' Número total de filas

        ' Recorrer cada fila en el rango de columnas especificado
        For row As Integer = 1 To totalRows
            For col As Integer = startColumn To endColumn
                Dim celda As ExcelRange = worksheet.Cells(row, col)
                ' Centrar el contenido horizontalmente
                celda.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center
                ' Centrar el contenido verticalmente (opcional)
                celda.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center
            Next
        Next

        For row As Integer = 1 To totalRows
            For col As Integer = 1 To endColumn - 1
                Dim celda As ExcelRange = worksheet.Cells(row, col)
                With celda.Style.Border
                    .Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin
                    .Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin
                    .Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin
                    .Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin
                End With
            Next
        Next

        'PINTA EL DIA DE LA SEMANA
        ' Definir el color de fondo
        Dim color As System.Drawing.Color = System.Drawing.Color.FromArgb(164, 180, 192)

        For row As Integer = 4 To 4
            For col As Integer = startColumn To endColumn - 1
                Dim celda As ExcelRange = worksheet.Cells(row, col)
                ' Establecer el patrón de relleno a sólido y el color de fondo
                celda.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid
                celda.Style.Fill.BackgroundColor.SetColor(color)
            Next
        Next
    End Sub

    Public Sub GestionSabadosYDomingos(ByVal worksheet As ExcelWorksheet, ByVal fecha1 As String)
        ' Rellenar la fila desde G6 con los números del 1 al 31 repetidos tres veces
        Dim columnaInicial As Integer = 7 ' Columna G
        Dim columnaFinal As Integer = columnaInicial + 31 * 3 - 1 ' CU
        Dim numeroDia As Integer = 1
        Dim fechaInicio As Date = New DateTime(DateTime.Now.Year, DateTime.Now.Month, 1) ' Primer día del mes actual

        For columna As Integer = columnaInicial To columnaFinal
            Dim celda As ExcelRange = worksheet.Cells(5, columna) ' Fila 5
            celda.Value = numeroDia

            Dim fechaActual As Date
            Try
                fechaActual = New DateTime(Year(fecha1), Month(fecha1), numeroDia)
                ' Verificar el día de la semana
                Dim diaSemana As DayOfWeek = fechaActual.DayOfWeek

                Select Case diaSemana
                    Case DayOfWeek.Saturday
                        worksheet.Cells(4, columna).Value = "S" ' Sábado
                        ' Cambiar el color de fondo de la columna a amarillo fosforescente
                        For fila As Integer = 4 To worksheet.Dimension.End.Row
                            worksheet.Cells(fila, columna).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid
                            worksheet.Cells(fila, columna).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow)
                        Next
                    Case DayOfWeek.Sunday
                        worksheet.Cells(4, columna).Value = "D" ' Domingo
                        ' Cambiar el color de fondo de la columna a amarillo fosforescente
                        For fila As Integer = 4 To worksheet.Dimension.End.Row
                            worksheet.Cells(fila, columna).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid
                            worksheet.Cells(fila, columna).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow)
                        Next
                    Case Else
                        worksheet.Cells(4, columna).Value = "L" ' Día laborable
                End Select
            Catch ex As Exception
                worksheet.Cells(4, columna).Value = ""
            End Try

            ' Incrementar el número del día
            numeroDia += 1

            If numeroDia > 31 Then
                numeroDia = 1
            End If
            ' Ajustar el ancho de la columna a 3
            worksheet.Column(columna).Width = 6
        Next
    End Sub

    Public Sub GestionarEstilos(ByVal worksheet As ExcelWorksheet, ByVal dtFinal As DataTable, ByVal fecha1 As String)
        SetInformacion(worksheet, fecha1, dtFinal)

        ' Aplicar formato negrita a la fila 1
        Dim fila1 As ExcelRange = worksheet.Cells(5, 1, 5, worksheet.Dimension.End.Column)
        fila1.Style.Font.Bold = True

        ' Pintar las filas en función de la fecha de  baja
        For i As Integer = 0 To dtFinal.Rows.Count - 1

            Dim filaActual As Integer = i + 6 ' Ajustar para comenzar en la fila 6
            Dim idOficio As String = dtFinal.Rows(i)("Skill").ToString()

            ' Primero, pintar en función del valor de IDOficio
            If idOficio = "STAFF" OrElse idOficio = "ENCARGADO" Then
                Dim rango As ExcelRange = worksheet.Cells(filaActual, 1, filaActual, 6)
                rango.Style.Fill.PatternType = ExcelFillStyle.Solid
                rango.Style.Fill.BackgroundColor.SetColor(Color.LightGray)
            End If

            ' Luego, pintar en función de la fecha de baja
            If Not IsDBNull(dtFinal.Rows(i)("Finish day:")) Then
                Dim rango As ExcelRange = worksheet.Cells(filaActual, 1, filaActual, 6)
                rango.Style.Fill.PatternType = ExcelFillStyle.Solid
                rango.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 0, 176, 240))

                ' Tachado del contenido de la columna "Name"
                Dim nombreCelda As ExcelRange = worksheet.Cells(filaActual, 2) ' Suponiendo que la columna "Name" es la segunda columna (B)
                nombreCelda.Style.Font.Strike = True
            End If
        Next

        ' Ajustar el ancho de las columnas B, D, E y F
        worksheet.Column(2).Width = 30 ' Columna B
        worksheet.Column(4).Width = 15
        worksheet.Column(5).Width = 15
        worksheet.Column(6).Width = 20

        ' Inmovilizar paneles desde la línea 5
        worksheet.View.FreezePanes(6, 1)

    End Sub

    Public Sub GestionDatosHoras(ByVal worksheet As ExcelWorksheet, ByVal dtFinal As DataTable, ByVal fecha1 As String)
        '1º Control de la A de ausencia para las personas que no están para esa fecha.
        Dim fila As Integer = 1
        For Each dr As DataRow In dtFinal.Rows
            Dim fechaBaja As String = dr("Finish day:").ToString
            Dim fechaAlta As String = dr("Start day:").ToString

            For dia As Integer = 1 To 31
                Try
                    Dim fechaComparar As New Date(Year(fecha1), Month(fecha1), dia)
                    ProcesarDia(worksheet, dr, fila, dia, fechaComparar, fechaBaja, fechaAlta)
                Catch ex As Exception
                    Continue For
                End Try
            Next
            fila += 1
        Next
    End Sub

    Private Sub ProcesarDia(ByVal worksheet As ExcelWorksheet, ByVal dr As DataRow, ByVal fila As Integer, ByVal dia As Integer, ByVal fechaComparar As Date, ByVal fechaBaja As String, ByVal fechaAlta As String)
        If Len(fechaBaja) <> 0 Then
            If fechaBaja <= fechaComparar Or fechaAlta > fechaComparar Then
                EscribirAusencia(worksheet, fila, dia, fechaComparar)
            Else
                EscribirHorasTrabajo(worksheet, dr, fila, dia, fechaComparar)
            End If
        Else
            If fechaAlta > fechaComparar Then
                EscribirAusencia(worksheet, fila, dia, fechaComparar)
            Else
                EscribirHorasTrabajo(worksheet, dr, fila, dia, fechaComparar)
            End If
        End If
    End Sub

    Private Sub EscribirAusencia(ByVal worksheet As ExcelWorksheet, ByVal fila As Integer, ByVal dia As Integer, ByVal fechaComparar As Date)
        If EsDiaLaboral(fechaComparar) Then
            Dim columna As Integer = dia + 6 ' Columna G es la columna 7, por lo que agregamos 6
            Dim celda As ExcelRange = worksheet.Cells(fila + 5, columna)
            celda.Value = "A"
            celda.Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(0, 176, 240))
        End If
    End Sub

    Private Sub EscribirHorasTrabajo(ByVal worksheet As ExcelWorksheet, ByVal dr As DataRow, ByVal fila As Integer, ByVal dia As Integer, ByVal fechaComparar As Date)
        If dr("Skill") = "STAFF" Then
            If EsDiaLaboral(fechaComparar) Then
                Dim columna As Integer = dia + 6 ' Columna G es la columna 7, por lo que agregamos 6
                Dim celda As ExcelRange = worksheet.Cells(fila + 5, columna)
                If String.IsNullOrEmpty(celda.Text) Then
                    celda.Value = CDbl("7,5")
                    celda.Style.Font.Color.SetColor(System.Drawing.Color.Black)
                End If
            End If
        Else
            EscribirHorasNoStaff(worksheet, dr, fila, dia, fechaComparar)
        End If
    End Sub

    Private Sub EscribirHorasNoStaff(ByVal worksheet As ExcelWorksheet, ByVal dr As DataRow, ByVal fila As Integer, ByVal dia As Integer, ByVal fechaComparar As Date)
        Dim dt As New DataTable
        Dim f As New Filter
        f.Add("FechaParte", FilterOperator.Equal, fechaComparar)
        f.Add("IDOperario", FilterOperator.Equal, dr("EXP."))
        dt = New BE.DataEngine().Filter(tablaDatos, f)

        If dt.Rows.Count > 0 Then
            Dim IDCausa As String = Nz(dt.Rows(0)("IDCausa").ToString, "")
            If Len(IDCausa) <> 0 Then
                EscribirIDCausa(worksheet, fila, dia, IDCausa)
            Else
                EscribirHorasProductivas(worksheet, dt, fila, dia, fechaComparar)
            End If

        Else
            If EsDiaLaboral(fechaComparar) Then
                Dim columna As Integer = dia + 6 ' Columna G es la columna 7, por lo que agregamos 6
                Dim celda As ExcelRange = worksheet.Cells(fila + 5, columna)
                celda.Value = CDbl("0.0")
                celda.Style.Font.Color.SetColor(System.Drawing.Color.Black)
            End If
        End If
    End Sub

    Private Sub EscribirIDCausa(ByVal worksheet As ExcelWorksheet, ByVal fila As Integer, ByVal dia As Integer, ByVal IDCausa As String)
        Dim columna As Integer = dia + 6 ' Columna G es la columna 7, por lo que agregamos 6
        Dim celda As ExcelRange = worksheet.Cells(fila + 5, columna)
        celda.Value = IDCausa

        Select Case IDCausa
            Case "B"
                celda.Style.Font.Color.SetColor(System.Drawing.Color.Orange)
            Case "UA"
                celda.Style.Font.Color.SetColor(System.Drawing.Color.Green)
            Case "V"
                celda.Style.Font.Color.SetColor(System.Drawing.Color.Red)
            Case Else
                celda.Style.Font.Color.SetColor(System.Drawing.Color.Black)
        End Select
    End Sub

    Private Sub EscribirHorasProductivas(ByVal worksheet As ExcelWorksheet, ByVal dt As DataTable, ByVal fila As Integer, ByVal dia As Integer, ByVal fechaComparar As Date)
        Dim columna As Integer = dia + 6 ' Columna G es la columna 7, por lo que agregamos 6
        Dim celda As ExcelRange = worksheet.Cells(fila + 5, columna)

        Dim horas As Double = 0

        Dim color As System.Drawing.Color = getColorTurnoTrabajador(dt)
        celda.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid
        celda.Style.Fill.BackgroundColor.SetColor(color)

        horas = dt.Rows(0)("Horas").ToString.Replace(".", ",")


        'If tipoExportacion = "ORIGINAL" Then
        '    If Not IsDBNull(dt.Rows(0)("Horas")) Then
        '        Double.TryParse(dt.Rows(0)("Horas").ToString(), horas)
        '        If EsFindeSemana(fechaComparar) Then
        '            'SI EL COLOR DE ESTA CELDA ES NARANJA(FFF4B484) Y VERDE:6,5
        '            'ROSA: FFFFC0CB
        '            'SI NO. QUE PONGA 0 Y MEJOR NULL QUE NO APAREZCA
        '            Dim colorHex As String = celda.Style.Fill.BackgroundColor.Rgb
        '            If colorHex = "FFF4B484" Or colorHex = "FF78D870" Then
        '                If horas > 6.5 Then
        '                    horas = 6.5
        '                End If
        '            Else
        '                celda.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow)
        '                horas = ""
        '            End If

        '        End If
        '    End If
        'ElseIf tipoExportacion = "TECOZAM" Then
        '    horas = dt.Rows(0)("Horas").ToString.Replace(".", ",")
        'End If


        celda.Value = horas

        If Len(dt.Rows(0)("Comentarios").ToString) <> 0 Then
            celda.AddComment(dt.Rows(0)("Comentarios").ToString)
        End If

        'AQUI TAMBIEN METO EL COLOR A LA SEGUNDA Y TERCERA PARTE DEL EXCEL
        ' Obtener la celda que está 31 posiciones a la derecha
        Dim celdaDerecha As ExcelRange = worksheet.Cells(fila + 5, columna + 31)

        ' Aplicar el color a la celda a la derecha
        celdaDerecha.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid
        celdaDerecha.Style.Fill.BackgroundColor.SetColor(color)

        ' Obtener la celda que está 31 posiciones a la derecha
        Dim celdaDerechaDerecha As ExcelRange = worksheet.Cells(fila + 5, columna + 62)

        ' Aplicar el color a la celda a la derecha
        celdaDerechaDerecha.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid
        celdaDerechaDerecha.Style.Fill.BackgroundColor.SetColor(color)
    End Sub

    Private Function EsDiaLaboral(ByVal fecha As Date) As Boolean
        Return fecha.DayOfWeek >= DayOfWeek.Monday AndAlso fecha.DayOfWeek <= DayOfWeek.Friday
    End Function
    Private Function EsFindeSemana(ByVal fecha As Date) As Boolean
        Return fecha.DayOfWeek = DayOfWeek.Saturday OrElse fecha.DayOfWeek = DayOfWeek.Sunday
    End Function

    Public Function getColorTurnoTrabajador(ByVal dt As DataTable) As System.Drawing.Color
        ' Supongamos que HoraEntrada y HoraSalida están en formato TimeSpan
        Dim turnoEntrada As TimeSpan = CType(dt.Rows(0)("HoraEntrada"), TimeSpan)
        Dim turnoSalida As TimeSpan = CType(dt.Rows(0)("HoraSalida"), TimeSpan)

        ' Convertir las cadenas de hora a TimeSpan para comparación
        Dim intervalo0Inicio As TimeSpan = TimeSpan.Parse("07:00")
        Dim intervalo0Fin As TimeSpan = TimeSpan.Parse("16:00")

        Dim intervalo1Inicio As TimeSpan = TimeSpan.Parse("07:00")
        Dim intervalo1Fin As TimeSpan = TimeSpan.Parse("14:00")

        Dim intervalo2Inicio As TimeSpan = TimeSpan.Parse("14:00")
        Dim intervalo2Fin As TimeSpan = TimeSpan.Parse("21:00")

        Dim intervalo3Inicio As TimeSpan = TimeSpan.Parse("10:30")
        Dim intervalo3Fin As TimeSpan = TimeSpan.Parse("20:00")

        Dim intervalo4Inicio As TimeSpan = TimeSpan.Parse("10:00")
        Dim intervalo4Fin As TimeSpan = TimeSpan.Parse("20:00")

        ' Comprobar en qué intervalo se encuentra el turno
        If turnoEntrada = intervalo0Inicio AndAlso turnoSalida = intervalo0Fin Then
            'MsgBox("Turno de 07:00 a 16:00")
            Return System.Drawing.Color.FromArgb(255, 255, 255)
        ElseIf turnoEntrada = intervalo1Inicio AndAlso turnoSalida = intervalo1Fin Then
            'MsgBox("Turno de 07:00 a 14:00")
            Return System.Drawing.Color.FromArgb(244, 180, 132)
        ElseIf turnoEntrada = intervalo2Inicio AndAlso turnoSalida = intervalo2Fin Then
            'MsgBox("Turno de 14:00 a 21:00")
            Return System.Drawing.Color.FromArgb(120, 216, 112)
        ElseIf turnoEntrada = intervalo3Inicio AndAlso turnoSalida = intervalo3Fin Then
            'MsgBox("Turno de 10:30 a 20:00")
            Return System.Drawing.Color.FromArgb(156, 196, 228)
        ElseIf turnoEntrada = intervalo4Inicio AndAlso turnoSalida = intervalo4Fin Then
            'MsgBox("Turno de 10:00 a 20:00")
            Return System.Drawing.Color.FromArgb(168, 164, 164)
        Else
            Return System.Drawing.Color.FromArgb(255, 192, 203)
        End If

    End Function
End Class
