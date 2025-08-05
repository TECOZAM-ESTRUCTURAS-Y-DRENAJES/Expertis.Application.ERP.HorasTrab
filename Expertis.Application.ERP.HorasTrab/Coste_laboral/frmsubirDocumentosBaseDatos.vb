Imports System.Windows.Forms
Imports System.IO
Imports OfficeOpenXml
Imports System.Data
Imports System.Data.SqlClient
Imports Solmicro.Expertis.Engine
Imports Solmicro.Expertis.Engine.DAL

Public Class frmsubirDocumentosBaseDatos

    Dim connectionString = "Data Source=stecodesarr;Initial Catalog=xTecozam50R2;User ID=sa;Password=180M296;"
    Dim baseDatos As String = "xTecozam50R2"

    Public Sub SubirDocumentacion()
        Dim ofd As New OpenFileDialog()

        ofd.Filter = "Archivos de Excel (*.xls;*.xlsx;*.xlsm)|*.xls;*.xlsx;*.xlsm"
        ofd.FilterIndex = 1
        ofd.Title = "Seleccione un fichero Excel"

        If ofd.ShowDialog() = DialogResult.OK Then

            Dim cursor As Cursor
            cursor = Cursors.WaitCursor

            Dim fichero As String = ofd.FileName
            Dim nombreFichero As String = ofd.SafeFileName
            Dim excelSeleccionado As New FileInfo(fichero)
            Dim dt As New DataTable()
            Dim dtCoincidenciasBDD As New DataTable()
            Dim dtCoincidenciasBDDIDGET As New DataTable()

            Dim mbRecibirFichero = MessageBox.Show("Has seleccionado el fichero " & nombreFichero & ". ¿Deseas continuar?", "Confirmar fichero", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

            If mbRecibirFichero = DialogResult.Yes Then
                Using package As New ExcelPackage(excelSeleccionado)
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial

                    'acceder a primera hoja
                    Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets(0)

                    Dim ncolumnas As Integer = worksheet.Dimension.End.Column
                    Dim nfilas As Integer = 1

                    While Not String.IsNullOrEmpty(worksheet.Cells(nfilas, 1).Text)
                        nfilas += 1
                    End While
                    nfilas -= 1

                    'crear estructura columnas datatable
                    For col As Integer = 1 To ncolumnas
                        Dim nombreCol As String = worksheet.Cells(1, col).Text
                        dt.Columns.Add(nombreCol)
                    Next

                    'recorrer filas excel y volcar info a dt
                    For row As Integer = 2 To nfilas
                        Dim newrow As DataRow = dt.NewRow()
                        For col As Integer = 1 To ncolumnas
                            If dt.Columns(col - 1).ColumnName = "CosteEmpresa" Or dt.Columns(col - 1).ColumnName = "Total" Then
                                'tratar para dejar solo cantidad
                                Dim costeEmpresa As String = worksheet.Cells(row, col).Text.Replace("€", "").Trim()
                                newrow(col - 1) = Convert.ToDecimal(costeEmpresa)
                            Else
                                newrow(col - 1) = worksheet.Cells(row, col).Text
                            End If
                        Next
                        dt.Rows.Add(newrow)
                    Next

                    'procesar nombre de fichero y decidir
                    Dim tipo As String = ""
                    Dim mes As String = ""
                    Dim anio As String = ""

                    ObtenerTipoMesAnio(tipo, mes, anio, nombreFichero)

                    Dim f As New Filter()

                    Using connection As New SqlConnection(connectionString)
                        connection.Open()

                        If tipo = "HORAS" Then

                            f.Add("MesNatural", FilterOperator.Equal, mes)
                            f.Add("AñoNatural", FilterOperator.Equal, anio)

                            dtCoincidenciasBDD = New BE.DataEngine().Filter(baseDatos & "..tbHorasCostesLaborales", f, "*")

                            Using bulkCopy As New SqlBulkCopy(connection)
                                bulkCopy.DestinationTableName = baseDatos & "..tbHorasCostesLaborales"

                                'mapeo de columnas para saltar la columna de ID
                                bulkCopy.ColumnMappings.Add("Empresa", "Empresa")
                                bulkCopy.ColumnMappings.Add("IDGET", "IDGET")
                                bulkCopy.ColumnMappings.Add("IDOperario", "IDOperario")
                                bulkCopy.ColumnMappings.Add("DescOperario", "DescOperario")
                                bulkCopy.ColumnMappings.Add("IDOficio", "IDOficio")
                                bulkCopy.ColumnMappings.Add("IDCategoriaProfesionalSCCP", "IDCategoriaProfesionalSCCP")
                                bulkCopy.ColumnMappings.Add("NObra", "NObra")
                                bulkCopy.ColumnMappings.Add("FechaInicio", "FechaInicio")
                                bulkCopy.ColumnMappings.Add("MesNatural", "MesNatural")
                                bulkCopy.ColumnMappings.Add("AñoNatural", "AñoNatural")
                                bulkCopy.ColumnMappings.Add("Horas", "Horas")
                                bulkCopy.ColumnMappings.Add("IDHora", "IDHora")
                                bulkCopy.ColumnMappings.Add("HorasAdministrativas", "HorasAdministrativas")
                                bulkCopy.ColumnMappings.Add("HorasBaja", "HorasBaja")
                                bulkCopy.ColumnMappings.Add("Turno", "Turno")

                                If dtCoincidenciasBDD.Rows.Count > 0 Then
                                    Dim mbBorrar = MessageBox.Show("El fichero que has seleccionado ya existe en base de datos. ¿Deseas continuar para actualizarlo? Se sobreescribirá con la nueva información", "Confirmar nueva subida de fichero existente", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

                                    If mbBorrar = DialogResult.Yes Then
                                        'documento existente en base de datos, borrar antes de volver a insertar
                                        Dim deleteCommand As String = "DELETE FROM " & baseDatos & "..tbHorasCostesLaborales WHERE MesNatural='" & mes & "' AND AñoNatural='" & anio & "'"
                                        Using Command As New SqlCommand(deleteCommand, connection)
                                            Command.ExecuteNonQuery()
                                        End Using
                                    Else
                                        MessageBox.Show("Proceso cancelado correctamente", "Información", MessageBoxButtons.OK)
                                        Exit Sub 'abortar proceso
                                    End If
                                End If
                                bulkCopy.WriteToServer(dt)
                            End Using

                        ElseIf tipo = "EXTRA" Then
                            FormatearDtExtra(dt, mes, anio)

                            f.Add("Mes", FilterOperator.Equal, mes)
                            f.Add("Anio", FilterOperator.Equal, anio)

                            dtCoincidenciasBDD = New BE.DataEngine().Filter(baseDatos & "..tbExtraCostesLaborales", f, "*")

                            Using bulkCopy As New SqlBulkCopy(connection)
                                bulkCopy.DestinationTableName = baseDatos & "..tbExtraCostesLaborales"

                                'mapeo de columnas para saltar la columna de ID
                                bulkCopy.ColumnMappings.Add("Empresa", "Empresa")
                                bulkCopy.ColumnMappings.Add("IDCategoriaProfesionalSCCP", "IDCategoriaProfesionalSCCP")
                                bulkCopy.ColumnMappings.Add("Total", "Total")
                                bulkCopy.ColumnMappings.Add("Mes", "Mes")
                                bulkCopy.ColumnMappings.Add("Anio", "Anio")

                                If dtCoincidenciasBDD.Rows.Count > 0 Then
                                    Dim mbBorrar = MessageBox.Show("El fichero que has seleccionado ya existe en base de datos. ¿Deseas continuar para actualizarlo? Se sobreescribirá con la nueva información", "Confirmar nueva subida de fichero existente", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

                                    If mbBorrar = DialogResult.Yes Then
                                        'documento existente en base de datos, borrar antes de volver a insertar
                                        Dim deleteCommand As String = "DELETE FROM " & baseDatos & "..tbExtraCostesLaborales WHERE Mes='" & mes & "' AND Anio='" & anio & "'"
                                        Using Command As New SqlCommand(deleteCommand, connection)
                                            Command.ExecuteNonQuery()
                                        End Using
                                    Else
                                        MessageBox.Show("Proceso cancelado correctamente", "Información", MessageBoxButtons.OK)
                                        Exit Sub 'abortar proceso
                                    End If
                                End If
                                bulkCopy.WriteToServer(dt)
                            End Using

                        ElseIf tipo = "REGULARIZACIONES" Then
                            Dim dtIDGET As New DataTable()
                            Dim worksheet2 As ExcelWorksheet = package.Workbook.Worksheets(1)

                            FormatearDtRegularizaciones(dt, dtIDGET, worksheet, worksheet2, mes, anio)

                            f.Add("Mes", FilterOperator.Equal, mes)
                            f.Add("Anio", FilterOperator.Equal, anio)

                            dtCoincidenciasBDD = New BE.DataEngine().Filter(baseDatos & "..tbRegularizacionesCostesLaborales", f, "*")
                            dtCoincidenciasBDDIDGET = New BE.DataEngine().Filter(baseDatos & "..tbRegularizacionesIDGETCostesLaborales", f, "*")

                            If dtCoincidenciasBDD.Rows.Count > 0 Or dtCoincidenciasBDDIDGET.Rows.Count > 0 Then
                                Dim mbBorrar = MessageBox.Show("El fichero que has seleccionado ya existe en base de datos. ¿Deseas continuar para actualizarlo? Se sobreescribirá con la nueva información", "Confirmar nueva subida de fichero existente", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

                                If mbBorrar = DialogResult.Yes Then

                                    If dtCoincidenciasBDD.Rows.Count > 0 Then
                                        Dim deleteCommand As String = "DELETE FROM " & baseDatos & "..tbRegularizacionesCostesLaborales WHERE Mes='" & mes & "' AND Anio='" & anio & "'"
                                        Using Command As New SqlCommand(deleteCommand, connection)
                                            Command.ExecuteNonQuery()
                                        End Using
                                    End If

                                    If dtCoincidenciasBDDIDGET.Rows.Count > 0 Then
                                        Dim deleteCommand As String = "DELETE FROM " & baseDatos & "..tbRegularizacionesIDGETCostesLaborales WHERE Mes='" & mes & "' AND Anio='" & anio & "'"
                                        Using Command As New SqlCommand(deleteCommand, connection)
                                            Command.ExecuteNonQuery()
                                        End Using
                                    End If
                                Else
                                    MessageBox.Show("Proceso cancelado correctamente", "Información", MessageBoxButtons.OK)
                                    Exit Sub
                                End If

                            End If

                            'primer volcado
                            Using bulkCopy As New SqlBulkCopy(connection)
                                bulkCopy.DestinationTableName = baseDatos & "..tbRegularizacionesCostesLaborales"

                                bulkCopy.ColumnMappings.Add("Empresa", "Empresa")
                                bulkCopy.ColumnMappings.Add("IDCategoriaProfesionalSCCP", "IDCategoriaProfesionalSCCP")
                                bulkCopy.ColumnMappings.Add("Total", "Total")
                                bulkCopy.ColumnMappings.Add("Observaciones", "Observaciones")
                                bulkCopy.ColumnMappings.Add("Mes", "Mes")
                                bulkCopy.ColumnMappings.Add("Anio", "Anio")


                                bulkCopy.WriteToServer(dt)

                            End Using

                            'segundo volcado
                            Using bulkCopy As New SqlBulkCopy(connection)
                                bulkCopy.DestinationTableName = baseDatos & "..tbRegularizacionesIDGETCostesLaborales"

                                bulkCopy.ColumnMappings.Add("Empresa", "Empresa")
                                bulkCopy.ColumnMappings.Add("IDGET", "IDGET")
                                bulkCopy.ColumnMappings.Add("Total", "Total")
                                bulkCopy.ColumnMappings.Add("Observaciones", "Observaciones")
                                bulkCopy.ColumnMappings.Add("Mes", "Mes")
                                bulkCopy.ColumnMappings.Add("Anio", "Anio")


                                bulkCopy.WriteToServer(dtIDGET)
                            End Using

                        ElseIf tipo = "A3" Then

                            f.Add("Mes", FilterOperator.Equal, mes)
                            f.Add("Anio", FilterOperator.Equal, anio)

                            dtCoincidenciasBDD = New BE.DataEngine().Filter(baseDatos & "..tbA3CostesLaborales", f, "*")

                            Using bulkCopy As New SqlBulkCopy(connection)
                                bulkCopy.DestinationTableName = baseDatos & "..tbA3CostesLaborales"

                                'mapeo de columnas para saltar la columna de ID
                                bulkCopy.ColumnMappings.Add("Empresa", "Empresa")
                                bulkCopy.ColumnMappings.Add("IDGET", "IDGET")
                                bulkCopy.ColumnMappings.Add("IDOperario", "IDOperario")
                                bulkCopy.ColumnMappings.Add("DescOperario", "DescOperario")
                                bulkCopy.ColumnMappings.Add("Mes", "Mes")
                                bulkCopy.ColumnMappings.Add("Anio", "Anio")
                                bulkCopy.ColumnMappings.Add("CosteEmpresa", "CosteEmpresa")
                                bulkCopy.ColumnMappings.Add("IDCategoriaProfesionalSCCP", "IDCategoriaProfesionalSCCP")
                                bulkCopy.ColumnMappings.Add("IDOficio", "IDOficio")
                                bulkCopy.ColumnMappings.Add("NObra", "NObra")

                                If dtCoincidenciasBDD.Rows.Count > 0 Then
                                    Dim mbBorrar = MessageBox.Show("El fichero que has seleccionado ya existe en base de datos. ¿Deseas continuar para actualizarlo? Se sobreescribirá con la nueva información", "Confirmar nueva subida de fichero existente", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

                                    If mbBorrar = DialogResult.Yes Then
                                        'documento existente en base de datos, borrar antes de volver a insertar
                                        Dim deleteCommand As String = "DELETE FROM " & baseDatos & "..tbA3CostesLaborales WHERE Mes='" & mes & "' AND Anio='" & anio & "'"
                                        Using Command As New SqlCommand(deleteCommand, connection)
                                            Command.ExecuteNonQuery()
                                        End Using
                                    Else
                                        MessageBox.Show("Proceso cancelado correctamente", "Información", MessageBoxButtons.OK)
                                        Exit Sub 'abortar proceso
                                    End If
                                End If
                                'insertar en base de datos
                                bulkCopy.WriteToServer(dt)
                            End Using
                        End If
                    End Using
                End Using
            End If
            MsgBox("Fichero cargado correctamente.", MsgBoxStyle.Information, "Exito")
            cursor = Cursors.Default
        End If
    End Sub

    Public Sub FormatearDtExtra(ByRef dt As DataTable, ByVal mes As Integer, ByVal anio As Integer)
        'agregar columnas
        dt.Columns.Add("Mes")
        dt.Columns.Add("Anio")

        'rellenar con datos
        For Each row As DataRow In dt.Rows
            For Each col As DataColumn In dt.Columns
                If col.ColumnName.Contains("Mes") Then
                    row(col) = mes
                ElseIf col.ColumnName.Contains("Anio") Then
                    row(col) = anio
                End If
            Next
        Next
    End Sub

    Public Sub FormatearDtRegularizaciones(ByRef dt As DataTable, ByRef dtIDGET As DataTable, ByVal worksheet As ExcelWorksheet, ByVal worksheet2 As ExcelWorksheet, ByVal mes As Integer, ByVal anio As Integer)
        'renombrar columna de observaciones

        If dt.Columns.Count > 3 Then
            ' Si ya existe una columna en la posición 3, le cambiamos el nombre
            dt.Columns(3).ColumnName = "Observaciones"
        ElseIf Not dt.Columns.Contains("Observaciones") Then
            ' Si no hay suficiente columnas, agregamos la columna "Observaciones"
            Dim colObservaciones As New DataColumn("Observaciones", GetType(String))
            dt.Columns.Add(colObservaciones)
        End If


        'agregar columnas
        dt.Columns.Add("Mes")
        dt.Columns.Add("Anio")

        'rellenar con datos
        For Each row As DataRow In dt.Rows
            For Each col As DataColumn In dt.Columns
                If col.ColumnName.Contains("Mes") Then
                    row(col) = mes
                ElseIf col.ColumnName.Contains("Anio") Then
                    row(col) = anio
                End If
            Next
        Next

        'crear segundo datatable
        Dim ncolumnas As Integer = worksheet2.Dimension.End.Column
        Dim nfilas As Integer = worksheet2.Dimension.End.Row

        'crear estructura columnas datatable
        For col As Integer = 1 To ncolumnas
            Dim nombreCol As String = worksheet2.Cells(1, col).Text
            dtIDGET.Columns.Add(nombreCol)
        Next

        'recorrer filas excel y volcar info a dt
        For row As Integer = 2 To nfilas
            Dim newrow As DataRow = dtIDGET.NewRow()
            For col As Integer = 1 To ncolumnas
                If Not String.IsNullOrEmpty(worksheet2.Cells(row, col).Text) Then
                    If dtIDGET.Columns(col - 1).ColumnName = "Total" Then
                        'tratar para dejar solo cantidad
                        Dim Total As String = worksheet2.Cells(row, col).Text.Replace("€", "").Trim()
                        newrow(col - 1) = Convert.ToDecimal(Total)
                    Else
                        newrow(col - 1) = worksheet2.Cells(row, col).Text
                    End If
                Else
                    newrow(col - 1) = DBNull.Value
                End If
            Next
            dtIDGET.Rows.Add(newrow)
        Next

        'añadir columnas mes y anio a nueva dt
        dtIDGET.Columns.Add("Mes")
        dtIDGET.Columns.Add("Anio")

        For Each row As DataRow In dtIDGET.Rows
            For Each col As DataColumn In dtIDGET.Columns
                If col.ColumnName.Contains("Mes") Then
                    row(col) = mes
                ElseIf col.ColumnName.Contains("Anio") Then
                    row(col) = anio
                End If
            Next
        Next
    End Sub

    Public Sub ObtenerTipoMesAnio(ByRef tipo As String, ByRef mes As String, ByRef anio As String, ByVal nombreFichero As String)
        'obtener mes y tipo
        Dim partes() As String = nombreFichero.Split(" "c)
        mes = partes(0)
        tipo = partes(1)

        'obtener año
        Dim puntoExtension As Integer = nombreFichero.LastIndexOf(".")
        anio = "20" & nombreFichero.Substring(nombreFichero.LastIndexOf(" "c, puntoExtension - 1) + 3, 2)

    End Sub
End Class