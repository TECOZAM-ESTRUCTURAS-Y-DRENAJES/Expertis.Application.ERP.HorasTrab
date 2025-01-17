﻿Imports Solmicro.Expertis.Engine
Imports Solmicro.Expertis.Engine.UI
Imports Solmicro.Expertis.Engine.DAL
Imports System.Windows.Forms
Imports Solmicro.Expertis.Business.ClasesTecozam
Imports OfficeOpenXml
Imports System.IO

Public Class CIMntoTrabObraMes

    Private Sub CIMntoTrabObraMes_QueryExecuting(ByVal sender As Object, ByRef e As Solmicro.Expertis.Engine.UI.QueryExecutingEventArgs) Handles Me.QueryExecuting
        e.Filter.Add("NObra", FilterOperator.Equal, advNObra.Text)
        e.Filter.Add("IDOperario", FilterOperator.Equal, advIDOperario.Text)
        e.Filter.Add("Mes", FilterOperator.Equal, cmbmes.Value)
        e.Filter.Add("Anio", FilterOperator.Equal, cmbanio.Value)
        e.Filter.Add("FechaInicio", FilterOperator.GreaterThanOrEqual, clbFecha.Value, FilterType.DateTime)
        e.Filter.Add("FechaInicio", FilterOperator.LessThanOrEqual, clbFecha1.Value, FilterType.DateTime)
    End Sub

    Public Sub cargarComboMes()
        Dim dtcombo As New DataTable
        dtcombo.Columns.Add("Codigo")
        dtcombo.Columns.Add("Descripcion")

        Dim dr As DataRow

        dr = dtcombo.NewRow()
        dr("Codigo") = "01"
        dr("Descripcion") = "Enero"
        dtcombo.Rows.Add(dr)

        dr = dtcombo.NewRow()
        dr("Codigo") = "02"
        dr("Descripcion") = "Febrero"
        dtcombo.Rows.Add(dr)

        dr = dtcombo.NewRow()
        dr("Codigo") = "03"
        dr("Descripcion") = "Marzo"
        dtcombo.Rows.Add(dr)

        dr = dtcombo.NewRow()
        dr("Codigo") = "04"
        dr("Descripcion") = "Abril"
        dtcombo.Rows.Add(dr)

        dr = dtcombo.NewRow()
        dr("Codigo") = "05"
        dr("Descripcion") = "Mayo"
        dtcombo.Rows.Add(dr)

        dr = dtcombo.NewRow()
        dr("Codigo") = "06"
        dr("Descripcion") = "Junio"
        dtcombo.Rows.Add(dr)

        dr = dtcombo.NewRow()
        dr("Codigo") = "07"
        dr("Descripcion") = "Julio"
        dtcombo.Rows.Add(dr)

        dr = dtcombo.NewRow()
        dr("Codigo") = "08"
        dr("Descripcion") = "Agosto"
        dtcombo.Rows.Add(dr)

        dr = dtcombo.NewRow()
        dr("Codigo") = "09"
        dr("Descripcion") = "Septiembre"
        dtcombo.Rows.Add(dr)

        dr = dtcombo.NewRow()
        dr("Codigo") = "10"
        dr("Descripcion") = "Octubre"
        dtcombo.Rows.Add(dr)

        dr = dtcombo.NewRow()
        dr("Codigo") = "11"
        dr("Descripcion") = "Noviembre"
        dtcombo.Rows.Add(dr)

        dr = dtcombo.NewRow()
        dr("Codigo") = "12"
        dr("Descripcion") = "Diciembre"
        dtcombo.Rows.Add(dr)

        cmbmes.DataSource = dtcombo
        cmbmes.ValueMember = "Codigo"
        cmbmes.DisplayMember = "Descripcion"
    End Sub
    Private Sub cargarComboAnio()
        Dim dtcombo As New DataTable
        dtcombo.Columns.Add("Anio")

        Dim dr As DataRow

        For i As Integer = 0 To 10
            Dim j As Integer = Year(Today)
            dr = dtcombo.NewRow()
            dr("Anio") = j - i
            dtcombo.Rows.Add(dr)
        Next
        cmbanio.DataSource = dtcombo

    End Sub

    Public Sub New()

        ' Llamada necesaria para el Diseñador de Windows Forms.
        InitializeComponent()

        ' Agregue cualquier inicialización después de la llamada a InitializeComponent().
        cargarComboAnio()
        cargarComboMes()
        LoadToolbarActions()

    End Sub

    Private Sub LoadToolbarActions()
        Try
            With Me.FormActions
                '.Add("Calcular dias por obra Entre fechas.", AddressOf calcular)
                .Add("Exportar excel horas acumulado.", AddressOf exportarExcel)
                '.Add("Leer un Fichero", AddressOf LeerFichero)
            End With
        Catch ex As Exception
            'ExpertisApp.GenerateMessage(ex.Message)
        End Try
    End Sub

    Public Sub exportarExcel()
        Dim fecha1 As String = clbFecha1.Value
        Dim fecha2 As String = clbFecha.Value

        If advNObra.Text.ToString.Length = 0 Or fecha1.Length = 0 Or fecha2.Length = 0 Then
            MsgBox("La obra y las fechas son datos obligatorios")
        Else
            exportaExcel(fecha1, fecha2)
        End If
    End Sub
    Public Sub exportaExcel(ByVal fecha1 As String, ByVal fecha2 As String)
        Dim dtObras As New DataTable
        Dim sql As String
        sql = "SELECT IDOperario, DescOperario, SUM(Horas) AS TotalHoras FROM  vSistLabListadoTrabajadoresObraMes "
        sql &= "WHERE FechaInicio >= '" & fecha2 & "' AND FechaInicio <= '" & fecha1 & "' AND NObra = '" & advNObra.Text.ToString & "'"
        sql &= "GROUP BY IDOperario, DescOperario"

        Dim s As New Solmicro.Expertis.Business.ClasesTecozam.ControlArticuloNSerie
        dtObras = s.EjecutarSqlSelect(sql)
        GeneraExcelHoras(dtObras)
    End Sub
    Public Sub GeneraExcelHoras(ByVal dtFinal As DataTable)

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial

        Dim saveFileDialog As New SaveFileDialog()

        ' Establecer el filtro para el cuadro de diálogo (solo archivos .xlsx)
        saveFileDialog.Filter = "Archivos de Excel (*.xlsx)|*.xlsx"
        saveFileDialog.FileName = " HORAS - " & advNObra.Text.ToString & ".xlsx" ' Establecer el nombre predeterminado del archivo

        ' Mostrar el cuadro de diálogo para que el usuario elija dónde guardar el archivo
        If saveFileDialog.ShowDialog() = DialogResult.OK Then
            ' Obtener la ruta seleccionada por el usuario
            Dim ruta As New FileInfo(saveFileDialog.FileName)
            Dim rutaCadena As String = ruta.FullName

            ' Ahora 'rutaCadena' tiene la ruta completa del archivo seleccionado
            MessageBox.Show("El archivo se ha guardado en: " & rutaCadena)
        End If


        Using package As New ExcelPackage(saveFileDialog.FileName)
            ' Crear una hoja de cálculo y obtener una referencia a ella.
            Dim worksheet = package.Workbook.Worksheets.Add("TOTAL HORAS")

            ' Copiar los datos de la DataTable a la hoja de cálculo.
            worksheet.Cells("A1").LoadFromDataTable(dtFinal, True)

            ' Aplicar formato negrita a la fila 1
            Dim fila1 As ExcelRange = worksheet.Cells(1, 1, 1, worksheet.Dimension.End.Column)
            fila1.Style.Font.Bold = True
            ' Congelar la primera columna
            worksheet.View.FreezePanes(2, 1)
            ' Guardar el archivo de Excel.
            package.Save()
        End Using

    End Sub
    Public Sub calcular()
        Dim fechadesde As Date
        Dim fechahasta As Date
        fechadesde = Nz(clbFecha.Value.ToString, "01/01/2000")
        fechahasta = Nz(clbFecha1.Value.ToString, "31/12/2050")

        creardt(fechadesde, fechahasta)
    End Sub
    Public Sub creardt(ByVal fechadesde As Date, ByVal fechahasta As Date)
        'Creacion Estructura Tabla
        Dim dt As DataTable
        dt = estructuraTabla()
        'Relleno tabla
        rellenoTabla(dt, fechadesde, fechahasta)

    End Sub
    Public Function estructuraTabla() As DataTable
        Dim dt As New DataTable
        Dim dc As New DataColumn("CodigoObra")
        dt.Columns.Add(dc)
        dc = New DataColumn("Dias")
        dt.Columns.Add(dc)
        Return dt
    End Function
    Public Sub rellenoTabla(ByVal dt As DataTable, ByVal fechadesde As Date, ByVal fechahasta As Date)
        'Creo todas las variables
        Dim dtError As New DataTable
        dtError = estructuraTabla()
        Try
            calculaObras(dt, fechadesde, fechahasta)

        Catch ex As Exception
            Grid.DataSource = dtError
            MsgBox("No existe ningún registro con estas características")
        End Try
    End Sub

    Public Sub calculaObras(ByVal dt As DataTable, ByVal fechadesde As Date, ByVal fechahasta As Date)

        Dim dtObras As New DataTable
        Dim dtHoras As New DataTable

        Dim sql As String = "select distinct(NObra) from vSistLabListadoTrabajadoresObraMes where FechaInicio>='" & fechadesde & "' And FechaInicio<='" & fechahasta & "'"
        Dim s As New Solmicro.Expertis.Business.ClasesTecozam.ControlArticuloNSerie
        dtObras = s.EjecutarSqlSelect(sql)
        Dim sql2 As String = ""

        Dim CodigoObra As String = ""
        Dim Dias As String = ""
        For Each dr As DataRow In dtObras.Rows

            CodigoObra = dr("NObra")

            sql2 = "select count(distinct(FechaInicio)) as Dias from vSistLabListadoTrabajadoresObraMes  where FechaInicio>='" & fechadesde & "' And FechaInicio<='" & fechahasta & "' and NObra='" & CodigoObra & "'"
            dtHoras = s.EjecutarSqlSelect(sql2)
            Dias = dtHoras.Rows(0)("Dias").ToString

            Dim drFinal As DataRow
            drFinal = dt.NewRow
            drFinal("NObra") = CodigoObra
            drFinal("Horas") = Dias
            dt.Rows.Add(drFinal)

        Next

        Grid.DataSource = dt
        dt = Nothing
    End Sub

#Region "Informes"

    Private Function generarCuadranteObrasDel20Al21(ByVal mes As Integer, ByVal anio As Integer, ByVal informe As String)
        Dim rp As New Report(informe)
        Dim dtObrasMes As New DataTable
        Dim strSelect1 As String = "select distinct nobra,descobra from vSistLabListadoTrabajadoresObraMes where anio=" & anio & " and (mes=" & mes & "OR mes=" & mes - 1 & ")"

        'Para obtener del 21 al 20
        Dim desde As Date
        Dim hasta As Date

        Dim mesanterior As String = ""
        Dim mesactual As String = ""
        Select Case mes
            Case "01"
                'Meses
                mesanterior = 12
                mesactual = mes
                'Fechas
                desde = CDate("21/12/" & CStr(anio - 1))
                hasta = CDate("20/" & mes & "/" & CStr(anio))
            Case Else
                'Meses
                Dim mesNum As Integer = CInt(mes)

                mesanterior = CStr(mesNum - 1)
                mesactual = CStr(mesNum)
                'Fechas
                desde = CDate("21/" & CStr(mesNum - 1) & "/" & CStr(anio))
                hasta = CDate("20/" & CStr(mesNum) & "/" & CStr(anio))
        End Select

        Dim DE As New BE.DataEngine
        'Listado de obras que tienen horas
        dtObrasMes = DE.RetrieveData(strSelect1, , , , False)
        Dim a As Integer = 0
        Dim strSelect2 As String = ""

        Try

            For Each drObrasMes As DataRow In dtObrasMes.Rows
                If a > 0 Then
                    strSelect2 &= " union "
                End If
                Dim obra1 As String = drObrasMes(0)
                Dim dobra As String = drObrasMes(1)

                strSelect2 &= "select b.fecha,b.Numdiasemana,b.diasemana,isnull(a.nobra,'" & obra1 & "') as NumObra,isnull(a.descobra,'" & dobra & "') as DObra,a.*"
                strSelect2 &= " from (select * from vSistLabListadoTrabajadoresObraMes where anio=" & anio & " and nobra='" & obra1 & "' AND FechaInicio>= '" & desde & "' AND FechaInicio<= '" & hasta & "') as a "
                strSelect2 &= " full outer join (select * from tiempo where Fecha>= '" & desde & "' AND Fecha<= '" & hasta & "') as b on a.fechainicio=b.fecha"
                a = a + 1

            Next
            Dim dt As New DataTable
            dt = DE.RetrieveData(strSelect2, , , "fecha,numobra", False)
            rp.DataSource = dt
            rp.Formulas("desde").Text = Format(desde, "dd/MM/yyyy")
            rp.Formulas("hasta").Text = Format(hasta, "dd/MM/yyyy")
            ExpertisApp.OpenReport(rp)


        Catch ex As SqlClient.SqlException


            MsgBox(ex.Message)

        End Try

    End Function

    Private Function generarCuadranteObras(ByVal mes As Integer, ByVal anio As Integer, ByVal informe As String)
        Dim rp As New Report(informe)
        Dim dtObrasMes As New DataTable
        Dim strSelect1 As String = "select distinct nobra,descobra from vSistLabListadoTrabajadoresObraMes where anio=" & anio & " and mes=" & mes
        Dim mesTiempo As String
        Dim mesT As String
        If (mes.ToString).Length = 1 Then
            mesTiempo = "0" & mes.ToString
            mesT = anio & "-" & mesTiempo
        Else
            mesTiempo = mes.ToString
            mesT = anio & "-" & mesTiempo
        End If
        Dim DE As New BE.DataEngine
        dtObrasMes = DE.RetrieveData(strSelect1, , , , False)
        'dtObrasMes = AdminData.GetData(strSelect1, False)
        Dim a As Integer = 0
        Dim strSelect2 As String = ""

        Try

            For Each drObrasMes As DataRow In dtObrasMes.Rows
                If a > 0 Then
                    strSelect2 &= " union "
                End If
                Dim obra1 As String = drObrasMes(0)
                Dim dobra As String = drObrasMes(1)

                strSelect2 &= "select b.fecha,b.Numdiasemana,b.diasemana,isnull(a.nobra,'" & obra1 & "') as NumObra,isnull(a.descobra,'" & dobra & "') as DObra,a.*"
                strSelect2 &= " from (select * from vSistLabListadoTrabajadoresObraMes where anio=" & anio & " and mes=" & mes & " and nobra='" & obra1 & "') as a "
                strSelect2 &= " full outer join (select * from tiempo where mes='" & mesT & "') as b on a.fechainicio=b.fecha"
                a = a + 1

            Next
            'strSelect2 &= " order by fecha,numobra"
            'MsgBox(strSelect2)
            'rp.DataSource = AdminData.GetData(strSelect2, False)

            'David Velasco 18/12/23 Solucion Error Dias FESTIVOS
            Dim dtGente As New DataTable
            dtGente = DE.RetrieveData(strSelect2, , , "fecha,numobra", False)


            Dim Fecha1 As String
            Fecha1 = DevuelvePrimerDia(mes, anio)

            Dim Fecha2 As String
            Fecha2 = DevuelveUltimoDia(mes, anio)

            Dim dtFestivos As New DataTable
            Dim f As New Filter
            f.Add("TipoDia", FilterOperator.Equal, 1)
            f.Add("Fecha", FilterOperator.GreaterThanOrEqual, Fecha1)
            f.Add("Fecha", FilterOperator.LessThanOrEqual, Fecha2)
            f.Add("IDCentro", FilterOperator.Equal, "00")
            dtFestivos = New BE.DataEngine().Filter("tbCalendarioCentro", f)


            For Each dr As DataRow In dtGente.Rows
                For Each filaFecha As DataRow In dtFestivos.Rows
                    If dr("fecha") = filaFecha("Fecha") Then
                        dr("Horas") = 0
                    End If
                Next
            Next


            rp.DataSource = dtGente
            'DE.RetrieveData(
            ExpertisApp.OpenReport(rp)
        Catch ex As SqlClient.SqlException
            MsgBox(ex.Message)
        End Try

    End Function
    Public Function DevuelvePrimerDia(ByVal mes As String, ByVal anio As String) As DateTime
        Dim txtSql As String
        txtSql = "SELECT MIN(fecha) AS Fecha FROM tbMaestroFechas WHERE YEAR(fecha) = " & anio & " AND MONTH(fecha) = " & mes & ""
        Dim s As New Solmicro.Expertis.Business.ClasesTecozam.ControlArticuloNSerie
        Dim dt As DataTable
        dt = s.EjecutarSqlSelect(txtSql)
        Return dt.Rows(0)("Fecha").ToString
    End Function

    Public Function DevuelveUltimoDia(ByVal mes As String, ByVal anio As String) As DateTime
        Dim txtSql As String
        txtSql = "SELECT MAX(fecha) AS Fecha FROM tbMaestroFechas WHERE YEAR(fecha) = " & anio & " AND MONTH(fecha) = " & mes & ""
        Dim s As New Solmicro.Expertis.Business.ClasesTecozam.ControlArticuloNSerie
        Dim dt As DataTable
        dt = s.EjecutarSqlSelect(txtSql)
        Return dt.Rows(0)("Fecha").ToString
    End Function

    Private Sub CIMntoTrabObraMes_SetReportDesignObjects(ByVal sender As Object, ByVal e As Solmicro.Expertis.Engine.UI.ReportDesignObjectsEventArgs) Handles MyBase.SetReportDesignObjects
        Dim mes As Integer
        Dim anio As Integer
        Dim Obra As String

        If e.Alias = "CUADOBRAS" Or e.Alias = "CUADOBRASHORAS" Then
            Dim frm As New frmDatosInforme
            Dim informe As String = e.Alias
            frm.ShowDialog()
            mes = frm.mes
            anio = frm.anio
            If frm.blEstado = True Then
                MessageBox.Show("Proceso Cancelado", "Información", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                e.Cancel = True
                Exit Sub
            End If
            generarCuadranteObras(mes, anio, informe)
            e.Cancel = True
            'David Velasco 15/05/22
            'Este informe saca las horas del 20 del mes anterio al 21 del mes actual
        ElseIf e.Alias = "PASIS2021" Then
            Dim frm As New frmDatosInforme
            Dim informe As String = e.Alias
            frm.ShowDialog()
            mes = frm.mes
            anio = frm.anio
            generarCuadranteObrasDel20Al21(mes, anio, informe)
            e.Cancel = True
            'Fin David Velasco 15/05/2022
        ElseIf e.Alias = "JORINDI" Or e.Alias = "REGHOR" Or e.Alias = "REGHORT" Then
            Dim informe As String = e.Alias
            Dim frm As New frmDatosInforme
            frm.ShowDialog()
            mes = frm.mes
            anio = frm.anio
            'Obra = frm.nobra
            If frm.blEstado = True Then
                MessageBox.Show("Proceso Cancelado", "Información", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                e.Cancel = True
                Exit Sub
            End If
            generarCuadranteIndividual(mes, anio, informe)
            e.Cancel = True
        ElseIf e.Alias = "REGJESPERSONAL" Then

            Dim informe As String = e.Alias
            Dim frm As New frmDatosInformePersonal
            Dim IDOperario As String

            frm.ShowDialog()
            mes = frm.mes
            anio = frm.anio
            IDOperario = frm.IDOperario

            'Obra = frm.nobra
            If frm.blEstado = True Then
                MessageBox.Show("Proceso Cancelado", "Información", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                e.Cancel = True
                Exit Sub
            End If
            generarCuadranteIndividualPersonal(mes, anio, informe, IDOperario)
            e.Cancel = True
        ElseIf e.Alias = "REGHOROF" Then
            Dim informe As String = e.Alias
            Dim frm As New frmDatosInforme
            frm.ShowDialog()
            mes = frm.mes
            anio = frm.anio
            'Obra = frm.nobra
            If frm.blEstado = True Then
                MessageBox.Show("Proceso Cancelado", "Información", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                e.Cancel = True
                Exit Sub
            End If
            generarCuadranteIndividualOF(mes, anio, informe)
            e.Cancel = True
            'David Velasco Herrero 11/7/23
            'INFORME OPERARIO CATEGORIA PROFESIONAL
        ElseIf e.Alias = "INFOO" Then
            Dim codigo As String
            Dim informe As String = e.Alias
            Dim frm As New frmFiltroOficio
            frm.ShowDialog()
            codigo = frm.codigo

            Dim dt As New DataTable
            Dim filtro As New Filter
            filtro.Add("CategoriaProfesional", FilterOperator.Equal, codigo)
            dt = New BE.DataEngine().Filter("vOperarioOficio", filtro)
            Dim rp As New Report(informe)
            rp.DataSource = dt
            ExpertisApp.OpenReport(rp)

        End If
    End Sub

    'Private Function generarCuadranteIndividual(ByVal mes As Integer, ByVal anio As Integer, ByVal NumObra As String)
    '    Dim rp As New Report("JORINDI")
    '    'Dim dtObrasMes As New DataTable
    '    Dim dtTrabMes As New DataTable
    '    'Dim strSelect1 As String = "select distinct nobra,descobra from vSistLabListadoTrabajadoresObraMes where anio=" & anio & " and mes=" & mes
    '    Dim strSelect3 As String = ""
    '    strSelect3 &= "select b.idOperario, b.descOperario, b.FechaInicio, a.Nobra from"
    '    strSelect3 &= "(select idOperario, nobra, FechaInicio from vSistLabListadoTrabajadoresObraMes where anio=" & anio & " and mes=" & mes & ") as a inner join"
    '    strSelect3 &= "(select distinct idOperario,descoperario,max(FechaInicio) as FechaInicio from vSistLabListadoTrabajadoresObraMes where anio='" & anio & "' and mes='" & mes & "' group by  idOperario,descoperario ) as b "
    '    strSelect3 &= "on b.IDOperario=a.IDOperario and b.FechaInicio =a.FechaInicio where NObra = '" & NumObra & "' order by IDOperario"

    '    Dim mesTiempo As String
    '    Dim mesT As String
    '    If (mes.ToString).Length = 1 Then
    '        mesTiempo = "0" & mes.ToString
    '        mesT = anio & "-" & mesTiempo
    '    Else
    '        mesTiempo = mes.ToString
    '        mesT = anio & "-" & mesTiempo
    '    End If

    '    'dtObrasMes = AdminData.GetData(strSelect1, False)
    '    dtTrabMes = AdminData.GetData(strSelect3, False)



    '    Dim a As Integer = 0
    '    Dim strSelect2 As String = ""
    '    Dim itrab As String
    '    Dim dtrab As String
    '    Try
    '        If dtTrabMes.Rows.Count > 0 Then
    '            For Each drTrabMes As DataRow In dtTrabMes.Rows

    '                If a > 0 Then
    '                    strSelect2 &= " union "
    '                End If
    '                itrab = drTrabMes(0)
    '                dtrab = drTrabMes(1)

    '                strSelect2 &= "select b.fecha,b.Numdiasemana,b.diasemana,isnull(a.IDOperario,'" & itrab & "') as IOperario,isnull(a.DescOperario,'" & dtrab & "') as DOperario,'" & NumObra & "' as NumObra,a.*"
    '                strSelect2 &= " from (select * from vSistLabListadoTrabajadoresObraMes where anio=" & anio & " and mes=" & mes & " and IDOperario='" & itrab & "') as a "
    '                strSelect2 &= " full outer join (select * from tiempo where mes='" & mesT & "') as b on a.fechainicio=b.fecha"

    '                a = a + 1

    '            Next
    '            'strSelect2 &= " order by fecha,idOperario"

    '            'MsgBox(strSelect2)

    '            rp.DataSource = AdminData.GetData(strSelect2, False)

    '            rp.Formulas("anio").Text = anio
    '            rp.Formulas("mes").Text = mesTiempo
    '            If (mes.ToString).Length = 1 Then
    '                rp.Formulas("Fecha Liquidacion").Text = "01/0" & mesTiempo + 1 & "/" & anio
    '            Else
    '                If mesTiempo = "12" Then
    '                    rp.Formulas("Fecha Liquidacion").Text = "01/01/" & anio + 1
    '                Else
    '                    rp.Formulas("Fecha Liquidacion").Text = "01/" & mes + 1 & "/" & anio
    '                End If
    '            End If


    '            ExpertisApp.OpenReport(rp)
    '        Else
    '            MsgBox("A fecha de liquidacion no hay ningun Trabajador en la obra " & NumObra)
    '        End If

    '    Catch ex As SqlClient.SqlException
    '        MsgBox("El error lo ha dado en el registro " & a & " y en el operario " & itrab)
    '        MsgBox(ex.Message)
    '    Catch ex As Exception
    '        MsgBox(ex.Message)


    '    End Try


    'End Function
    Private Function generarCuadranteIndividual(ByVal mes As Integer, ByVal anio As Integer, ByVal informe As String)
        Dim rp As New Report(informe)

        Dim hT As New horastrabajador

        Dim mesTiempo As String
        Dim mesT As String
        If (mes.ToString).Length = 1 Then
            mesTiempo = "0" & mes.ToString
            mesT = anio & "-" & mesTiempo
        Else
            mesTiempo = mes.ToString
            mesT = anio & "-" & mesTiempo
        End If

        Dim dtInforme = hT.datosCuadranteIndividual(mes, anio)

        'MsgBox("el informe tiene los siguientes registros " & dtInforme.rows.count)

        'David Velasco 18/12/23 Solucion Error Dias FESTIVOS
        Dim Fecha1 As String
        Fecha1 = DevuelvePrimerDia(mes, anio)

        Dim Fecha2 As String
        Fecha2 = DevuelveUltimoDia(mes, anio)

        Dim dtFestivos As New DataTable
        Dim f As New Filter
        f.Add("TipoDia", FilterOperator.Equal, 1)
        f.Add("Fecha", FilterOperator.GreaterThanOrEqual, Fecha1)
        f.Add("Fecha", FilterOperator.LessThanOrEqual, Fecha2)
        f.Add("IDCentro", FilterOperator.Equal, "00")
        dtFestivos = New BE.DataEngine().Filter("tbCalendarioCentro", f)


        For Each dr As DataRow In dtInforme.Rows
            For Each filaFecha As DataRow In dtFestivos.Rows
                If dr("fecha") = filaFecha("Fecha") Then
                    dr("Horas") = 0
                End If
            Next
        Next



        rp.DataSource = dtInforme

        rp.Formulas("anio").Text = anio
        rp.Formulas("mes").Text = mesTiempo
        If (mes.ToString).Length = 1 Then
            rp.Formulas("Fecha Liquidacion").Text = "01/0" & mesTiempo + 1 & "/" & anio
        Else
            If mesTiempo = "12" Then
                rp.Formulas("Fecha Liquidacion").Text = "01/01/" & anio + 1
            Else
                rp.Formulas("Fecha Liquidacion").Text = "01/" & mes + 1 & "/" & anio
            End If
        End If

        Try

            ExpertisApp.OpenReport(rp)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        'Dim dtObrasMes As New DataTable
        'Dim dtTrabMes As New DataTable
        ''Dim strSelect1 As String = "select distinct nobra,descobra from vSistLabListadoTrabajadoresObraMes where anio=" & anio & " and mes=" & mes
        '' Dim strSelect3 As String = ""
        ''strSelect3 &= "select b.idOperario, b.descOperario, b.FechaInicio, a.Nobra from"
        ''strSelect3 &= "(select idOperario, nobra, FechaInicio from vSistLabListadoTrabajadoresObraMes where anio=" & anio & " and mes=" & mes & ") as a inner join"
        ''strSelect3 &= "(select distinct idOperario,descoperario,max(FechaInicio) as FechaInicio from vSistLabListadoTrabajadoresObraMes where anio='" & anio & "' and mes='" & mes & "' group by  idOperario,descoperario ) as b "
        ''strSelect3 &= "on b.IDOperario=a.IDOperario and b.FechaInicio =a.FechaInicio"
        '' strSelect3 = "select distinct idOperario,descoperario,max(FechaInicio) as FechaInicio from vSistLabListadoTrabajadoresObraMes where anio=" & anio & " and mes=" & mes & "  group by  idOperario,descoperario "

        'Dim mesTiempo As String
        'Dim mesT As String
        'If (mes.ToString).Length = 1 Then
        '    mesTiempo = "0" & mes.ToString
        '    mesT = anio & "-" & mesTiempo
        'Else
        '    mesTiempo = mes.ToString
        '    mesT = anio & "-" & mesTiempo
        'End If

        ''Dim DE As New BE.DataEngine
        ''dtObrasMes = DE.RetrieveData(strSelect1, , , , False)

        ''dtObrasMes = AdminData.GetData(strSelect1, False)
        ''dtTrabMes = DE.RetrieveData(strSelect3, , , "idoperario", False)

        ''DE.RetrieveData(

        'Dim a As Integer = 0
        'Dim strSelect2 As String = ""
        'Dim itrab As String
        'Dim dtrab As String
        'Try
        '    If dtTrabMes.Rows.Count > 0 Then
        '        For Each drTrabMes As DataRow In dtTrabMes.Rows

        '            If a > 0 Then
        '                strSelect2 &= " union "
        '            End If
        '            itrab = drTrabMes(0)
        '            dtrab = drTrabMes(1)

        '            'If itrab = "T984" Then

        '            '    MsgBox(itrab)
        '            'End If

        '            strSelect2 &= "select (select id from tbDatosEmpresa) as idEmpresa, b.fecha,b.Numdiasemana,b.diasemana,isnull(a.IDOperario,'" & itrab & "') as IOperario,isnull(a.DescOperario,'" & dtrab & "') as DOperario,a.*"
        '            strSelect2 &= " from (select distinct FechaInicio, IDOperario,DescOperario,sum(horas) as horas from vSistLabOperarioObraMesSObra where anio=" & anio & " and mes=" & mes & " and IDOperario='" & itrab & "' group by FechaInicio, IDOperario,DescOperario) as a "
        '            strSelect2 &= " full outer join (select * from tiempo where mes='" & mesT & "') as b on a.fechainicio=b.fecha"

        '            a = a + 1

        '        Next
        '        'strSelect2 &= " order by fecha,idOperario"

        '        'MsgBox(strSelect2)

        '        'rp.DataSource = DE.RetrieveData(strSelect2, , , "fecha,idOperario", False)

        '        rp.Formulas("anio").Text = anio
        '        rp.Formulas("mes").Text = mesTiempo
        '        If (mes.ToString).Length = 1 Then
        '            rp.Formulas("Fecha Liquidacion").Text = "01/0" & mesTiempo + 1 & "/" & anio
        '        Else
        '            If mesTiempo = "12" Then
        '                rp.Formulas("Fecha Liquidacion").Text = "01/01/" & anio + 1
        '            Else
        '                rp.Formulas("Fecha Liquidacion").Text = "01/" & mes + 1 & "/" & anio
        '            End If
        '        End If


        '        ExpertisApp.OpenReport(rp)
        '    Else
        '        'MsgBox("A fecha de liquidacion no hay ningun Trabajador en la obra " & NumObra)
        '        MsgBox("No se ha trabajado este mes")
        '    End If

        'Catch ex As SqlClient.SqlException
        '    MsgBox("El error lo ha dado en el registro " & a & " y en el operario " & itrab)
        '    MsgBox(ex.Message)
        'Catch ex As Exception
        '    MsgBox(ex.Message)


        'End Try


    End Function
    Private Function generarCuadranteIndividualPersonal(ByVal mes As Integer, ByVal anio As Integer, ByVal informe As String, ByVal IDOperario As String)
        Dim rp As New Report(informe)

        Dim mesTiempo As String
        Dim mesT As String
        If (mes.ToString).Length = 1 Then
            mesTiempo = "0" & mes.ToString
            mesT = anio & "-" & mesTiempo
        Else
            mesTiempo = mes.ToString
            mesT = anio & "-" & mesTiempo
        End If

        Dim filtro As New Filter
        filtro.Add("Mes", FilterOperator.Equal, mes)
        filtro.Add("Anio", FilterOperator.Equal, anio)

        Dim dtInforme = New BE.DataEngine().Filter("vMaestroFechas", filtro)

        'MsgBox("el informe tiene los siguientes registros " & dtInforme.rows.count)
        Dim filtroPers As New Filter
        filtroPers.Add("IDOperario", FilterOperator.Equal, IDOperario)
        Dim dtPersonal As DataTable = New BE.DataEngine().Filter("vMaestroOperarioCompleta", filtroPers)


        rp.DataSource = dtInforme
        rp.Formulas("IDOperario").Text = dtPersonal.Rows(0)("IDOperario")
        rp.Formulas("DescOperario").Text = dtPersonal.Rows(0)("DescOperario")
        rp.Formulas("DNI").Text = dtPersonal.Rows(0)("DNI")
        rp.Formulas("NAF").Text = dtPersonal.Rows(0)("N_Social")

        rp.Formulas("anio").Text = anio
        rp.Formulas("mes").Text = mesTiempo
        If (mes.ToString).Length = 1 Then
            rp.Formulas("Fecha Liquidacion").Text = "01/0" & mesTiempo + 1 & "/" & anio
        Else
            If mesTiempo = "12" Then
                rp.Formulas("Fecha Liquidacion").Text = "01/01/" & anio + 1
            Else
                rp.Formulas("Fecha Liquidacion").Text = "01/" & mes + 1 & "/" & anio
            End If
        End If

        Try

            ExpertisApp.OpenReport(rp)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Function

    Private Function generarCuadranteIndividualOF(ByVal mes As Integer, ByVal anio As Integer, ByVal informe As String)
        Dim rp As New Report(informe)
        'Dim dtObrasMes As New DataTable
        Dim dtTrabMes As New DataTable
        'Dim strSelect1 As String = "select distinct nobra,descobra from vSistLabListadoTrabajadoresObraMes where anio=" & anio & " and mes=" & mes
        Dim strSelect3 As String = ""
        'strSelect3 &= "select b.idOperario, b.descOperario, b.FechaInicio, a.Nobra from"
        'strSelect3 &= "(select idOperario, nobra, FechaInicio from vSistLabListadoTrabajadoresObraMes where anio=" & anio & " and mes=" & mes & ") as a inner join"
        'strSelect3 &= "(select distinct idOperario,descoperario,max(FechaInicio) as FechaInicio from vSistLabListadoTrabajadoresObraMes where anio='" & anio & "' and mes='" & mes & "' group by  idOperario,descoperario ) as b "
        'strSelect3 &= "on b.IDOperario=a.IDOperario and b.FechaInicio =a.FechaInicio"


        Dim mesTiempo As String
        Dim mesT As String
        'mes=mes+1
        If (mes.ToString).Length = 1 Then
            mesTiempo = "0" & mes.ToString
            mesT = anio & "-" & mesTiempo
        Else
            mesTiempo = mes.ToString
            mesT = anio & "-" & mesTiempo
        End If
        Dim FechaLiquidacion As Date = "01/01/2000"
        'David Velasco 03/10/22
        'Segunda condicion mes<=8 porque daba error para mes de oct(10)//nov(11)
        If (mes.ToString).Length = 1 And mes <= 8 Then
            FechaLiquidacion = "01/0" & mesTiempo + 1 & "/" & anio
        Else
            If mesTiempo = "12" Then
                FechaLiquidacion = "01/01/" & anio + 1
            Else
                FechaLiquidacion = "01/" & mes + 1 & "/" & anio
            End If
        End If
        Dim FechaIniLiq As Date = "01/" & mesTiempo & "/" & anio



        Dim DE As New BE.DataEngine

        strSelect3 = "select IDOperario,Obra_Predeterminada,FechaAlta,isnull(Fecha_Baja,DATEADD(year,10,'01/02/2022')) as Fecha_Baja from vMaestroOperarioCompleta where Obra_Predeterminada in (select idObra from tbObraCabecera where nobra in ('OFIZAM','OFIMAD','STOBRA')) and " & _
        "(Fecha_Baja is null or Fecha_Baja > '" & FechaIniLiq & "') and FechaAlta < '" & FechaLiquidacion & "'"

        'dtObrasMes = AdminData.GetData(strSelect1, False)
        dtTrabMes = DE.RetrieveData(strSelect3, , , , False)



        Dim a As Integer = 0
        Dim strSelect2 As String = ""
        Dim itrab As String
        Dim Otrab As String
        Try
            If dtTrabMes.Rows.Count > 0 Then
                For Each drTrabMes As DataRow In dtTrabMes.Rows
                    'Dim strSelectObra As String = "select"

                    If a > 0 Then
                        strSelect2 &= " union "
                    End If
                    itrab = drTrabMes(0)
                    Otrab = drTrabMes(1)
                    Dim strSelObra As String = "Select nobra from tbobracabecera where idObra='" & drTrabMes(1) & "'"
                    Dim DtObraTrab As DataTable = DE.RetrieveData(strSelObra, , , , False)
                    Dim drObraTrab As DataRow = DtObraTrab.Rows(0)

                    strSelect2 &= "SELECT (SELECT id FROM tbDatosEmpresa) AS IDEmpresa, b.fecha, month(b.fecha) AS mes, year(b.fecha) AS anio, b.Numdiasemana, b.diasemana, '" & drTrabMes(0) & _
                    "' AS IDOperario,CONVERT (Date,'" & drTrabMes("FechaAlta") & "') as FechaAlta, CONVERT (Date,'" & drTrabMes("Fecha_Baja") & "') as Fecha_Baja, isNull(a.TipoDia, 0) AS TipoDia, isNull(c.TipoDia, 0) AS TipoDiaV FROM ((SELECT * FROM tbCalendarioCentro WHERE  idcentro ='" & drTrabMes(1) & "'  AND month(fecha) = " & mes & " AND year(fecha) = " & anio & ") " & _
                    " AS A FULL OUTER JOIN (SELECT * FROM tiempo WHERE month(fecha) = " & mes & " AND year(fecha) = " & anio & ") AS B ON a.fecha = b.fecha) " & _
                    "full outer join (select * from tbCalendarioOperario where IdOperario='" & drTrabMes(0) & "' AND month(fecha) = " & mes & " AND year(fecha) = " & anio & ") as C on b.Fecha = c.fecha"


                    'If drObraTrab(0) = "STOBRA" Then
                    '    strSelect2 &= "select c.*, vOperarioHorario.IdHorario, (select  Nhorario from tbHorariosOficina where idhorario=vOperarioHorario.IdHorario) as NHorario,"
                    '    strSelect2 &= " IIf(c.tipodia=0, vOperarioHorario.EntradaMañana,null) as EntradaMañana,"
                    '    strSelect2 &= " IIf(c.tipodia=0, vOperarioHorario.SalidaMañana,null)  as SalidaMañana, "
                    '    strSelect2 &= " IIf(c.tipodia=0, vOperarioHorario.EntradaTarde,null)  as EntradaTarde,"
                    '    strSelect2 &= " IIf(c.tipodia=0, vOperarioHorario.SalidaTarde,null)  as SalidaTarde ,'" & drObraTrab(0) & "' as NObra from "
                    '    strSelect2 &= " (select (select id from tbDatosEmpresa ) as IDEmpresa,  b.fecha, month(b.fecha) as mes, year(b.fecha) as anio, b.Numdiasemana, b.diasemana,"
                    '    strSelect2 &= " isnull(a.IDOperario, '" & drTrabMes(0) & "') as IDOperario,isNull(a.TipoDia,0) as TipoDia from "
                    '    strSelect2 &= " (select * from tbCalendarioOperario where idoperario='" & drTrabMes(0) & "' and month(fecha) = " & mes & " and year(fecha) =" & anio & " ) as A"
                    '    strSelect2 &= " full outer join (select * from tiempo where month(fecha) =" & mes & " and year(fecha) =" & anio & ") as B on a.fecha=b.fecha) as C inner join vOperarioHorario "
                    '    strSelect2 &= " on c.IDOperario=vOperarioHorario.IdOperario and c.mes=vOperarioHorario.mes and c.anio=vOperarioHorario.Anio"
                    'Else
                    '    strSelect2 &= "select c.*, vOperarioHorario.IdHorario,(select  Nhorario from tbHorariosOficina where idhorario=vOperarioHorario.IdHorario) as NHorario,"
                    '    strSelect2 &= " IIf(c.tipodia=0,IIF(c.Numdiasemana<>5, vOperarioHorario.EntradaMañana,vOperarioHorario.EntradaViernes),null) as EntradaMañana,"
                    '    strSelect2 &= " IIf(c.tipodia=0,IIf(c.Numdiasemana<>5, vOperarioHorario.SalidaMañana,vOperarioHorario.SalidaViernes),null)  as SalidaMañana,"
                    '    strSelect2 &= " IIf(c.tipodia=0,IIf(c.Numdiasemana<>5, vOperarioHorario.EntradaTarde,null) ,null) as EntradaTarde,"
                    '    strSelect2 &= " IIf(c.tipodia=0,IIf(c.Numdiasemana<>5, vOperarioHorario.SalidaTarde,null),null)  as SalidaTarde ,'" & drObraTrab(0) & "' as NObra from "
                    '    strSelect2 &= " (select (select id from tbDatosEmpresa ) as IDEmpresa,  b.fecha, month(b.fecha) as mes, year(b.fecha) as anio, b.Numdiasemana, b.diasemana,"
                    '    strSelect2 &= " isnull(a.IDOperario, '" & drTrabMes(0) & "') as IDOperario,isNull(a.TipoDia,0) as TipoDia from "
                    '    strSelect2 &= " (select * from tbCalendarioOperario where idoperario='" & drTrabMes(0) & "' and month(fecha) = " & mes & " and year(fecha) =" & anio & " ) as A"
                    '    strSelect2 &= " full outer join (select * from tiempo where month(fecha) =" & mes & " and year(fecha) =" & anio & ") as B on a.fecha=b.fecha) as C inner join vOperarioHorario "
                    '    strSelect2 &= " on c.IDOperario=vOperarioHorario.IdOperario and c.mes=vOperarioHorario.mes and c.anio=vOperarioHorario.Anio"
                    'End If


                    'strSelect2 &= "select c.*, vOperarioHorario.IdHorario,"
                    'strSelect2 &= " IIf(c.tipodia=0,IIF(c.Numdiasemana<>5, vOperarioHorario.EntradaMañana,vOperarioHorario.EntradaViernes),null) as EntradaMañana,"
                    'strSelect2 &= " IIf(c.tipodia=0,IIf(c.Numdiasemana<>5, vOperarioHorario.SalidaMañana,vOperarioHorario.SalidaViernes),null)  as SalidaMañana,"
                    'strSelect2 &= " IIf(c.tipodia=0,IIf(c.Numdiasemana<>5, vOperarioHorario.EntradaTarde,null) ,null) as EntradaTarde,"
                    'strSelect2 &= " IIf(c.tipodia=0,IIf(c.Numdiasemana<>5, vOperarioHorario.SalidaTarde,null),null)  as SalidaTarde from "
                    'strSelect2 &= " (select (select id from tbDatosEmpresa ) as IDEmpresa,  b.fecha, month(b.fecha) as mes, year(b.fecha) as anio, b.Numdiasemana, b.diasemana,"
                    'strSelect2 &= " isnull(a.IDOperario, '" & drTrabMes(0) & "') as IDOperario,isNull(a.TipoDia,0) as TipoDia from "
                    'strSelect2 &= " (select * from tbCalendarioOperario where idoperario='" & drTrabMes(0) & "' and month(fecha) = " & mes & " and year(fecha) =" & anio & " ) as A"
                    'strSelect2 &= " full outer join (select * from tiempo where month(fecha) =" & mes & " and year(fecha) =" & anio & ") as B on a.fecha=b.fecha) as C inner join vOperarioHorario "
                    'strSelect2 &= " on c.IDOperario=vOperarioHorario.IdOperario and c.mes=vOperarioHorario.mes and c.anio=vOperarioHorario.Anio"


                    a = a + 1

                Next
                'strSelect2 &= " order by fecha,idOperario"

                'MsgBox(strSelect2)

                rp.DataSource = DE.RetrieveData(strSelect2, , , , False)
                'rp.Formulas("nobra").Text = Otrab
                rp.Formulas("anio").Text = anio
                rp.Formulas("mes").Text = mesTiempo
                'David Velasco 03/10/22
                'Segunda condicion mes<=8 porque daba error para mes de oct(10)//nov(11)
                If (mes.ToString).Length = 1 And mes <= 8 Then
                    rp.Formulas("Fecha Liquidacion").Text = "01/0" & mesTiempo + 1 & "/" & anio
                Else
                    If mesTiempo = "12" Then
                        rp.Formulas("Fecha Liquidacion").Text = "01/01/" & anio + 1
                    Else
                        rp.Formulas("Fecha Liquidacion").Text = "01/" & mes + 1 & "/" & anio
                    End If
                End If


                ExpertisApp.OpenReport(rp)
            Else
                'MsgBox("A fecha de liquidacion no hay ningun Trabajador en la obra " & NumObra)
                MsgBox("No se ha trabajado este mes")
            End If

        Catch ex As SqlClient.SqlException
            MsgBox("El error lo ha dado en el registro " & a & " y en el operario " & itrab)
            MsgBox(ex.Message)
        Catch ex As Exception
            MsgBox(ex.Message)


        End Try


    End Function

#End Region
End Class
