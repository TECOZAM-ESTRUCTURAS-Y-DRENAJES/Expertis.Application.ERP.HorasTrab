Imports Solmicro.Expertis.Engine
Imports Solmicro.Expertis.Business.ClasesTecozam
Imports Solmicro.Expertis.Engine.UI
Imports System.Xml
Imports System.IO
Imports OfficeOpenXml

Public Class frmObraDias
    Public Sub New()
        MyBase.New()
        'El Diseñador de Windows Forms requiere esta llamada.
        InitializeComponent()
        'Agregar cualquier inicialización después de la llamada a InitializeComponent()
    End Sub

    Private Sub LoadToolbarActions()
        Try
            With Me.FormActions
                .Add("Calcular sumatorio por obra y mes.", AddressOf calcularSumatorio)
                .Add("GENERA EXCEL A TRAVÉS DE .TXT ", AddressOf generaExcel)
            End With
        Catch ex As Exception
            ExpertisApp.GenerateMessage(ex.Message)
        End Try
    End Sub
    Public Sub generaExcel()
        Try
            'Leo del XML
            Dim documentoxml As XmlDocument
            Dim nodelist As XmlNodeList
            Dim nodo As XmlNode
            documentoxml = New XmlDocument
            documentoxml.Load("N:\10. AUXILIARES\00. EXPERTIS\Turnos_MesProductivo.xml")
            nodelist = documentoxml.SelectNodes("/G/DatosObligatorios")
            Dim mesP As String = ""
            Dim anioP As String = ""

            For Each nodo In nodelist
                Dim IdImagen = nodo.Attributes.GetNamedItem("id").Value
                mesP = nodo.ChildNodes(0).InnerText
                anioP = nodo.ChildNodes(1).InnerText
            Next

            'Formo Fechas para sacar los turnos
            Dim Fecha1 As String
            Dim Fecha2 As String

            If mesP = "01" Then
                Fecha1 = "21/12/" & anioP - 1 & ""
                Fecha2 = "20/" & mesP & "/" & anioP
            Else
                Fecha1 = "21/" & mesP - 1 & "/" & anioP
                Fecha2 = "20/" & mesP & "/" & anioP
            End If

            'Formo la tabla para exportar la obra con los turnos
            Dim dt As New DataTable
            Dim dc As New DataColumn("NObra")
            dt.Columns.Add(dc)
            dc = New DataColumn("Turnos")
            dt.Columns.Add(dc)
            dc = New DataColumn("Dias")
            dt.Columns.Add(dc)
            dc = New DataColumn("Empresa")
            dt.Columns.Add(dc)

            'De momento solo lo hago para Tecozam
            Dim dtObras As New DataTable
            Dim dtTurnos As New DataTable
            Dim dtDias As New DataTable

            Dim sql As String = "select distinct(NObra) from xTecozam50R2..vSistLabListadoTrabajadoresObraMes where FechaInicio>='" & Fecha1 & "' And FechaInicio<='" & Fecha2 & "'"
            Dim s As New Solmicro.Expertis.Business.ClasesTecozam.ControlArticuloNSerie
            dtObras = s.EjecutarSqlSelect(sql)
            Dim sql2 As String = ""

            Dim CodigoObra As String = ""
            Dim Turnos As String = ""
            Dim Dias As String = ""
            Dim Maximo As Double = 0
            Dim Resta As Double = 0

            '***************TECOZAM******************
            For Each dr As DataRow In dtObras.Rows

                CodigoObra = dr("NObra")

                sql2 = "select sum(iif((idhora='HO' and horas >=8) or (idhora='HE' and horas>=4) , 1, 0)) as Turno from xTecozam50R2..vSistLabListadoTrabajadoresObraMes where FechaInicio>='" & Fecha1 & "' And FechaInicio<='" & Fecha2 & "' and NObra='" & CodigoObra & "'"
                dtTurnos = s.EjecutarSqlSelect(sql2)
                Turnos = dtTurnos.Rows(0)("Turno").ToString

                sql2 = "select top 1 IDOPerario, sum(iif((idhora='HO' and horas >=8) or (idhora='HE' and horas>=4), 1, 0)) as Dias from xTecozam50R2..vSistLabListadoTrabajadoresObraMes  where FechaInicio>='" & Fecha1 & "' And FechaInicio<='" & Fecha2 & "' and NObra='" & CodigoObra & "' group by IDOperario order by Dias desc"
                dtDias = s.EjecutarSqlSelect(sql2)
                Dias = dtDias.Rows(0)("Dias")

                Dim drFinal As DataRow
                drFinal = dt.NewRow
                drFinal("NObra") = CodigoObra
                drFinal("Turnos") = Turnos
                drFinal("Dias") = Dias
                drFinal("Empresa") = "T. ES."
                dt.Rows.Add(drFinal)
            Next
            Dim dtFinal As DataTable = dt
            '***************FERRALLAS******************

            sql = "select distinct(NObra) from xFerrallas50R2..vSistLabListadoTrabajadoresObraMes where FechaInicio>='" & Fecha1 & "' And FechaInicio<='" & Fecha2 & "'"
            dtObras = s.EjecutarSqlSelect(sql)

            For Each dr As DataRow In dtObras.Rows

                CodigoObra = dr("NObra")

                sql2 = "select sum(iif((idhora='HO' and horas >=8) or (idhora='HE' and horas>=4) , 1, 0)) as Turno from xFerrallas50R2..vSistLabListadoTrabajadoresObraMes where FechaInicio>='" & Fecha1 & "' And FechaInicio<='" & Fecha2 & "' and NObra='" & CodigoObra & "'"
                dtTurnos = s.EjecutarSqlSelect(sql2)
                Turnos = dtTurnos.Rows(0)("Turno").ToString

                sql2 = "select top 1 IDOPerario, sum(iif((idhora='HO' and horas >=8) or (idhora='HE' and horas>=4), 1, 0)) as Dias from xFerrallas50R2..vSistLabListadoTrabajadoresObraMes  where FechaInicio>='" & Fecha1 & "' And FechaInicio<='" & Fecha2 & "' and NObra='" & CodigoObra & "' group by IDOperario order by Dias desc"
                dtDias = s.EjecutarSqlSelect(sql2)
                Dias = dtDias.Rows(0)("Dias")

                Dim drFinal As DataRow
                drFinal = dtFinal.NewRow
                drFinal("NObra") = CodigoObra
                drFinal("Turnos") = Turnos
                drFinal("Dias") = Dias
                drFinal("Empresa") = "FERR. "
                dtFinal.Rows.Add(drFinal)
            Next

            '***************DCZ******************

            '***************TUK******************

            ' Borro las filas cuyo valor de Dias =0
            For i As Integer = dtFinal.Rows.Count - 1 To 0 Step -1
                ' Verificar si el valor de la columna "turnos" es cero.
                If dtFinal.Rows(i)("Dias") = 0 Then
                    ' Si el valor es cero, eliminar la fila.
                    dtFinal.Rows.RemoveAt(i)
                End If
            Next
            
            'Importar librería EPPLUS.dll
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial

            Dim ruta As New FileInfo("N:\10. AUXILIARES\00. EXPERTIS\" & mesP & " TURNOS " & anioP & ".xlsx")
            Dim rutaCadena As String = ""
            rutaCadena = ruta.FullName

            'Verificar si el archivo existe.
            If File.Exists(rutaCadena) Then
                'Si el archivo existe, eliminarlo.
                File.Delete(rutaCadena)
            End If

            Using package As New ExcelPackage(ruta)
                ' Crear una hoja de cálculo y obtener una referencia a ella.
                Dim worksheet = package.Workbook.Worksheets.Add(mesP & " TURNOS " & anioP)

                ' Copiar los datos de la DataTable a la hoja de cálculo.
                worksheet.Cells("A1").LoadFromDataTable(dtFinal, True)

                ' Guardar el archivo de Excel.
                package.Save()
            End Using

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
    Public Sub calcularSumatorio()
        Dim fechadesde As Date
        Dim fechahasta As Date
        Dim dt As New DataTable

        fechadesde = Nz(clbFecha.Value.ToString, "01/01/2000")
        fechahasta = Nz(clbFecha1.Value.ToString, "31/12/2050")
        dt = creardtSumatorio(fechadesde, fechahasta)

        rellenoTablaSumatorio(dt, fechadesde, fechahasta)

    End Sub

    Public Function creardtSumatorio(ByVal fechadesde As Date, ByVal fechahasta As Date)
        Dim dt As New DataTable
        Dim dc As New DataColumn("CodigoObra")
        dt.Columns.Add(dc)
        dc = New DataColumn("Dias")
        dt.Columns.Add(dc)
        dc = New DataColumn("Turnos")
        dt.Columns.Add(dc)
        Return dt
    End Function

    Public Sub calcular(ByVal fechadesde As Date, ByVal fechahasta As Date)
        creardt(fechadesde, fechahasta)
    End Sub


    Public Function estructuraTabla() As DataTable
        Dim dt As New DataTable
        Dim dc As New DataColumn("CodigoObra")
        dt.Columns.Add(dc)
        dc = New DataColumn("Dias")
        dt.Columns.Add(dc)
        Return dt
    End Function

    Public Sub creardt(ByVal fechadesde As Date, ByVal fechahasta As Date)
        'Creacion Estructura Tabla
        Dim dt As DataTable
        dt = estructuraTabla()
        'Relleno tabla
        rellenoTabla(dt, fechadesde, fechahasta)

    End Sub

    Public Sub rellenoTablaSumatorio(ByVal dt As DataTable, ByVal fechadesde As Date, ByVal fechahasta As Date)

        Try
            calculaObrasSumatorio(dt, fechadesde, fechahasta)

        Catch ex As Exception
            MsgBox("No existe ningún registro con estas características")
        End Try
    End Sub

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

    Private Sub frmObraDias_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadToolbarActions()
    End Sub

    Private Sub frmObraDias_QueryExecuting(ByVal sender As System.Object, ByRef e As Solmicro.Expertis.Engine.UI.QueryExecutingEventArgs) Handles MyBase.QueryExecuting
        Dim fechadesde As Date
        Dim fechahasta As Date
        fechadesde = Nz(clbFecha.Value.ToString, "01/01/2000")
        fechahasta = Nz(clbFecha1.Value.ToString, "31/12/2050")

        calcular(fechadesde, fechahasta)
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

            'sql2 = "select count(distinct(FechaInicio)) as Dias from vSistLabListadoTrabajadoresObraMes  where FechaInicio>='" & fechadesde & "' And FechaInicio<='" & fechahasta & "' and NObra='" & CodigoObra & "'"
            sql2 = "SELECT COUNT(DISTINCT CONVERT(DATE, FechaInicio)) AS Dias FROM vSistLabListadoTrabajadoresObraMes where ((idhora='HO' and horas >=1) or (idhora='HE' and horas>=1)) and FechaInicio >='" & fechadesde & "' And FechaInicio<='" & fechahasta & "' and NObra='" & CodigoObra & "'"
            dtHoras = s.EjecutarSqlSelect(sql2)
            Dias = dtHoras.Rows(0)("Dias").ToString

            Dim drFinal As DataRow
            drFinal = dt.NewRow
            drFinal("CodigoObra") = CodigoObra
            drFinal("Dias") = Dias
            dt.Rows.Add(drFinal)

        Next
        Grid.DataSource = dt
        dt = Nothing
    End Sub

    Public Sub calculaObrasSumatorio(ByVal dt As DataTable, ByVal fechadesde As Date, ByVal fechahasta As Date)

        Dim dtObras As New DataTable
        Dim dtTurnos As New DataTable
        Dim dtDias As New DataTable

        Dim sql As String = "select distinct(NObra) from vSistLabListadoTrabajadoresObraMes where FechaInicio>='" & fechadesde & "' And FechaInicio<='" & fechahasta & "'"
        Dim s As New Solmicro.Expertis.Business.ClasesTecozam.ControlArticuloNSerie
        dtObras = s.EjecutarSqlSelect(sql)
        Dim sql2 As String = ""

        Dim CodigoObra As String = ""
        Dim Turnos As String = ""
        Dim Dias As String = ""
        Dim Maximo As Double = 0
        Dim Resta As Double = 0

        For Each dr As DataRow In dtObras.Rows

            CodigoObra = dr("NObra")

            'sql2 = "select count(distinct(FechaInicio)) as Dias from vSistLabListadoTrabajadoresObraMes  where FechaInicio>='" & fechadesde & "' And FechaInicio<='" & fechahasta & "' and NObra='" & CodigoObra & "'"
            sql2 = "select sum(iif((idhora='HO' and horas >=8) or (idhora='HE' and horas>=4) , 1, 0)) as Turno from vSistLabListadoTrabajadoresObraMes where FechaInicio>='" & fechadesde & "' And FechaInicio<='" & fechahasta & "' and NObra='" & CodigoObra & "'"
            dtTurnos = s.EjecutarSqlSelect(sql2)
            Turnos = dtTurnos.Rows(0)("Turno").ToString

            'sql2 = "select top 1 IDOPerario, sum(iif((idhora='HO' and horas >=8) or (idhora='HE' and horas>=4), 1, 0)) as Dias from vSistLabListadoTrabajadoresObraMes  where FechaInicio>='" & fechadesde & "' And FechaInicio<='" & fechahasta & "' and NObra='" & CodigoObra & "' group by IDOperario order by Dias desc"
            sql2 = "SELECT COUNT(DISTINCT CONVERT(DATE, FechaInicio)) AS Dias FROM vSistLabListadoTrabajadoresObraMes where ((idhora='HO' and horas >=1) or (idhora='HE' and horas>=1)) and FechaInicio >='" & fechadesde & "' And FechaInicio<='" & fechahasta & "' and NObra='" & CodigoObra & "'"
            dtDias = s.EjecutarSqlSelect(sql2)
            Dias = dtDias.Rows(0)("Dias")

            Dim drFinal As DataRow
            drFinal = dt.NewRow
            drFinal("CodigoObra") = CodigoObra
            drFinal("Turnos") = Turnos
            drFinal("Dias") = Dias
            dt.Rows.Add(drFinal)

        Next
        Grid.DataSource = dt
        dt = Nothing

    End Sub
End Class
