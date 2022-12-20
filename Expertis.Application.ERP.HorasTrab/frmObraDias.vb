Imports Solmicro.Expertis.Engine
Imports Solmicro.Expertis.Business.ClasesTecozam
Imports Solmicro.Expertis.Engine.UI

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
            End With
        Catch ex As Exception
            ExpertisApp.GenerateMessage(ex.Message)
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

            sql2 = "select count(distinct(FechaInicio)) as Dias from vSistLabListadoTrabajadoresObraMes  where FechaInicio>='" & fechadesde & "' And FechaInicio<='" & fechahasta & "' and NObra='" & CodigoObra & "'"
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
            sql2 = "select sum(iif(idhora='HO' or idhora='HE' , 1, 0)) as Turno from vSistLabListadoTrabajadoresObraMes where FechaInicio>='" & fechadesde & "' And FechaInicio<='" & fechahasta & "' and NObra='" & CodigoObra & "'"
            dtTurnos = s.EjecutarSqlSelect(sql2)
            Turnos = dtTurnos.Rows(0)("Turno").ToString

            sql2 = "select top 1 IDOPerario, sum(iif(idhora='HO' or idhora='HE', 1, 0)) as Dias from vSistLabListadoTrabajadoresObraMes  where FechaInicio>='" & fechadesde & "' And FechaInicio<='" & fechahasta & "' and NObra='" & CodigoObra & "' group by IDOperario order by Dias desc"
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
