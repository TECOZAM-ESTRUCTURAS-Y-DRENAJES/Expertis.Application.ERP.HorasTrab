Imports Solmicro.Expertis.Engine

Public Class CIMntoTrabObraMesUnionBasesDatos

    Private Sub CIMntoTrabObraMesUnionBasesDatos_QueryExecuting(ByVal sender As System.Object, ByRef e As Solmicro.Expertis.Engine.UI.QueryExecutingEventArgs) Handles MyBase.QueryExecuting
        e.Filter.Add("NObra", FilterOperator.Equal, advNObra.Text)
        e.Filter.Add("IDOperario", FilterOperator.Equal, advIDOperario.Text)
        e.Filter.Add("MesNatural", FilterOperator.Equal, cmbmes.Value)
        e.Filter.Add("AñoNatural", FilterOperator.Equal, cmbanio.Value)

        e.Filter.Add("FechaInicio", FilterOperator.GreaterThanOrEqual, clbFecha.Value)
        e.Filter.Add("FechaInicio", FilterOperator.LessThanOrEqual, clbFecha1.Value)

        'e.Filter.Add("MesProductivo", FilterOperator.Equal, cbMesProductivo.Value)
        'e.Filter.Add("AñoProductivo", FilterOperator.Equal, cbAnioProductivo.Value)

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

        cbMesProductivo.DataSource = dtcombo
        cbMesProductivo.ValueMember = "Codigo"
        cbMesProductivo.DisplayMember = "Descripcion"
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
        cbAnioProductivo.DataSource = dtcombo
    End Sub

    Public Sub New()

        ' Llamada necesaria para el Diseñador de Windows Forms.
        InitializeComponent()

        ' Agregue cualquier inicialización después de la llamada a InitializeComponent().
        cargarComboAnio()
        cargarComboMes()
        'LoadToolbarActions()

    End Sub
End Class