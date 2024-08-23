Imports Solmicro.Expertis.Engine

Public Class CIMntoBusqHorasAPP



    Private Sub CIMntoBusqHorasAPP_QueryExecuting(ByVal sender As Object, ByRef e As Solmicro.Expertis.Engine.UI.QueryExecutingEventArgs) Handles Me.QueryExecuting
        e.Filter.Add("NObra", FilterOperator.Equal, advNObra.Text)
        e.Filter.Add("IDOperario", FilterOperator.Equal, advIDOperario.Text)
        e.Filter.Add("Fecha", FilterOperator.GreaterThanOrEqual, clbFecha.Value)
        e.Filter.Add("Fecha", FilterOperator.LessThanOrEqual, clbFecha1.Value)
        e.Filter.Add("Encargado", FilterOperator.Equal, advEncargado.Text)
        e.Filter.Add("matrVehiculo", FilterOperator.Equal, txtmatrVehiculo.Text)
    End Sub
End Class
