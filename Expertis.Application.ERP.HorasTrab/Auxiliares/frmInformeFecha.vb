Public Class frmInformeFecha

    Public blEstado As Boolean
    Public fecha1 As Date
    Public fecha2 As Date

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        fecha1 = FechaDesde.Value
        fecha2 = FechaHasta.Value
        blEstado = True
        Me.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        blEstado = False
        Me.Close()
    End Sub
End Class