Public Class frmFiltroOficio

    Friend codigo As String

    Public Sub cargarComboMes()
        Dim dtcombo As New DataTable
        dtcombo.Columns.Add("Codigo")
        dtcombo.Columns.Add("Oficio")

        Dim dr As DataRow

        dr = dtcombo.NewRow()
        dr("Codigo") = "1"
        dr("Oficio") = "Jefe de Producción"
        dtcombo.Rows.Add(dr)

        dr = dtcombo.NewRow()
        dr("Codigo") = "2"
        dr("Oficio") = "Encargado"
        dtcombo.Rows.Add(dr)

        dr = dtcombo.NewRow()
        dr("Codigo") = "3"
        dr("Oficio") = "Operario"
        dtcombo.Rows.Add(dr)

        dr = dtcombo.NewRow()
        dr("Codigo") = "4"
        dr("Oficio") = "Técnicos de Obra"
        dtcombo.Rows.Add(dr)

        dr = dtcombo.NewRow()
        dr("Codigo") = "5"
        dr("Oficio") = "STAFF"
        dtcombo.Rows.Add(dr)

        cbxMes.DataSource = dtcombo
        cbxMes.ValueMember = "Codigo"
        cbxMes.DisplayMember = "Oficio"

    End Sub

    Private Sub frmFiltroOficio_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        cargarComboMes()
    End Sub

    Private Sub btnAceptar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAceptar.Click
        codigo = cbxMes.Value.ToString
        Me.Close()
    End Sub
End Class