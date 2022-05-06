Public Class frmFechas

    Friend blnImprimir As Boolean
    Friend mes As String
    Friend anio As Integer

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

        cbxMes.DataSource = dtcombo
        cbxMes.ValueMember = "Codigo"
        cbxMes.DisplayMember = "Descripcion"
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
        cbAnio.DataSource = dtcombo

    End Sub

    Private Sub btnAceptar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAceptar.Click
        blnImprimir = True
        mes = cbxMes.Value
        anio = cbAnio.Value

        If Trim(mes).Length = 0 And Trim(anio).Length = 0 Then
            MsgBox("Debe de seleccionar un mes y un año")
        Else
            'MsgBox("Ha elegido mes " & mes & " año " & anio)
            Me.Close()
        End If
    End Sub

    Private Sub btnCancelar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancelar.Click
        blnImprimir = False
        Me.Close()
    End Sub

    Public Sub New()

        ' Llamada necesaria para el Diseñador de Windows Forms.
        InitializeComponent()
        cargarComboMes()
        cargarComboAnio()

        ' Agregue cualquier inicialización después de la llamada a InitializeComponent().

    End Sub
End Class