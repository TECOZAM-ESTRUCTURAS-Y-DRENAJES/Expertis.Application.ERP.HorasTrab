Public Class frmDatosInformePersonal

    Public blEstado As Boolean
    Public mes As Integer
    Public anio As Integer
    Public IDOperario As String



    Private Sub btnAceptar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAceptar.Click
        If cbxMes.Value < 1 And cbxMes.Value > 12 Then
            MsgBox("Debe introducir un valor entre 1 y 12")
        Else
            If cbAnio.Value < 2006 Then
                MsgBox("No existen valores anteriores a 2006")
            Else

                mes = cbxMes.Value
                anio = cbAnio.Value
                IDOperario = advIDOperario.Text
                blEstado = False
                Me.Close()

            End If
        End If

    End Sub

    Private Sub btbCancelar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancelar.Click
        blEstado = True
        Me.Close()
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

    Public Sub New()

        ' Llamada necesaria para el Diseñador de Windows Forms.
        InitializeComponent()


        ' Agregue cualquier inicialización después de la llamada a InitializeComponent().
        cargarComboMes()
        cargarComboAnio()
    End Sub
End Class