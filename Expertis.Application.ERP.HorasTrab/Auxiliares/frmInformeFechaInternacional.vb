Public Class frmInformeFechaInternacional

    Public blEstado As Boolean
    Public fecha1 As Date
    Public fecha2 As Date
    Public basedatos As String

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        fecha1 = FechaDesde.Value
        fecha2 = FechaHasta.Value
        basedatos = cbBasesDatos.Value.ToString()
        blEstado = True
        Me.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        blEstado = False
        Me.Close()
    End Sub

    Private Sub frmInformeFechaInternacional_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        cargaCombo()
    End Sub

    Public Sub cargaCombo()

        Dim dtcombo As New DataTable
        dtcombo.Columns.Add("Base de datos")
        dtcombo.Columns.Add("Pais")

        Dim dr As DataRow

        dr = dtcombo.NewRow()
        dr("Base de datos") = "xTecozamNorge50R2"
        dr("Pais") = "Noruega"
        dtcombo.Rows.Add(dr)

        dr = dtcombo.NewRow()
        dr("Base de datos") = "xDrenajesPortugal50R2"
        dr("Pais") = "Portugal"
        dtcombo.Rows.Add(dr)

        dr = dtcombo.NewRow()
        dr("Base de datos") = "xTecozamUnitedKingdom50R2"
        dr("Pais") = "Reino Unido"
        dtcombo.Rows.Add(dr)

        'dr = dtcombo.NewRow()
        'dr("Base de datos") = "3"
        'dr("Pais") = "Operario"
        'dtcombo.Rows.Add(dr)

        'dr = dtcombo.NewRow()
        'dr("Base de datos") = "4"
        'dr("Pais") = "Técnicos de Obra"
        'dtcombo.Rows.Add(dr)

        'dr = dtcombo.NewRow()
        'dr("Base de datos") = "5"
        'dr("Pais") = "STAFF"
        'dtcombo.Rows.Add(dr)

        cbBasesDatos.DataSource = dtcombo
        cbBasesDatos.ValueMember = "Base de datos"
        cbBasesDatos.DisplayMember = "Pais"
    End Sub
End Class