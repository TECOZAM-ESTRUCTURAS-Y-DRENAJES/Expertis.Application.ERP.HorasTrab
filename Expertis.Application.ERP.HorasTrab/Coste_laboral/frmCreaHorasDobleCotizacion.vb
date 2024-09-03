Imports Solmicro.Expertis.Engine
Imports Solmicro.Expertis.Business.ClasesTecozam
Imports Solmicro.Expertis.Engine.UI

Public Class frmCreaHorasDobleCotizacion
    Dim obraTrabajo As New Solmicro.Expertis.Business.Obra.ObraTrabajo
    Dim auto As New OperarioCalendario
    Dim aux As New Solmicro.Expertis.Business.ClasesTecozam.MetodosAuxiliares

    Private Sub bCancelar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bCancelar.Click
        Me.Close()
    End Sub

    Private Sub bAceptar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bAceptar.Click
        Dim NObra As String = advObra.Text.ToString
        Dim IDOperario As String = advPersona.Text.ToString
        Dim FechaMa1 As DateTime = Fecha1.Value.ToString


        If Len(NObra) = 0 Or Len(IDOperario) = 0 Or Len(FechaMa1) = 0 Or Len(txtHoras.Text) = 0 Then
            MsgBox("Los datos son obligatorios.", MsgBoxStyle.Critical, "Error")
            Exit Sub
        End If
        Dim IDHora As String = ""
        CreaHoras(IDOperario, NObra, FechaMa1, txtHoras.Text)
        MsgBox("Horas creadas correctamente.")
        Me.Close()
    End Sub
    Public Function DevuelveIDCategoriaProfesionalSCCP(ByVal IDOperario As String) As Integer
        Dim dt As New DataTable
        Dim f As New Filter

        f.Add("IDOperario", FilterOperator.Equal, IDOperario)
        dt = New BE.DataEngine().Filter("vOperarioCategoriaProf", f)
        If dt.Rows.Count > 0 Then
            Return dt(0)("Abreviatura")
        Else
            Return 0
        End If
    End Function
    Public Function DevuelveIDOficio(ByVal IDOperario As String) As String
        Dim dt As New DataTable
        Dim f As New Filter

        f.Add("IDOperario", FilterOperator.Equal, IDOperario)
        dt = New BE.DataEngine().Filter("tbMaestroOperario", f)
        If dt.Rows.Count > 0 Then
            Return dt(0)("IDOficio")
        Else
            Return ""
        End If
    End Function
    Public Function ObtieneIDTrabajo(ByVal IDObra As String, ByVal CodTrabajo As String) As String
        Dim dtTrabajo As New DataTable
        Dim filtro As New Filter
        filtro.Add("IDObra", FilterOperator.Equal, IDObra)
        filtro.Add("CodTrabajo", FilterOperator.Equal, CodTrabajo)

        dtTrabajo = New BE.DataEngine().Filter("tbObraTrabajo", filtro)
        'dtTrabajo = obraTrabajo.Filter(filtro)

        Return dtTrabajo.Rows(0)("IDTrabajo")

    End Function
    Public Function ObtieneIDObra(ByVal NObra As String) As String
        Dim dt As New DataTable
        Dim f As New Filter
        f.Add("NObra", FilterOperator.Equal, NObra)

        dt = New BE.DataEngine().Filter("tbObraCabecera", f)
        Return dt.Rows(0)("IDObra")
    End Function
    Public Function ObtieneCalendario(ByVal Fecha1 As String, ByVal Fecha2 As String) As DataTable
        Dim dtCalendario As New DataTable

        Dim filtro As New Filter
        filtro.Add("Fecha", FilterOperator.GreaterThanOrEqual, Fecha1)
        filtro.Add("Fecha", FilterOperator.LessThanOrEqual, Fecha2)
        dtCalendario = New BE.DataEngine().Filter("xTecozam50R2..tbMaestroFechas", filtro)

        Return dtCalendario
    End Function
    Public Function DevuelveIDObra(ByVal NObra As String) As String
        Dim dtObra As New DataTable

        Dim filtro As New Filter
        filtro.Add("NObra", FilterOperator.Equal, NObra)
        dtObra = New BE.DataEngine().Filter("tbObraCabecera", filtro)

        Return dtObra.Rows(0)("IDObra")
    End Function



    Public Sub CreaHoras(ByVal IDOperario As String, ByVal NObra As String, ByVal Fecha1 As String, ByVal Horas As String)
        Dim IDOficio As String
        Dim IDCategoriaProfesionalSCCP As String
        Dim IDObra As String
        Dim IDTrabajo As String
        Dim IDAutonumerico As String
        Dim CodTrabajo As String
        Dim txtSQL As String
        IDOficio = DevuelveIDOficio(IDOperario)
        IDCategoriaProfesionalSCCP = DevuelveIDCategoriaProfesionalSCCP(IDOperario)
        IDObra = DevuelveIDObra(NObra)
        IDTrabajo = ObtieneIDTrabajo(IDObra, "PT1")
        IDAutonumerico = auto.Autonumerico()

        Dim rsTrabajo As New DataTable : Dim filtro2 As New Filter
        filtro2.Add("IDObra", FilterOperator.Equal, IDObra) : filtro2.Add("IdTrabajo", FilterOperator.Equal, IDTrabajo)
        rsTrabajo = New BE.DataEngine().Filter("tbObraTrabajo", filtro2)
        'rsTrabajo = obraTrabajo.Filter(filtro2, , "IdTrabajo, CodTrabajo, DescTrabajo, IdTipoTrabajo, IdSubtipoTrabajo")

        IDTrabajo = rsTrabajo.Rows(0)("IdTrabajo") : CodTrabajo = rsTrabajo.Rows(0)("CodTrabajo")
        Dim DescTrabajo As String = "" : Dim IdTipoTrabajo As String = "" : Dim IdSubTipoTrabajo As String = ""
        DescTrabajo = rsTrabajo.Rows(0)("DescTrabajo") : IdTipoTrabajo = rsTrabajo.Rows(0)("IdTipoTrabajo") : IdSubTipoTrabajo = Nz(rsTrabajo.Rows(0)("IdSubtipotrabajo"), "")
        'Dim DescParte As String : DescParte = "OFICINA" & " " & Fecha1 & "-" & Fecha2 & "-OFI"

        Dim DescParte As String : DescParte = Fecha1 & "-DCotizacion"
        txtSQL = "Insert into tbObraMODControl(IdLineaModControl, IdTrabajo, IdObra, CodTrabajo, DescTrabajo, IdTipoTrabajo, " & _
            "IdSubTipoTrabajo, IdOperario, IdCategoria, IdHora, FechaInicio, HorasRealMod, TasaRealModA, " & _
             "ImpRealModA, HorasFactMod, ImpFactModA, DescParte, Facturable, FechaCreacionAudi, FechaModificacionAudi, Usuarioaudi, IDOficio, IdTipoTurno, HorasAdministrativas, IDCategoriaProfesionalSCCP) " & _
             "Values(" & IDAutonumerico & ", " & IDTrabajo & ", " & IDObra & ", '" & _
             CodTrabajo & "', '" & DescTrabajo & "', '" & IdTipoTrabajo & "', '" & _
             IdSubTipoTrabajo & "', '" & IDOperario & "', 'PREDET', '" & _
             "DC" & "', '" & Fecha1 & "', 0 , " & 0 & ", " & 0 & _
             ", 0 , " & 0 & _
             ", '" & DescParte & "', " & 0 & ", '" & Date.Now.Date & "', '" & Date.Now.Date & "', '" & ExpertisApp.UserName & "','" & IDOficio & "', 4, 0 ," & Nz(IDCategoriaProfesionalSCCP, "") & ")"
        auto.Ejecutar(txtSQL)
    End Sub

End Class