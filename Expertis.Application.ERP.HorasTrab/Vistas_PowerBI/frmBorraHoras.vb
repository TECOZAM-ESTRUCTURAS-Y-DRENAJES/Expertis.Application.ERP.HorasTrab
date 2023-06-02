Imports Solmicro.Expertis.Engine
Imports Solmicro.Expertis.Business.ClasesTecozam
Imports Solmicro.Expertis.Engine.UI

Public Class frmBorraHoras

    Dim obraTrabajo As New Solmicro.Expertis.Business.Obra.ObraTrabajo
    Dim auto As New OperarioCalendario
    Dim aux As New Solmicro.Expertis.Business.ClasesTecozam.MetodosAuxiliares

    Private Sub bAceptar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bAceptar.Click
        Dim NObra As String = advObra.Text.ToString
        Dim IDOperario As String = advPersona.Text.ToString
        Dim FechaMa1 As String = Fecha1.Value.ToString
        Dim FechaMe2 As String = Fecha2.Value.ToString
        Dim TipoHoras As String = cbTipoHoras.Text

        If Len(NObra) = 0 Or Len(IDOperario) = 0 Or Len(FechaMa1) = 0 Or Len(FechaMe2) = 0 Or Len(TipoHoras) = 0 Then
            MsgBox("Los datos son obligatorios.", MsgBoxStyle.Critical, "Error")
            Exit Sub
        End If
        Dim IDHora As String = ""

        If TipoHoras = "Horas Administrativas" Then
            IDHora = "HA"
        End If
        If TipoHoras = "Horas Productivas" Then
            IDHora = "HO"
        End If

        BorrarRegistros(IDOperario, NObra, FechaMa1, FechaMe2, IDHora)

    End Sub

    Private Sub frmBorraHoras_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        cargaTipoHoras()
    End Sub
    Public Sub cargaTipoHoras()
        Dim dtcombo As New DataTable
        dtcombo.Columns.Add("TipoHoras")
        Dim dr As DataRow
        dr = dtcombo.NewRow()
        dr("TipoHoras") = "Horas Administrativas"
        dtcombo.Rows.Add(dr)
        dr = dtcombo.NewRow()
        dr("TipoHoras") = "Horas Productivas"
        dtcombo.Rows.Add(dr)

        cbTipoHoras.DataSource = dtcombo
    End Sub

    Private Sub bCancelar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bCancelar.Click
        Me.Close()
    End Sub
    Public Sub BorrarRegistros(ByVal IDOperario As String, ByVal NObra As String, ByVal Fecha1 As String, ByVal Fecha2 As String, ByVal IDHora As String)

        Dim IDObra As String
        IDObra = DevuelveIDObra(NObra)

        Dim f As New Filter
        f.Add("IDOperario", FilterOperator.Equal, IDOperario)
        If IDHora = "HA" Then
            f.Add("IDHora", FilterOperator.Equal, IDHora)
        Else
            f.Add("IDHora", FilterOperator.NotEqual, "HA")
        End If

        f.Add("IDObra", FilterOperator.Equal, IDObra)
        f.Add("FechaInicio", FilterOperator.GreaterThanOrEqual, Fecha1)
        f.Add("FechaInicio", FilterOperator.LessThanOrEqual, Fecha2)
        Dim dt As New DataTable
        dt = New BE.DataEngine().Filter("tbObraModControl", f)

        Dim filas As Integer
        filas = dt.Rows.Count

        If filas = 0 Then
            MsgBox("No existe ningun registro con estas especificaciones.")
        Else
            Dim intResponse As Integer
            intResponse = MsgBox("Existen " & filas & " registros de horas para esta persona en esta obra entre  estas fechas. ¿Desea eliminarlas?", vbYesNo + vbQuestion, "Información")

            'El finde se mete horas y el 31/12 tambien(En resumen)
            If intResponse = vbYes Then
                Borrar(IDOperario, IDObra, Fecha1, Fecha2, IDHora)
            Else
            End If
            Me.Close()
        End If
    End Sub

    Public Function DevuelveIDObra(ByVal NObra As String) As String
        Dim dtObra As New DataTable

        Dim filtro As New Filter
        filtro.Add("NObra", FilterOperator.Equal, NObra)
        dtObra = New BE.DataEngine().Filter("tbObraCabecera", filtro)

        Return dtObra.Rows(0)("IDObra")
    End Function
    Public Sub Borrar(ByVal IDOperario As String, ByVal IDObra As String, ByVal Fecha1 As String, ByVal Fecha2 As String, ByVal IDHora As String)
        Dim sql As String

        If IDHora = "HA" Then
            sql = "delete from tbObraMODControl "
            sql &= "where IDOperario='" & IDOperario & "' and IDObra='" & IDObra & "'"
            sql &= "and IDHora='" & IDHora & "'"
            sql &= "and FechaInicio>='" & Fecha1 & "' and FechaInicio<='" & Fecha2 & "'"
        Else
            sql = "delete from tbObraMODControl "
            sql &= "where IDOperario='" & IDOperario & "' and IDObra='" & IDObra & "'"
            sql &= "and IDHora!='HA'"
            sql &= "and FechaInicio>='" & Fecha1 & "' and FechaInicio<='" & Fecha2 & "'"
        End If

        

        aux.EjecutarSql(sql)
    End Sub
End Class