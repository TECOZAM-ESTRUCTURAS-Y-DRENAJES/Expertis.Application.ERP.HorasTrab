Imports System.Data
Imports System.Data.SqlClient
Imports Solmicro.Expertis.Engine
Imports Solmicro.Expertis.Engine.DAL

Public Class frmBorrarFicheroBDD

    Dim connectionString = "Data Source=stecodesarr;Initial Catalog=xTecozam50R2;User ID=sa;Password=180M296;"
    Dim baseDatos = "xTecozam50R2"

    Private Sub frmBorrarFicheroBDD_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LlenarComboTipo()
        LlenarComboMes()
        LlenarComboAnio()
    End Sub

    Private Sub LlenarComboTipo()
        Dim dtTipo As New DataTable
        dtTipo.Columns.Add("Opciones")

        Dim dr As DataRow

        dr = dtTipo.NewRow()
        dr("Opciones") = "A3"
        dtTipo.Rows.Add(dr)

        dr = dtTipo.NewRow()
        dr("Opciones") = "Horas"
        dtTipo.Rows.Add(dr)

        dr = dtTipo.NewRow()
        dr("Opciones") = "Extra"
        dtTipo.Rows.Add(dr)

        dr = dtTipo.NewRow()
        dr("Opciones") = "Regularizaciones"
        dtTipo.Rows.Add(dr)

        dr = dtTipo.NewRow()
        dr("Opciones") = "RegularizacionesIDGET"
        dtTipo.Rows.Add(dr)

        cbTipo.DataSource = dtTipo
        cbTipo.DisplayMember = "Opciones"
    End Sub

    Private Sub LlenarComboMes()
        Dim dtMes As New DataTable()
        dtMes.Columns.Add("Mes")

        Dim dr As DataRow

        For i As Integer = 1 To 12
            dr = dtMes.NewRow()
            dr("Mes") = i
            dtMes.Rows.Add(dr)
        Next
        cbMes.DataSource = dtMes
        cbMes.DisplayMember = "Mes"
        cbMes.ValueMember = "Mes"
    End Sub

    Private Sub LlenarComboAnio()
        Dim dtAnio As New DataTable
        dtAnio.Columns.Add("Año")

        Dim dr As DataRow

        For i As Integer = 0 To 5
            Dim j As Integer = Year(Today)
            dr = dtAnio.NewRow()
            dr("Año") = j - i
            dtAnio.Rows.Add(dr)
        Next
        cbAnio.DataSource = dtAnio
        cbAnio.DisplayMember = "Año"
        cbAnio.ValueMember = "Año"
    End Sub

    Private Sub btnAceptar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAceptar.Click
        Borrar()
    End Sub

    Private Sub Borrar()
        'obtener datos
        Dim tipo As String = cbTipo.Value
        Dim mes As Integer = cbMes.Value
        Dim anio As Integer = cbAnio.Value

        'buscar coincidencias en BDD
        Dim dtCoincidenciasBDD As New DataTable
        Dim f As New Filter()

        Using connection As New SqlConnection(connectionString)
            connection.Open()
            If tipo = "A3" Then
                f.Add("Mes", FilterOperator.Equal, mes)
                f.Add("Anio", FilterOperator.Equal, anio)
                dtCoincidenciasBDD = New BE.DataEngine().Filter(baseDatos & "..tbA3CostesLaborales", f, "*")

                If dtCoincidenciasBDD.Rows.Count > 0 Then
                    Dim deteleCommand As String = "DELETE FROM " & baseDatos & "..tbA3CostesLaborales WHERE Mes='" & mes & "' AND Anio='" & anio & "'"
                    Using command As New SqlCommand(deteleCommand, connection)
                        command.ExecuteNonQuery()
                    End Using
                    MsgBox("Fichero borrado correctamente.", MsgBoxStyle.Information, "Exito")
                Else
                    MsgBox("No hay coincidencias en Base de Datos.", MsgBoxStyle.Information, "Sin coincidencias")
                End If
            ElseIf tipo = "Horas" Then
                f.Add("MesNatural", FilterOperator.Equal, mes)
                f.Add("AñoNatural", FilterOperator.Equal, anio)
                dtCoincidenciasBDD = New BE.DataEngine().Filter(baseDatos & "..tbHorasCostesLaborales", f, "*")

                If dtCoincidenciasBDD.Rows.Count > 0 Then
                    Dim deteleCommand As String = "DELETE FROM " & baseDatos & "..tbHorasCostesLaborales WHERE MesNatural='" & mes & "' AND AñoNatural='" & anio & "'"
                    Using command As New SqlCommand(deteleCommand, connection)
                        command.ExecuteNonQuery()
                    End Using
                    MsgBox("Fichero borrado correctamente.", MsgBoxStyle.Information, "Exito")
                Else
                    MsgBox("No hay coincidencias en Base de Datos.", MsgBoxStyle.Information, "Sin coincidencias")
                End If
            ElseIf tipo = "Extra" Then
                f.Add("Mes", FilterOperator.Equal, mes)
                f.Add("Anio", FilterOperator.Equal, anio)
                dtCoincidenciasBDD = New BE.DataEngine().Filter(baseDatos & "..tbExtraCostesLaborales", f, "*")

                If dtCoincidenciasBDD.Rows.Count > 0 Then
                    Dim deteleCommand As String = "DELETE FROM " & baseDatos & "..tbExtraCostesLaborales WHERE Mes='" & mes & "' AND Anio='" & anio & "'"
                    Using command As New SqlCommand(deteleCommand, connection)
                        command.ExecuteNonQuery()
                    End Using
                    MsgBox("Fichero borrado correctamente.", MsgBoxStyle.Information, "Exito")
                Else
                    MsgBox("No hay coincidencias en Base de Datos.", MsgBoxStyle.Information, "Sin coincidencias")
                End If
            ElseIf tipo = "Regularizaciones" Then
                f.Add("Mes", FilterOperator.Equal, mes)
                f.Add("Anio", FilterOperator.Equal, anio)
                dtCoincidenciasBDD = New BE.DataEngine().Filter(baseDatos & "..tbRegularizacionesCostesLaborales", f, "*")

                If dtCoincidenciasBDD.Rows.Count > 0 Then
                    Dim deteleCommand As String = "DELETE FROM " & baseDatos & "..tbRegularizacionesCostesLaborales WHERE Mes='" & mes & "' AND Anio='" & anio & "'"
                    Using command As New SqlCommand(deteleCommand, connection)
                        command.ExecuteNonQuery()
                    End Using
                    MsgBox("Fichero borrado correctamente.", MsgBoxStyle.Information, "Exito")
                Else
                    MsgBox("No hay coincidencias en Base de Datos.", MsgBoxStyle.Information, "Sin coincidencias")
                End If

            ElseIf tipo = "RegularizacionesIDGET" Then
                f.Add("Mes", FilterOperator.Equal, mes)
                f.Add("Anio", FilterOperator.Equal, anio)
                dtCoincidenciasBDD = New BE.DataEngine().Filter(baseDatos & "..tbRegularizacionesIDGETCostesLaborales", f, "*")

                If dtCoincidenciasBDD.Rows.Count > 0 Then
                    Dim deteleCommand As String = "DELETE FROM " & baseDatos & "..tbRegularizacionesIDGETCostesLaborales WHERE Mes='" & mes & "' AND Anio='" & anio & "'"
                    Using command As New SqlCommand(deteleCommand, connection)
                        command.ExecuteNonQuery()
                    End Using
                    MsgBox("Fichero borrado correctamente.", MsgBoxStyle.Information, "Exito")
                Else
                    MsgBox("No hay coincidencias en Base de Datos.", MsgBoxStyle.Information, "Sin coincidencias")
                End If
            End If
        End Using
    End Sub

    Private Sub btnCancelar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancelar.Click
        Me.Close()
    End Sub
End Class