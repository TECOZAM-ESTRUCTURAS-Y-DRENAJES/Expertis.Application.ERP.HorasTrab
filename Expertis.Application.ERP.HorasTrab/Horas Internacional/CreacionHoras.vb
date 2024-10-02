Imports Solmicro.Expertis.Engine
Imports Solmicro.Expertis.Engine.UI
Imports Solmicro.Expertis.Business.ClasesTecozam
Imports System.Windows.Forms

Public Class CreacionHoras

    Dim aux As New MetodosAuxiliares

#Region "COMUNES"
    Public Sub CrearHorasApp(ByVal fecha1 As String, ByVal fecha2 As String)
        Dim dtLineasPersonas As DataTable
        Dim f As New Filter
        f.Add("FechaParte", FilterOperator.GreaterThanOrEqual, fecha1)
        f.Add("FechaParte", FilterOperator.LessThanOrEqual, fecha2)
        f.Add("Insertado", FilterOperator.Equal, False)
        dtLineasPersonas = New BE.DataEngine().Filter("tbHorasInternacional", f)

        If dtLineasPersonas.Rows.Count < 0 Then
            MsgBox("No existen registros para insertar en este rango de fechas.", MsgBoxStyle.Information)
            Exit Sub
        End If
        Dim result As DialogResult = MessageBox.Show("¿Existen registros para insertar.¿Deseas insertarlos en la tabla definitiva?", "Confirmación", MessageBoxButtons.YesNo)

        If result = DialogResult.Yes Then
            InsertarHoras(fecha1, fecha2, dtLineasPersonas)
        Else
            Exit Sub
        End If

    End Sub
    Public Sub InsertarHoras(ByVal fecha1 As String, ByVal fecha2 As String, ByVal dtLineasPersonas As DataTable)
        If ExpertisApp.DataBaseName = "xTecozam50R2" Then

        Else
            InsertaInternacional(fecha1, fecha2, dtLineasPersonas)
        End If
    End Sub
#End Region
    
#Region "INTERNACIONAL"

    Public Sub InsertaInternacional(ByVal fecha1 As String, ByVal fecha2 As String, ByVal dtLineasPersonas As DataTable)

        For Each dr As DataRow In dtLineasPersonas.Rows
            ' Obtener datos de la obra y el trabajo
            Dim IDObra As String = dr("IDObra").ToString()
            Dim IDTrabajo As String = ObtieneIDTrabajo(IDObra, "PT1")
            Dim categoria As String = ObtieneCategoriaIDOficio(dr("Oficio"))
            Dim rsTrabajo As DataTable = ObtenerDatosTrabajo(IDObra, IDTrabajo)

            ' Verificar si se obtuvo algún trabajo
            If rsTrabajo IsNot Nothing AndAlso rsTrabajo.Rows.Count > 0 Then
                ' Extraer información del trabajo
                Dim codTrabajo As String = rsTrabajo.Rows(0)("CodTrabajo").ToString()
                Dim DescTrabajo As String = rsTrabajo.Rows(0)("DescTrabajo").ToString()
                Dim IdTipoTrabajo As String = rsTrabajo.Rows(0)("IdTipoTrabajo").ToString()
                Dim IdSubTipoTrabajo As String = Nz(rsTrabajo.Rows(0)("IdSubtipotrabajo"), "")
                Dim DescParte As String = "HORAS APP " & Month(fecha1) & "-" & Year(fecha2)

                ' Insertar según el tipo de horas
                If IsNumeric(dr("HorasProductivas")) Then
                    InsertarHorasProductivas(IDTrabajo, IDObra, codTrabajo, DescTrabajo, IdTipoTrabajo, IdSubTipoTrabajo, dr, DescParte, categoria)
                ElseIf EsCausaValida(dr("IDCausa").ToString()) Then
                    InsertarHorasBaja(IDTrabajo, IDObra, codTrabajo, DescTrabajo, IdTipoTrabajo, IdSubTipoTrabajo, dr, DescParte, categoria)
                Else
                    InsertaRegistroSinHoras(IDTrabajo, IDObra, codTrabajo, DescTrabajo, IdTipoTrabajo, IdSubTipoTrabajo, dr, DescParte, categoria, dr("IDCausa").ToString())
                End If
                ActualizaLineaTablaIntermedia(dr("IDLineaParte"))
            End If
        Next

        MsgBox("Las horas han sido insertadas correctamente.", MsgBoxStyle.Information)
    End Sub

    ' Método para obtener datos del trabajo
    Private Function ObtenerDatosTrabajo(ByVal IDObra As String, ByVal IDTrabajo As String) As DataTable
        Dim filtro As New Filter()
        filtro.Add("IDObra", FilterOperator.Equal, IDObra)
        filtro.Add("IdTrabajo", FilterOperator.Equal, IDTrabajo)
        Return New BE.DataEngine().Filter("tbObraTrabajo", filtro)
    End Function

    ' Método para insertar horas productivas
    Private Sub InsertarHorasProductivas(ByVal IDTrabajo As String, ByVal IDObra As String, ByVal codTrabajo As String, ByVal DescTrabajo As String, ByVal IdTipoTrabajo As String, ByVal IdSubTipoTrabajo As String, ByVal dr As DataRow, ByVal DescParte As String, ByVal categoria As String)
        Dim txtSQL As String = "INSERT INTO tbObraMODControl(IdLineaModControl, IdTrabajo, IdObra, CodTrabajo, DescTrabajo, IdTipoTrabajo, " & _
                               "IdSubTipoTrabajo, IdOperario, IdCategoria, IdHora, FechaInicio, HorasRealMod, TasaRealModA, ImpRealModA, " & _
                               "HorasFactMod, ImpFactModA, DescParte, Facturable, FechaCreacionAudi, FechaModificacionAudi, Usuarioaudi, " & _
                               "IDOficio, IdTipoTurno, HorasAdministrativas, IDCategoriaProfesionalSCCP) " & _
                               "VALUES(" & aux.devuelveAutonumeri() & ", " & IDTrabajo & ", " & IDObra & ", '" & codTrabajo & "', '" & _
                               DescTrabajo & "', '" & IdTipoTrabajo & "', '" & IdSubTipoTrabajo & "', '" & dr("IDOperario") & _
                               "', 'PREDET', 'HO', '" & dr("FechaParte") & "', '" & dr("HorasProductivas") & "', 0, 0, 0, 0, '" & _
                               DescParte & "', 0, '" & Date.Now.Date & "', '" & Date.Now.Date & "', '" & ExpertisApp.UserName & _
                               "', '" & dr("Oficio") & "', 4, 0, " & categoria & ")"
        aux.Ejecutar(txtSQL)
    End Sub

    ' Método para insertar horas de baja (según causa)
    Private Sub InsertarHorasBaja(ByVal IDTrabajo As String, ByVal IDObra As String, ByVal codTrabajo As String, ByVal DescTrabajo As String, ByVal IdTipoTrabajo As String, ByVal IdSubTipoTrabajo As String, ByVal dr As DataRow, ByVal DescParte As String, ByVal categoria As String)
        Dim txtSQL As String = "INSERT INTO tbObraMODControl(IdLineaModControl, IdTrabajo, IdObra, CodTrabajo, DescTrabajo, IdTipoTrabajo, " & _
                               "IdSubTipoTrabajo, IdOperario, IdCategoria, IdHora, FechaInicio, HorasRealMod, TasaRealModA, ImpRealModA, " & _
                               "HorasFactMod, ImpFactModA, DescParte, Facturable, FechaCreacionAudi, FechaModificacionAudi, Usuarioaudi, " & _
                               "IDOficio, IdTipoTurno, HorasBaja, IDCategoriaProfesionalSCCP) " & _
                               "VALUES(" & aux.devuelveAutonumeri() & ", " & IDTrabajo & ", " & IDObra & ", '" & codTrabajo & "', '" & _
                               DescTrabajo & "', '" & IdTipoTrabajo & "', '" & IdSubTipoTrabajo & "', '" & dr("IDOperario") & _
                               "', 'PREDET', 'HB', '" & dr("FechaParte") & "', 0, 0, 0, 0, 0, '" & DescParte & "', 0, '" & _
                               Date.Now.Date & "', '" & Date.Now.Date & "', '" & ExpertisApp.UserName & "', '" & dr("Oficio") & _
                               "', 4, 8, " & categoria & ")"
        aux.Ejecutar(txtSQL)
    End Sub

    ' Método para insertar registros que no sean horas productivas ni bajas
    Private Sub InsertaRegistroSinHoras(ByVal IDTrabajo As String, ByVal IDObra As String, ByVal codTrabajo As String, ByVal DescTrabajo As String, ByVal IdTipoTrabajo As String, ByVal IdSubTipoTrabajo As String, ByVal dr As DataRow, ByVal DescParte As String, ByVal categoria As String, ByVal IDHora As String)
        Dim txtSQL As String = "INSERT INTO tbObraMODControl(IdLineaModControl, IdTrabajo, IdObra, CodTrabajo, DescTrabajo, IdTipoTrabajo, " & _
                               "IdSubTipoTrabajo, IdOperario, IdCategoria, IdHora, FechaInicio, HorasRealMod, TasaRealModA, ImpRealModA, " & _
                               "HorasFactMod, ImpFactModA, DescParte, Facturable, FechaCreacionAudi, FechaModificacionAudi, Usuarioaudi, " & _
                               "IDOficio, IdTipoTurno, HorasBaja, IDCategoriaProfesionalSCCP) " & _
                               "VALUES(" & aux.devuelveAutonumeri() & ", " & IDTrabajo & ", " & IDObra & ", '" & codTrabajo & "', '" & _
                               DescTrabajo & "', '" & IdTipoTrabajo & "', '" & IdSubTipoTrabajo & "', '" & dr("IDOperario") & _
                               "', 'PREDET', '" & IDHora & "', '" & dr("FechaParte") & "', 0, 0, 0, 0, 0, '" & DescParte & "', 0, '" & _
                               Date.Now.Date & "', '" & Date.Now.Date & "', '" & ExpertisApp.UserName & "', '" & dr("Oficio") & _
                               "', 4, 0, " & categoria & ")"
        aux.Ejecutar(txtSQL)
    End Sub

    ' Método para verificar si la causa es válida
    Private Function EsCausaValida(ByVal IDCausa As String) As Boolean
        Dim causasValidas As String() = {"ACC", "CC", "acc", "cc", "SSP", "B"}
        Return causasValidas.Contains(IDCausa)
    End Function

    Public Sub ActualizaLineaTablaIntermedia(ByVal IDLineaParte As String)
        Dim sql As String
        sql = "UPDATE tbHorasInternacional set Insertado= 1 where IDLineaParte= '" & IDLineaParte & "'"
        aux.EjecutarSql(sql)
    End Sub
#End Region

#Region "TECOZAM ESPAÑA"

#End Region

#Region "GETTERS"
    Public Function ObtieneIDTrabajo(ByVal IDObra As String, ByVal CodTrabajo As String) As String
        Dim dtTrabajo As New DataTable
        Dim filtro As New Filter
        filtro.Add("IDObra", FilterOperator.Equal, IDObra)
        filtro.Add("CodTrabajo", FilterOperator.Equal, CodTrabajo)
        dtTrabajo = New BE.DataEngine().Filter("tbObraTrabajo", filtro)
        Return dtTrabajo.Rows(0)("IDTrabajo")
    End Function

    Public Function ObtieneCategoriaIDOficio(ByVal IDOficio As String) As String
        Dim dtOperario As New DataTable
        Dim filtro As New Filter
        filtro.Add("IDOficio", FilterOperator.Equal, IDOficio)
        dtOperario = New BE.DataEngine().Filter("tbMaestroOficio", filtro)

        Return dtOperario.Rows(0)("Abreviatura").ToString
    End Function
#End Region
    
End Class
