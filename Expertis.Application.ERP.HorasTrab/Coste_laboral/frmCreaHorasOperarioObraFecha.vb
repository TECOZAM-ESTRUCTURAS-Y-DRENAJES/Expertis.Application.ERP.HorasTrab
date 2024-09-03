Imports Solmicro.Expertis.Engine
Imports Solmicro.Expertis.Business.ClasesTecozam
Imports Solmicro.Expertis.Engine.UI

Public Class frmCreaHorasOperarioObraFecha

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
        Dim FechaMe2 As DateTime = Fecha2.Value.ToString
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
        If TipoHoras = "Horas Baja" Then
            IDHora = "HB"
        End If

        Dim intResponse As Integer
        intResponse = MsgBox("¿Quieres que meta horas también en fines de semana y festivos?", vbYesNo + vbQuestion, "Información")

        'El finde se mete horas y el 31/12 tambien(En resumen)
        If intResponse = vbYes Then
            CreaHorasConFestivos(IDOperario, NObra, FechaMa1, FechaMe2, IDHora)
        Else
            CreaHorasSinFestivos(IDOperario, NObra, FechaMa1, FechaMe2, IDHora)
        End If

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

    Public Function ObtieneDiasVacacionesYFestivos(ByVal basededatosteco As String, ByVal IDOperario As String, ByVal Fecha1 As String, ByVal Fecha2 As String) As DataTable
        Dim dtVacaciones As New DataTable
        Dim dtFestivos As New DataTable
        Dim dtTrabajados As New DataTable

        Dim filtro As New Filter
        'DIA DE VACACIONES = 2
        filtro.Add("TipoDia", FilterOperator.Equal, 2)
        filtro.Add("IDOperario", FilterOperator.Equal, IDOperario)
        filtro.Add("Fecha", FilterOperator.GreaterThanOrEqual, Fecha1)
        filtro.Add("Fecha", FilterOperator.LessThanOrEqual, Fecha2)

        dtVacaciones = New BE.DataEngine().Filter("tbCalendarioOperario", filtro, "Fecha, TipoDia")
        filtro.Clear()
        'FESTIVOS Y FINDES = 1
        filtro.Add("TipoDia", FilterOperator.Equal, 1)
        filtro.Add("IDCentro", FilterOperator.Equal, "00")
        filtro.Add("Fecha", FilterOperator.GreaterThanOrEqual, Fecha1)
        filtro.Add("Fecha", FilterOperator.LessThanOrEqual, Fecha2)
        dtFestivos = New BE.DataEngine().Filter(basededatosteco & "..tbCalendarioCentro", filtro, "Fecha, TipoDia")


        'FILTRO LOS DIAS TRABAJADOS
        filtro.Clear()
        filtro.Add("FechaInicio", FilterOperator.GreaterThanOrEqual, Fecha1)
        filtro.Add("FechaInicio", FilterOperator.LessThanOrEqual, Fecha2)
        filtro.Add("IDOperario", FilterOperator.Equal, IDOperario)
        dtTrabajados = New BE.DataEngine().Filter("tbObraModControl", filtro, "FechaInicio As Fecha")
        ' Crear un nuevo DataTable llamado dtCalendario
        Dim dtCalendario As New DataTable()

        ' Agregar las columnas Fecha y TipoDia al DataTable
        dtCalendario.Columns.Add("Fecha", GetType(Date))
        'dtCalendario.Columns.Add("TipoDia", GetType(Integer))

        ' Unir los DataTables dtVacaciones y dtFestivos en el DataTable dtCalendario
        dtCalendario.Merge(dtVacaciones)
        dtCalendario.Merge(dtFestivos)
        dtCalendario.Merge(dtTrabajados)

        Return dtCalendario
    End Function
    Public Function ObtieneDiasVacacionesYFestivos(ByVal IDOperario As String, ByVal Fecha1 As String, ByVal Fecha2 As String) As DataTable
        Dim dtVacaciones As New DataTable
        Dim dtFestivos As New DataTable
        Dim dtTrabajados As New DataTable

        Dim filtro As New Filter
        'DIA DE VACACIONES = 2
        filtro.Add("TipoDia", FilterOperator.Equal, 2)
        filtro.Add("IDOperario", FilterOperator.Equal, IDOperario)
        filtro.Add("Fecha", FilterOperator.GreaterThanOrEqual, Fecha1)
        filtro.Add("Fecha", FilterOperator.LessThanOrEqual, Fecha2)

        dtVacaciones = New BE.DataEngine().Filter("tbCalendarioOperario", filtro, "Fecha, TipoDia")
        filtro.Clear()
        'FESTIVOS Y FINDES = 1
        filtro.Add("TipoDia", FilterOperator.Equal, 1)
        filtro.Add("IDCentro", FilterOperator.Equal, "00")
        filtro.Add("Fecha", FilterOperator.GreaterThanOrEqual, Fecha1)
        filtro.Add("Fecha", FilterOperator.LessThanOrEqual, Fecha2)
        dtFestivos = New BE.DataEngine().Filter("xTecozam50R2..tbCalendarioCentro", filtro, "Fecha, TipoDia")


        'FILTRO LOS DIAS TRABAJADOS
        filtro.Clear()
        filtro.Add("FechaInicio", FilterOperator.GreaterThanOrEqual, Fecha1)
        filtro.Add("FechaInicio", FilterOperator.LessThanOrEqual, Fecha2)
        filtro.Add("IDOperario", FilterOperator.Equal, IDOperario)
        dtTrabajados = New BE.DataEngine().Filter("tbObraModControl", filtro, "FechaInicio As Fecha")
        ' Crear un nuevo DataTable llamado dtCalendario
        Dim dtCalendario As New DataTable()

        ' Agregar las columnas Fecha y TipoDia al DataTable
        dtCalendario.Columns.Add("Fecha", GetType(Date))
        'dtCalendario.Columns.Add("TipoDia", GetType(Integer))

        ' Unir los DataTables dtVacaciones y dtFestivos en el DataTable dtCalendario
        dtCalendario.Merge(dtVacaciones)
        dtCalendario.Merge(dtFestivos)
        dtCalendario.Merge(dtTrabajados)

        Return dtCalendario
    End Function

    Public Function ObtieneFechasInsertar(ByVal IDOperario As String, ByVal dtCalendario As DataTable, ByVal dtOperarioCalendarioNoProductivo As DataTable) As DataTable
        Dim dtDiasInsertar As New DataTable
        dtDiasInsertar.Columns.Add("Fecha", GetType(Date))

        'ESTE FOR FORMA LA TABLA CON FECHAS DONDE SE INSERTAN HORAS
        For Each rowCalendario As DataRow In dtCalendario.Rows
            Dim fechaCalendario As Date = rowCalendario.Field(Of Date)("Fecha")
            Dim encontrado As Boolean = False

            ' Busque si la fecha de la fila actual de dtCalendario está en dtOperarioCalendarioNoProductivo'
            For Each rowOperario As DataRow In dtOperarioCalendarioNoProductivo.Rows
                Dim fechaOperario As Date = rowOperario.Field(Of Date)("Fecha")
                If fechaCalendario = fechaOperario Then
                    encontrado = True
                    Exit For
                End If
            Next

            ' Si no se encontró la fecha en dtOperarioCalendarioNoProductivo, agregue una nueva fila a dtNuevaTabla'
            If Not encontrado Then
                Dim rowNuevaTabla As DataRow = dtDiasInsertar.NewRow()
                rowNuevaTabla("Fecha") = fechaCalendario
                dtDiasInsertar.Rows.Add(rowNuevaTabla)
            End If
        Next

        Dim dtOperario As New DataTable
        Dim f As New Filter
        f.Add("IDOperario", FilterOperator.Equal, IDOperario)
        dtOperario = New BE.DataEngine().Filter("tbMaestroOperarioSat", f)

        If Len(dtOperario.Rows(0)("Fecha_Baja").ToString) <> 0 Then
            Dim fechaBaja As String
            fechaBaja = dtOperario.Rows(0)("Fecha_Baja").ToString

            For i As Integer = dtDiasInsertar.Rows.Count - 1 To 0 Step -1
                Dim fila As DataRow = dtDiasInsertar.Rows(i)
                Dim fecha As Date = CDate(fila("Fecha"))

                If fecha >= fechaBaja Then
                    ' La fecha es mayor que la fecha límite, eliminamos la fila
                    dtDiasInsertar.Rows.RemoveAt(i)
                End If
            Next
            Return dtDiasInsertar
        Else
            Return dtDiasInsertar
        End If
    End Function


    Public Sub CreaHorasSinFestivos(ByVal IDOperario As String, ByVal NObra As String, ByVal Fecha1 As String, ByVal Fecha2 As String, ByVal IDHora As String)
        Dim IDOficio As String
        Dim IDCategoriaProfesionalSCCP As String
        Dim IDObra As String
        Dim IDTrabajo As String
        Dim IDAutonumerico As String
        Dim CodTrabajo As String
        Dim txtSQL As String

        Dim dtOperarioCalendarioNoProductivo As New DataTable
        Dim dtCalendario As New DataTable
        dtCalendario = ObtieneCalendario(Fecha1, Fecha2)

        'TABLA CON FECHAS DONDE SE INSERTAN HORAS
        Dim dtDiasInsertar As New DataTable
        dtDiasInsertar.Columns.Add("Fecha", GetType(Date))


        IDOficio = DevuelveIDOficio(IDOperario)
        IDCategoriaProfesionalSCCP = DevuelveIDCategoriaProfesionalSCCP(IDOperario)
        IDObra = DevuelveIDObra(NObra)
        IDTrabajo = ObtieneIDTrabajo(IDObra, "PT1")

        dtOperarioCalendarioNoProductivo = ObtieneDiasVacacionesYFestivos(IDOperario, Fecha1, Fecha2)
        dtDiasInsertar = ObtieneFechasInsertar(IDOperario, dtCalendario, dtOperarioCalendarioNoProductivo)


        For Each row As DataRow In dtDiasInsertar.Rows
            Dim fecha As Date = row.Field(Of Date)("Fecha")
            IDAutonumerico = auto.Autonumerico()

            Dim rsTrabajo As New DataTable : Dim filtro2 As New Filter
            filtro2.Add("IDObra", FilterOperator.Equal, IDObra) : filtro2.Add("IdTrabajo", FilterOperator.Equal, IDTrabajo)
            rsTrabajo = New BE.DataEngine().Filter("tbObraTrabajo", filtro2)
            'rsTrabajo = obraTrabajo.Filter(filtro2, , "IdTrabajo, CodTrabajo, DescTrabajo, IdTipoTrabajo, IdSubtipoTrabajo")

            IDTrabajo = rsTrabajo.Rows(0)("IdTrabajo") : CodTrabajo = rsTrabajo.Rows(0)("CodTrabajo")
            Dim DescTrabajo As String = "" : Dim IdTipoTrabajo As String = "" : Dim IdSubTipoTrabajo As String = ""
            DescTrabajo = rsTrabajo.Rows(0)("DescTrabajo") : IdTipoTrabajo = rsTrabajo.Rows(0)("IdTipoTrabajo") : IdSubTipoTrabajo = Nz(rsTrabajo.Rows(0)("IdSubtipotrabajo"), "")
            'Dim DescParte As String : DescParte = "OFICINA" & " " & Fecha1 & "-" & Fecha2 & "-OFI"

            If IDHora = "HA" Then
                Dim DescParte As String : DescParte = "OFICINA" & " " & Fecha1 & "-" & Fecha2 & "-OFI"
                txtSQL = "Insert into tbObraMODControl(IdLineaModControl, IdTrabajo, IdObra, CodTrabajo, DescTrabajo, IdTipoTrabajo, " & _
                    "IdSubTipoTrabajo, IdOperario, IdCategoria, IdHora, FechaInicio, HorasRealMod, TasaRealModA, " & _
                     "ImpRealModA, HorasFactMod, ImpFactModA, DescParte, Facturable, FechaCreacionAudi, FechaModificacionAudi, Usuarioaudi, IDOficio, IdTipoTurno, HorasAdministrativas, IDCategoriaProfesionalSCCP) " & _
                     "Values(" & IDAutonumerico & ", " & IDTrabajo & ", " & IDObra & ", '" & _
                     CodTrabajo & "', '" & DescTrabajo & "', '" & IdTipoTrabajo & "', '" & _
                     IdSubTipoTrabajo & "', '" & IDOperario & "', 'PREDET', '" & _
                     "HA" & "', '" & fecha & "', 0 , " & 0 & ", " & 0 & _
                     ", 0 , " & 0 & _
                     ", '" & DescParte & "', " & 0 & ", '" & Date.Now.Date & "', '" & Date.Now.Date & "', '" & ExpertisApp.UserName & "','" & IDOficio & "', 4, 8 ," & Nz(IDCategoriaProfesionalSCCP, "") & ")"
            ElseIf IDHora = "HB" Then
                Dim DescParte As String : DescParte = NObra & " " & Fecha1 & "-" & Fecha2 & "-HorasBaja"

                txtSQL = "Insert into tbObraMODControl(IdLineaModControl, IdTrabajo, IdObra, CodTrabajo, DescTrabajo, IdTipoTrabajo, " & _
                   "IdSubTipoTrabajo, IdOperario, IdCategoria, IdHora, FechaInicio, HorasRealMod, TasaRealModA, " & _
                    "ImpRealModA, HorasFactMod, ImpFactModA, DescParte, Facturable, FechaCreacionAudi, FechaModificacionAudi, Usuarioaudi, IDOficio, IdTipoTurno, HorasAdministrativas, IDCategoriaProfesionalSCCP, HorasBaja) " & _
                    "Values(" & IDAutonumerico & ", " & IDTrabajo & ", " & IDObra & ", '" & _
                    CodTrabajo & "', '" & DescTrabajo & "', '" & IdTipoTrabajo & "', '" & _
                    IdSubTipoTrabajo & "', '" & IDOperario & "', 'PREDET', '" & _
                    "HB" & "', '" & fecha & "', 0 , " & 0 & ", " & 0 & _
                    ", 0 , " & 0 & _
                    ", '" & DescParte & "', " & 0 & ", '" & Date.Now.Date & "', '" & Date.Now.Date & "', '" & ExpertisApp.UserName & "','" & IDOficio & "', 4, 0 ," & Nz(IDCategoriaProfesionalSCCP, "") & " ,8)"

            Else
                Dim DescParte As String : DescParte = NObra & " " & Fecha1 & "-" & Fecha2 & "-INDIVI"

                txtSQL = "Insert into tbObraMODControl(IdLineaModControl, IdTrabajo, IdObra, CodTrabajo, DescTrabajo, IdTipoTrabajo, " & _
                   "IdSubTipoTrabajo, IdOperario, IdCategoria, IdHora, FechaInicio, HorasRealMod, TasaRealModA, " & _
                    "ImpRealModA, HorasFactMod, ImpFactModA, DescParte, Facturable, FechaCreacionAudi, FechaModificacionAudi, Usuarioaudi, IDOficio, IdTipoTurno, HorasAdministrativas, IDCategoriaProfesionalSCCP) " & _
                    "Values(" & IDAutonumerico & ", " & IDTrabajo & ", " & IDObra & ", '" & _
                    CodTrabajo & "', '" & DescTrabajo & "', '" & IdTipoTrabajo & "', '" & _
                    IdSubTipoTrabajo & "', '" & IDOperario & "', 'PREDET', '" & _
                    "HO" & "', '" & fecha & "', 8 , " & 0 & ", " & 0 & _
                    ", 0 , " & 0 & _
                    ", '" & DescParte & "', " & 0 & ", '" & Date.Now.Date & "', '" & Date.Now.Date & "', '" & ExpertisApp.UserName & "','" & IDOficio & "', 4, 0 ," & Nz(IDCategoriaProfesionalSCCP, "") & ")"

            End If

            auto.Ejecutar(txtSQL)
        Next
    End Sub

    Public Sub CreaHorasConFestivos(ByVal IDOperario As String, ByVal NObra As String, ByVal Fecha1 As String, ByVal Fecha2 As String, ByVal IDHora As String)

        Dim IDOficio As String
        Dim IDCategoriaProfesionalSCCP As String
        Dim IDObra As String
        Dim IDTrabajo As String
        Dim IDAutonumerico As String
        Dim CodTrabajo As String
        Dim txtSQL As String

        Dim dtOperarioCalendarioNoProductivo As New DataTable
        Dim dtCalendario As New DataTable
        dtCalendario = ObtieneCalendario(Fecha1, Fecha2)

        'TABLA CON FECHAS DONDE SE INSERTAN HORAS
        Dim dtDiasInsertar As New DataTable
        dtDiasInsertar.Columns.Add("Fecha", GetType(Date))


        IDOficio = DevuelveIDOficio(IDOperario)
        IDCategoriaProfesionalSCCP = DevuelveIDCategoriaProfesionalSCCP(IDOperario)
        IDObra = DevuelveIDObra(NObra)
        IDTrabajo = ObtieneIDTrabajo(IDObra, "PT1")

        'dtOperarioCalendarioNoProductivo = ObtieneDiasVacacionesYFestivos("xTecozam50R2", IDOperario, Fecha1, Fecha2)
        dtDiasInsertar = ObtieneFechasInsertar(IDOperario, dtCalendario, dtOperarioCalendarioNoProductivo)


        For Each row As DataRow In dtDiasInsertar.Rows
            Dim fecha As Date = row.Field(Of Date)("Fecha")
            IDAutonumerico = auto.Autonumerico()

            Dim rsTrabajo As New DataTable : Dim filtro2 As New Filter
            filtro2.Add("IDObra", FilterOperator.Equal, IDObra) : filtro2.Add("IdTrabajo", FilterOperator.Equal, IDTrabajo)
            rsTrabajo = New BE.DataEngine().Filter("tbObraTrabajo", filtro2)
            'rsTrabajo = obraTrabajo.Filter(filtro2, , "IdTrabajo, CodTrabajo, DescTrabajo, IdTipoTrabajo, IdSubtipoTrabajo")

            IDTrabajo = rsTrabajo.Rows(0)("IdTrabajo") : CodTrabajo = rsTrabajo.Rows(0)("CodTrabajo")
            Dim DescTrabajo As String = "" : Dim IdTipoTrabajo As String = "" : Dim IdSubTipoTrabajo As String = ""
            DescTrabajo = rsTrabajo.Rows(0)("DescTrabajo") : IdTipoTrabajo = rsTrabajo.Rows(0)("IdTipoTrabajo") : IdSubTipoTrabajo = Nz(rsTrabajo.Rows(0)("IdSubtipotrabajo"), "")


            If IDHora = "HA" Then
                Dim DescParte As String : DescParte = NObra & " " & Fecha1 & "-" & Fecha2 & "-HorasAdministrativas"

                txtSQL = "Insert into tbObraMODControl(IdLineaModControl, IdTrabajo, IdObra, CodTrabajo, DescTrabajo, IdTipoTrabajo, " & _
                    "IdSubTipoTrabajo, IdOperario, IdCategoria, IdHora, FechaInicio, HorasRealMod, TasaRealModA, " & _
                     "ImpRealModA, HorasFactMod, ImpFactModA, DescParte, Facturable, FechaCreacionAudi, FechaModificacionAudi, Usuarioaudi, IDOficio, IdTipoTurno, HorasAdministrativas, IDCategoriaProfesionalSCCP) " & _
                     "Values(" & IDAutonumerico & ", " & IDTrabajo & ", " & IDObra & ", '" & _
                     CodTrabajo & "', '" & DescTrabajo & "', '" & IdTipoTrabajo & "', '" & _
                     IdSubTipoTrabajo & "', '" & IDOperario & "', 'PREDET', '" & _
                     "HA" & "', '" & fecha & "', 0 , " & 0 & ", " & 0 & _
                     ", 0 , " & 0 & _
                     ", '" & DescParte & "', " & 0 & ", '" & Date.Now.Date & "', '" & Date.Now.Date & "', '" & ExpertisApp.UserName & "','" & IDOficio & "', 4, 8 ," & Nz(IDCategoriaProfesionalSCCP, "") & ")"
            ElseIf IDHora = "HB" Then
                Dim DescParte As String : DescParte = NObra & " " & Fecha1 & "-" & Fecha2 & "-HorasBaja"

                txtSQL = "Insert into tbObraMODControl(IdLineaModControl, IdTrabajo, IdObra, CodTrabajo, DescTrabajo, IdTipoTrabajo, " & _
                   "IdSubTipoTrabajo, IdOperario, IdCategoria, IdHora, FechaInicio, HorasRealMod, TasaRealModA, " & _
                    "ImpRealModA, HorasFactMod, ImpFactModA, DescParte, Facturable, FechaCreacionAudi, FechaModificacionAudi, Usuarioaudi, IDOficio, IdTipoTurno, HorasAdministrativas, IDCategoriaProfesionalSCCP, HorasBaja) " & _
                    "Values(" & IDAutonumerico & ", " & IDTrabajo & ", " & IDObra & ", '" & _
                    CodTrabajo & "', '" & DescTrabajo & "', '" & IdTipoTrabajo & "', '" & _
                    IdSubTipoTrabajo & "', '" & IDOperario & "', 'PREDET', '" & _
                    "HB" & "', '" & fecha & "', 0 , " & 0 & ", " & 0 & _
                    ", 0 , " & 0 & _
                    ", '" & DescParte & "', " & 0 & ", '" & Date.Now.Date & "', '" & Date.Now.Date & "', '" & ExpertisApp.UserName & "','" & IDOficio & "', 4, 0 ," & Nz(IDCategoriaProfesionalSCCP, "") & " ,8)"

            Else
                Dim DescParte As String : DescParte = NObra & " " & Fecha1 & "-" & Fecha2 & "-INDIVI"

                txtSQL = "Insert into tbObraMODControl(IdLineaModControl, IdTrabajo, IdObra, CodTrabajo, DescTrabajo, IdTipoTrabajo, " & _
                   "IdSubTipoTrabajo, IdOperario, IdCategoria, IdHora, FechaInicio, HorasRealMod, TasaRealModA, " & _
                    "ImpRealModA, HorasFactMod, ImpFactModA, DescParte, Facturable, FechaCreacionAudi, FechaModificacionAudi, Usuarioaudi, IDOficio, IdTipoTurno, HorasAdministrativas, IDCategoriaProfesionalSCCP) " & _
                    "Values(" & IDAutonumerico & ", " & IDTrabajo & ", " & IDObra & ", '" & _
                    CodTrabajo & "', '" & DescTrabajo & "', '" & IdTipoTrabajo & "', '" & _
                    IdSubTipoTrabajo & "', '" & IDOperario & "', 'PREDET', '" & _
                    "HO" & "', '" & fecha & "', 8 , " & 0 & ", " & 0 & _
                    ", 0 , " & 0 & _
                    ", '" & DescParte & "', " & 0 & ", '" & Date.Now.Date & "', '" & Date.Now.Date & "', '" & ExpertisApp.UserName & "','" & IDOficio & "', 4, 0 ," & Nz(IDCategoriaProfesionalSCCP, "") & ")"

            End If
            
            auto.Ejecutar(txtSQL)
        Next
    End Sub


    Private Sub frmCreaHorasOperarioObraFecha_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
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
        dr = dtcombo.NewRow()
        dr("TipoHoras") = "Horas Baja"
        dtcombo.Rows.Add(dr)

        cbTipoHoras.DataSource = dtcombo
    End Sub
End Class