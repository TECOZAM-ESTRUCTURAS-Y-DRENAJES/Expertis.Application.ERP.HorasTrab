Imports System.Windows.Forms
Imports Solmicro.Expertis.Engine
Imports Solmicro.Expertis.Engine.UI
Imports Solmicro.Expertis.Business.ClasesTecozam
Imports Solmicro.Expertis.Business.Obra
Imports System.Math
Imports System.Data.SqlClient
Imports Solmicro.Expertis.Engine.DAL
Imports Solmicro.Expertis.Business
Imports iTextSharp.text.pdf
Imports iTextSharp.text.pdf.parser
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.IO
Imports OfficeOpenXml
Imports System.Data
Imports System.Globalization




Public Class CargaHorasJPSTAFF
    'CONSTANTES
    Const DB_TECOZAM As String = "xTecozam50R2"
    Const DB_FERRALLAS As String = "xFerrallas50R2"
    Const DB_SECOZAM As String = "xSecozam50R2"
    Const DB_DCZ As String = "xDrenajesPortugal50R2"
    Const DB_UK As String = "xTecozamUnitedKingdom50R2"
    Const DB_NO As String = "xTecozamNorge50R2"
    Const DB_SU As String = "xTecozamSuecia4"

    Dim obraTrabajo As New ObraTrabajo
    Dim auto As New OperarioCalendario
    Dim aux As New Solmicro.Expertis.Business.ClasesTecozam.MetodosAuxiliares

    Private Sub bBorrarExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bBorrarExcel.Click
        Dim DescParte As String

        DescParte = InputBox("Introduzca NObra mesNatural-añoNatural-JP" & vbCrLf & "Por ejemplo:T636 04-2023-JP" & _
                             vbCrLf & "Por ejemplo:OFICINA 04-2023-OFI", "Borrar horas administrativas")
        If DescParte = "" Then
            MsgBox("Faltan datos.")
        Else
            'Comentado por David Velasco 15/05/23

            Dim auto As New OperarioCalendario
            auto.BorraDatosObraMODControlHorasAdministrativas(DescParte, DB_TECOZAM)
            auto.BorraDatosObraMODControlHorasAdministrativas(DescParte, DB_FERRALLAS)
            auto.BorraDatosObraMODControlHorasAdministrativas(DescParte, DB_SECOZAM)
            'auto.BorraDatosObraMODControlHorasAdministrativas(DescParte, DB_DCZ)
            'auto.BorraDatosObraMODControlHorasAdministrativas(DescParte, DB_UK)
        End If
    End Sub

    Private Sub btnSalir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub

    Private Sub cmdUbicacion_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUbicacion.Click
        Dim CD As New OpenFileDialog()

        CD.Title = "Seleccionar archivos"
        CD.Filter = "Archivos Excel(*.xls;*.xlsx;*.xlsm)|*.xls;*xlsx;*.xlsm|Todos los archivos(*.*)|*.*"

        'CD.ShowOpen()
        CD.ShowDialog()

        If CD.FileName <> "" Then
            'lblRuta.Caption = CD.FileName
            lblRuta.Text = CD.FileName
        End If
    End Sub
    '12/05/2022
    Public Function CargaTablas(ByRef dtTecozam As DataTable, ByRef dtPortugal As DataTable, ByRef dtUK As DataTable, ByRef dtNO As DataTable, ByVal dt As DataTable) As Integer
        dtTecozam.Columns.Add("IDOperario", GetType(String))
        dtTecozam.Columns.Add("DescOperario", GetType(String))
        dtTecozam.Columns.Add("Empresa", GetType(String))
        dtTecozam.Columns.Add("CentroCoste", GetType(String))
        dtTecozam.Columns.Add("ProduccionSinVentas", GetType(String))
        dtTecozam.Columns.Add("Porcentaje", GetType(Double))

        dtPortugal.Columns.Add("IDOperario", GetType(String))
        dtPortugal.Columns.Add("DescOperario", GetType(String))
        dtPortugal.Columns.Add("Empresa", GetType(String))
        dtPortugal.Columns.Add("CentroCoste", GetType(String))
        dtPortugal.Columns.Add("ProduccionSinVentas", GetType(String))
        dtPortugal.Columns.Add("Porcentaje", GetType(Double))

        dtUK.Columns.Add("IDOperario", GetType(String))
        dtUK.Columns.Add("DescOperario", GetType(String))
        dtUK.Columns.Add("Empresa", GetType(String))
        dtUK.Columns.Add("CentroCoste", GetType(String))
        dtUK.Columns.Add("ProduccionSinVentas", GetType(String))
        dtUK.Columns.Add("Porcentaje", GetType(Double))

        dtNO.Columns.Add("IDOperario", GetType(String))
        dtNO.Columns.Add("DescOperario", GetType(String))
        dtNO.Columns.Add("Empresa", GetType(String))
        dtNO.Columns.Add("CentroCoste", GetType(String))
        dtNO.Columns.Add("ProduccionSinVentas", GetType(String))
        dtNO.Columns.Add("Porcentaje", GetType(Double))

        For Each dr As DataRow In dt.Rows
            If dr("Empresa") = "T. ES." Then
                dtTecozam.ImportRow(dr)
            ElseIf dr("Empresa") = "D. P." Then
                dtPortugal.ImportRow(dr)
            ElseIf dr("Empresa") = "T. UK." Then
                dtUK.ImportRow(dr)
            ElseIf dr("Empresa") = "T. NO." Then
                dtNO.ImportRow(dr)
            Else
                Return 0
            End If
        Next
        Return 1
    End Function

    Public Sub insertaHorasJPStaffTecozam(ByVal mes As String, ByVal año As String, ByVal Fecha1 As String, ByVal Fecha2 As String, ByVal dtTecozam As DataTable)
        Dim IDOperario As String
        Dim IDOficio As String
        Dim IDCategoriaProfesionalSCCP As String
        Dim IDObra As String
        Dim IDTrabajo As String
        Dim IDAutonumerico As String
        Dim CodTrabajo As String
        Dim txtSQL As String
        Dim horas As Double = 0

        'Tabla que recoje los dias que no se trabaja, ya sea por vacacion o por festivo/fin de semana
        Dim dtOperarioCalendarioNoProductivo As New DataTable
        Dim dtCalendario As New DataTable
        dtCalendario = ObtieneCalendario(Fecha1, Fecha2)

        'TABLA CON FECHAS DONDE SE INSERTAN HORAS
        Dim dtDiasInsertar As New DataTable
        dtDiasInsertar.Columns.Add("Fecha", GetType(Date))

        Dim filas As Integer = 0
        PvProgreso.Value = 0 : PvProgreso.Maximum = dtTecozam.Rows.Count : PvProgreso.Step = 1 : PvProgreso.Visible = True

        For Each fila As DataRow In dtTecozam.Rows
            IDOperario = fila("IDOperario")
            IDOficio = DevuelveIDOficio(DB_TECOZAM, IDOperario)
            IDCategoriaProfesionalSCCP = DevuelveIDCategoriaProfesionalSCCP(DB_TECOZAM, IDOperario)
            Dim filtro As New Filter
            Dim dtObra As New DataTable
            filtro.Add("NObra", FilterOperator.Equal, fila("CentroCoste"))
            dtObra = New BE.DataEngine().Filter(DB_TECOZAM & "..tbObraCabecera", filtro)
            IDObra = dtObra.Rows(0)("IDObra").ToString
            IDTrabajo = ObtieneIDTrabajo(DB_TECOZAM, IDObra, "PT1")
            horas = 8 * fila("Porcentaje")

            dtOperarioCalendarioNoProductivo = ObtieneDiasVacacionesYFestivosJP(DB_TECOZAM, DB_TECOZAM, IDOperario, Fecha1, Fecha2)
            dtDiasInsertar = ObtieneFechasInsertar(DB_TECOZAM, IDOperario, dtCalendario, dtOperarioCalendarioNoProductivo)

            Windows.Forms.Application.DoEvents()
            LProgreso.Text = "Importando : " & IDOperario & " - TECOZAM JP"
            Windows.Forms.Application.DoEvents()

            For Each row As DataRow In dtDiasInsertar.Rows
                Dim fecha As Date = row.Field(Of Date)("Fecha")
                IDAutonumerico = auto.Autonumerico()

                Dim rsTrabajo As New DataTable : Dim filtro2 As New Filter
                filtro2.Add("IDObra", FilterOperator.Equal, IDObra) : filtro2.Add("IdTrabajo", FilterOperator.Equal, IDTrabajo)
                rsTrabajo = New BE.DataEngine().Filter(DB_TECOZAM & "..tbObraTrabajo", filtro2)
                'rsTrabajo = obraTrabajo.Filter(filtro2, , "IdTrabajo, CodTrabajo, DescTrabajo, IdTipoTrabajo, IdSubtipoTrabajo")

                IDTrabajo = rsTrabajo.Rows(0)("IdTrabajo") : CodTrabajo = rsTrabajo.Rows(0)("CodTrabajo")
                Dim DescTrabajo As String = "" : Dim IdTipoTrabajo As String = "" : Dim IdSubTipoTrabajo As String = ""
                DescTrabajo = rsTrabajo.Rows(0)("DescTrabajo") : IdTipoTrabajo = rsTrabajo.Rows(0)("IdTipoTrabajo") : IdSubTipoTrabajo = Nz(rsTrabajo.Rows(0)("IdSubtipotrabajo"), "")
                Dim DescParte As String : DescParte = "JP STAFF " & mes & "-" & año & "-JP"

                txtSQL = "Insert into " & DB_TECOZAM & "..tbObraMODControl(IdLineaModControl, IdTrabajo, IdObra, CodTrabajo, DescTrabajo, IdTipoTrabajo, " & _
                        "IdSubTipoTrabajo, IdOperario, IdCategoria, IdHora, FechaInicio, HorasRealMod, TasaRealModA, " & _
                         "ImpRealModA, HorasFactMod, ImpFactModA, DescParte, Facturable, FechaCreacionAudi, FechaModificacionAudi, Usuarioaudi, IDOficio, IdTipoTurno, HorasAdministrativas, IDCategoriaProfesionalSCCP) " & _
                         "Values(" & IDAutonumerico & ", " & IDTrabajo & ", " & IDObra & ", '" & _
                         CodTrabajo & "', '" & DescTrabajo & "', '" & IdTipoTrabajo & "', '" & _
                         IdSubTipoTrabajo & "', '" & IDOperario & "', 'PREDET', '" & _
                         "HA" & "', '" & fecha & "', 0 , " & 0 & ", " & 0 & _
                         ", 0 , " & 0 & _
                         ", '" & DescParte & "', " & 0 & ", '" & Date.Now.Date & "', '" & Date.Now.Date & "', '" & ExpertisApp.UserName & "','" & IDOficio & "', 4, '" & Replace(horas, ",", ".") & " ' ," & Nz(IDCategoriaProfesionalSCCP, "") & ")"

                auto.Ejecutar(txtSQL)
            Next

            filas = filas + 1
            PvProgreso.Value = filas
        Next
        '3. Obtengo una tabla por persona de los días que no tengan que insertar horas
        'MsgBox("Horas de la gente de oficina de Tecozam han sido insertadas correctamente.", vbInformation + vbOKOnly, "STAFF OFICINA")

    End Sub

    Public Sub insertaHorasJPStaffPortugal(ByVal mes As String, ByVal año As String, ByVal Fecha1 As String, ByVal Fecha2 As String, ByVal dtTecozam As DataTable)
        Dim IDOperario As String
        Dim IDOficio As String
        Dim IDCategoriaProfesionalSCCP As String
        Dim IDObra As String
        Dim IDTrabajo As String
        Dim IDAutonumerico As String
        Dim CodTrabajo As String
        Dim txtSQL As String
        Dim horas As Double = 0

        'Tabla que recoje los dias que no se trabaja, ya sea por vacacion o por festivo/fin de semana
        Dim dtOperarioCalendarioNoProductivo As New DataTable
        Dim dtCalendario As New DataTable
        dtCalendario = ObtieneCalendario(Fecha1, Fecha2)

        'TABLA CON FECHAS DONDE SE INSERTAN HORAS
        Dim dtDiasInsertar As New DataTable
        dtDiasInsertar.Columns.Add("Fecha", GetType(Date))

        Dim filas As Integer = 0
        PvProgreso.Value = 0 : PvProgreso.Maximum = dtTecozam.Rows.Count : PvProgreso.Step = 1 : PvProgreso.Visible = True

        For Each fila As DataRow In dtTecozam.Rows
            IDOperario = fila("IDOperario")
            IDOficio = DevuelveIDOficio(DB_DCZ, IDOperario)
            IDCategoriaProfesionalSCCP = DevuelveIDCategoriaProfesionalSCCP(DB_DCZ, IDOperario)
            Dim filtro As New Filter
            Dim dtObra As New DataTable
            filtro.Add("NObra", FilterOperator.Equal, fila("CentroCoste"))
            dtObra = New BE.DataEngine().Filter(DB_DCZ & "..tbObraCabecera", filtro)
            IDObra = dtObra.Rows(0)("IDObra").ToString
            IDTrabajo = ObtieneIDTrabajo(DB_DCZ, IDObra, "PT1")
            horas = 8 * fila("Porcentaje")

            dtOperarioCalendarioNoProductivo = ObtieneDiasVacacionesYFestivosJP(DB_TECOZAM, DB_DCZ, IDOperario, Fecha1, Fecha2)
            dtDiasInsertar = ObtieneFechasInsertar(DB_DCZ, IDOperario, dtCalendario, dtOperarioCalendarioNoProductivo)

            Windows.Forms.Application.DoEvents()
            LProgreso.Text = "Importando : " & IDOperario & " - DCZ JP"
            Windows.Forms.Application.DoEvents()

            For Each row As DataRow In dtDiasInsertar.Rows
                Dim fecha As Date = row.Field(Of Date)("Fecha")
                IDAutonumerico = auto.Autonumerico()

                Dim rsTrabajo As New DataTable : Dim filtro2 As New Filter
                filtro2.Add("IDObra", FilterOperator.Equal, IDObra) : filtro2.Add("IdTrabajo", FilterOperator.Equal, IDTrabajo)
                rsTrabajo = New BE.DataEngine().Filter(DB_DCZ & "..tbObraTrabajo", filtro2)
                'rsTrabajo = obraTrabajo.Filter(filtro2, , "IdTrabajo, CodTrabajo, DescTrabajo, IdTipoTrabajo, IdSubtipoTrabajo")

                IDTrabajo = rsTrabajo.Rows(0)("IdTrabajo") : CodTrabajo = rsTrabajo.Rows(0)("CodTrabajo")
                Dim DescTrabajo As String = "" : Dim IdTipoTrabajo As String = "" : Dim IdSubTipoTrabajo As String = ""
                DescTrabajo = rsTrabajo.Rows(0)("DescTrabajo") : IdTipoTrabajo = rsTrabajo.Rows(0)("IdTipoTrabajo") : IdSubTipoTrabajo = Nz(rsTrabajo.Rows(0)("IdSubtipotrabajo"), "")
                Dim DescParte As String : DescParte = "JP STAFF " & mes & "-" & año & "-JP"

                txtSQL = "Insert into " & DB_DCZ & "..tbObraMODControl(IdLineaModControl, IdTrabajo, IdObra, CodTrabajo, DescTrabajo, IdTipoTrabajo, " & _
                        "IdSubTipoTrabajo, IdOperario, IdCategoria, IdHora, FechaInicio, HorasRealMod, TasaRealModA, " & _
                         "ImpRealModA, HorasFactMod, ImpFactModA, DescParte, Facturable, FechaCreacionAudi, FechaModificacionAudi, Usuarioaudi, IDOficio, IdTipoTurno, HorasAdministrativas, IDCategoriaProfesionalSCCP) " & _
                         "Values(" & IDAutonumerico & ", " & IDTrabajo & ", " & IDObra & ", '" & _
                         CodTrabajo & "', '" & DescTrabajo & "', '" & IdTipoTrabajo & "', '" & _
                         IdSubTipoTrabajo & "', '" & IDOperario & "', 'PREDET', '" & _
                         "HA" & "', '" & fecha & "', 0 , " & 0 & ", " & 0 & _
                         ", 0 , " & 0 & _
                         ", '" & DescParte & "', " & 0 & ", '" & Date.Now.Date & "', '" & Date.Now.Date & "', '" & ExpertisApp.UserName & "','" & IDOficio & "', 4, '" & Replace(horas, ",", ".") & " ' ," & Nz(IDCategoriaProfesionalSCCP, "") & ")"

                auto.Ejecutar(txtSQL)
            Next

            filas = filas + 1
            PvProgreso.Value = filas
        Next
        '3. Obtengo una tabla por persona de los días que no tengan que insertar horas
        'MsgBox("Horas de la gente de oficina de Tecozam han sido insertadas correctamente.", vbInformation + vbOKOnly, "STAFF OFICINA")

    End Sub

    Public Sub insertaHorasJPStaffUK(ByVal mes As String, ByVal año As String, ByVal Fecha1 As String, ByVal Fecha2 As String, ByVal dtTecozam As DataTable)
        Dim IDOperario As String
        Dim IDOficio As String
        Dim IDCategoriaProfesionalSCCP As String
        Dim IDObra As String
        Dim IDTrabajo As String
        Dim IDAutonumerico As String
        Dim CodTrabajo As String
        Dim txtSQL As String
        Dim horas As Double = 0

        'Tabla que recoje los dias que no se trabaja, ya sea por vacacion o por festivo/fin de semana
        Dim dtOperarioCalendarioNoProductivo As New DataTable
        Dim dtCalendario As New DataTable
        dtCalendario = ObtieneCalendario(Fecha1, Fecha2)

        'TABLA CON FECHAS DONDE SE INSERTAN HORAS
        Dim dtDiasInsertar As New DataTable
        dtDiasInsertar.Columns.Add("Fecha", GetType(Date))

        Dim filas As Integer = 0
        PvProgreso.Value = 0 : PvProgreso.Maximum = dtTecozam.Rows.Count : PvProgreso.Step = 1 : PvProgreso.Visible = True

        For Each fila As DataRow In dtTecozam.Rows
            IDOperario = fila("IDOperario")
            IDOficio = DevuelveIDOficio(DB_UK, IDOperario)
            IDCategoriaProfesionalSCCP = DevuelveIDCategoriaProfesionalSCCP(DB_UK, IDOperario)
            Dim filtro As New Filter
            Dim dtObra As New DataTable
            filtro.Add("NObra", FilterOperator.Equal, fila("CentroCoste"))
            dtObra = New BE.DataEngine().Filter(DB_UK & "..tbObraCabecera", filtro)
            IDObra = dtObra.Rows(0)("IDObra").ToString
            IDTrabajo = ObtieneIDTrabajo(DB_UK, IDObra, "PT1")
            horas = 8 * fila("Porcentaje")

            dtOperarioCalendarioNoProductivo = ObtieneDiasVacacionesYFestivosJP(DB_TECOZAM, DB_UK, IDOperario, Fecha1, Fecha2)
            dtDiasInsertar = ObtieneFechasInsertarUK(DB_UK, IDOperario, dtCalendario, dtOperarioCalendarioNoProductivo)

            Windows.Forms.Application.DoEvents()
            LProgreso.Text = "Importando : " & IDOperario & " - UK JP"
            Windows.Forms.Application.DoEvents()

            For Each row As DataRow In dtDiasInsertar.Rows
                Dim fecha As Date = row.Field(Of Date)("Fecha")
                IDAutonumerico = auto.Autonumerico()

                Dim rsTrabajo As New DataTable : Dim filtro2 As New Filter
                filtro2.Add("IDObra", FilterOperator.Equal, IDObra) : filtro2.Add("IdTrabajo", FilterOperator.Equal, IDTrabajo)
                rsTrabajo = New BE.DataEngine().Filter(DB_UK & "..tbObraTrabajo", filtro2)
                'rsTrabajo = obraTrabajo.Filter(filtro2, , "IdTrabajo, CodTrabajo, DescTrabajo, IdTipoTrabajo, IdSubtipoTrabajo")

                IDTrabajo = rsTrabajo.Rows(0)("IdTrabajo") : CodTrabajo = rsTrabajo.Rows(0)("CodTrabajo")
                Dim DescTrabajo As String = "" : Dim IdTipoTrabajo As String = "" : Dim IdSubTipoTrabajo As String = ""
                DescTrabajo = rsTrabajo.Rows(0)("DescTrabajo") : IdTipoTrabajo = rsTrabajo.Rows(0)("IdTipoTrabajo") : IdSubTipoTrabajo = Nz(rsTrabajo.Rows(0)("IdSubtipotrabajo"), "")
                Dim DescParte As String : DescParte = "JP STAFF " & mes & "-" & año & "-JP"

                txtSQL = "Insert into " & DB_UK & "..tbObraMODControl(IdLineaModControl, IdTrabajo, IdObra, CodTrabajo, DescTrabajo, IdTipoTrabajo, " & _
                         "IdSubTipoTrabajo, IdOperario, IdCategoria, IdHora, FechaInicio, HorasRealMod, TasaRealModA, " & _
                         "ImpRealModA, HorasFactMod, ImpFactModA, DescParte, Facturable, FechaCreacionAudi, FechaModificacionAudi, Usuarioaudi, IDOficio, IdTipoTurno, HorasAdministrativas, IDCategoriaProfesionalSCCP) " & _
                         "Values(" & IDAutonumerico & ", " & IDTrabajo & ", " & IDObra & ", '" & _
                         CodTrabajo & "', '" & DescTrabajo & "', '" & IdTipoTrabajo & "', '" & _
                         IdSubTipoTrabajo & "', '" & IDOperario & "', 'PREDET', '" & _
                         "HA" & "', '" & fecha & "', 0 , " & 0 & ", " & 0 & _
                         ", 0 , " & 0 & _
                         ", '" & DescParte & "', " & 0 & ", '" & Date.Now.Date & "', '" & Date.Now.Date & "', '" & ExpertisApp.UserName & "','" & IDOficio & "', 4, '" & Replace(horas, ",", ".") & " ' ," & Nz(IDCategoriaProfesionalSCCP, "") & ")"

                auto.Ejecutar(txtSQL)
            Next

            filas = filas + 1
            PvProgreso.Value = filas
        Next
        '3. Obtengo una tabla por persona de los días que no tengan que insertar horas
        'MsgBox("Horas de la gente de oficina de Tecozam han sido insertadas correctamente.", vbInformation + vbOKOnly, "STAFF OFICINA")

    End Sub

    Public Sub insertaHorasJPStaffNO(ByVal mes As String, ByVal año As String, ByVal Fecha1 As String, ByVal Fecha2 As String, ByVal dtTecozam As DataTable)
        Dim IDOperario As String
        Dim IDOficio As String
        Dim IDCategoriaProfesionalSCCP As String
        Dim IDObra As String
        Dim IDTrabajo As String
        Dim IDAutonumerico As String
        Dim CodTrabajo As String
        Dim txtSQL As String
        Dim horas As Double = 0

        'Tabla que recoje los dias que no se trabaja, ya sea por vacacion o por festivo/fin de semana
        Dim dtOperarioCalendarioNoProductivo As New DataTable
        Dim dtCalendario As New DataTable
        dtCalendario = ObtieneCalendario(Fecha1, Fecha2)

        'TABLA CON FECHAS DONDE SE INSERTAN HORAS
        Dim dtDiasInsertar As New DataTable
        dtDiasInsertar.Columns.Add("Fecha", GetType(Date))

        Dim filas As Integer = 0
        PvProgreso.Value = 0 : PvProgreso.Maximum = dtTecozam.Rows.Count : PvProgreso.Step = 1 : PvProgreso.Visible = True

        For Each fila As DataRow In dtTecozam.Rows
            IDOperario = fila("IDOperario")
            IDOficio = DevuelveIDOficio(DB_NO, IDOperario)
            IDCategoriaProfesionalSCCP = DevuelveIDCategoriaProfesionalSCCP(DB_NO, IDOperario)
            Dim filtro As New Filter
            Dim dtObra As New DataTable
            filtro.Add("NObra", FilterOperator.Equal, fila("CentroCoste"))
            dtObra = New BE.DataEngine().Filter(DB_NO & "..tbObraCabecera", filtro)
            IDObra = dtObra.Rows(0)("IDObra").ToString
            IDTrabajo = ObtieneIDTrabajo(DB_NO, IDObra, "PT1")
            horas = 8 * fila("Porcentaje")

            dtOperarioCalendarioNoProductivo = ObtieneDiasVacacionesYFestivosJP(DB_TECOZAM, DB_NO, IDOperario, Fecha1, Fecha2)
            dtDiasInsertar = ObtieneFechasInsertarUK(DB_NO, IDOperario, dtCalendario, dtOperarioCalendarioNoProductivo)

            Windows.Forms.Application.DoEvents()
            LProgreso.Text = "Importando : " & IDOperario & " - NO JP"
            Windows.Forms.Application.DoEvents()

            For Each row As DataRow In dtDiasInsertar.Rows
                Dim fecha As Date = row.Field(Of Date)("Fecha")
                IDAutonumerico = auto.Autonumerico()

                Dim rsTrabajo As New DataTable : Dim filtro2 As New Filter
                filtro2.Add("IDObra", FilterOperator.Equal, IDObra) : filtro2.Add("IdTrabajo", FilterOperator.Equal, IDTrabajo)
                rsTrabajo = New BE.DataEngine().Filter(DB_NO & "..tbObraTrabajo", filtro2)
                'rsTrabajo = obraTrabajo.Filter(filtro2, , "IdTrabajo, CodTrabajo, DescTrabajo, IdTipoTrabajo, IdSubtipoTrabajo")

                IDTrabajo = rsTrabajo.Rows(0)("IdTrabajo") : CodTrabajo = rsTrabajo.Rows(0)("CodTrabajo")
                Dim DescTrabajo As String = "" : Dim IdTipoTrabajo As String = "" : Dim IdSubTipoTrabajo As String = ""
                DescTrabajo = rsTrabajo.Rows(0)("DescTrabajo") : IdTipoTrabajo = rsTrabajo.Rows(0)("IdTipoTrabajo") : IdSubTipoTrabajo = Nz(rsTrabajo.Rows(0)("IdSubtipotrabajo"), "")
                Dim DescParte As String : DescParte = "JP STAFF " & mes & "-" & año & "-JP"

                txtSQL = "Insert into " & DB_NO & "..tbObraMODControl(IdLineaModControl, IdTrabajo, IdObra, CodTrabajo, DescTrabajo, IdTipoTrabajo, " & _
                         "IdSubTipoTrabajo, IdOperario, IdCategoria, IdHora, FechaInicio, HorasRealMod, TasaRealModA, " & _
                         "ImpRealModA, HorasFactMod, ImpFactModA, DescParte, Facturable, FechaCreacionAudi, FechaModificacionAudi, Usuarioaudi, IDOficio, IdTipoTurno, HorasAdministrativas, IDCategoriaProfesionalSCCP) " & _
                         "Values(" & IDAutonumerico & ", " & IDTrabajo & ", " & IDObra & ", '" & _
                         CodTrabajo & "', '" & DescTrabajo & "', '" & IdTipoTrabajo & "', '" & _
                         IdSubTipoTrabajo & "', '" & IDOperario & "', 'PREDET', '" & _
                         "HA" & "', '" & fecha & "', 0 , " & 0 & ", " & 0 & _
                         ", 0 , " & 0 & _
                         ", '" & DescParte & "', " & 0 & ", '" & Date.Now.Date & "', '" & Date.Now.Date & "', '" & ExpertisApp.UserName & "','" & IDOficio & "', 4, '" & Replace(horas, ",", ".") & " ' ," & Nz(IDCategoriaProfesionalSCCP, "") & ")"

                auto.Ejecutar(txtSQL)
            Next

            filas = filas + 1
            PvProgreso.Value = filas
        Next
        '3. Obtengo una tabla por persona de los días que no tengan que insertar horas
        'MsgBox("Horas de la gente de oficina de Tecozam han sido insertadas correctamente.", vbInformation + vbOKOnly, "STAFF OFICINA")

    End Sub

    Private Sub btnAceptar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAceptar.Click
        Dim hoja As String
        'hoja = "% HORAS J.P. y STAFF"
        hoja = "HORAS"
        Dim dt As New DataTable

        Dim ruta As String
        ruta = lblRuta.Text
        Dim rango As String = "A2:F10000"
        dt = ObtenerDatosExcel(ruta, hoja, rango)

        dt.Columns("F1").ColumnName = "IDOperario"
        dt.Columns("F2").ColumnName = "DescOperario"
        dt.Columns("F3").ColumnName = "Empresa"
        dt.Columns("F4").ColumnName = "CentroCoste"
        dt.Columns("F5").ColumnName = "ProduccionSinVentas"
        dt.Columns("F6").ColumnName = "Porcentaje"

        Dim mes As String = ""
        Dim año As String = ""

        mes = ruta.Substring(ruta.Length - 9, 2)
        año = ruta.Substring(ruta.Length - 7, 2)
        año = "20" & año
        'Formo Fechas para sacar los turnos
        Dim Fecha1 As String
        Dim Fecha2 As String
        Fecha1 = "01/" & mes & "/" & año
        Dim diaMes As String
        diaMes = ObtieneDiaUltimoMes(mes, año)
        Fecha2 = diaMes & "/" & mes & "/" & año & ""

        'Limpia donde el porcentaje hay 0 y acota la tabla por abajo
        dt = ChecksPrevios(dt)

        Dim dtTecozam As New DataTable
        Dim dtPortugal As New DataTable
        Dim dtUK As New DataTable
        Dim dtNO As New DataTable

        Dim flat As Integer
        'FILTRO LOS REGISTROS DE TECOZAM 'FILTRO LOS REGISTROS DE DCZ 'FILTROS LOS REGISTROS DE UK
        flat = CargaTablas(dtTecozam, dtPortugal, dtUK, dtNO, dt)

        If flat = 0 Then
            MsgBox("Existen registros que no coinciden con ninguna empresa.")
            Exit Sub
        End If

        Dim result As DialogResult = MessageBox.Show("Hay " & dtTecozam.Rows.Count & " registros de T. ES." & vbCrLf & _
        "Hay " & dtPortugal.Rows.Count & " registros de D. P." & vbCrLf & _
        "Hay " & dtUK.Rows.Count & " registros de T. UK." & vbCrLf & _
        "Hay " & dtNO.Rows.Count & " registros de T. NO." & vbCrLf, "¿Están correctos estos datos?", MessageBoxButtons.YesNo)
        If result = DialogResult.No Then
            Exit Sub
        End If

        '-------SOBREESCRIBIR HORAS POR MES Y AÑO NATURAL---------
        'If SobreescribirHoras(Fecha1, Fecha2) = True Then
        '    Dim result2 As DialogResult = MessageBox.Show("Exiten horas de JP y STAFF entre este rango de fechas, ¿desea sobreescribir los datos?", "Borrar e insertar datos.", MessageBoxButtons.YesNo)
        '    If result2 = DialogResult.Yes Then
        '        BorrarDatos(mes, año)
        '    Else
        '        Exit Sub
        '    End If
        'Else
        'End If
        '-------------------------------------------------------

        '--------------INICIO CHECKS---------------------------
        Dim bandera As Integer
        bandera = CheckRegistrosEmpresa(dtTecozam, DB_TECOZAM)
        If bandera = 0 Then
            Exit Sub
        End If

        bandera = CheckRegistrosEmpresa(dtPortugal, DB_DCZ)
        If bandera = 0 Then
            Exit Sub
        End If

        bandera = CheckRegistrosEmpresa(dtUK, DB_UK)
        If bandera = 0 Then
            Exit Sub
        End If

        bandera = CheckRegistrosEmpresa(dtNO, DB_NO)
        If bandera = 0 Then
            Exit Sub
        End If
        '---------------FIN CHECKS---------------------------

        'Inserta horas en Tecozam
        insertaHorasJPStaffTecozam(mes, año, Fecha1, Fecha2, dtTecozam)
        'Inserta horas en Portugal
        insertaHorasJPStaffPortugal(mes, año, Fecha1, Fecha2, dtPortugal)
        'Inserta horas en UK
        insertaHorasJPStaffUK(mes, año, Fecha1, Fecha2, dtUK)
        'Inserta horas en NO
        insertaHorasJPStaffNO(mes, año, Fecha1, Fecha2, dtNO)

        MsgBox("Horas desde Excel cargadas correctamente")
    End Sub
    Public Function ChecksPrevios(ByVal dt As DataTable) As DataTable
        'CHECK 1.SI ES =0 PORCENTAJE SE BORRA
        Dim nuevaTabla As New DataTable
        For Each columna As DataColumn In dt.Columns
            nuevaTabla.Columns.Add(columna.ColumnName, columna.DataType)
        Next

        ' Itera por las filas de la tabla original
        For Each fila As DataRow In dt.Rows
            ' Verifica si la celda en la columna "IDOperario" está vacía
            If Len(fila("IDOperario").ToString) = 0 Then
                ' La celda en la columna "IDOperario" está vacía, salir del bucle
                Exit For
            End If

            ' Agrega la fila a la nueva tabla
            nuevaTabla.Rows.Add(fila.ItemArray)
        Next

        'CHECK 2.CORTAR DT DONDE IDOPERARIO
        ' Itera por las filas de la tabla en reversa
        For i As Integer = nuevaTabla.Rows.Count - 1 To 0 Step -1
            ' Verifica si el valor de la celda en la columna es cero
            If nuevaTabla.Rows(i)("Porcentaje") = 0 Then
                ' Elimina la fila
                nuevaTabla.Rows.RemoveAt(i)
            End If
        Next
        Return nuevaTabla
    End Function

    Public Function CheckRegistrosEmpresa(ByVal dtTecozam As DataTable, ByVal basededatos As String) As Integer
        'En este For se hacen los CHECKS necesarios
        Dim IDOperario As String = ""
        Dim CategoriaSSCP As String = ""
        Dim IDObra As String = ""
        Dim dtObra As DataTable
        For Each dr As DataRow In dtTecozam.Rows
            IDOperario = dr("IDOperario")
            Try
                CategoriaSSCP = ObtieneCategoriaIDOperario(IDOperario, basededatos)
                If CategoriaSSCP.ToString.Length = 0 Or (CategoriaSSCP.ToString <> 1 And CategoriaSSCP.ToString <> 4 And CategoriaSSCP.ToString <> 5) Then
                    MsgBox("Existe error al asociar CategoriaSCCP en el operario " & IDOperario & " en " & basededatos & ".", vbOKCancel + vbCritical, "Aviso")
                    Return 0
                End If
            Catch ex As Exception
                MsgBox("Existe error al asociar CategoriaSCCP en el operario " & IDOperario & " en " & basededatos & ".", vbOKCancel + vbCritical, "Aviso")
                Return 0
            End Try

            Dim NObra As String = dr("CentroCoste").ToString
            Try
                Dim filtro As New Filter
                filtro.Add("NObra", FilterOperator.Equal, NObra)
                dtObra = New BE.DataEngine().Filter(basededatos & "..tbObraCabecera", filtro)
                If dtObra.Rows.Count = 0 Then
                    MsgBox("No existe la obra " & NObra & " en " & basededatos & ".", vbOKCancel + vbCritical, "Aviso")
                    Return 0
                End If
                IDObra = dtObra.Rows(0)("IDObra").ToString

                Dim dtTrab As New DataTable
                Dim fil As New Filter
                fil.Add("IDObra", FilterOperator.Equal, IDObra)
                fil.Add("CodTrabajo", FilterOperator.Equal, "PT1")
                dtTrab = New BE.DataEngine().Filter(basededatos & "..tbObraTrabajo", fil)

                If dtTrab.Rows.Count = 0 Then
                    MsgBox("No existe partes de horas asignado a la obra " & NObra & " en " & basededatos & ".", vbOKCancel + vbCritical, "Aviso")
                    Return 0
                End If
            Catch ex As Exception
                MsgBox("No existe partes de horas asignado a la obra " & NObra & " en " & basededatos & ".", vbOKCancel + vbCritical, "Aviso")
                Return 0
            End Try
        Next
        Return 1
    End Function

    Public Sub BorrarDatos(ByVal mesP As String, ByVal anioP As String)
        Dim DescParte As String
        DescParte = "%" & mesP & "-" & anioP & "-JP"

        auto.BorraDatosObraMODControlHorasAdministrativas(DescParte, DB_TECOZAM)
        auto.BorraDatosObraMODControlHorasAdministrativas(DescParte, DB_DCZ)
        auto.BorraDatosObraMODControlHorasAdministrativas(DescParte, DB_UK)
    End Sub

    Public Function SobreescribirHoras(ByVal Fecha1 As String, ByVal Fecha2 As String) As Boolean
        Dim dtHoras As New DataTable
        Dim filtro As New Filter
        filtro.Add("FechaInicio", FilterOperator.GreaterThanOrEqual, Fecha1)
        filtro.Add("FechaInicio", FilterOperator.LessThanOrEqual, Fecha2)
        filtro.Add("HorasAdministrativas", FilterOperator.NotEqual, DBNull.Value)
        dtHoras = New BE.DataEngine().Filter("tbObraModControl", filtro)

        If dtHoras.Rows.Count = 0 Then
            Return False
        Else
            Return True
        End If

    End Function

    Public Function ObtieneIDTrabajo(ByVal basededatos As String, ByVal IDObra As String, ByVal CodTrabajo As String) As String
        Dim dtTrabajo As New DataTable
        Dim filtro As New Filter
        filtro.Add("IDObra", FilterOperator.Equal, IDObra)
        filtro.Add("CodTrabajo", FilterOperator.Equal, CodTrabajo)

        dtTrabajo = New BE.DataEngine().Filter(basededatos & "..tbObraTrabajo", filtro)
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

    Public Function DevuelveIDCategoriaProfesionalSCCP(ByVal basededatos As String, ByVal IDOperario As String) As Integer
        Dim dt As New DataTable
        Dim f As New Filter

        f.Add("IDOperario", FilterOperator.Equal, IDOperario)
        dt = New BE.DataEngine().Filter(basededatos & "..vOperarioCategoriaProf", f)
        If dt.Rows.Count > 0 Then
            Return Nz(dt(0)("Abreviatura"), 0)
        Else
            Return 0
        End If
    End Function
    Public Function DevuelveIDOficio(ByVal basededatos As String, ByVal IDOperario As String) As String
        Dim dt As New DataTable
        Dim f As New Filter

        f.Add("IDOperario", FilterOperator.Equal, IDOperario)
        dt = New BE.DataEngine().Filter(basededatos & "..tbMaestroOperario", f)
        If dt.Rows.Count > 0 Then
            Return Nz(dt(0)("IDOficio"), "")
        Else
            Return ""
        End If
    End Function


    Public Function ObtieneFechasInsertar(ByVal basededatos As String, ByVal IDOperario As String, ByVal dtCalendario As DataTable, ByVal dtOperarioCalendarioNoProductivo As DataTable) As DataTable
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
        dtOperario = New BE.DataEngine().Filter(basededatos & "..frmMntoOperario", f)

        'SI SE DA DE ALTA EL 15 DE UN MES SE INSERTA A PARTIR DEL DIA 15
        If Len(dtOperario.Rows(0)("FechaAlta").ToString) <> 0 Then
            Dim fechaAlta As String
            fechaAlta = dtOperario.Rows(0)("FechaAlta").ToString

            For i As Integer = dtDiasInsertar.Rows.Count - 1 To 0 Step -1
                Dim fila As DataRow = dtDiasInsertar.Rows(i)
                Dim fecha As Date = CDate(fila("Fecha"))

                If fecha < fechaAlta Then
                    ' La fecha es mayor que la fecha límite, eliminamos la fila
                    dtDiasInsertar.Rows.RemoveAt(i)
                End If
            Next
        End If

        'SI TE DAS DE BAJA EL 15 DE UN MES BORRA LOS SIGUIENTE 15 DIAS A INSERTAR
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

    Public Function ObtieneFechasInsertarUK(ByVal basededatos As String, ByVal IDOperario As String, ByVal dtCalendario As DataTable, ByVal dtOperarioCalendarioNoProductivo As DataTable) As DataTable
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
        dtOperario = New BE.DataEngine().Filter(basededatos & "..frmMntoOperario", f)

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

    Public Function ObtieneCalendario(ByVal Fecha1 As String, ByVal Fecha2 As String) As DataTable
        Dim dtCalendario As New DataTable

        Dim filtro As New Filter
        filtro.Add("Fecha", FilterOperator.GreaterThanOrEqual, Fecha1)
        filtro.Add("Fecha", FilterOperator.LessThanOrEqual, Fecha2)
        dtCalendario = New BE.DataEngine().Filter(DB_TECOZAM & "..tbMaestroFechas", filtro)

        Return dtCalendario
    End Function

    Public Function ObtieneDiasVacacionesYFestivos(ByVal basededatosteco As String, ByVal basededatosoriginal As String, ByVal IDOperario As String, ByVal Fecha1 As String, ByVal Fecha2 As String, ByVal dtDias As DataTable) As DataTable
        Dim dtVacaciones As New DataTable
        Dim dtFestivos As New DataTable
        Dim dtTrabajados As New DataTable
        Dim dtDiasCambioDeObra As New DataTable

        Dim filtro As New Filter
        'DIA DE VACACIONES = 2
        filtro.Add("TipoDia", FilterOperator.Equal, 2)
        filtro.Add("IDOperario", FilterOperator.Equal, IDOperario)
        filtro.Add("Fecha", FilterOperator.GreaterThanOrEqual, Fecha1)
        filtro.Add("Fecha", FilterOperator.LessThanOrEqual, Fecha2)

        dtVacaciones = New BE.DataEngine().Filter(basededatosoriginal & "..tbCalendarioOperario", filtro, "Fecha, TipoDia")
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
        dtTrabajados = New BE.DataEngine().Filter(basededatosoriginal & "..tbObraModControl", filtro, "FechaInicio As Fecha")
        filtro.Clear()

        'PARA AQUELLAS PERSONAS QUE FORMAN PARTE DE LAS QUE HAN ESTADO HASTA ALGUN DIA EN OFICINA
        'Y LUEGO SE HAN CAMBIADO A OBRA Y POR TANTO NO HAY QUE CARGARLE DIAS A PARTIR DE LA FECHA
        'QUE ESTA EN TBHISTORICOPERSONAL
        filtro.Add("Fecha", FilterOperator.GreaterThanOrEqual, Fecha1)
        filtro.Add("Fecha", FilterOperator.LessThanOrEqual, Fecha2)
        filtro.Add("IDOperario", FilterOperator.Equal, IDOperario)
        dtDiasCambioDeObra = New BE.DataEngine().Filter(basededatosoriginal & "..tbHistoricoPersonal", filtro, "Fecha")


        If dtDiasCambioDeObra.Rows.Count <> 0 Then
            Dim diaLimite As String
            diaLimite = dtDiasCambioDeObra.Rows(0)("Fecha").ToString
            For Each dr As DataRow In dtDias.Rows
                If dr("Fecha") >= diaLimite Then
                    dtDiasCambioDeObra.ImportRow(dr)
                End If
            Next
        End If

        ' Crear un nuevo DataTable llamado dtCalendario
        Dim dtCalendario As New DataTable
        ' Agregar las columnas Fecha
        dtCalendario.Columns.Add("Fecha", GetType(Date))
        'dtCalendario.Columns.Add("TipoDia", GetType(Integer))
        ' Unir los DataTables dtVacaciones y dtFestivos, trabajador y con cambio de obra en el DataTable dtCalendario
        'dtCalendario.Merge(dtVacaciones)
        dtCalendario.Merge(dtFestivos)
        dtCalendario.Merge(dtTrabajados)
        dtCalendario.Merge(dtDiasCambioDeObra)

        Return dtCalendario
    End Function

    Public Function ObtieneDiasVacacionesYFestivosJP(ByVal basededatosteco As String, ByVal basededatosoriginal As String, ByVal IDOperario As String, ByVal Fecha1 As String, ByVal Fecha2 As String) As DataTable
        Dim dtVacaciones As New DataTable
        Dim dtFestivos As New DataTable
        Dim dtTrabajados As New DataTable

        Dim filtro As New Filter
        'DIA DE VACACIONES = 2
        filtro.Add("TipoDia", FilterOperator.Equal, 2)
        filtro.Add("IDOperario", FilterOperator.Equal, IDOperario)
        filtro.Add("Fecha", FilterOperator.GreaterThanOrEqual, Fecha1)
        filtro.Add("Fecha", FilterOperator.LessThanOrEqual, Fecha2)

        dtVacaciones = New BE.DataEngine().Filter(basededatosoriginal & "..tbCalendarioOperario", filtro, "Fecha, TipoDia")
        filtro.Clear()
        'FESTIVOS Y FINDES = 1
        filtro.Add("TipoDia", FilterOperator.Equal, 1)
        filtro.Add("IDCentro", FilterOperator.Equal, "00")
        filtro.Add("Fecha", FilterOperator.GreaterThanOrEqual, Fecha1)
        filtro.Add("Fecha", FilterOperator.LessThanOrEqual, Fecha2)
        dtFestivos = New BE.DataEngine().Filter(basededatosteco & "..tbCalendarioCentro", filtro, "Fecha, TipoDia")

        ' Crear un nuevo DataTable llamado dtCalendario
        Dim dtCalendario As New DataTable()

        ' Agregar las columnas Fecha y TipoDia al DataTable
        dtCalendario.Columns.Add("Fecha", GetType(Date))
        'dtCalendario.Columns.Add("TipoDia", GetType(Integer))

        ' Unir los DataTables dtVacaciones y dtFestivos en el DataTable dtCalendario
        'dtCalendario.Merge(dtVacaciones)
        dtCalendario.Merge(dtFestivos)

        Return dtCalendario
    End Function

    Public Function ObtieneIDOperario(ByVal DNI As String) As String
        Dim dtOperario As New DataTable
        Dim filtro As New Filter
        filtro.Add("DNI", FilterOperator.Equal, DNI)
        dtOperario = New BE.DataEngine().Filter("tbMaestroOperario", filtro)

        Return dtOperario.Rows(0)("IDOperario").ToString
    End Function

    Public Function ObtieneCategoriaIDOperario(ByVal IDOperario As String, ByVal basededatos As String) As String
        Dim dtOperario As New DataTable
        Dim filtro As New Filter
        filtro.Add("IDOperario", FilterOperator.Equal, IDOperario)
        dtOperario = New BE.DataEngine().Filter(basededatos & "..vOperarioCategoriaProf", filtro)

        Return dtOperario.Rows(0)("Abreviatura").ToString
    End Function
    Public Function ObtenerDatosExcel(ByVal ruta As String, ByVal hoja As String, ByVal rango As String) As DataTable
        Dim MyConnection As System.Data.OleDb.OleDbConnection
        Dim DtSet As System.Data.DataSet
        Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
        MyConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & ruta & "';Extended Properties='Excel 8.0;HDR=NO'")
        MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [" & hoja & "$" & rango & "]", MyConnection)
        'MyCommand.TableMappings.Add("Table", "Net-informations.com")
        DtSet = New System.Data.DataSet
        MyCommand.Fill(DtSet)
        Dim dt As DataTable = DtSet.Tables(0)
        MyConnection.Close()

        Return dt

    End Function

    Private Sub bHorasOficina_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bHorasOficina.Click
        '0. Saber que mes y año productivo hay que insertar las horas en la oficina
        'Dim mes As String
        'Dim año As String

        'mes = InputBox("Introduzca el mes natural", "Formato: mm")
        'año = InputBox("Introduzca el año natural", "Formato: aaaa")

        'Dim Fecha1 As String
        'Dim Fecha2 As String

        'Fecha1 = "01/" & mes & "/" & año
        'Dim diaMes As String
        'diaMes = ObtieneDiaUltimoMes(mes, año)
        'Fecha2 = diaMes & "/" & mes & "/" & año & ""
        Dim frm As New frmInformeFecha
        frm.ShowDialog()
        Dim Fecha1 As String : Dim Fecha2 As String
        Fecha1 = frm.fecha1 : Fecha2 = frm.fecha2

        Dim mes As String : mes = Month(Fecha1)
        If Length(mes) = 1 Then
            mes = "0" & mes
        End If

        Dim anio As String
        anio = Year(Fecha1)

        If frm.blEstado = True Then
            '-----------TECOZAM--------------
            setHorasOficinaTecozam(mes, anio, Fecha1, Fecha2)
            '-----------FERRALLAS------------
            setHorasOficinaFerrallas(mes, anio, Fecha1, Fecha2)
            '-----------SECOZAM--------------
            setHorasOficinaSecozam(mes, anio, Fecha1, Fecha2)
            '-----------DCZ(No hay nadie)------------------
            'setHorasOficinaDCZ(mes, año, Fecha1, Fecha2)
            '-----------UK--------------

            '-----------ESLOVAQUIA------------

            '-----------SUECIA--------------

            '-----------NORUEGA------------
        End If

    End Sub
    Public Function getListadoPersonasOfiFerrallas(ByVal Fecha1 As String, ByVal Fecha2 As String) As DataTable
        Dim dt As New DataTable
        Dim sql As String
        '------COMO YO LO DE DEJARIA---------
        sql = "select IDOperario, Obra_Predeterminada from " & DB_FERRALLAS & "..tbMaestroOperarioSat " & _
        "where (Obra_Predeterminada='12677838' Or Obra_Predeterminada='12677615' Or Obra_Predeterminada='12678141') and " & _
        "(Fecha_Baja is null or (Fecha_Baja>='" & Fecha1 & "' and Fecha_Baja<=GETDATE()))" & _
        " Union " & _
        "select IDOperario, Proyecto AS Obra_Predeterminada " & _
        "from " & DB_FERRALLAS & "..tbHistoricoPersonal " & _
        "where (Proyecto = '12677838' OR Proyecto = '12677615' OR Proyecto = '12678141') and " & _
        "((Fecha >= '" & Fecha1 & "' AND Fecha <=GETDATE()))"

        dt = aux.EjecutarSqlSelect(sql)
        Return dt
    End Function

    Public Function getListadoPersonasOfiDCZ(ByVal Fecha1 As String, ByVal Fecha2 As String) As DataTable
        Dim dt As New DataTable
        '----------FORMA BUENA'-------------
        Dim sql As String
        sql = "select IDOperario, Obra_Predeterminada from " & DB_DCZ & "..tbMaestroOperarioSat " & _
        "where Obra_Predeterminada='11860026' and " & _
        "(Fecha_Baja is null or (Fecha_Baja>='" & Fecha1 & "' and Fecha_Baja<=GETDATE()))" & _
        " Union " & _
        "select IDOperario, Proyecto AS Obra_Predeterminada " & _
        "from " & DB_DCZ & "..tbHistoricoPersonal " & _
        "where (Proyecto = '11860026') and " & _
        "((Fecha >= '" & Fecha1 & "' AND Fecha <=GETDATE()))"

        dt = aux.EjecutarSqlSelect(sql)
        Return dt
    End Function

    Public Function getListadoPersonasOfiTecozam(ByVal Fecha1 As String, ByVal Fecha2 As String) As DataTable
        Dim dt As New DataTable
        Dim sql As String
        '------COMO YO LO DE DEJARIA---------
        sql = "select IDOperario, Obra_Predeterminada from " & DB_TECOZAM & "..tbMaestroOperarioSat " & _
        "where (Obra_Predeterminada='16895681' Or Obra_Predeterminada='11984995') and " & _
        "(Fecha_Baja is null or (Fecha_Baja>='" & Fecha1 & "' and Fecha_Baja<=GETDATE()))" & _
        " Union " & _
        "select IDOperario, Proyecto AS Obra_Predeterminada " & _
        "from " & DB_TECOZAM & "..tbHistoricoPersonal " & _
        "where (Proyecto = '16895681' OR Proyecto = '11984995') and " & _
        "((Fecha >= '" & Fecha1 & "' AND Fecha <=GETDATE()))"


        'sql = "select IDOperario, Obra_Predeterminada from DB_TECOZAM..tbMaestroOperarioSat where idoperario='T3450'"
        dt = aux.EjecutarSqlSelect(sql)
        Return dt
    End Function

    Public Function getListadoPersonasOfiSecozam(ByVal Fecha1 As String, ByVal Fecha2 As String) As DataTable
        Dim dt As New DataTable
        Dim sql As String
        '------COMO YO LO DE DEJARIA---------
        sql = "select IDOperario, Obra_Predeterminada from " & DB_SECOZAM & "..tbMaestroOperarioSat " & _
        "where (Obra_Predeterminada='11854299' Or Obra_Predeterminada='11854231') and " & _
        "(Fecha_Baja is null or (Fecha_Baja>='" & Fecha1 & "' and Fecha_Baja<=GETDATE()))" & _
        " Union " & _
        "select IDOperario, Proyecto AS Obra_Predeterminada " & _
        "from " & DB_SECOZAM & "..tbHistoricoPersonal " & _
        "where (Proyecto = '11854299' OR Proyecto = '11854231') and " & _
        "((Fecha >= '" & Fecha1 & "' AND Fecha <=GETDATE()))"

        dt = aux.EjecutarSqlSelect(sql)
        Return dt
    End Function

    Public Function DevuelveUltimoCambioObra(ByVal IDOperario As String, ByVal bbdd As String) As String
        Dim f As New Filter
        Dim dt As New DataTable
        f.Add("IDOperario", FilterOperator.Equal, IDOperario)

        dt = New BE.DataEngine().Filter(bbdd & "..tbHistoricoPersonal", f, , "Fecha desc")

        If dt.Rows.Count > 0 Then
            Return dt.Rows(0)("Proyecto")
        Else
            Return ""
        End If
    End Function
    Public Sub setHorasOficinaTecozam(ByVal mes As String, ByVal año As String, ByVal Fecha1 As String, ByVal Fecha2 As String)
        '1. Obtengo la tabla de personas que estén en oficina
        Dim dtPersonasOfi As New DataTable
        dtPersonasOfi = getListadoPersonasOfiTecozam(Fecha1, Fecha2)
        '2. Recorro las personas 
        Dim IDOperario As String
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

        Dim filas As Integer = 0
        PvProgreso.Value = 0 : PvProgreso.Maximum = dtPersonasOfi.Rows.Count : PvProgreso.Step = 1 : PvProgreso.Visible = True

        For Each fila As DataRow In dtPersonasOfi.Rows
            IDOperario = fila("IDOperario")
            IDOficio = DevuelveIDOficio(DB_TECOZAM, IDOperario)
            IDCategoriaProfesionalSCCP = DevuelveIDCategoriaProfesionalSCCP(DB_TECOZAM, IDOperario)
            'IDObra = "15330631"
            'IDObra destino = OFICINA
            IDObra = fila("Obra_Predeterminada")
            'Si es distinto que oficina y almacen
            If IDObra <> "11984995" And IDObra <> "16895681" Then
                IDObra = DevuelveUltimoCambioObra(IDOperario, DB_TECOZAM)
            End If

            IDTrabajo = ObtieneIDTrabajo(DB_TECOZAM, IDObra, "PT1")

            dtOperarioCalendarioNoProductivo = ObtieneDiasVacacionesYFestivos(DB_TECOZAM, DB_TECOZAM, IDOperario, Fecha1, Fecha2, dtCalendario)
            dtDiasInsertar = ObtieneFechasInsertar(DB_TECOZAM, IDOperario, dtCalendario, dtOperarioCalendarioNoProductivo)

            Windows.Forms.Application.DoEvents()
            LProgreso.Text = "Importando : " & IDOperario & " - TECOZAM OFICINA"
            Windows.Forms.Application.DoEvents()

            For Each row As DataRow In dtDiasInsertar.Rows
                Dim fecha As Date = row.Field(Of Date)("Fecha")
                IDAutonumerico = auto.Autonumerico()

                'Check de si hay un registro para esta fecha y este operario que pase al siguiente dia.
                'ESTO SE DEBE PORQUE SI ESTA DE BAJA Y TIENE HORAS NO SE GENERAN ADMINISTRATIVAS
                Dim checkSeguir As Boolean
                checkSeguir = getSiInsertarONo(DB_TECOZAM, fecha, IDOperario)

                If checkSeguir Then
                    Dim rsTrabajo As New DataTable : Dim filtro2 As New Filter
                    filtro2.Add("IDObra", FilterOperator.Equal, IDObra) : filtro2.Add("IdTrabajo", FilterOperator.Equal, IDTrabajo)
                    rsTrabajo = New BE.DataEngine().Filter(DB_TECOZAM & "..tbObraTrabajo", filtro2)
                    'rsTrabajo = obraTrabajo.Filter(filtro2, , "IdTrabajo, CodTrabajo, DescTrabajo, IdTipoTrabajo, IdSubtipoTrabajo")

                    IDTrabajo = rsTrabajo.Rows(0)("IdTrabajo") : CodTrabajo = rsTrabajo.Rows(0)("CodTrabajo")
                    Dim DescTrabajo As String = "" : Dim IdTipoTrabajo As String = "" : Dim IdSubTipoTrabajo As String = ""
                    DescTrabajo = rsTrabajo.Rows(0)("DescTrabajo") : IdTipoTrabajo = rsTrabajo.Rows(0)("IdTipoTrabajo") : IdSubTipoTrabajo = Nz(rsTrabajo.Rows(0)("IdSubtipotrabajo"), "")
                    Dim DescParte As String : DescParte = "OFICINA" & " " & mes & "-" & año & "-OFI"
                    '08/11/23 Para corregir esos que tienen reducción de jornada
                    Dim horas As Double = 8
                    Dim porcentaje As Double
                    Dim dtPorcentaje As DataTable
                    Dim f As New Filter
                    f.Add("IDOperario", FilterOperator.Equal, IDOperario)
                    dtPorcentaje = New BE.DataEngine().Filter(DB_TECOZAM & "..frmMntoOperario", f)
                    porcentaje = Nz(dtPorcentaje.Rows(0)("JornadaParcial"), 0)
                    If porcentaje <> 0 Then
                        horas = (horas * porcentaje) / 100
                    End If

                    txtSQL = "Insert into " & DB_TECOZAM & "..tbObraMODControl(IdLineaModControl, IdTrabajo, IdObra, CodTrabajo, DescTrabajo, IdTipoTrabajo, " & _
                            "IdSubTipoTrabajo, IdOperario, IdCategoria, IdHora, FechaInicio, HorasRealMod, TasaRealModA, " & _
                             "ImpRealModA, HorasFactMod, ImpFactModA, DescParte, Facturable, FechaCreacionAudi, FechaModificacionAudi, Usuarioaudi, IDOficio, IdTipoTurno, HorasAdministrativas, IDCategoriaProfesionalSCCP) " & _
                             "Values(" & IDAutonumerico & ", " & IDTrabajo & ", " & IDObra & ", '" & _
                             CodTrabajo & "', '" & DescTrabajo & "', '" & IdTipoTrabajo & "', '" & _
                             IdSubTipoTrabajo & "', '" & IDOperario & "', 'PREDET', '" & _
                             "HA" & "', '" & fecha & "', 0 , " & 0 & ", " & 0 & _
                             ", 0 , " & 0 & _
                             ", '" & DescParte & "', " & 0 & ", '" & Date.Now.Date & "', '" & Date.Now.Date & "', '" & ExpertisApp.UserName & "','" & IDOficio & "', 4,'" & horas.ToString.Replace(",", ".") & "' ," & Nz(IDCategoriaProfesionalSCCP, "") & ")"

                    auto.Ejecutar(txtSQL)
                End If
            Next
            filas = filas + 1
            PvProgreso.Value = filas
        Next
        '3. Obtengo una tabla por persona de los días que no tengan que insertar horas
        'MsgBox("Horas de la gente de oficina de Tecozam han sido insertadas correctamente.", vbInformation + vbOKOnly, "STAFF OFICINA")

    End Sub
    Public Function getSiInsertarONo(ByVal bbdd As String, ByVal fecha As String, ByVal IDOperario As String) As Boolean
        Dim f As New Filter
        f.Add("FechaInicio", FilterOperator.Equal, fecha)
        f.Add("IDOperario", FilterOperator.Equal, IDOperario)

        Dim dt As New DataTable
        dt = New BE.DataEngine().Filter(bbdd & "..tbObraModControl", f)

        If dt.Rows.Count = 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub setHorasOficinaFerrallas(ByVal mes As String, ByVal año As String, ByVal Fecha1 As String, ByVal Fecha2 As String)
        '1. Obtengo la tabla de personas que estén en oficina
        Dim dtPersonasOfi As New DataTable
        dtPersonasOfi = getListadoPersonasOfiFerrallas(Fecha1, Fecha2)
        '2. Recorro las personas 
        Dim IDOperario As String
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

        Dim filas As Integer = 0
        PvProgreso.Value = 0 : PvProgreso.Maximum = dtPersonasOfi.Rows.Count : PvProgreso.Step = 1 : PvProgreso.Visible = True

        For Each fila As DataRow In dtPersonasOfi.Rows

            IDOperario = fila("IDOperario")
            IDOficio = DevuelveIDOficio(DB_FERRALLAS, IDOperario)
            IDCategoriaProfesionalSCCP = DevuelveIDCategoriaProfesionalSCCP(DB_FERRALLAS, IDOperario)
            'IDObra = "12677615"
            IDObra = fila("Obra_Predeterminada")
            'Si es distinto que  ferrallas, oficina y secozam
            If IDObra <> "12677838" And IDObra <> "12677615" And IDObra <> "12678141" Then
                IDObra = DevuelveUltimoCambioObra(IDOperario, DB_FERRALLAS)
            End If

            IDTrabajo = ObtieneIDTrabajo(DB_FERRALLAS, IDObra, "PT1")
            'Este es DB_TECOZAM porque coje el calendario de España
            dtOperarioCalendarioNoProductivo = ObtieneDiasVacacionesYFestivos(DB_TECOZAM, DB_FERRALLAS, IDOperario, Fecha1, Fecha2, dtCalendario)
            dtDiasInsertar = ObtieneFechasInsertar(DB_FERRALLAS, IDOperario, dtCalendario, dtOperarioCalendarioNoProductivo)

            Windows.Forms.Application.DoEvents()
            LProgreso.Text = "Importando : " & IDOperario & " - FERRALLAS OFICINA"
            Windows.Forms.Application.DoEvents()

            For Each row As DataRow In dtDiasInsertar.Rows

                Dim fecha As Date = row.Field(Of Date)("Fecha")
                IDAutonumerico = auto.Autonumerico()

                'Check de si hay un registro para esta fecha y este operario que pase al siguiente dia.
                'ESTO SE DEBE PORQUE SI ESTA DE BAJA Y TIENE HORAS NO SE GENERAN ADMINISTRATIVAS
                Dim checkSeguir As Boolean
                checkSeguir = getSiInsertarONo(DB_FERRALLAS, fecha, IDOperario)

                If checkSeguir Then
                    Dim rsTrabajo As New DataTable : Dim filtro2 As New Filter
                    filtro2.Add("IDObra", FilterOperator.Equal, IDObra) : filtro2.Add("IdTrabajo", FilterOperator.Equal, IDTrabajo)
                    rsTrabajo = New BE.DataEngine().Filter(DB_FERRALLAS & "..tbObraTrabajo", filtro2)
                    'rsTrabajo = obraTrabajo.Filter(filtro2, , "IdTrabajo, CodTrabajo, DescTrabajo, IdTipoTrabajo, IdSubtipoTrabajo")

                    IDTrabajo = rsTrabajo.Rows(0)("IdTrabajo") : CodTrabajo = rsTrabajo.Rows(0)("CodTrabajo")
                    Dim DescTrabajo As String = "" : Dim IdTipoTrabajo As String = "" : Dim IdSubTipoTrabajo As String = ""
                    DescTrabajo = rsTrabajo.Rows(0)("DescTrabajo") : IdTipoTrabajo = rsTrabajo.Rows(0)("IdTipoTrabajo") : IdSubTipoTrabajo = Nz(rsTrabajo.Rows(0)("IdSubtipotrabajo"), "")
                    Dim DescParte As String : DescParte = "OFICINA" & " " & mes & "-" & año & "-OFI"

                    Dim horas As Double = 8
                    Dim porcentaje As Double
                    Dim dtPorcentaje As DataTable
                    Dim f As New Filter
                    f.Add("IDOperario", FilterOperator.Equal, IDOperario)
                    dtPorcentaje = New BE.DataEngine().Filter(DB_FERRALLAS & "..frmMntoOperario", f)
                    porcentaje = Nz(dtPorcentaje.Rows(0)("JornadaParcial"), 0)
                    If porcentaje <> 0 Then
                        horas = (horas * porcentaje) / 100
                    End If

                    txtSQL = "Insert into " & DB_FERRALLAS & "..tbObraMODControl(IdLineaModControl, IdTrabajo, IdObra, CodTrabajo, DescTrabajo, IdTipoTrabajo, " & _
                            "IdSubTipoTrabajo, IdOperario, IdCategoria, IdHora, FechaInicio, HorasRealMod, TasaRealModA, " & _
                             "ImpRealModA, HorasFactMod, ImpFactModA, DescParte, Facturable, FechaCreacionAudi, FechaModificacionAudi, Usuarioaudi,IDOficio, IdTipoTurno, HorasAdministrativas, IDCategoriaProfesionalSCCP) " & _
                             "Values(" & IDAutonumerico & ", " & IDTrabajo & ", " & IDObra & ", '" & _
                             CodTrabajo & "', '" & DescTrabajo & "', '" & IdTipoTrabajo & "', '" & _
                             IdSubTipoTrabajo & "', '" & IDOperario & "', 'PREDET', '" & _
                             "HA" & "', '" & fecha & "', 0 , " & 0 & ", " & 0 & _
                             ", 0 , " & 0 & _
                             ", '" & DescParte & "', " & 0 & ", '" & Date.Now.Date & "', '" & Date.Now.Date & "', '" & ExpertisApp.UserName & "','" & IDOficio & "', 4,'" & horas.ToString.Replace(",", ".") & "'," & Nz(IDCategoriaProfesionalSCCP, "") & ")"

                    auto.Ejecutar(txtSQL)
                End If

            Next

            filas = filas + 1
            PvProgreso.Value = filas
        Next
        '3. Obtengo una tabla por persona de los días que no tengan que insertar horas
        'MsgBox("Horas de la gente de oficina de Tecozam han sido insertadas correctamente.", vbInformation + vbOKOnly, "STAFF OFICINA")

    End Sub

    Public Sub setHorasOficinaSecozam(ByVal mes As String, ByVal año As String, ByVal Fecha1 As String, ByVal Fecha2 As String)
        '1. Obtengo la tabla de personas que estén en oficina
        Dim dtPersonasOfi As New DataTable
        dtPersonasOfi = getListadoPersonasOfiSecozam(Fecha1, Fecha2)
        '2. Recorro las personas 
        Dim IDOperario As String
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

        Dim filas As Integer = 0
        PvProgreso.Value = 0 : PvProgreso.Maximum = dtPersonasOfi.Rows.Count : PvProgreso.Step = 1 : PvProgreso.Visible = True

        For Each fila As DataRow In dtPersonasOfi.Rows
            IDOperario = fila("IDOperario")
            IDOficio = DevuelveIDOficio(DB_SECOZAM, IDOperario)
            IDCategoriaProfesionalSCCP = DevuelveIDCategoriaProfesionalSCCP(DB_SECOZAM, IDOperario)
            'IDObra = "11854231"
            IDObra = fila("Obra_Predeterminada")
            'Si es distinto que oficina y secozam
            If IDObra <> "11854299" And IDObra <> "11854231" Then
                IDObra = DevuelveUltimoCambioObra(IDOperario, DB_SECOZAM)
            End If

            IDTrabajo = ObtieneIDTrabajo(DB_SECOZAM, IDObra, "PT1")
            'Este es DB_TECOZAM porque coje el calendario de España
            dtOperarioCalendarioNoProductivo = ObtieneDiasVacacionesYFestivos(DB_TECOZAM, DB_SECOZAM, IDOperario, Fecha1, Fecha2, dtCalendario)
            dtDiasInsertar = ObtieneFechasInsertar(DB_SECOZAM, IDOperario, dtCalendario, dtOperarioCalendarioNoProductivo)

            Windows.Forms.Application.DoEvents()
            LProgreso.Text = "Importando : " & IDOperario & " - SECOZAM OFICINA"
            Windows.Forms.Application.DoEvents()

            For Each row As DataRow In dtDiasInsertar.Rows
                Dim fecha As Date = row.Field(Of Date)("Fecha")
                IDAutonumerico = auto.Autonumerico()

                'Check de si hay un registro para esta fecha y este operario que pase al siguiente dia.
                'ESTO SE DEBE PORQUE SI ESTA DE BAJA Y TIENE HORAS NO SE GENERAN ADMINISTRATIVAS
                Dim checkSeguir As Boolean
                checkSeguir = getSiInsertarONo(DB_SECOZAM, fecha, IDOperario)

                If checkSeguir Then
                    Dim rsTrabajo As New DataTable : Dim filtro2 As New Filter
                    filtro2.Add("IDObra", FilterOperator.Equal, IDObra) : filtro2.Add("IdTrabajo", FilterOperator.Equal, IDTrabajo)
                    rsTrabajo = New BE.DataEngine().Filter(DB_SECOZAM & "..tbObraTrabajo", filtro2)
                    'rsTrabajo = obraTrabajo.Filter(filtro2, , "IdTrabajo, CodTrabajo, DescTrabajo, IdTipoTrabajo, IdSubtipoTrabajo")

                    IDTrabajo = rsTrabajo.Rows(0)("IdTrabajo") : CodTrabajo = rsTrabajo.Rows(0)("CodTrabajo")
                    Dim DescTrabajo As String = "" : Dim IdTipoTrabajo As String = "" : Dim IdSubTipoTrabajo As String = ""
                    DescTrabajo = rsTrabajo.Rows(0)("DescTrabajo") : IdTipoTrabajo = rsTrabajo.Rows(0)("IdTipoTrabajo") : IdSubTipoTrabajo = Nz(rsTrabajo.Rows(0)("IdSubtipotrabajo"), "")
                    Dim DescParte As String : DescParte = "OFICINA" & " " & mes & "-" & año & "-OFI"

                    Dim horas As Double = 8
                    Dim porcentaje As Double
                    Dim dtPorcentaje As DataTable
                    Dim f As New Filter
                    f.Add("IDOperario", FilterOperator.Equal, IDOperario)
                    dtPorcentaje = New BE.DataEngine().Filter(DB_SECOZAM & "..frmMntoOperario", f)
                    porcentaje = Nz(dtPorcentaje.Rows(0)("JornadaParcial"), 0)
                    If porcentaje <> 0 Then
                        horas = (horas * porcentaje) / 100
                    End If

                    txtSQL = "Insert into " & DB_SECOZAM & "..tbObraMODControl(IdLineaModControl, IdTrabajo, IdObra, CodTrabajo, DescTrabajo, IdTipoTrabajo, " & _
                            "IdSubTipoTrabajo, IdOperario, IdCategoria, IdHora, FechaInicio, HorasRealMod, TasaRealModA, " & _
                             "ImpRealModA, HorasFactMod, ImpFactModA, DescParte, Facturable, FechaCreacionAudi, FechaModificacionAudi, Usuarioaudi,IDOficio, IdTipoTurno, HorasAdministrativas, IDCategoriaProfesionalSCCP) " & _
                             "Values(" & IDAutonumerico & ", " & IDTrabajo & ", " & IDObra & ", '" & _
                             CodTrabajo & "', '" & DescTrabajo & "', '" & IdTipoTrabajo & "', '" & _
                             IdSubTipoTrabajo & "', '" & IDOperario & "', 'PREDET', '" & _
                             "HA" & "', '" & fecha & "', 0 , " & 0 & ", " & 0 & _
                             ", 0 , " & 0 & _
                             ", '" & DescParte & "', " & 0 & ", '" & Date.Now.Date & "', '" & Date.Now.Date & "', '" & ExpertisApp.UserName & "','" & IDOficio & "', 4," & horas & " ," & Nz(IDCategoriaProfesionalSCCP, "") & ")"

                    auto.Ejecutar(txtSQL)
                End If

            Next

            filas = filas + 1
            PvProgreso.Value = filas
        Next
        '3. Obtengo una tabla por persona de los días que no tengan que insertar horas
        MsgBox("Las horas de la gente de oficina de Tecozam, Ferrallas y Secozam han sido insertadas correctamente.", vbInformation + vbOKOnly, "STAFF OFICINA")

    End Sub

    Public Sub setHorasOficinaDCZ(ByVal mes As String, ByVal año As String, ByVal Fecha1 As String, ByVal Fecha2 As String)
        '1. Obtengo la tabla de personas que estén en oficina
        Dim dtPersonasOfi As New DataTable
        dtPersonasOfi = getListadoPersonasOfiDCZ(Fecha1, Fecha2)
        '2. Recorro las personas 
        Dim IDOperario As String
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

        Dim filas As Integer = 0
        PvProgreso.Value = 0 : PvProgreso.Maximum = dtPersonasOfi.Rows.Count : PvProgreso.Step = 1 : PvProgreso.Visible = True

        For Each fila As DataRow In dtPersonasOfi.Rows
            IDOperario = fila("IDOperario")
            IDOficio = DevuelveIDOficio(DB_DCZ, IDOperario)
            IDCategoriaProfesionalSCCP = DevuelveIDCategoriaProfesionalSCCP(DB_DCZ, IDOperario)
            IDObra = "11860026"

            If IDObra <> "11860026" Then
                IDObra = DevuelveUltimoCambioObra(IDOperario, DB_DCZ)
            End If

            IDTrabajo = ObtieneIDTrabajo(DB_DCZ, IDObra, "PT1")
            'Este es DB_TECOZAM porque coje el calendario de España
            dtOperarioCalendarioNoProductivo = ObtieneDiasVacacionesYFestivos(DB_TECOZAM, DB_DCZ, IDOperario, Fecha1, Fecha2, dtCalendario)
            dtDiasInsertar = ObtieneFechasInsertar(DB_DCZ, IDOperario, dtCalendario, dtOperarioCalendarioNoProductivo)

            Windows.Forms.Application.DoEvents()
            LProgreso.Text = "Importando : " & IDOperario & " - DCZ OFICINA"
            Windows.Forms.Application.DoEvents()

            For Each row As DataRow In dtDiasInsertar.Rows
                Dim fecha As Date = row.Field(Of Date)("Fecha")
                IDAutonumerico = auto.Autonumerico()

                Dim rsTrabajo As New DataTable : Dim filtro2 As New Filter
                filtro2.Add("IDObra", FilterOperator.Equal, IDObra) : filtro2.Add("IdTrabajo", FilterOperator.Equal, IDTrabajo)
                rsTrabajo = New BE.DataEngine().Filter(DB_DCZ & "..tbObraTrabajo", filtro2)
                'rsTrabajo = obraTrabajo.Filter(filtro2, , "IdTrabajo, CodTrabajo, DescTrabajo, IdTipoTrabajo, IdSubtipoTrabajo")

                IDTrabajo = rsTrabajo.Rows(0)("IdTrabajo") : CodTrabajo = rsTrabajo.Rows(0)("CodTrabajo")
                Dim DescTrabajo As String = "" : Dim IdTipoTrabajo As String = "" : Dim IdSubTipoTrabajo As String = ""
                DescTrabajo = rsTrabajo.Rows(0)("DescTrabajo") : IdTipoTrabajo = rsTrabajo.Rows(0)("IdTipoTrabajo") : IdSubTipoTrabajo = Nz(rsTrabajo.Rows(0)("IdSubtipotrabajo"), "")
                Dim DescParte As String : DescParte = "OFICINA" & " " & mes & "-" & año & "-OFI"

                txtSQL = "Insert into " & DB_DCZ & "..tbObraMODControl(IdLineaModControl, IdTrabajo, IdObra, CodTrabajo, DescTrabajo, IdTipoTrabajo, " & _
                        "IdSubTipoTrabajo, IdOperario, IdCategoria, IdHora, FechaInicio, HorasRealMod, TasaRealModA, " & _
                         "ImpRealModA, HorasFactMod, ImpFactModA, DescParte, Facturable, FechaCreacionAudi, FechaModificacionAudi, Usuarioaudi,IDOficio, IdTipoTurno, HorasAdministrativas, IDCategoriaProfesionalSCCP) " & _
                         "Values(" & IDAutonumerico & ", " & IDTrabajo & ", " & IDObra & ", '" & _
                         CodTrabajo & "', '" & DescTrabajo & "', '" & IdTipoTrabajo & "', '" & _
                         IdSubTipoTrabajo & "', '" & IDOperario & "', 'PREDET', '" & _
                         "HA" & "', '" & fecha & "', 0 , " & 0 & ", " & 0 & _
                         ", 0 , " & 0 & _
                         ", '" & DescParte & "', " & 0 & ", '" & Date.Now.Date & "', '" & Date.Now.Date & "', '" & ExpertisApp.UserName & "','" & IDOficio & "', 4, 8 ," & Nz(IDCategoriaProfesionalSCCP, "") & ")"

                auto.Ejecutar(txtSQL)
            Next

            filas = filas + 1
            PvProgreso.Value = filas
        Next
        '3. Obtengo una tabla por persona de los días que no tengan que insertar horas
        MsgBox("Las horas de la gente de oficina han sido insertadas correctamente.", vbInformation + vbOKOnly, "STAFF OFICINA")

    End Sub

    Public Function ObtieneDiaUltimoMes(ByVal mes As String, ByVal anio As String) As String
        Dim connectionString As String = "Data Source=stecodesarr;Initial Catalog=" & DB_TECOZAM & ";User ID=sa;Password=180M296;"
        Dim connection As New SqlConnection(connectionString)
        connection.Open()

        Dim queryString As String = "SELECT COUNT(*) AS Dias FROM tbMaestroFechas where Month(Fecha)='" & mes & "' and YEAR(Fecha)='" & anio & "'"
        Dim command As New SqlCommand(queryString, connection)

        Dim adapter As New SqlDataAdapter(command)
        Dim dt As New DataTable()
        adapter.Fill(dt)

        connection.Close()
        Return dt.Rows(0)("Dias")

    End Function

    Private Sub bNota_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MsgBox("JP Y STAFF: TECOZAM-DCK-UK " & vbCrLf _
               & "OFICINA: ESPAÑA " & vbCrLf _
               & "BAJA: ESPAÑA", MsgBoxStyle.OkOnly, "Ayuda")

        'Dim IDObra As String
        'IDObra = DevuelveIDObra(DB_UK, "Tuk08")
    End Sub

    Private Sub bAñadirHorasPersona_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bAñadirHorasPersona.Click
        Dim frmCrea As New frmCreaHorasOperarioObraFecha
        frmCrea.ShowDialog()

    End Sub

    Private Sub bBorrarOperarioObraFecha_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bBorrarOperarioObraFecha.Click
        Dim frmBorra As New frmBorraHoras
        frmBorra.ShowDialog()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        CD.Filter = "Archivos Excel(*.xls;*.xlsx)|*.xls;*xlsx|Todos los archivos(*.*)|*.*"

        'CD.ShowOpen()
        CD.ShowDialog()

        If CD.FileName <> "" Then
            'lblRuta.Caption = CD.FileName
            lblRuta.Text = CD.FileName
        End If
    End Sub

    Sub DeshacerTraspaso(ByVal sNombreGlobal)
        'Dim clsAdmin As New AdminEjecutor

        If sNombreGlobal <> "" Then
            'clsAdmin.Ejecutar("Delete from tbObraMODControl where DescParte like '" & sNombreGlobal & "'", False)
            AdminData.Execute("Delete from tbObraMODControl where DescParte like '" & sNombreGlobal & "'")
        End If

        'Libero memoria
        'clsAdmin = Nothing

    End Sub

    Private Sub bCreaHoras_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bCreaHoras.Click
        'importarExcelPorEmpresa()
        '----------NUEVO METODO CON EXCEL NUEVO -----------
        '1. EN LA HOJA "DATOS ARCHIVO" CELDA A1 ESTA EL NOBRA
        '2. EN LA HOJA "HORAS" ESTÁN TODAS LAS PERSONAS POR IDOPERARIO Y SUS HORAS
        Dim ruta As String = lblRuta.Text
        Dim hoja1 As String = "DATOS DE ARCHIVO"
        Dim hoja2 As String = "HORAS"
        Dim rango1 As String = "A1:B1"
        Dim rango2 As String = "A5:BR300"

        Dim dtObra As New DataTable
        Dim dtHoras As New DataTable
        Dim dtFechas As New DataTable
        dtObra = ObtenerDatosExcel(ruta, hoja1, rango1)
        dtHoras = ObtenerDatosExcel(ruta, hoja2, rango2)
        dtFechas = ObtenerDatosExcel(ruta, hoja1, "A1:B3")

        Dim basededatos As String
        Dim obra As String
        Dim fecha1 As String
        Dim fecha2 As String

        fecha1 = dtFechas.Rows(1)("F2").ToString.Trim()
        fecha2 = dtFechas.Rows(2)("F2").ToString.Trim()
        obra = dtObra.Rows(0)("F2").ToString.Trim()
        basededatos = DevuelveBaseDeDatosInternacional(obra)

        If basededatos = "0" Then
            Exit Sub
        End If

        dtHoras = dtFormaInternacional(dtHoras, fecha1)

        '-------CHECK IDOPERARIO
        Dim idoperario As String
        For Each fila As DataRow In dtHoras.Rows
            idoperario = fila("IDOperario").ToString
            If idoperario.Length = 0 Then
                Continue For
            End If
            Dim dt As New DataTable
            Dim f As New Filter : f.Add("IDOperario", FilterOperator.Equal, idoperario)
            dt = New BE.DataEngine().Filter("vUnionOperariosCategoriaProfesional", f)

            If dt.Rows.Count = 0 Then
                MsgBox("El operario " & idoperario & " no existe.")
                Exit Sub
            End If

            If dt.Rows(0)("CategoriaProfesionalSCCP") Is Nothing OrElse IsDBNull(dt.Rows(0)("CategoriaProfesionalSCCP")) Then
                MsgBox("El operario " & idoperario & " no tiene oficio asignado.")
                Exit Sub
            End If

        Next
        '----------------------
        Dim filas As Integer = 0
        PvProgreso.Value = 0 : PvProgreso.Maximum = dtHoras.Rows.Count : PvProgreso.Step = 1 : PvProgreso.Visible = True
        For Each dr As DataRow In dtHoras.Rows
            idoperario = dr("IDOperario").ToString
            If idoperario.Length = 0 Then
                filas = filas + 1
                Continue For
            End If
            Windows.Forms.Application.DoEvents()
            LProgreso.Text = "Importando : " & idoperario
            Windows.Forms.Application.DoEvents()

            ' Recorrer las columnas a partir de la tercera
            For i As Integer = 3 To dtHoras.Columns.Count - 1
                Dim value As Object = dr(i)
                Dim fecha As String
                ' Verificar si la cabecera de la columna es numérica
                If IsNumeric(dtHoras.Columns(i).ColumnName) Then
                    fecha = DevuelveFechaDeNumero(dtHoras.Columns(i).ColumnName)
                Else
                    fecha = DevuelveFechaConFormato(dtHoras.Columns(i).ColumnName)
                End If

                InsertaHorasBaseDeDatos(basededatos, obra, idoperario, fecha, value)
            Next
            filas = filas + 1
            PvProgreso.Value = filas
        Next

        MsgBox("Proceso finalizado correctamente", MsgBoxStyle.Information)
    End Sub
    Public Sub InsertaHorasBaseDeDatos(ByVal basededatos As String, ByVal obra As String, ByVal idoperario As String, ByVal fecha As String, ByVal value As Object)
        'CHECK SI EL OPERARIO ES CATEGORIA 2 O 3 ENTONCES INSERTA HORAS
        Dim IDCategoriaProfesionalSCCP As String
        IDCategoriaProfesionalSCCP = DevuelveIDCategoriaProfesionalSCCP(basededatos, idoperario)
        'CHECK SI existe registro de esta persona este dia en la base de datos
        Dim dtCheckRegistro As DataTable
        Dim f As New Filter
        Dim idobra As String
        f.Add("FechaInicio", FilterOperator.Equal, fecha)
        f.Add("IDOperario", FilterOperator.Equal, idoperario)
        idobra = DevuelveIDObra(basededatos, obra)
        f.Add("IDObra", FilterOperator.Equal, idobra)
        dtCheckRegistro = New BE.DataEngine().Filter(basededatos & "..tbObraModControl", f)
        Dim txtSQL As String
        Dim idoficio As String
        Dim IDTrabajo As String
        Dim CodTrabajo As String
        Dim IDAutonumerico As String

        If dtCheckRegistro.Rows.Count = 0 Then
            If IDCategoriaProfesionalSCCP = 2 Or IDCategoriaProfesionalSCCP = 3 Then
                'CHECK SI value es numerico va a HorasRealMod
                ' SI value es ACC o CC inserta 8 en horas baja.
                If IsNumeric(value) Then
                    Dim horas As Double = value
                    IDOficio = DevuelveIDOficio(basededatos, idoperario)
                    'idobra = DevuelveIDObra(basededatos, obra)
                    IDTrabajo = ObtieneIDTrabajo(basededatos, idobra, "PT1")
                    IDAutonumerico = auto.Autonumerico()

                    Dim rsTrabajo As New DataTable
                    Dim filtro2 As New Filter
                    filtro2.Add("IDObra", FilterOperator.Equal, idobra)
                    filtro2.Add("IdTrabajo", FilterOperator.Equal, IDTrabajo)

                    rsTrabajo = New BE.DataEngine().Filter(basededatos & "..tbObraTrabajo", filtro2)
                    'rsTrabajo = obraTrabajo.Filter(filtro2, , "IdTrabajo, CodTrabajo, DescTrabajo, IdTipoTrabajo, IdSubtipoTrabajo")
                    IDTrabajo = rsTrabajo.Rows(0)("IdTrabajo")
                    CodTrabajo = rsTrabajo.Rows(0)("CodTrabajo")
                    Dim DescTrabajo As String = "" : Dim IdTipoTrabajo As String = "" : Dim IdSubTipoTrabajo As String = ""
                    DescTrabajo = rsTrabajo.Rows(0)("DescTrabajo") : IdTipoTrabajo = rsTrabajo.Rows(0)("IdTipoTrabajo") : IdSubTipoTrabajo = Nz(rsTrabajo.Rows(0)("IdSubtipotrabajo"), "")


                    Dim DescParte As String : DescParte = "INTERNACIONAL PRODUCTIVAS"
                    txtSQL = "Insert into " & basededatos & "..tbObraMODControl(IdLineaModControl, IdTrabajo, IdObra, CodTrabajo, DescTrabajo, IdTipoTrabajo, " & _
                            "IdSubTipoTrabajo, IdOperario, IdCategoria, IdHora, FechaInicio, HorasRealMod, TasaRealModA, " & _
                             "ImpRealModA, HorasFactMod, ImpFactModA, DescParte, Facturable, FechaCreacionAudi, FechaModificacionAudi, Usuarioaudi, IDOficio, IdTipoTurno, HorasAdministrativas, IDCategoriaProfesionalSCCP) " & _
                             "Values(" & IDAutonumerico & ", " & IDTrabajo & ", " & idobra & ", '" & _
                             CodTrabajo & "', '" & DescTrabajo & "', '" & IdTipoTrabajo & "', '" & _
                             IdSubTipoTrabajo & "', '" & idoperario & "', 'PREDET', '" & _
                             "HO" & "', '" & fecha & "',  " & horas.ToString.Replace(",", ".") & "  , " & 0 & ", " & 0 & _
                             ", 0 , " & 0 & _
                             ", '" & DescParte & "', " & 0 & ", '" & Date.Now.Date & "', '" & Date.Now.Date & "', '" & ExpertisApp.UserName & "','" & idoficio & "', 4,0 ," & Nz(IDCategoriaProfesionalSCCP, "") & ")"

                    auto.Ejecutar(txtSQL)

                ElseIf value IsNot Nothing AndAlso TypeOf value Is String Then
                    If value = "ACC" Or value.ToString = "CC" Or value.ToString = "acc" Or value.ToString = "cc" Or value.ToString = "SSP" Or value.ToString = "B" Then
                        idoficio = DevuelveIDOficio(basededatos, idoperario)
                        idobra = DevuelveIDObra(basededatos, obra)
                        IDTrabajo = ObtieneIDTrabajo(basededatos, idobra, "PT1")
                        IDAutonumerico = auto.Autonumerico()

                        Dim rsTrabajo As New DataTable
                        Dim filtro2 As New Filter
                        filtro2.Add("IDObra", FilterOperator.Equal, idobra)
                        filtro2.Add("IdTrabajo", FilterOperator.Equal, IDTrabajo)

                        rsTrabajo = New BE.DataEngine().Filter(basededatos & "..tbObraTrabajo", filtro2)
                        'rsTrabajo = obraTrabajo.Filter(filtro2, , "IdTrabajo, CodTrabajo, DescTrabajo, IdTipoTrabajo, IdSubtipoTrabajo")
                        IDTrabajo = rsTrabajo.Rows(0)("IdTrabajo")
                        CodTrabajo = rsTrabajo.Rows(0)("CodTrabajo")
                        Dim DescTrabajo As String = "" : Dim IdTipoTrabajo As String = "" : Dim IdSubTipoTrabajo As String = ""
                        DescTrabajo = rsTrabajo.Rows(0)("DescTrabajo") : IdTipoTrabajo = rsTrabajo.Rows(0)("IdTipoTrabajo") : IdSubTipoTrabajo = Nz(rsTrabajo.Rows(0)("IdSubtipotrabajo"), "")


                        Dim DescParte As String : DescParte = "INTERNACIONAL BAJA"
                        txtSQL = "Insert into " & basededatos & "..tbObraMODControl(IdLineaModControl, IdTrabajo, IdObra, CodTrabajo, DescTrabajo, IdTipoTrabajo, " & _
                                "IdSubTipoTrabajo, IdOperario, IdCategoria, IdHora, FechaInicio, HorasRealMod, TasaRealModA, " & _
                                 "ImpRealModA, HorasFactMod, ImpFactModA, DescParte, Facturable, FechaCreacionAudi, FechaModificacionAudi, Usuarioaudi, IDOficio, IdTipoTurno, HorasBaja, IDCategoriaProfesionalSCCP) " & _
                                 "Values(" & IDAutonumerico & ", " & IDTrabajo & ", " & idobra & ", '" & _
                                 CodTrabajo & "', '" & DescTrabajo & "', '" & IdTipoTrabajo & "', '" & _
                                 IdSubTipoTrabajo & "', '" & idoperario & "', 'PREDET', '" & _
                                 value & "', '" & fecha & "',  0 , " & 0 & ", " & 0 & _
                                 ", 0 , " & 0 & _
                                 ", '" & DescParte & "', " & 0 & ", '" & Date.Now.Date & "', '" & Date.Now.Date & "', '" & ExpertisApp.UserName & "','" & idoficio & "', 4,8," & Nz(IDCategoriaProfesionalSCCP, "") & ")"

                        auto.Ejecutar(txtSQL)

                    End If
                End If
            End If
        Else
            Dim IDLineaModControl As String
            IDLineaModControl = dtCheckRegistro(0)("IDLineaModControl").ToString

            BorraRegistro(basededatos, IDLineaModControl)
            If IDCategoriaProfesionalSCCP = 2 Or IDCategoriaProfesionalSCCP = 3 Then
                'CHECK SI value es numerico va a HorasRealMod
                ' SI value es ACC o CC inserta 8 en horas baja.
                If IsNumeric(value) Then
                    Dim horas As Double = value
                    idoficio = DevuelveIDOficio(basededatos, idoperario)
                    'idobra = DevuelveIDObra(basededatos, obra)
                    IDTrabajo = ObtieneIDTrabajo(basededatos, idobra, "PT1")
                    IDAutonumerico = auto.Autonumerico()

                    Dim rsTrabajo As New DataTable
                    Dim filtro2 As New Filter
                    filtro2.Add("IDObra", FilterOperator.Equal, idobra)
                    filtro2.Add("IdTrabajo", FilterOperator.Equal, IDTrabajo)

                    rsTrabajo = New BE.DataEngine().Filter(basededatos & "..tbObraTrabajo", filtro2)
                    'rsTrabajo = obraTrabajo.Filter(filtro2, , "IdTrabajo, CodTrabajo, DescTrabajo, IdTipoTrabajo, IdSubtipoTrabajo")
                    IDTrabajo = rsTrabajo.Rows(0)("IdTrabajo")
                    CodTrabajo = rsTrabajo.Rows(0)("CodTrabajo")
                    Dim DescTrabajo As String = "" : Dim IdTipoTrabajo As String = "" : Dim IdSubTipoTrabajo As String = ""
                    DescTrabajo = rsTrabajo.Rows(0)("DescTrabajo") : IdTipoTrabajo = rsTrabajo.Rows(0)("IdTipoTrabajo") : IdSubTipoTrabajo = Nz(rsTrabajo.Rows(0)("IdSubtipotrabajo"), "")


                    Dim DescParte As String : DescParte = "INTERNACIONAL PRODUCTIVAS"
                    txtSQL = "Insert into " & basededatos & "..tbObraMODControl(IdLineaModControl, IdTrabajo, IdObra, CodTrabajo, DescTrabajo, IdTipoTrabajo, " & _
                            "IdSubTipoTrabajo, IdOperario, IdCategoria, IdHora, FechaInicio, HorasRealMod, TasaRealModA, " & _
                             "ImpRealModA, HorasFactMod, ImpFactModA, DescParte, Facturable, FechaCreacionAudi, FechaModificacionAudi, Usuarioaudi, IDOficio, IdTipoTurno, HorasAdministrativas, IDCategoriaProfesionalSCCP) " & _
                             "Values(" & IDAutonumerico & ", " & IDTrabajo & ", " & idobra & ", '" & _
                             CodTrabajo & "', '" & DescTrabajo & "', '" & IdTipoTrabajo & "', '" & _
                             IdSubTipoTrabajo & "', '" & idoperario & "', 'PREDET', '" & _
                             "HO" & "', '" & fecha & "',  " & horas.ToString.Replace(",", ".") & "  , " & 0 & ", " & 0 & _
                             ", 0 , " & 0 & _
                             ", '" & DescParte & "', " & 0 & ", '" & Date.Now.Date & "', '" & Date.Now.Date & "', '" & ExpertisApp.UserName & "','" & idoficio & "', 4,0 ," & Nz(IDCategoriaProfesionalSCCP, "") & ")"

                    auto.Ejecutar(txtSQL)

                ElseIf value IsNot Nothing AndAlso TypeOf value Is String Then
                    If value = "ACC" Or value.ToString = "CC" Or value.ToString = "acc" Or value.ToString = "cc" Or value.ToString = "SSP" Or value.ToString = "B" Then
                        idoficio = DevuelveIDOficio(basededatos, idoperario)
                        idobra = DevuelveIDObra(basededatos, obra)
                        IDTrabajo = ObtieneIDTrabajo(basededatos, idobra, "PT1")
                        IDAutonumerico = auto.Autonumerico()

                        Dim rsTrabajo As New DataTable
                        Dim filtro2 As New Filter
                        filtro2.Add("IDObra", FilterOperator.Equal, idobra)
                        filtro2.Add("IdTrabajo", FilterOperator.Equal, IDTrabajo)

                        rsTrabajo = New BE.DataEngine().Filter(basededatos & "..tbObraTrabajo", filtro2)
                        'rsTrabajo = obraTrabajo.Filter(filtro2, , "IdTrabajo, CodTrabajo, DescTrabajo, IdTipoTrabajo, IdSubtipoTrabajo")
                        IDTrabajo = rsTrabajo.Rows(0)("IdTrabajo")
                        CodTrabajo = rsTrabajo.Rows(0)("CodTrabajo")
                        Dim DescTrabajo As String = "" : Dim IdTipoTrabajo As String = "" : Dim IdSubTipoTrabajo As String = ""
                        DescTrabajo = rsTrabajo.Rows(0)("DescTrabajo") : IdTipoTrabajo = rsTrabajo.Rows(0)("IdTipoTrabajo") : IdSubTipoTrabajo = Nz(rsTrabajo.Rows(0)("IdSubtipotrabajo"), "")


                        Dim DescParte As String : DescParte = "INTERNACIONAL BAJA"
                        txtSQL = "Insert into " & basededatos & "..tbObraMODControl(IdLineaModControl, IdTrabajo, IdObra, CodTrabajo, DescTrabajo, IdTipoTrabajo, " & _
                                "IdSubTipoTrabajo, IdOperario, IdCategoria, IdHora, FechaInicio, HorasRealMod, TasaRealModA, " & _
                                 "ImpRealModA, HorasFactMod, ImpFactModA, DescParte, Facturable, FechaCreacionAudi, FechaModificacionAudi, Usuarioaudi, IDOficio, IdTipoTurno, HorasBaja, IDCategoriaProfesionalSCCP) " & _
                                 "Values(" & IDAutonumerico & ", " & IDTrabajo & ", " & idobra & ", '" & _
                                 CodTrabajo & "', '" & DescTrabajo & "', '" & IdTipoTrabajo & "', '" & _
                                 IdSubTipoTrabajo & "', '" & idoperario & "', 'PREDET', '" & _
                                 value & "', '" & fecha & "',  0 , " & 0 & ", " & 0 & _
                                 ", 0 , " & 0 & _
                                 ", '" & DescParte & "', " & 0 & ", '" & Date.Now.Date & "', '" & Date.Now.Date & "', '" & ExpertisApp.UserName & "','" & idoficio & "', 4,8," & Nz(IDCategoriaProfesionalSCCP, "") & ")"

                        auto.Ejecutar(txtSQL)

                    End If
                End If
            End If
        End If
    End Sub

    Public Sub BorraRegistro(ByVal basededatos As String, ByVal IDLineaModCOntrol As String)
        Dim sql As String = "Delete from " & basededatos & "..tbObraMODControl where IDLineaModControl = '" & IDLineaModCOntrol & "'"
        aux.EjecutarSql(sql)
    End Sub
    Public Function DevuelveFechaDeNumero(ByVal fechanumero As String) As String
        Dim fechaString As String = fechanumero.ToString()
        Dim fechaBase As New DateTime(1899, 12, 30)
        Dim fechaFinal As DateTime = fechaBase.AddDays(fechaString)
        ' Extraer día, mes y año
        Dim dia As Integer = fechaFinal.Day
        Dim mes As Integer = fechaFinal.Month
        Dim anio As Integer = fechaFinal.Year

        Return fechaFinal
    End Function

    Public Function DevuelveFechaConFormato(ByVal fechanumero As String) As String

        Dim fechaActual As DateTime = fechanumero
        Dim formatoPersonalizado As String = "dd/MM/yyyy"
        Dim fechaComoStringConFormato As String = fechaActual.ToString(formatoPersonalizado)

        Return fechaComoStringConFormato

    End Function

    Public Function DevuelveBaseDeDatosInternacional(ByVal obra As String) As String
        Dim basededatos As String
        If obra.StartsWith("TUK") Then
            basededatos = "xTecozamUnitedKingdom50R2"
        ElseIf obra.StartsWith("DP") Then
            basededatos = "xDrenajesPortugal50R2"
        ElseIf obra.StartsWith("TN") Then
            basededatos = "xTecozamNorge50R2"
        Else
            MsgBox("No se encuentra base de datos para esta obra")
            basededatos = "0"
        End If

        Return basededatos
    End Function

    Public Function dtFormaInternacional(ByRef dtHoras As DataTable, ByVal fecha1 As String) As DataTable
        dtHoras.Columns.Remove("F4")
        dtHoras.Columns.Remove("F5")
        dtHoras.Columns.Remove("F6")
        dtHoras.Columns.Remove("F7")
        dtHoras.Columns.Remove("F8")
        dtHoras.Columns.Remove("F10")
        dtHoras.Columns.Remove("F12")
        dtHoras.Columns.Remove("F14")
        dtHoras.Columns.Remove("F16")
        dtHoras.Columns.Remove("F18")
        dtHoras.Columns.Remove("F20")
        dtHoras.Columns.Remove("F22")
        dtHoras.Columns.Remove("F24")
        dtHoras.Columns.Remove("F26")
        dtHoras.Columns.Remove("F28")
        dtHoras.Columns.Remove("F30")
        dtHoras.Columns.Remove("F32")
        dtHoras.Columns.Remove("F34")
        dtHoras.Columns.Remove("F36")
        dtHoras.Columns.Remove("F38")
        dtHoras.Columns.Remove("F40")
        dtHoras.Columns.Remove("F42")
        dtHoras.Columns.Remove("F44")
        dtHoras.Columns.Remove("F46")
        dtHoras.Columns.Remove("F48")
        dtHoras.Columns.Remove("F50")
        dtHoras.Columns.Remove("F52")
        dtHoras.Columns.Remove("F54")
        dtHoras.Columns.Remove("F56")
        dtHoras.Columns.Remove("F58")
        dtHoras.Columns.Remove("F60")
        dtHoras.Columns.Remove("F62")
        dtHoras.Columns.Remove("F64")
        dtHoras.Columns.Remove("F66")
        dtHoras.Columns.Remove("F68")
        dtHoras.Columns.Remove("F70")

        ' Crear un nuevo DataTable para almacenar el resultado con la primera fila como cabecera
        Dim dtFinal As New DataTable()

        ' Añadir columnas al DataTable con base en la primera fila de dtHoras
        For Each columnaOriginal As DataColumn In dtHoras.Columns
            ' Obtener el nombre de la columna original y agregarla como columna al nuevo DataTable
            dtFinal.Columns.Add(dtHoras.Rows(0)(columnaOriginal.ColumnName).ToString(), columnaOriginal.DataType)
        Next

        ' Añadir filas al DataTable con base en las filas restantes de dtHoras
        For i As Integer = 1 To dtHoras.Rows.Count - 1
            Dim nuevaFila As DataRow = dtFinal.NewRow()

            If dtHoras.Rows(i)(0).ToString() = "Fin" Or dtHoras.Rows(i)(0).ToString() = "FIN" Or dtHoras.Rows(i)(0).ToString() = "TOTAL" Or dtHoras.Rows(i)(0).ToString() = "Total" Then
                Exit For
            End If
            ' Copiar los datos de las filas restantes a las filas del nuevo DataTable
            For j As Integer = 0 To dtHoras.Columns.Count - 1
                nuevaFila(j) = dtHoras.Rows(i)(j)
            Next

            ' Añadir la nueva fila al DataTable con cabecera
            dtFinal.Rows.Add(nuevaFila)
        Next

        '--------------BORRO LAS QUE EMPIEZAN POR COLUMN QUE SE COMPLETAN SOLAS
        ' Supongamos que tienes un DataTable llamado dtFinal
        Dim columnasAEliminar As New List(Of DataColumn)

        ' Identificar las columnas a eliminar
        For Each columna As DataColumn In dtFinal.Columns
            If columna.ColumnName.StartsWith("Colum", StringComparison.OrdinalIgnoreCase) Then
                columnasAEliminar.Add(columna)
            End If
        Next

        ' Eliminar las columnas identificadas
        For Each columnaAEliminar As DataColumn In columnasAEliminar
            dtFinal.Columns.Remove(columnaAEliminar)
        Next

        'A partir de la tercera columna y solo para la cabecera pongo el format de dd/mm/yyyy
        'David V 12/12/23
        Dim fecha As DateTime = fecha1
        fecha = fecha.AddDays(-1)
        For i As Integer = 3 To dtFinal.Columns.Count - 1
            If TypeOf dtFinal.Columns(i) Is DataColumn Then
                Dim fechaColumna As DataColumn = CType(dtFinal.Columns(i), DataColumn)
                ' Formatea la fecha y establece el nuevo formato en el nombre de la columna
                fecha = fecha.AddDays(1)
                fechaColumna.ColumnName = fecha
            End If
        Next

        Return dtFinal
    End Function
    Public Sub importarExcelPorEmpresa()
        Dim obraCab As New ObraCabecera

        Dim columna As Integer
        Dim ruta As String = lblRuta.Text
        Dim hoja As String = "Horas"
        Dim rango1 As String = "B1:B10"
        Dim rango2 As String = "A12:AG200"
        Dim rango3 As String = "A11:AG11"

        Dim empresa As String
        Dim estado As String
        Dim obra As String
        Dim trabajo As String
        Dim mes As String
        Dim numero As String
        Dim iRegistros As Integer
        Dim fecha As String
        Dim idOperario As String
        Dim basededatos1 As String

        Dim hora As Double
        Dim tipoHora As String

        Dim sNombreUnicoGlobal As String
        Dim iSQL As String
        Dim sSQL As String

        Dim rsnobra As New DataTable
        Dim rs As New DataTable
        Dim dtHoras As New DataTable
        Dim dtDatos As New DataTable
        Dim dtFecha As New DataTable

        Dim f As New Filter
        dtDatos = ObtenerDatosExcel(ruta, hoja, rango1)
        dtHoras = ObtenerDatosExcel(ruta, hoja, rango2)
        dtFecha = ObtenerDatosExcel(ruta, hoja, rango3)


        empresa = dtDatos.Rows(0)(0)
        estado = dtDatos.Rows(1)(0)
        obra = dtDatos.Rows(2)(0)
        trabajo = dtDatos.Rows(3)(0)
        '01/06/23
        basededatos1 = dtDatos.Rows(8)(0)
        mes = dtDatos.Rows(9)(0)

        Dim bbdd As String
        If Len(basededatos1) = 0 Or basededatos1 = "TECOZAM" Then
            bbdd = DB_TECOZAM
        ElseIf basededatos1 = "FERRALLAS" Then
            bbdd = DB_FERRALLAS
        ElseIf basededatos1 = "DCZ" Then
            bbdd = DB_DCZ
        ElseIf basededatos1 = "UK" Then
            bbdd = DB_UK
        ElseIf basededatos1 = "SUECIA" Then
            bbdd = DB_SU
        ElseIf basededatos1 = "NORUEGA" Then
            bbdd = DB_NO
        Else
            MsgBox("No coincide la empresa con ninguna base de datos habilitada.")
            Exit Sub
        End If

        Dim result As DialogResult = MessageBox.Show("¿Deseas aceptar el proceso de insertar horas en " & basededatos1 & " ?", "Confirmación datos", MessageBoxButtons.YesNo)
        If result = DialogResult.Yes Then
            'iRegistros = dtHoras.Rows.Count
            'David Velasco 09/11
            'Recorro la tabla y en cuanto haya un Codigo de Operario vacio deja de leer.
            Dim cont As Integer = 0
            For Each dr As DataRow In dtHoras.Rows
                If IsDBNull(dr(cont)) Then
                    'MsgBox("El Excel tiene " & cont & " filas.")
                    Exit For
                End If
                cont += 1
            Next

            iRegistros = dtHoras.Rows.Count - 1
            sNombreUnicoGlobal = obra & " " & mes

            If estado <> "REVISADO" Then
                MsgBox("El estado del archivo es: " & estado & ". Para Importar debe ser 'Revisado'. El proceso se cancelara", vbExclamation + vbOKOnly)

                rs = Nothing
                DeshacerTraspaso(sNombreUnicoGlobal)
                If Err.Description <> "" Then
                    MsgBox("Proceso cancelado. Error: '" & Err.Description & "'", vbCritical + vbOKOnly)
                End If
            End If

            f.Clear()
            f.Add("NObra", FilterOperator.Equal, obra)
            iSQL = "Nobra= '" & obra & "'"


            rsnobra = New BE.DataEngine().Filter(bbdd & "..tbObraCabecera", f)
            'numero = DevuelveIDObra(bbdd, obra)
            'rsnobra = obraCab.Filter(f, , "IDObra")
            numero = rsnobra(0)("IDObra")

            f.Clear()
            f.Add("IDObra", FilterOperator.Equal, numero)
            f.Add("CodTrabajo", FilterOperator.Equal, trabajo)

            sSQL = "IdObra=" & numero & " and Codtrabajo='" & trabajo & "'"

            Dim obraTrabajo As New ObraTrabajo

            rs = New BE.DataEngine().Filter(bbdd & "..tbObraTrabajo", f)
            'rs = obraTrabajo.Filter(f)

            Dim idtrab As String
            idtrab = rs(0)("IDTrabajo").ToString

            If rs.Rows.Count > 2 Then
                MsgBox("Ya hay datos insertados para este parte. Se cancela la importacion", vbCritical + vbOKOnly)
                sNombreUnicoGlobal = ""
                rs = Nothing
                DeshacerTraspaso(sNombreUnicoGlobal)
                If Err.Description <> "" Then
                    MsgBox("Proceso cancelado. Error: '" & Err.Description & "'", vbCritical + vbOKOnly)
                End If

                Exit Sub

            End If

            PvProgreso.Value = 0
            PvProgreso.Maximum = dtFecha.Columns.Count - 1
            PvProgreso.Step = 1
            PvProgreso.Visible = True

            columna = 2
            'Dim cuenta As Integer = 1
            'RECORRE LAS COLUMNAS HASTA AG
            While columna < dtFecha.Columns.Count

                'MessageBox.Show("fecha: " & contador)
                Try
                    fecha = dtFecha(0)(columna)
                Catch ex As Exception

                End Try


                For Each drHora As DataRow In dtHoras.Rows

                    'MessageBox.Show("hora: " & cuenta)

                    If Length(drHora(0)) > 0 Then
                        idOperario = drHora(0)
                        Windows.Forms.Application.DoEvents()
                        LProgreso.Text = "Importando : " & idOperario & " - " & fecha
                        Windows.Forms.Application.DoEvents()

                        If Length(drHora(columna)) > 0 Then

                            If IsNumeric(drHora(columna)) = True Then
                                hora = drHora(columna)
                                tipoHora = "HORAS"

                                InsertarPorBaseDeDatos(idOperario, numero, fecha, trabajo, tipoHora, hora, sNombreUnicoGlobal, numero, idtrab, bbdd)

                            Else
                                hora = 0
                                tipoHora = drHora(columna)
                                InsertarPorBaseDeDatos(idOperario, numero, fecha, trabajo, tipoHora, hora, sNombreUnicoGlobal, numero, idtrab, bbdd)
                            End If
                            'cuenta = cuenta + 1
                        Else
                            'cuenta = cuenta + 1
                            Continue For
                        End If
                    Else
                        Exit For
                    End If

                Next
                columna = columna + 1

                If columna < dtFecha.Columns.Count Then
                    PvProgreso.Value = columna
                End If


            End While
            If (PvProgreso.Value.Equals(dtFecha.Columns.Count - 1)) Then
                MsgBox("Se han insertado las filas correctamente.")
            End If
        Else
            Exit Sub
        End If
    End Sub
    Sub InsertarPorBaseDeDatos(ByVal Operario As String, ByVal IdObra As String, ByVal Fecha As Date, ByVal cboTrabajo As String, ByVal sTipoHora As String, ByVal N_Horas As Double, ByVal sNombreUnico As String, ByVal numero As String, ByVal idtrab As String, ByVal bbdd As String)

        Dim obj As New Solmicro.Expertis.Business.General.Operario
        Dim txtSQL As String
        Dim rs As New DataTable
        Dim rsTrabajo As New DataTable
        Dim rsOperario As New DataTable
        Dim rsCalendarioCentro As New DataTable
        Dim IdOperacion As String
        Dim CodTrabajo As String
        Dim DescTrabajo As String
        Dim IdTipoTrabajo As String
        Dim IdSubTipoTrabajo As Object
        Dim iVeces As Long
        Dim Coste_Hora As Double
        Dim Tipo_Hora As String
        Dim I As Long
        Dim IdAutonumerico As Long
        Dim HorasFacturables As Integer
        Dim IdTrabajo As Double
        Dim HorasOrigen As Double
        Dim dia As String
        dia = Date.Now.Date
        Dim f As New Filter

        'Antes de insertar compruebo si existe el Operario
        f.Add("IdOperario", FilterOperator.Equal, Operario)
        rs = New BE.DataEngine().Filter(bbdd & "..tbMaestroOperario", f)

        If rs.Rows.Count = 0 Then
            MsgBox("El operario: '" & Operario & "' no existe en la BBDD. Todo el proceso se cancelara", vbExclamation + vbOKOnly)
            iVeces = "Error Provocado"
        End If

        rs = Nothing

        IdOperacion = "Guardar Datos"
        HorasOrigen = N_Horas
        Dim objTrabajo As New ObraTrabajo
        Dim filtro2 As New Filter
        Dim filtro3 As New Filter
        'Guardos los datos
        If IdOperacion = "Guardar Datos" Then

            'txtSQL = "Select IdTrabajo, CodTrabajo, DescTrabajo, IdTipoTrabajo, IdSubtipoTrabajo from tbObraTrabajo where IdObra=" & numero & " and Codtrabajo='" & cboTrabajo & "'"
            filtro2.Add("IDObra", FilterOperator.Equal, numero)
            filtro2.Add("IdTrabajo", FilterOperator.Equal, idtrab)
            rsTrabajo = New BE.DataEngine().Filter(bbdd & "..tbObraTrabajo", filtro2)
            'rsTrabajo = objTrabajo.Filter(filtro2, , "IdTrabajo, CodTrabajo, DescTrabajo, IdTipoTrabajo, IdSubtipoTrabajo")

            If rsTrabajo.Rows.Count = 0 Then
                IdTrabajo = Nothing
                CodTrabajo = ""
                DescTrabajo = ""
                IdTipoTrabajo = Nothing
                IdSubTipoTrabajo = Nothing
            Else
                IdTrabajo = rsTrabajo.Rows(0)("IdTrabajo")
                CodTrabajo = rsTrabajo.Rows(0)("CodTrabajo")
                DescTrabajo = rsTrabajo.Rows(0)("DescTrabajo")
                IdTipoTrabajo = rsTrabajo.Rows(0)("IdTipoTrabajo")
                IdSubTipoTrabajo = rsTrabajo.Rows(0)("IdSubtipotrabajo")
            End If

            'Obtengo datos del Operario
            'txtSQL = "Select Jornada_Laboral, c_h_n, c_h_x, c_h_e from tbMaestroOperario where idoperario='" & Operario & "'"
            'rsOperario = Conexion.Execute(txtSQL)
            filtro3.Add("IDOperario", FilterOperator.Equal, Operario)
            rsOperario = New BE.DataEngine().Filter(bbdd & "..frmMntoOperario", filtro3)

            'rsOperario = obj.Filter(f, , "Jornada_Laboral, c_h_n, c_h_x, c_h_e")

            'Compruebo en el calendario
            'txtSQL = "Select * from tbCalendarioCentro where idcentro='" & numero & "' and Fecha='" & Fecha & "' and tipodia=1"
            'rsCalendarioCentro = Conexion.Execute(txtSQL)
            Dim calendario As New General.CalendarioCentro
            Dim filtro As New Filter
            'Se pone por defecto 100. Para que no haga falta crear el calendario centro siempre que se cree obra.
            '100 es vegas altas. Para que no haya errores.David velasco 06/05/22
            'Antes ponia numero en vez de vegas
            Dim vegas As String
            vegas = "100"
            filtro.Add("Fecha", FilterOperator.Equal, Fecha)
            filtro.Add("IDCentro", FilterOperator.Equal, vegas)
            filtro.Add("TipoDia", FilterOperator.Equal, 1)

            rsCalendarioCentro = New BE.DataEngine().Filter(DB_TECOZAM & "..tbCalendarioCentro", filtro)

            'David 15/11/21 En vez de <>0 ponia "=0"
            'Si tiene datos es que es festivo
            If rsCalendarioCentro.Rows.Count <> 0 Then
                iVeces = 1
                N_Horas = Nz(N_Horas, 0)
                Coste_Hora = Nz(rsOperario.Rows(0)("c_h_e"), 0)
                Tipo_Hora = "HE"
            Else
                'Si no es festivo
                If rsOperario.Rows(0)("Jornada_Laboral") >= N_Horas Then
                    'Todas son horas normales
                    iVeces = 1
                    N_Horas = N_Horas
                    Coste_Hora = Nz(rsOperario.Rows(0)("c_h_n"), 0)
                    Tipo_Hora = "HO"
                Else
                    'Hay horas normales y horas extras, primero pongo las horas normales
                    iVeces = 2
                    Coste_Hora = Nz(rsOperario.Rows(0)("c_h_n"))
                    N_Horas = Nz(rsOperario.Rows(0)("Jornada_Laboral"), 0)
                    Tipo_Hora = "HO"
                End If
            End If

            'Tipo de hora que se inserta
            If sTipoHora <> "HORAS" Then
                Tipo_Hora = sTipoHora
                iVeces = 1
            End If

            For I = 1 To iVeces
                Dim auto As New OperarioCalendario
                IdAutonumerico = auto.Autonumerico()

                'Horas Facturables
                If Trim(DescTrabajo) = "HORAS FACTURABLES" Then
                    HorasFacturables = 1
                Else
                    HorasFacturables = 0
                End If

                Dim IDCategoriaProfesionalSCCP As String = ""
                Dim IDOficio As String = ""

                IDCategoriaProfesionalSCCP = DevuelveIDCategoriaProfesionalSCCP(bbdd, Operario)
                IDOficio = DevuelveIDOficio(bbdd, Operario)

                If Tipo_Hora = "HB" Then
                    txtSQL = "Insert into " & bbdd & "..tbObraMODControl(IdLineaModControl, IdTrabajo, IdObra, CodTrabajo, DescTrabajo, IdTipoTrabajo, " & _
                        "IdSubTipoTrabajo, IdOperario, IdCategoria, IdHora, FechaInicio, HorasRealMod, TasaRealModA, " & _
                         "ImpRealModA, HorasFactMod, ImpFactModA, DescParte, Facturable, FechaCreacionAudi, FechaModificacionAudi, Usuarioaudi, IDOficio, IdTipoTurno, HorasAdministrativas, IDCategoriaProfesionalSCCP, HorasBaja) " & _
                         "Values(" & IdAutonumerico & ", " & IdTrabajo & ", " & IdObra & ", '" & _
                         CodTrabajo & "', '" & DescTrabajo & "', '" & IdTipoTrabajo & "', '" & _
                         IdSubTipoTrabajo & "', '" & Operario & "', 'PREDET', '" & _
                         "HB" & "', '" & Fecha & "', 0 , " & 0 & ", " & 0 & _
                         ", 0 , " & 0 & _
                         ", '" & sNombreUnico & "', " & 0 & ", '" & Date.Now.Date & "', '" & Date.Now.Date & "', '" & ExpertisApp.UserName & "','" & IDOficio & "', 4, 0 ," & Nz(IDCategoriaProfesionalSCCP, "") & ", 8)"

                ElseIf Tipo_Hora = "HA" Then
                    txtSQL = "Insert into " & bbdd & "..tbObraMODControl(IdLineaModControl, IdTrabajo, IdObra, CodTrabajo, DescTrabajo, IdTipoTrabajo, " & _
                        "IdSubTipoTrabajo, IdOperario, IdCategoria, IdHora, FechaInicio, HorasRealMod, TasaRealModA, " & _
                        "ImpRealModA, HorasFactMod, ImpFactModA, DescParte, Facturable, FechaCreacionAudi, FechaModificacionAudi, Usuarioaudi, IDOficio, IdTipoTurno, HorasAdministrativas, IDCategoriaProfesionalSCCP) " & _
                        "Values(" & IdAutonumerico & ", " & IdTrabajo & ", " & IdObra & ", '" & _
                        CodTrabajo & "', '" & DescTrabajo & "', '" & IdTipoTrabajo & "', '" & _
                        IdSubTipoTrabajo & "', '" & Operario & "', 'PREDET', '" & _
                        "HA" & "', '" & Fecha & "', 0 , " & 0 & ", " & 0 & _
                        ", 0 , " & 0 & _
                        ", '" & sNombreUnico & "', " & 0 & ", '" & Date.Now.Date & "', '" & Date.Now.Date & "', '" & ExpertisApp.UserName & "','" & IDOficio & "', 4, 8 ," & Nz(IDCategoriaProfesionalSCCP, "") & ")"

                Else
                    txtSQL = "Insert into " & bbdd & " ..tbObraMODControl(IdLineaModControl, IdTrabajo, IdObra, CodTrabajo, DescTrabajo, IdTipoTrabajo, " & _
                     "IdSubTipoTrabajo, IdOperario, IdCategoria, IdHora, FechaInicio, HorasRealMod, TasaRealModA, " & _
                     "ImpRealModA, HorasFactMod, ImpFactModA, DescParte, Facturable, FechaCreacionAudi, FechaModificacionAudi, Usuarioaudi, IdTipoTurno, IDCategoriaProfesionalSCCP, IDOficio) " & _
                     "Values(" & IdAutonumerico & ", " & IdTrabajo & ", " & IdObra & ", '" & _
                     CodTrabajo & "', '" & DescTrabajo & "', '" & IdTipoTrabajo & "', '" & _
                     IdSubTipoTrabajo & "', '" & Operario & "', 'PREDET', '" & _
                     Tipo_Hora & "', '" & Fecha & "', " & Replace(N_Horas, ",", ".") & _
                     ", " & Replace(Coste_Hora, ",", ".") & ", " & Replace(Round(CDbl(Coste_Hora) * CDbl(N_Horas), 2), ",", ".") & _
                     ", " & Replace(N_Horas, ",", ".") & ", " & Replace(Round(CDbl(Coste_Hora) * CDbl(N_Horas), 2), ",", ".") & _
                     ", '" & sNombreUnico & "', " & HorasFacturables & ", '" & dia & "', '" & dia & "', '" & ExpertisApp.UserName & "', 4," & Nz(IDCategoriaProfesionalSCCP, "") & ",'" & Nz(IDOficio, "") & "')"

                End If
                'If IDCategoriaProfesionalSCCP = 2 Or IDCategoriaProfesionalSCCP = 3 Then

                'Else
                '    txtSQL = "Insert into " & bbdd & "..tbObraMODControl(IdLineaModControl, IdTrabajo, IdObra, CodTrabajo, DescTrabajo, IdTipoTrabajo, " & _
                '        "IdSubTipoTrabajo, IdOperario, IdCategoria, IdHora, FechaInicio, HorasRealMod, TasaRealModA, " & _
                '         "ImpRealModA, HorasFactMod, ImpFactModA, DescParte, Facturable, FechaCreacionAudi, FechaModificacionAudi, Usuarioaudi, IdTipoTurno, HorasAdministrativas, IDCategoriaProfesionalSCCP, IDOficio) " & _
                '         "Values(" & IdAutonumerico & ", " & IdTrabajo & ", " & IdObra & ", '" & _
                '         CodTrabajo & "', '" & DescTrabajo & "', '" & IdTipoTrabajo & "', '" & _
                '         IdSubTipoTrabajo & "', '" & Operario & "', 'PREDET', '" & _
                '         Tipo_Hora & "', '" & Fecha & "', 0 , " & Replace(Coste_Hora, ",", ".") & ", " & Replace(Round(CDbl(Coste_Hora) * CDbl(N_Horas), 2), ",", ".") & _
                '         ", 0 , " & Replace(Round(CDbl(Coste_Hora) * CDbl(N_Horas), 2), ",", ".") & _
                '         ", '" & sNombreUnico & "', " & HorasFacturables & ", '" & dia & "', '" & dia & "', '" & ExpertisApp.UserName & "', 4," & Replace(N_Horas, ",", ".") & "," & IDCategoriaProfesionalSCCP & ",'" & IDOficio & "')"
                'End If

                'Inserto
                'Conexion.Execute(txtSQL)
                auto.Ejecutar(txtSQL)

                'Cambio valores, pongo las horas extras
                Coste_Hora = Nz(rsOperario.Rows(0)("c_h_x"), 0)
                N_Horas = Nz(CDbl(HorasOrigen) - CDbl(rsOperario.Rows(0)("Jornada_Laboral")), 0)
                Tipo_Hora = "HX"
            Next
            'Libero memoria
            'Conexion = Nothing
            rs = Nothing
            rsTrabajo = Nothing
            rsOperario = Nothing
            rsCalendarioCentro = Nothing
        End If
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
    Public Function DevuelveIDCategoriaProfesionalSCCPTodasBasesDeDatos(ByVal IDOperario As String) As Integer
        Dim dt As New DataTable
        Dim f As New Filter

        f.Add("IDOperario", FilterOperator.Equal, IDOperario)
        dt = New BE.DataEngine().Filter("vUnionOperariosCategoriaProfesional", f)
        If dt.Rows.Count > 0 Then
            Return dt(0)("CategoriaProfesionalSCCP")
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

    Public Function DevuelveIDObra(ByVal bbdd As String, ByVal NObra As String) As String
        Dim dtObra As New DataTable

        Dim filtro As New Filter
        Dim sql2 As String

        sql2 = "Select * from " & bbdd & "..tbObraCabecera where NObra='" & NObra & "'"
        dtObra = aux.EjecutarSqlSelect(sql2)

        Return dtObra.Rows(0)("IDObra")
    End Function

    Dim dtResumen As New DataTable

    Private Sub bA3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bA3.Click

        'CREO LA TABLA A MODO RESUMEN QUE VA EN EL A3 UNIFICADO(HOJA 2 DEL EXCEL)
        'FormaTablaResumen()
        '------------------
        Dim dtFinal As New DataTable
        FormaTablaFinal(dtFinal)
        Dim dtAuxiliar As New DataTable
        Do
            ' Aquí va el código que deseas ejecutar repetidamente
            dtAuxiliar = CargaExcelA3()
            If dtAuxiliar Is Nothing Then
                ExpertisApp.GenerateMessage("Proceso cancelado correctamente.")
                Exit Sub
            End If
            For Each row As DataRow In dtAuxiliar.Rows
                dtFinal.ImportRow(row)
            Next
            ' Preguntar al usuario si desea continuar
            Dim respuesta As DialogResult = MessageBox.Show("¿Deseas cargar algún Excel más?", "Continuar", MessageBoxButtons.YesNo)
            ' Salir del bucle si el usuario responde "No"
            If respuesta = DialogResult.No Then
                Exit Do
            End If
        Loop
        'VALORES IMPORTANTES
        Dim mes As String
        Dim Anio As String
        Dim ultimoCaracter As String = lblRuta.Text.Substring(lblRuta.Text.Length - 1)

        If ultimoCaracter = "x" Then
            mes = lblRuta.Text.Substring(lblRuta.Text.Length - 9, 2)
            Anio = lblRuta.Text.Substring(lblRuta.Text.Length - 7, 2)
        Else
            mes = lblRuta.Text.Substring(lblRuta.Text.Length - 8, 2)
            Anio = lblRuta.Text.Substring(lblRuta.Text.Length - 6, 2)
        End If

        Anio = "20" & Anio
        'GENERA EXCEL
        GeneraExcel(mes, Anio, dtFinal)
        MsgBox("Fichero generado correctamente en N:\10. AUXILIARES\00. EXPERTIS\02. A3.")
    End Sub
    Public Function CargaExtrasTabla(ByVal dtUnion As DataTable) As DataTable
        Dim CD As New OpenFileDialog()
        CD.Title = "Seleccionar archivos"
        CD.Filter = "Archivos Excel(*.xls;*.xlsx)|*.xls;*xlsx|Todos los archivos(*.*)|*.*"
        CD.ShowDialog()

        If CD.FileName <> "" Then
            lblRuta.Text = CD.FileName
        End If

        'La hoja siempre es 1
        Dim hoja As String = "1"
        Dim dt As New DataTable
        Dim ruta As String = lblRuta.Text
        Dim empresa As String = DevuelveValorEntreParentesis(ruta)
        Dim rango As String = ""
        Select Case empresa
            Case "T. ES.", "FERR.", "SEC."
                rango = "B10:Z10000"
        End Select

        dt = ObtenerDatosExcel(ruta, hoja, rango)

        Dim mes As String
        Dim anio As String

        'CHECK DE QUE EL FICHERO ACABA EN XLSX O XLS
        Dim ultimoCaracter As String = ruta.Substring(ruta.Length - 1)

        If ultimoCaracter = "x" Then
            mes = ruta.Substring(ruta.Length - 9, 2)
            anio = ruta.Substring(ruta.Length - 7, 2)
        Else
            mes = ruta.Substring(ruta.Length - 8, 2)
            anio = ruta.Substring(ruta.Length - 6, 2)
        End If
        anio = "20" & anio

        'FORMO LA TABLA FINAL
        dtUnion = FormarTablaExtraPorEmpresa(dt, mes, anio, empresa)
        Return dtUnion
    End Function

    Public Function CargaExcelA3() As DataTable
        Dim CD As New OpenFileDialog()
        CD.Title = "Seleccionar archivos"
        CD.Filter = "Archivos Excel(*.xls;*.xlsx)|*.xls;*xlsx|Todos los archivos(*.*)|*.*"
        CD.ShowDialog()

        If CD.FileName <> "" Then
            lblRuta.Text = CD.FileName
        End If

        'La hoja siempre es 1
        Dim hoja As String = "1"
        Dim dt As New DataTable
        Dim ruta As String = lblRuta.Text
        Dim empresa As String = DevuelveValorEntreParentesis(ruta)
        Dim rango As String = ""
        Select Case empresa
            Case "T. ES.", "FERR.", "SEC."
                rango = "B10:Z10000"
            Case "D. P."
                rango = "A3:F10000"
            Case "T. UK."
                rango = "A2:Q10000"
            Case "T. NO."
                rango = "A2:T10000"
            Case Else
                MsgBox("El nombre identificado entre parentesis no se reconoce pero funciona. Coje las 3 primeras columnas.")
                rango = "A2:C10000"
        End Select

        dt = ObtenerDatosExcel(ruta, hoja, rango)

        Dim mes As String
        Dim anio As String

        'CHECK DE QUE EL FICHERO ACABA EN XLSX O XLS
        Dim ultimoCaracter As String = ruta.Substring(ruta.Length - 1)

        If ultimoCaracter = "x" Then
            mes = ruta.Substring(ruta.Length - 9, 2)
            anio = ruta.Substring(ruta.Length - 7, 2)
        Else
            mes = ruta.Substring(ruta.Length - 8, 2)
            anio = ruta.Substring(ruta.Length - 6, 2)
        End If
        anio = "20" & anio

        'FORMO LA TABLA FINAL
        dt = FormarTablaPorEmpresa(dt, mes, anio, empresa)
        Return dt
    End Function

    Public Sub FormaTablaFinal(ByRef dtFinal As DataTable)

        dtFinal.Columns.Add("IDGET")
        dtFinal.Columns.Add("IDOperario")
        dtFinal.Columns.Add("DescOperario")
        dtFinal.Columns.Add("CosteEmpresa", System.Type.GetType("System.Double"))
        dtFinal.Columns.Add("Mes")
        dtFinal.Columns.Add("Anio")
        dtFinal.Columns.Add("Empresa")
    End Sub

    Public Sub FormaTablaResumen()
        dtResumen.Columns.Add("Sociedad")
        dtResumen.Columns.Add("Importe A3 origen", System.Type.GetType("System.Double"))
        dtResumen.Columns.Add("Tipo Moneda")
        dtResumen.Columns.Add("Cambio", System.Type.GetType("System.Double"))
        dtResumen.Columns.Add("Importe A3 final(€)", System.Type.GetType("System.Double"))
    End Sub

    Public Function FormarTablaExtraPorEmpresa(ByVal dt As DataTable, ByVal mes As String, ByVal anio As String, ByVal empresa As String) As DataTable

        Dim newDataTable As DataTable = New DataTable
        newDataTable.Columns.Add("Empresa")
        newDataTable.Columns.Add("IDGET")
        newDataTable.Columns.Add("IDOperario")
        newDataTable.Columns.Add("DescOperario")
        newDataTable.Columns.Add("IDCategoriaProfesionalSCCP", System.Type.GetType("System.String"))
        newDataTable.Columns.Add("SinIncentivos", System.Type.GetType("System.Double"))
        newDataTable.Columns.Add("ConIncentivos", System.Type.GetType("System.Double"))
        newDataTable.Columns.Add("Mes", System.Type.GetType("System.Double"))
        newDataTable.Columns.Add("Anio", System.Type.GetType("System.Double"))

        Dim bbdd As String = ""
        If empresa = "T. ES." Then
            bbdd = DB_TECOZAM
            newDataTable = FormaTablaExtraEspaña(dt, newDataTable, bbdd, mes, anio, empresa)
        ElseIf empresa = "FERR." Then
            bbdd = DB_FERRALLAS
            newDataTable = FormaTablaExtraEspaña(dt, newDataTable, bbdd, mes, anio, empresa)
        ElseIf empresa = "SEC." Then
            bbdd = DB_SECOZAM
            newDataTable = FormaTablaExtraEspaña(dt, newDataTable, bbdd, mes, anio, empresa)
        End If

        Return newDataTable
    End Function

    Public Function FormarTablaPorEmpresa(ByVal dt As DataTable, ByVal mes As String, ByVal anio As String, ByVal empresa As String) As DataTable

        Dim newDataTable As DataTable = New DataTable
        newDataTable.Columns.Add("IDGET")
        newDataTable.Columns.Add("IDOperario")
        newDataTable.Columns.Add("DescOperario")
        newDataTable.Columns.Add("CosteEmpresa", System.Type.GetType("System.Double"))
        newDataTable.Columns.Add("Mes")
        newDataTable.Columns.Add("Anio")
        newDataTable.Columns.Add("Empresa")

        Dim bbdd As String = ""
        If empresa = "T. ES." Then
            bbdd = DB_TECOZAM
            newDataTable = FormaTablaEspaña(dt, newDataTable, bbdd, mes, anio, empresa)
        ElseIf empresa = "FERR." Then
            bbdd = DB_FERRALLAS
            newDataTable = FormaTablaEspaña(dt, newDataTable, bbdd, mes, anio, empresa)
        ElseIf empresa = "SEC." Then
            bbdd = DB_SECOZAM
            newDataTable = FormaTablaEspaña(dt, newDataTable, bbdd, mes, anio, empresa)
        ElseIf empresa = "D. P." Then
            bbdd = DB_DCZ
            newDataTable = FormaTablaDCZ(dt, newDataTable, bbdd, mes, anio, empresa)
        ElseIf empresa = "T. UK." Then
            bbdd = DB_UK
            newDataTable = FormaTablaUK(dt, newDataTable, bbdd, mes, anio, empresa)
        ElseIf empresa = "T. NO." Then
            bbdd = DB_NO
            newDataTable = FormaTablaNO(dt, newDataTable, bbdd, mes, anio, empresa)
        Else
            newDataTable = FormaTablaTipo(dt, newDataTable, mes, anio)
        End If

        Return newDataTable
    End Function

    Public Function FormaTablaTipo(ByVal dt As DataTable, ByVal newDataTable As DataTable, ByVal mes As String, ByVal anio As String)
        Dim IDOperario As String
        Dim Diccionario As String
        Dim descOperario As String
        Dim bbdd As String
        For Each row As DataRow In dt.Rows
            If Len(row("F1").ToString) = 0 Then
                'Return newDataTable
                Exit For ' Salir del bucle si la celda está vacía
            End If
            Dim newRow As DataRow = newDataTable.NewRow()

            If row("F3").ToString = "T. ES." Then
                bbdd = DB_TECOZAM
                IDOperario = DevuelveIDOperario(bbdd, row("F1"))
                descOperario = DevuelveDescOperario(bbdd, IDOperario)
                newRow("IDOperario") = IDOperario
                newRow("DescOperario") = descOperario
                newRow("IDGET") = DevuelveIDGET(bbdd, IDOperario)

            ElseIf row("F3").ToString = "FERR." Then
                bbdd = DB_FERRALLAS
                IDOperario = DevuelveIDOperario(bbdd, row("F1"))
                descOperario = DevuelveDescOperario(bbdd, IDOperario)
                newRow("IDOperario") = IDOperario
                newRow("DescOperario") = descOperario
                newRow("IDGET") = DevuelveIDGET(bbdd, IDOperario)

            ElseIf row("F3").ToString = "SEC." Then
                bbdd = DB_SECOZAM
                IDOperario = DevuelveIDOperario(bbdd, row("F1"))
                descOperario = DevuelveDescOperario(bbdd, IDOperario)
                newRow("IDOperario") = IDOperario
                newRow("DescOperario") = descOperario
                newRow("IDGET") = DevuelveIDGET(bbdd, IDOperario)

            ElseIf row("F3").ToString = "T. UK." Then
                bbdd = DB_UK
                Diccionario = row("F1")
                IDOperario = DevuelveIDOperarioDiccionario(bbdd, Diccionario)
                descOperario = DevuelveDescOperario(bbdd, IDOperario)
                newRow("IDOperario") = IDOperario
                newRow("DescOperario") = descOperario
                newRow("IDGET") = DevuelveIDGET(bbdd, IDOperario)

            ElseIf row("F3").ToString = "D. P." Then
                bbdd = DB_DCZ
                Diccionario = row("F1")
                IDOperario = DevuelveIDOperarioDiccionario(bbdd, Diccionario)
                descOperario = DevuelveDescOperario(bbdd, IDOperario)
                newRow("IDOperario") = IDOperario
                newRow("DescOperario") = descOperario
                newRow("IDGET") = DevuelveIDGET(bbdd, IDOperario)
            End If

            newRow("CosteEmpresa") = row("F2")
            newRow("Mes") = mes
            newRow("Anio") = anio
            newRow("Empresa") = row("F3")
            newDataTable.Rows.Add(newRow)
        Next

        'CHECK DE QUE EL EXCEL RESULTANTE TIENE EL MISMO COSTE EMPRESA TOTAL
        Dim CosteE1 As Double = 0
        Dim CosteEFinal As Double = 0

        For Each dr As DataRow In dt.Rows
            If Len(dr("F1").ToString) = 0 Then
                'Return newDataTable
                Exit For ' Salir del bucle si la celda está vacía
            End If
            CosteE1 = CosteE1 + dr("F2")
        Next

        For Each dr As DataRow In newDataTable.Rows
            CosteEFinal = CosteEFinal + dr("CosteEmpresa")
        Next

        Dim result As DialogResult = MessageBox.Show("El coste del excel introducido es " & CosteE1.ToString("N2") & "€. El del excel resultante es " & CosteEFinal.ToString("N2") & "€.", "¿Desea Continuar?", MessageBoxButtons.YesNo)
        If result = DialogResult.No Then
            Return Nothing
            Exit Function
        End If
        Return newDataTable
    End Function
    Public Function FormaTablaDCZ(ByVal dt As DataTable, ByVal newDataTable As DataTable, ByVal bbdd As String, ByVal mes As String, ByVal anio As String, ByVal empresa As String)

        Dim IDOperario As String = ""
        Dim diccionario As String = ""
        Dim descOperario As String = ""

        ' Copiar los datos de las columnas seleccionadas al nuevo DataTable
        For Each row As DataRow In dt.Rows
            'Verificar si la celda está vacía
            If Len(row("F1").ToString) = 0 Or row("F1").ToString = "TOTAL" Then
                'Return newDataTable
                Exit For ' Salir del bucle si la celda está vacía
            End If

            Dim newRow As DataRow = newDataTable.NewRow()

            Dim parts() As String = row("F1").ToString.Split("-"c)

            ' Eliminar espacios adicionales en cada parte
            For i As Integer = 0 To parts.Length - 1
                parts(i) = parts(i).Trim()
            Next

            diccionario = parts(0)
            IDOperario = DevuelveIDOperarioDiccionario(bbdd, diccionario)
            newRow("IDOperario") = IDOperario
            newRow("DescOperario") = parts(1)
            newRow("IDGET") = DevuelveIDGET(bbdd, IDOperario)
            newRow("CosteEmpresa") = row("F3") + row("F4")
            newRow("Mes") = mes
            newRow("Anio") = anio
            newRow("Empresa") = empresa

            newDataTable.Rows.Add(newRow)
        Next

        Dim dtOrdenada As New DataTable
        newDataTable.DefaultView.Sort = "IDOperario asc"
        dtOrdenada = newDataTable.DefaultView.ToTable

        'CHECK DE QUE EL EXCEL RESULTANTE TIENE EL MISMO COSTE EMPRESA TOTAL
        Dim CosteE1 As Double = 0
        Dim CosteEFinal As Double = 0

        For Each dr As DataRow In dt.Rows
            If Len(dr("F1").ToString) = 0 Then
                'Return newDataTable
                Exit For ' Salir del bucle si la celda está vacía
            End If
            CosteE1 = CosteE1 + dr("F3") + dr("F4")
        Next

        For Each dr As DataRow In dtOrdenada.Rows
            CosteEFinal = CosteEFinal + dr("CosteEmpresa")
        Next

        Dim result As DialogResult = MessageBox.Show("El coste del excel introducido es " & CosteE1.ToString("N2") & "€. El del excel resultante es " & CosteEFinal.ToString("N2") & "€.", "¿Desea Continuar?", MessageBoxButtons.YesNo)
        If result = DialogResult.No Then
            Return Nothing
            Exit Function
        End If

        Dim fila As DataRow = dtResumen.NewRow()
        fila("Sociedad") = empresa
        fila("Importe A3 origen") = CosteE1
        fila("Tipo Moneda") = "€"
        fila("Cambio") = 1
        fila("Importe A3 final(€)") = CosteEFinal
        dtResumen.Rows.Add(fila)

        Return dtOrdenada
    End Function


    Public Function FormaTablaUK(ByVal dt As DataTable, ByVal newDataTable As DataTable, ByVal bbdd As String, ByVal mes As String, ByVal anio As String, ByVal empresa As String)

        Dim IDOperario As String = ""
        Dim diccionario As String = ""
        Dim totaleuros As Double = 0
        Dim totallibras As Double = 0

        'TABLA DE CAMBIO DE MONEDA LIBRAS
        Dim ruta As String
        ruta = "\\stor01\dg\SCCP_Prueba\03. COSTES\TIPO DE CAMBIO MONEDA.xlsx"
        Dim hoja As String = "TIPO DE CAMBIO"
        Dim rango As String = "A1:K10000"
        Dim dtCambioMoneda As New DataTable
        dtCambioMoneda = ObtenerDatosExcel(ruta, hoja, rango)


        ' Copiar los datos de las columnas seleccionadas al nuevo DataTable
        For Each row As DataRow In dt.Rows
            'Verificar si la celda está vacía
            If Len(row("F1").ToString) = 0 Then
                'Return newDataTable
                Exit For ' Salir del bucle si la celda está vacía
            End If

            Dim newRow As DataRow = newDataTable.NewRow()

            IDOperario = DevuelveIDOperarioDiccionario(bbdd, row("F1"))
            newRow("IDOperario") = IDOperario
            newRow("DescOperario") = row("F2")
            newRow("IDGET") = DevuelveIDGET(bbdd, IDOperario)
            totallibras = Nz(row("F9"), 0) + Nz(row("F10"), 0) + Nz(row("F15"), 0) + Nz(row("F16"), 0)
            totaleuros = CambioLibraAEuro(dtCambioMoneda, totallibras, mes, anio)
            newRow("CosteEmpresa") = totaleuros
            newRow("Mes") = mes
            newRow("Anio") = anio
            newRow("Empresa") = empresa

            newDataTable.Rows.Add(newRow)
        Next

        Dim dtOrdenada As New DataTable
        newDataTable.DefaultView.Sort = "IDOperario asc"
        dtOrdenada = newDataTable.DefaultView.ToTable

        'CHECK DE QUE EL EXCEL RESULTANTE TIENE EL MISMO COSTE EMPRESA TOTAL
        Dim CosteE1 As Double = 0
        Dim CosteEFinal As Double = 0

        For Each dr As DataRow In dt.Rows
            If Len(dr("F1").ToString) = 0 Then
                'Return newDataTable
                Exit For ' Salir del bucle si la celda está vacía
            End If
            CosteE1 = CosteE1 + Nz(dr("F9"), 0) + Nz(dr("F10"), 0) + Nz(dr("F15"), 0) + Nz(dr("F16"), 0)
        Next

        For Each dr As DataRow In dtOrdenada.Rows
            CosteEFinal = CosteEFinal + dr("CosteEmpresa")
        Next
        Dim result As DialogResult = MessageBox.Show("El coste del excel introducido es " & CosteE1.ToString("N2") & _
        " £ =" & CambioLibraAEuro(dtCambioMoneda, CosteE1, mes, anio).ToString("N2") & " €. El del excel resultante es " & CosteEFinal.ToString("N2") & _
        "€." & vbCrLf & "El cambio usado es: " & DevuelveCambioMoneda(dtCambioMoneda, mes, anio), "¿Desea Continuar?", MessageBoxButtons.YesNo)
        If result = DialogResult.No Then
            Return Nothing
            Exit Function
        End If

        Dim fila As DataRow = dtResumen.NewRow()
        fila("Sociedad") = empresa
        fila("Importe A3 origen") = CosteE1.ToString("N2")
        fila("Tipo Moneda") = dtCambioMoneda(0)("F4")
        fila("Cambio") = DevuelveCambioMoneda(dtCambioMoneda, mes, anio)
        fila("Importe A3 final(€)") = CosteEFinal
        dtResumen.Rows.Add(fila)

        Return dtOrdenada
    End Function
    Public Function FormaTablaNO(ByVal dt As DataTable, ByVal newDataTable As DataTable, ByVal bbdd As String, ByVal mes As String, ByVal anio As String, ByVal empresa As String)

        Dim IDOperario As String = ""
        Dim diccionario As String = ""
        Dim totaleuros As Double = 0
        Dim totalcoronas As Double = 0

        'TABLA DE CAMBIO DE MONEDA LIBRAS
        Dim ruta As String
        ruta = "\\stor01\dg\SCCP_Prueba\03. COSTES\TIPO DE CAMBIO MONEDA.xlsx"
        Dim hoja As String = "TIPO DE CAMBIO"
        Dim rango As String = "A1:K10000"
        Dim dtCambioMoneda As New DataTable
        dtCambioMoneda = ObtenerDatosExcel(ruta, hoja, rango)


        ' Copiar los datos de las columnas seleccionadas al nuevo DataTable
        For Each row As DataRow In dt.Rows
            'Verificar si la celda está vacía
            If Len(row("F1").ToString) = 0 Then
                'Return newDataTable
                Exit For ' Salir del bucle si la celda está vacía
            End If

            Dim newRow As DataRow = newDataTable.NewRow()

            IDOperario = DevuelveIDOperarioDiccionario(bbdd, row("F1"))
            newRow("IDOperario") = IDOperario
            newRow("DescOperario") = row("F2")
            newRow("IDGET") = DevuelveIDGET(bbdd, IDOperario)
            totalcoronas = Nz(row("F16"), 0)
            totaleuros = CambioCoronaAEuro(dtCambioMoneda, totalcoronas, mes, anio)
            newRow("CosteEmpresa") = totaleuros
            newRow("Mes") = mes
            newRow("Anio") = anio
            newRow("Empresa") = empresa

            newDataTable.Rows.Add(newRow)
        Next

        Dim dtOrdenada As New DataTable
        newDataTable.DefaultView.Sort = "IDOperario asc"
        dtOrdenada = newDataTable.DefaultView.ToTable

        'CHECK DE QUE EL EXCEL RESULTANTE TIENE EL MISMO COSTE EMPRESA TOTAL
        Dim CosteE1 As Double = 0
        Dim CosteEFinal As Double = 0

        For Each dr As DataRow In dt.Rows
            If Len(dr("F1").ToString) = 0 Then
                'Return newDataTable
                Exit For ' Salir del bucle si la celda está vacía
            End If
            CosteE1 = CosteE1 + Nz(dr("F16"), 0)
        Next

        For Each dr As DataRow In dtOrdenada.Rows
            CosteEFinal = CosteEFinal + dr("CosteEmpresa")
        Next
        Dim result As DialogResult = MessageBox.Show("El coste del excel introducido es " & CosteE1.ToString("N2") & _
        " NOK =" & CambioCoronaAEuro(dtCambioMoneda, CosteE1, mes, anio).ToString("N2") & " €. El del excel resultante es " & CosteEFinal.ToString("N2") & _
        "€." & vbCrLf & "El cambio usado es: " & DevuelveCambioMonedaCorona(dtCambioMoneda, mes, anio), "¿Desea Continuar?", MessageBoxButtons.YesNo)
        If result = DialogResult.No Then
            Return Nothing
            Exit Function
        End If

        Dim fila As DataRow = dtResumen.NewRow()
        fila("Sociedad") = empresa
        fila("Importe A3 origen") = CosteE1.ToString("N2")
        fila("Tipo Moneda") = dtCambioMoneda(0)("F10")
        fila("Cambio") = DevuelveCambioMonedaCorona(dtCambioMoneda, mes, anio)
        fila("Importe A3 final(€)") = CosteEFinal
        dtResumen.Rows.Add(fila)

        Return dtOrdenada
    End Function
    Public Function DevuelveCambioMoneda(ByVal dtCambioMoneda As DataTable, ByVal mes As String, ByVal anio As String) As Double
        Dim fecha As String
        Dim cambioMoneda As Double

        For Each dr As DataRow In dtCambioMoneda.Rows
            Try
                fecha = dr("F1")
                If Month(fecha) = mes And Year(fecha) = anio Then
                    cambioMoneda = dr("F4")
                    Return cambioMoneda
                End If
            Catch ex As Exception
            End Try
        Next
    End Function
    Public Function DevuelveCambioMonedaCorona(ByVal dtCambioMoneda As DataTable, ByVal mes As String, ByVal anio As String) As Double
        Dim fecha As String
        Dim cambioMoneda As Double

        For Each dr As DataRow In dtCambioMoneda.Rows
            Try
                fecha = dr("F1")
                If Month(fecha) = mes And Year(fecha) = anio Then
                    cambioMoneda = dr("F10")
                    Return cambioMoneda
                End If
            Catch ex As Exception
            End Try
        Next
    End Function

    Public Function CambioLibraAEuro(ByVal dtCambioMoneda As DataTable, ByVal totallibras As Double, ByVal mes As String, ByVal anio As String) As Double

        Dim fecha As String
        Dim cambioMoneda As Double

        For Each dr As DataRow In dtCambioMoneda.Rows
            Try
                fecha = dr("F1")
                If Month(fecha) = mes And Year(fecha) = anio Then
                    cambioMoneda = dr("F4")
                End If
            Catch ex As Exception
            End Try
        Next

        Return (totallibras * cambioMoneda)
    End Function


    Public Function CambioCoronaAEuro(ByVal dtCambioMoneda As DataTable, ByVal totalcoronas As Double, ByVal mes As String, ByVal anio As String) As Double

        Dim fecha As String
        Dim cambioMoneda As Double

        For Each dr As DataRow In dtCambioMoneda.Rows
            Try
                fecha = dr("F1")
                If Month(fecha) = mes And Year(fecha) = anio Then
                    cambioMoneda = dr("F10")
                End If
            Catch ex As Exception
            End Try
        Next

        Return (totalcoronas * cambioMoneda)
    End Function

    Public Function FormaTablaExtraEspaña(ByVal dt As DataTable, ByVal newDataTable As DataTable, ByVal bbdd As String, ByVal mes As String, ByVal anio As String, ByVal empresa As String)

        Try
            Dim IDOperario As String
            Dim incentivos As String
            For Each row As DataRow In dt.Rows
                'Verificar si la celda está vacía
                If Len(row("F1").ToString) = 0 Then
                    'Return newDataTable
                    Exit For ' Salir del bucle si la celda está vacía
                End If

                Dim newRow As DataRow = newDataTable.NewRow()
                IDOperario = DevuelveIDOperario(bbdd, row("F3"))
                newRow("IDOperario") = IDOperario
                newRow("DescOperario") = row("F2")
                newRow("IDCategoriaProfesionalSCCP") = DevuelveIDCategoriaProfesionalSCCPTodasBasesDeDatos(IDOperario)
                newRow("IDGET") = DevuelveIDGET(bbdd, IDOperario)
                incentivos = DevuelveIncentivos(bbdd, IDOperario)
                If incentivos = 1 Then
                    newRow("ConIncentivos") = row("F8")
                    newRow("SinIncentivos") = 0
                Else
                    If Len(row("F10").ToString()) = 0 Then
                        newRow("SinIncentivos") = 0
                    Else
                        newRow("SinIncentivos") = Math.Abs(row("F10"))
                    End If

                    newRow("ConIncentivos") = 0
                End If
                'newRow("ConIncentivos") = row("F8")
                If mes = 13 Then
                    mes = 6
                ElseIf mes = 14 Then
                    mes = 12
                End If
                newRow("Mes") = mes
                newRow("Anio") = anio
                newRow("Empresa") = empresa
                newDataTable.Rows.Add(newRow)
            Next
        Catch ex As Exception
            MsgBox("Error al leer las extras de " & empresa & " pero deja seguir.")
        End Try

        Return newDataTable

    End Function

    Public Function FormaTablaEspaña(ByVal dt As DataTable, ByVal newDataTable As DataTable, ByVal bbdd As String, ByVal mes As String, ByVal anio As String, ByVal empresa As String)

        Dim IDOperario As String = ""
        ' Copiar los datos de las columnas seleccionadas al nuevo DataTable
        For Each row As DataRow In dt.Rows
            'Verificar si la celda está vacía
            If Len(row("F1").ToString) = 0 Then
                'Return newDataTable
                Exit For ' Salir del bucle si la celda está vacía
            End If

            Dim newRow As DataRow = newDataTable.NewRow()
            IDOperario = DevuelveIDOperario(bbdd, row("F3"))
            newRow("IDOperario") = IDOperario
            newRow("DescOperario") = row("F2")
            newRow("IDGET") = DevuelveIDGET(bbdd, IDOperario)
            newRow("CosteEmpresa") = row("F8")
            newRow("Mes") = mes
            newRow("Anio") = anio
            newRow("Empresa") = empresa

            newDataTable.Rows.Add(newRow)
        Next

        Dim dtOrdenada As New DataTable
        newDataTable.DefaultView.Sort = "IDOperario asc"
        dtOrdenada = newDataTable.DefaultView.ToTable
        'AQUI RECORRO PARA UNIFICAR SI HUBIERA FINIQUITO
        dtOrdenada = CheckFiniquito(dtOrdenada)

        'CHECK DE QUE EL EXCEL RESULTANTE TIENE EL MISMO COSTE EMPRESA TOTAL
        Dim CosteE1 As Double = 0
        Dim CosteEFinal As Double = 0

        For Each dr As DataRow In dt.Rows
            If Len(dr("F1").ToString) = 0 Then
                'Return newDataTable
                Exit For ' Salir del bucle si la celda está vacía
            End If
            CosteE1 = CosteE1 + Nz(dr("F8"), 0)
        Next

        For Each dr As DataRow In dtOrdenada.Rows
            CosteEFinal = CosteEFinal + Nz(dr("CosteEmpresa"), 0)
        Next

        Dim result As DialogResult = MessageBox.Show("El coste del excel introducido es " & CosteE1.ToString("N2") & "€. El del excel resultante es " & CosteEFinal.ToString("N2") & "€.", "¿Desea Continuar?", MessageBoxButtons.YesNo)
        If result = DialogResult.No Then
            Return Nothing
            Exit Function
        End If
        'Añade la linea a modo resumen en la hoja 2 del Excel resultanto 09/10/23

        Dim fila As DataRow = dtResumen.NewRow()
        fila("Sociedad") = empresa
        fila("Importe A3 origen") = CosteE1
        fila("Tipo Moneda") = "€"
        fila("Cambio") = 1
        fila("Importe A3 final(€)") = CosteEFinal
        dtResumen.Rows.Add(fila)

        Return dtOrdenada
    End Function

    Public Function DevuelveIDOperario(ByVal bbdd As String, ByVal DNI As String) As String
        Dim f As New Filter
        f.Add("DNI", FilterOperator.Equal, DNI)
        Dim dt As DataTable
        dt = New BE.DataEngine().Filter(bbdd & "..frmMntoOperario", f)

        If dt.Rows.Count = 0 Then
            MsgBox("No existe este DNI " & DNI & " en " & bbdd)
            Exit Function
        End If
        Return dt.Rows(0)("IDOperario")
    End Function

    Public Function DevuelveDescOperario(ByVal bbdd As String, ByVal IDOperario As String) As String
        Dim f As New Filter
        f.Add("IDOperario", FilterOperator.Equal, IDOperario)
        Dim dt As DataTable
        dt = New BE.DataEngine().Filter(bbdd & "..frmMntoOperario", f)

        If dt.Rows.Count = 0 Then
            MsgBox("No existe este IDOperario " & IDOperario & " en " & bbdd)
            Exit Function
        End If
        Return dt.Rows(0)("DescOperario").ToString.ToUpper
    End Function

    Public Function DevuelveIDOperarioDiccionario(ByVal bbdd As String, ByVal Diccionario As String) As String
        Dim f As New Filter
        f.Add("Diccionario", FilterOperator.Equal, Diccionario)
        Dim dt As DataTable
        dt = New BE.DataEngine().Filter(bbdd & "..frmMntoOperario", f)

        If dt.Rows.Count = 0 Then
            MsgBox("No existe este Diccionario " & Diccionario & " en " & bbdd)
            Exit Function
        End If
        Return dt.Rows(0)("IDOperario")
    End Function


    Public Function DevuelveIDGET(ByVal bbdd As String, ByVal IDOperario As String) As String
        Dim f As New Filter
        f.Add("IDOperario", FilterOperator.Equal, IDOperario)
        Dim dt As DataTable
        dt = New BE.DataEngine().Filter(bbdd & "..frmMntoOperario", f)

        If Len(dt.Rows(0)("IDGET").ToString) = 0 Then
            MsgBox("No existe este IDGET para el IDOperario" & IDOperario & " en " & bbdd)
            Return ""
            Exit Function
        End If

        Return dt.Rows(0)("IDGET")
    End Function
    Public Function CheckFiniquito(ByVal dtOrdenada As DataTable) As DataTable
        'Recorro dtOrdenada si coinciden dos IDOperario sumo el costeempresa en una fila
        Dim dtFinal As DataTable = dtOrdenada.Clone()
        dtFinal.Clear()

        Dim contador As Integer = 0
        Dim acumulaFiniquito As Double = 0

        For Each fila As DataRow In dtOrdenada.Rows
            Try
                If dtOrdenada.Rows(contador)("IDOperario").ToString <> dtOrdenada.Rows(contador + 1)("IDOperario").ToString Then
                    dtOrdenada.Rows(contador)("CosteEmpresa") = Nz(dtOrdenada.Rows(contador)("CosteEmpresa"), 0) + acumulaFiniquito
                    dtFinal.ImportRow(fila)
                    acumulaFiniquito = 0
                Else
                    If dtOrdenada.Rows(contador)("IDOperario").ToString = dtOrdenada.Rows(contador + 1)("IDOperario").ToString Then
                        acumulaFiniquito = acumulaFiniquito + Nz(dtOrdenada(contador)("CosteEmpresa"), 0)
                    End If
                End If
            Catch ex As Exception
                dtOrdenada.Rows(contador)("CosteEmpresa") = Nz(dtOrdenada.Rows(contador)("CosteEmpresa"), 0) + acumulaFiniquito
                dtFinal.ImportRow(fila)
            End Try
            contador += 1
        Next

        Return dtFinal
    End Function

    Public Function DevuelveValorEntreParentesis(ByVal ruta As String)
        Dim input As String = ruta

        Dim startIndex As Integer = input.IndexOf("("c)
        Dim endIndex As Integer = input.IndexOf(")"c)

        If startIndex <> -1 AndAlso endIndex <> -1 AndAlso endIndex > startIndex + 1 Then
            ' Obtener el texto entre paréntesis
            Dim result As String = input.Substring(startIndex + 1, endIndex - startIndex - 1)
            Return result
        Else
            MsgBox("No se encontró ningún texto entre paréntesis.")
        End If
    End Function

    Public Sub GeneraExcel(ByVal mes As String, ByVal anio As String, ByVal dtFinal As DataTable)

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial

        Dim dtFinalOrdenado As New DataTable()
        dtFinalOrdenado.Columns.Add("Empresa", GetType(String))
        dtFinalOrdenado.Columns.Add("IDGET", GetType(String))
        dtFinalOrdenado.Columns.Add("IDOperario", GetType(String))
        dtFinalOrdenado.Columns.Add("DescOperario", GetType(String))
        dtFinalOrdenado.Columns.Add("Mes", GetType(String))
        dtFinalOrdenado.Columns.Add("Anio", GetType(String))
        dtFinalOrdenado.Columns.Add("CosteEmpresa", GetType(Decimal))
        dtFinalOrdenado.Columns.Add("IDCategoriaProfesionalSCCP", GetType(String))
        dtFinalOrdenado.Columns.Add("IDOficio", GetType(String))
        dtFinalOrdenado.Columns.Add("NObra", GetType(String))

        ' Copiar los datos del DataTable original al DataTable ordenado
        For Each dr As DataRow In dtFinal.Rows
            If dr("CosteEmpresa") > 0 Then
                Dim newRow As DataRow = dtFinalOrdenado.NewRow()
                newRow("Empresa") = dr("Empresa")
                newRow("IDGET") = dr("IDGET")
                newRow("IDOperario") = dr("IDOperario")
                'newRow("DescOperario") = dr("DescOperario")
                newRow("DescOperario") = DevuelveDescOperario(DevuelveBaseDeDatos(dr("Empresa")), dr("IDOperario"))
                newRow("Mes") = dr("Mes")
                newRow("Anio") = dr("Anio")
                newRow("CosteEmpresa") = dr("CosteEmpresa")
                newRow("IDCategoriaProfesionalSCCP") = DevuelveIDCategoriaProfesionalSCCP(DevuelveBaseDeDatos(dr("Empresa")), dr("IDOperario"))
                newRow("IDOficio") = DevuelveIDOficio(DevuelveBaseDeDatos(dr("Empresa")), dr("IDOperario"))
                dtFinalOrdenado.Rows.Add(newRow)
            End If
            
        Next

        Dim ruta As New FileInfo("N:\10. AUXILIARES\00. EXPERTIS\02. A3\" & mes & " A3 " & mes & anio.Substring(anio.Length - 2) & ".xlsx")
        'Dim ruta As New FileInfo("N:\01. A3\" & mes & " A3 " & mes & anio.Substring(anio.Length - 2) & ".xlsx")
        Dim rutaCadena As String = ""
        rutaCadena = ruta.FullName

        'Verificar si el archivo existe.
        If File.Exists(rutaCadena) Then
            'Si el archivo existe, eliminarlo.
            File.Delete(rutaCadena)
        End If

        Using package As New ExcelPackage(ruta)
            ' Crear una hoja de cálculo y obtener una referencia a ella.
            Dim worksheet = package.Workbook.Worksheets.Add("A3")

            ' Copiar los datos de la DataTable a la hoja de cálculo.
            worksheet.Cells("A1").LoadFromDataTable(dtFinalOrdenado, True)

            Dim columnaE As ExcelRange = worksheet.Cells("G2:G" & worksheet.Dimension.End.Row)
            columnaE.Style.Numberformat.Format = "#,##0.00€"

            ' Aplicar formato negrita a la fila 1
            Dim fila1 As ExcelRange = worksheet.Cells(1, 1, 1, worksheet.Dimension.End.Column)
            fila1.Style.Font.Bold = True
            worksheet.Column(4).Width = 30
            worksheet.Column(7).Width = 14

            'SEGUNDA HOJA DEL EXCEL QUE ES RESUMEN
            Dim resumenWorksheet = package.Workbook.Worksheets.Add("RESUMEN")
            resumenWorksheet.Cells("A1").LoadFromDataTable(dtResumen, True)

            Dim columnaBResumen As ExcelRange = resumenWorksheet.Cells("B2:B" & worksheet.Dimension.End.Row)
            columnaBResumen.Style.Numberformat.Format = "#,##0.00"

            Dim columnaDResumen As ExcelRange = resumenWorksheet.Cells("D2:D" & worksheet.Dimension.End.Row)
            columnaDResumen.Style.Numberformat.Format = "#,##0.00000€"

            Dim columnaEResumen As ExcelRange = resumenWorksheet.Cells("E2:E" & worksheet.Dimension.End.Row)
            columnaEResumen.Style.Numberformat.Format = "#,##0.00€"

            ' Aplicar formato negrita a la fila 1
            Dim filaResumen1 As ExcelRange = resumenWorksheet.Cells(1, 1, 1, resumenWorksheet.Dimension.End.Column)
            filaResumen1.Style.Font.Bold = True

            worksheet.Cells("A1:" & GetExcelColumnName(worksheet.Dimension.End.Column) & "1").AutoFilter = True
            resumenWorksheet.Column(2).Width = 15
            resumenWorksheet.Column(3).Width = 15
            resumenWorksheet.Column(5).Width = 15

            'TERCERA HOJA DEL EXCEL QUE ES EL RESUMEN POR CATEGORIA PROFESIONAL
            Dim dtA3CategoriaProfesional As DataTable
            dtA3CategoriaProfesional = DevuelveTablaA3PorCategoriaProf(dtFinal)

            Dim resumenCategoria = package.Workbook.Worksheets.Add("RESUMEN POR CATEGORIA")
            resumenCategoria.Cells("A1").LoadFromDataTable(dtA3CategoriaProfesional, True)

            Dim f1 As ExcelRange = resumenCategoria.Cells(1, 1, 1, resumenCategoria.Dimension.End.Column)
            f1.Style.Font.Bold = True

            Dim columnaB As ExcelRange = resumenCategoria.Cells("C2:C" & resumenCategoria.Dimension.End.Row)
            columnaB.Style.Numberformat.Format = "#,##0.00€"
            resumenCategoria.Column(3).Width = 15

            ' Congelar la primera fila
            worksheet.View.FreezePanes(2, 1)

            ' Guardar el archivo de Excel.
            package.Save()
        End Using

    End Sub
    Public Function DevuelveBaseDeDatos(ByVal empresa As String) As String
        Dim bbdd As String = ""
        If empresa = "T. ES." Then
            bbdd = "xTecozam50R2"
        ElseIf empresa = "FERR." Then
            bbdd = "xFerrallas50R2"
        ElseIf empresa = "SEC." Then
            bbdd = "xSecozam50R2"
        ElseIf empresa = "D. P." Then
            bbdd = "xDrenajesPortugal50R2"
        ElseIf empresa = "T. UK." Then
            bbdd = "xTecozamUnitedKingdom50R2"
        ElseIf empresa = "T. NO." Then
            bbdd = "xTecozamNorge50R2"
        End If
        Return bbdd
    End Function
    Public Function DevuelveTablaA3PorCategoriaProf(ByVal dtFinal As DataTable) As DataTable

        Dim dtResultado As New DataTable
        dtResultado.Columns.Add("Empresa", GetType(String))
        dtResultado.Columns.Add("IDCategoriaProfesionalSCCP", GetType(String))
        dtResultado.Columns.Add("Total", GetType(Double))

        Dim jprod As Double = 0 : Dim encar As Double = 0 : Dim operar As Double = 0 : Dim tecobra As Double = 0 : Dim staff As Double = 0 : Dim otros As Double = 0

        Dim IDOperario As String
        Dim cont As Integer = 0
        Dim empresa As String = ""
        Dim bbdd As String = ""
        Dim categoria As Integer
        Dim coste As Double = 0

        For Each dr As DataRow In dtFinal.Rows
            empresa = dr("Empresa").ToString
            bbdd = DevuelveBaseDeDatos(empresa)

            Try
                If dtFinal.Rows(cont)("Empresa") <> dtFinal.Rows(cont + 1)("Empresa") Or cont = dtFinal.Rows.Count Then
                    IDOperario = dr("IDOperario").ToString()
                    categoria = DevuelveIDCategoriaProfesionalSCCP(bbdd, IDOperario)
                    coste = Convert.ToDouble(dr("CosteEmpresa"))
                    Select Case categoria
                        Case 1
                            jprod = jprod + coste
                        Case 2
                            encar = encar + coste
                        Case 3
                            operar = operar + coste
                        Case 4
                            tecobra = tecobra + coste
                        Case 5
                            staff = staff + coste
                        Case Else
                            otros = otros + coste
                    End Select
                    Dim newRow As DataRow = dtResultado.NewRow()
                    newRow("Empresa") = empresa : newRow("IDCategoriaProfesionalSCCP") = 1 : newRow("Total") = jprod
                    dtResultado.Rows.Add(newRow)
                    newRow = dtResultado.NewRow() : newRow("Empresa") = empresa : newRow("IDCategoriaProfesionalSCCP") = 2 : newRow("Total") = encar
                    dtResultado.Rows.Add(newRow)
                    newRow = dtResultado.NewRow() : newRow("Empresa") = empresa : newRow("IDCategoriaProfesionalSCCP") = 3 : newRow("Total") = operar
                    dtResultado.Rows.Add(newRow)
                    newRow = dtResultado.NewRow() : newRow("Empresa") = empresa : newRow("IDCategoriaProfesionalSCCP") = 4 : newRow("Total") = tecobra
                    dtResultado.Rows.Add(newRow)
                    newRow = dtResultado.NewRow() : newRow("Empresa") = empresa : newRow("IDCategoriaProfesionalSCCP") = 5 : newRow("Total") = staff
                    dtResultado.Rows.Add(newRow)
                    newRow = dtResultado.NewRow() : newRow("Empresa") = empresa : newRow("IDCategoriaProfesionalSCCP") = 0 : newRow("Total") = otros
                    dtResultado.Rows.Add(newRow)
                    jprod = 0 : encar = 0 : operar = 0 : tecobra = 0 : staff = 0 : otros = 0
                Else
                    IDOperario = dr("IDOperario").ToString()
                    categoria = DevuelveIDCategoriaProfesionalSCCP(bbdd, IDOperario)
                    coste = Convert.ToDouble(dr("CosteEmpresa"))
                    Select Case categoria
                        Case 1
                            jprod = jprod + coste
                        Case 2
                            encar = encar + coste
                        Case 3
                            operar = operar + coste
                        Case 4
                            tecobra = tecobra + coste
                        Case 5
                            staff = staff + coste
                        Case Else
                            otros = otros + coste
                    End Select

                End If
            Catch ex As Exception
                IDOperario = dr("IDOperario").ToString()
                categoria = DevuelveIDCategoriaProfesionalSCCP(bbdd, IDOperario)
                coste = Convert.ToDouble(dr("CosteEmpresa"))
                Select Case categoria
                    Case 1
                        jprod = jprod + coste
                    Case 2
                        encar = encar + coste
                    Case 3
                        operar = operar + coste
                    Case 4
                        tecobra = tecobra + coste
                    Case 5
                        staff = staff + coste
                    Case Else
                        otros = otros + coste
                End Select
                Dim newRow As DataRow = dtResultado.NewRow()
                newRow("Empresa") = empresa : newRow("IDCategoriaProfesionalSCCP") = 1 : newRow("Total") = jprod
                dtResultado.Rows.Add(newRow)
                newRow = dtResultado.NewRow() : newRow("Empresa") = empresa : newRow("IDCategoriaProfesionalSCCP") = 2 : newRow("Total") = encar
                dtResultado.Rows.Add(newRow)
                newRow = dtResultado.NewRow() : newRow("Empresa") = empresa : newRow("IDCategoriaProfesionalSCCP") = 3 : newRow("Total") = operar
                dtResultado.Rows.Add(newRow)
                newRow = dtResultado.NewRow() : newRow("Empresa") = empresa : newRow("IDCategoriaProfesionalSCCP") = 4 : newRow("Total") = tecobra
                dtResultado.Rows.Add(newRow)
                newRow = dtResultado.NewRow() : newRow("Empresa") = empresa : newRow("IDCategoriaProfesionalSCCP") = 5 : newRow("Total") = staff
                dtResultado.Rows.Add(newRow)
                newRow = dtResultado.NewRow() : newRow("Empresa") = empresa : newRow("IDCategoriaProfesionalSCCP") = 0 : newRow("Total") = otros
                dtResultado.Rows.Add(newRow)
            End Try

            cont = cont + 1
        Next

        Return dtResultado
    End Function
    Private Sub bIDGET_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim vPersonas As New DataTable
        Dim f As New Filter
        Dim bbdd As String = ""
        vPersonas = New BE.DataEngine().Filter(DB_TECOZAM & "..vPersonasTFSD", f, , "FechaAlta asc")

        For Each dr As DataRow In vPersonas.Rows
            Dim valor As String = dr("IDOperario").ToString()
            If valor(0) = "T"c Or (Char.IsDigit(valor(0))) Then
                bbdd = DB_UK
            End If
            ActualizaOperario(bbdd, valor)
        Next
    End Sub
    Public Sub ActualizaOperario(ByVal bbdd As String, ByVal valor As String)
        Dim sql As String
        Dim IDGET As String
        IDGET = GetIDGET()
        sql = "UPDATE " & bbdd & "..tbMaestroOperario set IDGET= '" & IDGET & "' where IDOperario= '" & valor & "'"

        aux.EjecutarSql(sql)

        setIDGET()
    End Sub

    Public Function GetIDGET() As String
        Dim f As New Filter
        f.Add("IDContador", FilterOperator.Equal, "IDGET")
        Dim dt As New DataTable
        dt = New BE.DataEngine().Filter(DB_TECOZAM & "..tbMaestroContador", f)

        Dim texto As String
        texto = dt.Rows(0)("Texto")

        Dim numerico As String
        numerico = dt.Rows(0)("Contador")

        If Len(numerico) = 1 Then
            texto = texto & "0000" & numerico
        ElseIf Len(numerico) = 2 Then
            texto = texto & "000" & numerico
        ElseIf Len(numerico) = 3 Then
            texto = texto & "00" & numerico
        ElseIf Len(numerico) = 4 Then
            texto = texto & "0" & numerico
        Else
            texto = texto
        End If

        Return texto

    End Function

    Public Sub setIDGET()
        Dim f As New Filter
        f.Add("IDContador", FilterOperator.Equal, "IDGET")
        Dim dt As New DataTable
        dt = New BE.DataEngine().Filter(DB_TECOZAM & "..tbMaestroContador", f)

        Dim texto As String
        texto = dt.Rows(0)("Texto")

        Dim numerico As Integer
        numerico = dt.Rows(0)("Contador")

        numerico = numerico + 1

        Dim sql As String
        sql = "UPDATE " & DB_TECOZAM & "..tbMaestroContador set Contador= " & numerico & " Where IDContador='IDGET'"

        aux.EjecutarSql(sql)

    End Sub

    Private Sub bExportarHoras_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bExportarHoras.Click
        Dim frm As New frmInformeFecha
        frm.ShowDialog()
        Dim Fecha1 As String : Dim Fecha2 As String : Dim dt As New DataTable
        Fecha1 = frm.fecha1 : Fecha2 = frm.fecha2 : dt = ObtenerTabla(Fecha1, Fecha2)

        If frm.blEstado = False Then
            Exit Sub
        End If
        Dim mes As String : mes = Month(Fecha1)
        If Length(mes) = 1 Then
            mes = "0" & mes
        End If

        Dim anio As String
        anio = Year(Fecha1)
        GeneraExcelHoras(mes, anio, dt)
        MsgBox("Fichero generado correctamente en N:\10. AUXILIARES\00. EXPERTIS\01. HORAS.")
    End Sub
    Public Function ObtenerTabla(ByVal Fecha1 As String, ByVal Fecha2 As String) As DataTable
        Dim dt As New DataTable
        Dim f As New Filter
        f.Add("FechaInicio", FilterOperator.GreaterThanOrEqual, Fecha1)
        f.Add("FechaInicio", FilterOperator.LessThanOrEqual, Fecha2)

        dt = New BE.DataEngine().Filter(DB_TECOZAM & "..vUnionSistLabListadoTrabajadoresObraMes", f, "Empresa, IDGET, IDOperario, DescOperario, IDOficio," _
        & "IDCategoriaProfesionalSCCP, NObra, FechaInicio, MesNatural, AñoNatural, Horas, IDHora, HorasAdministrativas, HorasBaja, Turno", "FechaInicio")
        Return dt
    End Function
    Public Sub GeneraExcelHoras(ByVal mes As String, ByVal anio As String, ByVal dtFinal As DataTable)

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial

        'Dim ruta As New FileInfo("N:\10. AUXILIARES\00. EXPERTIS\02. A3\" & mes & " A3 " & mes & anio.Substring(anio.Length - 2) & ".xlsx")
        Dim ruta As New FileInfo("N:\10. AUXILIARES\00. EXPERTIS\01. HORAS\" & mes & " HORAS " & mes & anio.Substring(anio.Length - 2) & ".xlsx")
        Dim rutaCadena As String = ""
        rutaCadena = ruta.FullName

        'Verificar si el archivo existe.
        If File.Exists(rutaCadena) Then
            'Si el archivo existe, eliminarlo.
            File.Delete(rutaCadena)
        End If

        Using package As New ExcelPackage(ruta)
            ' Crear una hoja de cálculo y obtener una referencia a ella.
            Dim worksheet = package.Workbook.Worksheets.Add("HORAS")

            ' Copiar los datos de la DataTable a la hoja de cálculo.
            worksheet.Cells("A1").LoadFromDataTable(dtFinal, True)

            ' Aplicar formato negrita a la fila 1
            Dim fila1 As ExcelRange = worksheet.Cells(1, 1, 1, worksheet.Dimension.End.Column)
            fila1.Style.Font.Bold = True

            Dim columnaG As ExcelRange = worksheet.Cells("H2:H" & worksheet.Dimension.End.Row)
            columnaG.Style.Numberformat.Format = "dd/mm/yyyy"

            ' Agregar un filtro a la primera fila
            worksheet.Cells("A1:" & GetExcelColumnName(worksheet.Dimension.End.Column) & "1").AutoFilter = True
            worksheet.Column(4).Width = 30
            worksheet.Column(7).Width = 12
            worksheet.Column(8).Width = 12

            ' Congelar la primera columna
            worksheet.View.FreezePanes(2, 1)

            ' Guardar el archivo de Excel.
            package.Save()
        End Using

    End Sub
    Function GetExcelColumnName(ByVal columnNumber As Integer) As String
        Dim dividend As Integer = columnNumber
        Dim columnName As String = String.Empty
        Dim modulo As Integer

        While dividend > 0
            modulo = (dividend - 1) Mod 26
            columnName = Convert.ToChar(65 + modulo) & columnName
            dividend = CInt((dividend - modulo) / 26)
        End While

        Return columnName
    End Function

    Private Sub bDobleCotizacion_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim dt As New DataTable
        Dim filter As New Filter
        dt = New BE.DataEngine().Filter(DB_TECOZAM & "..vunionOperariodoblecotizacion", filter, , "IDGET")
        GeneraExcelDobleCoti(dt)
        MsgBox("Fichero generado correctamente en N:\10. AUXILIARES\00. EXPERTIS\03. DOBLE COTIZACION")
    End Sub
    Public Sub GeneraExcelDobleCoti(ByVal dtFinal As DataTable)

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial

        'Dim ruta As New FileInfo("N:\10. AUXILIARES\00. EXPERTIS\02. A3\" & mes & " A3 " & mes & anio.Substring(anio.Length - 2) & ".xlsx")
        Dim ruta As New FileInfo("N:\10. AUXILIARES\00. EXPERTIS\03. DOBLE COTIZACION\DOBLE COTIZACION.xlsx")
        Dim rutaCadena As String = ""
        rutaCadena = ruta.FullName

        'Verificar si el archivo existe.
        If File.Exists(rutaCadena) Then
            'Si el archivo existe, eliminarlo.
            File.Delete(rutaCadena)
        End If

        Using package As New ExcelPackage(ruta)
            ' Crear una hoja de cálculo y obtener una referencia a ella.
            Dim worksheet = package.Workbook.Worksheets.Add(" DOBLE COTIZACION ")

            ' Copiar los datos de la DataTable a la hoja de cálculo.
            worksheet.Cells("A1").LoadFromDataTable(dtFinal, True)

            ' Aplicar formato negrita a la fila 1
            Dim fila1 As ExcelRange = worksheet.Cells(1, 1, 1, worksheet.Dimension.End.Column)
            fila1.Style.Font.Bold = True

            ' Guardar el archivo de Excel.
            package.Save()
        End Using

    End Sub

    Private Sub bGetDatos_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim dtIDGet As New DataTable


        Dim CD As New OpenFileDialog()
        CD.Title = "Seleccionar archivos"
        CD.Filter = "Archivos Excel(*.xls;*.xlsx)|*.xls;*xlsx|Todos los archivos(*.*)|*.*"
        CD.ShowDialog()
        lblRuta.Text = CD.FileName

        Dim ruta As String = lblRuta.Text
        Dim hoja As String = "1"
        Dim rango As String = "A1:A10000"


        dtIDGet = ObtenerDatosExcel(ruta, hoja, rango)

        Dim dtOperarios As New DataTable
        dtOperarios = New BE.DataEngine().Filter("vUnionOperariosActivos", New Filter)

        Dim dtFinal As New DataTable
        FormaTablaFinalOperarios(dtFinal)

        For Each drIDGet As DataRow In dtIDGet.Rows
            For Each drOperario As DataRow In dtOperarios.Rows
                If drIDGet("F1").ToString = drOperario("IDGET").ToString Then
                    dtFinal.ImportRow(drOperario)
                End If
            Next
        Next

        GeneraExcelIDGET(dtFinal)
        MsgBox("Fichero generado correctamente en N:\10. AUXILIARES\00. EXPERTIS\04. IDGET\IDGET.xlsx")
    End Sub
    Public Sub FormaTablaFinalOperarios(ByRef dtFinal As DataTable)
        dtFinal.Columns.Add("IDOperario")
        dtFinal.Columns.Add("DescOperario")
        dtFinal.Columns.Add("Fecha_Baja")
        dtFinal.Columns.Add("IDGET")
        dtFinal.Columns.Add("Diccionario")
        dtFinal.Columns.Add("Empresa")
        dtFinal.Columns.Add("IDOficio")
    End Sub
    Public Sub FormaTablaGenteBaja(ByRef dtPersonasDeBaja As DataTable)
        dtPersonasDeBaja.Columns.Add("Empresa")
        dtPersonasDeBaja.Columns.Add("IDOperario")
        dtPersonasDeBaja.Columns.Add("Fecha_Baja")
        dtPersonasDeBaja.Columns.Add("Fecha_Alta")
        dtPersonasDeBaja.Columns.Add("nDias")
        dtPersonasDeBaja.Columns.Add("IDObra")
    End Sub

    Public Sub GeneraExcelIDGET(ByVal dtFinal As DataTable)
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial
        Dim ruta As New FileInfo("N:\10. AUXILIARES\00. EXPERTIS\04. IDGET\IDGET.xlsx")
        'Dim ruta As New FileInfo("N:\01. A3\" & mes & " A3 " & mes & anio.Substring(anio.Length - 2) & ".xlsx")
        Dim rutaCadena As String = ""
        rutaCadena = ruta.FullName

        'Verificar si el archivo existe.
        If File.Exists(rutaCadena) Then
            'Si el archivo existe, eliminarlo.
            File.Delete(rutaCadena)
        End If

        Using package As New ExcelPackage(ruta)
            ' Crear una hoja de cálculo y obtener una referencia a ella.
            Dim worksheet = package.Workbook.Worksheets.Add("Listado")

            ' Copiar los datos de la DataTable a la hoja de cálculo.
            worksheet.Cells("A1").LoadFromDataTable(dtFinal, True)
            ' Aplicar formato negrita a la fila 1
            Dim fila1 As ExcelRange = worksheet.Cells(1, 1, 1, worksheet.Dimension.End.Column)
            fila1.Style.Font.Bold = True

            ' Guardar el archivo de Excel.
            package.Save()
        End Using

    End Sub

    Private Sub bCrearHorasBaja_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bCrearHorasBaja.Click
        Dim frm As New frmInformeFecha
        frm.ShowDialog()
        Dim Fecha1 As String : Dim Fecha2 As String
        Fecha1 = frm.fecha1 : Fecha2 = frm.fecha2

        Dim mes As String : mes = Month(Fecha1)
        If Length(mes) = 1 Then
            mes = "0" & mes
        End If

        Dim anio As String
        anio = Year(Fecha1)

        If frm.blEstado = True Then
            'Horas Baja por Accidentes en España
            HorasBajaEspaña(Fecha1, Fecha2)
        End If
    End Sub

    Public Sub HorasBajaEspaña(ByVal Fecha1 As String, ByVal Fecha2 As String)
        'David Velasco 17/10/23

        '1. SACO LISTADO DE PERSONAS POR ACCIDENTE
        Dim dtPersonasBajaPorAccidente As New DataTable
        dtPersonasBajaPorAccidente = GetListadoPersonasDeBajaPorAccidente(Fecha1, Fecha2)

        '2. SACO LISTADO DE PERSONAS POR ENFERMEDAD
        Dim dtPersonasBajaPorEnfermedad As New DataTable
        dtPersonasBajaPorEnfermedad = GetListadoPersonasDeBajaPorEnfermedad(Fecha1, Fecha2)

        Dim dtPersonasDeBaja As New DataTable
        FormaTablaGenteBaja(dtPersonasDeBaja)

        UnirTablas(dtPersonasDeBaja, dtPersonasBajaPorAccidente, dtPersonasBajaPorEnfermedad)
        Dim result As DialogResult = MessageBox.Show("¿Desea introducir horas de bajas?", "¿Desea Continuar?", MessageBoxButtons.YesNo)
        If result = DialogResult.No Then
            Exit Sub
        End If
        CrearHorasBaja(dtPersonasDeBaja, Fecha1, Fecha2)

        MsgBox("Horas de baja creadas correctamente.")
        'MsgBox("Accidente: " & dtPersonasBajaPorAccidente.Rows.Count & " Enfermedad:" & dtPersonasBajaPorEnfermedad.Rows.Count)
        'MsgBox(dtPersonasDeBaja.Rows.Count)
    End Sub
    Public Sub CrearHorasBaja(ByVal dtPersonasDeBaja As DataTable, ByVal Fecha1 As DateTime, ByVal Fecha2 As DateTime)
        '1. Miro la empresa en la que está
        '2. Si pasa el día 60 pasa al centro de coste BAJAS
        '2. Si ya tiene alguna hora productiva ese dia no inserta nada
        '   Si tiene ACC o CC actualizo las horas baja a 8
        '3. Creo el registro IDHora="HB"
        Dim bbdd As String
        Dim idoperario As String
        Dim fechabaja As DateTime
        Dim fechaalta As DateTime
        Dim fechaCalculos As DateTime
        Dim fechaFin As DateTime
        Dim idobra As String
        Dim dias As Integer = 0

        Dim filas As Integer = 0
        PvProgreso.Value = 0 : PvProgreso.Maximum = dtPersonasDeBaja.Rows.Count : PvProgreso.Step = 1 : PvProgreso.Visible = True

        'FESTIVOS Y FINDES = 1
        Dim dtFestivos As New DataTable
        Dim f As New Filter
        f.Add("TipoDia", FilterOperator.Equal, 1)
        f.Add("IDCentro", FilterOperator.Equal, "00")
        f.Add("Fecha", FilterOperator.GreaterThanOrEqual, Fecha1)
        f.Add("Fecha", FilterOperator.LessThanOrEqual, Fecha2)
        dtFestivos = New BE.DataEngine().Filter(DB_TECOZAM & "..tbCalendarioCentro", f, "Fecha, TipoDia")

        For Each dr As DataRow In dtPersonasDeBaja.Rows
            'If dr("idoperario") <> "T3249" Then
            'Continue For ' Esto pasará a la siguiente iteración
            'End If
            bbdd = dr("Empresa") : idoperario = dr("idoperario") : fechabaja = dr("Fecha_Baja") : fechaalta = Nz(dr("Fecha_Alta"), Fecha2)
            fechaCalculos = Fecha1 : idobra = Nz(dr("IDObra").ToString, getObraBaja(bbdd, idoperario, fechabaja))

            Windows.Forms.Application.DoEvents()
            LProgreso.Text = "Importando : " & idoperario & " - HORAS DE BAJA"
            Windows.Forms.Application.DoEvents()
            'Para los que están de baja en el intervalo de fechas, por ejemplo 15/09/23
            If fechabaja > Fecha1 Then
                fechaCalculos = fechabaja
            End If
            'Para los que se dan de alta en el intervalo de fechas, por ejemplo 15/09/23
            If fechaalta < Fecha2 Then
                fechaFin = fechaalta
            Else
                fechaFin = Fecha2
            End If

            While fechaCalculos <= fechaFin
                If fechaalta = fechaCalculos Then
                    Exit While
                End If
                'Si es festivo pasa al siguiente
                For Each fila As DataRow In dtFestivos.Rows
                    If fila("Fecha") = fechaCalculos Then
                        fechaCalculos = fechaCalculos.AddDays(1)
                        Continue While
                    End If
                Next
                dias = (fechaCalculos - fechabaja).Days
                'Se insertan las horas en la ultima obra donde tuvo horas antes de la fecha1
                'Si ya han pasado 60 dias desde la baja pasa a un centro de coste que se llama BAJAS
                insertarOActualizar(bbdd, idoperario, idobra, dias, fechaCalculos)
                fechaCalculos = fechaCalculos.AddDays(1)
            End While
            filas = filas + 1
            PvProgreso.Value = filas
        Next
    End Sub
    Public Sub insertarOActualizar(ByVal bbdd As String, ByVal idoperario As String, ByVal idobra As String, ByVal dias As Integer, ByVal fechaInicio As String)
        'Asigno bases de datos
        If bbdd = "T. ES." Then
            bbdd = DB_TECOZAM
        ElseIf bbdd = "FERR." Then
            bbdd = DB_FERRALLAS
        ElseIf bbdd = "SEC." Then
            bbdd = DB_SECOZAM
        End If
        'Asigno el idobra
        If dias > 60 Then
            If bbdd = DB_TECOZAM Then
                idobra = "17152171"
            ElseIf bbdd = DB_FERRALLAS Then
                idobra = "12712406"
            ElseIf bbdd = DB_SECOZAM Then
                idobra = "11863745"
            End If
        Else

        End If

        'CHECKEO QUE NO HAYA REGISTRO EN ESTA FECHA PARA ESTE OPERARIO. 
        'SI ES CC O ACC QUE ACTUALICE HORASBAJA A 8
        'SI NO METO 8 HORASBAJAS CON IDHORA=HB

        Dim dtCheck As New DataTable
        Dim filtro As New Filter
        filtro.Add("FechaInicio", FilterOperator.Equal, fechaInicio)
        filtro.Add("IDOperario", FilterOperator.Equal, idoperario)
        dtCheck = New BE.DataEngine().Filter(bbdd & "..tbObraModControl", filtro)
        If dtCheck.Rows.Count = 0 Then
            'INSERTO HORAS ESE DIA
            insertarHorasBajas(bbdd, idoperario, idobra, fechaInicio)
        Else
            For Each dr As DataRow In dtCheck.Rows
                If dr("IDHora") = "ACC" Or dr("IDHora") = "CC" Or dr("IDHora") = "acc" Or dr("IDHora") = "cc" Then
                    'Actualizo horas baja a 8
                    actualizaHorasBajas(bbdd, idobra, idoperario, dtCheck.Rows(0)("IDLineaModControl").ToString)
                End If
            Next
        End If
    End Sub
    Public Sub actualizaHorasBajas(ByVal bbdd As String, ByVal idobra As String, ByVal idoperario As String, ByVal IDLineaModControl As String)
        Dim IDOficio As String
        Dim IDCategoriaProfesionalSCCP As String
        Dim IDTrabajo As String
        Dim txtSQL As String

        IDOficio = DevuelveIDOficio(bbdd, idoperario)
        IDCategoriaProfesionalSCCP = DevuelveIDCategoriaProfesionalSCCP(bbdd, idoperario)
        IDTrabajo = ObtieneIDTrabajo(bbdd, idobra, "PT1")
        txtSQL = "UPDATE " & bbdd & "..tbObraMODControl "
        txtSQL &= "SET IDObra=" & idobra & ", IDTrabajo= " & IDTrabajo & ","
        txtSQL &= "IDOficio= '" & IDOficio & "', IDCategoriaProfesionalSCCP= " & IDCategoriaProfesionalSCCP & ", HorasBaja=8 "
        txtSQL &= "WHERE IDLineaModControl = " & IDLineaModControl

        auto.Ejecutar(txtSQL)
    End Sub
    Public Sub insertarHorasBajas(ByVal bbdd As String, ByVal idoperario As String, ByVal idobra As String, ByVal fechaInicio As String)
        Dim IDTrabajo As String
        Dim CodTrabajo As String
        Dim txtSQL As String
        Dim IDAutonumerico = auto.Autonumerico()
        Dim IDOficio As String
        Dim IDCategoriaProfesionalSCCP As String

        IDOficio = DevuelveIDOficio(bbdd, idoperario)
        IDCategoriaProfesionalSCCP = DevuelveIDCategoriaProfesionalSCCP(bbdd, idoperario)

        IDTrabajo = ObtieneIDTrabajo(bbdd, idobra, "PT1")
        Dim rsTrabajo As New DataTable
        Dim filtro2 As New Filter
        filtro2.Add("IDObra", FilterOperator.Equal, idobra)
        filtro2.Add("IdTrabajo", FilterOperator.Equal, IDTrabajo)
        rsTrabajo = New BE.DataEngine().Filter(bbdd & "..tbObraTrabajo", filtro2)

        IDTrabajo = rsTrabajo.Rows(0)("IdTrabajo")
        CodTrabajo = rsTrabajo.Rows(0)("CodTrabajo")

        Dim DescTrabajo As String = "" : Dim IdTipoTrabajo As String = "" : Dim IdSubTipoTrabajo As String = ""
        DescTrabajo = rsTrabajo.Rows(0)("DescTrabajo") : IdTipoTrabajo = rsTrabajo.Rows(0)("IdTipoTrabajo") : IdSubTipoTrabajo = Nz(rsTrabajo.Rows(0)("IdSubtipotrabajo"), "")
        Dim DescParte As String : DescParte = "HORAS BAJA" & " " & Month(fechaInicio) & "-" & Year(fechaInicio) & "-HB"

        txtSQL = "Insert into " & bbdd & "..tbObraMODControl(IdLineaModControl, IdTrabajo, IdObra, CodTrabajo, DescTrabajo, IdTipoTrabajo, " & _
                "IdSubTipoTrabajo, IdOperario, IdCategoria, IdHora, FechaInicio, HorasRealMod, TasaRealModA, " & _
                 "ImpRealModA, HorasFactMod, ImpFactModA, DescParte, Facturable, FechaCreacionAudi, FechaModificacionAudi, Usuarioaudi, IDOficio, IdTipoTurno, HorasAdministrativas, IDCategoriaProfesionalSCCP, HorasBaja) " & _
                 "Values(" & IDAutonumerico & ", " & IDTrabajo & ", " & idobra & ", '" & _
                 CodTrabajo & "', '" & DescTrabajo & "', '" & IdTipoTrabajo & "', '" & _
                 IdSubTipoTrabajo & "', '" & idoperario & "', 'PREDET', '" & _
                 "HB" & "', '" & fechaInicio & "', 0 , " & 0 & ", " & 0 & _
                 ", 0 , " & 0 & _
                 ", '" & DescParte & "', " & 0 & ", '" & Date.Now.Date & "', '" & Date.Now.Date & "', '" & ExpertisApp.UserName & "','" & IDOficio & "', 4, 0 ," & Nz(IDCategoriaProfesionalSCCP, "") & ", 8)"

        auto.Ejecutar(txtSQL)
    End Sub
    Public Function getObraBaja(ByVal bbdd As String, ByVal idoperario As String, ByVal fechaBaja As String) As String
        Dim dtHoras As New DataTable
        Dim f As New Filter
        f.Add("IDOperario", FilterOperator.Equal, idoperario)
        f.Add("FechaInicio", FilterOperator.Equal, fechaBaja)
        If bbdd = "T. ES." Then
            bbdd = DB_TECOZAM
        ElseIf bbdd = "FERR." Then
            bbdd = DB_FERRALLAS
        ElseIf bbdd = "SEC." Then
            bbdd = DB_SECOZAM
        End If
        dtHoras = New BE.DataEngine().Filter(bbdd & "..tbObraModControl", f, , "FechaInicio desc")
        '---SI NO EXISTE PUES VA A BAJAS DIRECTAMENTE
        Dim dtObrasBaja As New DataTable
        Dim fil As New Filter
        fil.Add("NObra", FilterOperator.Equal, "BAJAS")
        dtObrasBaja = New BE.DataEngine().Filter(bbdd & "..tbObraCabecera", fil)

        If dtHoras.Rows.Count = 0 Then
            Return dtObrasBaja.Rows(0)("IDObra").ToString
        Else
            Return dtHoras.Rows(0)("IDObra").ToString
        End If
    End Function
    Public Sub UnirTablas(ByRef dtPersonasDeBaja As DataTable, ByVal dtPersonasBajaPorAccidente As DataTable, ByVal dtPersonasBajaPorEnfermedad As DataTable)
        For Each fila As DataRow In dtPersonasBajaPorAccidente.Rows
            Dim dr As DataRow
            dr = dtPersonasDeBaja.NewRow()
            dr("Empresa") = fila("Empresa")
            dr("IDOperario") = fila("IDOperario")
            dr("Fecha_Baja") = fila("fBaja")
            dr("Fecha_Alta") = fila("fAlta")
            dr("nDias") = fila("nDiasBaja")
            dr("IDObra") = fila("CodObra")
            dtPersonasDeBaja.Rows.Add(dr)
        Next

        For Each fila As DataRow In dtPersonasBajaPorEnfermedad.Rows
            Dim dr As DataRow
            dr = dtPersonasDeBaja.NewRow()
            dr("Empresa") = fila("Empresa")
            dr("IDOperario") = fila("IDOperario")
            dr("Fecha_Baja") = fila("fBaja")
            dr("Fecha_Alta") = fila("fAlta")
            dr("nDias") = fila("nDias")
            dtPersonasDeBaja.Rows.Add(dr)
        Next

    End Sub
    Public Function GetListadoPersonasDeBajaPorAccidente(ByVal Fecha1 As String, ByVal Fecha2 As String) As DataTable
        '1. SACO LAS PERSONAS QUE SE HAN DADO DE ALTA EN UN MES
        Dim filtro As New Filter
        Dim dtPersonas As New DataTable
        Dim sql As String

        sql = "select * from xTecozam50R2..vUnionOperariosAccidentes"
        sql &= " where ((fAlta >= '" & Fecha1 & "' AND fBaja >= '" & Fecha1 & "' AND fBaja <= '" & Fecha2 & "')"
        sql &= " or(fAlta >= '" & Fecha1 & "' AND fAlta <= '" & Fecha2 & "') or (fBaja<='" & Fecha1 & "' and fAlta>='" & Fecha2 & "') or fAlta is null)"
        sql &= " and fBaja is not null and nDiasBaja!=0"

        dtPersonas = aux.EjecutarSqlSelect(sql)
        Return dtPersonas
    End Function

    Public Function GetListadoPersonasDeBajaPorEnfermedad(ByVal Fecha1 As String, ByVal Fecha2 As String) As DataTable
        '1. SACO LAS PERSONAS QUE SE HAN DADO DE ALTA EN UN MES
        Dim filtro As New Filter
        Dim dtPersonas As New DataTable
        Dim sql As String

        sql = "select * from xTecozam50R2..vUnionOperariosEnfermedadesCC"
        sql &= " where ((fAlta >= '" & Fecha1 & "' AND fBaja >= '" & Fecha1 & "' AND fBaja <= '" & Fecha2 & "')"
        sql &= " or(fAlta >= '" & Fecha1 & "' AND fAlta <= '" & Fecha2 & "') or (fBaja<='" & Fecha1 & "' and fAlta>='" & Fecha2 & "') or fAlta is null)"
        sql &= " and fBaja is not null and nDias!=0"

        dtPersonas = aux.EjecutarSqlSelect(sql)
        Return dtPersonas
    End Function
    Private Sub bMixA3Horas_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bMixA3Horas.Click
        'DAVID VELASCO 10/10/23
        '----------------------
        'GENERO UN EXCEL CON 4 PESTAÑAS
        '----------------------
        '1 CON RATIO €/HORA CON CATEGORIAPROFESIONALSCCP
        '1 CON GENTE QUE TENGA HORAS QUE NO TENGA €
        '1 CON GENTE QUE TENGA € PERO NO TENGA HORAS
        '1 CON PERSONAS CON DOBLE COTIZACIÓN
        Dim mes As String
        Dim anio As String
        Dim frmFechas As New frmFechas
        frmFechas.ShowDialog()

        mes = frmFechas.mes
        anio = frmFechas.anio

        'Tabla de las personas con horas
        Dim dtHorasExpertis As New DataTable
        dtHorasExpertis = getHorasPersonas(mes, anio)

        'Tabla de las personas con €
        Dim dtA3 As New DataTable
        dtA3 = getEurosPersonas(mes, anio)

        '1. TABLA DE GENTE QUE TIENE HORAS TRABAJADAS Y NO TIENE EUROS
        Dim dtGenteSiHorasNoEuros As New DataTable
        dtGenteSiHorasNoEuros = getGenteSiHorasNoEuros(dtHorasExpertis, dtA3)

        '2. TABLA DE GENTE QUE TIENE € Y NO TIENE HORAS
        Dim dtGenteSiEurosNoHoras As New DataTable
        dtGenteSiEurosNoHoras = getGenteSiEurosNoHoras(dtHorasExpertis, dtA3)

        '3. TABLA DE RATIOS DE LA GENTE
        Dim dtRatiosGente As New DataTable
        dtRatiosGente = getRatiosGente(dtHorasExpertis, dtA3)

        '4. TABLA DE PERSONAS CON DOBLE COTIZACION
        Dim dtPersonasDobleCoti As New DataTable
        Dim filter As New Filter
        dtPersonasDobleCoti = New BE.DataEngine().Filter(DB_TECOZAM & "..vunionOperariodoblecotizacion", filter, , "IDGET")

        '5. TABLA DE RESUMEN
        Dim dtResumenCategoriaProfesional As New DataTable
        dtResumenCategoriaProfesional = getResumenCategoria(dtRatiosGente)



        'GENERACION EXCEL CON LAS 5 PESTAÑAS
        GeneraExcelHorasA3(dtGenteSiHorasNoEuros, dtGenteSiEurosNoHoras, dtRatiosGente, dtPersonasDobleCoti, dtResumenCategoriaProfesional, mes, anio)
    End Sub
    Public Function getResumenCategoria(ByVal dtRatiosGente As DataTable) As DataTable
        Dim dtResultado As New DataTable()
        dtResultado.Columns.Add("CategoriaProfesional", GetType(String))
        dtResultado.Columns.Add("CosteTotal", GetType(Double))

        Dim jprod As Double = 0
        Dim encar As Double = 0
        Dim operar As Double = 0
        Dim tecobra As Double = 0
        Dim staff As Double = 0
        Dim otros As Double = 0

        For Each dr As DataRow In dtRatiosGente.Rows
            Dim categoria As Double = Convert.ToDouble(dr("IDCategoriaProfesionalSCCP"))
            Dim coste As Double = Convert.ToDouble(dr("EurosTotales"))
            Select Case categoria
                Case 1
                    jprod = jprod + coste
                Case 2
                    encar = encar + coste
                Case 3
                    operar = operar + coste
                Case 4
                    tecobra = tecobra + coste
                Case 5
                    staff = staff + coste
                Case Else
                    otros = otros + coste
            End Select
        Next
        '-1.jefes de produccion
        Dim newRow As DataRow = dtResultado.NewRow()
        newRow("CategoriaProfesional") = 1
        newRow("CosteTotal") = jprod
        dtResultado.Rows.Add(newRow)

        '-2. Encargados
        newRow = dtResultado.NewRow()
        newRow("CategoriaProfesional") = 2
        newRow("CosteTotal") = encar
        dtResultado.Rows.Add(newRow)

        '-3.operarios
        newRow = dtResultado.NewRow()
        newRow("CategoriaProfesional") = 3
        newRow("CosteTotal") = operar
        dtResultado.Rows.Add(newRow)

        '-4. tecnicos de obra
        newRow = dtResultado.NewRow()
        newRow("CategoriaProfesional") = 4
        newRow("CosteTotal") = tecobra
        dtResultado.Rows.Add(newRow)

        '-5. staff
        newRow = dtResultado.NewRow()
        newRow("CategoriaProfesional") = 5
        newRow("CosteTotal") = staff
        dtResultado.Rows.Add(newRow)

        '-6. otros
        newRow = dtResultado.NewRow()
        newRow("CategoriaProfesional") = 0
        newRow("CosteTotal") = otros
        dtResultado.Rows.Add(newRow)
        Return dtResultado
    End Function

    Public Sub GeneraExcelHorasA3(ByVal dtGenteSiHorasNoEuros As DataTable, ByVal dtGenteSiEurosNoHoras As DataTable, ByVal dtRatiosGente As DataTable, _
                                  ByVal dtPersonasDobleCoti As DataTable, ByVal dtResumenCategoriaProfesional As DataTable, ByVal mes As String, ByVal anio As String)

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial

        Dim ruta As New FileInfo("N:\10. AUXILIARES\00. EXPERTIS\05. CHECK HORAS-A3\" & mes & " CHECK " & mes & anio.Substring(anio.Length - 2) & ".xlsx")
        'Dim ruta As New FileInfo("N:\01. A3\" & mes & " A3 " & mes & anio.Substring(anio.Length - 2) & ".xlsx")
        Dim rutaCadena As String = ""
        rutaCadena = ruta.FullName

        'Verificar si el archivo existe.
        If File.Exists(rutaCadena) Then
            'Si el archivo existe, eliminarlo.
            File.Delete(rutaCadena)
        End If

        Using package As New ExcelPackage(ruta)
            ' HOJA 1
            Dim worksheet = package.Workbook.Worksheets.Add(mes & " SI HORAS/NO € " & anio)
            worksheet.Cells("A1").LoadFromDataTable(dtGenteSiHorasNoEuros, True)
            worksheet.Column(2).Width = 30
            Dim fila1 As ExcelRange = worksheet.Cells(1, 1, 1, worksheet.Dimension.End.Column)
            fila1.Style.Font.Bold = True
            worksheet.Cells("A1:" & GetExcelColumnName(worksheet.Dimension.End.Column) & "1").AutoFilter = True
            ' HOJA 2
            Dim worksheet2 = package.Workbook.Worksheets.Add(mes & " SI €/NO HORAS " & anio)
            worksheet2.Cells("A1").LoadFromDataTable(dtGenteSiEurosNoHoras, True)
            Dim fila12 As ExcelRange = worksheet2.Cells(1, 1, 1, worksheet2.Dimension.End.Column)
            fila12.Style.Font.Bold = True
            worksheet2.Cells("A1:" & GetExcelColumnName(worksheet2.Dimension.End.Column) & "1").AutoFilter = True
            worksheet2.Column(2).Width = 30
            ' HOJA 3
            Dim worksheet3 = package.Workbook.Worksheets.Add(mes & " RATIOS " & anio)

            ' Ordenar el DataTable por la columna L
            dtRatiosGente.DefaultView.Sort = "Ratio ASC" ' Ajusta "ColumnaL" al nombre real de la columna
            Dim dtRatiosOrdenado = dtRatiosGente.DefaultView.ToTable()

            worksheet3.Cells("A1").LoadFromDataTable(dtRatiosOrdenado, True)

            Dim fila13 As ExcelRange = worksheet3.Cells(1, 1, 1, worksheet3.Dimension.End.Column)
            fila13.Style.Font.Bold = True

            Dim columnaE As ExcelRange = worksheet3.Cells("E2:E" & worksheet3.Dimension.End.Row)
            columnaE.Style.Numberformat.Format = "dd/mm/yyyy"

            Dim columnaF As ExcelRange = worksheet3.Cells("F2:F" & worksheet3.Dimension.End.Row)
            columnaF.Style.Numberformat.Format = "dd/mm/yyyy"

            Dim columnaK As ExcelRange = worksheet3.Cells("K2:K" & worksheet3.Dimension.End.Row)
            columnaK.Style.Numberformat.Format = "#,##0.00€"

            Dim columnaL As ExcelRange = worksheet3.Cells("L2:L" & worksheet3.Dimension.End.Row)
            columnaL.Style.Numberformat.Format = "#,##0.00€"

            worksheet3.Cells("A1:" & GetExcelColumnName(worksheet3.Dimension.End.Column) & "1").AutoFilter = True
            worksheet3.Column(3).Width = 30
            worksheet3.Column(5).Width = 15


            ' HOJA 4
            Dim worksheet4 = package.Workbook.Worksheets.Add(mes & " DOBLE COTIZACION " & anio)
            worksheet4.Cells("A1").LoadFromDataTable(dtPersonasDobleCoti, True)
            Dim fila14 As ExcelRange = worksheet4.Cells(1, 1, 1, worksheet4.Dimension.End.Column)
            fila14.Style.Font.Bold = True
            worksheet4.Cells("A1:" & GetExcelColumnName(worksheet4.Dimension.End.Column) & "1").AutoFilter = True
            worksheet4.Column(2).Width = 30
            ' HOJA 5
            Dim worksheet5 = package.Workbook.Worksheets.Add(mes & " RESUMEN POR CATEGORIA " & anio)
            worksheet5.Cells("A1").LoadFromDataTable(dtResumenCategoriaProfesional, True)
            Dim fila15 As ExcelRange = worksheet5.Cells(1, 1, 1, worksheet5.Dimension.End.Column)
            fila15.Style.Font.Bold = True

            Dim columnaBResumen As ExcelRange = worksheet5.Cells("B2:B" & worksheet5.Dimension.End.Row)
            columnaBResumen.Style.Numberformat.Format = "#,##0.00€"
            worksheet5.Cells("A1:" & GetExcelColumnName(worksheet5.Dimension.End.Column) & "1").AutoFilter = True
            worksheet.Column(2).Width = 15
            ' Guardar el archivo de Excel.
            package.Save()

            MsgBox("Fichero guardado en N:\10. AUXILIARES\00. EXPERTIS\05. CHECK HORAS-A3\")
        End Using
    End Sub

    Public Function getRatiosGente(ByVal dtHorasExpertis As DataTable, ByVal dtA3 As DataTable) As DataTable
        ' Crear una nueva tabla para el resultado
        Dim dtResultado As New DataTable()
        dtResultado.Columns.Add("IDGET", GetType(String))
        dtResultado.Columns.Add("IDOperario", GetType(String))
        dtResultado.Columns.Add("DescOperario", GetType(String))
        dtResultado.Columns.Add("IDCategoriaProfesionalSCCP", GetType(String))
        dtResultado.Columns.Add("Fecha_Alta", GetType(String))
        dtResultado.Columns.Add("Fecha_Baja", GetType(String))
        dtResultado.Columns.Add("HorasProductivas", GetType(Double))
        dtResultado.Columns.Add("HorasAdministrativas", GetType(Double))
        dtResultado.Columns.Add("HorasBaja", GetType(Double))
        dtResultado.Columns.Add("HorasTotales", GetType(Double))
        dtResultado.Columns.Add("EurosTotales", GetType(Double))
        dtResultado.Columns.Add("Ratio", GetType(Double))

        'DAVID VELASCO 21/12/2023
        'Este for lo que hace es unir las nominas de las personas que tengan mas de una linea
        ' Crear una nueva tabla para almacenar los resultados
        Dim dtResult As New DataTable()
        dtResult.Columns.Add("IDOperario", GetType(String))
        dtResult.Columns.Add("IDGet", GetType(String))
        dtResult.Columns.Add("DescOperario", GetType(String))
        dtResult.Columns.Add("TotalCosteEmpresa", GetType(Double))

        ' Iterar sobre las filas de la tabla original
        For Each row As DataRow In dtA3.Rows
            Dim currentIDOperario As String = CStr(row("IDOperario"))

            ' Buscar si ya existe una fila con el mismo IDOperario en la tabla de resultados
            Dim foundRow As DataRow = dtResult.Select("IDOperario = '" & currentIDOperario & "'").FirstOrDefault()

            If foundRow IsNot Nothing Then
                ' Si existe, sumar el CosteEmpresa
                foundRow("TotalCosteEmpresa") = CDbl(foundRow("TotalCosteEmpresa")) + CDbl(row("CosteEmpresa"))
            Else
                ' Si no existe, agregar una nueva fila
                Dim newRow As DataRow = dtResult.NewRow()
                newRow("IDOperario") = currentIDOperario
                newRow("IDGet") = CStr(row("IDGet"))
                newRow("DescOperario") = CStr(row("DescOperario"))
                newRow("TotalCosteEmpresa") = CDbl(row("CosteEmpresa"))
                dtResult.Rows.Add(newRow)
            End If
        Next

        ' Limpiar la tabla original
        dtA3.Rows.Clear()

        ' Copiar los resultados de la nueva tabla a la tabla original sin duplicados
        Dim uniqueRows As New List(Of DataRow)()
        For Each row As DataRow In dtResult.Rows
            ' Buscar las filas originales que coinciden con el IDOperario
            Dim originalRows As DataRow() = dtA3.Select("IDOperario = '" & row("IDOperario") & "'")

            ' Agregar solo una fila única a la tabla original
            If originalRows.Length = 0 Then
                Dim newRow As DataRow = dtA3.NewRow()
                newRow("IDOperario") = row("IDOperario")
                newRow("IDGet") = row("IDGet")
                newRow("DescOperario") = row("DescOperario")
                newRow("CosteEmpresa") = row("TotalCosteEmpresa")
                dtA3.Rows.Add(newRow)
            End If
        Next
        ' Recorrer las filas de dtHorasExpertis
        For Each rowHorasExpertis As DataRow In dtHorasExpertis.Rows

            Dim idGet As String = rowHorasExpertis.Field(Of String)("IDGET")
            Dim idOperario As String = rowHorasExpertis.Field(Of String)("IDOperario")
            Dim descOperario As String = rowHorasExpertis.Field(Of String)("DescOperario")
            Dim IDCategoriaProfesionalSCCP As String = rowHorasExpertis.Field(Of String)("IDCategoriaProfesionalSCCP")
            Dim horas As Double = rowHorasExpertis.Field(Of String)("Horas")
            Dim horasAdministrativas As Double = rowHorasExpertis.Field(Of String)("HorasAdministrativas")
            Dim horasBaja As Double = rowHorasExpertis.Field(Of String)("HorasBaja")
            Dim fechaAlta As String = Convert.ToDateTime(rowHorasExpertis("Fecha_Alta")).ToString("dd/MM/yyyy")
            Dim fechaBaja As String
            If Len(rowHorasExpertis("Fecha_Baja")).ToString <> 0 Then
                fechaBaja = Convert.ToDateTime(rowHorasExpertis("Fecha_Baja")).ToString("dd/MM/yyyy")
            Else
                fechaBaja = String.Empty
            End If

            ' Buscar una fila correspondiente en dtA3s
            Dim rowA3 As DataRow = dtA3.Rows.Cast(Of DataRow)().FirstOrDefault(Function(row) row.Field(Of String)("IDGET") = idGet)

            If rowA3 IsNot Nothing Then
                Dim costeEmpresa As Double = rowA3.Field(Of String)("CosteEmpresa")
                Dim horasTotales As Double
                horasTotales = (horas + horasAdministrativas + horasBaja)
                Dim ratio As Double = Math.Round((costeEmpresa / horasTotales), 2)

                ' Agregar los resultados a dtResultado
                Dim newRow As DataRow = dtResultado.NewRow()
                newRow("IDGET") = idGet
                newRow("IDOperario") = idOperario
                newRow("DescOperario") = descOperario
                newRow("IDCategoriaProfesionalSCCP") = IDCategoriaProfesionalSCCP
                newRow("Fecha_Alta") = fechaAlta
                newRow("Fecha_Baja") = fechaBaja
                newRow("HorasProductivas") = horas
                newRow("HorasAdministrativas") = horasAdministrativas
                newRow("HorasBaja") = horasBaja
                newRow("HorasTotales") = horasTotales
                newRow("EurosTotales") = costeEmpresa
                If horasTotales = 0 Then
                    newRow("Ratio") = 0
                Else
                    newRow("Ratio") = ratio
                End If
                dtResultado.Rows.Add(newRow)
                'Aqui van las personas de la 2ª pestaña. Para tener a todos sí o sí.
            End If
        Next

        '2ª PARTE. Que me salgan las personas de la segunda pestaña
        For Each rowA3 As DataRow In dtA3.Rows

            Dim idOperarioA3 As String = Convert.ToString(rowA3("IDOperario"))
            Dim costeEmpresa As Double = rowA3.Field(Of String)("CosteEmpresa")
            Dim idGet As String = rowA3.Field(Of String)("IDGET")
            Dim descOperario As String = rowA3.Field(Of String)("DescOperario")

            ' Verificar si el IDOperario existe en dtResultado
            Dim found As Boolean = False
            For Each rowResultado As DataRow In dtResultado.Rows
                Dim idOperarioResultado As String = Convert.ToString(rowResultado("IDOperario"))
                If idOperarioA3 = idOperarioResultado Then
                    found = True
                    Exit For
                End If
            Next

            If Not found Then
                Dim newRow As DataRow = dtResultado.NewRow()
                newRow("IDGET") = idGet
                newRow("IDOperario") = idOperarioA3
                newRow("DescOperario") = descOperario
                newRow("IDCategoriaProfesionalSCCP") = DevuelveIDCategoriaProfesionalSCCPTodasBasesDeDatos(idOperarioA3)
                newRow("Fecha_Alta") = devuelveFechaAlta(idGet)
                newRow("Fecha_Baja") = devuelveFechaBaja(idGet)
                newRow("HorasProductivas") = 0
                newRow("HorasAdministrativas") = 0
                newRow("HorasBaja") = 0
                newRow("HorasTotales") = 0
                newRow("EurosTotales") = costeEmpresa
                newRow("Ratio") = costeEmpresa
                dtResultado.Rows.Add(newRow)

            End If
        Next


        Return dtResultado
    End Function

    Public Function getGenteSiEurosNoHoras(ByVal dtHorasExpertis As DataTable, ByVal dtA3 As DataTable) As DataTable
        ' Obtén una lista de los IDGet que están en dtHorasExpertis
        Dim idGetEnHorasExpertis = dtHorasExpertis.AsEnumerable().Select(Function(row) row.Field(Of String)("IDGet")).ToList()

        ' Usa LINQ para encontrar las filas en dtA3 que no están en dtHorasExpertis
        Dim filasFaltantes = dtA3.AsEnumerable().Where(Function(row) Not idGetEnHorasExpertis.Contains(row.Field(Of String)("IDGet")))

        ' Crea una nueva tabla con los resultados
        Dim dtFaltantes As New DataTable()
        dtFaltantes.Columns.Add("IDGet", GetType(String))
        dtFaltantes.Columns.Add("DescOperario", GetType(String))
        dtFaltantes.Columns.Add("IDOperario", GetType(String))
        dtFaltantes.Columns.Add("IDCategoriaProfesionalSCCP", GetType(String))

        ' Agrega las filas faltantes a la nueva tabla
        For Each fila In filasFaltantes
            Dim newRow As DataRow = dtFaltantes.NewRow()
            newRow("IDGet") = fila.Field(Of String)("IDGet")
            newRow("DescOperario") = fila.Field(Of String)("DescOperario")
            newRow("IDOperario") = fila.Field(Of String)("IDOperario")
            newRow("IDCategoriaProfesionalSCCP") = DevuelveIDCategoriaProfesionalSCCPTodasBasesDeDatos(fila.Field(Of String)("IDOperario"))
            dtFaltantes.Rows.Add(newRow)
        Next
        Return dtFaltantes
    End Function

    Public Function getGenteSiHorasNoEuros(ByVal dtHorasExpertis As DataTable, ByVal dtA3 As DataTable) As DataTable
        ' Obtén una lista de los IDGet que están en dtA3
        Dim idGetEnA3 = dtA3.AsEnumerable().Select(Function(row) row.Field(Of String)("IDGet")).ToList()

        ' Usa LINQ para encontrar las filas en dtHorasExpertis que no están en dtA3
        Dim filasFaltantes = dtHorasExpertis.AsEnumerable().Where(Function(row) Not idGetEnA3.Contains(row.Field(Of String)("IDGet")))

        ' Crea una nueva tabla con los resultados
        Dim dtFaltantes As New DataTable()
        dtFaltantes.Columns.Add("IDGet", GetType(String))
        dtFaltantes.Columns.Add("DescOperario", GetType(String))

        ' Agrega las filas faltantes a la nueva tabla
        For Each fila In filasFaltantes
            Dim newRow As DataRow = dtFaltantes.NewRow()
            newRow("IDGet") = fila.Field(Of String)("IDGet")
            newRow("DescOperario") = fila.Field(Of String)("DescOperario")
            dtFaltantes.Rows.Add(newRow)
        Next

        Return dtFaltantes
    End Function

    Public Function getHorasPersonas(ByVal mes As String, ByVal anio As String) As DataTable
        Dim dtHorasExpertis As New DataTable

        Dim ruta As String = "N:\10. AUXILIARES\00. EXPERTIS\01. HORAS\" & mes & " HORAS " & mes & anio.Substring(anio.Length - 2) & ".xlsx"
        ' Nombre de la hoja
        Dim hoja As String = "HORAS"
        Dim rango As String
        rango = "A2:N1000000"
        dtHorasExpertis = ObtenerDatosExcel(ruta, hoja, rango)
        Dim dtHorasFinal As New DataTable
        FormarTablaHoras(dtHorasFinal)

        '1º LA ordeno por IDGET
        Dim expresion As String = "F2 asc"
        Dim rows As DataRow() = dtHorasExpertis.Select("", expresion)
        Dim dtOrdenado As DataTable = dtHorasExpertis.Clone()
        For Each row As DataRow In rows
            dtOrdenado.ImportRow(row)
        Next
        '2 º RECORRO LA TABLA Y VOY SUMANDO POR IDGET LAS HORAS, LAS HORAS ADMINISTRATIVAS Y LAS HORAS DE BAJA
        'MsgBox(dtOrdenado.Rows.Count)

        ' Variables para rastrear el IDGet actual y las sumas de Horas y HorasAdministrativas y de baja
        Dim currentIDGET As String = -1
        Dim currentDescOperario As String = -1
        Dim currentCategoriaP As String = -1
        Dim sumaHoras As Double = 0
        Dim sumaHorasAdmin As Double = 0
        Dim sumaHorasBaja As Double = 0

        ' Iterar a través de las filas de dtOrdenado
        For Each row As DataRow In dtOrdenado.Rows
            Dim rowIDGET As String = row("F2")
            Dim rowDescOperario As String = row("F4")
            Dim rowCategoriaProfesional As String = Nz(row("F6"), "")
            If rowCategoriaProfesional.ToString.Length = 0 Then
                ExpertisApp.GenerateMessage("El operario con IDGET no tiene categoria profesional: " & rowIDGET)
            End If
            Dim rowHoras As Double = Convert.ToDouble(row("F11"))
            Dim rowHorasAdmin As Double = Convert.ToDouble(row("F13"))
            Dim rowHorasBaja As Double = Convert.ToDouble(row("F14"))

            ' Verificar si el IDGet ha cambiado
            If rowIDGET <> currentIDGET Then
                ' Agregar las sumas a dtHorasFinal para el IDGet anterior
                If currentIDGET <> "-1" Then
                    Dim newRow As DataRow = dtHorasFinal.NewRow()
                    newRow("IDGet") = currentIDGET
                    'Calcula IDOperario por el IDGET
                    newRow("IDOperario") = DevuelveIDOperario(currentIDGET)
                    newRow("Fecha_Alta") = devuelveFechaAlta(currentIDGET)
                    newRow("Fecha_Baja") = devuelveFechaBaja(currentIDGET)
                    newRow("DescOperario") = currentDescOperario
                    newRow("IDCategoriaProfesionalSCCP") = currentCategoriaP
                    newRow("Horas") = sumaHoras
                    newRow("HorasAdministrativas") = sumaHorasAdmin
                    newRow("HorasBaja") = sumaHorasBaja
                    newRow("HorasTotales") = sumaHoras + sumaHorasAdmin + sumaHorasBaja
                    dtHorasFinal.Rows.Add(newRow)
                End If

                ' Reiniciar las sumas para el nuevo IDGet
                currentIDGET = rowIDGET
                currentDescOperario = rowDescOperario
                currentCategoriaP = rowCategoriaProfesional
                sumaHoras = 0
                sumaHorasAdmin = 0
                sumaHorasBaja = 0
            End If

            ' Sumar las horas de las columnas Horas y HorasAdministrativas
            sumaHoras += rowHoras
            sumaHorasAdmin += rowHorasAdmin
            sumaHorasBaja += rowHorasBaja
        Next

        ' Agregar la última suma a dtHorasFinal
        If currentIDGET <> "-1" Then
            Dim newRow As DataRow = dtHorasFinal.NewRow()
            newRow("IDGet") = currentIDGET
            newRow("IDOperario") = DevuelveIDOperario(currentIDGET)
            newRow("Fecha_Alta") = devuelveFechaAlta(currentIDGET)
            newRow("Fecha_Baja") = devuelveFechaBaja(currentIDGET)
            newRow("DescOperario") = currentDescOperario
            newRow("IDCategoriaProfesionalSCCP") = currentCategoriaP
            newRow("Horas") = sumaHoras
            newRow("HorasAdministrativas") = sumaHorasAdmin
            newRow("HorasBaja") = sumaHorasBaja
            newRow("HorasTotales") = sumaHoras + sumaHorasAdmin + sumaHorasBaja
            dtHorasFinal.Rows.Add(newRow)
        End If

        'Dim totalHoras As Double = 0
        'For Each row As DataRow In dtHorasFinal.Rows
        '    Dim horas As Double = Convert.ToDouble(row("Horas"))
        '    totalHoras += horas
        'Next
        'MsgBox(totalHoras)

        Return dtHorasFinal
    End Function

    Public Function devuelveIDOperario(ByVal IDGET As String) As String
        Dim dt As New DataTable
        Dim f As New Filter
        f.Add("IDGET", FilterOperator.Equal, IDGET)
        dt = New BE.DataEngine().Filter("vUnionOperariosCategoriaProfesional", f)
        Return dt.Rows(0)("IDOperario").ToString()
    End Function

    Public Function devuelveFechaAlta(ByVal IDGET As String) As String
        Dim dt As New DataTable
        Dim f As New Filter
        f.Add("IDGET", FilterOperator.Equal, IDGET)
        dt = New BE.DataEngine().Filter("vUnionOperariosCategoriaProfesional", f)
        Return Nz(dt.Rows(0)("FechaAlta").ToString, "")
    End Function

    Public Function devuelveFechaBaja(ByVal IDGET As String) As String
        Dim dt As New DataTable
        Dim f As New Filter
        f.Add("IDGET", FilterOperator.Equal, IDGET)
        dt = New BE.DataEngine().Filter("vUnionOperariosCategoriaProfesional", f)
        Return dt.Rows(0)("Fecha_Baja").ToString
    End Function

    Public Function getEurosPersonas(ByVal mes As String, ByVal anio As String) As DataTable
        Dim dtEurosA3 As New DataTable

        Dim ruta As String = "N:\10. AUXILIARES\00. EXPERTIS\02. A3\" & mes & " A3 " & mes & anio.Substring(anio.Length - 2) & ".xlsx"
        ' Nombre de la hoja
        Dim hoja As String = "A3"
        Dim rango As String
        rango = "A2:G100000"
        dtEurosA3 = ObtenerDatosExcel(ruta, hoja, rango)
        Dim dtEurosFinal As New DataTable
        FormarTablaEuros(dtEurosFinal)

        For Each dr As DataRow In dtEurosA3.Rows
            Dim newRow As DataRow = dtEurosFinal.NewRow()
            newRow("IDGet") = dr("F2")
            newRow("IDOperario") = dr("F3")
            newRow("DescOperario") = dr("F4")
            newRow("CosteEmpresa") = dr("F7")
            dtEurosFinal.Rows.Add(newRow)
        Next

        Return dtEurosFinal
    End Function

    Public Sub FormarTablaHoras(ByRef dtHorasFinal As DataTable)
        dtHorasFinal.Columns.Add("IDGET")
        dtHorasFinal.Columns.Add("IDOperario")
        dtHorasFinal.Columns.Add("DescOperario")
        dtHorasFinal.Columns.Add("IDCategoriaProfesionalSCCP")
        dtHorasFinal.Columns.Add("Horas")
        dtHorasFinal.Columns.Add("HorasAdministrativas")
        dtHorasFinal.Columns.Add("HorasBaja")
        dtHorasFinal.Columns.Add("HorasTotales")
        dtHorasFinal.Columns.Add("Fecha_Alta")
        dtHorasFinal.Columns.Add("Fecha_Baja")
    End Sub

    Public Sub FormarTablaEuros(ByRef dtEurosFinal As DataTable)
        dtEurosFinal.Columns.Add("IDGET")
        dtEurosFinal.Columns.Add("IDOperario")
        dtEurosFinal.Columns.Add("DescOperario")
        dtEurosFinal.Columns.Add("CosteEmpresa")
    End Sub

    Private Sub bDocumentacion_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bDocumentacion.Click
        'ABRE EL WORD DEL MANUAL DEL USUARIO QUE TENGO EN LA RUTA
        'N:\DOCUMENTACION_OFICIAL\ManualDelUsuario.docx
        Dim filePath As String = "N:\1000. DOCUMENTACION OFICIAL\01. DOCUMENTOS\1000. DOCUMENTACION OFICIAL\Manual_Del_Usuario.docx"

        If System.IO.File.Exists(filePath) Then
            Process.Start(filePath)
        Else
            MessageBox.Show("El archivo no existe en la ruta especificada.", "Archivo no encontrado", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub
    Public Sub FormaTablaExtras(ByRef dtUnion As DataTable)
        dtUnion.Columns.Add("Empresa")
        dtUnion.Columns.Add("IDGET")
        dtUnion.Columns.Add("IDOperario")
        dtUnion.Columns.Add("DescOperario")
        dtUnion.Columns.Add("IDCategoriaProfesionalSCCP", System.Type.GetType("System.String"))
        dtUnion.Columns.Add("SinIncentivos", System.Type.GetType("System.Double"))
        dtUnion.Columns.Add("ConIncentivos", System.Type.GetType("System.Double"))
        dtUnion.Columns.Add("Mes", System.Type.GetType("System.Double"))
        dtUnion.Columns.Add("Anio", System.Type.GetType("System.Double"))
    End Sub
    Private Sub bExtras_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bExtras.Click
        '------------------
        Dim dtUnion As New DataTable
        FormaTablaExtras(dtUnion)
        Dim dtAuxiliar As New DataTable
        Do
            ' Aquí va el código que deseas ejecutar repetidamente
            dtAuxiliar = CargaExtrasTabla(dtUnion)
            If dtAuxiliar Is Nothing Then
                ExpertisApp.GenerateMessage("Proceso cancelado correctamente.")
                Exit Sub
            End If
            For Each row As DataRow In dtAuxiliar.Rows
                dtUnion.ImportRow(row)
            Next
            ' Preguntar al usuario si desea continuar
            Dim respuesta As DialogResult = MessageBox.Show("¿Deseas cargar algún Excel más?", "Continuar", MessageBoxButtons.YesNo)
            ' Salir del bucle si el usuario responde "No"
            If respuesta = DialogResult.No Then
                Exit Do
            End If
        Loop

        '---
        Dim mes As String
        Dim anio As String
        'CHECK DE QUE EL FICHERO ACABA EN XLSX O XLS
        Dim ruta As String = lblRuta.Text
        Dim ultimoCaracter As String = ruta.Substring(ruta.Length - 1)
        If ultimoCaracter = "x" Then
            mes = ruta.Substring(ruta.Length - 9, 2)
            anio = ruta.Substring(ruta.Length - 7, 2)
        Else
            mes = ruta.Substring(ruta.Length - 8, 2)
            anio = ruta.Substring(ruta.Length - 6, 2)
        End If

        anio = "20" & anio
        If mes = 13 Then
            mes = 6
        ElseIf mes = 14 Then
            mes = 12
        End If

        Dim dtImprimirCategorias As DataTable = FormaTablaImprimirExtrasCategoriasAMJ(dtUnion)
        GenerarExcelExtrasResumen(dtUnion, dtImprimirCategorias, mes, anio)
        GenerarExcelExtras(dtUnion, dtImprimirCategorias, mes, anio)
        MsgBox("Excel creado correctamente en N:\10. AUXILIARES\00. EXPERTIS\03. PAGAS EXTRA\")
    End Sub
    Public Sub GenerarExcelExtras(ByVal dtUnion As DataTable, ByVal dtImprimirCategorias As DataTable, ByVal mes As String, ByVal anio As String)
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial

        For Each row As DataRow In dtUnion.Rows
            row("SinIncentivos") = CDbl(row("SinIncentivos")) / 6
            row("ConIncentivos") = CDbl(row("ConIncentivos")) / 6
        Next

        For Each row As DataRow In dtImprimirCategorias.Rows
            row("Total") = CDbl(row("Total")) / 6
        Next

        Dim primero As Integer
        Dim ultimo As Integer
        Dim texto = ""
        If mes = 12 Then
            primero = 1
            ultimo = 6
            anio = anio + 1
        Else
            primero = 7
            ultimo = 12
        End If
        For i As Integer = primero To ultimo
            If Len(primero.ToString()) = 1 Then
                texto = "0" & primero
            Else
                texto = primero
            End If
            Dim ruta As New FileInfo("N:\10. AUXILIARES\00. EXPERTIS\03. PAGAS EXTRA\" & texto & " EXTRA " & texto & anio.Substring(anio.Length - 2) & ".xlsx")
            Dim rutaCadena As String = ""
            rutaCadena = ruta.FullName

            If File.Exists(rutaCadena) Then
                File.Delete(rutaCadena)
            End If

            Using package As New ExcelPackage(ruta)

                ' HOJA 1
                Dim worksheet2 = package.Workbook.Worksheets.Add("EXTRAS POR CATEGORIA PROF")
                worksheet2.Cells("A1").LoadFromDataTable(dtImprimirCategorias, True)
                Dim fila2 As ExcelRange = worksheet2.Cells(1, 1, 1, worksheet2.Dimension.End.Column)
                fila2.Style.Font.Bold = True

                worksheet2.Column(1).Width = 12
                worksheet2.Column(2).Width = 12
                worksheet2.Column(3).Width = 12

                Dim columnaC As ExcelRange = worksheet2.Cells("C2:C" & worksheet2.Dimension.End.Row)
                columnaC.Style.Numberformat.Format = "#,##0.00€"
                worksheet2.Cells("A1:" & GetExcelColumnName(worksheet2.Dimension.End.Column) & "1").AutoFilter = True

                ' HOJA 2
                Dim worksheet = package.Workbook.Worksheets.Add("EXTRAS POR PERSONA")
                worksheet.Cells("A1").LoadFromDataTable(dtUnion, True)
                Dim fila1 As ExcelRange = worksheet.Cells(1, 1, 1, worksheet.Dimension.End.Column)
                fila1.Style.Font.Bold = True
                Dim columnaF As ExcelRange = worksheet.Cells("F2:F" & worksheet.Dimension.End.Row)
                columnaF.Style.Numberformat.Format = "#,##0.00€"
                Dim columnaG As ExcelRange = worksheet.Cells("G2:G" & worksheet.Dimension.End.Row)
                columnaG.Style.Numberformat.Format = "#,##0.00€"
                worksheet.Cells("A1:" & GetExcelColumnName(worksheet.Dimension.End.Column) & "1").AutoFilter = True
                ' Guardar el archivo de Excel.
                worksheet.Column(4).Width = 30
                package.Save()
            End Using
            primero = primero + 1
        Next


    End Sub

    Public Sub GenerarExcelExtrasResumen(ByVal dtUnion As DataTable, ByVal dtImprimirCategorias As DataTable, ByVal mes As String, ByVal anio As String)
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial

        If mes = 6 Then
            mes = 13
        ElseIf mes = 12 Then
            mes = 14
        End If

        Dim ruta As New FileInfo("N:\10. AUXILIARES\00. EXPERTIS\03. PAGAS EXTRA\" & mes & " EXTRA REAL " & mes & anio.Substring(anio.Length - 2) & ".xlsx")
        Dim rutaCadena As String = ""
        rutaCadena = ruta.FullName

        If File.Exists(rutaCadena) Then
            File.Delete(rutaCadena)
        End If

        Using package As New ExcelPackage(ruta)

            ' HOJA 1
            Dim worksheet2 = package.Workbook.Worksheets.Add("EXTRAS POR CATEGORIA PROF")
            worksheet2.Cells("A1").LoadFromDataTable(dtImprimirCategorias, True)
            Dim fila2 As ExcelRange = worksheet2.Cells(1, 1, 1, worksheet2.Dimension.End.Column)
            fila2.Style.Font.Bold = True
            worksheet2.Column(1).Width = 12
            worksheet2.Column(2).Width = 12
            worksheet2.Column(3).Width = 12

            Dim columnaC As ExcelRange = worksheet2.Cells("C2:C" & worksheet2.Dimension.End.Row)
            columnaC.Style.Numberformat.Format = "#,##0.00€"
            worksheet2.Cells("A1:" & GetExcelColumnName(worksheet2.Dimension.End.Column) & "1").AutoFilter = True

            ' HOJA 2
            Dim worksheet = package.Workbook.Worksheets.Add("EXTRAS POR PERSONA")
            worksheet.Cells("A1").LoadFromDataTable(dtUnion, True)
            Dim fila1 As ExcelRange = worksheet.Cells(1, 1, 1, worksheet.Dimension.End.Column)
            fila1.Style.Font.Bold = True
            Dim columnaF As ExcelRange = worksheet.Cells("F2:F" & worksheet.Dimension.End.Row)
            columnaF.Style.Numberformat.Format = "#,##0.00€"
            Dim columnaG As ExcelRange = worksheet.Cells("G2:G" & worksheet.Dimension.End.Row)
            columnaG.Style.Numberformat.Format = "#,##0.00€"
            ' Guardar el archivo de Excel.
            worksheet.Cells("A1:" & GetExcelColumnName(worksheet.Dimension.End.Column) & "1").AutoFilter = True
            worksheet.Column(4).Width = 30
            package.Save()
        End Using

    End Sub

    Public Function FormaTablaImprimirExtrasCategoriasAMJ(ByVal dtUnion As DataTable) As DataTable
        Dim dtResultado As New DataTable()
        dtResultado.Columns.Add("Empresa", GetType(String))
        dtResultado.Columns.Add("IDCategoriaProfesionalSCCP", GetType(String))
        dtResultado.Columns.Add("Total", GetType(Double))

        Dim jprod As Double = 0 : Dim encar As Double = 0 : Dim operar As Double = 0 : Dim tecobra As Double = 0 : Dim staff As Double = 0 : Dim otros As Double = 0
        Dim jprodf As Double = 0 : Dim encarf As Double = 0 : Dim operarf As Double = 0 : Dim tecobraf As Double = 0 : Dim stafff As Double = 0 : Dim otrosf As Double = 0
        Dim jprods As Double = 0 : Dim encars As Double = 0 : Dim operars As Double = 0 : Dim tecobras As Double = 0 : Dim staffs As Double = 0 : Dim otross As Double = 0

        For Each dr As DataRow In dtUnion.Rows
            Dim empresa As String = dr("Empresa").ToString
            Dim categoria As Integer = Convert.ToInt64(dr("IDCategoriaProfesionalSCCP"))
            Dim coste As Double = Convert.ToDouble(dr("ConIncentivos"))
            Dim incentivos As Double = Convert.ToDouble(dr("SinIncentivos"))
            If empresa = "T. ES." Then
                Select Case categoria
                    Case 1
                        jprod = jprod + coste + incentivos
                    Case 2
                        encar = encar + coste + incentivos
                    Case 3
                        operar = operar + coste + incentivos
                    Case 4
                        tecobra = tecobra + coste + incentivos
                    Case 5
                        staff = staff + coste + incentivos
                    Case Else
                        otros = otros + coste + incentivos
                End Select
            ElseIf empresa = "FERR." Then
                Select Case categoria
                    Case 1
                        jprodf = jprodf + coste + incentivos
                    Case 2
                        encarf = encarf + coste + incentivos
                    Case 3
                        operarf = operarf + coste + incentivos
                    Case 4
                        tecobraf = tecobraf + coste + incentivos
                    Case 5
                        stafff = stafff + coste + incentivos
                    Case Else
                        otrosf = otrosf + coste + incentivos
                End Select
            ElseIf empresa = "SEC." Then
                Select Case categoria
                    Case 1
                        jprods = jprods + coste + incentivos
                    Case 2
                        encars = encars + coste + incentivos
                    Case 3
                        operars = operars + coste + incentivos
                    Case 4
                        tecobras = tecobras + coste + incentivos
                    Case 5
                        staffs = staffs + coste + incentivos
                    Case Else
                        otross = otross + coste + incentivos
                End Select
            End If

        Next
        '-1.TECOZAM
        Dim newRow As DataRow = dtResultado.NewRow()
        newRow("Empresa") = "T. ES." : newRow("IDCategoriaProfesionalSCCP") = 1 : newRow("Total") = jprod : dtResultado.Rows.Add(newRow)
        newRow = dtResultado.NewRow() : newRow("Empresa") = "T. ES." : newRow("IDCategoriaProfesionalSCCP") = 2 : newRow("Total") = encar
        dtResultado.Rows.Add(newRow)
        newRow = dtResultado.NewRow() : newRow("Empresa") = "T. ES." : newRow("IDCategoriaProfesionalSCCP") = 3 : newRow("Total") = operar
        dtResultado.Rows.Add(newRow)
        newRow = dtResultado.NewRow() : newRow("Empresa") = "T. ES." : newRow("IDCategoriaProfesionalSCCP") = 4 : newRow("Total") = tecobra
        dtResultado.Rows.Add(newRow)
        newRow = dtResultado.NewRow() : newRow("Empresa") = "T. ES." : newRow("IDCategoriaProfesionalSCCP") = 5 : newRow("Total") = staff
        dtResultado.Rows.Add(newRow)
        newRow = dtResultado.NewRow() : newRow("Empresa") = "T. ES." : newRow("IDCategoriaProfesionalSCCP") = 0 : newRow("Total") = otros
        dtResultado.Rows.Add(newRow)
        '-2. FERRALLAS
        newRow = dtResultado.NewRow()
        newRow("Empresa") = "FERR." : newRow("IDCategoriaProfesionalSCCP") = 1 : newRow("Total") = jprodf : dtResultado.Rows.Add(newRow)
        newRow = dtResultado.NewRow() : newRow("Empresa") = "FERR." : newRow("IDCategoriaProfesionalSCCP") = 2 : newRow("Total") = encarf
        dtResultado.Rows.Add(newRow)
        newRow = dtResultado.NewRow() : newRow("Empresa") = "FERR." : newRow("IDCategoriaProfesionalSCCP") = 3 : newRow("Total") = operarf
        dtResultado.Rows.Add(newRow)
        newRow = dtResultado.NewRow() : newRow("Empresa") = "FERR." : newRow("IDCategoriaProfesionalSCCP") = 4 : newRow("Total") = tecobraf
        dtResultado.Rows.Add(newRow)
        newRow = dtResultado.NewRow() : newRow("Empresa") = "FERR." : newRow("IDCategoriaProfesionalSCCP") = 5 : newRow("Total") = stafff
        dtResultado.Rows.Add(newRow)
        newRow = dtResultado.NewRow() : newRow("Empresa") = "FERR." : newRow("IDCategoriaProfesionalSCCP") = 0 : newRow("Total") = otrosf
        dtResultado.Rows.Add(newRow)
        '-3. SECOZAM
        newRow = dtResultado.NewRow()
        newRow("Empresa") = "SEC." : newRow("IDCategoriaProfesionalSCCP") = 1 : newRow("Total") = jprods : dtResultado.Rows.Add(newRow)
        newRow = dtResultado.NewRow() : newRow("Empresa") = "SEC." : newRow("IDCategoriaProfesionalSCCP") = 2 : newRow("Total") = encars
        dtResultado.Rows.Add(newRow)
        newRow = dtResultado.NewRow() : newRow("Empresa") = "SEC." : newRow("IDCategoriaProfesionalSCCP") = 3 : newRow("Total") = operars
        dtResultado.Rows.Add(newRow)
        newRow = dtResultado.NewRow() : newRow("Empresa") = "SEC." : newRow("IDCategoriaProfesionalSCCP") = 4 : newRow("Total") = tecobras
        dtResultado.Rows.Add(newRow)
        newRow = dtResultado.NewRow() : newRow("Empresa") = "SEC." : newRow("IDCategoriaProfesionalSCCP") = 5 : newRow("Total") = staffs
        dtResultado.Rows.Add(newRow)
        newRow = dtResultado.NewRow() : newRow("Empresa") = "SEC." : newRow("IDCategoriaProfesionalSCCP") = 0 : newRow("Total") = otross
        dtResultado.Rows.Add(newRow)

        Return dtResultado
    End Function

    Public Function DevuelveIncentivos(ByVal bbdd As String, ByVal IDOperario As String) As String
        Dim f As New Filter
        f.Add("IDOperario", FilterOperator.Equal, IDOperario)
        Dim dt As DataTable
        dt = New BE.DataEngine().Filter(bbdd & "..frmMntoOperario", f)

        Try
            If dt.Rows(0)("Incentivos") = False Or String.IsNullOrEmpty(dt.Rows(0)("Incentivos").ToString()) Then
                Return "0"
                Exit Function
            End If
        Catch ex As Exception
            Return "0"
        End Try


        Return "1"
    End Function

    Private Sub bPisarFicheroExtra_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bPisarFicheroExtra.Click
        'EN DICIEMBRE ME GENERA DEL 1 AL 6
        'EN JUNIO ME GENERA DEL 7 AL 12
        '------
        'EN JUNIO NORMALIZO EL FICHERO 6 GENERADO AL METER EL 12 RESTANDO DEL A3 DE JUNIO PRORRATEADO
        Dim CD As New OpenFileDialog()
        MsgBox("Selecciona el fichero de extras entero sin prorratear." & vbCrLf & "JUNIO->13 EXTRA REAL 13YY" & vbCrLf & "DICIEMBRE->14 EXTRA REAL 14YY", vbInformation)
        CD.Title = "Seleccionar archivos"
        CD.Filter = "Archivos Excel(*.xls;*.xlsx)|*.xls;*xlsx|Todos los archivos(*.*)|*.*"
        CD.ShowDialog()

        If CD.FileName <> "" Then
            lblRuta.Text = CD.FileName
        End If

        Dim hoja As String = "EXTRAS POR CATEGORIA PROF"
        Dim dtFicheroExtra As New DataTable
        Dim ruta As String = lblRuta.Text

        Dim rango As String = ""
        rango = "A2:C19"
        dtFicheroExtra = ObtenerDatosExcel(ruta, hoja, rango)


        MsgBox("Selecciona el fichero de extras a normalizar. JUNIO -> 06 EXTRA 06YY Y DICIEMBRE -> 12 EXTRA 12YY", vbInformation)
        CD.Title = "Seleccionar archivos"
        CD.Filter = "Archivos Excel(*.xls;*.xlsx)|*.xls;*xlsx|Todos los archivos(*.*)|*.*"
        CD.ShowDialog()

        Dim cadena As String = ""

        If CD.FileName <> "" Then
            lblRuta.Text = CD.FileName
            cadena = CD.FileName
        End If

        hoja = "EXTRAS POR CATEGORIA PROF"
        Dim dtNormalizar As New DataTable
        ruta = lblRuta.Text
        rango = "A2:C19"
        dtNormalizar = ObtenerDatosExcel(ruta, hoja, rango)

        Dim dtGenerar As New DataTable
        dtGenerar = generarTablaFicheroExtra(dtFicheroExtra, dtNormalizar)
        dtGenerar.Columns(0).ColumnName = "Empresa"
        dtGenerar.Columns(1).ColumnName = "IDCategoriaProfesionalSCCP"
        dtGenerar.Columns(2).ColumnName = "Total"

        GuardarExcel(dtGenerar, cadena)
    End Sub
    Public Sub GuardarExcel(ByVal dtGenerar As DataTable, ByVal cadena As String)
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial

        'Dim ruta As New FileInfo("N:\10. AUXILIARES\00. EXPERTIS\02. A3\" & mes & " A3 " & mes & anio.Substring(anio.Length - 2) & ".xlsx")
        Dim ruta As New FileInfo(cadena)
        Dim rutaCadena As String = ""
        rutaCadena = ruta.FullName

        'Verificar si el archivo existe.
        If File.Exists(rutaCadena) Then
            'Si el archivo existe, eliminarlo.
            File.Delete(rutaCadena)
            MsgBox("Fichero anterior normalizado.")
        End If

        Using package As New ExcelPackage(ruta)
            ' Crear una hoja de cálculo y obtener una referencia a ella.
            Dim worksheet2 = package.Workbook.Worksheets.Add("EXTRAS POR CATEGORIA PROF")

            ' Copiar los datos de la DataTable a la hoja de cálculo.
            worksheet2.Cells("A1").LoadFromDataTable(dtGenerar, True)

            ' Aplicar formato negrita a la fila 1
            Dim fila1 As ExcelRange = worksheet2.Cells(1, 1, 1, worksheet2.Dimension.End.Column)
            fila1.Style.Font.Bold = True

            worksheet2.Column(1).Width = 12
            worksheet2.Column(2).Width = 12
            worksheet2.Column(3).Width = 12

            Dim columnaC As ExcelRange = worksheet2.Cells("C2:C" & worksheet2.Dimension.End.Row)
            columnaC.Style.Numberformat.Format = "#,##0.00€"
            worksheet2.Cells("A1:" & GetExcelColumnName(worksheet2.Dimension.End.Column) & "1").AutoFilter = True
            ' Guardar el archivo de Excel.
            package.Save()
        End Using
    End Sub
    Public Function generarTablaFicheroExtra(ByVal dtFicheroExtra As DataTable, ByVal dtNormalizar As DataTable) As DataTable
        '1. MULTIPLIC DTNORMALIZAR POR 5
        '2. LOS RESTO DEL TOTAL
        For Each row As DataRow In dtNormalizar.Rows
            row("F3") = CDbl(row("F3")) * 5
        Next

        ' Realizar la resta de las tablas
        For i As Integer = 0 To dtFicheroExtra.Rows.Count - 1
            dtNormalizar.Rows(i)("F3") = CDbl(dtFicheroExtra.Rows(i)("F3")) - CDbl(dtNormalizar.Rows(i)("F3"))
        Next

        Return dtNormalizar
    End Function

    Private Sub bRegularizarSemestral_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bRegularizarSemestral.Click
        'Fecha2 = diaMes & "/" & mes & "/" & año & ""
        Dim frm As New frmInformeFecha
        frm.ShowDialog()
        Dim Fecha1 As String : Dim Fecha2 As String
        Fecha1 = frm.fecha1 : Fecha2 = frm.fecha2

        Dim mes1 As String : mes1 = Month(Fecha1)
        Dim mes2 As String : mes2 = Month(Fecha2)
        If Length(mes1) = 1 Then
            mes1 = "0" & mes1
        End If
        If Length(mes2) = 1 Then
            mes2 = "0" & mes2
        End If

        Dim anio As String
        anio = Year(Fecha1)
        'Hay que obtener los ficheros entre los meses mes1 y mes2
        Dim dtA3EntreFechasPowerBi As DataTable

        '1º. ESTA ES LA 3ª HOJA QUE ES LA UNION DE TODOS LOS A3 QUE SE HAN INSERTADO
        '------------------
        Dim dtFinalA3 As New DataTable
        FormaTablaA3(dtFinalA3)
        Dim dtAuxiliar As New DataTable
        Do
            ' Aquí va el código que deseas ejecutar repetidamente
            dtAuxiliar = CargaExcelA3Semestral()
            If dtAuxiliar Is Nothing Then
                ExpertisApp.GenerateMessage("Proceso cancelado correctamente.")
                Exit Sub
            End If
            For Each row As DataRow In dtAuxiliar.Rows
                dtFinalA3.ImportRow(row)
            Next
            ' Preguntar al usuario si desea continuar
            Dim respuesta As DialogResult = MessageBox.Show("¿Deseas cargar algún Excel más?", "Continuar", MessageBoxButtons.YesNo)
            ' Salir del bucle si el usuario responde "No"
            If respuesta = DialogResult.No Then
                Exit Do
            End If
        Loop

        Dim dtSemestral As DataTable
        dtSemestral = getTablaSemestral(dtFinalA3)


        '2º. ESTA ES LA 2ª HOJA DEL EXCEL QUE SUMA LOS A3 QUE SE HAN GENERADO ENTRE DOS FECHAS
        dtA3EntreFechasPowerBi = getA3EntreFechasPowerBi(mes1, mes2, anio)


        '3º. ESTA ES LA 1ª HOJA QUE ES LA DIFERENCIA ENTRE LOS FICHEROS QUE SE HAN SUMADO ENTRE DOS FECHAS Y UNO QUE AGRUPE AL RESTO
        'PRINCIPALMENTE SE HACE LA HERRAMIENTA PARA NORMALIZAR A CARACTER SEMESTRAL O ANUAL
        Dim dtRegularizar As DataTable
        dtRegularizar = generarTablaRegularizar(dtA3EntreFechasPowerBi, dtSemestral)

        GeneraExcelRegularizarA3(dtRegularizar, dtA3EntreFechasPowerBi, dtSemestral, mes2, anio)
    End Sub
    Public Sub GeneraExcelRegularizarA3(ByVal dtRegularizar As DataTable, ByVal dtA3EntreFechasPowerBi As DataTable, ByVal dtSemestral As DataTable, _
                                        ByVal mes As String, ByVal anio As String)
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial

        Dim ruta As New FileInfo("N:\10. AUXILIARES\00. EXPERTIS\04. REGULARIZACIONES\" & mes & " REGULARIZACIONES " & mes & anio.Substring(anio.Length - 2) & ".xlsx")
        Dim rutaCadena As String = ""
        rutaCadena = ruta.FullName

        'Verificar si el archivo existe.
        If File.Exists(rutaCadena) Then
            'Si el archivo existe, eliminarlo.
            File.Delete(rutaCadena)
        End If

        Using package As New ExcelPackage(ruta)
            ' Crear una hoja de cálculo y obtener una referencia a ella.
            Dim worksheet = package.Workbook.Worksheets.Add(mes & " REGULARIZACIONES " & anio)
            ' Copiar los datos de la DataTable a la hoja de cálculo.
            worksheet.Cells("A1").LoadFromDataTable(dtRegularizar, True)

            Dim columnaE As ExcelRange = worksheet.Cells("C2:C" & worksheet.Dimension.End.Row)
            columnaE.Style.Numberformat.Format = "#,##0.00€"

            ' Aplicar formato negrita a la fila 1
            Dim fila1 As ExcelRange = worksheet.Cells(1, 1, 1, worksheet.Dimension.End.Column)
            fila1.Style.Font.Bold = True
            worksheet.Column(1).Width = 14
            worksheet.Column(2).Width = 14
            worksheet.Column(3).Width = 14

            'SEGUNDA HOJA DEL EXCEL= SUMA DE LOS FICHEROS DE POWER BI
            Dim resumenWorksheet = package.Workbook.Worksheets.Add("RESUMEN FICHEROS MENSUALES")
            resumenWorksheet.Cells("A1").LoadFromDataTable(dtA3EntreFechasPowerBi, True)

            Dim columnaBResumen As ExcelRange = resumenWorksheet.Cells("C2:C" & resumenWorksheet.Dimension.End.Row)
            columnaBResumen.Style.Numberformat.Format = "#,##0.00€"

            ' Aplicar formato negrita a la fila 1
            Dim filaResumen1 As ExcelRange = resumenWorksheet.Cells(1, 1, 1, resumenWorksheet.Dimension.End.Column)
            filaResumen1.Style.Font.Bold = True

            resumenWorksheet.Column(1).Width = 14
            resumenWorksheet.Column(2).Width = 14
            resumenWorksheet.Column(3).Width = 14

            'TERCERA HOJA DEL EXCEL QUE SON LOS A3 INSERTADOS CON CARACTER SEMESTRAL
            Dim resumenCategoria = package.Workbook.Worksheets.Add("RESUMEN FICHEROS SEMESTRAL")
            resumenCategoria.Cells("A1").LoadFromDataTable(dtSemestral, True)

            Dim f1 As ExcelRange = resumenCategoria.Cells(1, 1, 1, resumenCategoria.Dimension.End.Column)
            f1.Style.Font.Bold = True

            Dim columnaB As ExcelRange = resumenCategoria.Cells("C2:C" & resumenCategoria.Dimension.End.Row)
            columnaB.Style.Numberformat.Format = "#,##0.00€"
            resumenCategoria.Column(1).Width = 14
            resumenCategoria.Column(2).Width = 14
            resumenCategoria.Column(3).Width = 14

            ' Guardar el archivo de Excel.
            package.Save()
        End Using

        MsgBox("Fichero creado correctamente en N:\10. AUXILIARES\00. EXPERTIS\04. REGULARIZACIONES")
    End Sub
    Public Function generarTablaRegularizar(ByVal dtA3EntreFechasPowerBi As DataTable, ByVal dtSemestral As DataTable) As DataTable
        ' Crear un nuevo DataTable para almacenar los resultados finales
        Dim dtFinal As New DataTable
        dtFinal.Columns.Add("Empresa", GetType(String))
        dtFinal.Columns.Add("IDCategoriaProfesionalSCCP", GetType(Integer))
        dtFinal.Columns.Add("Total", GetType(Double))

        ' Recorrer la primera tabla (dtA3EntreFechasPowerBi)
        For Each filaPowerBi As DataRow In dtA3EntreFechasPowerBi.Rows
            Dim empresaPowerBi As String = CStr(filaPowerBi("Empresa"))
            Dim categoriaPowerBi As Integer = CInt(filaPowerBi("IDCategoriaProfesionalSCCP"))
            Dim totalPowerBi As Double = CDbl(filaPowerBi("Total"))

            ' Recorrer la segunda tabla (dtSemestral) para buscar coincidencias
            For Each filaSemestral As DataRow In dtSemestral.Rows
                Dim empresaSemestral As String = CStr(filaSemestral("Empresa"))
                Dim categoriaSemestral As Integer = CInt(filaSemestral("IDCategoriaProfesionalSCCP"))
                Dim totalSemestral As Double = CDbl(filaSemestral("Total"))

                ' Comprobar si la empresa y categoría coinciden en ambas tablas
                If empresaPowerBi = empresaSemestral And categoriaPowerBi = categoriaSemestral Then
                    ' Restar el valor de la segunda tabla al valor de la primera tabla
                    totalPowerBi -= totalSemestral
                End If
            Next

            ' Agregar una nueva fila al resultado con los valores actualizados
            Dim nuevaFila As DataRow = dtFinal.NewRow()
            nuevaFila("Empresa") = empresaPowerBi
            nuevaFila("IDCategoriaProfesionalSCCP") = categoriaPowerBi
            nuevaFila("Total") = -totalPowerBi 'EN NEGATIVO PARA CAMBIAR EL VALOR PORQUE REALMENTE SE RESTA LA TABLA DE LA PESTAÑA 3 - LA 2

            ' Agregar la fila al DataTable final
            dtFinal.Rows.Add(nuevaFila)
        Next
        Return dtFinal
    End Function
    Public Function getTablaSemestral(ByVal dtFinalA3 As DataTable) As DataTable
        Dim tablaResultado As New DataTable
        tablaResultado.Columns.Add("Empresa", GetType(String))
        tablaResultado.Columns.Add("IDCategoriaProfesionalSCCP", GetType(Integer))
        tablaResultado.Columns.Add("Total", GetType(Double))

        ' Obtener la cantidad de categorías posibles (1 a 5)
        Dim numCategorias As Integer = 5

        ' Recorrer la tabla para cada categoría y empresa
        For categoria As Integer = 1 To numCategorias
            For Each fila As DataRow In dtFinalA3.Rows
                Dim idCategoria As Integer = CInt(fila("IDCategoriaProfesionalSCCP"))
                Dim empresa As String = CStr(fila("Empresa"))
                Dim costoEmpresa As Double = CDbl(Nz(fila("CosteEmpresa"), 0))

                ' Comprobar si la fila coincide con la categoría actual
                If idCategoria = categoria Then
                    ' Buscar si ya existe una fila para esta categoría y empresa en la tabla de resultados
                    Dim filasExistentes() As DataRow = tablaResultado.Select("IDCategoriaProfesionalSCCP = " & categoria & " AND Empresa = '" & empresa & "'")

                    If filasExistentes.Length > 0 Then
                        ' Si existe una fila para esta categoría y empresa, acumular el costo
                        filasExistentes(0)("Total") = CDbl(filasExistentes(0)("Total")) + costoEmpresa
                    Else
                        ' Si no existe una fila, agregar una nueva
                        Dim filaResultado As DataRow = tablaResultado.NewRow()
                        filaResultado("IDCategoriaProfesionalSCCP") = categoria
                        filaResultado("Empresa") = empresa
                        filaResultado("Total") = costoEmpresa
                        tablaResultado.Rows.Add(filaResultado)
                    End If
                End If
            Next
        Next

        Return tablaResultado
    End Function
    Public Function CargaExcelA3Semestral() As DataTable
        Dim CD As New OpenFileDialog()
        CD.Title = "Seleccionar archivos"
        CD.Filter = "Archivos Excel(*.xls;*.xlsx)|*.xls;*xlsx|Todos los archivos(*.*)|*.*"
        CD.ShowDialog()

        If CD.FileName <> "" Then
            lblRuta.Text = CD.FileName
        End If

        'La hoja siempre es 1
        Dim hoja As String = "1"
        Dim dt As New DataTable
        Dim ruta As String = lblRuta.Text
        Dim empresa As String = DevuelveValorEntreParentesis(ruta)
        Dim rango As String = ""
        Select Case empresa
            Case "T. ES.", "FERR.", "SEC."
                rango = "B10:Z10000"
            Case "D. P."
                rango = "A2:F500"
            Case "T. UK."
                rango = "A2:F500"
            Case Else
                MsgBox("El nombre identificado entre parentesis no se reconoce pero funciona. Coje las 3 primeras columnas.")
                rango = "A2:C10000"
        End Select

        dt = ObtenerDatosExcel(ruta, hoja, rango)

        Dim mes As String
        Dim anio As String

        'CHECK DE QUE EL FICHERO ACABA EN XLSX O XLS
        Dim ultimoCaracter As String = ruta.Substring(ruta.Length - 1)

        If ultimoCaracter = "x" Then
            mes = ruta.Substring(ruta.Length - 9, 2)
            anio = ruta.Substring(ruta.Length - 7, 2)
        Else
            mes = ruta.Substring(ruta.Length - 8, 2)
            anio = ruta.Substring(ruta.Length - 6, 2)
        End If
        anio = "20" & anio

        'FORMO LA TABLA FINAL
        dt = FormarTablaPorEmpresaRegularizacion(dt, mes, anio, empresa)
        Return dt
    End Function

    Public Function FormarTablaPorEmpresaRegularizacion(ByVal dt As DataTable, ByVal mes As String, ByVal anio As String, ByVal empresa As String) As DataTable

        Dim newDataTable As DataTable = New DataTable
        newDataTable.Columns.Add("IDGET")
        newDataTable.Columns.Add("IDOperario")
        newDataTable.Columns.Add("DescOperario")
        newDataTable.Columns.Add("IDCategoriaProfesionalSCCP")
        newDataTable.Columns.Add("CosteEmpresa", System.Type.GetType("System.Double"))
        newDataTable.Columns.Add("Mes")
        newDataTable.Columns.Add("Anio")
        newDataTable.Columns.Add("Empresa")

        Dim bbdd As String = ""
        If empresa = "T. ES." Then
            bbdd = DB_TECOZAM
        ElseIf empresa = "FERR." Then
            bbdd = DB_FERRALLAS
        ElseIf empresa = "SEC." Then
            bbdd = DB_SECOZAM
        End If

        Dim f As New Filter
        Dim dtEstatica As DataTable
        dtEstatica = New BE.DataEngine().Filter(bbdd & "..frmMntoOperario", f)

        Dim fil As New Filter
        Dim dtTodasBasesDatos As DataTable
        dtTodasBasesDatos = New BE.DataEngine().Filter("vUnionOperariosCategoriaProfesional", fil)

        newDataTable = FormaTablaEspañaRegularizar(dt, newDataTable, bbdd, mes, anio, empresa, dtEstatica, dtTodasBasesDatos)
        Return newDataTable
    End Function
    Public Function FormaTablaEspañaRegularizar(ByVal dt As DataTable, ByVal newDataTable As DataTable, ByVal bbdd As String, ByVal mes As String, ByVal anio As String, _
                ByVal empresa As String, ByVal dtEstatica As DataTable, ByVal dtTodasBasesDatos As DataTable)

        Dim filas As Integer = 0
        PvProgreso.Value = 0 : PvProgreso.Maximum = dt.Rows.Count
        PvProgreso.Step = 1 : PvProgreso.Visible = True

        Dim IDOperario As String = ""
        ' Copiar los datos de las columnas seleccionadas al nuevo DataTable
        For Each row As DataRow In dt.Rows
            'Verificar si la celda está vacía
            If Len(row("F1").ToString) = 0 Then
                'Return newDataTable
                Exit For ' Salir del bucle si la celda está vacía
            End If

            Dim newRow As DataRow = newDataTable.NewRow()
            Dim filaEncontrada() As DataRow = dtEstatica.Select("DNI = '" & row("F3") & "'")
            'IDOperario = DevuelveIDOperario(bbdd, row("F3"))
            IDOperario = filaEncontrada(0)("IDOperario").ToString

            Dim filaEncontradaCategoriaProf() As DataRow = dtTodasBasesDatos.Select("IDOperario = '" & IDOperario & "'")
            '-------
            Windows.Forms.Application.DoEvents()
            LProgreso.Text = "OBTENIENDO DATOS DE : " & IDOperario
            Windows.Forms.Application.DoEvents()
            '-------
            newRow("IDOperario") = IDOperario
            newRow("DescOperario") = row("F2")
            newRow("IDCategoriaProfesionalSCCP") = filaEncontradaCategoriaProf(0)("CategoriaProfesionalSCCP").ToString
            'newRow("IDCategoriaProfesionalSCCP") = DevuelveIDCategoriaProfesionalSCCPTodasBasesDeDatos(IDOperario)
            newRow("IDGET") = filaEncontrada(0)("IDGET").ToString
            'newRow("IDGET") = DevuelveIDGET(bbdd, IDOperario)
            newRow("CosteEmpresa") = row("F8")
            newRow("Mes") = mes
            newRow("Anio") = anio
            newRow("Empresa") = empresa

            newDataTable.Rows.Add(newRow)
            '-----
            filas = filas + 1
            PvProgreso.Value = filas
        Next

        Return newDataTable
    End Function

    Public Sub FormaTablaA3(ByRef dtFinal As DataTable)
        dtFinal.Columns.Add("IDGET")
        dtFinal.Columns.Add("IDOperario")
        dtFinal.Columns.Add("DescOperario")
        dtFinal.Columns.Add("IDCategoriaProfesionalSCCP")
        dtFinal.Columns.Add("CosteEmpresa", System.Type.GetType("System.Double"))
        dtFinal.Columns.Add("Mes")
        dtFinal.Columns.Add("Anio")
        dtFinal.Columns.Add("Empresa")
    End Sub
    Public Function getA3EntreFechasPowerBi(ByVal mes1 As String, ByVal mes2 As String, ByVal anio As String)
        Dim mes1Inte As Integer
        mes1Inte = Integer.Parse(mes1)

        Dim mes2Inte As Integer
        mes2Inte = Integer.Parse(mes2)

        Dim ruta As String
        Dim hoja As String = "RESUMEN POR CATEGORIA"
        Dim rango As String = ""
        rango = "A2:C100"
        Dim dtAuxiliar As DataTable

        Dim dtTotal As New DataTable()
        dtTotal.Columns.Add("Empresa", GetType(String))
        dtTotal.Columns.Add("IDCategoriaProfesionalSCCP", GetType(String))
        dtTotal.Columns.Add("Total", GetType(Double))

        For i = mes1Inte To mes2Inte
            ruta = ""
            If Length(i) = 1 Then
                ruta = "0" & i
            Else
                ruta = i
            End If
            ruta = "N:\01. A3\" & ruta & " A3 " & ruta & "" & anio.Substring(anio.Length - 2) & ".xlsx"
            dtAuxiliar = ObtenerDatosExcel(ruta, hoja, rango)

            Dim newRow As DataRow
            For Each dr As DataRow In dtAuxiliar.Rows
                newRow = dtTotal.NewRow()
                newRow("Empresa") = dr("F1")
                newRow("IDCategoriaProfesionalSCCP") = dr("F2")
                newRow("Total") = dr("F3")
                dtTotal.Rows.Add(newRow)
            Next
        Next

        ' Supongamos que tienes un DataTable llamado dtTotal con la estructura adecuada

        Dim dtConsolidado As New DataTable()
        dtConsolidado.Columns.Add("Empresa", GetType(String))
        dtConsolidado.Columns.Add("IDCategoriaProfesionalSCCP", GetType(Integer))
        dtConsolidado.Columns.Add("Total", GetType(Decimal))

        ' Recorrer las filas del DataTable original
        For i As Integer = 0 To dtTotal.Rows.Count - 1
            Dim row As DataRow = dtTotal.Rows(i)
            Dim empresa As String = row.Field(Of String)("Empresa")
            Dim categoria As String = row.Field(Of String)("IDCategoriaProfesionalSCCP").ToString
            Dim total As Double = row.Field(Of Double)("Total")

            ' Verificar si ya existe una fila con la misma combinación de Empresa e IDCategoriaProfesionalSCCP
            Dim found As Boolean = False
            For j As Integer = 0 To dtConsolidado.Rows.Count - 1
                If dtConsolidado.Rows(j)("Empresa").ToString() = empresa AndAlso Convert.ToInt32(dtConsolidado.Rows(j)("IDCategoriaProfesionalSCCP")) = categoria Then
                    ' Si existe, sumar el Total al registro existente
                    dtConsolidado.Rows(j)("Total") = Convert.ToDecimal(dtConsolidado.Rows(j)("Total")) + total
                    found = True
                    Exit For
                End If
            Next

            ' Si no se encontró una fila existente, agregar una nueva fila
            If Not found Then
                Dim newRow As DataRow = dtConsolidado.NewRow()
                newRow("Empresa") = empresa
                newRow("IDCategoriaProfesionalSCCP") = categoria
                newRow("Total") = total
                dtConsolidado.Rows.Add(newRow)
            End If
        Next

        Return dtConsolidado
    End Function

    Private Sub bDCZ_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bDCZ.Click
        Dim dtPersonasPortugal As DataTable = SeleccionarPDFyLeerDataTable()
        dtPersonasPortugal = darFormaTabla(dtPersonasPortugal)
        ExportaExcel(dtPersonasPortugal)
    End Sub

    Public Sub ExportaExcel(ByVal dtPersonasPortugal As DataTable)
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial

        'Dim ruta As New FileInfo("N:\10. AUXILIARES\00. EXPERTIS\02. A3\" & mes & " A3 " & mes & anio.Substring(anio.Length - 2) & ".xlsx")
        Dim ruta As New FileInfo("N:\10. AUXILIARES\00. EXPERTIS\02. A3\" & rutaPDF.Substring(0, rutaPDF.Length - 4) & ".xlsx")
        Dim rutaCadena As String = ""
        rutaCadena = ruta.FullName

        'Verificar si el archivo existe.
        If File.Exists(rutaCadena) Then
            'Si el archivo existe, eliminarlo.
            File.Delete(rutaCadena)
        End If

        Using package As New ExcelPackage(ruta)
            ' Crear una hoja de cálculo y obtener una referencia a ella.
            Dim worksheet = package.Workbook.Worksheets.Add("1")

            ' Copiar los datos de la DataTable a la hoja de cálculo.
            worksheet.Cells("A1").LoadFromDataTable(dtPersonasPortugal, True)

            Dim columnaA As ExcelRange = worksheet.Cells("A2:A" & worksheet.Dimension.End.Row)
            columnaA.Style.Numberformat.Format = "@"

            Dim rangoMoneda As ExcelRange = worksheet.Cells("B2:G" & worksheet.Dimension.End.Row)
            rangoMoneda.Style.Numberformat.Format = "#,##0.00€"

            ' Aplicar formato negrita a la fila 1
            Dim fila1 As ExcelRange = worksheet.Cells(1, 1, 1, worksheet.Dimension.End.Column)
            fila1.Style.Font.Bold = True

            ' Agregar un filtro a la primera fila
            worksheet.Cells("A1:" & GetExcelColumnName(worksheet.Dimension.End.Column) & "1").AutoFilter = True
            worksheet.Column(2).Width = 30
            worksheet.Column(3).Width = 12
            worksheet.Column(4).Width = 12
            worksheet.Column(5).Width = 12
            worksheet.Column(6).Width = 12
            worksheet.Column(7).Width = 12

            ' Guardar el archivo de Excel.
            package.Save()
        End Using

        MsgBox("Fichero creado correctamente en N:\10. AUXILIARES\00. EXPERTIS\02. A3\")
    End Sub

    Public Function darFormaTabla(ByVal dtPersonasPortugal As DataTable) As DataTable
        Dim dataTable As New DataTable()
        dataTable.Columns.Add("Diccionario")
        dataTable.Columns.Add("Operario")
        dataTable.Columns.Add("Venciminetos", System.Type.GetType("System.Double"))
        dataTable.Columns.Add("Patronal", System.Type.GetType("System.Double"))
        dataTable.Columns.Add("Remuneraciones", System.Type.GetType("System.Double"))
        dataTable.Columns.Add("Descuentos", System.Type.GetType("System.Double"))
        dataTable.Columns.Add("Liquidas", System.Type.GetType("System.Double"))

        ' Recorrer cada fila de dtPersonaPortugal
        For Each fila As DataRow In dtPersonasPortugal.Rows
            ' Obtener los valores de la columna "valores" y dividirlos
            Dim valores As String() = fila("valores").ToString().Split(" "c)

            ' Añadir una nueva fila a la nueva tabla
            Dim nuevaFila As DataRow = dataTable.NewRow()

            ' Asignar los valores a las columnas correspondientes
            nuevaFila("Venciminetos") = valores(0)
            nuevaFila("Patronal") = valores(1)
            nuevaFila("Remuneraciones") = valores(2)
            nuevaFila("Descuentos") = valores(3)
            Dim indiceComa As Integer = valores(4).IndexOf(",")
            nuevaFila("Liquidas") = valores(4).Substring(0, indiceComa + 3)
            'Separo el diccionario, dos digitos despues de la coma hasta el final
            Dim diccionario As String = valores(4).Substring(indiceComa + 3)
            nuevaFila("Diccionario") = diccionario
            nuevaFila("Operario") = fila("operario").ToString()

            ' Agregar la nueva fila a la nueva tabla
            dataTable.Rows.Add(nuevaFila)
        Next
        Return dataTable
    End Function

    Dim rutaPDF As String

    Function SeleccionarPDFyLeerDataTable() As DataTable
        ' Crear un cuadro de diálogo para seleccionar el archivo PDF
        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Filter = "Archivos PDF|*.pdf"
        openFileDialog.Title = "Selecciona un archivo PDF"

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            ' Llamada a la función para leer el PDF y convertirlo a DataTable
            rutaPDF = Path.GetFileName(openFileDialog.FileName)
            Return LeerPDFaDataTableDCZ(openFileDialog.FileName)
        Else
            ' El usuario canceló la selección
            Return Nothing
        End If
    End Function

    Sub SeleccionarPDFyLeerDataTableUK(ByVal fichero As Integer)
        ' Crear un cuadro de diálogo para seleccionar el archivo PDF
        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Filter = "Archivos PDF|*.pdf"
        openFileDialog.Title = "Selecciona un archivo PDF"

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            ' Llamada a la función para leer el PDF y convertirlo a DataTable
            LeerPDFaDataTableUK(openFileDialog.FileName, fichero)
        End If
    End Sub
    Sub SeleccionarPDFyLeerDataTableUKNuevo(ByVal fichero As Integer, ByVal dtUkPersonas As DataTable)
        ' Crear un cuadro de diálogo para seleccionar el archivo PDF
        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Filter = "Archivos PDF|*.pdf"
        openFileDialog.Title = "Selecciona un archivo PDF"

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            ' Llamada a la función para leer el PDF y convertirlo a DataTable
            LeerPDFaDataTableUKNuevo(openFileDialog.FileName, fichero, dtUkPersonas)
        End If
    End Sub

    Function LeerPDFaDataTableDCZ(ByVal pdfPath As String) As DataTable
        ' Crear un DataTable
        Dim dataTable As New DataTable()

        ' Crear columnas en el DataTable (puedes ajustar esto según tu PDF)
        dataTable.Columns.Add("Valores")
        dataTable.Columns.Add("Operario")

        ' ... Agrega más columnas según sea necesario

        ' Crear un lector de PDF
        Dim pdfReader As New PdfReader(pdfPath)
        ' Recorrer las páginas del PDF
        For page As Integer = 1 To pdfReader.NumberOfPages
            ' Obtener el texto de la página
            Dim texto As String = PdfTextExtractor.GetTextFromPage(pdfReader, page)
            ' Buscar la posición de "Liquidas"
            Dim posicionLiquidas As Integer = texto.IndexOf("Liquidas")
            ' Buscar la posición de "©"
            Dim posicionCopyright As Integer = texto.IndexOf("©")

            Dim resultado As String
            resultado = texto.Substring(posicionLiquidas + "Liquidas".Length, posicionCopyright - posicionLiquidas - "Liquidas".Length).Trim()

            ' Contar la cantidad de guiones en la variable resultado
            Dim cantidadGuiones As Integer = resultado.Split("-"c).Length - 1

            ' Recorrer un bucle for desde 0 hasta la cantidad de guiones
            For i As Integer = 0 To cantidadGuiones
                ' Encontrar la posición del guion "-"
                Dim posicionGuion As Integer = resultado.IndexOf("-")

                ' Verificar si se encontró el guion
                If posicionGuion >= 0 Then
                    ' Inicializar la variable para almacenar la posición del primer dígito después del guion
                    Dim posicionDigito As Integer = -1

                    ' Recorrer desde la posición del guion + 1 hasta el final de la cadena
                    For j As Integer = posicionGuion + 1 To resultado.Length - 1
                        ' Verificar si el carácter en la posición actual es un dígito
                        If Char.IsDigit(resultado(j)) Then
                            ' Almacenar la posición del primer dígito
                            posicionDigito = j
                            Exit For
                        End If
                    Next

                    ' Verificar si se encontró un dígito después del guion
                    If posicionDigito >= 0 Then
                        ' Extraer las subcadenas
                        Dim izquierda As String = resultado.Substring(0, posicionGuion).Trim()
                        Dim derecha As String = resultado.Substring(posicionGuion + 1, posicionDigito - posicionGuion - 1).Trim()

                        ' Agregar las subcadenas al DataTable
                        dataTable.Rows.Add(izquierda, derecha)

                        ' Actualizar la variable resultado para continuar desde donde terminó la última búsqueda
                        resultado = resultado.Substring(posicionDigito)

                    Else
                        ' Si no se encontró un dígito después del guion, agregar la última parte al DataTable
                        Dim izquierda As String = resultado.Substring(0, posicionGuion).Trim()
                        Dim derecha As String = resultado.Substring(posicionGuion + 1).Trim()

                        ' Agregar las subcadenas al DataTable
                        dataTable.Rows.Add(izquierda, derecha)

                        ' Salir del bucle ya que estamos al final de la cadena
                        Exit For
                    End If
                End If
            Next
        Next

        Return dataTable
    End Function

    Sub LeerPDFaDataTableUK(ByVal pdfPath As String, ByVal fichero As Integer)
        ruta = pdfPath
        UnificaFichero(fichero)
    End Sub
    Sub LeerPDFaDataTableUKNuevo(ByVal pdfPath As String, ByVal fichero As Integer, ByVal dtUKPersonas As DataTable)
        ruta = pdfPath
        UnificaFichero(fichero)
    End Sub
    Public Sub UnificaFichero(ByVal fichero As Integer)

        Dim pdfReader As New PdfReader(ruta)
        Dim texto_entero As String = PdfTextExtractor.GetTextFromPage(pdfReader, 1)
        Dim palabras As String() = texto_entero.Split(" "c)
        Dim ultimaPalabra As String = palabras(palabras.Length - 1)
        Dim opcion As Integer

        If ultimaPalabra.Contains("6") Then
            opcion = 2
        ElseIf ultimaPalabra.Contains("7") Then
            opcion = 1
        Else
            MsgBox("Informe nuevo, hablar con David Velasco")
            Exit Sub
        End If

        Select Case opcion
            'La opcion 1 es para cuando detecta Employee Name en el algoritmo al leer el pdf
            Case 1
                ' Crear un lector de PDF
                ' Recorrer las páginas del PDF
                For page As Integer = 1 To pdfReader.NumberOfPages
                    If page = pdfReader.NumberOfPages Then
                        Dim texto As String = PdfTextExtractor.GetTextFromPage(pdfReader, page)
                        Dim posicionName As Integer = texto.IndexOf("Employee Name")
                        ' Buscar la posición de "©"
                        Dim total As Integer = texto.IndexOf("Dept")
                        Dim resultado As String
                        resultado = texto.Substring(posicionName + "Employee Name".Length, total - posicionName - "Employee Name".Length).Trim()
                        resultado &= vbCrLf
                        cadenaFinal &= resultado
                    Else
                        ' Obtener el texto de la página
                        Dim texto As String = PdfTextExtractor.GetTextFromPage(pdfReader, page)
                        Dim posicionName As Integer = texto.IndexOf("Employee Name")
                        ' Buscar la posición de "©"
                        Dim posicionCopyright As Integer = texto.IndexOf("©")

                        Dim resultado As String
                        resultado = texto.Substring(posicionName + "Employee Name".Length, posicionCopyright - posicionName - "Employee Name".Length).Trim()
                        resultado &= vbCrLf
                        cadenaFinal &= resultado
                    End If
                Next
            Case 2
                ' Recorrer las páginas del PDF
                For page As Integer = 1 To pdfReader.NumberOfPages
                    If page = pdfReader.NumberOfPages Then
                        Dim texto As String = PdfTextExtractor.GetTextFromPage(pdfReader, page)
                        Dim posicionName As Integer = texto.IndexOf("Class 1A Pension")
                        ' Buscar la posición de "©"
                        Dim total As Integer = texto.IndexOf("Dept")
                        Dim resultado As String
                        resultado = texto.Substring(posicionName + "Class 1A Pension".Length, total - posicionName - "Class 1A Pension".Length).Trim()
                        resultado &= vbCrLf
                        cadenaFinal &= resultado
                    Else
                        ' Obtener el texto de la página
                        Dim texto As String = PdfTextExtractor.GetTextFromPage(pdfReader, page)
                        Dim posicionName As Integer = texto.IndexOf("Class 1A Pension")
                        ' Buscar la posición de "©"
                        Dim posicionCopyright As Integer = texto.IndexOf("©")

                        Dim resultado As String
                        resultado = texto.Substring(posicionName + "Class 1A Pension".Length, posicionCopyright - posicionName - "Class 1A Pension".Length).Trim()
                        resultado &= vbCrLf
                        cadenaFinal &= resultado
                    End If
                Next
            Case 3
            Case Else
                Console.WriteLine("Falta un digito (")
        End Select
    End Sub
    Dim cadenaFinal As String
    Dim ruta As String
    Public Sub GuardaFicheroUkTxt()
        Dim rutaArchivo As String = "N:\100. GESTION\01. A3\00. Pruebas\temp.txt"
        ' Realizar los reemplazos
        cadenaFinal = cadenaFinal.Replace("  ", " ").Replace(", ", ",")
        ' Escribe la cadena en el archivo
        File.WriteAllText(rutaArchivo, cadenaFinal)
    End Sub

    Public Sub darFormaTablaUK(ByRef dtPersonasUK As DataTable)
        'En total 17 columnas + tipo de fichero
        dtPersonasUK.Columns.Add("Diccionario")
        dtPersonasUK.Columns.Add("Operario") 'Aqui se incluye hasta la columna que hay una A o una M
        dtPersonasUK.Columns.Add("Pre tax")
        dtPersonasUK.Columns.Add("Gu Costs")
        dtPersonasUK.Columns.Add("Abstence Pay")
        dtPersonasUK.Columns.Add("Holiday Pay")
        dtPersonasUK.Columns.Add("Pre Tax Pension")
        dtPersonasUK.Columns.Add("Taxable Pay")
        dtPersonasUK.Columns.Add("Tax")
        dtPersonasUK.Columns.Add("Net NI")
        dtPersonasUK.Columns.Add("Post Tax Add")
        dtPersonasUK.Columns.Add("Post Tax Pension")
        dtPersonasUK.Columns.Add("AEO")
        dtPersonasUK.Columns.Add("Students Loans")
        dtPersonasUK.Columns.Add("Net Pay")
        dtPersonasUK.Columns.Add("Net Er NI")
        dtPersonasUK.Columns.Add("Er Pension")
        dtPersonasUK.Columns.Add("Fichero")
    End Sub
    Private Sub bUk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bUk.Click
        cadenaFinal = ""
        'Esta variable determinará el nº de fichero que se ha insertado
        Dim fichero As Integer = 1
        Do
            SeleccionarPDFyLeerDataTableUK(fichero)
            Dim respuesta As DialogResult = MessageBox.Show("¿Deseas cargar algún Excel más?", "Continuar", MessageBoxButtons.YesNo)
            ' Salir del bucle si el usuario responde "No"
            If respuesta = DialogResult.No Then
                Exit Do
            End If
            fichero = fichero + 1
        Loop
        GuardaFicheroUkTxt()
        LeeFicheroYGuardaEnExcel()
    End Sub
    Public Sub LeeFicheroYGuardaEnExcel()
        Dim nombreArchivo As String = "N:\100. GESTION\01. A3\00. Pruebas\temp.txt"

        If File.Exists(nombreArchivo) Then
            ' Lee todas las líneas del archivo y las guarda en un array de String
            Dim lineas As String() = File.ReadAllLines(nombreArchivo)
            File.Delete(nombreArchivo)
            FormaTablaFinal(lineas)
        End If
    End Sub
    Public Sub LeeFicheroYGuardaEnExcelNuevo(ByVal dtUKPersonas As DataTable, ByVal fichero As Integer)
        Dim nombreArchivo As String = "N:\100. GESTION\01. A3\00. Pruebas\temp.txt"

        If File.Exists(nombreArchivo) Then
            ' Lee todas las líneas del archivo y las guarda en un array de String
            Dim lineas As String() = File.ReadAllLines(nombreArchivo)
            File.Delete(nombreArchivo)
            FormaTablaFinalNuevo(lineas, dtUKPersonas, fichero)
        End If
    End Sub

    Public Sub FormaTablaFinal(ByVal lineas As String())
        Dim dtUkPersonas As New DataTable
        darFormaTablaUK(dtUkPersonas)

        For Each fila As String In lineas
            ' Añadir una nueva fila a la nueva tabla
            Dim nuevaFila As DataRow = dtUkPersonas.NewRow()

            nuevaFila("Diccionario") = fila.Substring(0, fila.IndexOf(" "))

            ' Buscar la letra "A" o "M" y extraer el segundo valor (columna "Operario")
            Dim letras() As String = {" A ", ")A ", " M ", ")M"}
            Dim indiceEspacio1 As Integer = fila.IndexOf(" ")

            ' Encontrar la posición de la letra después del primer espacio
            Dim indiceLetra As Integer = -1

            ' Buscar la letra en el array
            For Each letra As String In letras
                indiceLetra = fila.IndexOf(letra, indiceEspacio1)
                If indiceLetra >= 0 Then
                    Exit For ' Salir del bucle si se encontró la letra
                End If
            Next

            ' Verificar si se encontró la letra
            If indiceLetra >= 0 AndAlso indiceLetra > indiceEspacio1 Then
                ' Extraer el segundo valor entre el primer espacio y la letra
                Dim segundoValor As String = fila.Substring(indiceEspacio1 + 1, indiceLetra - indiceEspacio1 - 1 + 3).Trim()
                ' Asignar el segundo valor a la columna "Operario"
                nuevaFila("Operario") = segundoValor
                '--------------------------
                ' Ahora, extraer los valores desde el segundo hasta el próximo espacio y asignarlos a las columnas adicionales
                Dim posInicio As Integer = indiceLetra + 3 ' Para empezar después de la letra y el espacio
                Dim posFin As Integer = fila.IndexOf(" ", posInicio)

                If posFin > posInicio Then
                    ' Iterar sobre las columnas adicionales y asignar valores
                    For Each columna As DataColumn In dtUkPersonas.Columns
                        If columna.ColumnName <> "Diccionario" AndAlso columna.ColumnName <> "Operario" Then
                            ' Extraer el valor entre posInicio y posFin
                            Dim valorColumna As String = fila.Substring(posInicio, posFin - posInicio).Trim()
                            nuevaFila(columna.ColumnName) = valorColumna
                            ' Actualizar posInicio para el próximo ciclo
                            posInicio = posFin + 1

                            If posInicio < fila.Length Then
                                posFin = fila.IndexOf(" ", posInicio)

                                ' Si no se encuentra un espacio, establecer posFin al final de la cadena
                                If posFin = -1 Then
                                    posFin = fila.Length
                                End If

                                ' Resto del código...
                            End If
                            If posFin = -1 Then
                                posFin = fila.Length ' Para manejar el último valor en la fila
                            End If
                        End If
                    Next
                End If
            Else
                ' Manejar el caso donde no se encuentra ninguna letra del array
                MessageBox.Show("No se encontró ninguna letra 'A' o 'M' en la fila.")
            End If


            ' Agregar la nueva fila a la nueva tabla
            dtUkPersonas.Rows.Add(nuevaFila)
        Next

        GeneraExcelUKUnificado(dtUkPersonas)
    End Sub

    Public Sub FormaTablaFinalNuevo(ByVal lineas As String(), ByVal dtUKPersonas As DataTable, ByVal fichero As Integer)

        For Each fila As String In lineas
            ' Añadir una nueva fila a la nueva tabla
            Dim nuevaFila As DataRow = dtUKPersonas.NewRow()

            nuevaFila("Diccionario") = fila.Substring(0, fila.IndexOf(" "))
            nuevaFila("Fichero") = fichero
            ' Buscar la letra "A" o "M" y extraer el segundo valor (columna "Operario")
            Dim letras() As String = {" A ", ")A ", " M ", ")M"}
            Dim indiceEspacio1 As Integer = fila.IndexOf(" ")

            ' Encontrar la posición de la letra después del primer espacio
            Dim indiceLetra As Integer = -1

            ' Buscar la letra en el array
            For Each letra As String In letras
                indiceLetra = fila.IndexOf(letra, indiceEspacio1)
                If indiceLetra >= 0 Then
                    Exit For ' Salir del bucle si se encontró la letra
                End If
            Next

            ' Verificar si se encontró la letra
            If indiceLetra >= 0 AndAlso indiceLetra > indiceEspacio1 Then
                ' Extraer el segundo valor entre el primer espacio y la letra
                Dim segundoValor As String = fila.Substring(indiceEspacio1 + 1, indiceLetra - indiceEspacio1 - 1 + 3).Trim()
                ' Asignar el segundo valor a la columna "Operario"
                nuevaFila("Operario") = segundoValor
                '--------------------------
                ' Ahora, extraer los valores desde el segundo hasta el próximo espacio y asignarlos a las columnas adicionales
                Dim posInicio As Integer = indiceLetra + 3 ' Para empezar después de la letra y el espacio
                Dim posFin As Integer = fila.IndexOf(" ", posInicio)

                If posFin > posInicio Then
                    ' Iterar sobre las columnas adicionales y asignar valores
                    For Each columna As DataColumn In dtUKPersonas.Columns
                        If columna.ColumnName <> "Diccionario" AndAlso columna.ColumnName <> "Operario" AndAlso columna.ColumnName <> "Fichero" Then
                            ' Extraer el valor entre posInicio y posFin
                            Dim valorColumna As String = fila.Substring(posInicio, posFin - posInicio).Trim()
                            nuevaFila(columna.ColumnName) = valorColumna
                            ' Actualizar posInicio para el próximo ciclo
                            posInicio = posFin + 1

                            If posInicio < fila.Length Then
                                posFin = fila.IndexOf(" ", posInicio)

                                ' Si no se encuentra un espacio, establecer posFin al final de la cadena
                                If posFin = -1 Then
                                    posFin = fila.Length
                                End If

                                ' Resto del código...
                            End If
                            If posFin = -1 Then
                                posFin = fila.Length ' Para manejar el último valor en la fila
                            End If
                        End If
                    Next
                End If
            Else
                ' Manejar el caso donde no se encuentra ninguna letra del array
                MessageBox.Show("No se encontró ninguna letra 'A' o 'M' en la fila.")
            End If


            ' Agregar la nueva fila a la nueva tabla
            dtUKPersonas.Rows.Add(nuevaFila)
        Next
    End Sub

    Public Sub GeneraExcelUKUnificado(ByVal dtUkPersonas As DataTable)
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial

        Dim saveFileDialog1 As New SaveFileDialog()

        For Each fila As DataRow In dtUkPersonas.Rows
            For Each columna As DataColumn In dtUkPersonas.Columns
                ' Si el valor es de tipo Double, formatearlo con coma en lugar de punto
                Try
                    fila(columna) = (DirectCast(fila(columna), String).Replace(".", ","))
                Catch ex As Exception
                    fila(columna) = ""
                End Try

                Try
                    fila(columna) = (DirectCast(fila(columna), String).Replace("(", "-"))
                Catch ex As Exception
                    fila(columna) = ""
                End Try

                Try
                    fila(columna) = (DirectCast(fila(columna), String).Replace(")", ""))
                Catch ex As Exception
                    fila(columna) = ""
                End Try


                ' Si la columna es "Diccionario", eliminar letras y dejar solo dígitos
                If columna.ColumnName = "Diccionario" Then
                    Dim valorOriginal As String = DirectCast(fila(columna), String)
                    Dim valorSinLetras As String = New String(valorOriginal.Where(Function(c) Char.IsDigit(c)).ToArray())
                    fila(columna) = valorSinLetras
                End If
            Next
        Next
        ' Configurar propiedades del cuadro de diálogo

        saveFileDialog1.Filter = "Archivos de texto|*.xlsx|Todos los archivos|*.*"
        saveFileDialog1.Title = "Guardar archivo"

        ' Mostrar el cuadro de diálogo y verificar si el usuario hizo clic en "Guardar"
        If saveFileDialog1.ShowDialog() = DialogResult.OK Then
            ' Obtener la ruta seleccionada por el usuario
            Dim rutaArchivo As String = saveFileDialog1.FileName

            'Verificar si el archivo existe.
            If File.Exists(rutaArchivo) Then
                'Si el archivo existe, eliminarlo.
                File.Delete(rutaArchivo)
            End If

            Dim result As DialogResult = MessageBox.Show("¿Deseas ordenar los registros?", "Confirmacion", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

            ' Check the user's choice
            If result = DialogResult.Yes Then
                dtUkPersonas.DefaultView.Sort = "Diccionario ASC" ' Ajusta "ColumnaL" al nombre real de la columna
            End If


            Dim dtUkPersonasOrdenado = dtUkPersonas.DefaultView.ToTable()

            Using package As New ExcelPackage(rutaArchivo)

                ' Crear una hoja de cálculo y obtener una referencia a ella.
                Dim worksheet = package.Workbook.Worksheets.Add("1")

                ' Copiar los datos de la DataTable a la hoja de cálculo.
                worksheet.Cells("A1").LoadFromDataTable(dtUkPersonasOrdenado, True)

                Dim columnaA As ExcelRange = worksheet.Cells("A2:A" & worksheet.Dimension.End.Row)
                columnaA.Style.Numberformat.Format = "@"

                ' Aplicar formato negrita a la fila 1
                Dim fila1 As ExcelRange = worksheet.Cells(1, 1, 1, worksheet.Dimension.End.Column)
                fila1.Style.Font.Bold = True

                For row As Integer = 2 To worksheet.Dimension.End.Row
                    For col As Integer = 2 To 17
                        Dim valorCelda As String = worksheet.Cells(row, col).Text
                        Dim valorNumerico As Double

                        ' Intentar convertir el valor a un número
                        If Double.TryParse(valorCelda, valorNumerico) Then
                            ' Si el valor está entre paréntesis, multiplicarlo por -1
                            ' Asignar el valor numérico a la celda
                            worksheet.Cells(row, col).Value = valorNumerico
                        End If
                    Next
                Next
                Dim rangoMoneda As ExcelRange = worksheet.Cells("B2:Q" & worksheet.Dimension.End.Row)
                rangoMoneda.Style.Numberformat.Format = "#,##0.00£"

                ' Agregar un filtro a la primera fila
                worksheet.Cells("A1:" & GetExcelColumnName(worksheet.Dimension.End.Column) & "1").AutoFilter = True
                worksheet.Column(2).Width = 30 : worksheet.Column(3).Width = 12 : worksheet.Column(4).Width = 12
                worksheet.Column(5).Width = 12 : worksheet.Column(6).Width = 12 : worksheet.Column(7).Width = 12 : worksheet.Column(8).Width = 12
                worksheet.Column(9).Width = 12 : worksheet.Column(10).Width = 12 : worksheet.Column(11).Width = 12
                worksheet.Column(12).Width = 12 : worksheet.Column(13).Width = 12 : worksheet.Column(14).Width = 12
                worksheet.Column(15).Width = 12 : worksheet.Column(16).Width = 12 : worksheet.Column(17).Width = 12

                ' Congelar la primera fila
                worksheet.View.FreezePanes(2, 1)

                ' Guardar el archivo de Excel.
                package.Save()
            End Using

            MsgBox("Fichero creado correctamente en N:\10. AUXILIARES\00. EXPERTIS\02. A3\")
        End If
    End Sub

    Private Sub bNO_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bNO.Click
        Dim dtPersonasPortugal As DataTable = SeleccionarPDFyLeerDataTableNO()
        GeneraExcel(dtPersonasPortugal)
    End Sub

    Function SeleccionarPDFyLeerDataTableNO() As DataTable
        ' Crear un cuadro de diálogo para seleccionar el archivo PDF
        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Filter = "Archivos PDF|*.pdf"
        openFileDialog.Title = "Selecciona un archivo PDF"

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            ' Llamada a la función para leer el PDF y convertirlo a DataTable
            Return LeerPDFaDataTableNO(openFileDialog.FileName)
        Else
            ' El usuario canceló la selección
            Return Nothing
        End If
    End Function
    Dim accrued_holiday_pay As String = "10,20"
    Dim accrued_holiday_pay_over_60_old As String = "2,30"
    Dim accrued_EC_holiday As String = "14,10"
    Dim employer_contribution As String = "14,10"



    Function LeerPDFaDataTableNO(ByVal pdfPath As String) As DataTable
        ' Crear un DataTable
        Dim dataTable As New DataTable()
        dataTable.Columns.Add("Diccionario")
        dataTable.Columns.Add("Operario")
        dataTable.Columns.Add("TAX")
        dataTable.Columns.Add("Cash Benefit (gross)")
        dataTable.Columns.Add("Food Allowance")
        dataTable.Columns.Add("Payment in kind (gross)")
        dataTable.Columns.Add("Other payouts")
        dataTable.Columns.Add("Other deductions")
        dataTable.Columns.Add("GROSS OVER TAX")
        dataTable.Columns.Add("NET TO PAY")
        dataTable.Columns.Add("Accrued holiday pay")
        dataTable.Columns.Add("Accrued holiday pay over 60 yr old")
        dataTable.Columns.Add("Accrued EC of holiday pay")
        dataTable.Columns.Add("Employer's Contribution")
        dataTable.Columns.Add("Withholding taxes")
        dataTable.Columns.Add("COSTE EMPRESA")

        'AÑADO LA NUEVA LINEA DE ANGEL
        Dim nuevaFila As DataRow = dataTable.NewRow()

        ' Crear un lector de PDF
        Dim pdfReader As New PdfReader(pdfPath)
        ' Recorrer las páginas del PDF
        For page As Integer = 1 To pdfReader.NumberOfPages
            ' Obtener el texto de la página
            Dim texto As String = PdfTextExtractor.GetTextFromPage(pdfReader, page)
            Dim tax As Double
            Dim accruedholidaypay As Double
            Dim accruedEC As Double
            Dim paymentinkind As Double
            Dim deductions As Double
            Dim cashBenefit As Double
            Dim impuestos As Double
            Dim contribucion As Double
            Dim old As Double
            Dim bruto As Double
            ' Añadir una nueva fila a la nueva tabla
            nuevaFila = dataTable.NewRow()

            ' Asignar los valores a las columnas correspondientes
            nuevaFila("Diccionario") = devuelveDiccionarioNO(texto)
            nuevaFila("Operario") = devuelveOperarioNO(texto)
            tax = devuelveTAX(texto)

            paymentinkind = Nz(devuelvePaymentInKind(texto), 0)
            deductions = Nz(devuelveDeductions(texto), 0)
            cashBenefit = devuelveCashBenefit(texto)
            nuevaFila("TAX") = tax
            nuevaFila("Cash Benefit (gross)") = cashBenefit

            Dim foodAllowance As Double
            foodAllowance = devuelveFoodAllowance(texto)

            Dim other_payouts As Double
            other_payouts = devuelveOtherPayouts(texto)

            nuevaFila("Food Allowance") = foodAllowance
            nuevaFila("Other payouts") = other_payouts
            nuevaFila("Payment in kind (gross)") = paymentinkind

            impuestos = (cashBenefit + paymentinkind) * (tax / 100)

            Dim grossOverTAX As Double
            grossOverTAX = cashBenefit + paymentinkind
            nuevaFila("GROSS OVER TAX") = grossOverTAX
            
            If deductions >= 0 Then
                'deductions = deductions * -1
                deductions = deductions * -1
            Else
                deductions = Math.Abs(deductions)
            End If
            nuevaFila("Other deductions") = deductions

            Dim netToPay As Double
            netToPay = (cashBenefit - (cashBenefit * (tax / 100))) + other_payouts

            If other_payouts > 0 Then
                netToPay = netToPay - deductions
            End If

            If deductions < 0 Then
                netToPay = netToPay + deductions
            End If

            If other_payouts > 0 And deductions < 0 Then
                netToPay = netToPay + deductions
            End If
            nuevaFila("NET TO PAY") = netToPay


            'David V 09/02/24
            'HOLIDAY PAY = (CASH BENEFIT - FOOD ALLOWANCE)* 0.102
            accruedholidaypay = (cashBenefit - foodAllowance) * 0.102
            nuevaFila("Accrued holiday pay") = accruedholidaypay
            old = DevuelvePagoPorViejo(devuelveDiccionarioNO(texto))
            Dim viejos As Double = 0
            If old = 1 Then
                viejos = (cashBenefit - foodAllowance) * 0.023
                nuevaFila("Accrued holiday pay over 60 yr old") = viejos
            End If
            accruedEC = (accruedholidaypay + viejos) * 0.141
            nuevaFila("Accrued EC of holiday pay") = accruedEC
            contribucion = (cashBenefit + paymentinkind) * 0.141
            nuevaFila("Employer's Contribution") = contribucion
            nuevaFila("Withholding taxes") = impuestos
            nuevaFila("COSTE EMPRESA") = cashBenefit + accruedholidaypay + viejos + accruedEC + contribucion

            ' Agregar la nueva fila a la nueva tabla
            dataTable.Rows.Add(nuevaFila)
        Next

        Return dataTable
    End Function

    Public Function devuelveDiccionarioNO(ByVal texto As String) As String
        Dim buscar As String = "Employee No."
        Dim startIndex As Integer = texto.IndexOf(buscar)

        ' Obtener las 5 posiciones después de "Employee No."
        Dim resultado As String = texto.Substring(startIndex + buscar.Length, 7)

        Dim soloNumeros As String = New String(resultado.Where(Function(c) Char.IsDigit(c)).ToArray())

        Return soloNumeros
    End Function
    Public Function DevuelvePagoPorViejo(ByVal diccionario As String) As Double
        Dim dt As New DataTable
        Dim f As New Filter
        f.Add("Diccionario", FilterOperator.Equal, diccionario)

        dt = New BE.DataEngine().Filter(DB_NO & "..frmMntoOperario", f)

        If dt.Rows.Count = 0 Then
            MsgBox("El operario con diccionario " & diccionario & " no está dado de alta en Expertis.")
            Return 0
        End If
        Dim fecha_nacimiento As String
        Try
            fecha_nacimiento = dt.Rows(0)("Fecha_Nacimiento")
        Catch ex As Exception
            Return 0
        End Try


        If fecha_nacimiento.Length = 0 Then
            Return 0
        Else
            Dim fechaNacimiento As DateTime
            If DateTime.TryParse(fecha_nacimiento, fechaNacimiento) Then
                ' Calcular la edad
                Dim edad As Integer = CalcularEdad(fechaNacimiento)

                ' Verificar si la edad supera los 60 años
                If edad > 60 Then
                    Return 1
                Else
                    Return 0
                End If
            End If
        End If

    End Function

    Function CalcularEdad(ByVal fechaNacimiento As DateTime) As Integer
        ' Calcular la diferencia en años entre la fecha de nacimiento y la fecha actual
        Dim edad As Integer = DateTime.Now.Year - fechaNacimiento.Year

        ' Ajustar la edad si aún no ha llegado el cumpleaños en este año
        If DateTime.Now < fechaNacimiento.AddYears(edad) Then
            edad -= 1
        End If

        Return edad
    End Function
    Public Function devuelveOperarioNO(ByVal texto As String) As String
        Dim startString As String = "931198114"
        Dim endString As String = "PAYSLIP"

        ' Encontrar las posiciones de inicio y fin
        Dim startIndex As Integer = texto.IndexOf(startString) + startString.Length
        Dim endIndex As Integer = texto.IndexOf(endString)

        ' Extraer la subcadena
        Dim operario As String = texto.Substring(startIndex, endIndex - startIndex)

        Return operario.Trim
    End Function

    Public Function devuelveTAX(ByVal texto As String) As String
        Dim startString As String = "Percentage deduction"
        ' Encontrar la posición de la cadena de búsqueda
        Dim startIndex As Integer = texto.IndexOf(startString)
        ' Buscar la posición del primer "%" después de "Percentage deduction" a partir del índice donde se encuentra "Percentage deduction"
        Dim porcentajeIndex As Integer = texto.IndexOf("%", startIndex + startString.Length)
        ' Obtener la subcadena deseada
        Dim porcentaje As String = texto.Substring(startIndex + startString.Length, porcentajeIndex - (startIndex + startString.Length) + 1)

        Return porcentaje.Trim.Replace(" %", "")
    End Function

    Public Function devuelveWages(ByVal texto As String) As String
        ' Cadena de búsqueda
        Dim searchString As String = "Fixed salary "
        ' Encontrar la posición de la cadena de búsqueda
        Dim startIndex As Integer = texto.IndexOf(searchString)
        ' Encontrar la posición de la primera coma después de "Fixed salary"
        Dim comaIndex As Integer = texto.IndexOf(",", startIndex)
        ' Obtener la subcadena deseada
        Dim resultado As String = texto.Substring(startIndex + searchString.Length, comaIndex - (startIndex + searchString.Length) + 3)

        'Comprobar si tiene "Wage deduction for holiday"
        'Si se encuentra se resta el valor
        searchString = "Wage deduction for holiday "

        If texto.Contains(searchString) Then
            Dim holiday As String
            ' Encontrar la posición de la cadena de búsqueda
            startIndex = texto.IndexOf(searchString)
            Dim guionIndex As Integer = texto.IndexOf("-", startIndex)
            ' Buscar la posición de la primera coma "," después del guión "-"
            comaIndex = texto.IndexOf(",", guionIndex)
            ' Obtener la subcadena deseada
            holiday = texto.Substring(guionIndex + 1, comaIndex - (guionIndex + 1) + 3)
            'Y ahora hago la resta
            Dim resultadoFinal As Double
            resultadoFinal = resultado.Replace(" ", "") - holiday.Replace(" ", "")
            Return resultadoFinal
        Else
            Return resultado.Trim.Replace(" ", "")
        End If

    End Function

    Public Function devuelveFood(ByVal texto As String) As String
        ' Cadena de búsqueda
        Dim searchString As String = "Monthly Food Allowance "
        ' Encontrar la posición de la cadena de búsqueda
        Dim startIndex As Integer = texto.IndexOf(searchString)
        ' Encontrar la posición de la primera coma después de "Fixed salary"
        Dim comaIndex As Integer = texto.IndexOf(",", startIndex)
        ' Obtener la subcadena deseada
        Dim resultado As String = texto.Substring(startIndex + searchString.Length, comaIndex - (startIndex + searchString.Length) + 3)
        Return resultado.Trim.Replace(" ", "")
    End Function


    Public Function devuelveOtherPayouts(ByVal texto As String) As String
        ' Cadena de búsqueda
        Dim searchString As String = "Other payouts "
        ' Encontrar la posición de la cadena de búsqueda
        Dim startIndex As Integer = texto.IndexOf(searchString)
        If startIndex = "-1" Then
            Return 0
        End If
        ' Encontrar la posición de la primera coma después de "Fixed salary"
        Dim comaIndex As Integer = texto.IndexOf(",", startIndex)
        ' Obtener la subcadena deseada
        Dim resultado As String = texto.Substring(startIndex + searchString.Length, comaIndex - (startIndex + searchString.Length) + 3)
        Return resultado.Trim.Replace(" ", "")
    End Function

    Public Function devuelveFoodAllowance(ByVal texto As String) As String
        ' Cadena de búsqueda
        Dim searchString As String = "Monthly Food Allowance "
        ' Encontrar la posición de la cadena de búsqueda
        Dim startIndex As Integer = texto.IndexOf(searchString)
        If startIndex = "-1" Then
            Return 0
        End If
        ' Encontrar la posición de la primera coma después de "Fixed salary"
        Dim comaIndex As Integer = texto.IndexOf(",", startIndex)
        ' Obtener la subcadena deseada
        Dim resultado As String = texto.Substring(startIndex + searchString.Length, comaIndex - (startIndex + searchString.Length) + 3)
        Return resultado.Trim.Replace(" ", "")
    End Function

    Public Function devuelvePaymentInKind(ByVal texto As String) As String
        ' Cadena de búsqueda
        Dim searchString As String = "Payment in kind "
        ' Encontrar la posición de la cadena de búsqueda
        Dim startIndex As Integer = texto.IndexOf(searchString)
        If startIndex = "-1" Then
            Return 0
        End If
        ' Encontrar la posición de la primera coma después de "Fixed salary"
        Dim comaIndex As Integer = texto.IndexOf(",", startIndex)
        ' Obtener la subcadena deseada
        Dim resultado As String = texto.Substring(startIndex + searchString.Length, comaIndex - (startIndex + searchString.Length) + 3)
        Return resultado.Trim.Replace(" ", "")
    End Function

    Public Function devuelveDeductions(ByVal texto As String) As String
        ' Cadena de búsqueda
        Dim searchString As String = "Other deductions "
        ' Encontrar la posición de la cadena de búsqueda
        Dim startIndex As Integer = texto.IndexOf(searchString)
        If startIndex = "-1" Then
            Return 0
        End If
        ' Encontrar la posición de la primera coma después de "Fixed salary"
        Dim comaIndex As Integer = texto.IndexOf(",", startIndex)
        ' Obtener la subcadena deseada
        Dim resultado As String = texto.Substring(startIndex + searchString.Length, comaIndex - (startIndex + searchString.Length) + 3)
        Return resultado.Trim.Replace(" ", "")
    End Function

    Public Function devuelvePayBasis(ByVal texto As String) As String
        ' Cadena de búsqueda
        Dim searchString As String = "Holiday pay basis "
        ' Encontrar la posición de la cadena de búsqueda
        Dim startIndex As Integer = texto.IndexOf(searchString)
        If startIndex = "-1" Then
            Return 0
        End If
        ' Encontrar la posición de la primera coma después de "Fixed salary"
        Dim comaIndex As Integer = texto.IndexOf(",", startIndex)
        ' Obtener la subcadena deseada
        Dim resultado As String = texto.Substring(startIndex + searchString.Length, comaIndex - (startIndex + searchString.Length) + 3)
        Return resultado.Trim.Replace(" ", "")
    End Function

    Public Function devuelveCashBenefit(ByVal texto As String) As String
        ' Cadena de búsqueda
        Dim searchString As String = "Cash benefit "
        ' Encontrar la posición de la cadena de búsqueda
        Dim startIndex As Integer = texto.IndexOf(searchString)
        If startIndex = "-1" Then
            Return 0
        End If
        ' Encontrar la posición de la primera coma después de "Fixed salary"
        Dim comaIndex As Integer = texto.IndexOf(",", startIndex)
        ' Obtener la subcadena deseada
        Dim resultado As String = texto.Substring(startIndex + searchString.Length, comaIndex - (startIndex + searchString.Length) + 3)
        Return resultado.Trim.Replace(" ", "")
    End Function
    Public Function devuelveComplementos(ByVal texto As String) As Double
        Dim overtime As Double = 0
        Dim bonus As Double = 0
        Dim other_complements As Double = 0

        If texto.Contains("Overtime") Then
            overtime = devuelveOvertime(texto)
        End If
        If texto.Contains("Bonus") Then
            bonus = devuelveBonus(texto)
        End If

        If texto.Contains("Other supplements") Then
            other_complements = devuelveOtrosComplementos(texto)
        End If

        Return (overtime + bonus + other_complements)
    End Function

    Public Function devuelveOvertime(ByVal texto As String) As Double
        'Overtime va a multiplicar 20 * 333,68
        Dim valor As String

        ' Cadena de búsqueda
        Dim searchString As String = "Overtime"

        ' Encontrar la posición de la cadena de búsqueda
        Dim startIndex As Integer = texto.IndexOf(searchString)

        ' Buscar la posición del carácter "%" después de "Overtime"
        Dim porcentajeIndex As Integer = texto.IndexOf("% ", startIndex)

        ' Buscar la posición del siguiente espacio después de la coma
        Dim espacioIndex As Integer = texto.IndexOf(" ", porcentajeIndex + 2)

        ' Obtener la subcadena deseada
        Dim porcentaje As String = texto.Substring(porcentajeIndex + 1, espacioIndex - (porcentajeIndex + 1))

        '--------------OBTENGO SEGUNDO PARAMETRO

        startIndex = texto.IndexOf(porcentaje.Trim)
        Dim valorIndex As Integer = texto.IndexOf(" ", startIndex)
        espacioIndex = texto.IndexOf(" ", valorIndex + 2)
        valor = texto.Substring(valorIndex + 1, espacioIndex - (valorIndex + 1))

        Dim total As Double
        total = (porcentaje.Trim.Replace(" ", "") * valor.Replace(" ", ""))
        Return total
    End Function

    Public Function devuelveBonus(ByVal texto As String) As String
        ' Cadena de búsqueda
        Dim searchString As String = "Bonus "
        ' Encontrar la posición de la cadena de búsqueda
        Dim startIndex As Integer = texto.IndexOf(searchString)
        ' Buscar la posición del carácter "," después de "Bonus "
        Dim comaIndex As Integer = texto.IndexOf(",", startIndex)

        ' Obtener la subcadena deseada
        Dim valor As String = texto.Substring(startIndex + searchString.Length, comaIndex - (startIndex + searchString.Length) + 3)

        Return valor.Trim.Replace(" ", "")
    End Function

    Public Function devuelveOtrosComplementos(ByVal texto As String) As String
        Dim searchString As String = "Other supplements "
        ' Encontrar la posición de la cadena de búsqueda
        Dim startIndex As Integer = texto.IndexOf(searchString)
        ' Buscar la posición del carácter "," después de "Other supplements "
        Dim comaIndex As Integer = texto.IndexOf(",", startIndex)

        ' Obtener la subcadena deseada
        Dim valor As String = texto.Substring(startIndex + searchString.Length, comaIndex - (startIndex + searchString.Length) + 3)

        Return valor.Trim.Replace(" ", "")
    End Function

    Public Sub GeneraExcel(ByVal dtFinal As DataTable)

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial

        Dim saveFileDialog1 As New SaveFileDialog()
        saveFileDialog1.Filter = "Archivos de texto|*.xlsx|Todos los archivos|*.*"
        saveFileDialog1.Title = "Guardar archivo"

        ' Mostrar el cuadro de diálogo y verificar si el usuario hizo clic en "Guardar"
        If saveFileDialog1.ShowDialog() = DialogResult.OK Then
            ' Obtener la ruta seleccionada por el usuario
            Dim rutaArchivo As String = saveFileDialog1.FileName
            Dim ruta As New FileInfo(rutaArchivo)
            Dim rutaCadena As String = ""
            rutaCadena = ruta.FullName

            'Verificar si el archivo existe.
            If File.Exists(rutaCadena) Then
                'Si el archivo existe, eliminarlo.
                File.Delete(rutaCadena)
            End If

            Using package As New ExcelPackage(ruta)
                ' Crear una hoja de cálculo y obtener una referencia a ella.
                Dim worksheet = package.Workbook.Worksheets.Add("1")

                ' Copiar los datos de la DataTable a la hoja de cálculo.
                worksheet.Cells("A1").LoadFromDataTable(dtFinal, True)

                ' Aplicar formato negrita a la fila 1
                Dim fila1 As ExcelRange = worksheet.Cells(1, 1, 1, worksheet.Dimension.End.Column)
                fila1.Style.Font.Bold = True

                For row As Integer = 2 To worksheet.Dimension.End.Row
                    For col As Integer = 3 To 20
                        Dim valorCelda As String = worksheet.Cells(row, col).Text
                        Dim valorNumerico As Double

                        If Double.TryParse(valorCelda, valorNumerico) Then
                            ' Redondear el valor numérico a dos decimales
                            valorNumerico = Math.Round(valorNumerico, 2)

                            ' Si el valor está entre paréntesis, multiplicarlo por -1
                            ' Asignar el valor numérico redondeado a la celda
                            worksheet.Cells(row, col).Value = valorNumerico
                        End If
                    Next
                Next

                ' Establecer el formato de moneda para la columna N
                worksheet.Column(16).Width = 18

                Dim rangoMoneda As ExcelRange = worksheet.Cells("P2:P" & worksheet.Dimension.End.Row)
                rangoMoneda.Style.Numberformat.Format = "_-[$NOK] * #,##0.00_-;_-[$NOK] * -#,##0.00_-;_-[$NOK] * ""-""??_-;_-@_-"

                ' Establecer el color de fondo de la columna H a un amarillo claro
                Dim rangoColumnaH As ExcelRange = worksheet.Cells(2, 9, worksheet.Dimension.End.Row, 9) ' Columna H es la 8
                rangoColumnaH.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid
                rangoColumnaH.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 255, 153)) ' Amarillo claro

                ' Establecer el color de fondo de la columna H a un amarillo claro
                Dim rangoColumnaI As ExcelRange = worksheet.Cells(2, 10, worksheet.Dimension.End.Row, 10) ' Columna I es la 9
                rangoColumnaI.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid
                rangoColumnaI.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 181, 82)) ' Amarillo claro

                ' Establecer el color de fondo de la columna N a verde
                Dim rangoColumnaN As ExcelRange = worksheet.Cells(2, 16, worksheet.Dimension.End.Row, 16) ' Columna O es la 15
                rangoColumnaN.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid
                rangoColumnaN.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(144, 238, 144)) ' Utilizando un tono de verde claro

                Dim dtExplicacionNo As DataTable = devuelveExplicacionNO()
                Dim worksheet2 = package.Workbook.Worksheets.Add("2")

                ' Copiar los datos de la DataTable a la hoja de cálculo.
                worksheet2.Cells("A1").LoadFromDataTable(dtExplicacionNo, True)
                worksheet2.Column(1).Width = 30
                worksheet2.Column(2).Width = 40

                Dim fila11 As ExcelRange = worksheet2.Cells(1, 1, 1, worksheet2.Dimension.End.Column)
                fila11.Style.Font.Bold = True

                ' Congelar la primera fila
                worksheet.View.FreezePanes(2, 1)

                ' Guardar el archivo de Excel.
                package.Save()
            End Using

            MsgBox("Excel generado correctamente")
        End If


    End Sub
    Public Function devuelveExplicacionNO() As DataTable
        Dim dtExplicacionNoruega As New DataTable()
        dtExplicacionNoruega.Columns.Add("VALOR")
        dtExplicacionNoruega.Columns.Add("EXPLICACION")
        Dim nuevaFil As DataRow = dtExplicacionNoruega.NewRow()
        nuevaFil("VALOR") = "CASH BENEFIT"
        nuevaFil("EXPLICACION") = "NOMINA"
        dtExplicacionNoruega.Rows.Add(nuevaFil)
        '---
        nuevaFil = dtExplicacionNoruega.NewRow()
        nuevaFil("VALOR") = "PAYMENT IN KIND"
        nuevaFil("EXPLICACION") = "NOMINA"
        dtExplicacionNoruega.Rows.Add(nuevaFil)
        '---
        nuevaFil = dtExplicacionNoruega.NewRow()
        nuevaFil("VALOR") = "FOOD ALLOWANCE + OTHER PAYOUTS"
        nuevaFil("EXPLICACION") = "NOMINA"
        dtExplicacionNoruega.Rows.Add(nuevaFil)
        '---
        nuevaFil = dtExplicacionNoruega.NewRow()
        nuevaFil("VALOR") = "GROSS OVER TAX ="
        nuevaFil("EXPLICACION") = "CASH BENEFIT + PAYMENT IN KIND"
        dtExplicacionNoruega.Rows.Add(nuevaFil)
        '---
        nuevaFil = dtExplicacionNoruega.NewRow()
        nuevaFil("VALOR") = "NET TO PAY = "
        nuevaFil("EXPLICACION") = "(CASH BENEFIT - (CASH BENEFIT * (tax / 100))) + OTHER PAYOUTS"
        dtExplicacionNoruega.Rows.Add(nuevaFil)
        '---
        nuevaFil = dtExplicacionNoruega.NewRow()
        nuevaFil("VALOR") = "ACCRUED HOLIDAY PAY ="
        nuevaFil("EXPLICACION") = "(CASH BENEFIT - FOOD ALLOWANCE)* 0,102"
        dtExplicacionNoruega.Rows.Add(nuevaFil)
        '---
        nuevaFil = dtExplicacionNoruega.NewRow()
        nuevaFil("VALOR") = "ACCRUED HOLIDAY > 60 OLD ="
        nuevaFil("EXPLICACION") = "(CASH BENEFIT - FOOD ALLOWANCE) * 0,023"
        dtExplicacionNoruega.Rows.Add(nuevaFil)
        '---
        nuevaFil = dtExplicacionNoruega.NewRow()
        nuevaFil("VALOR") = "ACCRUED EC ="
        nuevaFil("EXPLICACION") = "(ACCRUED HOLIDAY PAY + ACCRUED HOLIDAY > 60 OLD ) * 0,141"
        dtExplicacionNoruega.Rows.Add(nuevaFil)
        '---
        nuevaFil = dtExplicacionNoruega.NewRow()
        nuevaFil("VALOR") = "EMPLOYER'S CONTIBUTION ="
        nuevaFil("EXPLICACION") = "(CASH BENEFIT + PAYMENT IN KIND) * 0,141"
        dtExplicacionNoruega.Rows.Add(nuevaFil)
        '---
        nuevaFil = dtExplicacionNoruega.NewRow()
        nuevaFil("VALOR") = "WITHOLDING TAXES ="
        nuevaFil("EXPLICACION") = "(CASH BENEFIT + PAYMENT IN KIND)* (TAX /100)"
        dtExplicacionNoruega.Rows.Add(nuevaFil)
        '---
        nuevaFil = dtExplicacionNoruega.NewRow()
        nuevaFil("VALOR") = "COSTE EMPRESA ="
        nuevaFil("EXPLICACION") = "(BRUTO + ACCRUED HOLIDAY PAY + ACCRUED HOLIDAY > 60 OLD + ACCRUED EC + EMPLOYER'S CONTIBUTION)"
        dtExplicacionNoruega.Rows.Add(nuevaFil)

        nuevaFil = dtExplicacionNoruega.NewRow()
        nuevaFil("VALOR") = "----"
        nuevaFil("EXPLICACION") = "----"
        dtExplicacionNoruega.Rows.Add(nuevaFil)

        nuevaFil = dtExplicacionNoruega.NewRow()
        nuevaFil("VALOR") = "IF OTHER_PAYOUTS>0"
        nuevaFil("EXPLICACION") = "NET TO PAY = NET TO PAY - DEDUCTIONS"
        dtExplicacionNoruega.Rows.Add(nuevaFil)

        nuevaFil = dtExplicacionNoruega.NewRow()
        nuevaFil("VALOR") = "IF DEDUCTIONS <0"
        nuevaFil("EXPLICACION") = "NET TO PAY = NET TO PAY + DEDUCTIONS"
        dtExplicacionNoruega.Rows.Add(nuevaFil)
        Return dtExplicacionNoruega
    End Function

    Private Sub bMatriz_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '6. TABLA DE HORAS POR PERSONAS
        'Dim dtHorasPersonasDias As New DataTable
        'FormaTablaMatriz(dtHorasPersonasDias)

        Dim frm As New frmInformeFecha
        frm.ShowDialog()
        Dim Fecha1 As String : Dim Fecha2 As String
        Fecha1 = frm.fecha1 : Fecha2 = frm.fecha2
        Dim dtHoras As New DataTable
        Dim f As New Filter
        f.Add("FechaInicio", FilterOperator.GreaterThanOrEqual, Fecha1)
        f.Add("FechaInicio", FilterOperator.LessThanOrEqual, Fecha2)

        dtHoras = New BE.DataEngine().Filter("vUniontbObraModControl", f, , "Empresa asc")

        ' Crear la estructura de la tabla dtHorasPersonasDias
        Dim dtHorasPersonasDias As New DataTable()
        dtHorasPersonasDias.Columns.Add("Empresa")
        dtHorasPersonasDias.Columns.Add("IDGET")
        dtHorasPersonasDias.Columns.Add("IDOperario")
        dtHorasPersonasDias.Columns.Add("DescOperario")
        dtHorasPersonasDias.Columns.Add("IDCategoriaProfesionalSCCP", System.Type.GetType("System.Double"))

        For i As Integer = 1 To 31 ' Suponiendo que tu tabla tiene columnas para cada día del mes
            dtHorasPersonasDias.Columns.Add(i.ToString(), System.Type.GetType("System.Double"))
        Next

        ' Iterar sobre las filas de dtHoras y calcular la suma por día y operario
        For Each filaHoras As DataRow In dtHoras.Rows
            Dim fechaTrabajo As DateTime = DateTime.Parse(filaHoras("FechaInicio").ToString())

            ' Verificar si la fecha está dentro del rango especificado
            If fechaTrabajo >= Fecha1 AndAlso fechaTrabajo <= Fecha2 Then
                Dim idOperario As String = filaHoras("IDOperario").ToString()
                Dim empresa As String = filaHoras("Empresa").ToString()
                Dim totalHoras As Double = Convert.ToDouble(Nz(filaHoras("HorasRealMod"), 0)) + Convert.ToDouble(Nz(filaHoras("HorasAdministrativas"), 0)) + Convert.ToDouble(Nz(filaHoras("HorasBaja"), 0))

                ' Buscar la fila correspondiente en dtHorasPersonasDias y actualizar el valor
                Dim fila As DataRow = dtHorasPersonasDias.Rows.Cast(Of DataRow)().FirstOrDefault(Function(row) row("IDOperario").ToString() = idOperario)
                If fila IsNot Nothing Then
                    Dim diaDelMes As Integer = fechaTrabajo.Day
                    'fila(diaDelMes.ToString()) = totalHoras
                    fila(diaDelMes.ToString()) = Convert.ToDouble(Nz(fila(diaDelMes.ToString()), 0)) + totalHoras
                Else
                    ' Si la fila no existe, puedes agregarla
                    fila = dtHorasPersonasDias.NewRow()
                    fila("Empresa") = empresa

                    Dim bbdd As String
                    bbdd = DevuelveBaseDeDatos(empresa)
                    fila("IDGET") = DevuelveIDGET(bbdd, idOperario)
                    fila("IDOperario") = idOperario
                    fila("DescOperario") = DevuelveDescOperario(bbdd, idOperario)
                    fila("IDCategoriaProfesionalSCCP") = DevuelveIDCategoriaProfesionalSCCPTodasBasesDeDatos(idOperario)
                    Dim diaDelMes As Integer = fechaTrabajo.Day
                    fila(diaDelMes.ToString()) = totalHoras
                    dtHorasPersonasDias.Rows.Add(fila)
                End If
            End If
        Next

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial

        Dim ruta As New FileInfo("N:\10. AUXILIARES\00. EXPERTIS\05. CHECK HORAS-A3\" & Month(Fecha1) & " MATRIZ HORAS " & Year(Fecha1) & ".xlsx")
        'Dim ruta As New FileInfo("N:\01. A3\" & mes & " A3 " & mes & anio.Substring(anio.Length - 2) & ".xlsx")
        Dim rutaCadena As String = ""
        rutaCadena = ruta.FullName

        'Verificar si el archivo existe.
        If File.Exists(rutaCadena) Then
            'Si el archivo existe, eliminarlo.
            File.Delete(rutaCadena)
        End If

        Using package As New ExcelPackage(ruta)
            ' HOJA 1
            Dim worksheet = package.Workbook.Worksheets.Add("MATRIZ HORAS")
            worksheet.Cells("A1").LoadFromDataTable(dtHorasPersonasDias, True)
            Dim fila1 As ExcelRange = worksheet.Cells(1, 1, 1, worksheet.Dimension.End.Column)
            fila1.Style.Font.Bold = True
            worksheet.Cells("A1:" & GetExcelColumnName(worksheet.Dimension.End.Column) & "1").AutoFilter = True

            ' Guardar el archivo de Excel.
            package.Save()

            MsgBox("Fichero guardado en N:\10. AUXILIARES\00. EXPERTIS\05. CHECK HORAS-A3\")
        End Using
    End Sub
    Public Sub FormaTablaMatriz(ByRef dtHorasPersonasDias As DataTable)
        'dtHorasPersonasDias.Columns.Add("Empresa")
        'dtHorasPersonasDias.Columns.Add("IDGET")
        dtHorasPersonasDias.Columns.Add("IDOperario")
        'dtHorasPersonasDias.Columns.Add("DescOperario")
        dtHorasPersonasDias.Columns.Add("1")
        dtHorasPersonasDias.Columns.Add("2")
        dtHorasPersonasDias.Columns.Add("3")
        dtHorasPersonasDias.Columns.Add("4")
        dtHorasPersonasDias.Columns.Add("5")
        dtHorasPersonasDias.Columns.Add("7")
        dtHorasPersonasDias.Columns.Add("8")
        dtHorasPersonasDias.Columns.Add("9")
        dtHorasPersonasDias.Columns.Add("10")
        dtHorasPersonasDias.Columns.Add("11")
        dtHorasPersonasDias.Columns.Add("12")
        dtHorasPersonasDias.Columns.Add("13")
        dtHorasPersonasDias.Columns.Add("14")
        dtHorasPersonasDias.Columns.Add("15")
        dtHorasPersonasDias.Columns.Add("16")
        dtHorasPersonasDias.Columns.Add("17")
        dtHorasPersonasDias.Columns.Add("18")
        dtHorasPersonasDias.Columns.Add("19")
        dtHorasPersonasDias.Columns.Add("20")
        dtHorasPersonasDias.Columns.Add("21")
        dtHorasPersonasDias.Columns.Add("22")
        dtHorasPersonasDias.Columns.Add("23")
        dtHorasPersonasDias.Columns.Add("24")
        dtHorasPersonasDias.Columns.Add("25")
        dtHorasPersonasDias.Columns.Add("26")
        dtHorasPersonasDias.Columns.Add("27")
        dtHorasPersonasDias.Columns.Add("28")
        dtHorasPersonasDias.Columns.Add("29")
        dtHorasPersonasDias.Columns.Add("30")
        dtHorasPersonasDias.Columns.Add("31")
    End Sub

    Private Sub CargaHorasJPSTAFF_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        FormaTablaResumen()
    End Sub

    Private Sub bDuplicados_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim dt As New DataTable
        Dim f As New Filter

        dt = New BE.DataEngine().Filter("vUniontbObraMod", f)
        If dt.Rows.Count > 0 Then
            MsgBox("El operario " & dt.Rows(0)("IDGET").ToString & " tiene horas en mas de una empresa.")
        Else
            MsgBox("No hay registros duplicados con misma fecha en distintas empresas.")
        End If
    End Sub


    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim dt As New DataTable
        Dim f As New Filter

        Dim frm As New frmInformeFecha
        frm.ShowDialog()
        Dim Fecha1 As String : Dim Fecha2 As String
        Fecha1 = frm.fecha1 : Fecha2 = frm.fecha2

        f.Add("FechaInicio", FilterOperator.GreaterThanOrEqual, Fecha1)
        f.Add("FechaInicio", FilterOperator.LessThanOrEqual, Fecha2)

        dt = New BE.DataEngine().Filter("vUniontbObraMod", f)
        Dim cont As Integer = 0
        If dt.Rows.Count > 0 Then
            For Each dr As DataRow In dt.Rows
                If dr("IDGET") = "GET03540" Or dr("IDGET") = "GET03605" Then
                    cont = 1
                Else
                    MsgBox("El operario " & dt.Rows(0)("IDGET").ToString & " tiene horas en mas de una empresa.")
                End If
            Next

        Else
            MsgBox("No hay registros duplicados con misma fecha en distintas empresas.", MsgBoxStyle.Information, "Check duplicidad horas")
            Exit Sub
        End If
        If cont = 1 Then
            MsgBox("No hay registros duplicados con misma fecha en distintas empresas.")
        End If

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '6. TABLA DE HORAS POR PERSONAS
        'Dim dtHorasPersonasDias As New DataTable
        'FormaTablaMatriz(dtHorasPersonasDias)

        Dim frm As New frmInformeFecha
        frm.ShowDialog()
        Dim Fecha1 As String : Dim Fecha2 As String
        Fecha1 = frm.fecha1 : Fecha2 = frm.fecha2
        Dim dtHoras As New DataTable
        Dim f As New Filter
        f.Add("FechaInicio", FilterOperator.GreaterThanOrEqual, Fecha1)
        f.Add("FechaInicio", FilterOperator.LessThanOrEqual, Fecha2)

        dtHoras = New BE.DataEngine().Filter("vUniontbObraModControl", f, , "Empresa asc")

        ' Crear la estructura de la tabla dtHorasPersonasDias
        Dim dtHorasPersonasDias As New DataTable()
        dtHorasPersonasDias.Columns.Add("Empresa")
        dtHorasPersonasDias.Columns.Add("IDGET")
        dtHorasPersonasDias.Columns.Add("IDOperario")
        dtHorasPersonasDias.Columns.Add("DescOperario")
        dtHorasPersonasDias.Columns.Add("IDCategoriaProfesionalSCCP", System.Type.GetType("System.Double"))

        For i As Integer = 1 To 31 ' Suponiendo que tu tabla tiene columnas para cada día del mes
            dtHorasPersonasDias.Columns.Add(i.ToString(), System.Type.GetType("System.Double"))
        Next

        ' Iterar sobre las filas de dtHoras y calcular la suma por día y operario
        For Each filaHoras As DataRow In dtHoras.Rows
            Dim fechaTrabajo As DateTime = DateTime.Parse(filaHoras("FechaInicio").ToString())

            ' Verificar si la fecha está dentro del rango especificado
            If fechaTrabajo >= Fecha1 AndAlso fechaTrabajo <= Fecha2 Then
                Dim idOperario As String = filaHoras("IDOperario").ToString()
                Dim empresa As String = filaHoras("Empresa").ToString()
                Dim totalHoras As Double = Convert.ToDouble(Nz(filaHoras("HorasRealMod"), 0)) + Convert.ToDouble(Nz(filaHoras("HorasAdministrativas"), 0)) + Convert.ToDouble(Nz(filaHoras("HorasBaja"), 0))

                ' Buscar la fila correspondiente en dtHorasPersonasDias y actualizar el valor
                Dim fila As DataRow = dtHorasPersonasDias.Rows.Cast(Of DataRow)().FirstOrDefault(Function(row) row("IDOperario").ToString() = idOperario)
                If fila IsNot Nothing Then
                    Dim diaDelMes As Integer = fechaTrabajo.Day
                    'fila(diaDelMes.ToString()) = totalHoras
                    fila(diaDelMes.ToString()) = Convert.ToDouble(Nz(fila(diaDelMes.ToString()), 0)) + totalHoras
                Else
                    ' Si la fila no existe, puedes agregarla
                    fila = dtHorasPersonasDias.NewRow()
                    fila("Empresa") = empresa

                    Dim bbdd As String
                    bbdd = DevuelveBaseDeDatos(empresa)
                    fila("IDGET") = DevuelveIDGET(bbdd, idOperario)
                    fila("IDOperario") = idOperario
                    fila("DescOperario") = DevuelveDescOperario(bbdd, idOperario)
                    fila("IDCategoriaProfesionalSCCP") = DevuelveIDCategoriaProfesionalSCCPTodasBasesDeDatos(idOperario)
                    Dim diaDelMes As Integer = fechaTrabajo.Day
                    fila(diaDelMes.ToString()) = totalHoras
                    dtHorasPersonasDias.Rows.Add(fila)
                End If
            End If
        Next

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial

        Dim ruta As New FileInfo("N:\10. AUXILIARES\00. EXPERTIS\05. CHECK HORAS-A3\" & Month(Fecha1) & " MATRIZ HORAS " & Year(Fecha1) & ".xlsx")
        'Dim ruta As New FileInfo("N:\01. A3\" & mes & " A3 " & mes & anio.Substring(anio.Length - 2) & ".xlsx")
        Dim rutaCadena As String = ""
        rutaCadena = ruta.FullName

        'Verificar si el archivo existe.
        If File.Exists(rutaCadena) Then
            'Si el archivo existe, eliminarlo.
            File.Delete(rutaCadena)
        End If

        Using package As New ExcelPackage(ruta)
            ' HOJA 1
            Dim worksheet = package.Workbook.Worksheets.Add("MATRIZ HORAS")
            worksheet.Cells("A1").LoadFromDataTable(dtHorasPersonasDias, True)
            Dim fila1 As ExcelRange = worksheet.Cells(1, 1, 1, worksheet.Dimension.End.Column)
            fila1.Style.Font.Bold = True
            worksheet.Cells("A1:" & GetExcelColumnName(worksheet.Dimension.End.Column) & "1").AutoFilter = True

            ' Congelar la primera columna
            worksheet.View.FreezePanes(2, 1)

            ' Guardar el archivo de Excel.
            package.Save()

            MsgBox("Fichero guardado en N:\10. AUXILIARES\00. EXPERTIS\05. CHECK HORAS-A3\")
        End Using
    End Sub


    Private Sub bDobleCotizacion_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim frmCrea As New frmCreaHorasDobleCotizacion
        frmCrea.ShowDialog()

    End Sub

    Private Sub bUKNuevo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bUKNuevo.Click
        cadenaFinal = ""
        'Esta variable determinará el nº de fichero que se ha insertado
        Dim fichero As Integer = 0
        Dim dtUkPersonas As New DataTable
        darFormaTablaUK(dtUkPersonas)

        Do
            fichero = fichero + 1 : cadenaFinal = ""
            SeleccionarPDFyLeerDataTableUK(fichero)
            GuardaFicheroUkTxt()
            LeeFicheroYGuardaEnExcelNuevo(dtUkPersonas, fichero)
            Dim respuesta As DialogResult = MessageBox.Show("¿Deseas cargar algún Excel más?", "Continuar", MessageBoxButtons.YesNo)
            ' Salir del bucle si el usuario responde "No"
            If respuesta = DialogResult.No Then
                Exit Do
            End If
        Loop
        FormaExcelUkPorFicheros(dtUkPersonas)
    End Sub

    Public Sub FormaExcelUkPorFicheros(ByVal dtUkPersonas As DataTable)
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial

        Dim saveFileDialog1 As New SaveFileDialog()

        For Each fila As DataRow In dtUkPersonas.Rows
            For Each columna As DataColumn In dtUkPersonas.Columns
                ' Si el valor es de tipo Double, formatearlo con coma en lugar de punto
                Try
                    fila(columna) = (DirectCast(fila(columna), String).Replace(".", ","))
                Catch ex As Exception
                    fila(columna) = ""
                End Try

                Try
                    fila(columna) = (DirectCast(fila(columna), String).Replace("(", "-"))
                Catch ex As Exception
                    fila(columna) = ""
                End Try

                Try
                    fila(columna) = (DirectCast(fila(columna), String).Replace(")", ""))
                Catch ex As Exception
                    fila(columna) = ""
                End Try


                ' Si la columna es "Diccionario", eliminar letras y dejar solo dígitos
                If columna.ColumnName = "Diccionario" Then
                    Dim valorOriginal As String = DirectCast(fila(columna), String)
                    Dim valorSinLetras As String = New String(valorOriginal.Where(Function(c) Char.IsDigit(c)).ToArray())
                    fila(columna) = valorSinLetras
                End If
            Next
        Next
        '-----AQUI SEPARO LA TABLA ENTRE EL NUMERO DE FICHEROS INSERTADOS-----

        saveFileDialog1.Filter = "Archivos de texto|*.xlsx|Todos los archivos|*.*"
        saveFileDialog1.Title = "Guardar archivo"

        If saveFileDialog1.ShowDialog() = DialogResult.OK Then
            ' Obtener la ruta seleccionada por el usuario
            Dim rutaArchivo As String = saveFileDialog1.FileName

            ' Verificar si el archivo existe.
            If File.Exists(rutaArchivo) Then
                ' Si el archivo existe, eliminarlo.
                File.Delete(rutaArchivo)
            End If

            Dim dtUkPersonasOrdenado = dtUkPersonas.DefaultView.ToTable()

            Using package As New ExcelPackage(rutaArchivo)
                ' Crear una hoja de cálculo y obtener una referencia a ella.
                Dim worksheetFinal = package.Workbook.Worksheets.Add("1")
                'worksheetFinal.Cells("A1:" & GetExcelColumnName(worksheetFinal.Dimension.End.Column) & "1").AutoFilter = True
                ' Copiar los datos de la DataTable completa a la hoja de cálculo "FINAL".
                worksheetFinal.Cells("A1").LoadFromDataTable(dtUkPersonasOrdenado, True)
                worksheetFinal.Cells("A1:" & GetExcelColumnName(worksheetFinal.Dimension.End.Column) & "1").AutoFilter = True

                ' Aplicar formato negrita a la fila 1
                Dim fila1Final As ExcelRange = worksheetFinal.Cells(1, 1, 1, worksheetFinal.Dimension.End.Column)
                fila1Final.Style.Font.Bold = True

                ' Iterar sobre las celdas de la hoja "FINAL" y aplicar formato
                For row As Integer = 2 To worksheetFinal.Dimension.End.Row
                    For col As Integer = 2 To worksheetFinal.Dimension.End.Column
                        Dim valorCelda As String = worksheetFinal.Cells(row, col).Text
                        Dim valorNumerico As Double

                        ' Intentar convertir el valor a un número
                        If Double.TryParse(valorCelda, valorNumerico) Then
                            ' Si el valor está entre paréntesis, multiplicarlo por -1
                            ' Asignar el valor numérico a la celda
                            worksheetFinal.Cells(row, col).Value = valorNumerico
                        End If
                    Next
                Next

                ' Aplicar formato de moneda a la hoja "FINAL"
                Dim rangoMonedaFinal As ExcelRange = worksheetFinal.Cells("B2:" & GetExcelColumnName(worksheetFinal.Dimension.End.Column) & worksheetFinal.Dimension.End.Row)
                rangoMonedaFinal.Style.Numberformat.Format = "#,##0.00£"

                ' Congelar la primera fila y la primera columna en la hoja "FINAL"
                worksheetFinal.View.FreezePanes(2, 1)

                ' Crear DataTable de resumen
                Dim resumenDataTable As New DataTable("RESUMEN")
                resumenDataTable.Columns.Add("Fichero", GetType(Integer))
                resumenDataTable.Columns.Add("Tax", GetType(Double))
                resumenDataTable.Columns.Add("Net NI", GetType(Double))
                resumenDataTable.Columns.Add("Net Pay", GetType(Double))
                resumenDataTable.Columns.Add("Net Er NI", GetType(Double))

                ' Obtener valores únicos de la columna "Fichero"
                Dim ficherosUnicos = dtUkPersonasOrdenado.AsEnumerable().Select(Function(row) Convert.ToInt32(row("Fichero"))).Distinct()

                ' Iterar sobre los valores únicos de "Fichero" para crear el resumen
                For Each fichero As Integer In ficherosUnicos
                    ' Filtrar el DataTable por el valor actual de "Fichero"
                    Dim dtFiltrado = dtUkPersonasOrdenado.AsEnumerable().Where(Function(row) Convert.ToInt32(row("Fichero")) = fichero).CopyToDataTable()

                    ' Sumar las cantidades de las columnas 9, 10, 15 y 16
                    Dim sumaColumna9 As Double = dtFiltrado.AsEnumerable().Sum(Function(row) GetDoubleValue(row, "Tax"))
                    Dim sumaColumna10 As Double = dtFiltrado.AsEnumerable().Sum(Function(row) GetDoubleValue(row, "Net NI"))
                    Dim sumaColumna15 As Double = dtFiltrado.AsEnumerable().Sum(Function(row) GetDoubleValue(row, "Net Pay"))
                    Dim sumaColumna16 As Double = dtFiltrado.AsEnumerable().Sum(Function(row) GetDoubleValue(row, "Net Er NI"))

                    ' Agregar fila al DataTable de resumen
                    resumenDataTable.Rows.Add(fichero, sumaColumna9, sumaColumna10, sumaColumna15, sumaColumna16)

                    ' Crear una hoja de cálculo para el valor actual de "Fichero"
                    Dim worksheetFichero = package.Workbook.Worksheets.Add("F" & fichero.ToString())

                    ' Copiar los datos filtrados a la hoja de cálculo correspondiente.
                    worksheetFichero.Cells("A1").LoadFromDataTable(dtFiltrado, True)
                    worksheetFichero.Cells("A1:" & GetExcelColumnName(worksheetFichero.Dimension.End.Column) & "1").AutoFilter = True
                    ' Aplicar formato negrita a la fila 1
                    Dim fila1Fichero As ExcelRange = worksheetFichero.Cells(1, 1, 1, worksheetFichero.Dimension.End.Column)
                    fila1Fichero.Style.Font.Bold = True

                    ' Congelar la primera fila
                    worksheetFichero.View.FreezePanes(2, 1)

                    For row As Integer = 2 To worksheetFichero.Dimension.End.Row
                        For col As Integer = 2 To worksheetFichero.Dimension.End.Column
                            Dim valorCelda As String = worksheetFichero.Cells(row, col).Text
                            Dim valorNumerico As Double

                            ' Intentar convertir el valor a un número
                            If Double.TryParse(valorCelda, valorNumerico) Then
                                ' Si el valor está entre paréntesis, multiplicarlo por -1
                                ' Asignar el valor numérico a la celda
                                worksheetFichero.Cells(row, col).Value = valorNumerico
                            End If
                        Next
                    Next
                Next

                ' Crear hoja "RESUMEN" y cargar los datos del DataTable de resumen
                Dim worksheetResumen = package.Workbook.Worksheets.Add("RESUMEN")

                ' Copiar los datos de la DataTable de resumen a la hoja de cálculo "RESUMEN".
                worksheetResumen.Cells("A1").LoadFromDataTable(resumenDataTable, True)

                ' Aplicar formato negrita a la fila 1
                Dim fila1Resumen As ExcelRange = worksheetResumen.Cells(1, 1, 1, worksheetResumen.Dimension.End.Column)
                fila1Resumen.Style.Font.Bold = True

                ' Iterar sobre las celdas de la hoja "RESUMEN" y aplicar formato
                For row As Integer = 2 To worksheetResumen.Dimension.End.Row
                    For col As Integer = 2 To 5 ' Cambia estos valores según las columnas específicas que desees formatear
                        Dim valorCelda As String = worksheetResumen.Cells(row, col).Text
                        Dim valorNumerico As Double

                        ' Intentar convertir el valor a un número
                        If Double.TryParse(valorCelda, valorNumerico) Then
                            ' Si el valor está entre paréntesis, multiplicarlo por -1
                            ' Asignar el valor numérico a la celda
                            worksheetResumen.Cells(row, col).Value = valorNumerico
                        End If
                    Next
                Next

                ' Aplicar formato de moneda a la hoja "RESUMEN"
                Dim rangoMonedaResumen As ExcelRange = worksheetResumen.Cells("B2:E" & worksheetResumen.Dimension.End.Row)
                rangoMonedaResumen.Style.Numberformat.Format = "#,##0.00£"


                worksheetResumen.Column(1).Width = 14 : worksheetResumen.Column(2).Width = 14 : worksheetResumen.Column(3).Width = 14 : worksheetResumen.Column(4).Width = 14 : worksheetResumen.Column(5).Width = 14

                ' Guardar el archivo de Excel.
                package.Save()
            End Using
        End If



        MsgBox("Fichero creado correctamente en N:\10. AUXILIARES\00. EXPERTIS\02. A3\")
    End Sub

    Function GetDoubleValue(ByVal row As DataRow, ByVal columnName As String) As Double
        Dim value As Object = row(columnName)
        Dim result As Double

        If value IsNot DBNull.Value AndAlso Double.TryParse(value.ToString(), result) Then
            Return result
        Else
            Return 0 ' Puedes cambiar esto a cualquier valor predeterminado que desees cuando la celda sea vacía o no numérica
        End If
    End Function

End Class
