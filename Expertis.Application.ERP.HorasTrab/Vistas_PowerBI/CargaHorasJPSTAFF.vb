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
        CD.Filter = "Archivos Excel(*.xls;*.xlsx)|*.xls;*xlsx|Todos los archivos(*.*)|*.*"

        'CD.ShowOpen()
        CD.ShowDialog()

        If CD.FileName <> "" Then
            'lblRuta.Caption = CD.FileName
            lblRuta.Text = CD.FileName
        End If
    End Sub
    '12/05/2022
    Public Function CargaTablas(ByRef dtTecozam As DataTable, ByRef dtPortugal As DataTable, ByRef dtUK As DataTable, ByVal dt As DataTable) As Integer
        dtTecozam.Columns.Add("IDOperario", GetType(String))
        dtTecozam.Columns.Add("DescOperario", GetType(String))
        dtTecozam.Columns.Add("DNI", GetType(String))
        dtTecozam.Columns.Add("Empresa", GetType(String))
        dtTecozam.Columns.Add("CentroCoste", GetType(String))
        dtTecozam.Columns.Add("ProduccionSinVentas", GetType(String))
        dtTecozam.Columns.Add("Porcentaje", GetType(Double))

        dtPortugal.Columns.Add("IDOperario", GetType(String))
        dtPortugal.Columns.Add("DescOperario", GetType(String))
        dtPortugal.Columns.Add("DNI", GetType(String))
        dtPortugal.Columns.Add("Empresa", GetType(String))
        dtPortugal.Columns.Add("CentroCoste", GetType(String))
        dtPortugal.Columns.Add("ProduccionSinVentas", GetType(String))
        dtPortugal.Columns.Add("Porcentaje", GetType(Double))

        dtUK.Columns.Add("IDOperario", GetType(String))
        dtUK.Columns.Add("DescOperario", GetType(String))
        dtUK.Columns.Add("DNI", GetType(String))
        dtUK.Columns.Add("Empresa", GetType(String))
        dtUK.Columns.Add("CentroCoste", GetType(String))
        dtUK.Columns.Add("ProduccionSinVentas", GetType(String))
        dtUK.Columns.Add("Porcentaje", GetType(Double))

        For Each dr As DataRow In dt.Rows
            If dr("Empresa") = "T. ES. " Then
                dtTecozam.ImportRow(dr)
            ElseIf dr("Empresa") = "D. P. " Then
                dtPortugal.ImportRow(dr)
            ElseIf dr("Empresa") = "T. UK. " Then
                dtUK.ImportRow(dr)
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
    Private Sub btnAceptar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAceptar.Click
        Dim hoja As String
        'hoja = "% HORAS J.P. y STAFF"
        hoja = "HORAS JP y STAFF"
        Dim dt As New DataTable

        Dim ruta As String
        ruta = lblRuta.Text
        Dim rango As String = "A2:G10000"
        dt = ObtenerDatosExcel(ruta, hoja, rango)

        dt.Columns("F1").ColumnName = "IDOperario"
        dt.Columns("F2").ColumnName = "DescOperario"
        dt.Columns("F3").ColumnName = "DNI"
        dt.Columns("F4").ColumnName = "Empresa"
        dt.Columns("F5").ColumnName = "CentroCoste"
        dt.Columns("F6").ColumnName = "ProduccionSinVentas"
        dt.Columns("F7").ColumnName = "Porcentaje"

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

        Dim flat As Integer
        'FILTRO LOS REGISTROS DE TECOZAM 'FILTRO LOS REGISTROS DE DCZ 'FILTROS LOS REGISTROS DE UK
        flat = CargaTablas(dtTecozam, dtPortugal, dtUK, dt)

        If flat = 0 Then
            MsgBox("Existen registros que no coinciden con ninguna empresa.")
            Exit Sub
        End If

        Dim result As DialogResult = MessageBox.Show("Hay " & dtTecozam.Rows.Count & " registros de T. ES. " & vbCrLf & _
        "Hay " & dtPortugal.Rows.Count & " registros de D. P." & vbCrLf & _
        "Hay " & dtUK.Rows.Count & " registros de T. UK." & vbCrLf, "¿Están correctos estos datos?", MessageBoxButtons.YesNo)
        If result = DialogResult.No Then
            Exit Sub
        End If

        '-------SOBREESCRIBIR HORAS POR MES Y AÑO NATURAL---------
        If SobreescribirHoras(Fecha1, Fecha2) = True Then
            Dim result2 As DialogResult = MessageBox.Show("Exiten horas de JP y STAFF entre este rango de fechas, ¿desea sobreescribir los datos?", "Borrar e insertar datos.", MessageBoxButtons.YesNo)
            If result2 = DialogResult.Yes Then
                BorrarDatos(mes, año)
            Else
                Exit Sub
            End If
        Else
        End If
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
        '---------------FIN CHECKS---------------------------

        'Inserta horas en Tecozam
        insertaHorasJPStaffTecozam(mes, año, Fecha1, Fecha2, dtTecozam)
        'Inserta horas en Portugal
        insertaHorasJPStaffPortugal(mes, año, Fecha1, Fecha2, dtPortugal)
        'Inserta horas en UK
        insertaHorasJPStaffUK(mes, año, Fecha1, Fecha2, dtPortugal)
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
                If CategoriaSSCP.ToString.Length = 0 Or (CategoriaSSCP.ToString <> 1 And CategoriaSSCP.ToString <> 5) Then
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
            Return dt(0)("Abreviatura")
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
        dtOperario = New BE.DataEngine().Filter(basededatos & "..tbMaestroOperarioSat", f)

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
        dtOperario = New BE.DataEngine().Filter(basededatos & "..tbMaestroOperario", f)

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
        dtCalendario.Merge(dtVacaciones)
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
    End Sub
    Public Function getListadoPersonasOfiFerrallas(ByVal Fecha1 As String, ByVal Fecha2 As String) As DataTable
        Dim dt As New DataTable
        Dim sql As String
        '------COMO YO LO DE DEJARIA---------
        sql = "select IDOperario, Obra_Predeterminada from " & DB_FERRALLAS & "..tbMaestroOperarioSat " & _
        "where (Obra_Predeterminada='12677838' Or Obra_Predeterminada='12677615' Or Obra_Predeterminada='12678141') and " & _
        "(Fecha_Baja is null or (Fecha_Baja>='" & Fecha1 & "' and Fecha_Baja<='" & Fecha2 & "'))" & _
        " Union " & _
        "select IDOperario, Proyecto AS Obra_Predeterminada " & _
        "from " & DB_FERRALLAS & "..tbHistoricoPersonal " & _
        "where (Proyecto = '12677838' OR Proyecto = '12677615' OR Proyecto = '12678141') and " & _
        "((Fecha >= '" & Fecha1 & "' AND Fecha <= '" & Fecha2 & "'))"

        dt = aux.EjecutarSqlSelect(sql)
        Return dt
    End Function

    Public Function getListadoPersonasOfiDCZ(ByVal Fecha1 As String, ByVal Fecha2 As String) As DataTable
        Dim dt As New DataTable
        '----------FORMA BUENA'-------------
        Dim sql As String
        sql = "select IDOperario, Obra_Predeterminada from " & DB_DCZ & "..tbMaestroOperarioSat " & _
        "where Obra_Predeterminada='11860026' and " & _
        "(Fecha_Baja is null or (Fecha_Baja>='" & Fecha1 & "' and Fecha_Baja<='" & Fecha2 & "'))" & _
        " Union " & _
        "select IDOperario, Proyecto AS Obra_Predeterminada " & _
        "from " & DB_DCZ & "..tbHistoricoPersonal " & _
        "where (Proyecto = '11860026') and " & _
        "((Fecha >= '" & Fecha1 & "' AND Fecha <= '" & Fecha2 & "'))"

        dt = aux.EjecutarSqlSelect(sql)
        Return dt
    End Function

    Public Function getListadoPersonasOfiTecozam(ByVal Fecha1 As String, ByVal Fecha2 As String) As DataTable
        Dim dt As New DataTable
        Dim sql As String
        '------COMO YO LO DE DEJARIA---------
        sql = "select IDOperario, Obra_Predeterminada from " & DB_TECOZAM & "..tbMaestroOperarioSat " & _
        "where (Obra_Predeterminada='16895681' Or Obra_Predeterminada='11984995') and " & _
        "(Fecha_Baja is null or (Fecha_Baja>='" & Fecha1 & "' and Fecha_Baja<='" & Fecha2 & "'))" & _
        " Union " & _
        "select IDOperario, Proyecto AS Obra_Predeterminada " & _
        "from " & DB_TECOZAM & "..tbHistoricoPersonal " & _
        "where (Proyecto = '16895681' OR Proyecto = '11984995') and " & _
        "((Fecha >= '" & Fecha1 & "' AND Fecha <= '" & Fecha2 & "'))"


        'sql = "select IDOperario, Obra_Predeterminada from DB_TECOZAM..tbMaestroOperarioSat where idoperario='T3450'"
        dt = aux.EjecutarSqlSelect(sql)
        Return dt
    End Function

    Public Function getListadoPersonasOfiSecozam(ByVal Fecha1 As String, ByVal Fecha2 As String) As DataTable
        Dim dt As New DataTable
        Dim sql As String
        '------COMO YO LO DE DEJARIA---------
        sql = "select IDOperario, Proyecto from " & DB_SECOZAM & "..tbMaestroOperarioSat " & _
        "where (Obra_Predeterminada='11854299' Or Obra_Predeterminada='11854231') and " & _
        "(Fecha_Baja is null or (Fecha_Baja>='" & Fecha1 & "' and Fecha_Baja<='" & Fecha2 & "'))" & _
        " Union " & _
        "select IDOperario, Proyecto AS Obra_Predeterminada " & _
        "from " & DB_SECOZAM & "..tbHistoricoPersonal " & _
        "where (Proyecto = '11854299' OR Proyecto = '1198118542314995') and " & _
        "((Fecha >= '" & Fecha1 & "' AND Fecha <= '" & Fecha2 & "'))"

        dt = aux.EjecutarSqlSelect(sql)
        Return dt
    End Function

    Public Function DevuelveUltimoCambioObra(ByVal IDOperario As String, ByVal bbdd As String) As String
        Dim f As New Filter
        Dim dt As New DataTable
        f.Add("IDOperario", FilterOperator.Equal, IDOperario)

        dt = New BE.DataEngine().Filter(bbdd & "..tbHistoricoPersonal", f, , "Fecha desc")

        Return dt.Rows(0)("Proyecto")
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

                    txtSQL = "Insert into " & DB_TECOZAM & "..tbObraMODControl(IdLineaModControl, IdTrabajo, IdObra, CodTrabajo, DescTrabajo, IdTipoTrabajo, " & _
                            "IdSubTipoTrabajo, IdOperario, IdCategoria, IdHora, FechaInicio, HorasRealMod, TasaRealModA, " & _
                             "ImpRealModA, HorasFactMod, ImpFactModA, DescParte, Facturable, FechaCreacionAudi, FechaModificacionAudi, Usuarioaudi, IDOficio, IdTipoTurno, HorasAdministrativas, IDCategoriaProfesionalSCCP) " & _
                             "Values(" & IDAutonumerico & ", " & IDTrabajo & ", " & IDObra & ", '" & _
                             CodTrabajo & "', '" & DescTrabajo & "', '" & IdTipoTrabajo & "', '" & _
                             IdSubTipoTrabajo & "', '" & IDOperario & "', 'PREDET', '" & _
                             "HA" & "', '" & fecha & "', 0 , " & 0 & ", " & 0 & _
                             ", 0 , " & 0 & _
                             ", '" & DescParte & "', " & 0 & ", '" & Date.Now.Date & "', '" & Date.Now.Date & "', '" & ExpertisApp.UserName & "','" & IDOficio & "', 4, 8 ," & Nz(IDCategoriaProfesionalSCCP, "") & ")"

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
            If IDObra <> "12677838" Or IDObra <> "12677615" Or IDObra <> "12678141" Then
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

                    txtSQL = "Insert into " & DB_FERRALLAS & "..tbObraMODControl(IdLineaModControl, IdTrabajo, IdObra, CodTrabajo, DescTrabajo, IdTipoTrabajo, " & _
                            "IdSubTipoTrabajo, IdOperario, IdCategoria, IdHora, FechaInicio, HorasRealMod, TasaRealModA, " & _
                             "ImpRealModA, HorasFactMod, ImpFactModA, DescParte, Facturable, FechaCreacionAudi, FechaModificacionAudi, Usuarioaudi,IDOficio, IdTipoTurno, HorasAdministrativas, IDCategoriaProfesionalSCCP) " & _
                             "Values(" & IDAutonumerico & ", " & IDTrabajo & ", " & IDObra & ", '" & _
                             CodTrabajo & "', '" & DescTrabajo & "', '" & IdTipoTrabajo & "', '" & _
                             IdSubTipoTrabajo & "', '" & IDOperario & "', 'PREDET', '" & _
                             "HA" & "', '" & fecha & "', 0 , " & 0 & ", " & 0 & _
                             ", 0 , " & 0 & _
                             ", '" & DescParte & "', " & 0 & ", '" & Date.Now.Date & "', '" & Date.Now.Date & "', '" & ExpertisApp.UserName & "','" & IDOficio & "', 4, 8 ," & Nz(IDCategoriaProfesionalSCCP, "") & ")"

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
            If IDObra <> "11854299" Or IDObra <> "11854231" Then
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

                    txtSQL = "Insert into " & DB_SECOZAM & "..tbObraMODControl(IdLineaModControl, IdTrabajo, IdObra, CodTrabajo, DescTrabajo, IdTipoTrabajo, " & _
                            "IdSubTipoTrabajo, IdOperario, IdCategoria, IdHora, FechaInicio, HorasRealMod, TasaRealModA, " & _
                             "ImpRealModA, HorasFactMod, ImpFactModA, DescParte, Facturable, FechaCreacionAudi, FechaModificacionAudi, Usuarioaudi,IDOficio, IdTipoTurno, HorasAdministrativas, IDCategoriaProfesionalSCCP) " & _
                             "Values(" & IDAutonumerico & ", " & IDTrabajo & ", " & IDObra & ", '" & _
                             CodTrabajo & "', '" & DescTrabajo & "', '" & IdTipoTrabajo & "', '" & _
                             IdSubTipoTrabajo & "', '" & IDOperario & "', 'PREDET', '" & _
                             "HA" & "', '" & fecha & "', 0 , " & 0 & ", " & 0 & _
                             ", 0 , " & 0 & _
                             ", '" & DescParte & "', " & 0 & ", '" & Date.Now.Date & "', '" & Date.Now.Date & "', '" & ExpertisApp.UserName & "','" & IDOficio & "', 4, 8 ," & Nz(IDCategoriaProfesionalSCCP, "") & ")"

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

    Private Sub bNota_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bNota.Click
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
        importarExcelPorEmpresa()
    End Sub

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

        sql2 = "Select * from " & DB_UK & "..tbObraCabecera where NObra='" & NObra & "'"
        dtObra = aux.EjecutarSqlSelect(sql2)

        Return dtObra.Rows(0)("IDObra")
    End Function

    Dim dtResumen As New DataTable

    Private Sub bA3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bA3.Click

        'CREO LA TABLA A MODO RESUMEN QUE VA EN EL A3 UNIFICADO(HOJA 2 DEL EXCEL)
        FormaTablaResumen()
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
            Case "T. ES. ", "FERR. ", "SEC. "
                rango = "B10:Z10000"
            Case "D. P. "
                rango = "A2:F500"
            Case "T. UK. "
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
        If empresa = "T. ES. " Then
            bbdd = DB_TECOZAM
            newDataTable = FormaTablaEspaña(dt, newDataTable, bbdd, mes, anio, empresa)
        ElseIf empresa = "FERR. " Then
            bbdd = DB_FERRALLAS
            newDataTable = FormaTablaEspaña(dt, newDataTable, bbdd, mes, anio, empresa)
        ElseIf empresa = "SEC. " Then
            bbdd = DB_SECOZAM
            newDataTable = FormaTablaEspaña(dt, newDataTable, bbdd, mes, anio, empresa)
        ElseIf empresa = "D. P. " Then
            bbdd = DB_DCZ
            newDataTable = FormaTablaDCZ(dt, newDataTable, bbdd, mes, anio, empresa)
        ElseIf empresa = "T. UK. " Then
            bbdd = DB_UK
            newDataTable = FormaTablaUK(dt, newDataTable, bbdd, mes, anio, empresa)
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

            If row("F3").ToString = "T. ES. " Then
                bbdd = DB_TECOZAM
                IDOperario = DevuelveIDOperario(bbdd, row("F1"))
                descOperario = DevuelveDescOperario(bbdd, IDOperario)
                newRow("IDOperario") = IDOperario
                newRow("DescOperario") = descOperario
                newRow("IDGET") = DevuelveIDGET(bbdd, IDOperario)

            ElseIf row("F3").ToString = "FERR. " Then
                bbdd = DB_FERRALLAS
                IDOperario = DevuelveIDOperario(bbdd, row("F1"))
                descOperario = DevuelveDescOperario(bbdd, IDOperario)
                newRow("IDOperario") = IDOperario
                newRow("DescOperario") = descOperario
                newRow("IDGET") = DevuelveIDGET(bbdd, IDOperario)

            ElseIf row("F3").ToString = "SEC. " Then
                bbdd = DB_SECOZAM
                IDOperario = DevuelveIDOperario(bbdd, row("F1"))
                descOperario = DevuelveDescOperario(bbdd, IDOperario)
                newRow("IDOperario") = IDOperario
                newRow("DescOperario") = descOperario
                newRow("IDGET") = DevuelveIDGET(bbdd, IDOperario)

            ElseIf row("F3").ToString = "T. UK. " Then
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
        Dim partes() As String
        ' Copiar los datos de las columnas seleccionadas al nuevo DataTable
        For Each row As DataRow In dt.Rows
            'Verificar si la celda está vacía
            If Len(row("F1").ToString) = 0 Or row("F1").ToString = "TOTAL" Then
                'Return newDataTable
                Exit For ' Salir del bucle si la celda está vacía
            End If

            Dim newRow As DataRow = newDataTable.NewRow()
            partes = row("F1").Split("-"c)

            diccionario = partes(0).Trim()
            descOperario = partes(1).Trim()

            IDOperario = DevuelveIDOperarioDiccionario(bbdd, diccionario)
            newRow("IDOperario") = IDOperario
            newRow("DescOperario") = descOperario
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
            If Len(dr("F1").ToString) = 0 Or dr("F1").ToString = "TOTAL" Then
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
            totallibras = row("F3") + row("F4") + row("F5") + row("F6")
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
            CosteE1 = CosteE1 + dr("F3") + dr("F4") + dr("F5") + dr("F6")
        Next

        For Each dr As DataRow In dtOrdenada.Rows
            CosteEFinal = CosteEFinal + dr("CosteEmpresa")
        Next

        Dim result As DialogResult = MessageBox.Show("El coste del excel introducido es " & CosteE1 & _
        " libras =" & CambioLibraAEuro(dtCambioMoneda, CosteE1, mes, anio) & " €. El del excel resultante es " & CosteEFinal & _
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
                    cambioMoneda = dr("F8")
                End If
            Catch ex As Exception
            End Try
        Next

        Return (totalcoronas * cambioMoneda)
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
        Return dt.Rows(0)("DescOperario")
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
                        acumulaFiniquito = Nz(dtOrdenada(contador)("CosteEmpresa"), 0)
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

        ' Copiar los datos del DataTable original al DataTable ordenado
        For Each dr As DataRow In dtFinal.Rows
            Dim newRow As DataRow = dtFinalOrdenado.NewRow()
            newRow("Empresa") = dr("Empresa")
            newRow("IDGET") = dr("IDGET")
            newRow("IDOperario") = dr("IDOperario")
            newRow("DescOperario") = dr("DescOperario")
            newRow("Mes") = dr("Mes")
            newRow("Anio") = dr("Anio")
            newRow("CosteEmpresa") = dr("CosteEmpresa")
            dtFinalOrdenado.Rows.Add(newRow)
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
            Dim worksheet = package.Workbook.Worksheets.Add(mes & " A3 " & anio)

            ' Copiar los datos de la DataTable a la hoja de cálculo.
            worksheet.Cells("A1").LoadFromDataTable(dtFinalOrdenado, True)

            Dim columnaE As ExcelRange = worksheet.Cells("G2:G" & worksheet.Dimension.End.Row)
            columnaE.Style.Numberformat.Format = "#,##0.00€"

            ' Aplicar formato negrita a la fila 1
            Dim fila1 As ExcelRange = worksheet.Cells(1, 1, 1, worksheet.Dimension.End.Column)
            fila1.Style.Font.Bold = True


            'SEGUNDA HOJA DEL EXCEL QUE ES RESUMEN
            Dim resumenWorksheet = package.Workbook.Worksheets.Add("RESUMEN")
            resumenWorksheet.Cells("A1").LoadFromDataTable(dtResumen, True)

            Dim columnaBResumen As ExcelRange = resumenWorksheet.Cells("B2:B" & worksheet.Dimension.End.Row)
            columnaBResumen.Style.Numberformat.Format = "#,##0.00"

            Dim columnaEResumen As ExcelRange = resumenWorksheet.Cells("E2:E" & worksheet.Dimension.End.Row)
            columnaEResumen.Style.Numberformat.Format = "#,##0.00€"

            ' Aplicar formato negrita a la fila 1
            Dim filaResumen1 As ExcelRange = resumenWorksheet.Cells(1, 1, 1, resumenWorksheet.Dimension.End.Column)
            filaResumen1.Style.Font.Bold = True

            ' Guardar el archivo de Excel.
            package.Save()
        End Using

    End Sub

    Private Sub bIDGET_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bIDGET.Click
        Dim vPersonas As New DataTable
        Dim f As New Filter
        Dim bbdd As String
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
            Dim worksheet = package.Workbook.Worksheets.Add(mes & " HORAS " & anio)

            ' Copiar los datos de la DataTable a la hoja de cálculo.
            worksheet.Cells("A1").LoadFromDataTable(dtFinal, True)

            ' Aplicar formato negrita a la fila 1
            Dim fila1 As ExcelRange = worksheet.Cells(1, 1, 1, worksheet.Dimension.End.Column)
            fila1.Style.Font.Bold = True

            Dim columnaG As ExcelRange = worksheet.Cells("H2:H" & worksheet.Dimension.End.Row)
            columnaG.Style.Numberformat.Format = "dd/mm/yyyy"
            ' Guardar el archivo de Excel.
            package.Save()
        End Using

    End Sub

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

        'Horas Baja por Accidentes en España
        HorasBajaEspaña(Fecha1, Fecha2)
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
        If bbdd = "T. ES. " Then
            bbdd = DB_TECOZAM
        ElseIf bbdd = "FERR. " Then
            bbdd = DB_FERRALLAS
        ElseIf bbdd = "SEC. " Then
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
                If dr("IDHora") = "ACC" Or dr("IDHora") = "CC" Then
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
        rsTrabajo = New BE.DataEngine().Filter(DB_TECOZAM & "..tbObraTrabajo", filtro2)

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
        If bbdd = "T. ES. " Then
            bbdd = DB_TECOZAM
        ElseIf bbdd = "FERR. " Then
            bbdd = DB_FERRALLAS
        ElseIf bbdd = "SEC. " Then
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
        sql &= " where ((fAlta >= '" & Fecha1 & "' AND fAlta <= '" & Fecha2 & "') or fAlta is null)"
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
        sql &= " where ((fAlta >= '" & Fecha1 & "' AND fAlta <= '" & Fecha2 & "') or fAlta is null)"
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

        'GENERACION EXCEL CON LAS 4 PESTAÑAS
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

        Dim ruta As New FileInfo("N:\10. AUXILIARES\00. EXPERTIS\04. CHECK HORAS-A3\" & mes & " CHECK " & mes & anio.Substring(anio.Length - 2) & ".xlsx")
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
            Dim fila1 As ExcelRange = worksheet.Cells(1, 1, 1, worksheet.Dimension.End.Column)
            fila1.Style.Font.Bold = True

            ' HOJA 2
            Dim worksheet2 = package.Workbook.Worksheets.Add(mes & " SI €/NO HORAS " & anio)
            worksheet2.Cells("A1").LoadFromDataTable(dtGenteSiEurosNoHoras, True)
            Dim fila12 As ExcelRange = worksheet2.Cells(1, 1, 1, worksheet2.Dimension.End.Column)
            fila12.Style.Font.Bold = True

            ' HOJA 3
            Dim worksheet3 = package.Workbook.Worksheets.Add(mes & " RATIOS " & anio)
            worksheet3.Cells("A1").LoadFromDataTable(dtRatiosGente, True)

            Dim fila13 As ExcelRange = worksheet3.Cells(1, 1, 1, worksheet3.Dimension.End.Column)
            fila13.Style.Font.Bold = True

            Dim columnaE As ExcelRange = worksheet3.Cells("E2:E" & worksheet3.Dimension.End.Row)
            columnaE.Style.Numberformat.Format = "dd/mm/yyyy"

            Dim columnaF As ExcelRange = worksheet3.Cells("F2:F" & worksheet3.Dimension.End.Row)
            columnaF.Style.Numberformat.Format = "dd/mm/yyyy"

            ' HOJA 4
            Dim worksheet4 = package.Workbook.Worksheets.Add(mes & " DOBLE COTIZACION " & anio)
            worksheet4.Cells("A1").LoadFromDataTable(dtPersonasDobleCoti, True)
            Dim fila14 As ExcelRange = worksheet4.Cells(1, 1, 1, worksheet4.Dimension.End.Column)
            fila14.Style.Font.Bold = True

            ' HOJA 5
            Dim worksheet5 = package.Workbook.Worksheets.Add(mes & " RESUMEN POR CATEGORIA " & anio)
            worksheet5.Cells("A1").LoadFromDataTable(dtResumenCategoriaProfesional, True)
            Dim fila15 As ExcelRange = worksheet5.Cells(1, 1, 1, worksheet5.Dimension.End.Column)
            fila15.Style.Font.Bold = True

            Dim columnaBResumen As ExcelRange = worksheet5.Cells("B2:B" & worksheet5.Dimension.End.Row)
            columnaBResumen.Style.Numberformat.Format = "#,##0.00€"

            ' Guardar el archivo de Excel.
            package.Save()

            MsgBox("Fichero guardado en N:\10. AUXILIARES\00. EXPERTIS\04. CHECK HORAS-A3\")
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
                Dim ratio As Double = Math.Round((costeEmpresa / (horas + horasAdministrativas + horasBaja)), 2)

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
                newRow("HorasTotales") = horas + horasAdministrativas + horasBaja
                newRow("EurosTotales") = costeEmpresa
                If horas = 0 Then
                    newRow("Ratio") = 0
                Else
                    newRow("Ratio") = ratio
                End If
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
        Dim hoja As String = mes & " HORAS " & anio
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
                ExpertisApp.GenerateMessage("El operario con IDGET no tiene categoria profesional" & rowIDGET)
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
        Return dt.Rows(0)("FechaAlta").ToString
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
        Dim hoja As String = mes & " A3 " & anio
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
        Dim filePath As String = "N:\DOCUMENTACION_OFICIAL\Manual_Del_Usuario.docx"

        If System.IO.File.Exists(filePath) Then
            Process.Start(filePath)
        Else
            MessageBox.Show("El archivo no existe en la ruta especificada.", "Archivo no encontrado", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub bExtras_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bExtras.Click
        CD.Filter = "Archivos Excel(*.xls;*.xlsx)|*.xls;*xlsx|Todos los archivos(*.*)|*.*"

        'CD.ShowOpen()
        CD.ShowDialog()

        If CD.FileName <> "" Then
            'lblRuta.Caption = CD.FileName
            lblRuta.Text = CD.FileName
        End If

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

        If mes <> 6 And mes <> 12 Then
            MessageBox.Show("El mes no es ni el 6 ni el 12.")
            Exit Sub
        End If
        anio = "20" & anio


        'UNIFICO TABLAS EN UNA
        Dim dtUnion As New DataTable
        dtUnion.Columns.Add("Empresa")
        dtUnion.Columns.Add("IDGET")
        dtUnion.Columns.Add("IDOperario")
        dtUnion.Columns.Add("DescOperario")
        dtUnion.Columns.Add("IDCategoriaProfesionalSCCP", System.Type.GetType("System.String"))
        dtUnion.Columns.Add("Incentivos", System.Type.GetType("System.Double"))
        dtUnion.Columns.Add("CosteExtra", System.Type.GetType("System.Double"))
        dtUnion.Columns.Add("Mes", System.Type.GetType("System.Double"))
        dtUnion.Columns.Add("Anio", System.Type.GetType("System.Double"))

        'TABLA DE EXTRAS PARA TECOZAM
        Dim dtTeco As New DataTable
        Dim hoja As String = "1"
        Dim rango As String = ""
        rango = "B10:Z10000"
        Try
            dtTeco = ObtenerDatosExcel(ruta, hoja, rango)
            Dim IDOperario As String
            Dim incentivos As String
            For Each row As DataRow In dtTeco.Rows
                'Verificar si la celda está vacía
                If Len(row("F1").ToString) = 0 Then
                    'Return newDataTable
                    Exit For ' Salir del bucle si la celda está vacía
                End If

                Dim newRow As DataRow = dtUnion.NewRow()
                IDOperario = DevuelveIDOperario(DB_TECOZAM, row("F3"))
                newRow("IDOperario") = IDOperario
                newRow("DescOperario") = row("F2")
                newRow("IDCategoriaProfesionalSCCP") = DevuelveIDCategoriaProfesionalSCCPTodasBasesDeDatos(IDOperario)
                newRow("IDGET") = DevuelveIDGET(DB_TECOZAM, IDOperario)
                incentivos = DevuelveIncentivos(DB_TECOZAM, IDOperario)
                If incentivos = 1 Then
                    newRow("CosteExtra") = row("F8")
                    newRow("Incentivos") = 0
                Else
                    If Len(row("F10").ToString()) = 0 Then
                        newRow("Incentivos") = 0
                    Else
                        newRow("Incentivos") = Math.Abs(row("F10"))
                    End If

                    newRow("CosteExtra") = 0
                End If
                newRow("Mes") = mes
                newRow("Anio") = anio
                newRow("Empresa") = "T. ES. "
                dtUnion.Rows.Add(newRow)
            Next
        Catch ex As Exception
            MsgBox(ex.ToString)
            MsgBox("Error al leer las extras de Tecozam pero deja seguir.")
        End Try
        'TABLA DE EXTRAS PARA FERRALLAS
        Dim dtFerr As New DataTable
        hoja = "2"
        rango = "B10:Z10000"
        Try
            dtFerr = ObtenerDatosExcel(ruta, hoja, rango)
            Dim IDOperario As String
            Dim incentivos As String
            For Each row As DataRow In dtTeco.Rows
                'Verificar si la celda está vacía
                If Len(row("F1").ToString) = 0 Then
                    'Return newDataTable
                    Exit For ' Salir del bucle si la celda está vacía
                End If

                Dim newRow As DataRow = dtUnion.NewRow()
                IDOperario = DevuelveIDOperario(DB_FERRALLAS, row("F3"))
                newRow("IDOperario") = IDOperario
                newRow("DescOperario") = row("F2")
                newRow("IDCategoriaProfesionalSCCP") = DevuelveIDCategoriaProfesionalSCCPTodasBasesDeDatos(IDOperario)
                newRow("IDGET") = DevuelveIDGET(DB_FERRALLAS, IDOperario)
                incentivos = DevuelveIncentivos(DB_FERRALLAS, IDOperario)
                If incentivos = 1 Then
                    newRow("CosteExtra") = row("F8")
                    newRow("Incentivos") = 0
                Else
                    If Len(row("F10").ToString()) = 0 Then
                        newRow("Incentivos") = 0
                    Else
                        newRow("Incentivos") = Math.Abs(row("F10"))
                    End If

                    newRow("CosteExtra") = 0
                End If
                newRow("CosteExtra") = row("F8")
                newRow("Mes") = mes
                newRow("Anio") = anio
                newRow("Empresa") = "FERR. "
                dtUnion.Rows.Add(newRow)
            Next
        Catch ex As Exception
            MsgBox("Error al leer las extras de Ferrallas pero deja seguir.")
        End Try


        'TABLA DE EXTRAS PARA FERRALLAS
        Dim dtSeco As New DataTable
        hoja = "3"
        rango = "B10:Z10000"
        Try
            dtSeco = ObtenerDatosExcel(ruta, hoja, rango)
            Dim IDOperario As String
            Dim incentivos As String
            For Each row As DataRow In dtTeco.Rows
                'Verificar si la celda está vacía
                If Len(row("F1").ToString) = 0 Then
                    'Return newDataTable
                    Exit For ' Salir del bucle si la celda está vacía
                End If

                Dim newRow As DataRow = dtUnion.NewRow()
                IDOperario = DevuelveIDOperario(DB_SECOZAM, row("F3"))
                newRow("IDOperario") = IDOperario
                newRow("DescOperario") = row("F2")
                newRow("IDCategoriaProfesionalSCCP") = DevuelveIDCategoriaProfesionalSCCPTodasBasesDeDatos(IDOperario)
                newRow("IDGET") = DevuelveIDGET(DB_SECOZAM, IDOperario)
                incentivos = DevuelveIncentivos(DB_SECOZAM, IDOperario)
                If incentivos = 1 Then
                    newRow("CosteExtra") = row("F8")
                    newRow("Incentivos") = 0
                Else
                    If Len(row("F10").ToString()) = 0 Then
                        newRow("Incentivos") = 0
                    Else
                        newRow("Incentivos") = Math.Abs(row("F10"))
                    End If

                    newRow("CosteExtra") = 0
                End If
                newRow("CosteExtra") = row("F8")
                newRow("Mes") = mes
                newRow("Anio") = anio
                newRow("Empresa") = "SEC. "
                dtUnion.Rows.Add(newRow)
            Next
        Catch ex As Exception
            MsgBox("Error al leer las extras de Secozam pero deja seguir.")
        End Try
        Dim dtImprimirCategorias As DataTable = FormaTablaImprimirExtrasCategorias(dtUnion)

        GenerarExcelExtrasResumen(dtUnion, dtImprimirCategorias, mes, anio)
        GenerarExcelExtras(dtUnion, dtImprimirCategorias, mes, anio)
        MsgBox("Excel creado correctamente en N:\10. AUXILIARES\00. EXPERTIS\03. PAGAS EXTRA\")
    End Sub
    Public Sub GenerarExcelExtras(ByVal dtUnion As DataTable, ByVal dtImprimirCategorias As DataTable, ByVal mes As String, ByVal anio As String)
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial

        For Each row As DataRow In dtUnion.Rows
            row("Incentivos") = CDbl(row("Incentivos")) / 6
            row("CosteExtra") = CDbl(row("CosteExtra")) / 6
        Next

        For Each row As DataRow In dtImprimirCategorias.Rows
            row("1") = CDbl(row("1")) / 6
            row("2") = CDbl(row("2")) / 6
            row("3") = CDbl(row("3")) / 6
            row("4") = CDbl(row("4")) / 6
            row("5") = CDbl(row("5")) / 6
            row("0") = CDbl(row("0")) / 6
        Next

        Dim primero As Integer
        Dim ultimo As Integer
        Dim texto = ""
        If mes = 12 Then
            primero = 1
            ultimo = 6
            anio = anio + 1
        Else
            primero = 6
            ultimo = 12
        End If
        For i As Integer = primero To ultimo
            If Len(primero.ToString()) = 1 Then
                texto = "0" & primero
            Else
                texto = primero
            End If
            Dim ruta As New FileInfo("N:\10. AUXILIARES\00. EXPERTIS\03. PAGAS EXTRA\" & texto & " EXTRAS " & texto & anio.Substring(anio.Length - 2) & ".xlsx")
            Dim rutaCadena As String = ""
            rutaCadena = ruta.FullName

            If File.Exists(rutaCadena) Then
                File.Delete(rutaCadena)
            End If

            Using package As New ExcelPackage(ruta)
                ' HOJA 1
                Dim worksheet = package.Workbook.Worksheets.Add(" EXTRAS POR PERSONA ")
                worksheet.Cells("A1").LoadFromDataTable(dtUnion, True)
                Dim fila1 As ExcelRange = worksheet.Cells(1, 1, 1, worksheet.Dimension.End.Column)
                fila1.Style.Font.Bold = True
                Dim columnaF As ExcelRange = worksheet.Cells("F2:F" & worksheet.Dimension.End.Row)
                columnaF.Style.Numberformat.Format = "#,##0.00€"
                Dim columnaG As ExcelRange = worksheet.Cells("G2:G" & worksheet.Dimension.End.Row)
                columnaG.Style.Numberformat.Format = "#,##0.00€"

                ' HOJA 2
                Dim worksheet2 = package.Workbook.Worksheets.Add(" EXTRAS POR CATEGORIA PROF ")
                worksheet2.Cells("A1").LoadFromDataTable(dtImprimirCategorias, True)
                Dim fila2 As ExcelRange = worksheet2.Cells(1, 1, 1, worksheet2.Dimension.End.Column)
                fila2.Style.Font.Bold = True
                Dim columnaB As ExcelRange = worksheet2.Cells("B2:B" & worksheet2.Dimension.End.Row)
                columnaB.Style.Numberformat.Format = "#,##0.00€"
                Dim columnaC As ExcelRange = worksheet2.Cells("C2:C" & worksheet2.Dimension.End.Row)
                columnaC.Style.Numberformat.Format = "#,##0.00€"
                Dim columnaD As ExcelRange = worksheet2.Cells("D2:D" & worksheet2.Dimension.End.Row)
                columnaD.Style.Numberformat.Format = "#,##0.00€"
                Dim columnaE As ExcelRange = worksheet2.Cells("E2:E" & worksheet2.Dimension.End.Row)
                columnaE.Style.Numberformat.Format = "#,##0.00€"
                Dim columnaF2 As ExcelRange = worksheet2.Cells("F2:F" & worksheet2.Dimension.End.Row)
                columnaF2.Style.Numberformat.Format = "#,##0.00€"
                Dim columnaG2 As ExcelRange = worksheet2.Cells("G2:G" & worksheet2.Dimension.End.Row)
                columnaG2.Style.Numberformat.Format = "#,##0.00€"
                ' Guardar el archivo de Excel.
                package.Save()
            End Using
            primero = primero + 1
        Next
        

    End Sub
    Public Sub GenerarExcelExtrasResumen(ByVal dtUnion As DataTable, ByVal dtImprimirCategorias As DataTable, ByVal mes As String, ByVal anio As String)
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial

        Dim ruta As New FileInfo("N:\10. AUXILIARES\00. EXPERTIS\03. PAGAS EXTRA\" & mes & " EXTRAS RESUMEN" & mes & anio.Substring(anio.Length - 2) & ".xlsx")
        Dim rutaCadena As String = ""
        rutaCadena = ruta.FullName

        If File.Exists(rutaCadena) Then
            File.Delete(rutaCadena)
        End If

        Using package As New ExcelPackage(ruta)

            ' HOJA 1
            Dim worksheet = package.Workbook.Worksheets.Add(" EXTRAS POR PERSONA ")
            worksheet.Cells("A1").LoadFromDataTable(dtUnion, True)
            Dim fila1 As ExcelRange = worksheet.Cells(1, 1, 1, worksheet.Dimension.End.Column)
            fila1.Style.Font.Bold = True
            Dim columnaF As ExcelRange = worksheet.Cells("F2:F" & worksheet.Dimension.End.Row)
            columnaF.Style.Numberformat.Format = "#,##0.00€"
            Dim columnaG As ExcelRange = worksheet.Cells("G2:G" & worksheet.Dimension.End.Row)
            columnaG.Style.Numberformat.Format = "#,##0.00€"

            ' HOJA 2
            Dim worksheet2 = package.Workbook.Worksheets.Add(" EXTRAS POR CATEGORIA PROF ")
            worksheet2.Cells("A1").LoadFromDataTable(dtImprimirCategorias, True)
            Dim fila2 As ExcelRange = worksheet2.Cells(1, 1, 1, worksheet2.Dimension.End.Column)
            fila2.Style.Font.Bold = True
            Dim columnaB As ExcelRange = worksheet2.Cells("B2:B" & worksheet2.Dimension.End.Row)
            columnaB.Style.Numberformat.Format = "#,##0.00€"
            Dim columnaC As ExcelRange = worksheet2.Cells("C2:C" & worksheet2.Dimension.End.Row)
            columnaC.Style.Numberformat.Format = "#,##0.00€"
            Dim columnaD As ExcelRange = worksheet2.Cells("D2:D" & worksheet2.Dimension.End.Row)
            columnaD.Style.Numberformat.Format = "#,##0.00€"
            Dim columnaE As ExcelRange = worksheet2.Cells("E2:E" & worksheet2.Dimension.End.Row)
            columnaE.Style.Numberformat.Format = "#,##0.00€"
            Dim columnaF2 As ExcelRange = worksheet2.Cells("F2:F" & worksheet2.Dimension.End.Row)
            columnaF2.Style.Numberformat.Format = "#,##0.00€"
            Dim columnaG2 As ExcelRange = worksheet2.Cells("G2:G" & worksheet2.Dimension.End.Row)
            columnaG2.Style.Numberformat.Format = "#,##0.00€"
            ' Guardar el archivo de Excel.
            package.Save()
        End Using

    End Sub
    Public Function FormaTablaImprimirExtrasCategorias(ByVal dtUnion As DataTable) As DataTable
        Dim dtResultado As New DataTable()
        dtResultado.Columns.Add("Empresa", GetType(String))
        dtResultado.Columns.Add("1", GetType(Double))
        dtResultado.Columns.Add("2", GetType(Double))
        dtResultado.Columns.Add("3", GetType(Double))
        dtResultado.Columns.Add("4", GetType(Double))
        dtResultado.Columns.Add("5", GetType(Double))
        dtResultado.Columns.Add("0", GetType(Double))

        Dim jprod As Double = 0 : Dim encar As Double = 0 : Dim operar As Double = 0 : Dim tecobra As Double = 0 : Dim staff As Double = 0 : Dim otros As Double = 0
        Dim jprodf As Double = 0 : Dim encarf As Double = 0 : Dim operarf As Double = 0 : Dim tecobraf As Double = 0 : Dim stafff As Double = 0 : Dim otrosf As Double = 0
        Dim jprods As Double = 0 : Dim encars As Double = 0 : Dim operars As Double = 0 : Dim tecobras As Double = 0 : Dim staffs As Double = 0 : Dim otross As Double = 0

        For Each dr As DataRow In dtUnion.Rows
            Dim empresa As String = dr("Empresa").ToString
            Dim categoria As Integer = Convert.ToInt64(dr("IDCategoriaProfesionalSCCP"))
            Dim coste As Double = Convert.ToDouble(dr("CosteExtra"))
            Dim incentivos As Double = Convert.ToDouble(dr("Incentivos"))
            If empresa = "T. ES. " Then
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
            ElseIf empresa = "FERR. " Then
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
            ElseIf empresa = "SEC. " Then
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
        newRow("Empresa") = "T. ES. "
        newRow("1") = jprod
        newRow("2") = encar
        newRow("3") = operar
        newRow("4") = tecobra
        newRow("5") = staff
        newRow("0") = otros
        dtResultado.Rows.Add(newRow)

        '-2.FERRALLAS
        newRow = dtResultado.NewRow()
        newRow("Empresa") = "FERR. "
        newRow("1") = jprodf
        newRow("2") = encarf
        newRow("3") = operarf
        newRow("4") = tecobraf
        newRow("5") = stafff
        newRow("0") = otrosf
        dtResultado.Rows.Add(newRow)

        '-3. SECOZAM
        newRow = dtResultado.NewRow()
        newRow("Empresa") = "SEC. "
        newRow("1") = jprods
        newRow("2") = encars
        newRow("3") = operars
        newRow("4") = tecobras
        newRow("5") = staffs
        newRow("0") = otross
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
End Class
