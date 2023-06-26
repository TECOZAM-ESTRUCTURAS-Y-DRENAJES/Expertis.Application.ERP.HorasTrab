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



Public Class CargaHorasJPSTAFF

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
            auto.BorraDatosObraMODControlHorasAdministrativas(DescParte, "xTecozam50R2")
            auto.BorraDatosObraMODControlHorasAdministrativas(DescParte, "xDrenajesPortugal50R2")
            auto.BorraDatosObraMODControlHorasAdministrativas(DescParte, "xTecozamUnitedKingdom4")
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
            If dr("Empresa") = "T. ES." Then
                dtTecozam.ImportRow(dr)
            ElseIf dr("Empresa") = "D. P." Then
                dtPortugal.ImportRow(dr)
            ElseIf dr("Empresa") = "T. UK." Then
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
            IDOficio = DevuelveIDOficio("xTecozam50R2", IDOperario)
            IDCategoriaProfesionalSCCP = DevuelveIDCategoriaProfesionalSCCP("xTecozam50R2", IDOperario)
            Dim filtro As New Filter
            Dim dtObra As New DataTable
            filtro.Add("NObra", FilterOperator.Equal, fila("CentroCoste"))
            dtObra = New BE.DataEngine().Filter("xTecozam50R2" & "..tbObraCabecera", filtro)
            IDObra = dtObra.Rows(0)("IDObra").ToString
            IDTrabajo = ObtieneIDTrabajo("xTecozam50R2", IDObra, "PT1")
            horas = 8 * fila("Porcentaje")

            dtOperarioCalendarioNoProductivo = ObtieneDiasVacacionesYFestivosJP("xTecozam50R2", "xTecozam50R2", IDOperario, Fecha1, Fecha2)
            dtDiasInsertar = ObtieneFechasInsertar("xTecozam50R2", IDOperario, dtCalendario, dtOperarioCalendarioNoProductivo)

            Windows.Forms.Application.DoEvents()
            LProgreso.Text = "Importando : " & IDOperario & " - TECOZAM JP"
            Windows.Forms.Application.DoEvents()

            For Each row As DataRow In dtDiasInsertar.Rows
                Dim fecha As Date = row.Field(Of Date)("Fecha")
                IDAutonumerico = auto.Autonumerico()

                Dim rsTrabajo As New DataTable : Dim filtro2 As New Filter
                filtro2.Add("IDObra", FilterOperator.Equal, IDObra) : filtro2.Add("IdTrabajo", FilterOperator.Equal, IDTrabajo)
                rsTrabajo = New BE.DataEngine().Filter("xTecozam50R2..tbObraTrabajo", filtro2)
                'rsTrabajo = obraTrabajo.Filter(filtro2, , "IdTrabajo, CodTrabajo, DescTrabajo, IdTipoTrabajo, IdSubtipoTrabajo")

                IDTrabajo = rsTrabajo.Rows(0)("IdTrabajo") : CodTrabajo = rsTrabajo.Rows(0)("CodTrabajo")
                Dim DescTrabajo As String = "" : Dim IdTipoTrabajo As String = "" : Dim IdSubTipoTrabajo As String = ""
                DescTrabajo = rsTrabajo.Rows(0)("DescTrabajo") : IdTipoTrabajo = rsTrabajo.Rows(0)("IdTipoTrabajo") : IdSubTipoTrabajo = Nz(rsTrabajo.Rows(0)("IdSubtipotrabajo"), "")
                Dim DescParte As String : DescParte = "JP STAFF " & mes & "-" & año & "-JP"

                txtSQL = "Insert into xTecozam50R2..tbObraMODControl(IdLineaModControl, IdTrabajo, IdObra, CodTrabajo, DescTrabajo, IdTipoTrabajo, " & _
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
            IDOficio = DevuelveIDOficio("xDrenajesPortugal50R2", IDOperario)
            IDCategoriaProfesionalSCCP = DevuelveIDCategoriaProfesionalSCCP("xDrenajesPortugal50R2", IDOperario)
            Dim filtro As New Filter
            Dim dtObra As New DataTable
            filtro.Add("NObra", FilterOperator.Equal, fila("CentroCoste"))
            dtObra = New BE.DataEngine().Filter("xDrenajesPortugal50R2" & "..tbObraCabecera", filtro)
            IDObra = dtObra.Rows(0)("IDObra").ToString
            IDTrabajo = ObtieneIDTrabajo("xDrenajesPortugal50R2", IDObra, "PT1")
            horas = 8 * fila("Porcentaje")

            dtOperarioCalendarioNoProductivo = ObtieneDiasVacacionesYFestivosJP("xTecozam50R2", "xDrenajesPortugal50R2", IDOperario, Fecha1, Fecha2)
            dtDiasInsertar = ObtieneFechasInsertar("xDrenajesPortugal50R2", IDOperario, dtCalendario, dtOperarioCalendarioNoProductivo)

            Windows.Forms.Application.DoEvents()
            LProgreso.Text = "Importando : " & IDOperario & " - DCZ JP"
            Windows.Forms.Application.DoEvents()

            For Each row As DataRow In dtDiasInsertar.Rows
                Dim fecha As Date = row.Field(Of Date)("Fecha")
                IDAutonumerico = auto.Autonumerico()

                Dim rsTrabajo As New DataTable : Dim filtro2 As New Filter
                filtro2.Add("IDObra", FilterOperator.Equal, IDObra) : filtro2.Add("IdTrabajo", FilterOperator.Equal, IDTrabajo)
                rsTrabajo = New BE.DataEngine().Filter("xDrenajesPortugal50R2..tbObraTrabajo", filtro2)
                'rsTrabajo = obraTrabajo.Filter(filtro2, , "IdTrabajo, CodTrabajo, DescTrabajo, IdTipoTrabajo, IdSubtipoTrabajo")

                IDTrabajo = rsTrabajo.Rows(0)("IdTrabajo") : CodTrabajo = rsTrabajo.Rows(0)("CodTrabajo")
                Dim DescTrabajo As String = "" : Dim IdTipoTrabajo As String = "" : Dim IdSubTipoTrabajo As String = ""
                DescTrabajo = rsTrabajo.Rows(0)("DescTrabajo") : IdTipoTrabajo = rsTrabajo.Rows(0)("IdTipoTrabajo") : IdSubTipoTrabajo = Nz(rsTrabajo.Rows(0)("IdSubtipotrabajo"), "")
                Dim DescParte As String : DescParte = "JP STAFF " & mes & "-" & año & "-JP"

                txtSQL = "Insert into xDrenajesPortugal50R2..tbObraMODControl(IdLineaModControl, IdTrabajo, IdObra, CodTrabajo, DescTrabajo, IdTipoTrabajo, " & _
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
            IDOficio = DevuelveIDOficio("xTecozamUnitedKingdom4", IDOperario)
            IDCategoriaProfesionalSCCP = DevuelveIDCategoriaProfesionalSCCP("xTecozamUnitedKingdom4", IDOperario)
            Dim filtro As New Filter
            Dim dtObra As New DataTable
            filtro.Add("NObra", FilterOperator.Equal, fila("CentroCoste"))
            dtObra = New BE.DataEngine().Filter("xTecozamUnitedKingdom4" & "..tbObraCabecera", filtro)
            IDObra = dtObra.Rows(0)("IDObra").ToString
            IDTrabajo = ObtieneIDTrabajo("xTecozamUnitedKingdom4", IDObra, "PT1")
            horas = 8 * fila("Porcentaje")

            dtOperarioCalendarioNoProductivo = ObtieneDiasVacacionesYFestivosJP("xTecozam50R2", "xTecozamUnitedKingdom4", IDOperario, Fecha1, Fecha2)
            dtDiasInsertar = ObtieneFechasInsertarUK("xTecozamUnitedKingdom4", IDOperario, dtCalendario, dtOperarioCalendarioNoProductivo)

            Windows.Forms.Application.DoEvents()
            LProgreso.Text = "Importando : " & IDOperario & " - UK JP"
            Windows.Forms.Application.DoEvents()

            For Each row As DataRow In dtDiasInsertar.Rows
                Dim fecha As Date = row.Field(Of Date)("Fecha")
                IDAutonumerico = auto.Autonumerico()

                Dim rsTrabajo As New DataTable : Dim filtro2 As New Filter
                filtro2.Add("IDObra", FilterOperator.Equal, IDObra) : filtro2.Add("IdTrabajo", FilterOperator.Equal, IDTrabajo)
                rsTrabajo = New BE.DataEngine().Filter("xTecozamUnitedKingdom4..tbObraTrabajo", filtro2)
                'rsTrabajo = obraTrabajo.Filter(filtro2, , "IdTrabajo, CodTrabajo, DescTrabajo, IdTipoTrabajo, IdSubtipoTrabajo")

                IDTrabajo = rsTrabajo.Rows(0)("IdTrabajo") : CodTrabajo = rsTrabajo.Rows(0)("CodTrabajo")
                Dim DescTrabajo As String = "" : Dim IdTipoTrabajo As String = "" : Dim IdSubTipoTrabajo As String = ""
                DescTrabajo = rsTrabajo.Rows(0)("DescTrabajo") : IdTipoTrabajo = rsTrabajo.Rows(0)("IdTipoTrabajo") : IdSubTipoTrabajo = Nz(rsTrabajo.Rows(0)("IdSubtipotrabajo"), "")
                Dim DescParte As String : DescParte = "JP STAFF " & mes & "-" & año & "-JP"

                txtSQL = "Insert into xTecozamUnitedKingdom4..tbObraMODControl(IdLineaModControl, IdTrabajo, IdObra, CodTrabajo, DescTrabajo, IdTipoTrabajo, " & _
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

        Dim result As DialogResult = MessageBox.Show("Hay " & dtTecozam.Rows.Count & " registros de T. ES." & vbCrLf & _
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
        bandera = CheckRegistrosEmpresa(dtTecozam, "xTecozam50R2")
        If bandera = 0 Then
            Exit Sub
        End If

        bandera = CheckRegistrosEmpresa(dtPortugal, "xDrenajesPortugal50R2")
        If bandera = 0 Then
            Exit Sub
        End If

        bandera = CheckRegistrosEmpresa(dtUK, "xTecozamUnitedKingdom4")
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

        auto.BorraDatosObraMODControlHorasAdministrativas(DescParte, "xTecozam50R2")
        auto.BorraDatosObraMODControlHorasAdministrativas(DescParte, "xDrenajesPortugal50R2")
        auto.BorraDatosObraMODControlHorasAdministrativas(DescParte, "xTecozamUnitedKingdom4")
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
            Return dt(0)("IDOficio")
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
        dtCalendario = New BE.DataEngine().Filter("xTecozam50R2..tbMaestroFechas", filtro)

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

        Dim diaLimite As DateTime
        diaLimite = dtDiasCambioDeObra.Rows(0)("Fecha").ToString


        For Each dr As DataRow In dtDias.Rows
            If dr("Fecha") >= diaLimite Then
                dtDiasCambioDeObra.ImportRow(dr)
            End If
        Next

        ' Crear un nuevo DataTable llamado dtCalendario
        Dim dtCalendario As New DataTable
        ' Agregar las columnas Fecha
        dtCalendario.Columns.Add("Fecha", GetType(Date))
        'dtCalendario.Columns.Add("TipoDia", GetType(Integer))
        ' Unir los DataTables dtVacaciones y dtFestivos, trabajador y con cambio de obra en el DataTable dtCalendario
        dtCalendario.Merge(dtVacaciones)
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
        Dim mes As String
        Dim año As String

        mes = InputBox("Introduzca el mes natural", "Formato: mm")
        año = InputBox("Introduzca el año natural", "Formato: aaaa")

        Dim Fecha1 As String
        Dim Fecha2 As String
        Fecha1 = "01/" & mes & "/" & año
        Dim diaMes As String
        diaMes = ObtieneDiaUltimoMes(mes, año)
        Fecha2 = diaMes & "/" & mes & "/" & año & ""

        '-----------TECOZAM--------------
        setHorasOficinaTecozam(mes, año, Fecha1, Fecha2)
        '-----------FERRALLAS------------
        setHorasOficinaFerrallas(mes, año, Fecha1, Fecha2)
        '-----------SECOZAM--------------
        setHorasOficinaSecozam(mes, año, Fecha1, Fecha2)
        '-----------DCZ(No hay nadie)------------------
        setHorasOficinaDCZ(mes, año, Fecha1, Fecha2)
        '-----------UK--------------

        '-----------ESLOVAQUIA------------

        '-----------SUECIA--------------

        '-----------NORUEGA------------
    End Sub
    Public Function getListadoPersonasOfiFerrallas(ByVal Fecha1 As String, ByVal Fecha2 As String) As DataTable
        Dim dt As New DataTable
        Dim sql As String
        '------COMO YO LO DE DEJARIA---------
        sql = "select IDOperario, Obra_Predeterminada from xFerrallas50R2..tbMaestroOperarioSat " & _
        "where (Obra_Predeterminada='12677838' Or Obra_Predeterminada='12677615' Or Obra_Predeterminada='12678141') and " & _
        "(Fecha_Baja is null or (Fecha_Baja>='" & Fecha1 & "' and Fecha_Baja<='" & Fecha2 & "'))" & _
        " Union " & _
        "select IDOperario, Proyecto AS Obra_Predeterminada " & _
        "from xFerrallas50R2..tbHistoricoPersonal " & _
        "where (Proyecto = '12677838' OR Proyecto = '12677615' OR Proyecto = '12678141') and " & _
        "((Fecha >= '" & Fecha1 & "' AND Fecha <= '" & Fecha2 & "'))"

        dt = aux.EjecutarSqlSelect(sql)
        Return dt
    End Function

    Public Function getListadoPersonasOfiDCZ(ByVal Fecha1 As String, ByVal Fecha2 As String) As DataTable
        Dim dt As New DataTable
        '----------FORMA BUENA'-------------
        Dim sql As String
        sql = "select IDOperario, Obra_Predeterminada from xDrenajesPortugal50R2..tbMaestroOperarioSat " & _
        "where Obra_Predeterminada='11860026' and " & _
        "(Fecha_Baja is null or (Fecha_Baja>='" & Fecha1 & "' and Fecha_Baja<='" & Fecha2 & "'))" & _
        " Union " & _
        "select IDOperario, Proyecto AS Obra_Predeterminada " & _
        "from xDrenajesPortugal50R2..tbHistoricoPersonal " & _
        "where (Proyecto = '11860026') and " & _
        "((Fecha >= '" & Fecha1 & "' AND Fecha <= '" & Fecha2 & "'))"

        dt = aux.EjecutarSqlSelect(sql)
        Return dt
    End Function

    Public Function getListadoPersonasOfiTecozam(ByVal Fecha1 As String, ByVal Fecha2 As String) As DataTable
        Dim dt As New DataTable
        Dim sql As String
        '------COMO YO LO DE DEJARIA---------
        sql = "select IDOperario, Obra_Predeterminada from xTecozam50R2..tbMaestroOperarioSat " & _
        "where (Obra_Predeterminada='16895681' Or Obra_Predeterminada='11984995') and " & _
        "(Fecha_Baja is null or (Fecha_Baja>='" & Fecha1 & "' and Fecha_Baja<='" & Fecha2 & "'))" & _
        " Union " & _
        "select IDOperario, Proyecto AS Obra_Predeterminada " & _
        "from xTecozam50R2..tbHistoricoPersonal " & _
        "where (Proyecto = '16895681' OR Proyecto = '11984995') and " & _
        "((Fecha >= '" & Fecha1 & "' AND Fecha <= '" & Fecha2 & "'))"


        'sql = "select IDOperario, Obra_Predeterminada from xTecozam50R2..tbMaestroOperarioSat where idoperario='T3450'"
        dt = aux.EjecutarSqlSelect(sql)
        Return dt
    End Function

    Public Function getListadoPersonasOfiSecozam(ByVal Fecha1 As String, ByVal Fecha2 As String) As DataTable
        Dim dt As New DataTable
        Dim sql As String
        '------COMO YO LO DE DEJARIA---------
        sql = "select IDOperario, Proyecto from xSecozam50R2..tbMaestroOperarioSat " & _
        "where (Obra_Predeterminada='11854299' Or Obra_Predeterminada='11854231') and " & _
        "(Fecha_Baja is null or (Fecha_Baja>='" & Fecha1 & "' and Fecha_Baja<='" & Fecha2 & "'))" & _
        " Union " & _
        "select IDOperario, Proyecto AS Obra_Predeterminada " & _
        "from xSecozam50R2..tbHistoricoPersonal " & _
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
            IDOficio = DevuelveIDOficio("xTecozam50R2", IDOperario)
            IDCategoriaProfesionalSCCP = DevuelveIDCategoriaProfesionalSCCP("xTecozam50R2", IDOperario)
            'IDObra = "15330631"
            'IDObra destino = OFICINA
            IDObra = fila("Obra_Predeterminada")
            'Si es distinto que oficina y almacen
            If IDObra <> "11984995" Or IDObra <> "16895681" Then
                IDObra = DevuelveUltimoCambioObra(IDOperario, "xTecozam50R2")
            End If

            IDTrabajo = ObtieneIDTrabajo("xTecozam50R2", IDObra, "PT1")

            dtOperarioCalendarioNoProductivo = ObtieneDiasVacacionesYFestivos("xTecozam50R2", "xTecozam50R2", IDOperario, Fecha1, Fecha2, dtCalendario)
            dtDiasInsertar = ObtieneFechasInsertar("xTecozam50R2", IDOperario, dtCalendario, dtOperarioCalendarioNoProductivo)

            Windows.Forms.Application.DoEvents()
            LProgreso.Text = "Importando : " & IDOperario & " - TECOZAM OFICINA"
            Windows.Forms.Application.DoEvents()

            For Each row As DataRow In dtDiasInsertar.Rows
                Dim fecha As Date = row.Field(Of Date)("Fecha")
                IDAutonumerico = auto.Autonumerico()

                Dim rsTrabajo As New DataTable : Dim filtro2 As New Filter
                filtro2.Add("IDObra", FilterOperator.Equal, IDObra) : filtro2.Add("IdTrabajo", FilterOperator.Equal, IDTrabajo)
                rsTrabajo = New BE.DataEngine().Filter("xTecozam50R2..tbObraTrabajo", filtro2)
                'rsTrabajo = obraTrabajo.Filter(filtro2, , "IdTrabajo, CodTrabajo, DescTrabajo, IdTipoTrabajo, IdSubtipoTrabajo")

                IDTrabajo = rsTrabajo.Rows(0)("IdTrabajo") : CodTrabajo = rsTrabajo.Rows(0)("CodTrabajo")
                Dim DescTrabajo As String = "" : Dim IdTipoTrabajo As String = "" : Dim IdSubTipoTrabajo As String = ""
                DescTrabajo = rsTrabajo.Rows(0)("DescTrabajo") : IdTipoTrabajo = rsTrabajo.Rows(0)("IdTipoTrabajo") : IdSubTipoTrabajo = Nz(rsTrabajo.Rows(0)("IdSubtipotrabajo"), "")
                Dim DescParte As String : DescParte = "OFICINA" & " " & mes & "-" & año & "-OFI"

                txtSQL = "Insert into xTecozam50R2..tbObraMODControl(IdLineaModControl, IdTrabajo, IdObra, CodTrabajo, DescTrabajo, IdTipoTrabajo, " & _
                        "IdSubTipoTrabajo, IdOperario, IdCategoria, IdHora, FechaInicio, HorasRealMod, TasaRealModA, " & _
                         "ImpRealModA, HorasFactMod, ImpFactModA, DescParte, Facturable, FechaCreacionAudi, FechaModificacionAudi, Usuarioaudi, IDOficio, IdTipoTurno, HorasAdministrativas, IDCategoriaProfesionalSCCP) " & _
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
        'MsgBox("Horas de la gente de oficina de Tecozam han sido insertadas correctamente.", vbInformation + vbOKOnly, "STAFF OFICINA")

    End Sub

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
            IDOficio = DevuelveIDOficio("xFerrallas50R2", IDOperario)
            IDCategoriaProfesionalSCCP = DevuelveIDCategoriaProfesionalSCCP("xFerrallas50R2", IDOperario)
            'IDObra = "12677615"
            IDObra = fila("Obra_Predeterminada")
            'Si es distinto que  ferrallas, oficina y secozam
            If IDObra <> "12677838" Or IDObra <> "12677615" Or IDObra <> "12678141" Then
                IDObra = DevuelveUltimoCambioObra(IDOperario, "xFerrallas50R2")
            End If

            IDTrabajo = ObtieneIDTrabajo("xFerrallas50R2", IDObra, "PT1")
            'Este es xTecozam50R2 porque coje el calendario de España
            dtOperarioCalendarioNoProductivo = ObtieneDiasVacacionesYFestivos("xTecozam50R2", "xFerrallas50R2", IDOperario, Fecha1, Fecha2, dtCalendario)
            dtDiasInsertar = ObtieneFechasInsertar("xFerrallas50R2", IDOperario, dtCalendario, dtOperarioCalendarioNoProductivo)

            Windows.Forms.Application.DoEvents()
            LProgreso.Text = "Importando : " & IDOperario & " - FERRALLAS OFICINA"
            Windows.Forms.Application.DoEvents()

            For Each row As DataRow In dtDiasInsertar.Rows
                Dim fecha As Date = row.Field(Of Date)("Fecha")
                IDAutonumerico = auto.Autonumerico()

                Dim rsTrabajo As New DataTable : Dim filtro2 As New Filter
                filtro2.Add("IDObra", FilterOperator.Equal, IDObra) : filtro2.Add("IdTrabajo", FilterOperator.Equal, IDTrabajo)
                rsTrabajo = New BE.DataEngine().Filter("xFerrallas50R2..tbObraTrabajo", filtro2)
                'rsTrabajo = obraTrabajo.Filter(filtro2, , "IdTrabajo, CodTrabajo, DescTrabajo, IdTipoTrabajo, IdSubtipoTrabajo")

                IDTrabajo = rsTrabajo.Rows(0)("IdTrabajo") : CodTrabajo = rsTrabajo.Rows(0)("CodTrabajo")
                Dim DescTrabajo As String = "" : Dim IdTipoTrabajo As String = "" : Dim IdSubTipoTrabajo As String = ""
                DescTrabajo = rsTrabajo.Rows(0)("DescTrabajo") : IdTipoTrabajo = rsTrabajo.Rows(0)("IdTipoTrabajo") : IdSubTipoTrabajo = Nz(rsTrabajo.Rows(0)("IdSubtipotrabajo"), "")
                Dim DescParte As String : DescParte = "OFICINA" & " " & mes & "-" & año & "-OFI"

                txtSQL = "Insert into xFerrallas50R2..tbObraMODControl(IdLineaModControl, IdTrabajo, IdObra, CodTrabajo, DescTrabajo, IdTipoTrabajo, " & _
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
            IDOficio = DevuelveIDOficio("xSecozam50R2", IDOperario)
            IDCategoriaProfesionalSCCP = DevuelveIDCategoriaProfesionalSCCP("xSecozam50R2", IDOperario)
            'IDObra = "11854231"
            IDObra = fila("Obra_Predeterminada")
            'Si es distinto que oficina y secozam
            If IDObra <> "11854299" Or IDObra <> "11854231" Then
                IDObra = DevuelveUltimoCambioObra(IDOperario, "xSecozam50R2")
            End If

            IDTrabajo = ObtieneIDTrabajo("xSecozam50R2", IDObra, "PT1")
            'Este es xTecozam50R2 porque coje el calendario de España
            dtOperarioCalendarioNoProductivo = ObtieneDiasVacacionesYFestivos("xTecozam50R2", "xSecozam50R2", IDOperario, Fecha1, Fecha2, dtCalendario)
            dtDiasInsertar = ObtieneFechasInsertar("xSecozam50R2", IDOperario, dtCalendario, dtOperarioCalendarioNoProductivo)

            Windows.Forms.Application.DoEvents()
            LProgreso.Text = "Importando : " & IDOperario & " - SECOZAM OFICINA"
            Windows.Forms.Application.DoEvents()

            For Each row As DataRow In dtDiasInsertar.Rows
                Dim fecha As Date = row.Field(Of Date)("Fecha")
                IDAutonumerico = auto.Autonumerico()

                Dim rsTrabajo As New DataTable : Dim filtro2 As New Filter
                filtro2.Add("IDObra", FilterOperator.Equal, IDObra) : filtro2.Add("IdTrabajo", FilterOperator.Equal, IDTrabajo)
                rsTrabajo = New BE.DataEngine().Filter("xSecozam50R2..tbObraTrabajo", filtro2)
                'rsTrabajo = obraTrabajo.Filter(filtro2, , "IdTrabajo, CodTrabajo, DescTrabajo, IdTipoTrabajo, IdSubtipoTrabajo")

                IDTrabajo = rsTrabajo.Rows(0)("IdTrabajo") : CodTrabajo = rsTrabajo.Rows(0)("CodTrabajo")
                Dim DescTrabajo As String = "" : Dim IdTipoTrabajo As String = "" : Dim IdSubTipoTrabajo As String = ""
                DescTrabajo = rsTrabajo.Rows(0)("DescTrabajo") : IdTipoTrabajo = rsTrabajo.Rows(0)("IdTipoTrabajo") : IdSubTipoTrabajo = Nz(rsTrabajo.Rows(0)("IdSubtipotrabajo"), "")
                Dim DescParte As String : DescParte = "OFICINA" & " " & mes & "-" & año & "-OFI"

                txtSQL = "Insert into xSecozam50R2..tbObraMODControl(IdLineaModControl, IdTrabajo, IdObra, CodTrabajo, DescTrabajo, IdTipoTrabajo, " & _
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
            IDOficio = DevuelveIDOficio("xDrenajesPortugal50R2", IDOperario)
            IDCategoriaProfesionalSCCP = DevuelveIDCategoriaProfesionalSCCP("xDrenajesPortugal50R2", IDOperario)
            IDObra = "11860026"

            If IDObra <> "11860026" Then
                IDObra = DevuelveUltimoCambioObra(IDOperario, "xDrenajesPortugal50R2")
            End If

            IDTrabajo = ObtieneIDTrabajo("xDrenajesPortugal50R2", IDObra, "PT1")
            'Este es xTecozam50R2 porque coje el calendario de España
            dtOperarioCalendarioNoProductivo = ObtieneDiasVacacionesYFestivos("xTecozam50R2", "xDrenajesPortugal50R2", IDOperario, Fecha1, Fecha2, dtCalendario)
            dtDiasInsertar = ObtieneFechasInsertar("xDrenajesPortugal50R2", IDOperario, dtCalendario, dtOperarioCalendarioNoProductivo)

            Windows.Forms.Application.DoEvents()
            LProgreso.Text = "Importando : " & IDOperario & " - DCZ OFICINA"
            Windows.Forms.Application.DoEvents()

            For Each row As DataRow In dtDiasInsertar.Rows
                Dim fecha As Date = row.Field(Of Date)("Fecha")
                IDAutonumerico = auto.Autonumerico()

                Dim rsTrabajo As New DataTable : Dim filtro2 As New Filter
                filtro2.Add("IDObra", FilterOperator.Equal, IDObra) : filtro2.Add("IdTrabajo", FilterOperator.Equal, IDTrabajo)
                rsTrabajo = New BE.DataEngine().Filter("xDrenajesPortugal50R2..tbObraTrabajo", filtro2)
                'rsTrabajo = obraTrabajo.Filter(filtro2, , "IdTrabajo, CodTrabajo, DescTrabajo, IdTipoTrabajo, IdSubtipoTrabajo")

                IDTrabajo = rsTrabajo.Rows(0)("IdTrabajo") : CodTrabajo = rsTrabajo.Rows(0)("CodTrabajo")
                Dim DescTrabajo As String = "" : Dim IdTipoTrabajo As String = "" : Dim IdSubTipoTrabajo As String = ""
                DescTrabajo = rsTrabajo.Rows(0)("DescTrabajo") : IdTipoTrabajo = rsTrabajo.Rows(0)("IdTipoTrabajo") : IdSubTipoTrabajo = Nz(rsTrabajo.Rows(0)("IdSubtipotrabajo"), "")
                Dim DescParte As String : DescParte = "OFICINA" & " " & mes & "-" & año & "-OFI"

                txtSQL = "Insert into xDrenajesPortugal50R2..tbObraMODControl(IdLineaModControl, IdTrabajo, IdObra, CodTrabajo, DescTrabajo, IdTipoTrabajo, " & _
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
        Dim connectionString As String = "Data Source=stecodesarr;Initial Catalog=xTecozam50R2;User ID=sa;Password=180M296;"
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
               & "OFICINA: TECOZAM-FERRALLAS-SECOZAM ", MsgBoxStyle.OkOnly, "Ayuda")

        'Dim IDObra As String
        'IDObra = DevuelveIDObra("xTecozamUnitedKingdom4", "Tuk08")
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
        CD.Filter = "Excel (*.xls)|*.xls"

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
        Dim rango2 As String = "A12:AG100"
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
            bbdd = "xTecozam50R2"
        ElseIf basededatos1 = "FERRALLAS" Then
            bbdd = "xFerrallas50R2"
        ElseIf basededatos1 = "DCZ" Then
            bbdd = "xDrenajesPortugal50R2"
        ElseIf basededatos1 = "UK" Then
            bbdd = "xTecozamUnitedKingdom4"
        ElseIf basededatos1 = "SUECIA" Then
            bbdd = "xTecozamSuecia4"
        ElseIf basededatos1 = "NORUEGA" Then
            bbdd = "xTecozamNoruega4"
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
                        IDOperario = drHora(0)
                        Windows.Forms.Application.DoEvents()
                        LProgreso.Text = "Importando : " & IDOperario & " - " & fecha
                        Windows.Forms.Application.DoEvents()

                        If Length(drHora(columna)) > 0 Then

                            If IsNumeric(drHora(columna)) = True Then
                                hora = drHora(columna)
                                tipoHora = "HORAS"

                                InsertarPorBaseDeDatos(IDOperario, numero, fecha, trabajo, tipoHora, hora, sNombreUnicoGlobal, numero, idtrab, bbdd)

                            Else
                                hora = 0
                                tipoHora = drHora(columna)
                                InsertarPorBaseDeDatos(IDOperario, numero, fecha, trabajo, tipoHora, hora, sNombreUnicoGlobal, numero, idtrab, bbdd)
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
            rsOperario = New BE.DataEngine().Filter(bbdd & "..tbMaestroOperario", filtro3)

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

            rsCalendarioCentro = New BE.DataEngine().Filter("xTecozam50R2..tbCalendarioCentro", filtro)

            'David 15/11/21 En vez de <>0 ponia "=0"
            'Si tiene datos es que es festivo
            If rsCalendarioCentro.Rows.Count <> 0 Then
                iVeces = 1
                N_Horas = N_Horas
                Coste_Hora = rsOperario.Rows(0)("c_h_e")
                Tipo_Hora = "HE"
            Else
                'Si no es festivo
                If rsOperario.Rows(0)("Jornada_Laboral") >= N_Horas Then
                    'Todas son horas normales
                    iVeces = 1
                    N_Horas = N_Horas
                    Coste_Hora = rsOperario.Rows(0)("c_h_n")
                    Tipo_Hora = "HO"
                Else
                    'Hay horas normales y horas extras, primero pongo las horas normales
                    iVeces = 2
                    Coste_Hora = rsOperario.Rows(0)("c_h_n")
                    N_Horas = rsOperario.Rows(0)("Jornada_Laboral")
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

                IDCategoriaProfesionalSCCP = DevuelveIDCategoriaProfesionalSCCP(Operario)
                IDOficio = DevuelveIDOficio(Operario)

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

                'Inserto
                'Conexion.Execute(txtSQL)
                auto.Ejecutar(txtSQL)

                'Cambio valores, pongo las horas extras
                Coste_Hora = rsOperario.Rows(0)("c_h_x")
                N_Horas = CDbl(HorasOrigen) - CDbl(rsOperario.Rows(0)("Jornada_Laboral"))
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

        sql2 = "Select * from xTecozamUnitedKingdom4..tbObraCabecera where NObra='" & NObra & "'"
        dtObra = aux.EjecutarSqlSelect(sql2)

        Return dtObra.Rows(0)("IDObra")
    End Function

    Private Sub bA3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bA3.Click

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
        MsgBox("El excel se ha guardado correctamente.")
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
            bbdd = "xTecozam50R2"
            newDataTable = FormaTablaEspaña(dt, newDataTable, bbdd, mes, anio, empresa)
        ElseIf empresa = "FERR." Then
            bbdd = "xFerrallas50R2"
            newDataTable = FormaTablaEspaña(dt, newDataTable, bbdd, mes, anio, empresa)
        ElseIf empresa = "SEC." Then
            bbdd = "xSecozam50R2"
            newDataTable = FormaTablaEspaña(dt, newDataTable, bbdd, mes, anio, empresa)
        ElseIf empresa = "D. P." Then
            bbdd = "xDrenajesPortugal50R2"
            newDataTable = FormaTablaDCZ(dt, newDataTable, bbdd, mes, anio, empresa)
        ElseIf empresa = "T. UK." Then
            bbdd = "xTecozamUnitedKingdom4"
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

            If row("F3").ToString = "T. ES." Then
                bbdd = "xTecozam50R2"
                IDOperario = DevuelveIDOperario(bbdd, row("F1"))
                descOperario = DevuelveDescOperario(bbdd, IDOperario)
                newRow("IDOperario") = IDOperario
                newRow("DescOperario") = descOperario
                newRow("IDGET") = DevuelveIDGET(bbdd, IDOperario)
                
            ElseIf row("F3").ToString = "FERR." Then
                bbdd = "xFerrallas50R2"
                IDOperario = DevuelveIDOperario(bbdd, row("F1"))
                descOperario = DevuelveDescOperario(bbdd, IDOperario)
                newRow("IDOperario") = IDOperario
                newRow("DescOperario") = descOperario
                newRow("IDGET") = DevuelveIDGET(bbdd, IDOperario)

            ElseIf row("F3").ToString = "SEC." Then
                bbdd = "xSecozam50R2"
                IDOperario = DevuelveIDOperario(bbdd, row("F1"))
                descOperario = DevuelveDescOperario(bbdd, IDOperario)
                newRow("IDOperario") = IDOperario
                newRow("DescOperario") = descOperario
                newRow("IDGET") = DevuelveIDGET(bbdd, IDOperario)

            ElseIf row("F3").ToString = "T. UK." Then
                bbdd = "xTecozamUnitedKingdom4"
                Diccionario = row("F1")
                IDOperario = DevuelveIDOperarioDiccionario(bbdd, Diccionario)
                descOperario = DevuelveDescOperario(bbdd, IDOperario)
                newRow("IDOperario") = IDOperario
                newRow("DescOperario") = descOperario
                newRow("IDGET") = DevuelveIDGET(bbdd, IDOperario)

            ElseIf row("F3").ToString = "D. P." Then
                bbdd = "xDrenajesPortugal50R2"
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

        Dim result As DialogResult = MessageBox.Show("El coste del excel introducido es " & CosteE1 & "€. El del excel resultante es " & CosteEFinal & "€.", "¿Desea Continuar?", MessageBoxButtons.YesNo)
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

        Dim result As DialogResult = MessageBox.Show("El coste del excel introducido es " & CosteE1 & "€. El del excel resultante es " & CosteEFinal & "€.", "¿Desea Continuar?", MessageBoxButtons.YesNo)
        If result = DialogResult.No Then
            Return Nothing
            Exit Function
        End If
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

        Dim result As DialogResult = MessageBox.Show("El coste del excel introducido es " & CosteE1 & " libras =" & CambioLibraAEuro(dtCambioMoneda, CosteE1, mes, anio) & " €. El del excel resultante es " & CosteEFinal & "€.", "¿Desea Continuar?", MessageBoxButtons.YesNo)
        If result = DialogResult.No Then
            Return Nothing
            Exit Function
        End If
        Return dtOrdenada
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
            CosteE1 = CosteE1 + dr("F8")
        Next

        For Each dr As DataRow In dtOrdenada.Rows
            CosteEFinal = CosteEFinal + dr("CosteEmpresa")
        Next

        Dim result As DialogResult = MessageBox.Show("El coste del excel introducido es " & CosteE1 & "€. El del excel resultante es " & CosteEFinal & "€.", "¿Desea Continuar?", MessageBoxButtons.YesNo)
        If result = DialogResult.No Then
            Return Nothing
            Exit Function
        End If
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
                    dtOrdenada.Rows(contador)("CosteEmpresa") = dtOrdenada.Rows(contador)("CosteEmpresa") + acumulaFiniquito
                    dtFinal.ImportRow(fila)
                    acumulaFiniquito = 0
                Else
                    If dtOrdenada.Rows(contador)("IDOperario").ToString = dtOrdenada.Rows(contador + 1)("IDOperario").ToString Then
                        acumulaFiniquito = dtOrdenada(contador)("CosteEmpresa")
                    End If
                End If
            Catch ex As Exception
                dtOrdenada.Rows(contador)("CosteEmpresa") = dtOrdenada.Rows(contador)("CosteEmpresa") + acumulaFiniquito
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

        Dim ruta As New FileInfo("N:\10. AUXILIARES\00. EXPERTIS\02. A3\" & mes & " A3 " & mes & anio.Substring(anio.Length - 2) & ".xlsx")
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
            worksheet.Cells("A1").LoadFromDataTable(dtFinal, True)

            Dim columnaE As ExcelRange = worksheet.Cells("D2:E" & worksheet.Dimension.End.Row)
            columnaE.Style.Numberformat.Format = "#,##0.00€"

            ' Aplicar formato negrita a la fila 1
            Dim fila1 As ExcelRange = worksheet.Cells(1, 1, 1, worksheet.Dimension.End.Column)
            fila1.Style.Font.Bold = True

            ' Guardar el archivo de Excel.
            package.Save()
        End Using

    End Sub

    Private Sub bIDGET_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bIDGET.Click
        Dim vPersonas As New DataTable
        Dim f As New Filter
        Dim bbdd As String
        vPersonas = New BE.DataEngine().Filter("xTecozam50R2..vPersonasTFSD", f, , "FechaAlta asc")

        For Each dr As DataRow In vPersonas.Rows
            Dim valor As String = dr("IDOperario").ToString()
            If valor(0) = "T"c Or (Char.IsDigit(valor(0))) Then
                bbdd = "xTecozamUnitedKingdom4"
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
        dt = New BE.DataEngine().Filter("xTecozam50R2..tbMaestroContador", f)

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
        dt = New BE.DataEngine().Filter("xTecozam50R2..tbMaestroContador", f)

        Dim texto As String
        texto = dt.Rows(0)("Texto")

        Dim numerico As Integer
        numerico = dt.Rows(0)("Contador")

        numerico = numerico + 1

        Dim sql As String
        sql = "UPDATE xTecozam50R2..tbMaestroContador set Contador= " & numerico & " Where IDContador='IDGET'"

        aux.EjecutarSql(sql)

    End Sub
End Class