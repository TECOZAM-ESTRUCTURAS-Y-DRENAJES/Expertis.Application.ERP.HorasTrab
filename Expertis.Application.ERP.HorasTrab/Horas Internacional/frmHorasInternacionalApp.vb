Imports Solmicro.Expertis.Engine
Imports System.Windows.Forms
Imports OfficeOpenXml
Imports OfficeOpenXml.Style
Imports System.Drawing
Imports Solmicro.Expertis.Engine.UI
Imports Solmicro.Expertis.Business.ClasesTecozam

Public Class frmHorasInternacionalApp

    Public Sub New()
        ' Llamada necesaria para el Diseñador de Windows Forms.
        InitializeComponent()
        LoadToolbarActions()
    End Sub

    Private Sub frmHorasInternacionalApp_QueryExecuting(ByVal sender As System.Object, ByRef e As Solmicro.Expertis.Engine.UI.QueryExecutingEventArgs) Handles MyBase.QueryExecuting
        e.Filter.Add("NObra", FilterOperator.Equal, advNObra.Text)
        e.Filter.Add("IDOperario", FilterOperator.Equal, advIDOperario.Text)
        e.Filter.Add("FechaParte", FilterOperator.GreaterThanOrEqual, clbFecha1.Value, FilterType.DateTime)
        e.Filter.Add("FechaParte", FilterOperator.LessThanOrEqual, clbFecha2.Value, FilterType.DateTime)
    End Sub

    Private Sub LoadToolbarActions()
        'David V 01/08/2024
        If ExpertisApp.DataBaseName = "xTecozamNorge50R2" Then
            Me.FormActions.Add("Generar excel horas Noruega.", AddressOf exportacionNO)
        End If
        'David V 01/10/2024
        If ExpertisApp.DataBaseName = "xTecozamUnitedKingdom50R2" Then
            Me.FormActions.Add("Generar excel horas Reino Unido.", AddressOf exportacionUK)
            Me.FormActions.Add("Generar excel vacaciones Reino Unido.", AddressOf exportacionUKVacaciones)
        End If
        'David V 02/10/2024
        If ExpertisApp.DataBaseName = "xDrenajesPortugal50R2" Then
            Me.FormActions.Add("Generar excel horas Portugal.", AddressOf exportacionDCZ)
        End If

        Me.FormActions.Add("Insertar horas.", AddressOf insertarHoras)
    End Sub

    Public Sub exportacionNO()
        ' Crear una instancia de la clase ExportacionCuadranteNoruega
        Dim tablaOriginal As String = "frmMntoHorasInternacionalTecozam"
        Dim exportacion As New ExportacionNoruegaCuadrante()
        ' Llamar al método generaExcelNoruega
        exportacion.tablaDatos = tablaOriginal
        exportacion.tipoExportacion = "TECOZAM"
        exportacion.generaExcelNoruega()
    End Sub

    Public Sub exportacionUK()
        ' Crear una instancia de la clase ExportacionCuadranteNoruega
        Dim tablaOriginal As String = "frmMntoHorasInternacionalTecozam"
        Dim exportacion As New ExportacionUKCuadrante()
        ' Llamar al método generaExcelNoruega
        exportacion.tablaDatos = tablaOriginal
        exportacion.generaExcel()
    End Sub

    Public Sub exportacionDCZ()
        ' Crear una instancia de la clase ExportacionCuadranteNoruega
        Dim tablaOriginal As String = "frmMntoHorasInternacionalTecozam"
        Dim exportacion As New ExportacionPortugalCuadrante()
        ' Llamar al método generaExcelNoruega
        exportacion.tablaDatos = tablaOriginal
        exportacion.generaExcel()
    End Sub

    Public Sub exportacionUKVacaciones()
        ' Crear una instancia de la clase ExportacionCuadranteNoruega
        Dim tablaOriginal As String = "frmMntoHorasInternacionalTecozam"
        Dim exportacion As New ExportacionUKVacaciones()
        ' Llamar al método generaExcelNoruega
        exportacion.tablaDatos = tablaOriginal
        exportacion.generaExcel()
    End Sub

    Public Sub insertarHoras()
        Dim fecha1 As String
        Dim fecha2 As String

        Dim frmFechas As New frmInformeFecha
        frmFechas.ShowDialog()

        fecha1 = frmFechas.fecha1
        fecha2 = frmFechas.fecha2

        Dim ClaseCreacion As New CreacionHoras()
        ClaseCreacion.CrearHorasApp(fecha1, fecha2)
    End Sub

End Class