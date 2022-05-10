Imports Solmicro.Expertis.Engine.DAL
Imports Solmicro.Expertis.Business.Negocio
Imports Solmicro.Expertis.Business.Obra
Imports Solmicro.Expertis.Engine.UI
Imports Solmicro.Expertis.Engine
Imports Solmicro.Expertis.Business.ClasesTecozam
Imports System.Math
Imports System.Text

Public Class frmHorasAPP

    Private Sub LoadToolBarActions()
        Me.FormActions.Add("Insertar Horas", AddressOf InsertarPartes)
        Me.AddSeparator()
        Me.FormActions.Add("Comprueba horas y las modifica si ya están insertadas.", AddressOf CompruebaeInsertaPartes)
    End Sub


    ' Sergio Blanco Tecozam 25/04/2017
    Public Sub InsertarPartes()
        Dim clsHor As New HorasTrabajador
        Dim idobra As String = ""
        Dim Nobra As String = ""
        Dim idOperario As String = ""
        Dim Fecha As Date = "1/1/1900"
        Dim i As Integer = 0
        Dim salida As Boolean = True
        Try
            Dim frmFechas As New frmFechas
            frmFechas.ShowDialog()
            Dim mes As String = frmFechas.mes
            Dim anio As String = frmFechas.anio
            Dim desde As Date
            Dim hasta As Date
            Dim descParte As String
            Dim codTrabajo As String = "PT1"
            Dim sTipoHora As String = ""
            Dim DE As New BE.DataEngine

            Select Case mes
                Case "01"
                    desde = CDate("21/12/" & CStr(anio - 1))
                    hasta = CDate("20/" & mes & "/" & CStr(anio))
                Case Else
                    Dim mesNum As Integer = CInt(mes)
                    desde = CDate("21/" & CStr(mesNum - 1) & "/" & CStr(anio))
                    hasta = CDate("20/" & CStr(mesNum) & "/" & CStr(anio))
            End Select

            'Insertamos las horas del personal de oficina
            'clsHor.insertarHorasOficina(desde, hasta, ExpertisApp.UserName)

            'Insertamos las horas del personal de obra

            ' Obtenemos las Horas por trabajador en el intervalo de fechas seleccionados
            'Dim strSelect As String = "select IdParte,IdObra, Fecha, IdOperario,Nhoras, IDHora from tbParteProdHoras where fecha between '" & desde & "' and '" & hasta & "' and (Insertado is null or Insertado=0)"
            Dim dtHoras As DataTable
            Dim Fhoras As New Filter
            Fhoras.Add("fecha", FilterOperator.GreaterThanOrEqual, desde)
            Fhoras.Add("fecha", FilterOperator.LessThanOrEqual, hasta)
            'Fhoras.Add("Insertado", False Or DBNull.Value)
            'Fhoras.Add("Insertado", FilterUnionOperator.Or(False, DBNull.Value))
            Dim FhorasOr As New Filter(FilterUnionOperator.Or)
            FhorasOr.Add("Insertado", DBNull.Value)
            FhorasOr.Add("Insertado", False)
            Dim strWhere As String = " (Insertado is null or insertado='false') and fecha between '" & desde & "' and '" & hasta & "'"
            'MsgBox(strWhere)

            'dtHoras = DE.Filter("tbParteProdHoras", Fhoras, "IdParte,IdObra, Fecha, IdOperario,Nhoras, IDHora, Insertado")
            dtHoras = DE.Filter("tbParteProdHoras", "IdParte,IdObra, Fecha, IdOperario,Nhoras, IDHora, Insertado,idtipoturno", strWhere)

            'dtHoras = AdminData.GetData(strSelect, False)
            If dtHoras.Rows.Count > 0 Then
                For Each dr As DataRow In dtHoras.Rows
                    'Obtenemos la el código de la obra para ponerle el nombre al parte
                    idobra = dr("IDObra")
                    idOperario = dr("idoperario")
                    Fecha = dr("fecha")

                    Dim dtObra As DataTable '= AdminData.GetData("select NObra from tbObraCabecera where IDObra='" & idobra & "'", False)
                    Dim fObra As New Filter
                    fObra.Add("IDObra", FilterOperator.Equal, idobra)
                    dtObra = DE.Filter("tbObraCabecera", fObra, "NObra")

                    If dtObra.Rows.Count > 1 Then
                        MsgBox("no puede haber dos obras con el mismo Código")
                    ElseIf dtObra.Rows.Count = 0 Then
                        MsgBox("la obra no existe")
                    Else
                        For Each drObra As DataRow In dtObra.Rows
                            Nobra = drObra("Nobra")
                        Next
                        descParte = Nobra & " " & desde & " - " & hasta
                        'MsgBox(" el nombre del parte es: " & descParte)
                    End If
                    'Compruebo si existe un parte para el trabajador en dicho dia

                    'Dim sSql = "idoperario='" & dr("idoperario") & "' and FechaInicio='" & dr("Fecha") & "' and IdObra='" & dr("IdObra") & "'"
                    Dim fPRep As New Filter
                    fPRep.Add("idoperario", FilterOperator.Equal, dr("idoperario"))
                    fPRep.Add("FechaInicio", FilterOperator.Equal, dr("Fecha"))
                    fPRep.Add("idobra", FilterOperator.Equal, dr("IdObra"))

                    Dim dtrep As DataTable = DE.Filter("tbObraMODControl", fPRep)
                    'DE.Filter("tbObraMODControl", "*", sSql)
                    'Dim isql2 = "insert into tbParteProdHoras (Insertado) values 'true' where IdParte='" & dr("IdParte") & "'"
                    'Dim isql = "UPDATE tbParteProdHoras SET Insertado = 'true' where IdParte='" & dr("IdParte") & "'"
                    If dtrep.Rows.Count > 0 Then
                        clsHor.ActualizarInsercionParte(dr("IdParte"))
                        'AdminData.Execute(isql)
                    Else

                        'Compruebo si hay horas de trabajo
                        If (Len(dr("idhora")) > 0 And dr("Nhoras") > 0) Then
                            sTipoHora = "HORAS"
                            salida = clsHor.Insertar(dr("IdParte"), idOperario, idobra, dr("Fecha"), codTrabajo, sTipoHora, dr("Nhoras"), descParte, dr("IdObra"), ExpertisApp.UserName, Nobra, Nz(dr("IDTipoTurno"), 4))
                            If salida = False Then
                                Exit For
                            End If
                            
                        Else
                            sTipoHora = dr("IdHora")
                            salida = clsHor.Insertar(dr("IdParte"), idOperario, idobra, dr("Fecha"), codTrabajo, sTipoHora, dr("Nhoras"), descParte, dr("IdObra"), ExpertisApp.UserName, Nobra, Nz(dr("IDTipoTurno"), 4))
                            If salida = False Then
                                Exit For
                            End If
                        End If
                        i = i + 1
                        'clsHor.ActualizarInsercionParte(dr("IdParte"))

                    End If


                Next
                Me.Grid.Refresh()
            End If

        Catch ex As Exception
            MsgBox("Error en el siguiente valor: " & idOperario & " - " & Nobra & " - " & Fecha & "")
            MsgBox("Error: " & ex.Message)


        End Try
        If i > 0 Then
            MsgBox("Se han instertado todos los partes correctamente")
        Else
            MsgBox("No hay partes que insertar", MsgBoxStyle.Information)
        End If


    End Sub

    Public Sub New()
        ' Llamada necesaria para el Diseñador de Windows Forms.
        InitializeComponent()
        LoadToolBarActions()

        ' Agregue cualquier inicialización después de la llamada a InitializeComponent().
    End Sub

    'David Velasco 09/05/22
    'Este metodo lo que hace es comprobar que si la linea de ese mes y ese año no está insertada la inserta
    'Si está insertada pero esta desmarcada en el grid tambien se gestiona que contenga el mismo numero de horas o no.
    'Si está insertada se comprueba que para el idobra, idoperario,fecha que el nhoras sea igual
    'si es igual no hace nada
    'si es distinto se borra esa linea y se inserta la nueva
    Public Sub CompruebaeInsertaPartes()
        Dim clsHor As New HorasTrabajador
        Dim idobra As String = ""
        Dim Nobra As String = ""
        Dim idOperario As String = ""
        Dim nhoras As Integer
        Dim Fecha As Date = "1/1/1900"
        Dim i As Integer = 0
        Dim salida As Boolean = True
        Dim contadorInsercciones As Integer = 0
        Dim pila As New Stack
        Try
            Dim frmFechas As New frmFechas
            frmFechas.ShowDialog()
            Dim mes As String = frmFechas.mes
            Dim anio As String = frmFechas.anio
            Dim desde As Date
            Dim hasta As Date
            Dim descParte As String = ""
            Dim codTrabajo As String = "PT1"
            Dim sTipoHora As String = ""
            Dim DE As New BE.DataEngine

            Select Case mes
                Case "01"
                    desde = CDate("21/12/" & CStr(anio - 1))
                    hasta = CDate("20/" & mes & "/" & CStr(anio))
                Case Else
                    Dim mesNum As Integer = CInt(mes)
                    desde = CDate("21/" & CStr(mesNum - 1) & "/" & CStr(anio))
                    hasta = CDate("20/" & CStr(mesNum) & "/" & CStr(anio))
            End Select
            '1. Hago un select en tbParteProdHoras de todos los operarios que el campo Fecha
            ' esté entre desde y hasta
            Dim dtHoras As DataTable
            Dim Fhoras As New Filter
            Fhoras.Add("fecha", FilterOperator.GreaterThanOrEqual, desde)
            Fhoras.Add("fecha", FilterOperator.LessThanOrEqual, hasta)

            Dim strWhere As String = "fecha between '" & desde & "' and '" & hasta & "'"
            dtHoras = DE.Filter("tbParteProdHoras", "IdParte,IdObra, Fecha, IdOperario,Nhoras, IDHora, Insertado,idtipoturno", strWhere)

            '2. Recorro la tabla. Si no esta insertado lo inserto
            'Si está insertado compruebo que el numero de horas coincide y sino borro e inserto.
            Dim cont As Integer = 0


            For Each dr As DataRow In dtHoras.Rows
                'Obtenemos el código de la obra para ponerle el nombre al parte
                idobra = dr("IDObra")
                idOperario = dr("idoperario")
                Fecha = dr("fecha")
                nhoras = dr("nhoras")

                'Obtengo el NObra del IdObra para formar DescParte
                Dim dtObra As DataTable
                Dim fObra As New Filter
                fObra.Add("IDObra", FilterOperator.Equal, idobra)
                dtObra = DE.Filter("tbObraCabecera", fObra, "NObra")

                If dtObra.Rows.Count > 1 Then
                    MsgBox("No puede haber dos obras con el mismo Código")
                ElseIf dtObra.Rows.Count = 0 Then
                    MsgBox("La obra no existe")
                Else
                    For Each drObra As DataRow In dtObra.Rows
                        Nobra = drObra("Nobra")
                    Next
                    descParte = Nobra & " " & desde & " - " & hasta
                End If

                If dtHoras.Rows(cont)("Insertado").Equals(0) Or IsDBNull(dtHoras.Rows(cont)("Insertado")) Or dtHoras.Rows(cont)("Insertado").ToString().Equals("False") Then
                    'Si no está insertado(IF) lo inserto
                    'Le modificado el valor de no insertado a insertado

                    Dim fPRep As New Filter
                    fPRep.Add("idoperario", FilterOperator.Equal, dr("idoperario"))
                    fPRep.Add("FechaInicio", FilterOperator.Equal, dr("Fecha"))
                    fPRep.Add("idobra", FilterOperator.Equal, dr("IdObra"))
                    Dim dtrep As DataTable = DE.Filter("tbObraMODControl", fPRep)

                    'Aunque este desmarcada, puede que esté insertada. Entonces compruebo que en tbObraModControl no haya
                    'un registro con distinto nº de horas. Si existe un registro y la suma es distinta, me borras y me insertas
                    'si las horas es igual solo me lo marcas para futuras ocasiones.

                    Dim suma2 As Integer = 0
                    For Each drRep As DataRow In dtrep.Rows
                        suma2 += drRep("HorasRealMod")
                    Next

                    If dtrep.Rows.Count > 0 And suma2 = dr("nhoras") Then
                        clsHor.ActualizarInsercionParte(dr("IdParte"))
                        'AdminData.Execute(isql)
                    Else
                        Dim opeCal As New OperarioCalendario
                        Dim strDel As String = "delete from tbObraModControl where IDOperario = '" & idOperario & "' AND FechaInicio = '" & Fecha & "' AND idObra = '" & idobra & "'"
                        'Borro
                        opeCal.Ejecutar(strDel)

                        If (Len(dr("idhora")) > 0 And dr("Nhoras") > 0) Then
                            sTipoHora = "HORAS"
                            salida = clsHor.Insertar(dr("IdParte"), idOperario, idobra, dr("Fecha"), codTrabajo, sTipoHora, dr("Nhoras"), descParte, dr("IdObra"), ExpertisApp.UserName, Nobra, Nz(dr("IDTipoTurno"), 4))
                            If salida = False Then
                                Exit For
                            End If

                        Else
                            sTipoHora = dr("IdHora")
                            salida = clsHor.Insertar(dr("IdParte"), idOperario, idobra, dr("Fecha"), codTrabajo, sTipoHora, dr("Nhoras"), descParte, dr("IdObra"), ExpertisApp.UserName, Nobra, Nz(dr("IDTipoTurno"), 4))
                            If salida = False Then
                                Exit For
                            End If
                        End If
                        i = i + 1
                        clsHor.ActualizarInsercionParte(dr("IdParte"))
                        contadorInsercciones += 1
                    End If
                Else
                    'Si está insertado.Compruebo NHora en tbObraModControl y tiene que coincidir. Si no coinciden que me borre
                    'de tbObraModControl esas horas y me inserte las que estan en tbParteProdHoras
                    Dim fPRep As New Filter
                    fPRep.Add("idoperario", FilterOperator.Equal, idOperario)
                    fPRep.Add("FechaInicio", FilterOperator.Equal, Fecha)
                    fPRep.Add("idobra", FilterOperator.Equal, idobra)


                    Dim dtrep As DataTable = DE.Filter("tbObraMODControl", fPRep)
                    'Recorro la tabla(dtrep) y compruebo si la suma del campo HorasRealMod es igual a
                    'El campo nhoras de la fila que estoy actualmente
                    Dim suma As Integer = 0
                    For Each drRep As DataRow In dtrep.Rows
                        suma += drRep("HorasRealMod")
                    Next


                    If suma = dr("nhoras") Then
                        'Si es igual no se hace nada
                    Else
                        'Si no: Se borran las filas de tbobramodcontrol y se realiza la inserccion de nuevo
                        pila.Push("El operario " & idOperario & " - " & Nobra & " - " & Fecha & " tenía horas diferentes.")
                        'MsgBox("El operario " & idOperario & " - " & Nobra & " - " & Fecha & " tenía horas diferentes.")

                        Dim opeCal As New OperarioCalendario
                        Dim strDel As String = "delete from tbObraModControl where IDOperario = '" & idOperario & "' AND FechaInicio = '" & Fecha & "' AND idObra = '" & idobra & "'"
                        'Borro
                        opeCal.Ejecutar(strDel)

                        'Inserto
                        'Compruebo si hay horas de trabajo
                        If (Len(dr("idhora")) > 0 And dr("Nhoras") > 0) Then
                            sTipoHora = "HORAS"
                            salida = clsHor.Insertar(dr("IdParte"), idOperario, idobra, dr("Fecha"), codTrabajo, sTipoHora, dr("Nhoras"), descParte, dr("IdObra"), ExpertisApp.UserName, Nobra, Nz(dr("IDTipoTurno"), 4))
                            If salida = False Then
                                Exit For
                            End If

                        Else
                            sTipoHora = dr("IdHora")
                            salida = clsHor.Insertar(dr("IdParte"), idOperario, idobra, dr("Fecha"), codTrabajo, sTipoHora, dr("Nhoras"), descParte, dr("IdObra"), ExpertisApp.UserName, Nobra, Nz(dr("IDTipoTurno"), 4))
                            If salida = False Then
                                Exit For
                            End If
                        End If
                        i = i + 1
                        'clsHor.Insertar(dr("IdParte"), idOperario, idobra, dr("Fecha"), codTrabajo, sTipoHora, dr("Nhoras"), descParte, dr("IdObra"), ExpertisApp.UserName, Nobra, Nz(dr("IDTipoTurno"), 4))
                        contadorInsercciones += 1
                    End If
                End If
                    cont += 1
            Next


        Catch ex As Exception
            MsgBox("Error en el siguiente valor: " & idOperario & " - " & Nobra & " - " & Fecha & "")
            MsgBox("Error: " & ex.Message)


        End Try
        If contadorInsercciones > 0 Then
            MsgBox("Se han insertado todos los partes correctamente.")
        Else
            MsgBox("No hay partes que insertar", MsgBoxStyle.Information)
        End If
        Dim pila2 As New Stack
        pila2 = pila
        MsgBox("Se han modificado " & pila2.Count & " registros.")
        If (pila2.Count = 0) Then

        Else
            Dim lista As New StringBuilder()
            For Each value As String In pila2
                lista.Append(value & "" & vbCrLf)
            Next
            MsgBox(lista.ToString())
        End If

    End Sub

    Private Sub frmHorasAPP_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim DE As New BE.DataEngine
        Dim fhoras As New Filter
        fhoras.Add("Fecha", FilterOperator.GreaterThanOrEqual, CDate("21/" & Month(DateAdd(DateInterval.Month, -2, Now))))
        fhoras.Add("Fecha", FilterOperator.LessThanOrEqual, Now)
        Grid.DataSource = DE.Filter(Me.ViewName, fhoras)

    End Sub

End Class