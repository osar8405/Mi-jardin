Public Class cEmpleados

    Dim oEjecutaConsulta As New cConexionBD.cConexionBD("localhost", "Bituaj", "sa", "Mio34")
    Dim sSentenciaSql As String = ""

    Dim oDt As DataTable

    Public Function InsertaEmpleados(oDatosEmpleado As cDatosEmpleado, iIdUsuario As Integer) As cDatosEmpleado
        Dim oDatosRespuesta As New cDatosEmpleado
        Try

            sSentenciaSql = "SELECT count(*) from empleado where correo_electronico='" & oDatosEmpleado.Correo & "' and correo_electronico<>''"
            If (oEjecutaConsulta.EjecutaExecuteScalar(sSentenciaSql)) > 0 Then
                oDatosRespuesta.MensajeError = "Este correo electrónico ya esta registrado"
            Else

                sSentenciaSql = "INSERT INTO empleado (nombre, apellido_paterno, apellido_materno, correo_electronico, fecha_alta, id_departamento, estatus, usuario) 
values('" & oDatosEmpleado.NombreEmpleado & "', '" & oDatosEmpleado.ApellidoPaterno & "','" & oDatosEmpleado.ApellidoMaterno & "', '" & oDatosEmpleado.Correo & "', '" & oDatosEmpleado.FechaAlta & "', " & oDatosEmpleado.IdDepartamento & ", " & oDatosEmpleado.Estatus & ", " & iIdUsuario & ")"

                If (oEjecutaConsulta.EjecutaExecuteNonQuery(sSentenciaSql)) = True Then
                    oDatosRespuesta.IdEmpleado = 2
                End If
            End If
        Catch ex As Exception
            oDatosRespuesta.MensajeError = ex.Message
        Finally

        End Try
        Return oDatosRespuesta
    End Function

    Public Function ActualizaEmpleados(oDatosEmpleado As cDatosEmpleado, ByVal iIdUsuario As Integer) As cDatosEmpleado
        Dim oDatosRespuesta As New cDatosEmpleado
        Try
            sSentenciaSql = "SELECT count(*) from empleado where (correo_electronico='" & oDatosEmpleado.Correo & "' and id_empleado<>" & oDatosEmpleado.IdEmpleado & ") and correo_electronico<>''"
            If (oEjecutaConsulta.EjecutaExecuteScalar(sSentenciaSql)) > 0 Then
                oDatosRespuesta.MensajeError = "Este correo electrónico ya fue registrado con anterioridad favor de verificar"

            Else

                sSentenciaSql = "Update empleado set nombre='" & oDatosEmpleado.NombreEmpleado & "', apellido_paterno='" & oDatosEmpleado.ApellidoPaterno & "', apellido_materno='" & oDatosEmpleado.ApellidoMaterno & "', correo_electronico='" & oDatosEmpleado.Correo & "', estatus=" & oDatosEmpleado.Estatus & ", rfc='" & oDatosEmpleado.RFC & "', id_departamento=" & oDatosEmpleado.IdDepartamento & ", fecha_alta='" & oDatosEmpleado.FechaAlta & "', usuario_modificacion=" & iIdUsuario & ", fecha_modificacion=GetDate() where id_empleado=" & oDatosEmpleado.IdEmpleado & ""
                If (oEjecutaConsulta.EjecutaExecuteNonQuery(sSentenciaSql)) = True Then
                    oDatosRespuesta.IdEmpleado = 1
                Else
                    oDatosRespuesta.MensajeError = "Ocurrio un error al actulizar el empleado"
                End If
            End If
        Catch ex As Exception
            oDatosRespuesta.MensajeError = ex.Message
        End Try
        Return oDatosRespuesta
    End Function

    Public Function ConsultaListaEmpleados() As cDatosEmpleado()
        Dim arrDatosRespuesta() As cDatosEmpleado = Nothing
        Try

            sSentenciaSql = "Select id_empleado As 'ID', 
E.nombre AS 'Nombre', 
E.apellido_paterno AS 'Apellido_Paterno', 
E.apellido_materno AS 'Apellido_Materno', 
E.estatus AS 'Estatus', 
D.id_departamento AS 'ID_Departamento', 
D.nombre as 'Departamento',
E.correo_electronico as 'Correo_Electrónico',
CONVERT(VARCHAR, E.fecha_alta, 23) AS 'Fecha_Alta' 
FROM empleado E

INNER JOIN departamento D on D.id_departamento = E.id_departamento"

            'Retorna un DT
            oDt = New DataTable
            oDt = oEjecutaConsulta.EjecutaDataAdapter(sSentenciaSql)

            If oDt.Rows.Count Then
                Dim i As Integer = 0
                For Each Row In oDt.Rows
                    ReDim Preserve arrDatosRespuesta(i)
                    Dim oDatosEmpleado As New cDatosEmpleado
                    oDatosEmpleado.IdEmpleado = Row("ID")
                    oDatosEmpleado.NombreEmpleado = Row("Nombre").ToString
                    oDatosEmpleado.ApellidoPaterno = Row("Apellido_Paterno").ToString
                    oDatosEmpleado.ApellidoMaterno = Row("Apellido_Materno").ToString
                    oDatosEmpleado.Estatus = Row("Estatus").ToString
                    oDatosEmpleado.IdDepartamento = Row("ID_Departamento").ToString
                    oDatosEmpleado.NombreDepartamento = Row("Departamento").ToString
                    oDatosEmpleado.Correo = Row("Correo_Electrónico").ToString
                    oDatosEmpleado.FechaAlta = Row("Fecha_Alta").ToString


                    arrDatosRespuesta(i) = oDatosEmpleado
                    oDatosEmpleado = Nothing
                    i += 1
                Next
            End If

        Catch ex As Exception

        End Try
        Return arrDatosRespuesta
    End Function
    Public Function BuscaEmpleados(oDatosEmpleado As cDatosEmpleado)
        Try
            sSentenciaSql = "SELECT id_empleado AS 'ID', 
E.nombre AS 'Nombre', 
E.apellido_paterno AS 'Apellido_Paterno', 
E.apellido_materno AS 'Apellido_Materno', 
IIF(E.estatus=1,'Activo', 'Inactivo') AS 'Estatus', 
D.id_departamento AS 'ID_Departamento', 
D.nombre as 'Departamento',
E.correo_electronico as 'Correo_Electrónico',
CONVERT(VARCHAR, E.fecha_alta, 103) AS 'Fecha_Alta' 
FROM empleado E
INNER JOIN departamento D on D.id_departamento = E.id_departamento

WHERE E.nombre='" & oDatosEmpleado.Busqueda & "' or
E.apellido_paterno='" & oDatosEmpleado.Busqueda & "' or
E.apellido_materno='" & oDatosEmpleado.Busqueda & "' or
D.nombre='" & oDatosEmpleado.Busqueda & "' or
E.correo_electronico='" & oDatosEmpleado.Busqueda & "'"

            oDt = New DataTable
            oDt = oEjecutaConsulta.EjecutaDataAdapter(sSentenciaSql)
        Catch ex As Exception
            oDatosEmpleado.MensajeError = ex.Message
        Finally

        End Try
        Return oDt
    End Function

    Public Function EliminaEmpleados(DatosEmpleado As cDatosEmpleado) As Boolean
        Dim bRespuesta As Boolean = False
        Try
            sSentenciaSql = "DELETE FROM empleado where id_empleado=" & DatosEmpleado.IdEmpleado & ""

            If (oEjecutaConsulta.EjecutaExecuteNonQuery(sSentenciaSql)) = True Then
                bRespuesta = True
            Else
                bRespuesta = False
            End If

        Catch ex As Exception
            DatosEmpleado.MensajeError = ex.Message
            bRespuesta = False
        End Try
        Return bRespuesta
    End Function
End Class



Public Class cDatosEmpleado
    Private iIdEmpleado As Integer
    Private sNombreEmpleado As String
    Private sApellidoPaterno As String
    Private sApellidoMaterno As String
    Private iEstatus As Integer
    Private sCorreo As String
    Private sRFC As String
    Private sFechaAlta As String
    Private iIdDepartamento As Integer
    Private sNombreDepartamento As String
    Private sMensajeError As String
    Private sBusqueda As String

    Public Property IdEmpleado As Integer
        Get
            Return iIdEmpleado
        End Get
        Set(value As Integer)
            iIdEmpleado = value
        End Set
    End Property

    Public Property Correo As String
        Get
            Return sCorreo
        End Get
        Set(value As String)
            sCorreo = value
        End Set
    End Property

    Public Property NombreEmpleado As String
        Get
            Return sNombreEmpleado
        End Get
        Set(value As String)
            sNombreEmpleado = value
        End Set
    End Property

    Public Property ApellidoPaterno As String
        Get
            Return sApellidoPaterno
        End Get
        Set(value As String)
            sApellidoPaterno = value
        End Set
    End Property

    Public Property ApellidoMaterno As String
        Get
            Return sApellidoMaterno
        End Get
        Set(value As String)
            sApellidoMaterno = value
        End Set
    End Property

    Public Property Estatus As Integer
        Get
            Return iEstatus
        End Get
        Set(value As Integer)
            iEstatus = value
        End Set
    End Property

    Public Property RFC As String
        Get
            Return sRFC
        End Get
        Set(value As String)
            sRFC = value
        End Set
    End Property

    Public Property NombreDepartamento As String
        Get
            Return sNombreDepartamento
        End Get
        Set(value As String)
            sNombreDepartamento = value
        End Set
    End Property

    Public Property MensajeError As String
        Get
            Return sMensajeError
        End Get
        Set(value As String)
            sMensajeError = value
        End Set
    End Property

    Public Property IdDepartamento As Integer
        Get
            Return iIdDepartamento
        End Get
        Set(value As Integer)
            iIdDepartamento = value
        End Set
    End Property

    Public Property FechaAlta As String
        Get
            Return sFechaAlta
        End Get
        Set(value As String)
            sFechaAlta = value
        End Set
    End Property

    Public Property Busqueda As String
        Get
            Return sBusqueda
        End Get
        Set(value As String)
            sBusqueda = value
        End Set
    End Property
End Class

