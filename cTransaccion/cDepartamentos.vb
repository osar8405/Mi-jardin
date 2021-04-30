Public Class cDepartamentos

    Dim oEjecutaConsulta As New cConexionBD.cConexionBD("localhost", "Bituaj", "sa", "Mio34")
    Dim sSentenciaSql As String = ""

    Dim oDt As DataTable
    Dim oDataAdapter As SqlClient.SqlDataAdapter

    Public Function ObtieneListaDepartamentos() As cDatosDepartamento()
        Dim oDatosRespuesta() As cDatosDepartamento = Nothing
        Try
            sSentenciaSql = "SELECT id_departamento as 'ID',
nombre as 'Departamento' 
from departamento
where estatus=1
UNION ALL SELECT 0,' Elije una opción..'
order by nombre"

            'recupera en un DataTable
            oDt = New DataTable
            oDt = oEjecutaConsulta.EjecutaDataAdapter(sSentenciaSql)
            If oDt.Rows.Count > 0 Then
                Dim i As Integer = 0
                For Each Row In oDt.Rows
                    ReDim Preserve oDatosRespuesta(i)
                    Dim oDatosDepartamento = New cDatosDepartamento

                    oDatosDepartamento.IdDepartamento = Row("ID")
                    oDatosDepartamento.NomDepartamento = Row("Departamento").ToString

                    oDatosRespuesta(i) = oDatosDepartamento
                    oDatosDepartamento = Nothing

                    i += 1
                Next

            End If

        Catch ex As Exception

        End Try

        Return oDatosRespuesta
    End Function
End Class

Public Class cDatosDepartamento
    Private iIdDepartamento As Integer
    Private sNomDepartamento As String
    Private iEstatus As Integer
    Private sMensajeError As String

    Public Property IdDepartamento As Integer
        Get
            Return iIdDepartamento
        End Get
        Set(value As Integer)
            iIdDepartamento = value
        End Set
    End Property

    Public Property NomDepartamento As String
        Get
            Return sNomDepartamento
        End Get
        Set(value As String)
            sNomDepartamento = value
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

    Public Property MensajeError As String
        Get
            Return sMensajeError
        End Get
        Set(value As String)
            sMensajeError = value
        End Set
    End Property


End Class
