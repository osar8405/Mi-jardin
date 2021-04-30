' NOTE: You can use the "Rename" command on the context menu to change the class name "wcfUsuarios" in code, svc and config file together.
' NOTE: In order to launch WCF Test Client for testing this service, please select wcfUsuarios.svc or wcfUsuarios.svc.vb at the Solution Explorer and start debugging.



Imports wcfUsuarios

Public Class wcfUsuarios
    Implements IwcfUsuarios


    Public Function AutenticaUsuarios(ByVal oDatos As cTransaccion.cDatosUsuario) As cTransaccion.cDatosUsuario Implements IwcfUsuarios.AutenticaUsuarios

        Dim oDatosRespuesta As New cTransaccion.cDatosUsuario, oCredenciales As New cTransaccion.cAutenticaUsuarios
        Try
            oDatosRespuesta = oCredenciales.ValidaCredencial(oDatos)
            Return oDatosRespuesta
        Catch ex As Exception

        End Try
    End Function

    Public Function RecuperaContrasenas(ByVal sCorreo As String) As cTransaccion.cDatosUsuario Implements IwcfUsuarios.RecuperaContrasenas
        Dim oDatosRespuesta As New cTransaccion.cDatosUsuario
        Dim oCorreo As New cTransaccion.cAutenticaUsuarios

        oDatosRespuesta = oCorreo.RecuperaContrasena(sCorreo)

        Return oDatosRespuesta
    End Function

    Public Function InsertaUsuarios(ByVal oDatos As cTransaccion.cDatosUsuario) As Long Implements IwcfUsuarios.InsertaUsuarios
        Dim respuesta As Long = 0
        Dim oUsuario As New cTransaccion.cUsuarios


        respuesta = oUsuario.InsertaUsuarios(oDatos)

        Return respuesta
    End Function

    Public Function ActualizaUsuarios(oDatos As cTransaccion.cDatosUsuario) As String Implements IwcfUsuarios.ActualizaUsuarios
        Dim respuesta As String = ""
        Dim oUsuario As New cTransaccion.cUsuarios

        respuesta = oUsuario.ActualizaUsuarios(oDatos)

        Return respuesta
    End Function

    Public Function ElinimaUsuario(iIdUsuario As Integer) As Boolean Implements IwcfUsuarios.ElinimaUsuario
        Dim bRespuesta As Boolean = False
        Dim oUsuario = New cTransaccion.cUsuarios, oDatosUsuario As New cTransaccion.cDatosUsuario

        oDatosUsuario.IdUsuario = iIdUsuario

        bRespuesta = oUsuario.EliminaUsuarios(oDatosUsuario)


        Return bRespuesta
    End Function

    Public Function ConsultaListaUsuarios() As cTransaccion.cDatosUsuario() Implements IwcfUsuarios.ConsultaListaUsuarios

        Dim oUsuario As New cTransaccion.cUsuarios
        Dim arrDatos() As cTransaccion.cDatosUsuario


        arrDatos = oUsuario.ConsultaListaUsuarios()

        Return arrDatos
    End Function


End Class
