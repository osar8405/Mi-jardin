' NOTE: You can use the "Rename" command on the context menu to change the class name "wcfEmpleados" in code, svc and config file together.
' NOTE: In order to launch WCF Test Client for testing this service, please select wcfEmpleados.svc or wcfEmpleados.svc.vb at the Solution Explorer and start debugging.
Imports System.Data.SqlClient
Imports cTransaccion
Imports wcfEmpleados

Public Class wcfEmpleados
    Implements IwcfEmpleados

    Public Function InsertaEmpleados(oDatosEmpleado As cTransaccion.cDatosEmpleado, iIdUsuario As Integer) As cTransaccion.cDatosEmpleado Implements IwcfEmpleados.InsertaEmpleados
        Dim oDatosRespuesta As cTransaccion.cDatosEmpleado
        Dim oEmpleado As New cTransaccion.cEmpleados

        oDatosRespuesta = oEmpleado.InsertaEmpleados(oDatosEmpleado, iIdUsuario)

        Return oDatosRespuesta
    End Function


    Public Function Actualiza_Empleados(oDatosEmpleado As cTransaccion.cDatosEmpleado, iIdUsuario As Integer) As cTransaccion.cDatosEmpleado Implements IwcfEmpleados.Actualiza_Empleados
        Dim oDatosRespuesta As cTransaccion.cDatosEmpleado
        Dim oEmpleado As New cTransaccion.cEmpleados

        oDatosRespuesta = oEmpleado.ActualizaEmpleados(oDatosEmpleado, iIdUsuario)

        Return oDatosRespuesta
    End Function

    Public Function EliminaEmpleados(iIdEmpleado As Integer) As Boolean Implements IwcfEmpleados.EliminaEmpleados

        Dim bRespuesta As Boolean = False
            Dim oDatosEmpleado As New cTransaccion.cDatosEmpleado, oEmpleado As New cTransaccion.cEmpleados

            oDatosEmpleado.IdEmpleado = iIdEmpleado

            bRespuesta = oEmpleado.EliminaEmpleados(oDatosEmpleado)

            Return bRespuesta

    End Function

    Public Function LlenaComboDepartamentos() As cTransaccion.cDatosDepartamento() Implements IwcfEmpleados.LlenaComboDepartamentos
        Dim oDatosRespuesta() As cTransaccion.cDatosDepartamento = Nothing
        Dim oDepartamento As New cTransaccion.cDepartamentos

        oDatosRespuesta = oDepartamento.ObtieneListaDepartamentos

        Return oDatosRespuesta
    End Function

    Public Function ConsultaListaEmpleados() As cTransaccion.cDatosEmpleado() Implements IwcfEmpleados.ConsultaListaEmpleados
        Dim arrDatosRespuesta() As cTransaccion.cDatosEmpleado
        Dim oEmpleado As New cTransaccion.cEmpleados

        arrDatosRespuesta = oEmpleado.ConsultaListaEmpleados

        Return arrDatosRespuesta
    End Function

End Class
