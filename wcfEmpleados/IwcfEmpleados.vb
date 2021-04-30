Imports System.ServiceModel

' NOTE: You can use the "Rename" command on the context menu to change the interface name "IwcfEmpleados" in both code and config file together.
<ServiceContract()>
Public Interface IwcfEmpleados

    <OperationContract>
    Function InsertaEmpleados(ByVal oDatosEmpleado As cTransaccion.cDatosEmpleado, ByVal iIdUsuario As Integer) As cTransaccion.cDatosEmpleado

    <OperationContract>
    Function Actualiza_Empleados(oDatosEmpleado As cTransaccion.cDatosEmpleado, iIdUsuario As Integer) As cTransaccion.cDatosEmpleado

    <OperationContract>
    Function EliminaEmpleados(ByVal iIdEmpleado As Integer) As Boolean

    <OperationContract>
    Function LlenaComboDepartamentos() As cTransaccion.cDatosDepartamento()

    <OperationContract>
    Function ConsultaListaEmpleados() As cTransaccion.cDatosEmpleado()

End Interface