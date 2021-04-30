Imports System.ServiceModel

' NOTE: You can use the "Rename" command on the context menu to change the interface name "IwcfUsuarios" in both code and config file together.
<ServiceContract()>
Public Interface IwcfUsuarios

    <OperationContract>
    Function AutenticaUsuarios(ByVal oDatos As cTransaccion.cDatosUsuario) As cTransaccion.cDatosUsuario

    '<OperationContract>
    'Function RecuperaContrasenas(ByVal sCorreo As String)
    <OperationContract>
    Function RecuperaContrasenas(ByVal sCorreo As String) As cTransaccion.cDatosUsuario

    <OperationContract()>
    Function InsertaUsuarios(ByVal oDatos As cTransaccion.cDatosUsuario) As Long

    <OperationContract>
    Function ActualizaUsuarios(ByVal oDatos As cTransaccion.cDatosUsuario) As String

    <OperationContract>
    Function ElinimaUsuario(ByVal iIdUsuario As Integer) As Boolean

    <OperationContract>
    Function ConsultaListaUsuarios() As cTransaccion.cDatosUsuario()

End Interface

