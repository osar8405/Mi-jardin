Imports System.Data.SqlClient
Public Class cConexionBD
    Private sUsuario As String
    Private sContrasena As String
    Private sCatalogo As String
    Private sServidor As String
    Private sMensajeError As String

    Dim oConexion As New SqlConnection
    Dim oCommand As New SqlCommand
    Dim oDataAdapter As New SqlDataAdapter
    Dim oDataTable As DataTable



    Public Property Usuario As String
        Get
            Return sUsuario
        End Get
        Set(value As String)
            sUsuario = value
        End Set
    End Property

    Public Property Contrasena As String
        Get
            Return sContrasena
        End Get
        Set(value As String)
            sContrasena = value
        End Set
    End Property

    Public Property Catalogo As String
        Get
            Return sCatalogo
        End Get
        Set(value As String)
            sCatalogo = value
        End Set
    End Property

    Public Property Servidor As String
        Get
            Return sServidor
        End Get
        Set(value As String)
            sServidor = value
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

    Public Sub New(ByVal sServer As String, ByVal sCatalog As String, ByVal sUser As String, ByVal sPwd As String)
        sServidor = sServer
        sCatalogo = sCatalog
        sUsuario = sUser
        sContrasena = sPwd
    End Sub
    Public Function Conectar() As Boolean
        Try
            'oConexion = New SqlConnection("data source='localhost', initial catalog='pruebas', user='sa', pwd='Mio34' ")
            'oConexion.Open()
            Return True
        Catch ex As Exception
            MensajeError = ex.Message
        End Try
        Return False
    End Function

    Public Function EjecutaExecuteNonQuery(ByVal sSentenciaSQL As String)
        Try
            Using oConexion As New SqlConnection("Data Source='" & Servidor & "'; Initial Catalog='" & sCatalogo & "';User Id='" & sUsuario & "';Password='" & sContrasena & "'")
                oConexion.Open()
                Using oCommand
                    oCommand.Connection = oConexion
                    oCommand.CommandText = sSentenciaSQL
                    oCommand.ExecuteNonQuery()
                End Using

                oConexion.Close()
            End Using

            Return True
        Catch ex As Exception
            MensajeError = ex.Message
            Return False
        End Try
    End Function

    Public Function EjecutaExecuteScalar(ByVal sSentenciaSQL As String)
        Try
            Dim sRespuesta As String
            Using oConexion As New SqlConnection("Data Source='" & Servidor & "'; Initial Catalog='" & sCatalogo & "';User Id='" & sUsuario & "';Password='" & sContrasena & "'")
                oConexion.Open()
                Using oCommand
                    oCommand.Connection = oConexion
                    oCommand.CommandText = sSentenciaSQL
                    sRespuesta = oCommand.ExecuteScalar
                End Using

                oConexion.Close()
            End Using

            Return sRespuesta
        Catch ex As Exception
            MensajeError = ex.Message
            Return False
        End Try
    End Function
    Public Function EjecutaDataAdapter(ByVal sSentenciaSQL As String)
        Try
            oConexion = New SqlConnection("Data Source='" & sServidor & "'; Initial Catalog='" & sCatalogo & "';User Id='" & sUsuario & "';Password='" & sContrasena & "'")
            oDataAdapter = New SqlDataAdapter(sSentenciaSQL, oConexion)
            oDataTable = New DataTable
            oDataAdapter.Fill(oDataTable)
            Return oDataTable
        Catch ex As Exception
            MensajeError = ex.Message
        End Try
        Return oDataTable
    End Function
End Class
