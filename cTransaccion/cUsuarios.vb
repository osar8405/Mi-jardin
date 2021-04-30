Imports System.Text

Public Class cDatosUsuario
    Private iIdUsuario As Integer
    Private sNombre As String
    Private sApellidoPaterno As String
    Private sApellidoMaterno As String
    Private sCorreo As String
    Private sUsuario As String
    Private sContrasena As String
    Private sMensajeError As String
    Private iEstatus As Integer

    Public Property IdUsuario As Integer
        Get
            Return iIdUsuario
        End Get
        Set(value As Integer)
            iIdUsuario = value
        End Set
    End Property
    Public Property Nombre As String
        Get
            Return sNombre
        End Get
        Set(value As String)
            sNombre = value
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

    Public Property Correo As String
        Get
            Return sCorreo
        End Get
        Set(value As String)
            sCorreo = value
        End Set
    End Property

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

    Public Property MensajeError As String
        Get
            Return sMensajeError
        End Get
        Set(value As String)
            sMensajeError = value
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


End Class

Public Class cUsuarios

    Dim oEjecutaConsulta As New cConexionBD.cConexionBD("localhost", "Bituaj", "sa", "Mio34")
    Dim sSentenciaSql As String
    Dim oDt = New DataTable

    Public Function ConsultaListaUsuarios() As cDatosUsuario()
        Dim oDatosRespuesta() As cDatosUsuario = Nothing
        Try

            sSentenciaSql = "Select id_usuario As 'ID', 
usuario as 'Usuario',
nombre AS 'Nombre', 
apellido_paterno AS 'Apellido_Paterno', 
apellido_materno AS 'Apellido_Materno', 
--IIF(estatus=1,'Activo', 'Inactivo') AS 'Estatus', 
estatus as 'Estatus',
correo_electronico as 'Correo_Electrónico'
FROM usuario"

            'Retorna un DT
            oDt = New DataTable
            oDt = oEjecutaConsulta.EjecutaDataAdapter(sSentenciaSql)

            If oDt.Rows.Count > 0 Then

                Dim i As Integer = 0
                For Each Row In oDt.Rows
                    ReDim Preserve oDatosRespuesta(i)
                    Dim oDatosUsuario = New cDatosUsuario
                    oDatosUsuario.IdUsuario = Row("ID")
                    oDatosUsuario.Usuario = Row("Usuario")
                    oDatosUsuario.Nombre = Row("Nombre")
                    oDatosUsuario.ApellidoPaterno = Row("Apellido_Paterno")
                    oDatosUsuario.ApellidoMaterno = Row("Apellido_Materno")
                    oDatosUsuario.Estatus = Row("Estatus")
                    oDatosUsuario.Correo = Row("Correo_Electrónico")


                    oDatosRespuesta(i) = oDatosUsuario
                    oDatosUsuario = Nothing

                    i += 1
                Next

            Else
                Throw New Exception(oEjecutaConsulta.MensajeError)
            End If

        Catch ex As Exception

        End Try
        Return oDatosRespuesta
    End Function

    Public Function InsertaUsuarios(oDatosUsuario As cDatosUsuario) As Long
        Dim oRes As Long = 0
        Try
            Dim oSeguridadContrasenas As New cSeguridadContrasenas
            Dim Contrasenia As String = ""

            Contrasenia = oSeguridadContrasenas.Encriptar(oDatosUsuario.Contrasena)
            sSentenciaSql = "INSERT INTO usuario (nombre, apellido_paterno, apellido_materno, correo_electronico, usuario, contrasena) 
values ('" & oDatosUsuario.Nombre & "', '" & oDatosUsuario.ApellidoPaterno & "', '" & oDatosUsuario.ApellidoMaterno & "', '" & oDatosUsuario.Correo & "', '" & oDatosUsuario.Usuario & "', '" & Contrasenia & "' )"

            If (oEjecutaConsulta.EjecutaExecuteNonQuery(sSentenciaSql)) = True Then
                oRes = 1
            Else
                Return oRes
            End If
        Catch ex As Exception
            oDatosUsuario.MensajeError = ex.Message
            Return oRes
        End Try
        Return oRes
    End Function

    Public Function ActualizaUsuarios(oDatosUsuario As cDatosUsuario) As String
        Dim sRespuesta As String = ""
        Try
            sSentenciaSql = "SELECT count(*) from usuario where (correo_electronico='" & oDatosUsuario.Correo & "' and id_usuario<>" & oDatosUsuario.IdUsuario & ") and correo_electronico<>''"
            If (oEjecutaConsulta.EjecutaExecuteScalar(sSentenciaSql)) > 0 Then
                sRespuesta = "Este correo electrónico ya está registrado en el sistema."
            Else

                sSentenciaSql = "Update usuario set nombre='" & oDatosUsuario.Nombre & "', apellido_paterno='" & oDatosUsuario.ApellidoPaterno & "', apellido_materno='" & oDatosUsuario.ApellidoMaterno & "', correo_electronico='" & oDatosUsuario.Correo & "', estatus=" & oDatosUsuario.Estatus & ", usuario='" & oDatosUsuario.Usuario & "' where id_usuario=" & oDatosUsuario.IdUsuario & ""
                If (oEjecutaConsulta.EjecutaExecuteNonQuery(sSentenciaSql)) = True Then
                    sRespuesta = "Ok"
                Else
                    oDatosUsuario.MensajeError = oDatosUsuario.MensajeError
                    sRespuesta = "Error al actualizar Usuario"
                End If
            End If
        Catch ex As Exception
            sRespuesta = ex.Message
        End Try
        Return sRespuesta
    End Function

    Public Function EliminaUsuarios(oDatosUsuario As cDatosUsuario) As Boolean
        Dim bRespuesta As Boolean = False
        Try
            sSentenciaSql = "DELETE FROM usuario where id_usuario=" & oDatosUsuario.IdUsuario & ""

            If (oEjecutaConsulta.EjecutaExecuteNonQuery(sSentenciaSql)) = True Then
                bRespuesta = True
            Else
                bRespuesta = False
            End If

        Catch ex As Exception
            oDatosUsuario.MensajeError = ex.Message
            bRespuesta = False
        End Try
        Return bRespuesta
    End Function
End Class

Public Class cSeguridadContrasenas
    'función que decodifica el dato  
    Function Desencriptar(ByVal DataValue As Object) As Object

        Dim x As Long
        Dim Temp As String
        Dim HexByte As String

        For x = 1 To Len(DataValue) Step 2

            HexByte = Mid(DataValue, x, 2)
            Temp = Temp & Chr(ConvToInt(HexByte))

        Next x
        ' retorno  
        Desencriptar = Temp

    End Function

    Private Function ConvToInt(ByVal x As String) As Integer

        Dim x1 As String
        Dim x2 As String
        Dim Temp As Integer

        x1 = Mid(x, 1, 1)
        x2 = Mid(x, 2, 1)

        If IsNumeric(x1) Then
            Temp = 16 * Int(x1)
        Else
            Temp = (Asc(x1) - 55) * 16
        End If

        If IsNumeric(x2) Then
            Temp = Temp + Int(x2)
        Else
            Temp = Temp + (Asc(x2) - 55)
        End If

        ' retorno  
        ConvToInt = Temp

    End Function

    Private Function ConvToHex(ByVal x As Integer) As String
        If x > 9 Then
            ConvToHex = Chr(x + 55)
        Else
            ConvToHex = CStr(x)
        End If
    End Function

    ' función que codifica el dato  
    Function Encriptar(ByVal DataValue As Object) As Object

        Dim x As Long
        Dim Temp As String
        Dim TempNum As Integer
        Dim TempChar As String
        Dim TempChar2 As String

        For x = 1 To Len(DataValue)
            TempChar2 = Mid(DataValue, x, 1)
            TempNum = Int(Asc(TempChar2) / 16)

            If ((TempNum * 16) < Asc(TempChar2)) Then

                TempChar = ConvToHex(Asc(TempChar2) - (TempNum * 16))
                Temp = Temp & ConvToHex(TempNum) & TempChar
            Else
                Temp = Temp & ConvToHex(TempNum) & "0"

            End If
        Next x
        Encriptar = Temp
    End Function

End Class

Public Class cAutenticaUsuarios
    Private sUsuario As String
    Private sContrasena As String
    Private sMensajeError As String

    Dim oDa As New SqlClient.SqlDataAdapter
    Dim oDt As DataTable
    Dim sSentenciaSql As String = ""
    Dim oEjecutaConsulta As New cConexionBD.cConexionBD("localhost", "Bituaj", "sa", "Mio34")
    Dim oSeguridadContrasenas As New cSeguridadContrasenas

    Dim oDatosCorreo As New cEnviaCorreo.cEnviaCorreo

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

    Public Property MensajeError As String
        Get
            Return sMensajeError
        End Get
        Set(value As String)
            sMensajeError = value
        End Set
    End Property

    Public Function ValidaCredencial(oDatosUsuario As cDatosUsuario) As cDatosUsuario
        Dim oDatosRespuesta As New cDatosUsuario
        Dim Contrasena As String = ""
        Try

            sSentenciaSql = "SELECT Id_usuario as 'Id Usuario',
nombre as 'Nombre',
apellido_paterno as 'Apellido Paterno',
apellido_materno as 'Apellido Materno',
estatus as 'Estatus',
contrasena as 'Contraseña',
correo_electronico as 'Email'
FROM usuario 
where usuario='" & oDatosUsuario.Usuario & "'"
            oDt = New DataTable
            oDt = oEjecutaConsulta.EjecutaDataAdapter(sSentenciaSql)

            If oDt.Rows.Count > 0 Then
                For Each Row In oDt.Rows
                    Contrasena = Row("Contraseña").ToString

                    If oSeguridadContrasenas.Desencriptar(Contrasena) = oDatosUsuario.Contrasena Then
                        If Row("Estatus").ToString > 0 Then
                            'oDatosUsuario.IdUsuario = Row("Id Usuario")
                            'oDatosUsuario.Nombre = Row("Nombre").ToString
                            'oDatosUsuario.ApellidoPaterno = Row("Apellido Paterno").ToString
                            'oDatosUsuario.ApellidoMaterno = Row("Apellido Materno").ToString
                            'oDatosUsuario.Estatus = Row("Estatus").ToString
                            'oDatosUsuario.Correo = Row("Email").ToString

                            oDatosRespuesta.IdUsuario = Row("Id Usuario")
                            oDatosRespuesta.Nombre = Row("Nombre")
                            oDatosRespuesta.ApellidoPaterno = Row("Apellido Paterno")
                            oDatosRespuesta.ApellidoMaterno = Row("Apellido Materno")
                            oDatosRespuesta.Estatus = Row("Estatus")
                            oDatosRespuesta.Correo = Row("Email")
                            oDatosRespuesta.MensajeError = ""

                        Else
                            oDatosRespuesta.MensajeError = "Usuario inactivo"
                        End If
                    Else
                        oDatosRespuesta.MensajeError = "Contraseña no valida, favor de intentar de nuevo"
                    End If
                Next
            Else
                oDatosRespuesta.MensajeError = "Usuario no valido, favor de verificar"
            End If

        Catch ex As Exception
            oDatosRespuesta.MensajeError = ex.Message
        Finally

        End Try
        Return oDatosRespuesta
    End Function

    Public Function RecuperaContrasena(ByVal sCorreo As String) As cDatosUsuario
        Dim oDatosUsuario As New cDatosUsuario, oDatosRespuesta As New cDatosUsuario
        Try


            sSentenciaSql = "SELECT nombre as 'Nombre', apellido_paterno as 'Apellido Paterno', apellido_materno as 'Apellido Materno', usuario as 'Usuario', contrasena as 'Contraseña'  from usuario where correo_electronico='" & sCorreo & "'"
            oDt = New DataTable
            oDt = oEjecutaConsulta.EjecutaDataAdapter(sSentenciaSql)

            If oDt.Rows.Count > 0 Then
                For Each Row In oDt.Rows

                    oDatosUsuario.Nombre = Row("Nombre").ToString
                    oDatosUsuario.Usuario = Row("Usuario").ToString
                    oDatosUsuario.Contrasena = Row("Contraseña").ToString
                    oDatosUsuario.Contrasena = oSeguridadContrasenas.Desencriptar(Row("Contraseña").ToString)

                    oDatosRespuesta.Nombre = Row("Nombre").ToString
                    oDatosRespuesta.Usuario = Row("Usuario").ToString


                Next

                oDatosCorreo.CorreoDe = "oscar.mireles@grupobituaj.com.mx"
                oDatosCorreo.CorreoDeContraseña = "Sow58517"
                oDatosCorreo.CorreoPara = oDatosUsuario.Correo
                oDatosCorreo.CorreoAlias = "Recuperación de contraseña"
                oDatosCorreo.CorreoAsunto = "Recuperación de contraseña"

                Dim strB As New StringBuilder


                strB.AppendLine("<p><span style='color:#003366'>Hola, " & oDatosUsuario.Nombre & "</span></p>")
                strB.AppendLine("<p><span style='color:#003366'>Puedes iniciar sesión en el sistema con los siguientes datos; </span></p>")
                strB.AppendLine("<span style='color:#003366'>Usuario: " & oDatosUsuario.Usuario & "</span> <br/>")
                strB.AppendLine("<span style='color:#003366'>Contraseña: " & oDatosUsuario.Contrasena & "</span>")



                oDatosCorreo.StrB = strB
                oDatosCorreo.SmtpPort = "587"
                oDatosCorreo.SmtpHost = "SMTP.Office365.com"

                oDatosCorreo.Envia(oDatosCorreo)

            Else
                oDatosRespuesta.MensajeError = "No existe ese correo asociado a un usuario de nuestro sistema. Por favor, introduzca un correo válido."

            End If
        Catch ex As Exception
            oDatosRespuesta.MensajeError = ex.Message
        End Try
        Return oDatosRespuesta
    End Function
End Class


