Imports System.Net
Imports System.Web.Http
Imports System.Web.Http.Cors

Imports System.IO


Namespace Controllers
    <EnableCors("*", "*", "*")>
    Public Class LoginController
        Inherits ApiController
        Public Function PostValue(<FromBody()> ByVal value As LogginParameter) As Usuario





            Dim usuario As String, password As String
            usuario = value.usuario

            password = value.password
            Dim user As New Usuario







            Try


                user.mensaje = "Usuario no esta registrado"

                Dim pwd As String

                For Each UsRow As DataRow In Ejecutar("LSp_lab_user '" & usuario & "'")(0).Rows
                    pwd = UsRow!clave.ToString
                    user.mensaje = ""


                    If UCase(password) <> UCase(pwd) Then
                        user.mensaje = "Password Incorrecto"

                        Exit Try
                    End If
                    If UsRow!estatus.ToString <> "A" Then
                        user.mensaje = "Usuario Inactivo"

                        Exit Try
                    End If



                    'Creo la sesion en la base de datos y genero una llave
                    user.mensaje = ""
                    With MySWebtoredProcedure("LSp_web_USersesion")
                        .Parameters(1).SqlValue = usuario
                        .ExecuteNonQuery()
                        user.token = .Parameters(2).Value
                    End With
                    user.valido = "S"




                Next




            Catch ex As Exception
                user.valido = "N"
                user.mensaje = "[Login Error]" & ex.Message
            End Try
            Return user
        End Function
    End Class

    <EnableCors("*", "*", "*")>
    Public Class SetHostController
        Inherits ApiController

        ' GET: api/Menu
        Public Function PostValues(<FromBody()> arguments As SessionIDParameter) As ReturnInfo

            Dim Ret As New ReturnInfo
            With Ret
                .MensajeError = ""
                .GeneroError = True
                .Data = ""
            End With

            Try
                Dim Usuario As String = sesionactiva(arguments.sessionid).codigousuario
                If Usuario = "." Then
                    Ret.SessionClosed = True
                    Return Ret
                End If


                Ejecutar("LSp_web_SetHost '" & arguments.sessionid & "','" & Usuario & "'")

                Ret.GeneroError = False
                Ret.Data = ""




            Catch ex As Exception
                Ret.GeneroError = True
                Ret.MensajeError = "[Set  Host ]" & ex.Message
            End Try
            Return Ret
        End Function

    End Class


    <EnableCors("*", "*", "*")>
    Public Class LoggoutController
        Inherits ApiController

        ' GET: api/Menu
        Public Function PostValues(<FromBody()> arguments As SessionIDParameter) As ReturnInfo

            Dim Ret As New ReturnInfo
            With Ret
                .MensajeError = ""
                .GeneroError = True
                .Data = ""
            End With
            Try



                Ejecutar("LSp_web_borrrarsesion '" & arguments.sessionid & "'")
                Ret.GeneroError = False
                Ret.Data = ""


            Catch ex As Exception
                Ret.GeneroError = True
                Ret.MensajeError = ex.Message
            End Try
            Return Ret
        End Function

    End Class


End Namespace


Public Class LogginParameter
    Public usuario As String
    Public password As String

End Class

Public Class SessionIDParameter
    Public sessionid As String
End Class
