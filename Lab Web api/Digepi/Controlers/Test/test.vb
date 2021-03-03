Imports System.Net
Imports System.Web.Http
Imports System.Web.Script.Serialization
Imports System.IO
Imports System.Web.Http.Cors
Imports System.Xml

Namespace Pruebas
    <EnableCors("*", "*", "*")> Public Class testController
        Inherits ApiController

        ' GET: api/vitico
        Public Function GetValues() As ReturnInfo


            Dim Ret As New ReturnInfo
            Try

                With Ret
                    .MensajeError = ""
                    .GeneroError = True
                    .Data = ""

                End With

                Ret.Data = "Servicion de notificacion de Laboratorios" & Now.ToString
                Ret.GeneroError = False

            Catch ex As Exception
                Ret.GeneroError = True
                Ret.MensajeError = ex.Message
            End Try
            Return Ret
        End Function

        ' GET: api/vitico/5
        Public Function GetValue(ByVal id As Integer) As String
            Return "value"
        End Function

        ' POST: api/vitico
        Public Sub PostValue(<FromBody()> ByVal value As String)

        End Sub

        ' PUT: api/vitico/5
        Public Sub PutValue(ByVal id As Integer, <FromBody()> ByVal value As String)

        End Sub

        ' DELETE: api/vitico/5
        Public Sub DeleteValue(ByVal id As Integer)

        End Sub
    End Class


End Namespace

