Imports System.Web.Routing
Imports System.Web.Http

Imports System.Globalization
Imports System.Threading

Public Class WebApiApplication
    Inherits System.Web.HttpApplication

    Public Sub New()

    End Sub

    Protected Sub Application_Start()
        System.Threading.Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US")
        GlobalConfiguration.Configure(AddressOf WebApiConfig.Register)
        'Cualquier carga inicial  o Configuracion de parametros






    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Sub WebApiApplication_PostAuthorizeRequest(sender As Object, e As EventArgs) Handles Me.PostAuthorizeRequest
        'System.Web.HttpContext.Current.SetSessionStateBehavior(System.Web.SessionState.SessionStateBehavior.Required)
    End Sub

    Private Sub WebApiApplication_LogRequest(sender As Object, e As EventArgs) Handles Me.LogRequest

    End Sub
End Class
