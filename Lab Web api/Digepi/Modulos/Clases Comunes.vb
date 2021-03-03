
Public Class ReturnInfo
    Public GeneroError As Boolean = True
    Public SessionClosed As Boolean = False
    Public MensajeError As String = ""
    Public ItemsError As Object
    Public Data As Object
    Public RegistrosValidos As Long = 0
    Public RegistrosInvalidos As Long = 0
    Public RegistrosCargados As Long = 0

End Class

Public Class ReturnCatalog
    Public GeneroError As Boolean = True
    Public SessionClosed As Boolean = False
    Public MensajeError As String = ""

    Public Data As Object


End Class

Public Class ResultadosArguments
    Public token As String
    Public datos As Object

End Class

Public Class CatalogoArguments
    Public token As String
    Public tipo As Object

End Class

Public Class ParametrosArguments
    Public token As String
    Public FechaI As String
    Public FechaF As String

End Class





Public Class Usuario


    Public token As String = ""
    Public valido As String = "N"
    Public mensaje As String = "Usuario no existe"


End Class
Public Class sesion
    Public sesionid As String
    Public codigousuario As String
    Public nombre As String
    Public fecha As String

    Public Ultimo_Acceso As String = ""
End Class
