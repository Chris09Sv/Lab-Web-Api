Imports System.Net
Imports System.Data.SqlClient
Imports System.Web.Http
Imports System.Drawing

Imports System.Web.Http.Cors


Imports System.Net.Mail
Imports System.IO
Namespace Controllers


#Region "Importar Datos"


    <EnableCors("*", "*", "*")> Public Class ResultadosController
        Inherits ApiController


        Public Function PostValues(<FromBody()> arguments As ResultadosArguments) As Object

            Dim Ret As New ReturnInfo
            With Ret
                Try
                    .MensajeError = ""
                Catch ex As Exception

                End Try

                .GeneroError = True
                .Data = " "

            End With




            Try



                Dim Usuario As String = sesionactiva(arguments.token).codigousuario
                If Usuario = "." Then
                    Ret.SessionClosed = True
                    Exit Try
                End If
                Dim Tabla_Errores As New DataTable

                With Tabla_Errores
                    .Columns.Add("linea", System.Type.GetType("System.Int32"))
                    .Columns.Add("msgerror")

                End With
                Dim Rowerror As DataRow
                Dim CCount As Long = 0, ErrorCount As Long = 0, ValidosCount As Long = 0
                Dim ErrorRetornado As String = ""



                With MySWebtoredProcedure("Sp_Registro_resultados")

                    For Each rec In arguments.datos
                        CCount = CCount + 1

                        .Parameters(1).SqlValue = Usuario
                        .Parameters(2).SqlValue = rec!codigo_muestra
                        .Parameters(3).SqlValue = rec!nombre_paciente
                        .Parameters(4).SqlValue = rec!apellido_paciente
                        .Parameters(5).SqlValue = rec!sexo
                        .Parameters(6).SqlValue = rec!fecha_nacimiento
                        .Parameters(7).SqlValue = rec!tipo_documento
                        .Parameters(8).SqlValue = rec!numero_documento
                        .Parameters(9).SqlValue = rec!codigo_nacionalidad
                        .Parameters(10).SqlValue = rec!codigo_tipo_atencion
                        .Parameters(11).SqlValue = rec!codigo_provincia
                        .Parameters(12).SqlValue = rec!codigo_municipio
                        .Parameters(13).SqlValue = rec!codigo_sector
                        .Parameters(14).SqlValue = rec!direccion
                        .Parameters(15).SqlValue = rec!telefono
                        .Parameters(16).SqlValue = rec!celular
                        .Parameters(17).SqlValue = rec!signos_sintomas
                        .Parameters(18).SqlValue = rec!fecha_inicio_sintomas
                        .Parameters(19).SqlValue = rec!embarazada
                        .Parameters(20).SqlValue = rec!fecha_toma_muestra
                        .Parameters(21).SqlValue = rec!fecha_reporte
                        .Parameters(22).SqlValue = rec!resultado
                        .Parameters(23).SqlValue = rec!tipo_muestra
                        .Parameters(24).SqlValue = rec!tipo_prueba
                        .Parameters(25).SqlValue = rec!agente
                        .Parameters(26).SqlValue = rec!subtipo
                        .Parameters(27).SqlValue = rec!trab_salud
                        .Parameters(28).SqlValue = rec!motivo_prueba
                        .Parameters(29).SqlValue = rec!codigo_pais_procedencia
                        .Parameters(30).SqlValue = rec!primera_muestra
                        .Parameters(31).SqlValue = rec!comentarios

                        .ExecuteNonQuery()

                        ErrorRetornado = .Parameters(32).Value
                        'Si el registro ha retornado un error
                        If ErrorRetornado <> "" Then
                            ErrorCount += 1
                            Rowerror = Tabla_Errores.NewRow
                            With Rowerror
                                !linea = CCount
                                !msgerror = ErrorRetornado
                            End With
                            Tabla_Errores.Rows.Add(Rowerror)
                        Else
                            ValidosCount += 1
                        End If

                    Next
                End With
                Ret.RegistrosInvalidos = ErrorCount
                Ret.RegistrosValidos = ValidosCount
                Ret.RegistrosCargados = CCount
                Ret.ItemsError = serializatabla(Tabla_Errores)
                Ret.GeneroError = False
                Ret.Data = ""




            Catch ex As Exception
                Ret.GeneroError = True
                Ret.MensajeError = "Carga de Resultados " & ex.Message
            End Try
            Return Ret
        End Function

    End Class


#End Region


#Region "Catalogos"


    <EnableCors("*", "*", "*")> Public Class CatalogosController
        Inherits ApiController


        Public Function PostValues(<FromBody()> arguments As CatalogoArguments) As Object

            Dim Ret As New ReturnCatalog
            With Ret
                Try
                    .MensajeError = ""
                Catch ex As Exception

                End Try

                .GeneroError = True
                .Data = " "

            End With




            Try



                Dim Usuario As String = sesionactiva(arguments.token).codigousuario
                If Usuario = "." Then
                    Ret.SessionClosed = True
                    Exit Try
                End If

                With Ejecutar("Sp_catalogos '" & arguments.tipo & "'")

                    Ret.GeneroError = False
                    Ret.Data = serializatabla(.Item(0))
                End With


            Catch ex As Exception
                Ret.GeneroError = True
                Ret.MensajeError = "Consulta de Catalogos " & ex.Message
            End Try
            Return Ret
        End Function

    End Class

#End Region

#Region "Consultas"


    <EnableCors("*", "*", "*")> Public Class ConsultaController
        Inherits ApiController


        Public Function PostValues(<FromBody()> arguments As ParametrosArguments) As Object

            Dim Ret As New ReturnCatalog
            With Ret
                Try
                    .MensajeError = ""
                Catch ex As Exception

                End Try

                .GeneroError = True
                .Data = " "

            End With




            Try



                Dim Usuario As String = sesionactiva(arguments.token).codigousuario
                If Usuario = "." Then
                    Ret.SessionClosed = True
                    Exit Try
                End If

                With Ejecutar("spGetResultados '" & arguments.FechaI & "'," & "'" & arguments.FechaF & "'")

                    Ret.GeneroError = False
                    Ret.Data = serializatabla(.Item(0))
                    'Ret.Data = .Item(0) 'serializatabla(.item(0))

                End With


            Catch ex As Exception
                Ret.GeneroError = True
                Ret.MensajeError = "Consulta de resultados " & ex.Message
            End Try
            Return Ret
        End Function

    End Class

#End Region
End Namespace
