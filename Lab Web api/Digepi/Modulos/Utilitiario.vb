
Imports System.Data
Imports System.Data.SqlClient
Imports System.Text
Imports System.Security.Cryptography
Public Class Utilitiario
    Public WebSqlConectionString As String

    ''' <summary>
    ''' Utilizado para traer la estructura de los stored procedures.
    ''' de esta manera para hacer la llamada no hay que crear el parametro.
    ''' solamente enviar los datos.
    ''' 
    ''' </summary>
    ''' <param name="StoredProcName"></param>
    ''' <returns></returns>
    Public Function MyStoredProcedure(ByVal StoredProcName As String) As SqlCommand
        Dim Cmd As New SqlCommand
        Dim Cnn As New SqlConnection(SQL_CONNECTION_STRING)

        Dim cmSQL As New SqlCommand
        Dim drSQL As SqlDataReader



        Dim Dbname As String = ""
        Dim Pos As Long
        Pos = InStr(StoredProcName, "dbo.") + 3
        If Pos > 3 Then
            Dbname = Mid(StoredProcName, 1, Pos)
            StoredProcName = Mid(StoredProcName, Pos + 1, 1000)
        End If

        'Dbname= mid(1,StoredProcName.Subs 

        Dim strSQL As String
        Try
            Cnn.Open()
            strSQL = Dbname & "LSp_ColumnasStoredProcedure '" & StoredProcName & "'"
            cmSQL = New SqlCommand(strSQL, Cnn)
            Dim MsqlDataAdapter As New SqlDataAdapter(cmSQL)
            With cmSQL



                drSQL = cmSQL.ExecuteReader()






                .Connection = Cnn
                .CommandText = Dbname & StoredProcName
                .CommandType = CommandType.StoredProcedure

                .CommandTimeout = 200

                With .Parameters.Add("return", SqlDbType.Int)
                    .Precision = 10
                    .Size = 0
                    .Scale = 0
                    .Direction = ParameterDirection.ReturnValue

                End With

                Do While drSQL.Read()

                    If (drSQL.Item("nombre_campo").ToString()) <> "n/a" Then


                        .Parameters.Add(drSQL.Item("nombre_campo").ToString(), SqlDbType.Variant)

                        With .Parameters(drSQL.Item("nombre_campo").ToString())
                            '.DbType = DbType.Int32
                            '.ParameterName = drSQL.Item("nombre_campo").ToString()

                            If drSQL.Item("salida") Then
                                .Direction = ParameterDirection.InputOutput
                            Else
                                .Direction = ParameterDirection.Input
                            End If




                            Select Case drSQL.Item("tipo_dato").ToString()

                                Case "VARCHAR", "CHAR"
                                    '.DbType = DbType.String
                                    .SqlDbType = SqlDbType.VarChar

                                    .Size = drSQL.Item("longitud")

                                Case "IMAGE"
                                    .SqlDbType = SqlDbType.Image


                                Case "NTEXT"
                                    '.DbType = DbType.StringFixedLength
                                    .SqlDbType = SqlDbType.NText

                                    .Size = 1073741823

                                Case "DATETIME"
                                    .DbType = DbType.DateTime


                                Case "SYSNAME"
                                    .SqlDbType = SqlDbType.Variant

                                    'lo asumo como numerico

                                Case Else
                                    .DbType = DbType.Double
                                    .SqlDbType = SqlDbType.Decimal
                                    .Size = 18
                                    .Precision = 19 ' drSQL.Item("precision")
                                    .Scale = drSQL.Item("scale")




                            End Select

                        End With
                    End If
                    'drSQL.Item("nombres").ToString()
                    'MsgBox(drSQL.Item("nombre_campo").ToStringGGGGGGG
                Loop
                drSQL.Close()
            End With
        Catch ex As SqlException
            ' MsgBox(ex.Message, MsgBoxStyle.Critical, "Error de Sql Server.")
        Catch ex As Exception
            'MsgBox(ex.Message, MsgBoxStyle.Critical, "Ha Ocurrido un Error Inesperado(005 car)")

        End Try



        MyStoredProcedure = cmSQL

    End Function


    Public Function Ejecutar(ByVal strSQL As String, Optional DisplayError As Boolean = True, Optional tablename As String = "empleado") As DataTableCollection

        Dim Cnn As New SqlConnection(SQL_CONNECTION_STRING)
        Using Cnn

            Try
                Cnn.Open()

                'Creo el comando   
                Dim SqlCmd As New SqlCommand(strSQL, Cnn)
                Dim SqlCmd_Fecha_Proceso As New SqlCommand("Select convert(varchar(10), ", Cnn)

                SqlCmd.CommandTimeout = 3600

                'El Sql Adaptador, utilizar el objeto Sql, para Llenar el Dataset
                Dim MsqlDataAdapter As New SqlDataAdapter(SqlCmd)

                'MsqlDataAdapter.UpdateCommand  
                'Verificar mas luego el uso de este
                Dim SqlCmdBuilder As New SqlCommandBuilder(MsqlDataAdapter)




                Try
                    'Verifico si tengo activo el control de fecha de proceso

                Catch ex As Exception

                End Try

                Dim _DataSetFormulario As New DataSet()
                MsqlDataAdapter.Fill(_DataSetFormulario, tablename)

                Cnn.Close()
                Cnn.Dispose()



                '- destruyo los objetos
                SqlCmd.Dispose()
                SqlCmd = Nothing


                'El Sql Adaptador, utilizar el objeto Sql, para Llenar el Dataset
                MsqlDataAdapter.Dispose()
                MsqlDataAdapter = Nothing
                SqlCmdBuilder.Dispose()
                SqlCmdBuilder = Nothing
                SqlConnection.ClearAllPools()
                Return _DataSetFormulario.Tables
                Exit Function

            Catch ex As Exception


            End Try
            Return Nothing


        End Using
    End Function
    ''' <summary>
    ''' Utilizado para sabser que el token esta activo
    ''' </summary>
    ''' <param name="sessionid"></param>
    ''' <returns></returns>
    Public Function sesionactiva(sessionid As String) As sesion
        Dim mysess As New sesion
        mysess.codigousuario = "."
        For Each UsRow As DataRow In Ejecutar("LSp_web_versesion '" & sessionid & "'")(0).Rows
            With mysess
                .sesionid = sessionid
                .codigousuario = UsRow!user_name
                .nombre = UsRow!nombre
                .fecha = FormatearFecha(UsRow!fecha_ultima_entrada, FormatosFechas.Lsf_YYYY_MM_dd_h_m_s_sql120)


            End With

        Next


        Return mysess


    End Function

End Class
