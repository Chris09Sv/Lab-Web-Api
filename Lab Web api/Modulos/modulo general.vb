Imports System.Data.SqlClient
Imports System.Net.Mail
Imports System.IO
Public Module ModuloGeneral


    Public SQL_CONNECTION_STRING As String = My.Settings.SqlConnection

    Public Structure Usr
        Dim ID_paciente As String
        Dim Email As String
        Public Name As String

    End Structure

    Function SetFecha112(Fecha As String) As String
        Dim Partes = Fecha.Split("-")

        Dim Dia As String = Partes(0)
        Dim Mes As String = Partes(1)
        Dim Ano As String = Partes(2)


        Return Ano & Mes & Dia

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

    Public Function MySWebtoredProcedure(ByVal StoredProcName As String) As SqlCommand


        Dim Cmd As New SqlCommand
        Dim Cnn As New SqlConnection(SQL_CONNECTION_STRING)


        Dim cmSQL As New SqlCommand
        Dim drSQL As SqlDataReader
        'Dim scmd As New SqlCommand("GetProducts", scnnNorthwind)
        'Dim sda As New SqlDataAdapter(scmd)
        'Dim dsProducts As New DataSet()
        Try

            ' dsProducts.Tables.
            Dim Dbname As String = ""
            Dim Pos As Long
            Pos = InStr(StoredProcName, "dbo.") + 3
            If Pos > 3 Then
                Dbname = Mid(StoredProcName, 1, Pos)
                StoredProcName = Mid(StoredProcName, Pos + 1, 1000)
            End If

            'Dbname= mid(1,StoredProcName.Subs 

            Dim strSQL As String

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
                '.Parameters.Add("codigo", SqlDbType.Decimal)
                '.Parameters.Add("nombre", SqlDbType.Decimal)



                'Reccorro todos los registros
                'Dim PrmNumber As Integer
                Do While drSQL.Read()
                    'PrmNumber = +1
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
            'Catch ex As SqlException
            '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Error de Sql Server.")
            'Catch ex As Exception
            '    MsgBox(ex.Message, MsgBoxStyle.Critical, "Ha Ocurrido un Error Inesperado(005 car)")

            'End Try

        Catch ex As Exception

        End Try


        MySWebtoredProcedure = cmSQL

    End Function


    Public Sub SendMail(ByVal recepientEmail As String, ByVal subject As String, ByVal body As String)

        'Dim correo As New System.Net.Mail.MailMessage
        'correo.From = New System.Net.Mail.MailAddress(My.Settings.mail_UserName)
        'correo.To.Add(New MailAddress(recepientEmail))
        'correo.Subject = subject
        'correo.Body = body
        'correo.IsBodyHtml = True
        'correo.Priority = System.Net.Mail.MailPriority.Normal
        'Dim smtp As New System.Net.Mail.SmtpClient
        'smtp.Host = My.Settings.mail_host
        'smtp.Port = Integer.Parse(My.Settings.mail_Port)
        'smtp.EnableSsl = My.Settings.mail_EnableSsl = "S"
        'smtp.Credentials = New System.Net.NetworkCredential(My.Settings.mail_UserName, My.Settings.mail_Password)
        'smtp.Send(correo)



    End Sub
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

    Public Function serializatabla(Tbl As DataTable) As Object
        Dim serializer As New System.Web.Script.Serialization.JavaScriptSerializer()
        serializer.MaxJsonLength = Int32.MaxValue
        Dim packet As New List(Of Dictionary(Of String, Object))()

        Dim row As Dictionary(Of String, Object) = Nothing

        Tbl.AsEnumerable.ToList


        For Each dr As DataRow In Tbl.Rows
            row = New Dictionary(Of String, Object)()
            For Each dc As DataColumn In Tbl.Columns
                row.Add(dc.ColumnName.Trim(), dr(dc))
            Next
            packet.Add(row)

        Next


        Dim jsonString As String = serializer.Serialize(packet)
        Return jsonString
    End Function

End Module
