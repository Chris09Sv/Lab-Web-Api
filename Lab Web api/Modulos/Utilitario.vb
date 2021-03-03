

Imports System.Data
Imports System.Data.SqlClient
Imports System.Text
Imports System.Security.Cryptography


Public Class ConfiguracionModulos


    ''' <summary>
    ''' Identifica el nombre de la base de datos
    ''' Que utiliza este Módulo
    ''' </summary>
    ''' <remarks></remarks>
    Public Bd_Nomina As String
    ''' <summary>
    ''' Utilizado para saber si en este cliente, se utiliza una copia local para acelerar los reportes
    ''' PAra esto el usuario debe tener acceso a los reportes
    ''' </summary>
    ''' <remarks></remarks>
    Public CopiarReportesLocalmente As Boolean
    ''' <summary>
    ''' Utilizado para saber si los reportes han sido copiados localmente
    ''' si ocurriese algun error copiandose, se asume los reportes de la red
    ''' </summary>
    ''' <remarks></remarks>
    Public ReportesCopiadosLocalmente As Boolean
    Public RutaLocal As String


    Public Sub New()
        Bd_Nomina = ""
        CopiarReportesLocalmente = ""
        RutaLocal = My.Computer.FileSystem.SpecialDirectories.AllUsersApplicationData & "\Reportes"
    End Sub
End Class



Public Module Utilitario


    Public Event RunFormulario(ByVal Id_Formulario As String)
    Public _ConfiguracionModulos As ConfiguracionModulos
    Public RutaIpWebServices As String = ""
    Public Lic_ServerName As String
    Public Lic_Fecha_Bd As String
    Public Ruta_Log_Errores As String
    Public CoreDatabase As String
    Public ProcedureDatabase As String



    'Public AppSettings_Nombre_Aplicacion As String = ""

    Public PublicDataParametro As DataTable
    Public PublicDataUsuario As DataTable
    Public FechaActual As Date
    Public Codigo_Usuario As String
    Public Nombre_Usuario As String
    Public Nombre_Modulo As String = "MILLEMNIUm"


    Public XmlActualizacionBd As String
    Public Tipo_Encabezado_Reportes As String 'Esta es para tener manejo de los encabezados de los reportes, para poder setarlo por cliente

    Public Notificar_Exportar_File As Boolean = True
    Public Directorio_Exportar_Crystar_Pdf As String = ""
    Public Ruta_Alterna_Reportes As String = ""
    Public Ruta_Fotos As String = ""


    Public Bd_Version_Requerida As Long = 0
    Public Bd_Version_Actual As Long = 0

    Private Declare Function GetVolumeInformation Lib "kernel32" _
         Alias "GetVolumeInformationA" _
    (ByVal PathName As String, ByVal VolumeNameBuffer As StringBuilder, ByVal VolumeNameSize As UInt32, ByRef VolumeSerialNumber As UInt32, ByRef MaximumComponentLength As UInt32, ByRef FileSystemFlags As UInt32, ByVal FileSystemNameBuffer As StringBuilder, ByVal FileSystemNameSize As UInt32) As Boolean



    <Microsoft.SqlServer.Server.SqlFunction()>
    Public Function DiskVolumeSerial(Optional ByVal strDrive As String = "C:\") As String
        Dim VolSer As UInt32, MaxComL As UInt32, VolFlags As UInt32
        Dim VolNameB As StringBuilder = New StringBuilder(256)
        Dim FileSysNB As StringBuilder = New StringBuilder(256)
        Dim renB As Boolean = GetVolumeInformation(strDrive, VolNameB, 256, VolSer, MaxComL, VolFlags, FileSysNB, 256)
        Return ""

    End Function

    Public Function Quitanulo(Monto As Object) As Double
        If IsDBNull(Monto) Then
            Return 0
        Else
            Return Monto

        End If
    End Function
    ''Public Function Quitanulo(Monto As String) As String
    ''    If IsDBNull(Monto) Then
    ''        Return ""
    ''    End If
    ''End Function

    Public Enum TipoDatos
        mNumeric = 2
        mString = 1
        mDate = 0
    End Enum
    Public Class Fecha_Final_Inicial
        Public Fecha_Inicial As Date
        Public Fecha_Final As Date




    End Class
    Public Function Buscar_Fecha_Mes(ByVal Mes As String, ByVal Ano As String) As Fecha_Final_Inicial
        Dim Fm As New Fecha_Final_Inicial
        Fm.Fecha_Inicial = Now
        Fm.Fecha_Final = Now

        Dim Separador_Fecha As String = IIf(InStr(Global.Microsoft.VisualBasic.DateString, "-") > 0, "-", "/")
        Dim Cmd As New SqlCommand


        Dim Cnn As New SqlConnection(SQL_CONNECTION_STRING)

        Try



            Cnn.Open()
            Cmd.Connection = Cnn
            Cmd.CommandText = IIf(Mes = 13, "select convert(datetime,'" & Ano & "/01/01',111)  as Fecha_Inicial,  convert(datetime,'" & Ano & "/12/31" & DateTime.DaysInMonth(Val(Ano), Val(Mes)) & "',111) As Fecha_Final", "select convert(datetime,'" & Ano & "/" & Mes & "/" & "01',111)  as Fecha_Inicial,  convert(datetime,'" & Ano & "/" & Mes & "/" & DateTime.DaysInMonth(Val(Ano), Val(Mes)) & "',111) As Fecha_Final")


            With Cmd.ExecuteReader
                If .Read Then

                    Fm.Fecha_Inicial = .Item("Fecha_Inicial")
                    Fm.Fecha_Final = .Item("Fecha_Final")

                End If
            End With
            'Dim scmd As New SqlCommand("GetProducts", scnnNorthwind)
            'Dim sda As New SqlDataAdapter(scmd)
            'Dim dsProducts As New DataSet()


            ' dsProducts.Tables.





        Catch ex As Exception

        End Try
        Return Fm


    End Function


    Public Enum FormatosFechas

        'Para manejar los diferentes formatos de fechas, que me intereses
        Lsf_AAAA_MM_DD_Sql_111 = 0 'Por default AAAA/MM/DD, el sql es el 111 en el convert
        Lsf_MM_DD_AAAA = 1
        Lsf_DD_MM_AAAA = 2
        Lsf_AAAAMMDD = 3
        Lsf_dd_de_MMM_de_AAAA = 4
        Lsf_MMMM_dd_AAAA_en_formato_Ingles = 5
        Lsf_Crystal_Report = 6              '
        Lsf_YYYY_MM_dd_h_m_s_sql120 = 120   'en Sql me utiliza el formato 120
        LSF_AAAAMMDDHHMMSS = 7 'Utilizado en los intercambios de informacion HLt
        Lsf_DD_MM_AAAA_HH_MM = 8
        Lsf_DD_MM_AAAA_T00 = 9
        Lsf_DD_MM_AAAA_105 = 105
        Lsf_01_MM_AAAA = 28
    End Enum
    Public IsConectado As Boolean








    Sub main()
        'Dim a As New mp 
        FechaActual = Now

    End Sub


    Public Function FormatearFecha(ByVal Fecha As Date, Optional ByVal formato As FormatosFechas = FormatosFechas.Lsf_YYYY_MM_dd_h_m_s_sql120) As String

        'Formatando la fecha a string,para manipular los datos mas facil...


        Try

            Select Case formato 'Para manejar los diferentes formatos de fechas, que me intereses
                Case 0 'Por default AAAA/MM/DD, el sql es el 111 en el convert
                    FormatearFecha = Format(Fecha, "yyyy") & "/" & Format(Fecha, "MM") & "/" & Format(Fecha, "dd")
                Case 1 'MM/DD/AAAA
                    FormatearFecha = Format(Fecha, "MM") & "/" & Format(Fecha, "dd") & "/" & Format(Fecha, "yyyy")
                Case 2 'DD/MM/AAAA
                    FormatearFecha = Format(Fecha, "dd") & "/" & Format(Fecha, "MM") & "/" & Format(Fecha, "yyyy")
                Case 3 'AAAAMMDD
                    FormatearFecha = Format(Fecha, "yyyy") & Format(Fecha, "MM") & Format(Fecha, "dd")
                Case 4 'dd de MMM de AAAA`
                    FormatearFecha = Format(Fecha, "dd") & " de " & Convertir_Mes(Month(Fecha)) & " del " & Format(Fecha, "yyyy")
                Case 5 ' MMMM dd,AAAA` en formato Ingles
                    FormatearFecha = Convertir_Mes_ingles(Month(Fecha)) & " " & Format(Fecha, "dd") & "," & Format(Fecha, "yyyy")
                Case 6 ' AAAA,MM,DD` en formato Para Crystal Report
                    FormatearFecha = "date(" & Format(Fecha, "yyyy") & "," & Format(Fecha, "MM") & "," & Format(Fecha, "dd") & ")"
                Case 7 'LSF_AAAAMMDDHHMMSS=7 

                    FormatearFecha = Format(Fecha, "yyyy") & Format(Fecha, "MM") & Format(Fecha, "dd") & Format(Fecha, "HHmmss")
                Case 120    '"YYYY-MM-dd h:m:s"   'en Sql me utiliza el formato 120
                    FormatearFecha = Format(Fecha, "yyyy") & "-" & Format(Fecha, "MM") & "-" & Format(Fecha, "dd") & " " & Format(Fecha, "HH:mm:ss")
                Case FormatosFechas.Lsf_DD_MM_AAAA_T00
                    FormatearFecha = Format(Fecha, "yyyy") & "-" & Format(Fecha, "MM") & "-" & Format(Fecha, "dd") & "T24:00:00.000"
                Case FormatosFechas.Lsf_DD_MM_AAAA_HH_MM
                    FormatearFecha = Format(Fecha, "yyyy") & "-" & Format(Fecha, "MM") & "-" & Format(Fecha, "dd") & " " & Format(Fecha, "HH:mm:ss")
                    FormatearFecha = Format(Fecha, "dd") & " de " & Convertir_Mes(Month(Fecha)) & " del " & Format(Fecha, "yyyy") & " a las " & Format(Fecha, "t")
                Case FormatosFechas.Lsf_DD_MM_AAAA_105 'DD-MM-AAAA
                    FormatearFecha = Format(Fecha, "dd") & "-" & Format(Fecha, "MM") & "-" & Format(Fecha, "yyyy")

                Case FormatosFechas.Lsf_01_MM_AAAA 'DD-MM-AAAA
                    FormatearFecha = "01-" & Format(Fecha, "MM") & "-" & Format(Fecha, "yyyy")


                Case Else
                    FormatearFecha = ""
            End Select

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error formateando la fecha")
            FormatearFecha = ""
        End Try

    End Function

    Public Function Convertir_Mes(ByVal Mes As Long) As String
        Select Case Mes
            Case 1
                Convertir_Mes = "Enero"
            Case 2
                Convertir_Mes = "Febrero"
            Case 3
                Convertir_Mes = "Marzo"
            Case 4
                Convertir_Mes = "Abril"
            Case 5
                Convertir_Mes = "Mayo"
            Case 6
                Convertir_Mes = "Junio"
            Case 7
                Convertir_Mes = "Julio"
            Case 8
                Convertir_Mes = "Agosto"
            Case 9
                Convertir_Mes = "Septiembre"
            Case 10
                Convertir_Mes = "Octubre"
            Case 11
                Convertir_Mes = "Noviembre"
            Case Else
                Convertir_Mes = "Diciembre"
        End Select

    End Function
    Public Function Convertir_Mes_ingles(ByVal Mes As Long) As String
        Select Case Mes
            Case 1
                Convertir_Mes_ingles = "January"
            Case 2
                Convertir_Mes_ingles = "February"
            Case 3
                Convertir_Mes_ingles = "March"
            Case 4
                Convertir_Mes_ingles = "April"
            Case 5
                Convertir_Mes_ingles = "May"
            Case 6
                Convertir_Mes_ingles = "June"
            Case 7
                Convertir_Mes_ingles = "July"
            Case 8
                Convertir_Mes_ingles = "August"
            Case 9
                Convertir_Mes_ingles = "September"
            Case 10
                Convertir_Mes_ingles = "October"
            Case 11
                Convertir_Mes_ingles = "November"
            Case Else
                Convertir_Mes_ingles = "December"
        End Select

    End Function
    Public Function CompletarDatosEspacios(ByVal Text As String, ByVal Tamano As Long, Optional ByVal Isnumeric As Boolean = False) As String
        If Tamano <= 0 Then
            CompletarDatosEspacios = ""
            Exit Function
        End If
        Text = Trim(Text)
        If Len(Trim(Text)) >= Tamano Then
            CompletarDatosEspacios = Mid(Text, 1, Tamano)

        Else
            ' Dim Replicar As Long
            '   

            CompletarDatosEspacios = Text & Space(Tamano - Len(Trim(Text)))
            'End If
        End If

    End Function


    Public Function CompletarDatosEspacios(ByVal Text As String, ByVal Tamano As Long, ByVal Isnumeric As Boolean, ByVal RellenarConCeros As Boolean) As String
        If Tamano <= 0 Then
            CompletarDatosEspacios = ""
            Exit Function
        End If
        Text = Trim(Text)
        If Len(Trim(Text)) >= Tamano Then
            CompletarDatosEspacios = Mid(Text, 1, Tamano)

        Else
            If Isnumeric Then
                Dim replicar As Integer = (Tamano - Len(Trim(Text)))
                'CompletarDatosEspacios = StrDup(replicar, "0") & Trim(Text)
                Dim ret As String = StrDup(replicar, IIf(RellenarConCeros, "0", " ")) & Trim(Text)

                CompletarDatosEspacios = ret 'Space(Tamano - Len(Trim(Text))) & Trim(Text)
            Else

                CompletarDatosEspacios = Space(Tamano - Len(Trim(Text))) & Text
            End If
        End If

    End Function

    Public Function AgregarEspacios(ByVal Text As String, ByVal Tamano As Long, ByVal Isnumeric As Boolean, ByVal RellenarConCeros As Boolean) As String
        If Tamano <= 0 Then
            AgregarEspacios = ""
            Exit Function
        End If
        Text = Trim(Text)
        If Len(Trim(Text)) >= Tamano Then
            AgregarEspacios = Mid(Text, 1, Tamano)

        Else
            If Isnumeric Then
                Dim replicar As Integer = (Tamano - Len(Trim(Text)))
                'CompletarDatosEspacios = StrDup(replicar, "0") & Trim(Text)
                Dim ret As String = Trim(Text) & StrDup(replicar, IIf(RellenarConCeros, " ", " "))

                AgregarEspacios = ret 'Space(Tamano - Len(Trim(Text))) & Trim(Text)
            Else

                AgregarEspacios = Text & Space(Tamano - Len(Trim(Text)))
            End If
        End If

    End Function

End Module


Public Module Modulo_Key_Registro


    'Esta clase es para controlar la pirateria del Software
    'esta genera un Id. unico, para cada nombre. la función de esta es verificar que se
    'haya registrado un nombre correcto
    '




    Public Function CompareKey(ByVal Name As String, ByVal Key_Registro As String) As Key_Compare
        Dim Retorno As New Key_Compare

        'Esta función determina, si el sistema ha sido registrado correctamente, y a su vez, determina la cantidad de licencias registradas
        Try

            Dim Str_Lic_Key As String
            Dim Str_Key As String
            Dim Str As String, Dig_Verif As String, Digito_Verificado As Boolean

            Name = UCase(Name)
            Key_Registro = UCase(Key_Registro)

            'Determino la Proporcion que equivale a las licencias


            Dig_Verif = AgregarVerificacion(Mid(Key_Registro, 1, 15))
            Digito_Verificado = Dig_Verif = Mid(Key_Registro, 16, 1)

            '\Obtengo la parte, que Corresponde al nombre
            Str_Lic_Key = Mid(Key_Registro, 2, 4) &
                    Mid(Key_Registro, 7, 8)

            'Obtengo la parte, que corresponde a las licencias

            Str_Key = Mid(Key_Registro, 15, 1) &
                   Mid(Key_Registro, 6, 1) &
                   Mid(Key_Registro, 1, 1)


            Key_Registro = Key_Registro


            Dim lPart1 As Long
            Dim lPart2 As Long
            Dim ch As Long
            Dim i As Long

            Dim X As Long

            Dim n As Long

            Dim mName As String = "", U_1 As String, U_2 As String
            'Trabajo solo con las primeras 40 posiciones a Registrar
            'Invierto el el String por Ejemplo si pedro, el string resultante es ordep
            Name = UCase(Trim$(Name))

            For X = 0 To Len(Name) - 1
                mName = mName & Mid(Name, Len(Name) - X, 1)
                If X = 5 Then
                    mName = mName & "A"
                End If

                If X = 10 Then
                    mName = mName & "L" 'C
                End If

                If X = 12 Then
                    mName = mName & "A" 'L
                End If

                If X = 26 Then
                    mName = mName & "N" 'G
                End If

            Next
            Name = Mid$(mName, 1, 40)



            If Len(Name) = 0 Then
                Exit Try
            End If

            For i = 1 To Len(Name)
                'Si es un Ascii mayor que 127, da un error, por eso lo convierto

                If Asc(Mid$(Name, i, 1)) < 128 Then
                    ch = Asc(Mid$(Name, i, 1)) * &H100
                Else
                    ch = Asc(Mid$(Name, i, 1)) - 126 * &H100

                End If
                'Hago el recorrido 8 veces, para cada Caracter



                For n = 1 To 8


                    If (((ch Xor lPart1) Mod &H10000) And &H8000&) = 0 Then
                        'Bit shift 1 bit a la izquierda
                        lPart1 = lPart1 And &HFFFFFFF
                        lPart1 = lPart1 * 2
                    Else
                        'Bit shift 1 bit a la izquierda
                        lPart1 = lPart1 And &HFFFFFFF
                        lPart1 = lPart1 * 2
                        'Xor con el Numerito Magico
                        lPart1 = lPart1 Xor &H1021&
                    End If
                    ch = ch * 2
                Next n
            Next i
            'Le agrego un Bit
            lPart1 = lPart1 + &H63&
            'Como solo quiero 4 Digitos
            lPart1 = lPart1 Mod &H10000
            'Part2 es simple


            For i = 1 To Len(Name)
                lPart2 = lPart2 + (Asc(Mid$(Name, i, 1)) * (i - 1))
            Next i

            U_1 = Chr((90 - (lPart1 Mod 10)))
            U_2 = Chr(65 + (lPart2 Mod 26))



            Str = Replicate(4 - Len(CStr(Hex(lPart1))), "0") _
             & CStr(Hex(lPart1)) _
             & Replicate(4 - Len(CStr(Hex(lPart2))), "0") _
             & CStr(Hex(lPart2)) _
             & CStr(Hex(((lPart2 + lPart2) Mod 26) + 1308)) '_1544

            Dig_Verif = AgregarVerificacion(Str)
            Str = Str & Dig_Verif
            '& CStr(Hex(((lPart2 + lPart2) Mod 1011) + 1972))
            'ya tengo el Registro por el Nombre, verifico, si se trata de un registro valido

            Retorno.Key_Valido = (Str_Lic_Key = Str) And Digito_Verificado
            'ahora verifico la cantidad de licencias

            Retorno.Key_Licencias = Obtener_Numero(Str_Key)



        Catch ex As Exception
            Retorno.Key_Valido = False
        End Try

        Return Retorno

    End Function

    Public Function Descomp(ByVal Str As String) As String
        Descomp = ""
        Dim Tot = 0
        Dim i
        For i = 1 To Len(Str)
            Tot = Tot + Asc(Mid(Str, i, 1))
        Next
        Descomp = Chr((Tot Mod 26) + 65)
    End Function


    Private Function Obtener_Numero(ByVal Str As String) As String
        On Error GoTo err
        'Esta funcion me devuelve el no. de licencia, a partir
        Dim S_1 As String, S_2 As String, S_ID
        S_1 = Mid(Str, 1, 1)
        S_ID = Mid(Str, 2, 1)
        S_2 = Mid(Str, 3, 1)

        Obtener_Numero = (((90 - (Asc(S_1) Mod 9)) - Asc(S_2)) * 20) + Asc(S_ID) - 65 - (((Asc(S_1) - 65)) Mod 6)
err:

    End Function

    Private Function AgregarVerificacion(ByVal Str As String) As String
        'Para Agregarle un Digito de Verificación

        AgregarVerificacion = ""

        Dim Tot = 0
        Dim i
        For i = 1 To Len(Str)
            Tot = Tot + Asc(Mid(Str, i, 1))
        Next
        AgregarVerificacion = Chr(85 - (Tot Mod 19))
    End Function

    Public Function Replicate(ByVal Veces As Long, ByVal Caracter As String) As String
        Dim a As String = "", i As Long = 0
        While i < Veces
            a = a & Caracter
            i += 1
        End While
        Return a
    End Function

End Module
Public Class Key_Compare

    Public Key_Licencias As String
    Public Key_Valido As Boolean

    Public Sub New()

    End Sub
End Class

Public Class Key_Reg
    Public Key_Registro As String, Key_Licencias As String, Key_Nombre As String

End Class

Public Module KeyGen


    Public Function GenerateKey(ByVal Name As String, Optional ByVal No_Licencias As Long = 0) As Key_Reg
        Dim MGenerateKey As New Key_Reg
        On Error GoTo err
        Dim lPart1 As Long
        Dim lPart2 As Long
        Dim ch As Long
        Dim I As Long

        Dim x As Long

        Dim n As Long

        Dim Mname As String = "", U_1 As String, U_2 As String, S_1 As String, S_2 As String, S_ID
        'Trabajo solo con las primeras 40 posiciones a Registrar
        'Invierto el el String por Ejemplo si pedro, el string resultante es ordep
        Name = UCase(Trim$(Name))
        For x = 0 To Len(Name) - 1
            Mname = Mname & Mid(Name, Len(Name) - x, 1)
            If x = 5 Then
                Mname = Mname & "A"
            End If

            If x = 10 Then
                Mname = Mname & "L" '"C"
            End If

            If x = 12 Then
                Mname = Mname & "A" '"L"
            End If

            If x = 26 Then
                Mname = Mname & "N" '"G"
            End If

        Next
        Name = Mid$(Mname, 1, 40)



        If Len(Name) = 0 Then
            Return MGenerateKey
            Exit Function
        End If
        Dim M_Ascii

        For I = 1 To Len(Name)
            'Si es un Ascii mayor que 127, da un error, por eso lo convierto

            If Asc(Mid$(Name, I, 1)) < 128 Then
                ch = Asc(Mid$(Name, I, 1)) * &H100
            Else
                ch = Asc(Mid$(Name, I, 1)) - 126 * &H100

            End If
            'Hago el recorrido 8 veces, para cada Caracter



            For n = 1 To 8


                If (((ch Xor lPart1) Mod &H10000) And &H8000&) = 0 Then
                    'Bit shift 1 bit a la izquierda
                    lPart1 = lPart1 And &HFFFFFFF
                    lPart1 = lPart1 * 2
                Else
                    'Bit shift 1 bit a la izquierda
                    lPart1 = lPart1 And &HFFFFFFF
                    lPart1 = lPart1 * 2
                    'Xor con el Numerito Magico
                    lPart1 = lPart1 Xor &H1021&
                End If
                ch = ch * 2
            Next n
        Next I
        'Le agrego un Bit
        lPart1 = lPart1 + &H63&
        'Como solo quiero 4 Digitos
        lPart1 = lPart1 Mod &H10000
        'Part2 es simple


        For I = 1 To Len(Name)
            lPart2 = lPart2 + (Asc(Mid$(Name, I, 1)) * (I - 1))
        Next I
        Dim Users As String
        U_1 = Chr((90 - (lPart1 Mod 10)))
        U_2 = Chr(65 + (lPart2 Mod 26))


        Dim Str As String, Dig_Verif As String

        Str = Replicate(4 - Len(CStr(Hex(lPart1))), "0") _
         & CStr(Hex(lPart1)) _
         & Replicate(4 - Len(CStr(Hex(lPart2))), "0") _
         & CStr(Hex(lPart2)) _
         & CStr(Hex(((lPart2 + lPart2) Mod 26) + 1308)) '_ 1544

        Dig_Verif = AgregarVerificacion(Str)
        Str = Str & Dig_Verif
        '& CStr(Hex(((lPart2 + lPart2) Mod 1011) + 1972))
        'ya tengo el Registro por el Nobre
        MGenerateKey.Key_Nombre = Str
        'Si voy a Registrar la cantidad de licencia
        If No_Licencias > 0 Then
            'Busco una cantidad de
            S_1 = Descomp(Str)
            S_2 = Chr(90 - (Asc(S_1) Mod 9) - Int(No_Licencias / 20))
            S_ID = Chr((65 + (No_Licencias Mod 20)) + (Asc(S_1) - 65) Mod 6)


            Str = S_2 & Replicate(4 - Len(CStr(Hex(lPart1))), "0") _
           & CStr(Hex(lPart1)) & S_ID _
          & Replicate(4 - Len(CStr(Hex(lPart2))), "0") _
          & CStr(Hex(lPart2)) _
          & CStr(Hex(((lPart2 + lPart2) Mod 26) + 1308)) & Dig_Verif & S_1  '1544

            Dig_Verif = AgregarVerificacion(Str)
            MGenerateKey.Key_Licencias = S_1 &
   S_ID _
   & S_2 & Dig_Verif
            'Le agrego un numero de verificacion al final
            'para asegurarme de que no se le cambie la cantidad de licencias
            Str = Str & Dig_Verif

            'ZCC9FF0FA915B1


            MGenerateKey.Key_Registro = UCase(Str)
        Else
            MGenerateKey.Key_Licencias = ""
            MGenerateKey.Key_Registro = ""
        End If
        Return MGenerateKey
err:
        MsgBox("No pudo generarse la clave de instalacion " & vbCr & Err.Description, vbCritical)
    End Function




    Function AgregarVerificacion(ByVal Str As String) As String
        'Para Agregarle un Digito de Verificación

        AgregarVerificacion = ""

        Dim Tot
        Dim I
        For I = 1 To Len(Str)
            Tot = Tot + Asc(Mid(Str, I, 1))
        Next
        AgregarVerificacion = Chr(85 - (Tot Mod 19))
    End Function


    Public Function VerifiDigit(ByVal Mstr As String) As String
        Dim Total_Ascii As Long, Newstr As String
        Dim MCheckSum As String = ""
        'En este caso, estoy Obviando los caracteres que marcan el inicio y el fin de texto

        For Each Letra As String In Mstr
            Total_Ascii += (Asc(Letra) * Asc(Letra))
        Next
        Newstr = Trim(Hex(Total_Ascii))
        Newstr = Mid(Newstr, 1, 1) & Right(Newstr, 1)

        'Newstr = Right(Trim(Hex(Total_Ascii)), 2)
        'Select Case mChekPosition
        'Case ChekPosition.Cp_Antepenultimo
        '    MCheckSum = Mid(Mstr, 1, Len(Mstr) - 1) & Newstr & Right(Mstr, 1)
        'Case ChekPosition.Cp_final
        MCheckSum = Mstr & Newstr
        'End Select
        Return MCheckSum
    End Function



    Public Function GetDriveSerialNumber() As String
        Dim DriveSerial As Integer
        'Create a FileSystemObject object
        Dim fso As Object = CreateObject("Scripting.FileSystemObject")
        Dim Drv As Object = fso.GetDrive(fso.GetDriveName(My.Computer.FileSystem.SpecialDirectories.ProgramFiles))
        With Drv
            If .IsReady Then
                DriveSerial = .SerialNumber
            Else    '"Drive Not Ready!"
                DriveSerial = -1
            End If

        End With

        Return GenerateKey(DriveSerial.ToString("X2") & Now.ToLongTimeString(), 0).Key_Nombre

    End Function
End Module
Public Class JSONArguments
    Public sessionid As String
    Public Datos As Dictionary(Of String, Object)
End Class
