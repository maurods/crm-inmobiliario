Imports MySql.Data.MySqlClient
Imports System.Text
Imports System.Security.Cryptography
Imports System.IO
Public Module conexionBD
    Public conexion As New MySqlConnection

    'Metodo para la conexion a la base de datos
    Public Sub conectar()
        Try
            conexion.Close()
            'conexion.ConnectionString = "server=10.0.29.13;database=enigma;user id=mhegui1;password=mhegui123*;"

            conexion.ConnectionString = "server=localhost;database=enigma;user id=root;password=Enigma.2019;"
            'conexion.ConnectionString = "server=localhost;database=enigma;user id=root;password=Veleda4193;"

       


            conexion.Open()

        Catch ex As StackOverflowException
            MsgBox("Error al conectar a la base de datos")
        End Try
    End Sub




    'FUNCION PARA LOGUEARSE AL SISTEMA
    Function existeusuario(ByVal usuario As String, ByVal contraseña As String, ByVal tipo As String) As Boolean

        Dim oDataAdapter As MySqlDataAdapter
        Dim oDataSet As New DataSet
        Dim sSql As String
        Dim sw As Boolean = False
        Dim contador As Integer = 0
        Dim contraseñaencriptada As String
        Try


            If tipo = "Cliente" Then

                sSql = "SELECT personas.IdPersona , personas.Nombre, personas.Apellido, personas.Funcion, Contraseña, Usuario , Cel FROM personas , usuarios where personas.IdPersona = usuarios.IdPersona and UPPER(Usuario) = UPPER('" & usuario & "') AND UPPER(Contraseña) = UPPER('" & contraseña & "')"


            Else

                contraseñaencriptada = generarClaveSHA1(contraseña)

                sSql = "SELECT IdSucursal , personas.IdPersona, personas.Nombre, personas.Apellido, personas.Funcion, Contraseña, Usuario FROM personas , usuarios where personas.IdPersona = usuarios.IdPersona and usuarios.Usuario='" & usuario & "'AND usuarios.Contraseña='" & contraseña & "' OR usuarios.Contraseña='" & contraseñaencriptada & "';"


            End If


            oDataAdapter = New MySqlDataAdapter(sSql, conexion)
            conectar()


            oDataAdapter.Fill(oDataSet, "datosusuario")

            If (oDataSet.Tables("datosusuario").Rows.Count = 0) Then

                MsgBox("usuario o contraseña incorrectos")

            Else


                contraseña = oDataSet.Tables("datosusuario").Rows(0)("Contraseña").ToString()


                If tipo <> "Cliente" Then


                    'AGREGA LA FUNCION QUE TIENE EL USUARIO PARA INICIAR SESION
                    FormLogin.tipoacceso = oDataSet.Tables("datosusuario").Rows(0)("Funcion").ToString()

                    'AGREGA EL ID DEL FUNCIONARIO PARA QUE LUEGO ASOCIARLO A LAS NOTIFICACIONES QUE HACE
                    FormPrincipal.IdPersonal = oDataSet.Tables("datosusuario").Rows(0)("IdPersona").ToString()

                    'AGREGA EL ID DEL FUNCIONARIO PARA QUE LUEGO ASOCIARLO A LAS NOTIFICACIONES QUE HACE
                    FormPrincipal.IdSucursal = oDataSet.Tables("datosusuario").Rows(0)("IdSucursal").ToString()



                Else

                    BusquedasPropiedadesInteresados.IdCliente = oDataSet.Tables("datosusuario").Rows(0)("IdPersona").ToString()

                    BusquedasPropiedadesInteresados.Celular = oDataSet.Tables("datosusuario").Rows(0)("Cel").ToString()
                    BusquedasPropiedadesInteresados.cliente = oDataSet.Tables("datosusuario").Rows(0)("Nombre").ToString()



                End If




                sw = True

            End If

        Catch ex As Exception

            MsgBox(ex.Message)

        End Try
        Return (sw)
    End Function



    'ENCRIPTA LAS CONTRASEÑAS PARA QUE NO SEAN VISIBLES DESDE LA BD
    Function generarClaveSHA1(ByVal cadena As String) As String

        Dim enc As New UTF8Encoding
        Dim data() As Byte = enc.GetBytes(cadena)

        Dim result() As Byte
        Dim sha As New SHA1CryptoServiceProvider
        result = sha.ComputeHash(data)

        Dim sb As New StringBuilder
        Dim max As Int32 = result.Length

        For i As Integer = 0 To max - 1

            If (result(i) < 16) Then

                sb.Append("0")

            End If

            sb.Append(result(i).ToString("x"))


        Next

        generarClaveSHA1 = sb.ToString().ToUpper


    End Function










    'FUNCION CONSULTA SQL

    'Dim consultaUsuarios As String = "SELECT * FROM usuarios;"
    'consultaSQL(consultaUsuarios).Tables("usuarios").Rows(0)("apellido")
    Public Function consultaSQL(ByRef SQL As String) As DataSet
        Dim datos As New DataSet
        Try


            conectar()

            Dim adaptador As New MySqlDataAdapter(SQL, conexion)

            adaptador.Fill(datos, "soncontraseñasiguales")
            adaptador.Fill(datos, "quebarrioshay")
            adaptador.Fill(datos, "consultainmueble")
            adaptador.Fill(datos, "img_inmuebles")
            adaptador.Fill(datos, "personas")
            adaptador.Fill(datos, "inmueble")
            adaptador.Fill(datos, "sucursal")

            conexion.Close()

        Catch ex As Exception


            MsgBox(ex.Message)

        End Try




        Return datos
    End Function




    'Procedimiento para agregar, modificar y eliminar en MySQL
    Sub updateSQL(ByVal sql As String)


        'Declaramos nuestro objeto de tipo SQLiteCommand para ejecutar la consulta
        Dim cmd As New MySqlCommand(sql, conexion)

        Try

            conectar()
            cmd.ExecuteNonQuery()
            conexion.Close()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub



    'Procedimiento para agregar, modificar y eliminar en MySQL
    Sub agregarmodificarImagenes(ByVal sql As String, ByVal imagen As PictureBox)
        'Creamos una variable de tipo MemoryStram
        Dim MS As New MemoryStream
        'Guardamos en la variable MS el contenido del PictureBox
        imagen.Image.Save(MS, imagen.Image.RawFormat)
        'Pasamos a Byte nuestra imagen
        Dim Imagenes() As Byte = MS.GetBuffer

        'Declaramos nuestro objeto de tipo SQLiteCommand para ejecutar la consulta
        Dim cmd As New MySqlCommand(sql, conexion)
        'Agregamos el parametro de nuestra varible imagenes que es de tipo Byte
        cmd.Parameters.AddWithValue("@imagen", Imagenes)
        Try


            conectar()
            cmd.ExecuteNonQuery()
            conexion.Close()
            ' MsgBox("Operación realizada con exito")
        Catch ex As Exception
            MsgBox("No se pueden guardar los registro")
        End Try
    End Sub













End Module


