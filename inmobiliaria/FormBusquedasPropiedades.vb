Imports System.IO
Imports System.Runtime.InteropServices
Imports MySql.Data.MySqlClient
Imports QRCoder

Public Class BusquedasPropiedadesInteresados
    Dim mostrarImagenes As Boolean = True
    Dim buscar As Boolean = False
    Dim cambiartamaño As Integer = 0
    Dim otrosFiltrosImagenes As Boolean = False
    Dim inactivo As Boolean = False
    Dim contadorCerrar As Integer = 0
    Dim cantidadimg As Integer = 0
    Public cliente As String = ""
    Public Idinmueble = 0
    Dim fechaactual As Date = Today
    Dim horaactual As DateTime = TimeString
    Public Celular As String
    Public IdCliente As String
    Public barrio1 As String
    Public barrio2 As String
    Public barrio3 As String
    Public barrio4 As String
    Public barrio5 As String
    Public barrio6 As String
    Public barrio7 As String

    Public jardin As String
    Public piscina As String
    Public garage As String
    Public amoblada As String
    Public barbacoa As String
    Public mascotas As String
    Public busc As Boolean = False

    '------------------------------------------------LOAD-------------------------------------------------------------------
    Private Sub BusquedasPropiedadesInteresados_Load(sender As Object, e As EventArgs) Handles Me.Load



        'REDIMENSIONA LOS OBJETOS AL INICIAR
        With Me
            .PanelBusquedaPropiedadesInteresadosInmuebles.Height = 50
            .PanelBusquedaPropiedadesInteresadosInmuebles.Width = PanelBusquedasPropiedades.Size.Width - 50
            .PanelBusquedaPropiedadesInteresadosInmuebles.Location = New System.Drawing.Point((PanelBusquedasPropiedades.Size.Width - PanelBusquedasPropiedades.Size.Width) + 20, PanelBusquedasPropiedades.Size.Height - 50)

            'REDIMENCIONA LOS PANELES SEGUN EL ANCHO Y LARGO DE LA PANTALLA MAXIMIZADA
            .PanelBusquedaPropiedadesInteresadosCentral.Size = New System.Drawing.Size(PanelBusquedasPropiedades.Size.Width - PanelBusquedasPropiedades.Size.Width / 3 + 30, PanelBusquedasPropiedades.Size.Height / 8)
            .PanelBusquedaPropiedadesInteresadosCentral.Location = New System.Drawing.Point(PanelBusquedasPropiedades.Size.Width / 6, PanelBusquedasPropiedades.Size.Height / 2)
            .LabelBusquedaPropiedadesInteresadosTitulo.Location = New System.Drawing.Point(PanelBusquedasPropiedades.Size.Width / 2.6, PanelBusquedasPropiedades.Size.Height / 4)
            'Le pone de fondo al label la pick
            .LabelBusquedaPropiedadesInteresadosTitulo.Parent = PickBusquedaPropiedadesInteresados
            'Selecciona el titulo para que no lo seleccione al combo
            .LabelBusquedaPropiedadesInteresadosTitulo.Select()

            .DataBuscarPropiedadesImagenes.Size = New System.Drawing.Size(PanelBusquedaPropiedadesInteresadosInmuebles.Size.Width - 270, 402)
        End With





    End Sub






    '----------------------------------------------BOTON BUSCAR-------------------------------------------------------------------------
    Private Sub BtnBusquedaPropiedadesInteresadosBuscar_Click(sender As Object, e As EventArgs) Handles BtnBusquedaPropiedadesInteresadosBuscar.Click

        busc = True



        barrio1 = "Si"
        barrio2 = "Si"
        barrio3 = "Si"
        barrio4 = "Si"
        barrio5 = "Si"
        barrio6 = "Si"
        barrio7 = "Si"


        DataBuscarPropiedadesImagenes.DataSource = Nothing

        '----------------------DETECTAR ACTIVIDAD------------------------------------------
        INPUT.cbSize = Marshal.SizeOf(INPUT) '¿? PERO ES NECESARIO                       '---
        TimerInactividad.Interval = 1000 'MONITORIZAREMOS GetLastInputInfo CADA SEGUNDO       '---
        TimerInactividad.Start()                                                              '---
        '-----------------------------------------------------------------------------------







        'MUESTRA EL NOMBRE DE LOS CHECK BOX DE BARRIOS SEGUN SELECCIONE EL DEPARTAMENTO
        Dim listadoBarrios As String = "select distinct Barrios from barrios , ubicacion where barrios.IdUbicacion = ubicacion.IdUbicacion and ubicacion.Departamento = '" & BoxBusquedaPropiedadesInteresadosDepartamento.SelectedItem & "';"

        Try
            CheckBusquedasPropiedadesBarrio1.Text = consultaSQL(listadoBarrios).Tables("quebarrioshay").Rows(0)("Barrios").ToString


        Catch ex As Exception
            CheckBusquedasPropiedadesBarrio1.Visible = False
        End Try

        Try
            CheckBusquedasPropiedadesBarrio2.Text = consultaSQL(listadoBarrios).Tables("quebarrioshay").Rows(1)("Barrios").ToString

        Catch ex As Exception
            CheckBusquedasPropiedadesBarrio2.Visible = False
        End Try

        Try
            CheckBusquedasPropiedadesBarrio3.Text = consultaSQL(listadoBarrios).Tables("quebarrioshay").Rows(2)("Barrios").ToString

        Catch ex As Exception
            CheckBusquedasPropiedadesBarrio3.Visible = False
        End Try

        Try
            CheckBusquedasPropiedadesBarrio4.Text = consultaSQL(listadoBarrios).Tables("quebarrioshay").Rows(3)("Barrios").ToString

        Catch ex As Exception
            CheckBusquedasPropiedadesBarrio4.Visible = False
        End Try

        Try
            CheckBusquedasPropiedadesBarrio5.Text = consultaSQL(listadoBarrios).Tables("quebarrioshay").Rows(4)("Barrios").ToString

        Catch ex As Exception
            CheckBusquedasPropiedadesBarrio5.Visible = False
        End Try

        Try
            CheckBusquedasPropiedadesBarrio6.Text = consultaSQL(listadoBarrios).Tables("quebarrioshay").Rows(5)("Barrios").ToString

        Catch ex As Exception
            CheckBusquedasPropiedadesBarrio6.Visible = False
        End Try

        Try
            CheckBusquedasPropiedadesBarrio7.Text = consultaSQL(listadoBarrios).Tables("quebarrioshay").Rows(6)("Barrios").ToString

        Catch ex As Exception
            CheckBusquedasPropiedadesBarrio7.Visible = False
        End Try




        'MUESTRA EL PANEL QUE VA A MOSTRAR LAS IMAGENES DE LOS INMUEBLES (porque en un principio estaba solo en menu de busqueda)
        PanelBusquedaPropiedadesInteresadosInmuebles.Visible = True

        'cambia la posicion del titulo (porque el panel de imagnes necesita mas espacio para mostrar todas las imagenes)
        LabelBusquedaPropiedadesInteresadosTitulo.Location = New System.Drawing.Point(PanelBusquedasPropiedades.Size.Width / 2.6, 30)

        PanelBusquedaPropiedadesInteresadosCentral.Location = New System.Drawing.Point(PanelBusquedasPropiedades.Size.Width / 6, PanelBusquedasPropiedades.Size.Height / 8)

        PanelBusquedaPropiedadesInteresadosInmuebles.Location = New System.Drawing.Point(20, PanelBusquedasPropiedades.Size.Height / 3.5)

        PanelBusquedaPropiedadesInteresadosInmuebles.Height = PanelBusquedasPropiedades.Size.Height - 230
        PanelBusquedasPropiedadesIzquierdo.Height = PanelBusquedasPropiedades.Size.Height - 250
        DataBuscarPropiedadesImagenes.Height = PanelBusquedasPropiedades.Size.Height - 260





        DataBuscarPropiedadesImagenes.DataSource = consultaSQL("SELECT IdInmueble,	Barrios, NumDormitorios, NumBaños , Jardin , Amoblada , Garage , barbacoa , Piscina , Moneda, Precio  From Inmueble, Ubicacion, caracteristicas_inm, publicacion , barrios
where ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio And inmueble.IdInm = caracteristicas_inm.idInm And Inmueble.IdPubli = publicacion.IdPubli And
	ubicacion.Departamento = '" & BoxBusquedaPropiedadesInteresadosDepartamento.SelectedItem & "' and caracteristicas_inm.TipoInm = '" & BoxBusquedaPropiedadesInteresadosTipoInmueble.SelectedItem & "' and publicacion.TipoPubli = '" & BoxBusquedaPropiedadesInteresadosCompraVenta.SelectedItem & "';
").Tables("consultainmueble")

        DataBuscarPropiedadesImagenes.Columns(1).Visible = False


        For i = 0 To DataBuscarPropiedadesImagenes.Rows.Count - 1

            Dim r As DataGridViewRow = DataBuscarPropiedadesImagenes.Rows(i)
            r.Height = 100


        Next







    End Sub



    'BOTON INGRESAR SI USUARIO Y CONTRASEÑA SON CORRECTOS
    Private Sub BtnBuscarPropiedadesIngresar_Click(sender As Object, e As EventArgs) Handles BtnBuscarPropiedadesIngresar.Click
        PanelFormBusquedaPropiedadesTeclado.Visible = False


        If BtnBuscarPropiedadesIngresar.Text = "Ingresar" Then



            If existeusuario(TxtBusquedasPropiedadesUsuario.Text, TxtBusquedasPropiedadesContraseña.Text, "Cliente") = True Then



                BtnBusquedasPropiedadesAgendar.Visible = True
                TxtBusquedasPropiedadesUsuario.Enabled = False
                TxtBusquedasPropiedadesContraseña.Enabled = False
                BtnBuscarPropiedadesIngresar.Text = "Cerrar sesion"




            End If


        Else



            If BtnBuscarPropiedadesIngresar.Text = "Cerrar sesion" Then

                BtnBuscarPropiedadesIngresar.Text = "Ingresar"
                TxtBusquedasPropiedadesUsuario.Enabled = True
                TxtBusquedasPropiedadesContraseña.Enabled = True
                TxtBusquedasPropiedadesUsuario.Text = "Usuario"
                TxtBusquedasPropiedadesContraseña.Text = "Contraseña"



                BtnBusquedasPropiedadesAgendar.Visible = False

            End If



        End If





    End Sub




    'BOTON AGENDAR
    Private Sub BtnBusquedasPropiedadesAgendar_Click(sender As Object, e As EventArgs)



    End Sub








    '----TIMER PARA REFRESCAR LA PANTALLA SI NO HAY ACTIVIDAD EN EL SISTEMA--------------------------
    Private Sub Inactividad_(sender As System.Object, e As System.EventArgs) Handles TimerInactividad.Tick

        GetLastInputInfo(INPUT) 'COMPROBAMOS LA FUNCION CADA SEGUNDO

        Dim TOTAL As Integer = Environment.TickCount        'MILISEGUNDOS DESDE QUE SE INICIO LA SESION
        Dim ULTIMO As Integer = INPUT.dwTime 'MILISEGUNDOS DESDE LA ULTIMA ACTIVIDAD (TECLADO Y MOUSE)
        Dim INTERVALO As Integer = (TOTAL - ULTIMO) / 1000 'DIFERENCIA EN SEGUNDOS 

        If INTERVALO >= 100000 Then 'SI LA INACTIVIDAD ES 10 SEGUNDOS O MÁS   (Numero 10 para 10 segundos)
            TimerInactividad.Stop()
            inactivo = True


            '-------------------VUELVE A POSICIONES ORIGINALES Y CONTADORES EN CERO -------------------------------------
            imagenesFullPantalla.Close()
            PanelFormBusquedasPropiedadesPanelImagenes.Visible = False
            PanelFormBusquedaPropiedadesTeclado.Visible = False


            With Me

                .PanelBusquedaPropiedadesInteresadosInmuebles.Height = 50
                .PanelBusquedaPropiedadesInteresadosInmuebles.Width = PanelBusquedasPropiedades.Size.Width - 50
                .PanelBusquedaPropiedadesInteresadosInmuebles.Location = New System.Drawing.Point((PanelBusquedasPropiedades.Size.Width - PanelBusquedasPropiedades.Size.Width) + 20, PanelBusquedasPropiedades.Size.Height - 50)

                'REDIMENCIONA LOS PANELES SEGUN EL ANCHO Y LARGO DE LA PANTALLA MAXIMIZADA
                .PanelBusquedaPropiedadesInteresadosCentral.Size = New System.Drawing.Size(PanelBusquedasPropiedades.Size.Width - PanelBusquedasPropiedades.Size.Width / 3 + 30, PanelBusquedasPropiedades.Size.Height / 8)
                .PanelBusquedaPropiedadesInteresadosCentral.Location = New System.Drawing.Point(PanelBusquedasPropiedades.Size.Width / 6, PanelBusquedasPropiedades.Size.Height / 2)
                .LabelBusquedaPropiedadesInteresadosTitulo.Location = New System.Drawing.Point(PanelBusquedasPropiedades.Size.Width / 2.6, PanelBusquedasPropiedades.Size.Height / 4)

                'Le pone de fondo al label la pick
                .LabelBusquedaPropiedadesInteresadosTitulo.Parent = PickBusquedaPropiedadesInteresados

                'Selecciona el titulo para que no lo seleccione al combo
                .LabelBusquedaPropiedadesInteresadosTitulo.Select()

            End With

            contadorCerrar = 0
            cambiartamaño = 0
            otrosFiltrosImagenes = False
            PanelBusquedaPropiedadesInteresadosInmuebles.Visible = False
            PanelBusquedaPropiedadesInteresadosCentral.Location = New System.Drawing.Point(PanelBusquedasPropiedades.Size.Width / 6, PanelBusquedasPropiedades.Size.Height / 2)
            MsgBox("10 MINUTOS DE INACTIVIDAD (ESTE TEXTBOX NO SE MOSTRARA TERMINADO EL PROYECTO)")


            '---------------------------FIN VALORES POR DEFACTO PANEL--------------------------------------------------

        End If
    End Sub

    '----------------------------------------EVENTOS----------------------------------------------------------------------------------------


    'SI SE CLICKEA EN EL PANEL INMUEBLE TAMBIEN SE RESTA 10 MINUTOS A LA INACTIVIDAD
    Private Sub PanelInmuebles_MouseClick(sender As Object, e As MouseEventArgs) Handles PanelBusquedaPropiedadesInteresadosInmuebles.MouseClick
        '----------------------DETECTAR ACTIVIDAD------------------------------------------
        INPUT.cbSize = Marshal.SizeOf(INPUT) '¿? PERO ES NECESARIO                       '---
        TimerInactividad.Interval = 1000 'MONITORIZAREMOS GetLastInputInfo CADA SEGUNDO       '---
        TimerInactividad.Start()                                                              '---
        '-----------------------------------------------------------------------------------
    End Sub







    'SELECCIONA EL SIGUIENTE COMBO SI YA SE ELEGIO UNA OPCION DEL ANTERIOR
    Private Sub ComboTipoInmueble_SelectedIndexChanged(sender As Object, e As EventArgs) Handles BoxBusquedaPropiedadesInteresadosTipoInmueble.SelectedIndexChanged

        BoxBusquedaPropiedadesInteresadosCompraVenta.Select()
    End Sub

    'EVENTOS PARA DESPLEGAR LOS COMBOS AL HACER CLICK EN EL CONTENIDO
    Private Sub ComboDepartamento_MouseClick(sender As Object, e As MouseEventArgs) Handles BoxBusquedaPropiedadesInteresadosDepartamento.MouseClick
        BoxBusquedaPropiedadesInteresadosDepartamento.DroppedDown = Enabled

    End Sub

    'EVENTOS PARA DESPLEGAR LOS COMBOS AL HACER CLICK EN EL CONTENIDO
    Private Sub ComboTipo_MouseClick(sender As Object, e As MouseEventArgs) Handles BoxBusquedaPropiedadesInteresadosTipoInmueble.MouseClick
        BoxBusquedaPropiedadesInteresadosTipoInmueble.DroppedDown = Enabled
    End Sub

    'EVENTOS PARA DESPLEGAR LOS COMBOS AL HACER CLICK EN EL CONTENIDO
    Private Sub ComboCompraVenta_MouseClick(sender As Object, e As MouseEventArgs) Handles BoxBusquedaPropiedadesInteresadosCompraVenta.MouseClick
        BoxBusquedaPropiedadesInteresadosCompraVenta.DroppedDown = Enabled
    End Sub

    'SELECCIONA EL SIGUIENTE COMBO SI YA SE ELEGIO UNA OPCION DEL ANTERIOR
    Private Sub ComboDepartamento_SelectedIndexChanged(sender As Object, e As EventArgs) Handles BoxBusquedaPropiedadesInteresadosDepartamento.SelectedIndexChanged
        BoxBusquedaPropiedadesInteresadosTipoInmueble.Select()
    End Sub


    'CODIGO PARA INGRESAR , MOSTRAR TECLADO LETRAS Y OCULTAR TECLADO NUMEROS AL ENTRAR EN TEXTBOX
    Private Sub TxtBusquedasPropiedadesUsuario_MouseEnter(sender As Object, e As EventArgs) Handles TxtBusquedasPropiedadesUsuario.MouseEnter
        PanelFormBusquedasPropiedadesPanelImagenes.Visible = False
        Me.ButtonFormBusquedaPropiedadesBORRAR.Location = New System.Drawing.Point(237, 214)

        Me.ButtonFormBusquedaPropiedadesBORRAR.Size = New System.Drawing.Size(197, 68)
        PanelFormBusquedaPropiedadesTeclado.Visible = True

        ButtonFormBusquedaPropiedades1.Visible = False
        ButtonFormBusquedaPropiedades2.Visible = False
        ButtonFormBusquedaPropiedades3.Visible = False
        ButtonFormBusquedaPropiedades4.Visible = False
        ButtonFormBusquedaPropiedades5.Visible = False
        ButtonFormBusquedaPropiedades6.Visible = False
        ButtonFormBusquedaPropiedades7.Visible = False
        ButtonFormBusquedaPropiedades8.Visible = False
        ButtonFormBusquedaPropiedades9.Visible = False
        ButtonFormBusquedaPropiedades0.Visible = False
        ButtonFormBusquedaPropiedadesA.Visible = True
        ButtonFormBusquedaPropiedadesB.Visible = True
        ButtonFormBusquedaPropiedadesC.Visible = True
        ButtonFormBusquedaPropiedadesD.Visible = True
        ButtonFormBusquedaPropiedadesE.Visible = True
        ButtonFormBusquedaPropiedadesF.Visible = True
        ButtonFormBusquedaPropiedadesG.Visible = True
        ButtonFormBusquedaPropiedadesH.Visible = True
        ButtonFormBusquedaPropiedadesI.Visible = True
        ButtonFormBusquedaPropiedadesJ.Visible = True
        ButtonFormBusquedaPropiedadesK.Visible = True
        ButtonFormBusquedaPropiedadesL.Visible = True
        ButtonFormBusquedaPropiedadesM.Visible = True
        ButtonFormBusquedaPropiedadesN.Visible = True
        ButtonFormBusquedaPropiedadesÑ.Visible = True
        ButtonFormBusquedaPropiedadesO.Visible = True
        ButtonFormBusquedaPropiedadesP.Visible = True
        ButtonFormBusquedaPropiedadesQ.Visible = True
        ButtonFormBusquedaPropiedadesR.Visible = True
        ButtonFormBusquedaPropiedadesS.Visible = True
        ButtonFormBusquedaPropiedadesT.Visible = True
        ButtonFormBusquedaPropiedadesU.Visible = True
        ButtonFormBusquedaPropiedadesV.Visible = True
        ButtonFormBusquedaPropiedadesW.Visible = True
        ButtonFormBusquedaPropiedadesX.Visible = True
        ButtonFormBusquedaPropiedadesY.Visible = True
        ButtonFormBusquedaPropiedadesZ.Visible = True


        If TxtBusquedasPropiedadesUsuario.Text = "Usuario" Then
            TxtBusquedasPropiedadesUsuario.Text = ""
        End If
    End Sub



    'INGRESAR CONTRASEÑA AL MOSTRAR TECLADO NUMERICO ENTRANDO EN TEXBOX
    Private Sub TxtBusquedasPropiedadesContraseña_MouseEnter(sender As Object, e As EventArgs) Handles TxtBusquedasPropiedadesContraseña.MouseEnter
        PanelFormBusquedasPropiedadesPanelImagenes.Visible = False
        Me.ButtonFormBusquedaPropiedadesBORRAR.Location = New System.Drawing.Point(12, 214)
        Me.ButtonFormBusquedaPropiedadesBORRAR.Size = New System.Drawing.Size(422, 68)

        ButtonFormBusquedaPropiedadesA.Visible = False
        ButtonFormBusquedaPropiedadesB.Visible = False
        ButtonFormBusquedaPropiedadesC.Visible = False
        ButtonFormBusquedaPropiedadesD.Visible = False
        ButtonFormBusquedaPropiedadesE.Visible = False
        ButtonFormBusquedaPropiedadesF.Visible = False
        ButtonFormBusquedaPropiedadesG.Visible = False
        ButtonFormBusquedaPropiedadesH.Visible = False
        ButtonFormBusquedaPropiedadesI.Visible = False
        ButtonFormBusquedaPropiedadesJ.Visible = False
        ButtonFormBusquedaPropiedadesK.Visible = False
        ButtonFormBusquedaPropiedadesL.Visible = False
        ButtonFormBusquedaPropiedadesM.Visible = False
        ButtonFormBusquedaPropiedadesN.Visible = False
        ButtonFormBusquedaPropiedadesÑ.Visible = False
        ButtonFormBusquedaPropiedadesO.Visible = False
        ButtonFormBusquedaPropiedadesP.Visible = False
        ButtonFormBusquedaPropiedadesQ.Visible = False
        ButtonFormBusquedaPropiedadesR.Visible = False
        ButtonFormBusquedaPropiedadesS.Visible = False
        ButtonFormBusquedaPropiedadesT.Visible = False
        ButtonFormBusquedaPropiedadesU.Visible = False
        ButtonFormBusquedaPropiedadesV.Visible = False
        ButtonFormBusquedaPropiedadesW.Visible = False
        ButtonFormBusquedaPropiedadesX.Visible = False
        ButtonFormBusquedaPropiedadesY.Visible = False
        ButtonFormBusquedaPropiedadesZ.Visible = False
        ButtonFormBusquedaPropiedades1.Visible = True
        ButtonFormBusquedaPropiedades2.Visible = True
        ButtonFormBusquedaPropiedades3.Visible = True
        ButtonFormBusquedaPropiedades4.Visible = True
        ButtonFormBusquedaPropiedades5.Visible = True
        ButtonFormBusquedaPropiedades6.Visible = True
        ButtonFormBusquedaPropiedades7.Visible = True
        ButtonFormBusquedaPropiedades8.Visible = True
        ButtonFormBusquedaPropiedades9.Visible = True
        ButtonFormBusquedaPropiedades0.Visible = True


        PanelFormBusquedaPropiedadesTeclado.Visible = True
        If TxtBusquedasPropiedadesContraseña.Text = "Contraseña" Then
            TxtBusquedasPropiedadesContraseña.Text = ""
        End If
    End Sub

    Private Sub TxtBusquedasPropiedadesUsuario_MouseLeave(sender As Object, e As EventArgs) Handles TxtBusquedasPropiedadesUsuario.MouseLeave

        If TxtBusquedasPropiedadesUsuario.Text = "" Then
            TxtBusquedasPropiedadesUsuario.Text = "Usuario"
        End If
    End Sub

    Private Sub TxtBusquedasPropiedadesContraseña_MouseLeave(sender As Object, e As EventArgs) Handles TxtBusquedasPropiedadesContraseña.MouseLeave
        If TxtBusquedasPropiedadesContraseña.Text = "" Then
            TxtBusquedasPropiedadesContraseña.Text = "Contraseña"
        End If
    End Sub



    'SI SE HACE CLICK 10 VECES AL TITULO SE CIERRA LA APLICACION  - FUNCIONALIDAD PARA EL AGENTE, GERENTE, ADMIN
    Private Sub LabelBusquedaPropiedadesInteresadosTitulo_MouseClick(sender As Object, e As MouseEventArgs) Handles LabelBusquedaPropiedadesInteresadosTitulo.MouseClick

        contadorCerrar += 1

        If contadorCerrar = 10 Then
            Application.Exit()
        End If

    End Sub



    'FUNCION PARA CONTAR LOS MINUTOS LUEGO DE HACER UN CLICK EN UN OBJETO
    Private Declare Function GetLastInputInfo Lib "user32" (ByRef plii As LASTINPUTINFO) As Boolean
    Structure LASTINPUTINFO
        Public cbSize As Integer 'TAMAÑO DE LA ESTRUCTURA EN BYTES ¿?
        Public dwTime As Integer 'MOMENTO (MILISEGUNDOS DESDE QUE SE INICIO LA SESION) EN QUE SE PULSA UNA TECLA O SE ACTIVA EL MOUSE
    End Structure
    Dim INPUT As New LASTINPUTINFO() 'PARA USAR LA ESTRUCTURA EN LA FUNCION GetLastInputInfo



    Private Sub PanelFormBusquedaPropiedadesTeclado_Paint(sender As Object, e As PaintEventArgs) Handles PanelFormBusquedaPropiedadesTeclado.Paint
        PanelFormBusquedaPropiedadesTeclado.Visible = True
    End Sub

    'CIERRA EL TECLADO
    Private Sub Button25_Click(sender As Object, e As EventArgs) Handles ButtonFormBusquedaPropiedadesCERRAR.Click
        PanelFormBusquedaPropiedadesTeclado.Visible = False

    End Sub


    'AGREGAR LETRAS Y NUMEROS LOS TEXBOX AL PRESIONAR SU CORRESPONDIENTE LETRA O NUMERO
    Private Sub ButtonFormBusquedaPropiedadesA_Click(sender As Object, e As EventArgs) Handles ButtonFormBusquedaPropiedadesA.Click
        If TxtBusquedasPropiedadesUsuario.Text = "Usuario" Then
            TxtBusquedasPropiedadesUsuario.Text = ""
        End If
        TxtBusquedasPropiedadesUsuario.Text = TxtBusquedasPropiedadesUsuario.Text & "A"
    End Sub
    Private Sub ButtonFormBusquedaPropiedadesB_Click(sender As Object, e As EventArgs) Handles ButtonFormBusquedaPropiedadesB.Click
        If TxtBusquedasPropiedadesUsuario.Text = "Usuario" Then
            TxtBusquedasPropiedadesUsuario.Text = ""
        End If
        TxtBusquedasPropiedadesUsuario.Text = TxtBusquedasPropiedadesUsuario.Text & "B"
    End Sub
    Private Sub ButtonFormBusquedaPropiedadesC_Click(sender As Object, e As EventArgs) Handles ButtonFormBusquedaPropiedadesC.Click
        If TxtBusquedasPropiedadesUsuario.Text = "Usuario" Then
            TxtBusquedasPropiedadesUsuario.Text = ""
        End If
        TxtBusquedasPropiedadesUsuario.Text = TxtBusquedasPropiedadesUsuario.Text & "C"
    End Sub
    Private Sub ButtonFormBusquedaPropiedadesD_Click(sender As Object, e As EventArgs) Handles ButtonFormBusquedaPropiedadesD.Click
        If TxtBusquedasPropiedadesUsuario.Text = "Usuario" Then
            TxtBusquedasPropiedadesUsuario.Text = ""
        End If
        TxtBusquedasPropiedadesUsuario.Text = TxtBusquedasPropiedadesUsuario.Text & "D"
    End Sub
    Private Sub ButtonFormBusquedaPropiedadesE_Click(sender As Object, e As EventArgs) Handles ButtonFormBusquedaPropiedadesE.Click
        If TxtBusquedasPropiedadesUsuario.Text = "Usuario" Then
            TxtBusquedasPropiedadesUsuario.Text = ""
        End If
        TxtBusquedasPropiedadesUsuario.Text = TxtBusquedasPropiedadesUsuario.Text & "E"
    End Sub
    Private Sub ButtonFormBusquedaPropiedadesF_Click(sender As Object, e As EventArgs) Handles ButtonFormBusquedaPropiedadesF.Click
        If TxtBusquedasPropiedadesUsuario.Text = "Usuario" Then
            TxtBusquedasPropiedadesUsuario.Text = ""
        End If
        TxtBusquedasPropiedadesUsuario.Text = TxtBusquedasPropiedadesUsuario.Text & "F"
    End Sub
    Private Sub ButtonFormBusquedaPropiedadesG_Click(sender As Object, e As EventArgs) Handles ButtonFormBusquedaPropiedadesG.Click
        If TxtBusquedasPropiedadesUsuario.Text = "Usuario" Then
            TxtBusquedasPropiedadesUsuario.Text = ""
        End If
        TxtBusquedasPropiedadesUsuario.Text = TxtBusquedasPropiedadesUsuario.Text & "G"
    End Sub
    Private Sub ButtonFormBusquedaPropiedadesH_Click(sender As Object, e As EventArgs) Handles ButtonFormBusquedaPropiedadesH.Click
        If TxtBusquedasPropiedadesUsuario.Text = "Usuario" Then
            TxtBusquedasPropiedadesUsuario.Text = ""
        End If
        TxtBusquedasPropiedadesUsuario.Text = TxtBusquedasPropiedadesUsuario.Text & "H"
    End Sub
    Private Sub ButtonFormBusquedaPropiedadesI_Click(sender As Object, e As EventArgs) Handles ButtonFormBusquedaPropiedadesI.Click
        If TxtBusquedasPropiedadesUsuario.Text = "Usuario" Then
            TxtBusquedasPropiedadesUsuario.Text = ""
        End If
        TxtBusquedasPropiedadesUsuario.Text = TxtBusquedasPropiedadesUsuario.Text & "I"
    End Sub
    Private Sub ButtonFormBusquedaPropiedadesJ_Click(sender As Object, e As EventArgs) Handles ButtonFormBusquedaPropiedadesJ.Click
        If TxtBusquedasPropiedadesUsuario.Text = "Usuario" Then
            TxtBusquedasPropiedadesUsuario.Text = ""
        End If
        TxtBusquedasPropiedadesUsuario.Text = TxtBusquedasPropiedadesUsuario.Text & "J"
    End Sub
    Private Sub ButtonFormBusquedaPropiedadesK_Click(sender As Object, e As EventArgs) Handles ButtonFormBusquedaPropiedadesK.Click
        If TxtBusquedasPropiedadesUsuario.Text = "Usuario" Then
            TxtBusquedasPropiedadesUsuario.Text = ""
        End If
        TxtBusquedasPropiedadesUsuario.Text = TxtBusquedasPropiedadesUsuario.Text & "K"
    End Sub
    Private Sub ButtonFormBusquedaPropiedadesL_Click(sender As Object, e As EventArgs) Handles ButtonFormBusquedaPropiedadesL.Click
        If TxtBusquedasPropiedadesUsuario.Text = "Usuario" Then
            TxtBusquedasPropiedadesUsuario.Text = ""
        End If
        TxtBusquedasPropiedadesUsuario.Text = TxtBusquedasPropiedadesUsuario.Text & "L"
    End Sub
    Private Sub ButtonFormBusquedaPropiedadesM_Click(sender As Object, e As EventArgs) Handles ButtonFormBusquedaPropiedadesM.Click
        If TxtBusquedasPropiedadesUsuario.Text = "Usuario" Then
            TxtBusquedasPropiedadesUsuario.Text = ""
        End If
        TxtBusquedasPropiedadesUsuario.Text = TxtBusquedasPropiedadesUsuario.Text & "M"
    End Sub
    Private Sub ButtonFormBusquedaPropiedadesN_Click(sender As Object, e As EventArgs) Handles ButtonFormBusquedaPropiedadesN.Click
        If TxtBusquedasPropiedadesUsuario.Text = "Usuario" Then
            TxtBusquedasPropiedadesUsuario.Text = ""
        End If
        TxtBusquedasPropiedadesUsuario.Text = TxtBusquedasPropiedadesUsuario.Text & "N"
    End Sub
    Private Sub ButtonFormBusquedaPropiedadesÑ_Click(sender As Object, e As EventArgs) Handles ButtonFormBusquedaPropiedadesÑ.Click
        If TxtBusquedasPropiedadesUsuario.Text = "Usuario" Then
            TxtBusquedasPropiedadesUsuario.Text = ""
        End If
        TxtBusquedasPropiedadesUsuario.Text = TxtBusquedasPropiedadesUsuario.Text & "Ñ"
    End Sub
    Private Sub ButtonFormBusquedaPropiedadesO_Click(sender As Object, e As EventArgs) Handles ButtonFormBusquedaPropiedadesO.Click
        If TxtBusquedasPropiedadesUsuario.Text = "Usuario" Then
            TxtBusquedasPropiedadesUsuario.Text = ""
        End If
        TxtBusquedasPropiedadesUsuario.Text = TxtBusquedasPropiedadesUsuario.Text & "O"
    End Sub
    Private Sub ButtonFormBusquedaPropiedadesP_Click(sender As Object, e As EventArgs) Handles ButtonFormBusquedaPropiedadesP.Click
        If TxtBusquedasPropiedadesUsuario.Text = "Usuario" Then
            TxtBusquedasPropiedadesUsuario.Text = ""
        End If
        TxtBusquedasPropiedadesUsuario.Text = TxtBusquedasPropiedadesUsuario.Text & "P"
    End Sub
    Private Sub ButtonFormBusquedaPropiedadesQ_Click(sender As Object, e As EventArgs) Handles ButtonFormBusquedaPropiedadesQ.Click
        If TxtBusquedasPropiedadesUsuario.Text = "Usuario" Then
            TxtBusquedasPropiedadesUsuario.Text = ""
        End If
        TxtBusquedasPropiedadesUsuario.Text = TxtBusquedasPropiedadesUsuario.Text & "Q"
    End Sub
    Private Sub ButtonFormBusquedaPropiedadesR_Click(sender As Object, e As EventArgs) Handles ButtonFormBusquedaPropiedadesR.Click
        If TxtBusquedasPropiedadesUsuario.Text = "Usuario" Then
            TxtBusquedasPropiedadesUsuario.Text = ""
        End If
        TxtBusquedasPropiedadesUsuario.Text = TxtBusquedasPropiedadesUsuario.Text & "R"
    End Sub
    Private Sub ButtonFormBusquedaPropiedadesS_Click(sender As Object, e As EventArgs) Handles ButtonFormBusquedaPropiedadesS.Click
        If TxtBusquedasPropiedadesUsuario.Text = "Usuario" Then
            TxtBusquedasPropiedadesUsuario.Text = ""
        End If
        TxtBusquedasPropiedadesUsuario.Text = TxtBusquedasPropiedadesUsuario.Text & "S"
    End Sub
    Private Sub ButtonFormBusquedaPropiedadesT_Click(sender As Object, e As EventArgs) Handles ButtonFormBusquedaPropiedadesT.Click
        If TxtBusquedasPropiedadesUsuario.Text = "Usuario" Then
            TxtBusquedasPropiedadesUsuario.Text = ""
        End If
        TxtBusquedasPropiedadesUsuario.Text = TxtBusquedasPropiedadesUsuario.Text & "T"
    End Sub
    Private Sub ButtonFormBusquedaPropiedadesU_Click(sender As Object, e As EventArgs) Handles ButtonFormBusquedaPropiedadesU.Click
        If TxtBusquedasPropiedadesUsuario.Text = "Usuario" Then
            TxtBusquedasPropiedadesUsuario.Text = ""
        End If
        TxtBusquedasPropiedadesUsuario.Text = TxtBusquedasPropiedadesUsuario.Text & "U"
    End Sub
    Private Sub ButtonFormBusquedaPropiedadesV_Click(sender As Object, e As EventArgs) Handles ButtonFormBusquedaPropiedadesV.Click
        If TxtBusquedasPropiedadesUsuario.Text = "Usuario" Then
            TxtBusquedasPropiedadesUsuario.Text = ""
        End If
        TxtBusquedasPropiedadesUsuario.Text = TxtBusquedasPropiedadesUsuario.Text & "V"
    End Sub
    Private Sub ButtonFormBusquedaPropiedadesW_Click(sender As Object, e As EventArgs) Handles ButtonFormBusquedaPropiedadesW.Click
        If TxtBusquedasPropiedadesUsuario.Text = "Usuario" Then
            TxtBusquedasPropiedadesUsuario.Text = ""
        End If
        TxtBusquedasPropiedadesUsuario.Text = TxtBusquedasPropiedadesUsuario.Text & "W"
    End Sub
    Private Sub ButtonFormBusquedaPropiedadesX_Click(sender As Object, e As EventArgs) Handles ButtonFormBusquedaPropiedadesX.Click
        If TxtBusquedasPropiedadesUsuario.Text = "Usuario" Then
            TxtBusquedasPropiedadesUsuario.Text = ""
        End If
        TxtBusquedasPropiedadesUsuario.Text = TxtBusquedasPropiedadesUsuario.Text & "X"
    End Sub
    Private Sub ButtonFormBusquedaPropiedadesZ_Click(sender As Object, e As EventArgs) Handles ButtonFormBusquedaPropiedadesZ.Click
        If TxtBusquedasPropiedadesUsuario.Text = "Usuario" Then
            TxtBusquedasPropiedadesUsuario.Text = ""
        End If
        TxtBusquedasPropiedadesUsuario.Text = TxtBusquedasPropiedadesUsuario.Text & "Z"
    End Sub
    Private Sub ButtonFormBusquedaPropiedadesBORRAR_Click(sender As Object, e As EventArgs) Handles ButtonFormBusquedaPropiedadesBORRAR.Click
        If ButtonFormBusquedaPropiedadesA.Visible = True Then
            TxtBusquedasPropiedadesUsuario.Text = ""
        Else
            TxtBusquedasPropiedadesContraseña.Text = ""
        End If
    End Sub
    Private Sub ButtonFormBusquedaPropiedades1_Click(sender As Object, e As EventArgs) Handles ButtonFormBusquedaPropiedades1.Click
        If TxtBusquedasPropiedadesContraseña.Text = "Contraseña" Then
            TxtBusquedasPropiedadesContraseña.Text = ""
        End If
        TxtBusquedasPropiedadesContraseña.Text = TxtBusquedasPropiedadesContraseña.Text & "1"
    End Sub
    Private Sub ButtonFormBusquedaPropiedades2_Click(sender As Object, e As EventArgs) Handles ButtonFormBusquedaPropiedades2.Click
        If TxtBusquedasPropiedadesContraseña.Text = "Contraseña" Then
            TxtBusquedasPropiedadesContraseña.Text = ""
        End If
        TxtBusquedasPropiedadesContraseña.Text = TxtBusquedasPropiedadesContraseña.Text & "2"
    End Sub
    Private Sub ButtonFormBusquedaPropiedades3_Click(sender As Object, e As EventArgs) Handles ButtonFormBusquedaPropiedades3.Click
        If TxtBusquedasPropiedadesContraseña.Text = "Contraseña" Then
            TxtBusquedasPropiedadesContraseña.Text = ""
        End If
        TxtBusquedasPropiedadesContraseña.Text = TxtBusquedasPropiedadesContraseña.Text & "3"
    End Sub
    Private Sub ButtonFormBusquedaPropiedades4_Click(sender As Object, e As EventArgs) Handles ButtonFormBusquedaPropiedades4.Click
        If TxtBusquedasPropiedadesContraseña.Text = "Contraseña" Then
            TxtBusquedasPropiedadesContraseña.Text = ""
        End If
        TxtBusquedasPropiedadesContraseña.Text = TxtBusquedasPropiedadesContraseña.Text & "4"
    End Sub
    Private Sub ButtonFormBusquedaPropiedades5_Click(sender As Object, e As EventArgs) Handles ButtonFormBusquedaPropiedades5.Click
        If TxtBusquedasPropiedadesContraseña.Text = "Contraseña" Then
            TxtBusquedasPropiedadesContraseña.Text = ""
        End If
        TxtBusquedasPropiedadesContraseña.Text = TxtBusquedasPropiedadesContraseña.Text & "5"
    End Sub
    Private Sub ButtonFormBusquedaPropiedades6_Click(sender As Object, e As EventArgs) Handles ButtonFormBusquedaPropiedades6.Click
        If TxtBusquedasPropiedadesContraseña.Text = "Contraseña" Then
            TxtBusquedasPropiedadesContraseña.Text = ""
        End If
        TxtBusquedasPropiedadesContraseña.Text = TxtBusquedasPropiedadesContraseña.Text & "6"
    End Sub
    Private Sub ButtonFormBusquedaPropiedades7_Click(sender As Object, e As EventArgs) Handles ButtonFormBusquedaPropiedades7.Click
        If TxtBusquedasPropiedadesContraseña.Text = "Contraseña" Then
            TxtBusquedasPropiedadesContraseña.Text = ""
        End If
        TxtBusquedasPropiedadesContraseña.Text = TxtBusquedasPropiedadesContraseña.Text & "7"
    End Sub
    Private Sub ButtonFormBusquedaPropiedades8_Click(sender As Object, e As EventArgs) Handles ButtonFormBusquedaPropiedades8.Click
        If TxtBusquedasPropiedadesContraseña.Text = "Contraseña" Then
            TxtBusquedasPropiedadesContraseña.Text = ""
        End If
        TxtBusquedasPropiedadesContraseña.Text = TxtBusquedasPropiedadesContraseña.Text & "8"
    End Sub
    Private Sub ButtonFormBusquedaPropiedades9_Click(sender As Object, e As EventArgs) Handles ButtonFormBusquedaPropiedades9.Click
        If TxtBusquedasPropiedadesContraseña.Text = "Contraseña" Then
            TxtBusquedasPropiedadesContraseña.Text = ""
        End If
        TxtBusquedasPropiedadesContraseña.Text = TxtBusquedasPropiedadesContraseña.Text & "9"
    End Sub
    Private Sub ButtonFormBusquedaPropiedades0_Click(sender As Object, e As EventArgs) Handles ButtonFormBusquedaPropiedades0.Click
        If TxtBusquedasPropiedadesContraseña.Text = "Contraseña" Then
            TxtBusquedasPropiedadesContraseña.Text = ""
        End If
        TxtBusquedasPropiedadesContraseña.Text = TxtBusquedasPropiedadesContraseña.Text & "0"
    End Sub




    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles ButtonlFormBusquedasPropiedadesPanelImagenesCerrar.Click
        PanelFormBusquedasPropiedadesPanelImagenes.Visible = False

    End Sub


    'SIGUIENTE
    Private Sub ButtonlFormBusquedasPropiedadesPanelImagenesSiguiente_Click(sender As Object, e As EventArgs) Handles ButtonlFormBusquedasPropiedadesPanelImagenesSiguiente.Click



        'Calcula la cantidad de imagenes que tiene el inmueble
        Dim cantidadimagenes As String = consultaSQL("Select count(Imagen) from img_inmuebles where IdInmueble=" & DataBuscarPropiedadesImagenes.CurrentRow.Cells("IdInmueble").Value & " ").Tables("img_inmuebles").Rows(0)("count(Imagen)").ToString()
            Dim cantidadimagenesInt As Integer = CInt(cantidadimagenes)


            If cantidadimg <= cantidadimagenesInt - 2 Then

                cantidadimg = cantidadimg + 1

                mostrarimagen(cantidadimg)
            mostrarimagenfull(cantidadimg)

        Else
                mostrarimagen(0)
                cantidadimg = 0
            End If














    End Sub



    'ANTERIOR
    Private Sub ButtonlFormBusquedasPropiedadesPanelImagenesAnterior_Click(sender As Object, e As EventArgs) Handles ButtonlFormBusquedasPropiedadesPanelImagenesAnterior.Click


        Dim cantidadimagenes As String = consultaSQL("Select count(Imagen) from img_inmuebles where IdInmueble=" & DataBuscarPropiedadesImagenes.CurrentRow.Cells("IdInmueble").Value & " ").Tables("img_inmuebles").Rows(0)("count(Imagen)").ToString()
            Dim cantidadimagenesInt As Integer = CInt(cantidadimagenes)



        If cantidadimg > 0 Then
            cantidadimg = cantidadimg - 1

            mostrarimagen(cantidadimg)
            mostrarimagenfull(cantidadimg)


        Else

            cantidadimg = cantidadimagenesInt - 1
            mostrarimagen(cantidadimg)


        End If













    End Sub





    'MUESTRA LAS IMAGENES QUE TIENE UN INMUEBLE
    Public Sub mostrarimagen(ByVal numimagen As Integer)


        Dim sql As String = "Select Imagen from img_inmuebles where IdInmueble=" & DataBuscarPropiedadesImagenes.CurrentRow.Cells("IdInmueble").Value & " "
        conectar()
        Dim adaptador As New MySqlDataAdapter(sql, conexion)
        Dim tabla As New DataTable
        adaptador.Fill(tabla)
        Dim imgByte() As Byte



        Try
            imgByte = tabla(numimagen)(0)
            Dim ms As New MemoryStream(imgByte)

            PictureFormBusquedasPropiedadesPanelImagenes.Image = Image.FromStream(ms)
        Catch ex As Exception

            MsgBox("No hay imagenes")

        End Try





    End Sub




    'MUESTRA LAS IMAGENES QUE TIENE UN INMUEBLE
    Public Sub mostrarimagenfull(ByVal numimagen As Integer)


        Dim sql As String = "Select Imagen from img_inmuebles where IdInmueble=" & DataBuscarPropiedadesImagenes.CurrentRow.Cells("IdInmueble").Value & " "
        conectar()
        Dim adaptador As New MySqlDataAdapter(sql, conexion)
        Dim tabla As New DataTable
        adaptador.Fill(tabla)
        Dim imgByte() As Byte



        Try
            imgByte = tabla(numimagen)(0)
            Dim ms As New MemoryStream(imgByte)

            imagenesFullPantalla.PicImagenesFullPantalla.Image = Image.FromStream(ms)
        Catch ex As Exception

            MsgBox("No hay imagenes")

        End Try





    End Sub






    Private Sub CheckBusquedasPropiedadesBarrio1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBusquedasPropiedadesBarrio1.CheckedChanged
        ButtonlFormBusquedasPropiedadesPanelImagenesCerrar.PerformClick()

        If CheckBusquedasPropiedadesBarrio1.Checked = True Then

            barrio1 = CheckBusquedasPropiedadesBarrio1.Text


            DataBuscarPropiedadesImagenes.DataSource = consultaSQL("SELECT IdInmueble,	Barrios, NumDormitorios, NumBaños , Jardin , Amoblada , Garage , barbacoa , Moneda, Precio  From Inmueble, Ubicacion, caracteristicas_inm, publicacion , barrios
where ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio And inmueble.IdInm = caracteristicas_inm.idInm And Inmueble.IdPubli = publicacion.IdPubli And
	ubicacion.Departamento = '" & BoxBusquedaPropiedadesInteresadosDepartamento.SelectedItem & "' and caracteristicas_inm.TipoInm = '" & BoxBusquedaPropiedadesInteresadosTipoInmueble.SelectedItem & "' and publicacion.TipoPubli = '" & BoxBusquedaPropiedadesInteresadosCompraVenta.SelectedItem & "'

and ( barrios.Barrios = '" & barrio1 & "'  or barrios.Barrios = '" & barrio2 & "' or barrios.Barrios = '" & barrio3 & "' or barrios.Barrios = '" & barrio4 & "' 
or barrios.Barrios = '" & barrio5 & "' or barrios.Barrios = '" & barrio6 & "' or barrios.Barrios = '" & barrio7 & "')


;").Tables("consultainmueble")

            For i = 0 To DataBuscarPropiedadesImagenes.Rows.Count - 1

                Dim r As DataGridViewRow = DataBuscarPropiedadesImagenes.Rows(i)
                r.Height = 100


            Next

        End If




        If CheckBusquedasPropiedadesBarrio1.Checked = False Then

            barrio1 = ""

            DataBuscarPropiedadesImagenes.DataSource = consultaSQL("SELECT IdInmueble,	Barrios, NumDormitorios, NumBaños , Jardin , Amoblada , Garage , barbacoa , Moneda, Precio  From Inmueble, Ubicacion, caracteristicas_inm, publicacion , barrios
where ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio And inmueble.IdInm = caracteristicas_inm.idInm And Inmueble.IdPubli = publicacion.IdPubli And
	ubicacion.Departamento = '" & BoxBusquedaPropiedadesInteresadosDepartamento.SelectedItem & "' and caracteristicas_inm.TipoInm = '" & BoxBusquedaPropiedadesInteresadosTipoInmueble.SelectedItem & "' and publicacion.TipoPubli = '" & BoxBusquedaPropiedadesInteresadosCompraVenta.SelectedItem & "'

and ( barrios.Barrios = '" & barrio1 & "'  or barrios.Barrios = '" & barrio2 & "' or barrios.Barrios = '" & barrio3 & "' or barrios.Barrios = '" & barrio4 & "' 
or barrios.Barrios = '" & barrio5 & "' or barrios.Barrios = '" & barrio6 & "' or barrios.Barrios = '" & barrio7 & "')


;").Tables("consultainmueble")

            For i = 0 To DataBuscarPropiedadesImagenes.Rows.Count - 1

                Dim r As DataGridViewRow = DataBuscarPropiedadesImagenes.Rows(i)
                r.Height = 100


            Next

        End If




        If barrio1 = "" And barrio2 = "" And barrio3 = "" And barrio4 = "" And barrio5 = "" And barrio6 = "" And barrio7 = "" Then

            DataBuscarPropiedadesImagenes.DataSource = consultaSQL("SELECT IdInmueble,	Barrios, NumDormitorios, NumBaños , Jardin , Amoblada , Garage , barbacoa , Moneda, Precio  From Inmueble, Ubicacion, caracteristicas_inm, publicacion , barrios
where ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio And inmueble.IdInm = caracteristicas_inm.idInm And Inmueble.IdPubli = publicacion.IdPubli And
	ubicacion.Departamento = '" & BoxBusquedaPropiedadesInteresadosDepartamento.SelectedItem & "' and caracteristicas_inm.TipoInm = '" & BoxBusquedaPropiedadesInteresadosTipoInmueble.SelectedItem & "' and publicacion.TipoPubli = '" & BoxBusquedaPropiedadesInteresadosCompraVenta.SelectedItem & "'

;").Tables("consultainmueble")

            For i = 0 To DataBuscarPropiedadesImagenes.Rows.Count - 1

                Dim r As DataGridViewRow = DataBuscarPropiedadesImagenes.Rows(i)
                r.Height = 100


            Next

        End If





    End Sub

    Private Sub CheckBusquedasPropiedadesBarrio2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBusquedasPropiedadesBarrio2.CheckedChanged
        ButtonlFormBusquedasPropiedadesPanelImagenesCerrar.PerformClick()


        If CheckBusquedasPropiedadesBarrio2.Checked = True Then

            barrio2 = CheckBusquedasPropiedadesBarrio2.Text


            DataBuscarPropiedadesImagenes.DataSource = consultaSQL("SELECT IdInmueble,	Barrios, NumDormitorios, NumBaños , Jardin , Amoblada , Garage , barbacoa , Moneda, Precio  From Inmueble, Ubicacion, caracteristicas_inm, publicacion , barrios
where ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio And inmueble.IdInm = caracteristicas_inm.idInm And Inmueble.IdPubli = publicacion.IdPubli And
	ubicacion.Departamento = '" & BoxBusquedaPropiedadesInteresadosDepartamento.SelectedItem & "' and caracteristicas_inm.TipoInm = '" & BoxBusquedaPropiedadesInteresadosTipoInmueble.SelectedItem & "' and publicacion.TipoPubli = '" & BoxBusquedaPropiedadesInteresadosCompraVenta.SelectedItem & "'
and ( barrios.Barrios = '" & barrio1 & "'  or barrios.Barrios = '" & barrio2 & "'  or barrios.Barrios = '" & barrio3 & "' or barrios.Barrios = '" & barrio4 & "' 
or barrios.Barrios = '" & barrio5 & "' or barrios.Barrios = '" & barrio6 & "' or barrios.Barrios = '" & barrio7 & "'        

)

;").Tables("consultainmueble")

            For i = 0 To DataBuscarPropiedadesImagenes.Rows.Count - 1

                Dim r As DataGridViewRow = DataBuscarPropiedadesImagenes.Rows(i)
                r.Height = 100


            Next



        End If



        If CheckBusquedasPropiedadesBarrio2.Checked = False Then

            barrio2 = ""

            DataBuscarPropiedadesImagenes.DataSource = consultaSQL("SELECT IdInmueble,	Barrios, NumDormitorios, NumBaños , Jardin , Amoblada , Garage , barbacoa , Moneda, Precio  From Inmueble, Ubicacion, caracteristicas_inm, publicacion , barrios
where ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio And inmueble.IdInm = caracteristicas_inm.idInm And Inmueble.IdPubli = publicacion.IdPubli And
	ubicacion.Departamento = '" & BoxBusquedaPropiedadesInteresadosDepartamento.SelectedItem & "' and caracteristicas_inm.TipoInm = '" & BoxBusquedaPropiedadesInteresadosTipoInmueble.SelectedItem & "' and publicacion.TipoPubli = '" & BoxBusquedaPropiedadesInteresadosCompraVenta.SelectedItem & "'
and ( barrios.Barrios = '" & barrio1 & "'  or barrios.Barrios = '" & barrio2 & "'  or barrios.Barrios = '" & barrio3 & "' or barrios.Barrios = '" & barrio4 & "' 
or barrios.Barrios = '" & barrio5 & "' or barrios.Barrios = '" & barrio6 & "' or barrios.Barrios = '" & barrio7 & "'        

)

;").Tables("consultainmueble")

            For i = 0 To DataBuscarPropiedadesImagenes.Rows.Count - 1

                Dim r As DataGridViewRow = DataBuscarPropiedadesImagenes.Rows(i)
                r.Height = 100


            Next


        End If




        If barrio1 = "" And barrio2 = "" And barrio3 = "" And barrio4 = "" And barrio5 = "" And barrio6 = "" And barrio7 = "" Then

            DataBuscarPropiedadesImagenes.DataSource = consultaSQL("SELECT IdInmueble,	Barrios, NumDormitorios, NumBaños , Jardin , Amoblada , Garage , barbacoa , Moneda, Precio  From Inmueble, Ubicacion, caracteristicas_inm, publicacion , barrios
where ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio And inmueble.IdInm = caracteristicas_inm.idInm And Inmueble.IdPubli = publicacion.IdPubli And
	ubicacion.Departamento = '" & BoxBusquedaPropiedadesInteresadosDepartamento.SelectedItem & "' and caracteristicas_inm.TipoInm = '" & BoxBusquedaPropiedadesInteresadosTipoInmueble.SelectedItem & "' and publicacion.TipoPubli = '" & BoxBusquedaPropiedadesInteresadosCompraVenta.SelectedItem & "'

;").Tables("consultainmueble")

            For i = 0 To DataBuscarPropiedadesImagenes.Rows.Count - 1

                Dim r As DataGridViewRow = DataBuscarPropiedadesImagenes.Rows(i)
                r.Height = 100


            Next

        End If




    End Sub

    Private Sub CheckBusquedasPropiedadesBarrio3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBusquedasPropiedadesBarrio3.CheckedChanged
        ButtonlFormBusquedasPropiedadesPanelImagenesCerrar.PerformClick()

        If CheckBusquedasPropiedadesBarrio3.Checked = True Then

            barrio3 = CheckBusquedasPropiedadesBarrio3.Text


            DataBuscarPropiedadesImagenes.DataSource = consultaSQL("SELECT IdInmueble,	Barrios, NumDormitorios, NumBaños , Jardin , Amoblada , Garage , barbacoa , Moneda, Precio  From Inmueble, Ubicacion, caracteristicas_inm, publicacion , barrios
where ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio And inmueble.IdInm = caracteristicas_inm.idInm And Inmueble.IdPubli = publicacion.IdPubli And
	ubicacion.Departamento = '" & BoxBusquedaPropiedadesInteresadosDepartamento.SelectedItem & "' and caracteristicas_inm.TipoInm = '" & BoxBusquedaPropiedadesInteresadosTipoInmueble.SelectedItem & "' and publicacion.TipoPubli = '" & BoxBusquedaPropiedadesInteresadosCompraVenta.SelectedItem & "'
and ( barrios.Barrios = '" & barrio1 & "'  or barrios.Barrios = '" & barrio2 & "'  or barrios.Barrios = '" & barrio3 & "' or barrios.Barrios = '" & barrio4 & "' 
or barrios.Barrios = '" & barrio5 & "' or barrios.Barrios = '" & barrio6 & "' or barrios.Barrios = '" & barrio7 & "'       

)

;").Tables("consultainmueble")

            For i = 0 To DataBuscarPropiedadesImagenes.Rows.Count - 1

                Dim r As DataGridViewRow = DataBuscarPropiedadesImagenes.Rows(i)
                r.Height = 100


            Next


        End If


        If CheckBusquedasPropiedadesBarrio3.Checked = False Then
            barrio3 = ""

            DataBuscarPropiedadesImagenes.DataSource = consultaSQL("SELECT IdInmueble,	Barrios, NumDormitorios, NumBaños , Jardin , Amoblada , Garage , barbacoa , Moneda, Precio  From Inmueble, Ubicacion, caracteristicas_inm, publicacion , barrios
where ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio And inmueble.IdInm = caracteristicas_inm.idInm And Inmueble.IdPubli = publicacion.IdPubli And
	ubicacion.Departamento = '" & BoxBusquedaPropiedadesInteresadosDepartamento.SelectedItem & "' and caracteristicas_inm.TipoInm = '" & BoxBusquedaPropiedadesInteresadosTipoInmueble.SelectedItem & "' and publicacion.TipoPubli = '" & BoxBusquedaPropiedadesInteresadosCompraVenta.SelectedItem & "'
and ( barrios.Barrios = '" & barrio1 & "'  or barrios.Barrios = '" & barrio2 & "'  or barrios.Barrios = '" & barrio3 & "' or barrios.Barrios = '" & barrio4 & "' 
or barrios.Barrios = '" & barrio5 & "' or barrios.Barrios = '" & barrio6 & "' or barrios.Barrios = '" & barrio7 & "'       

)

;").Tables("consultainmueble")

            For i = 0 To DataBuscarPropiedadesImagenes.Rows.Count - 1

                Dim r As DataGridViewRow = DataBuscarPropiedadesImagenes.Rows(i)
                r.Height = 100


            Next


        End If



        If barrio1 = "" And barrio2 = "" And barrio3 = "" And barrio4 = "" And barrio5 = "" And barrio6 = "" And barrio7 = "" Then

            DataBuscarPropiedadesImagenes.DataSource = consultaSQL("SELECT IdInmueble,	Barrios, NumDormitorios, NumBaños , Jardin , Amoblada , Garage , barbacoa , Moneda, Precio  From Inmueble, Ubicacion, caracteristicas_inm, publicacion , barrios
where ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio And inmueble.IdInm = caracteristicas_inm.idInm And Inmueble.IdPubli = publicacion.IdPubli And
	ubicacion.Departamento = '" & BoxBusquedaPropiedadesInteresadosDepartamento.SelectedItem & "' and caracteristicas_inm.TipoInm = '" & BoxBusquedaPropiedadesInteresadosTipoInmueble.SelectedItem & "' and publicacion.TipoPubli = '" & BoxBusquedaPropiedadesInteresadosCompraVenta.SelectedItem & "'

;").Tables("consultainmueble")

            For i = 0 To DataBuscarPropiedadesImagenes.Rows.Count - 1

                Dim r As DataGridViewRow = DataBuscarPropiedadesImagenes.Rows(i)
                r.Height = 100


            Next

        End If




    End Sub

    Private Sub CheckBusquedasPropiedadesBarrio4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBusquedasPropiedadesBarrio4.CheckedChanged
        ButtonlFormBusquedasPropiedadesPanelImagenesCerrar.PerformClick()


        If CheckBusquedasPropiedadesBarrio4.Checked = True Then

            barrio4 = CheckBusquedasPropiedadesBarrio4.Text


            DataBuscarPropiedadesImagenes.DataSource = consultaSQL("SELECT IdInmueble,	Barrios, NumDormitorios, NumBaños , Jardin , Amoblada , Garage , barbacoa , Moneda, Precio  From Inmueble, Ubicacion, caracteristicas_inm, publicacion , barrios
where ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio And inmueble.IdInm = caracteristicas_inm.idInm And Inmueble.IdPubli = publicacion.IdPubli And
	ubicacion.Departamento = '" & BoxBusquedaPropiedadesInteresadosDepartamento.SelectedItem & "' and caracteristicas_inm.TipoInm = '" & BoxBusquedaPropiedadesInteresadosTipoInmueble.SelectedItem & "' and publicacion.TipoPubli = '" & BoxBusquedaPropiedadesInteresadosCompraVenta.SelectedItem & "'
and ( barrios.Barrios = '" & barrio1 & "'  or barrios.Barrios = '" & barrio2 & "'  or barrios.Barrios = '" & barrio3 & "' or barrios.Barrios = '" & barrio4 & "' 
or barrios.Barrios = '" & barrio5 & "' or barrios.Barrios = '" & barrio6 & "' or barrios.Barrios = '" & barrio7 & "'       

)

;").Tables("consultainmueble")

            For i = 0 To DataBuscarPropiedadesImagenes.Rows.Count - 1

                Dim r As DataGridViewRow = DataBuscarPropiedadesImagenes.Rows(i)
                r.Height = 100


            Next


        End If


        If CheckBusquedasPropiedadesBarrio4.Checked = False Then
            barrio4 = ""

            DataBuscarPropiedadesImagenes.DataSource = consultaSQL("SELECT IdInmueble,	Barrios, NumDormitorios, NumBaños , Jardin , Amoblada , Garage , barbacoa , Moneda, Precio  From Inmueble, Ubicacion, caracteristicas_inm, publicacion , barrios
where ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio And inmueble.IdInm = caracteristicas_inm.idInm And Inmueble.IdPubli = publicacion.IdPubli And
	ubicacion.Departamento = '" & BoxBusquedaPropiedadesInteresadosDepartamento.SelectedItem & "' and caracteristicas_inm.TipoInm = '" & BoxBusquedaPropiedadesInteresadosTipoInmueble.SelectedItem & "' and publicacion.TipoPubli = '" & BoxBusquedaPropiedadesInteresadosCompraVenta.SelectedItem & "'
and ( barrios.Barrios = '" & barrio1 & "'  or barrios.Barrios = '" & barrio2 & "'  or barrios.Barrios = '" & barrio3 & "' or barrios.Barrios = '" & barrio4 & "' 
or barrios.Barrios = '" & barrio5 & "' or barrios.Barrios = '" & barrio6 & "' or barrios.Barrios = '" & barrio7 & "'       

)

;").Tables("consultainmueble")

            For i = 0 To DataBuscarPropiedadesImagenes.Rows.Count - 1

                Dim r As DataGridViewRow = DataBuscarPropiedadesImagenes.Rows(i)
                r.Height = 100


            Next

        End If


        If barrio1 = "" And barrio2 = "" And barrio3 = "" And barrio4 = "" And barrio5 = "" And barrio6 = "" And barrio7 = "" Then

            DataBuscarPropiedadesImagenes.DataSource = consultaSQL("SELECT IdInmueble,	Barrios, NumDormitorios, NumBaños , Jardin , Amoblada , Garage , barbacoa , Moneda, Precio  From Inmueble, Ubicacion, caracteristicas_inm, publicacion , barrios
where ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio And inmueble.IdInm = caracteristicas_inm.idInm And Inmueble.IdPubli = publicacion.IdPubli And
	ubicacion.Departamento = '" & BoxBusquedaPropiedadesInteresadosDepartamento.SelectedItem & "' and caracteristicas_inm.TipoInm = '" & BoxBusquedaPropiedadesInteresadosTipoInmueble.SelectedItem & "' and publicacion.TipoPubli = '" & BoxBusquedaPropiedadesInteresadosCompraVenta.SelectedItem & "'

;").Tables("consultainmueble")

            For i = 0 To DataBuscarPropiedadesImagenes.Rows.Count - 1

                Dim r As DataGridViewRow = DataBuscarPropiedadesImagenes.Rows(i)
                r.Height = 100


            Next

        End If





    End Sub

    Private Sub CheckBusquedasPropiedadesBarrio5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBusquedasPropiedadesBarrio5.CheckedChanged
        ButtonlFormBusquedasPropiedadesPanelImagenesCerrar.PerformClick()



        If CheckBusquedasPropiedadesBarrio5.Checked = True Then

            barrio5 = CheckBusquedasPropiedadesBarrio5.Text



            DataBuscarPropiedadesImagenes.DataSource = consultaSQL("SELECT IdInmueble,	Barrios, NumDormitorios, NumBaños , Jardin , Amoblada , Garage , barbacoa , Moneda, Precio  From Inmueble, Ubicacion, caracteristicas_inm, publicacion , barrios
where ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio And inmueble.IdInm = caracteristicas_inm.idInm And Inmueble.IdPubli = publicacion.IdPubli And
	ubicacion.Departamento = '" & BoxBusquedaPropiedadesInteresadosDepartamento.SelectedItem & "' and caracteristicas_inm.TipoInm = '" & BoxBusquedaPropiedadesInteresadosTipoInmueble.SelectedItem & "' and publicacion.TipoPubli = '" & BoxBusquedaPropiedadesInteresadosCompraVenta.SelectedItem & "'
and ( barrios.Barrios = '" & barrio1 & "'  or barrios.Barrios = '" & barrio2 & "'  or barrios.Barrios = '" & barrio3 & "' or barrios.Barrios = '" & barrio4 & "' 
or barrios.Barrios = '" & barrio5 & "' or barrios.Barrios = '" & barrio6 & "' or barrios.Barrios = '" & barrio7 & "'        

)

;").Tables("consultainmueble")

            For i = 0 To DataBuscarPropiedadesImagenes.Rows.Count - 1

                Dim r As DataGridViewRow = DataBuscarPropiedadesImagenes.Rows(i)
                r.Height = 100


            Next


        End If



        If CheckBusquedasPropiedadesBarrio5.Checked = False Then
            barrio5 = ""

            DataBuscarPropiedadesImagenes.DataSource = consultaSQL("SELECT IdInmueble,	Barrios, NumDormitorios, NumBaños , Jardin , Amoblada , Garage , barbacoa , Moneda, Precio  From Inmueble, Ubicacion, caracteristicas_inm, publicacion , barrios
where ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio And inmueble.IdInm = caracteristicas_inm.idInm And Inmueble.IdPubli = publicacion.IdPubli And
	ubicacion.Departamento = '" & BoxBusquedaPropiedadesInteresadosDepartamento.SelectedItem & "' and caracteristicas_inm.TipoInm = '" & BoxBusquedaPropiedadesInteresadosTipoInmueble.SelectedItem & "' and publicacion.TipoPubli = '" & BoxBusquedaPropiedadesInteresadosCompraVenta.SelectedItem & "'
and ( barrios.Barrios = '" & barrio1 & "'  or barrios.Barrios = '" & barrio2 & "'  or barrios.Barrios = '" & barrio3 & "' or barrios.Barrios = '" & barrio4 & "' 
or barrios.Barrios = '" & barrio5 & "' or barrios.Barrios = '" & barrio6 & "' or barrios.Barrios = '" & barrio7 & "'        

)

;").Tables("consultainmueble")

            For i = 0 To DataBuscarPropiedadesImagenes.Rows.Count - 1

                Dim r As DataGridViewRow = DataBuscarPropiedadesImagenes.Rows(i)
                r.Height = 100


            Next

        End If



        If barrio1 = "" And barrio2 = "" And barrio3 = "" And barrio4 = "" And barrio5 = "" And barrio6 = "" And barrio7 = "" Then

            DataBuscarPropiedadesImagenes.DataSource = consultaSQL("SELECT IdInmueble,	Barrios, NumDormitorios, NumBaños , Jardin , Amoblada , Garage , barbacoa , Moneda, Precio  From Inmueble, Ubicacion, caracteristicas_inm, publicacion , barrios
where ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio And inmueble.IdInm = caracteristicas_inm.idInm And Inmueble.IdPubli = publicacion.IdPubli And
	ubicacion.Departamento = '" & BoxBusquedaPropiedadesInteresadosDepartamento.SelectedItem & "' and caracteristicas_inm.TipoInm = '" & BoxBusquedaPropiedadesInteresadosTipoInmueble.SelectedItem & "' and publicacion.TipoPubli = '" & BoxBusquedaPropiedadesInteresadosCompraVenta.SelectedItem & "'

;").Tables("consultainmueble")

            For i = 0 To DataBuscarPropiedadesImagenes.Rows.Count - 1

                Dim r As DataGridViewRow = DataBuscarPropiedadesImagenes.Rows(i)
                r.Height = 100


            Next

        End If


    End Sub

    Private Sub CheckBusquedasPropiedadesBarrio6_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBusquedasPropiedadesBarrio6.CheckedChanged
        ButtonlFormBusquedasPropiedadesPanelImagenesCerrar.PerformClick()



        If CheckBusquedasPropiedadesBarrio6.Checked = True Then

            barrio6 = CheckBusquedasPropiedadesBarrio6.Text



            DataBuscarPropiedadesImagenes.DataSource = consultaSQL("SELECT IdInmueble,	Barrios, NumDormitorios, NumBaños , Jardin , Amoblada , Garage , barbacoa , Moneda, Precio  From Inmueble, Ubicacion, caracteristicas_inm, publicacion , barrios
where ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio And inmueble.IdInm = caracteristicas_inm.idInm And Inmueble.IdPubli = publicacion.IdPubli And
	ubicacion.Departamento = '" & BoxBusquedaPropiedadesInteresadosDepartamento.SelectedItem & "' and caracteristicas_inm.TipoInm = '" & BoxBusquedaPropiedadesInteresadosTipoInmueble.SelectedItem & "' and publicacion.TipoPubli = '" & BoxBusquedaPropiedadesInteresadosCompraVenta.SelectedItem & "'
and ( barrios.Barrios = '" & barrio1 & "'  or barrios.Barrios = '" & barrio2 & "'  or barrios.Barrios = '" & barrio3 & "' or barrios.Barrios = '" & barrio4 & "' 
or barrios.Barrios = '" & barrio5 & "' or barrios.Barrios = '" & barrio6 & "' or barrios.Barrios = '" & barrio7 & "'        

)

;").Tables("consultainmueble")

            For i = 0 To DataBuscarPropiedadesImagenes.Rows.Count - 1

                Dim r As DataGridViewRow = DataBuscarPropiedadesImagenes.Rows(i)
                r.Height = 100


            Next


        End If


        If CheckBusquedasPropiedadesBarrio6.Checked = False Then
            barrio6 = ""

            DataBuscarPropiedadesImagenes.DataSource = consultaSQL("SELECT IdInmueble,	Barrios, NumDormitorios, NumBaños , Jardin , Amoblada , Garage , barbacoa , Moneda, Precio  From Inmueble, Ubicacion, caracteristicas_inm, publicacion , barrios
where ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio And inmueble.IdInm = caracteristicas_inm.idInm And Inmueble.IdPubli = publicacion.IdPubli And
	ubicacion.Departamento = '" & BoxBusquedaPropiedadesInteresadosDepartamento.SelectedItem & "' and caracteristicas_inm.TipoInm = '" & BoxBusquedaPropiedadesInteresadosTipoInmueble.SelectedItem & "' and publicacion.TipoPubli = '" & BoxBusquedaPropiedadesInteresadosCompraVenta.SelectedItem & "'
and ( barrios.Barrios = '" & barrio1 & "'  or barrios.Barrios = '" & barrio2 & "'  or barrios.Barrios = '" & barrio3 & "' or barrios.Barrios = '" & barrio4 & "' 
or barrios.Barrios = '" & barrio5 & "' or barrios.Barrios = '" & barrio6 & "' or barrios.Barrios = '" & barrio7 & "'        

)

;").Tables("consultainmueble")

            For i = 0 To DataBuscarPropiedadesImagenes.Rows.Count - 1

                Dim r As DataGridViewRow = DataBuscarPropiedadesImagenes.Rows(i)
                r.Height = 100


            Next

        End If



        If barrio1 = "" And barrio2 = "" And barrio3 = "" And barrio4 = "" And barrio5 = "" And barrio6 = "" And barrio7 = "" Then

            DataBuscarPropiedadesImagenes.DataSource = consultaSQL("SELECT IdInmueble,	Barrios, NumDormitorios, NumBaños , Jardin , Amoblada , Garage , barbacoa , Moneda, Precio  From Inmueble, Ubicacion, caracteristicas_inm, publicacion , barrios
where ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio And inmueble.IdInm = caracteristicas_inm.idInm And Inmueble.IdPubli = publicacion.IdPubli And
	ubicacion.Departamento = '" & BoxBusquedaPropiedadesInteresadosDepartamento.SelectedItem & "' and caracteristicas_inm.TipoInm = '" & BoxBusquedaPropiedadesInteresadosTipoInmueble.SelectedItem & "' and publicacion.TipoPubli = '" & BoxBusquedaPropiedadesInteresadosCompraVenta.SelectedItem & "'

;").Tables("consultainmueble")

            For i = 0 To DataBuscarPropiedadesImagenes.Rows.Count - 1

                Dim r As DataGridViewRow = DataBuscarPropiedadesImagenes.Rows(i)
                r.Height = 100


            Next

        End If



    End Sub

    Private Sub CheckBusquedasPropiedadesBarrio7_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBusquedasPropiedadesBarrio7.CheckedChanged

        ButtonlFormBusquedasPropiedadesPanelImagenesCerrar.PerformClick()


        If CheckBusquedasPropiedadesBarrio7.Checked = True Then

            barrio7 = CheckBusquedasPropiedadesBarrio7.Text



            DataBuscarPropiedadesImagenes.DataSource = consultaSQL("SELECT IdInmueble,	Barrios, NumDormitorios, NumBaños , Jardin , Amoblada , Garage , barbacoa , Moneda, Precio  From Inmueble, Ubicacion, caracteristicas_inm, publicacion , barrios
where ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio And inmueble.IdInm = caracteristicas_inm.idInm And Inmueble.IdPubli = publicacion.IdPubli And
	ubicacion.Departamento = '" & BoxBusquedaPropiedadesInteresadosDepartamento.SelectedItem & "' and caracteristicas_inm.TipoInm = '" & BoxBusquedaPropiedadesInteresadosTipoInmueble.SelectedItem & "' and publicacion.TipoPubli = '" & BoxBusquedaPropiedadesInteresadosCompraVenta.SelectedItem & "'
and ( barrios.Barrios = '" & barrio1 & "'  or barrios.Barrios = '" & barrio2 & "'  or barrios.Barrios = '" & barrio3 & "' or barrios.Barrios = '" & barrio4 & "' 
or barrios.Barrios = '" & barrio5 & "' or barrios.Barrios = '" & barrio6 & "' or barrios.Barrios = '" & barrio7 & "'         

)

;").Tables("consultainmueble")

            For i = 0 To DataBuscarPropiedadesImagenes.Rows.Count - 1

                Dim r As DataGridViewRow = DataBuscarPropiedadesImagenes.Rows(i)
                r.Height = 100


            Next


        End If



        If CheckBusquedasPropiedadesBarrio7.Checked = False Then
            barrio7 = ""

            DataBuscarPropiedadesImagenes.DataSource = consultaSQL("SELECT IdInmueble,	Barrios, NumDormitorios, NumBaños , Jardin , Amoblada , Garage , barbacoa , Moneda, Precio  From Inmueble, Ubicacion, caracteristicas_inm, publicacion , barrios
where ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio And inmueble.IdInm = caracteristicas_inm.idInm And Inmueble.IdPubli = publicacion.IdPubli And
	ubicacion.Departamento = '" & BoxBusquedaPropiedadesInteresadosDepartamento.SelectedItem & "' and caracteristicas_inm.TipoInm = '" & BoxBusquedaPropiedadesInteresadosTipoInmueble.SelectedItem & "' and publicacion.TipoPubli = '" & BoxBusquedaPropiedadesInteresadosCompraVenta.SelectedItem & "'
and ( barrios.Barrios = '" & barrio1 & "'  or barrios.Barrios = '" & barrio2 & "'  or barrios.Barrios = '" & barrio3 & "' or barrios.Barrios = '" & barrio4 & "' 
or barrios.Barrios = '" & barrio5 & "' or barrios.Barrios = '" & barrio6 & "' or barrios.Barrios = '" & barrio7 & "'         

)

;").Tables("consultainmueble")

            For i = 0 To DataBuscarPropiedadesImagenes.Rows.Count - 1

                Dim r As DataGridViewRow = DataBuscarPropiedadesImagenes.Rows(i)
                r.Height = 100


            Next

        End If



        If barrio1 = "" And barrio2 = "" And barrio3 = "" And barrio4 = "" And barrio5 = "" And barrio6 = "" And barrio7 = "" Then

            DataBuscarPropiedadesImagenes.DataSource = consultaSQL("SELECT IdInmueble,	Barrios, NumDormitorios, NumBaños , Jardin , Amoblada , Garage , barbacoa , Moneda, Precio  From Inmueble, Ubicacion, caracteristicas_inm, publicacion , barrios
where ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio And inmueble.IdInm = caracteristicas_inm.idInm And Inmueble.IdPubli = publicacion.IdPubli And
	ubicacion.Departamento = '" & BoxBusquedaPropiedadesInteresadosDepartamento.SelectedItem & "' and caracteristicas_inm.TipoInm = '" & BoxBusquedaPropiedadesInteresadosTipoInmueble.SelectedItem & "' and publicacion.TipoPubli = '" & BoxBusquedaPropiedadesInteresadosCompraVenta.SelectedItem & "'

;").Tables("consultainmueble")

            For i = 0 To DataBuscarPropiedadesImagenes.Rows.Count - 1

                Dim r As DataGridViewRow = DataBuscarPropiedadesImagenes.Rows(i)
                r.Height = 100


            Next

        End If



    End Sub

    Private Sub CheckBusquedasPropiedadesJardin_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBusquedasPropiedadesJardin.CheckedChanged
        'ButtonlFormBusquedasPropiedadesPanelImagenesCerrar.PerformClick()

        If busc = True Then


            If CheckBusquedasPropiedadesJardin.Checked = True Then
                jardin = "Si"


                DataBuscarPropiedadesImagenes.DataSource = consultaSQL("SELECT IdInmueble,	Barrios, NumDormitorios, NumBaños , Jardin , Amoblada , Garage , barbacoa , Piscina, Moneda, Precio  From Inmueble, Ubicacion, caracteristicas_inm, publicacion , barrios
where ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio And inmueble.IdInm = caracteristicas_inm.idInm And Inmueble.IdPubli = publicacion.IdPubli And
	ubicacion.Departamento = '" & BoxBusquedaPropiedadesInteresadosDepartamento.Text & "' and caracteristicas_inm.TipoInm = '" & BoxBusquedaPropiedadesInteresadosTipoInmueble.Text & "' and publicacion.TipoPubli = '" & BoxBusquedaPropiedadesInteresadosCompraVenta.Text & "'

and   Jardin = '" & jardin & "' and ( Amoblada = '" & amoblada & "' or Garage = '" & garage & "' or Piscina = '" & piscina & "' or Barbacoa = '" & barbacoa & "' )
      
       
;").Tables("consultainmueble")

                DataBuscarPropiedadesImagenes.Columns(1).Visible = False


                For i = 0 To DataBuscarPropiedadesImagenes.Rows.Count - 1

                    Dim r As DataGridViewRow = DataBuscarPropiedadesImagenes.Rows(i)
                    r.Height = 100


                Next





            End If


            If CheckBusquedasPropiedadesJardin.Checked = False Then
                jardin = "No"



                DataBuscarPropiedadesImagenes.DataSource = consultaSQL("SELECT IdInmueble,	Barrios, NumDormitorios, NumBaños , Jardin , Amoblada , Garage , barbacoa , Piscina , Moneda, Precio  From Inmueble, Ubicacion, caracteristicas_inm, publicacion , barrios
where ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio And inmueble.IdInm = caracteristicas_inm.idInm And Inmueble.IdPubli = publicacion.IdPubli And
	ubicacion.Departamento = '" & BoxBusquedaPropiedadesInteresadosDepartamento.Text & "' and caracteristicas_inm.TipoInm = '" & BoxBusquedaPropiedadesInteresadosTipoInmueble.Text & "' and publicacion.TipoPubli = '" & BoxBusquedaPropiedadesInteresadosCompraVenta.Text & "'

and   Jardin = '" & jardin & "' and ( Amoblada = '" & amoblada & "' or Garage = '" & garage & "' or Piscina = '" & piscina & "' or Barbacoa = '" & barbacoa & "' )
      
       
;").Tables("consultainmueble")

                DataBuscarPropiedadesImagenes.Columns(1).Visible = False


                For i = 0 To DataBuscarPropiedadesImagenes.Rows.Count - 1

                    Dim r As DataGridViewRow = DataBuscarPropiedadesImagenes.Rows(i)
                    r.Height = 100


                Next





            End If







        End If





    End Sub

    Private Sub CheckBusquedasPropiedadesPiscina_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBusquedasPropiedadesPiscina.CheckedChanged
        'ButtonlFormBusquedasPropiedadesPanelImagenesCerrar.PerformClick()

        If busc = True Then
            If CheckBusquedasPropiedadesPiscina.Checked = True Then
                piscina = "Si"
                DataBuscarPropiedadesImagenes.DataSource = consultaSQL("SELECT IdInmueble,	Barrios, NumDormitorios, NumBaños , Jardin , Amoblada , Garage , barbacoa , Piscina , Moneda, Precio  From Inmueble, Ubicacion, caracteristicas_inm, publicacion , barrios
where ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio And inmueble.IdInm = caracteristicas_inm.idInm And Inmueble.IdPubli = publicacion.IdPubli And
	ubicacion.Departamento = '" & BoxBusquedaPropiedadesInteresadosDepartamento.Text & "' and caracteristicas_inm.TipoInm = '" & BoxBusquedaPropiedadesInteresadosTipoInmueble.Text & "' and publicacion.TipoPubli = '" & BoxBusquedaPropiedadesInteresadosCompraVenta.Text & "'

and  Piscina = '" & piscina & "'  and ( Amoblada = '" & amoblada & "' or Garage = '" & garage & "' or Jardin = '" & jardin & "' or Barbacoa = '" & barbacoa & "' )
      
       
;").Tables("consultainmueble")

                DataBuscarPropiedadesImagenes.Columns(1).Visible = False


                For i = 0 To DataBuscarPropiedadesImagenes.Rows.Count - 1

                    Dim r As DataGridViewRow = DataBuscarPropiedadesImagenes.Rows(i)
                    r.Height = 100


                Next


            End If

            If CheckBusquedasPropiedadesPiscina.Checked = False Then
                piscina = "No"

                DataBuscarPropiedadesImagenes.DataSource = consultaSQL("SELECT IdInmueble,	Barrios, NumDormitorios, NumBaños , Jardin , Amoblada , Garage , barbacoa , Piscina , Moneda, Precio  From Inmueble, Ubicacion, caracteristicas_inm, publicacion , barrios
where ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio And inmueble.IdInm = caracteristicas_inm.idInm And Inmueble.IdPubli = publicacion.IdPubli And
	ubicacion.Departamento = '" & BoxBusquedaPropiedadesInteresadosDepartamento.Text & "' and caracteristicas_inm.TipoInm = '" & BoxBusquedaPropiedadesInteresadosTipoInmueble.Text & "' and publicacion.TipoPubli = '" & BoxBusquedaPropiedadesInteresadosCompraVenta.Text & "'

and  Piscina = '" & piscina & "'  and ( Amoblada = '" & amoblada & "' or Garage = '" & garage & "' or Jardin = '" & jardin & "' or Barbacoa = '" & barbacoa & "' )
      
       
;").Tables("consultainmueble")

                DataBuscarPropiedadesImagenes.Columns(1).Visible = False


                For i = 0 To DataBuscarPropiedadesImagenes.Rows.Count - 1

                    Dim r As DataGridViewRow = DataBuscarPropiedadesImagenes.Rows(i)
                    r.Height = 100


                Next


            End If

        End If



    End Sub

    Private Sub CheckBusquedasPropiedadesGarage_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBusquedasPropiedadesGarage.CheckedChanged
        'ButtonlFormBusquedasPropiedadesPanelImagenesCerrar.PerformClick()


        If busc = True Then

            If CheckBusquedasPropiedadesGarage.Checked = True Then
                garage = "Si"

                DataBuscarPropiedadesImagenes.DataSource = consultaSQL("SELECT IdInmueble,	Barrios, NumDormitorios, NumBaños , Jardin , Amoblada , Garage , barbacoa , Piscina , Moneda, Precio  From Inmueble, Ubicacion, caracteristicas_inm, publicacion , barrios
where ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio And inmueble.IdInm = caracteristicas_inm.idInm And Inmueble.IdPubli = publicacion.IdPubli And
	ubicacion.Departamento = '" & BoxBusquedaPropiedadesInteresadosDepartamento.Text & "' and caracteristicas_inm.TipoInm = '" & BoxBusquedaPropiedadesInteresadosTipoInmueble.Text & "' and publicacion.TipoPubli = '" & BoxBusquedaPropiedadesInteresadosCompraVenta.Text & "'

and  Garage = '" & garage & "'and ( Amoblada = '" & amoblada & "' or  Jardin = '" & jardin & "'  or Piscina = '" & piscina & "' or Barbacoa = '" & barbacoa & "' )
      
       
;").Tables("consultainmueble")

                DataBuscarPropiedadesImagenes.Columns(1).Visible = False


                For i = 0 To DataBuscarPropiedadesImagenes.Rows.Count - 1

                    Dim r As DataGridViewRow = DataBuscarPropiedadesImagenes.Rows(i)
                    r.Height = 100


                Next

            End If


            If CheckBusquedasPropiedadesGarage.Checked = False Then
                garage = "No"

                DataBuscarPropiedadesImagenes.DataSource = consultaSQL("SELECT IdInmueble,	Barrios, NumDormitorios, NumBaños , Jardin , Amoblada , Garage , barbacoa , Piscina , Moneda, Precio  From Inmueble, Ubicacion, caracteristicas_inm, publicacion , barrios
where ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio And inmueble.IdInm = caracteristicas_inm.idInm And Inmueble.IdPubli = publicacion.IdPubli And
	ubicacion.Departamento = '" & BoxBusquedaPropiedadesInteresadosDepartamento.Text & "' and caracteristicas_inm.TipoInm = '" & BoxBusquedaPropiedadesInteresadosTipoInmueble.Text & "' and publicacion.TipoPubli = '" & BoxBusquedaPropiedadesInteresadosCompraVenta.Text & "'

and  Garage = '" & garage & "'and ( Amoblada = '" & amoblada & "' or  Jardin = '" & jardin & "'  or Piscina = '" & piscina & "' or Barbacoa = '" & barbacoa & "' )
      
       
;").Tables("consultainmueble")

                DataBuscarPropiedadesImagenes.Columns(1).Visible = False


                For i = 0 To DataBuscarPropiedadesImagenes.Rows.Count - 1

                    Dim r As DataGridViewRow = DataBuscarPropiedadesImagenes.Rows(i)
                    r.Height = 100


                Next
            End If



        End If





    End Sub

    Private Sub CheckBusquedasPropiedadesAmoblada_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBusquedasPropiedadesAmoblada.CheckedChanged
        'ButtonlFormBusquedasPropiedadesPanelImagenesCerrar.PerformClick()



        If busc = True Then

            If CheckBusquedasPropiedadesAmoblada.Checked = True Then
                amoblada = "Si"

                DataBuscarPropiedadesImagenes.DataSource = consultaSQL("SELECT IdInmueble,	Barrios, NumDormitorios, NumBaños , Jardin , Amoblada , Garage , barbacoa , Piscina ,  Moneda, Precio  From Inmueble, Ubicacion, caracteristicas_inm, publicacion , barrios
where ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio And inmueble.IdInm = caracteristicas_inm.idInm And Inmueble.IdPubli = publicacion.IdPubli And
	ubicacion.Departamento = '" & BoxBusquedaPropiedadesInteresadosDepartamento.Text & "' and caracteristicas_inm.TipoInm = '" & BoxBusquedaPropiedadesInteresadosTipoInmueble.Text & "' and publicacion.TipoPubli = '" & BoxBusquedaPropiedadesInteresadosCompraVenta.Text & "'

and   Amoblada = '" & amoblada & "'  and ( Jardin = '" & jardin & "' or Garage = '" & garage & "' or Piscina = '" & piscina & "' or Barbacoa = '" & barbacoa & "' )
      
       
;").Tables("consultainmueble")

                DataBuscarPropiedadesImagenes.Columns(1).Visible = False


                For i = 0 To DataBuscarPropiedadesImagenes.Rows.Count - 1

                    Dim r As DataGridViewRow = DataBuscarPropiedadesImagenes.Rows(i)
                    r.Height = 100


                Next



            End If


            If CheckBusquedasPropiedadesAmoblada.Checked = False Then
                amoblada = "No"

                DataBuscarPropiedadesImagenes.DataSource = consultaSQL("SELECT IdInmueble,	Barrios, NumDormitorios, NumBaños , Jardin , Amoblada , Garage , barbacoa , Piscina , Moneda, Precio  From Inmueble, Ubicacion, caracteristicas_inm, publicacion , barrios
where ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio And inmueble.IdInm = caracteristicas_inm.idInm And Inmueble.IdPubli = publicacion.IdPubli And
	ubicacion.Departamento = '" & BoxBusquedaPropiedadesInteresadosDepartamento.Text & "' and caracteristicas_inm.TipoInm = '" & BoxBusquedaPropiedadesInteresadosTipoInmueble.Text & "' and publicacion.TipoPubli = '" & BoxBusquedaPropiedadesInteresadosCompraVenta.Text & "'

and   Amoblada = '" & amoblada & "'  and ( Jardin = '" & jardin & "' or Garage = '" & garage & "' or Piscina = '" & piscina & "' or Barbacoa = '" & barbacoa & "' )
      
       
;").Tables("consultainmueble")

                DataBuscarPropiedadesImagenes.Columns(1).Visible = False


                For i = 0 To DataBuscarPropiedadesImagenes.Rows.Count - 1

                    Dim r As DataGridViewRow = DataBuscarPropiedadesImagenes.Rows(i)
                    r.Height = 100


                Next


            End If

        End If








    End Sub

    Private Sub CheckBusquedasPropiedadesBarbacoa_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBusquedasPropiedadesBarbacoa.CheckedChanged
        'ButtonlFormBusquedasPropiedadesPanelImagenesCerrar.PerformClick()

        If busc = True Then


            If CheckBusquedasPropiedadesBarbacoa.Checked = True Then
                barbacoa = "Si"

                DataBuscarPropiedadesImagenes.DataSource = consultaSQL("SELECT IdInmueble,	Barrios, NumDormitorios, NumBaños , Jardin , Amoblada , Garage , barbacoa , Piscina , Moneda, Precio  From Inmueble, Ubicacion, caracteristicas_inm, publicacion , barrios
where ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio And inmueble.IdInm = caracteristicas_inm.idInm And Inmueble.IdPubli = publicacion.IdPubli And
	ubicacion.Departamento = '" & BoxBusquedaPropiedadesInteresadosDepartamento.Text & "' and caracteristicas_inm.TipoInm = '" & BoxBusquedaPropiedadesInteresadosTipoInmueble.Text & "' and publicacion.TipoPubli = '" & BoxBusquedaPropiedadesInteresadosCompraVenta.Text & "'

and Barbacoa = '" & barbacoa & "'  and ( Jardin = '" & jardin & "' or Garage = '" & garage & "' or Piscina = '" & piscina & "' or Amoblada = '" & amoblada & "')
      
       
;").Tables("consultainmueble")

                DataBuscarPropiedadesImagenes.Columns(1).Visible = False


                For i = 0 To DataBuscarPropiedadesImagenes.Rows.Count - 1

                    Dim r As DataGridViewRow = DataBuscarPropiedadesImagenes.Rows(i)
                    r.Height = 100


                Next


            End If


            If CheckBusquedasPropiedadesBarbacoa.Checked = False Then
                barbacoa = "No"

                DataBuscarPropiedadesImagenes.DataSource = consultaSQL("SELECT IdInmueble,	Barrios, NumDormitorios, NumBaños , Jardin , Amoblada , Garage , barbacoa , Piscina , Moneda, Precio  From Inmueble, Ubicacion, caracteristicas_inm, publicacion , barrios
where ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio And inmueble.IdInm = caracteristicas_inm.idInm And Inmueble.IdPubli = publicacion.IdPubli And
	ubicacion.Departamento = '" & BoxBusquedaPropiedadesInteresadosDepartamento.Text & "' and caracteristicas_inm.TipoInm = '" & BoxBusquedaPropiedadesInteresadosTipoInmueble.Text & "' and publicacion.TipoPubli = '" & BoxBusquedaPropiedadesInteresadosCompraVenta.Text & "'

and Barbacoa = '" & barbacoa & "'  and ( Jardin = '" & jardin & "' or Garage = '" & garage & "' or Piscina = '" & piscina & "' or Amoblada = '" & amoblada & "')
      
       
;").Tables("consultainmueble")

                DataBuscarPropiedadesImagenes.Columns(1).Visible = False


                For i = 0 To DataBuscarPropiedadesImagenes.Rows.Count - 1

                    Dim r As DataGridViewRow = DataBuscarPropiedadesImagenes.Rows(i)
                    r.Height = 100


                Next


            End If




        End If




    End Sub



    'Muestra la imagenes para el inmueble seleccionado
    Private Sub DataBuscarPropiedadesImagenes_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataBuscarPropiedadesImagenes.CellClick
        PanelFormBusquedasPropiedadesPanelImagenes.Visible = True
        PanelFormBusquedasPropiedadesPanelImagenes.Location = New System.Drawing.Point(PanelBusquedasPropiedades.Size.Width / 4, PanelBusquedasPropiedades.Size.Height / 12)
        PanelFormBusquedasPropiedadesPanelImagenes.Size = New System.Drawing.Size(PanelBusquedaPropiedadesInteresadosInmuebles.Size.Width - 450, PanelBusquedaPropiedadesInteresadosInmuebles.Size.Height - 120)
        'MUESTRA LAS IMAGENES DEL INMUEBLE
        PictureFormBusquedasPropiedadesPanelImagenes.Image = Nothing


        Idinmueble = DataBuscarPropiedadesImagenes.CurrentRow.Cells("IdInmueble").Value


        'llama al metodo para mostrar la primera imagen
        mostrarimagen(cantidadimg)

        'OBTIENE EL DEPARTAMENTO ASOCIADO AL INMUEBLE



        Dim Depinm As String = consultaSQL("SELECT Departamento FROM ubicacion , barrios , inmueble where ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio and IdInmueble = " & Idinmueble & ";").Tables("inmueble").Rows(0)("Departamento").ToString()

        Dim TelSuc As String = consultaSQL("SELECT Tels FROM tel_sucursal , barrios , sucursales , ubicacion where tel_sucursal.IdSucursal = sucursales.IdSucursal
        and sucursales.IdBarrio = barrios.IdBarrio and barrios.IdUbicacion = ubicacion.IdUbicacion and Departamento = '" & Depinm & "';").Tables("inmueble").Rows(0)("Tels").ToString()


        Idinmueble = Idinmueble + 100
        Dim idinmst As String = Idinmueble & ""

        'GENERAR CODIGO QR
        Dim mensaje As String = "Numero contacto: " + TelSuc + "      Numero inmueble:" + idinmst
        Dim gen As New QRCodeGenerator
        Dim data = gen.CreateQrCode(mensaje, QRCodeGenerator.ECCLevel.Q)
        Dim code As New QRCode(data)
        PicQR.Image = code.GetGraphic(8)



    End Sub


    Private Sub BtnBusquedasPropiedadesAgendar_Click_1(sender As Object, e As EventArgs) Handles BtnBusquedasPropiedadesAgendar.Click

        Idinmueble = Idinmueble - 100

        Dim fechahoraactual = fechaactual.Year & "-" & fechaactual.Month & "-" & fechaactual.Day & " " & horaactual

        updateSQL("insert into reserva ( IdInmueble , IdAgenda , IdInteresado , FechaHraReserva , EstadoReserva ) values
                                    ( " & Idinmueble & " ,  1  , " & IdCliente & " , '" & fechahoraactual & "' , 'Pendiente');")




        Dim idreserva As String = consultaSQL("select IdReserva from  reserva where IdInmueble = " & Idinmueble & " and IdAgenda = 1 and 
             IdInteresado = " & IdCliente & "  and FechaHraReserva =   '" & fechahoraactual & "' and  EstadoReserva = 'Pendiente'
            ;").Tables("consultainmueble").Rows(0)("IdReserva").ToString()




        updateSQL("insert into notifica ( IdEnvia , IdRecibe , FechaNotifi , TipoNotif , Texto , IdReserva ) values 
           (  1 , " & IdCliente & " , '" & fechahoraactual & "' , 'SMS' , 'Hola " & cliente & " , un agente estará contactandose en breve para
        coordinar la visita.' , " & idreserva & " ) ;")



        Try
            'SE ENVIA SMS AL INTERESADO
            Dim SMSMasivos As New SMSMasivosWS.SMSMasivosAPISoapClient("SMSMasivosAPISoap")


            'El último parámetro indica si es TEST o no, poner el valor en FALSE para enviar realmente el mensaje
            MsgBox(SMSMasivos.EnviarSMS("DASILVA", "DASILVA908", Celular, "Hola " & cliente & " , un agente estará contactandose en breve para
        coordinar la visita.", True))

        Catch ex As Exception

        End Try




        PanelFormBusquedasPropiedadesPanelImagenes.Visible = False


    End Sub

    Private Sub PictureFormBusquedasPropiedadesPanelImagenes_Click(sender As Object, e As EventArgs) Handles PictureFormBusquedasPropiedadesPanelImagenes.Click


        imagenesFullPantalla.Show()

        imagenesFullPantalla.PicImagenesFullPantalla.Image = PictureFormBusquedasPropiedadesPanelImagenes.Image


    End Sub



    Private Sub PanelBusquedasPropiedadesIzquierdoInterno_MouseMove(sender As Object, e As MouseEventArgs) Handles PanelBusquedasPropiedadesIzquierdoInterno.MouseMove








    End Sub

End Class