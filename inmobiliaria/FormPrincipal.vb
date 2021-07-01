Imports System.Runtime.InteropServices
Imports System.IO
Imports System.Drawing.Imaging
Imports System.Text.RegularExpressions
Imports System.Globalization
Imports QRCoder
Imports MySql.Data.MySqlClient
Imports System.Windows.Forms.DateTimePickerFormat.Time










Public Class FormPrincipal
    Public IdPersonal As Integer
    Public inicial As Date
    Public inicialagenda As Date
    Public finalagenda As Date
    Public inicialhistorico As Date
    Public inicialfechaagente As Date
    Public finalfechaagente As Date
    Public fechainicialtrabajo As Date
    Public mifecha As Date
    Public BarrioInmueble As String
    Public Departamentoinmueble As String
    Public finalhistorico As Date
    Public Diasemana As String = ""
    Public Fechareserva As String = ""
    Public añoreserva As String = ""
    Public mesreserva As String = ""
    Public diareserva As String = ""
    Public FechareservaModificada As String = ""
    Public Fechaagendaini As String = ""
    Public Fechaagendafin As String = ""
    Public Fechahistoricoini As String = ""
    Public Fechahinicialagente As String = ""
    Public Fechafinalagente As String = ""
    Public Fechahistoricofin As String = ""
    Public fechatrabajonicial As String = ""

    Public IdSucursal As String
    Dim IdBarriopersonaInt As Integer
    Dim IdBarrioinmuebleInt As Integer
    Public IdPersonaInt As Integer = 0
    Dim IdInmuebleInt As Integer
    Dim idsucusralInt As Integer
    Dim IdReserva As Integer = 0
    Dim fechaactual As Date = Today
    Dim horaactual As DateTime = TimeString
    Public jardin As String = "Si"
    Public piscina As String = "Si"
    Public garage As String = "Si"
    Public amoblada As String = "Si"
    Public barbacoa As String = "Si"
    Public mascotas As String = "Si"
    Public cantidadimagenes As Integer = 0
    Public cantidadimagenesedit As Integer = 0
    Public cantidadimageneseditupdate As Integer = 0
    Public primerdiames As String
    Public restameses As Integer = 0
    Public sumameses As Integer = 0


    Public horainicialagente As String = ""
    Public horafinalagente As String = ""






    Public testingvalidaringresousuario As Boolean
    Public buscarcaract As Boolean
    Public agenda As String
    Public finalhorareserva As Integer
    Public count As String = ""
    Public countfechas As String
    Public fecha As String = ""
    Public inicioreserva As Integer
    Public cantidadhoras As Integer = 0
    Public departamentoagente As String
    Public inicioreservamodificada As Integer
    Public finreservamodificada As Integer
    Public Idinmueble As String = ""
    Public cantidadimg As Integer = 0





    Private Sub FormularioPrincipal_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: esta línea de código carga datos en la tabla 'Proyecto2DataSet.ubicacion' Puede moverla o quitarla según sea necesario.
        'Me.UbicacionTableAdapter.Fill(Me.Proyecto2DataSet.ubicacion)











        departamentoagente = consultaSQL("SELECT Departamento FROM ubicacion , barrios , personas
where personas.IdBarrio = barrios.IdBarrio and barrios.IdUbicacion = ubicacion.IdUbicacion
and personas.IdPersona = " & IdPersonal & "
;").Tables("personas").Rows(0)("Departamento").ToString()
        'Convierte en integer el id de la ultima persona ingresada


        'RECORRE EL FORMULARIO PARA QUITAR Y MOSTRAR OBJETOS DEPENDIENDO EL TIPO DE USUARIO
        With Me

            If FormLogin.tipoacceso = "Administrador" Then

                PanelFormPrincipalConsultaragenda.Visible = True
                PanelFormPrincipalRegistrarUsuarioDatos.Visible = True
                PanelFormPrincipalRegistrarUsuarioUsuarios.Visible = True

                ComboFormPrincipalRegistroUsuarioTipoPersona.Visible = False
                LabelFormPrincipalRegistroUsuarioTipoPers.Visible = False
                BoxFormPrincipalConsultaAgendaAgente.Visible = True
                LabelFormPrincipalConsultaragendaAgente.Visible = True
                ComboFormPrincipalCrearreservaAgentes.Visible = True
                BtnFormPrincipalCrearHorarios.Visible = True
                LabelFormPrincipalCrearreservaAgentes.Visible = True

                'CARGA LOS NOMBRES DE LOS AGENTES
                BoxFormPrincipalConsultaAgendaAgente.DataSource = (consultaSQL("select concat(concat(Nombre , (' ')) , Apellido ) from personas where Funcion = 'Agente';").Tables(0))
                BoxFormPrincipalConsultaAgendaAgente.DisplayMember = "concat(concat(Nombre , (' ')) , Apellido )"



                'MUESTRA LAS AGENDAS PARA EL DIA DE HOY
                DataFormPrincipalConsultaragendaDatos.DataSource = consultaSQL("select EstadoReserva ,  FehaHraVisita as Fecha , concat(concat(Nombre , (' '))  , Apellido ) as Nombre_Apellido , Tel as Telefono , Cel as Celular , Email , Barrios as Barrio , inmueble.Calle , inmueble.NroPuerta as Numero_Puerta, inmueble.NroApto as Numero_Apartamento , reserva.IdReserva from agenda ,  reserva , personas , inmueble , barrios  where  agenda.IdAgenda = reserva.IdAgenda and reserva.IdInteresado = personas.IdPersona and reserva.IdInmueble = inmueble.IdInmueble and barrios.IdBarrio = inmueble.IdBarrio
                and EstadoReserva = 'Pendiente' and reserva.IdAgenda = 1
                ;").Tables("consultainmueble")

                DataFormPrincipalConsultaragendaDatos.Columns(10).Visible = False

                DataFormPrincipalConsultaragendaDatos.ClearSelection()


            End If



            If FormLogin.tipoacceso = "Gerente" Then
                PanelFormPrincipalConsultaragenda.Visible = True
                PanelFormPrincipalRegistrarUsuarioDatos.Visible = True
                PanelFormPrincipalRegistrarUsuarioUsuarios.Visible = True

                ComboFormPrincipalRegistroUsuarioTipoPersona.Visible = False
                LabelFormPrincipalRegistroUsuarioTipoPers.Visible = False
                BoxFormPrincipalConsultaAgendaAgente.Visible = True
                LabelFormPrincipalConsultaragendaAgente.Visible = True
                BtnFormPrincipalCrearHorarios.Visible = True
                ComboFormPrincipalCrearreservaAgentes.Visible = True
                LabelFormPrincipalCrearreservaAgentes.Visible = True
                'CARGA LOS NOMBRES DE LOS AGENTES
                BoxFormPrincipalConsultaAgendaAgente.DataSource = (consultaSQL("select concat(concat(Nombre , (' ')) , Apellido ) from personas where Funcion = 'Agente';").Tables(0))
                BoxFormPrincipalConsultaAgendaAgente.DisplayMember = "concat(concat(Nombre , (' ')) , Apellido )"



                'MUESTRA LAS AGENDAS PARA EL DIA DE HOY
                DataFormPrincipalConsultaragendaDatos.DataSource = consultaSQL("select EstadoReserva ,  FehaHraVisita as Fecha , concat(concat(Nombre , (' '))  , Apellido ) as Nombre_Apellido , Tel as Telefono , Cel as Celular , Email , Barrios as Barrio , inmueble.Calle , inmueble.NroPuerta as Numero_Puerta, inmueble.NroApto as Numero_Apartamento , reserva.IdReserva from agenda ,  reserva , personas , inmueble , barrios  where  agenda.IdAgenda = reserva.IdAgenda and reserva.IdInteresado = personas.IdPersona and reserva.IdInmueble = inmueble.IdInmueble and barrios.IdBarrio = inmueble.IdBarrio
                and EstadoReserva = 'Pendiente' and reserva.IdAgenda = 1
                ;").Tables("consultainmueble")

                DataFormPrincipalConsultaragendaDatos.Columns(10).Visible = False

                DataFormPrincipalConsultaragendaDatos.ClearSelection()

            End If



            If FormLogin.tipoacceso = "Agente" Then

                BtnFormPrincipalHistorialVisitas.Visible = False
                PanelFormPrincipalRegistrarUsuarioUsuarios.Visible = False
                PanelFormPrincipalConsultaragenda.Visible = True
                PanelFormPrincipalRegistrarUsuarioDatos.Visible = True

                Dim fecha As String = fechaactual.Year & "-" & fechaactual.Month & "-" & fechaactual.Day



                DataFormPrincipalConsultaragendaDatos.DataSource = consultaSQL("select EstadoReserva , FehaHraVisita as Fecha , concat(concat(Nombre , (' '))  , Apellido ) as Nombre_Apellido , Tel as Telefono , Cel as Celular , Email , Barrios as Barrio , inmueble.Calle , inmueble.NroPuerta as Numero_Puerta, inmueble.NroApto as Numero_Apartamento , reserva.IdReserva from agenda ,  reserva , personas , inmueble , barrios  where  agenda.IdAgenda = reserva.IdAgenda and reserva.IdInteresado = personas.IdPersona and reserva.IdInmueble = inmueble.IdInmueble and barrios.IdBarrio = inmueble.IdBarrio
                and reserva.FehaHraVisita like '%" & fecha & "%' and  agenda.IdAgente = " & IdPersonal & "
                AND EstadoReserva = 'Activo'
                union all select EstadoReserva ,  FehaHraVisita as Fecha , concat(concat(Nombre , (' '))  , Apellido ) as Nombre_Apellido , Tel as Telefono , Cel as Celular , Email , Barrios as Barrio , inmueble.Calle , inmueble.NroPuerta as Numero_Puerta, inmueble.NroApto as Numero_Apartamento , reserva.IdReserva from agenda ,  reserva , personas , inmueble , barrios  where  agenda.IdAgenda = reserva.IdAgenda and reserva.IdInteresado = personas.IdPersona and reserva.IdInmueble = inmueble.IdInmueble and barrios.IdBarrio = inmueble.IdBarrio
                and EstadoReserva = 'Pendiente' and reserva.IdAgenda = 1
                ;").Tables("consultainmueble")

                DataFormPrincipalConsultaragendaDatos.Columns(10).Visible = False


                DataFormPrincipalConsultaragendaDatos.ClearSelection()

            End If




            If FormLogin.tipoacceso = "AgenteFijo" Then

                LabelFormPrincipalConsultaragendaAgente.Visible = True
                BoxFormPrincipalConsultaAgendaAgente.Visible = True
                BtnFormPrincipalHistorialVisitas.Visible = False
                LabelFormPrincipalCrearreservaAgentes.Visible = True
                ComboFormPrincipalCrearreservaAgentes.Visible = True
                BtnFormPrincipalConsultarAgenda.Visible = False
                BtnFormPrincipalConsultarAgenda.Visible = True
                PanelFormPrincipalRegistrarUsuarioUsuarios.Visible = False
                PanelFormPrincipalConsultaragenda.Visible = True
                PanelFormPrincipalRegistrarUsuarioDatos.Visible = True

                Dim fecha As String = fechaactual.Year & "-" & fechaactual.Month & "-" & fechaactual.Day

                DataFormPrincipalConsultaragendaDatos.DataSource = consultaSQL("select EstadoReserva , FehaHraVisita as Fecha , concat(concat(Nombre , (' '))  , Apellido ) as Nombre_Apellido , Tel as Telefono , Cel as Celular , Email , Barrios as Barrio , inmueble.Calle , inmueble.NroPuerta as Numero_Puerta, inmueble.NroApto as Numero_Apartamento , reserva.IdReserva from agenda ,  reserva , personas , inmueble , barrios  where  agenda.IdAgenda = reserva.IdAgenda and reserva.IdInteresado = personas.IdPersona and reserva.IdInmueble = inmueble.IdInmueble and barrios.IdBarrio = inmueble.IdBarrio
                and reserva.FehaHraVisita like '%" & fecha & "%' and  agenda.IdAgente = " & IdPersonal & "
                AND EstadoReserva = 'Activo'
                union all select EstadoReserva ,  FehaHraVisita as Fecha , concat(concat(Nombre , (' '))  , Apellido ) as Nombre_Apellido , Tel as Telefono , Cel as Celular , Email , Barrios as Barrio , inmueble.Calle , inmueble.NroPuerta as Numero_Puerta, inmueble.NroApto as Numero_Apartamento , reserva.IdReserva from agenda ,  reserva , personas , inmueble , barrios  where  agenda.IdAgenda = reserva.IdAgenda and reserva.IdInteresado = personas.IdPersona and reserva.IdInmueble = inmueble.IdInmueble and barrios.IdBarrio = inmueble.IdBarrio
                and EstadoReserva = 'Pendiente' and reserva.IdAgenda = 1
                ;").Tables("consultainmueble")

                DataFormPrincipalConsultaragendaDatos.Columns(10).Visible = False


                'CARGA LOS NOMBRES DE LOS AGENTES
                BoxFormPrincipalConsultaAgendaAgente.DataSource = (consultaSQL("select concat(concat(Nombre , (' ')) , Apellido ) from personas where Funcion = 'Agente';").Tables(0))
                BoxFormPrincipalConsultaAgendaAgente.DisplayMember = "concat(concat(Nombre , (' ')) , Apellido )"



                'MUESTRA LAS AGENDAS PARA EL DIA DE HOY
                DataFormPrincipalConsultaragendaDatos.DataSource = consultaSQL("select EstadoReserva ,  FehaHraVisita as Fecha , concat(concat(Nombre , (' '))  , Apellido ) as Nombre_Apellido , Tel as Telefono , Cel as Celular , Email , Barrios as Barrio , inmueble.Calle , inmueble.NroPuerta as Numero_Puerta, inmueble.NroApto as Numero_Apartamento , reserva.IdReserva from agenda ,  reserva , personas , inmueble , barrios  where  agenda.IdAgenda = reserva.IdAgenda and reserva.IdInteresado = personas.IdPersona and reserva.IdInmueble = inmueble.IdInmueble and barrios.IdBarrio = inmueble.IdBarrio
                and EstadoReserva = 'Pendiente' and reserva.IdAgenda = 1
                ;").Tables("consultainmueble")

                DataFormPrincipalConsultaragendaDatos.Columns(10).Visible = False

                DataFormPrincipalConsultaragendaDatos.ClearSelection()


            End If











        End With
    End Sub



    '-------------------------------------BOTONES MENU----------------------------------------------------------

    'BOTON BUSCAR
    Private Sub BtnBuscar_Click(sender As Object, e As EventArgs) Handles BtnFormPrincipalBuscar.Click
        Idinmueble = ""
        PanelBuscarInmueblesImagenes.Visible = False
        With Me



            .PanelFormPrincipalCrearreserva.Visible = False
            .BtnPrincipalCrearreserva.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
            .PanelFormPrincipalConsultaragenda.Visible = False
            .BtnFormPrincipalConsultarAgenda.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
            .FormPrincipalRegistroFechanacimiento.Visible = False
            .BtnFormPrincipalRegistrarInteresado.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
            .PanelFormPrincipalHistorialreservas.Visible = False
            .BtnFormPrincipalHistorialVisitas.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
            .PanelIngresoPropiedades.Visible = False
            .BtnFormPrincipalIngresoPropiedades.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
            .PanelFormPrincipalHorarios.Visible = False
            .BtnFormPrincipalCrearHorarios.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))



            .PanelFormPrincipalBuscarInmuebles.Visible = True
            .BtnFormPrincipalBuscar.ForeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))

        End With

        CheckFormPrincipalBuscarInmueblesPatio.Checked = True
        CheckFormPrincipalBuscarInmueblesAmoblada.Checked = True
        CheckFormPrincipalBuscarInmueblesGarage.Checked = True
        CheckFormPrincipalBuscarInmueblesPiscina.Checked = True
        CheckFormPrincipalBuscarInmueblesBarbacoa.Checked = True

        BoxFormPrincipalBuscarInmueblesAlquilerUbicacion.Text = departamentoagente

        DataFormPrincipalBuscarInmuebles.DataSource = consultaSQL("select IdInmueble , TipoInm, Departamento,  TipoPubli , NumDormitorios , Precio , Moneda , Jardin , Amoblada , Garage , Piscina , Barbacoa
                                                                    from ubicacion, caracteristicas_inm, inmueble, barrios, publicacion 	
                                                                    where caracteristicas_inm.IdInm = inmueble.IdInm and ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio and publicacion.IdPubli = inmueble.IdPubli
                                                                    and Departamento = '" & departamentoagente & "'  and publicacion.EstadoInm = 'Activo' limit 50;").Tables("consultainmueble")

        DataFormPrincipalBuscarInmuebles.Columns(0).Visible = False

        buscarcaract = True


    End Sub


    'BOTON AGENDAR VISITA
    Private Sub BtnAgendar_Click(sender As Object, e As EventArgs) Handles BtnPrincipalCrearreserva.Click

        With Me


            .PanelFormPrincipalConsultaragenda.Visible = False
            .BtnFormPrincipalConsultarAgenda.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
            .FormPrincipalRegistroFechanacimiento.Visible = False
            .BtnFormPrincipalRegistrarInteresado.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
            .PanelFormPrincipalHistorialreservas.Visible = False
            .BtnFormPrincipalHistorialVisitas.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
            .PanelIngresoPropiedades.Visible = False
            .BtnFormPrincipalIngresoPropiedades.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
            .PanelFormPrincipalHorarios.Visible = False
            .BtnFormPrincipalCrearHorarios.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
            .PanelFormPrincipalBuscarInmuebles.Visible = False
            .BtnFormPrincipalBuscar.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))


            .PanelFormPrincipalCrearreserva.Visible = True
            .BtnPrincipalCrearreserva.ForeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))


            'CARGA LOS NOMBRES DE LOS AGENTES
            ComboFormPrincipalCrearreservaAgentes.DataSource = (consultaSQL("select concat(concat(Nombre , (' ')) , Apellido ) from personas where Funcion = 'Agente';").Tables(0))
            ComboFormPrincipalCrearreservaAgentes.DisplayMember = "concat(concat(Nombre , (' ')) , Apellido )"


            DataFormPrincipalCrearreserva.DataSource = consultaSQL("select IdInmueble , TipoInm, Departamento,  TipoPubli , NumDormitorios , Precio , Moneda , Jardin , Amoblada , Garage , Piscina , Barbacoa
                                                                    from ubicacion, caracteristicas_inm, inmueble, barrios, publicacion 	
                                                                    where caracteristicas_inm.IdInm = inmueble.IdInm and ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio and publicacion.IdPubli = inmueble.IdPubli
                                                                    and ubicacion.Departamento = '" & departamentoagente & "' and publicacion.EstadoInm = 'Activo'                                                          
                                                                    limit 50;").Tables("consultainmueble")

            DataFormPrincipalCrearreserva.Columns(0).Visible = False


        End With






    End Sub


    'BOTON CONSULTAR AGENDA DE VISITAS
    Private Sub BtnConsultarAgenda_Click(sender As Object, e As EventArgs) Handles BtnFormPrincipalConsultarAgenda.Click


        With Me
            .PanelFormPrincipalHistorialreservas.Visible = False
            .BtnFormPrincipalHistorialVisitas.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
            .PanelFormPrincipalHorarios.Visible = False
            .BtnFormPrincipalCrearHorarios.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
            .PanelFormPrincipalBuscarInmuebles.Visible = False
            .BtnFormPrincipalBuscar.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
            .PanelFormPrincipalCrearreserva.Visible = False
            .BtnPrincipalCrearreserva.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
            .FormPrincipalRegistroFechanacimiento.Visible = False
            .BtnFormPrincipalRegistrarInteresado.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
            .PanelIngresoPropiedades.Visible = False
            .BtnFormPrincipalIngresoPropiedades.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))



            .PanelFormPrincipalConsultaragenda.Visible = True
            .BtnFormPrincipalConsultarAgenda.ForeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))



        End With

        'CAMBIA EL FORMATO DE LOS DATEPICKER
        DateFormPrincipalCrearreservaFechaReserva.CustomFormat = "yyy-MM-dd"
        DateFormPrincipalCrearreservaFechaReserva.Format = System.Windows.Forms.DateTimePickerFormat.Custom


        Dim fechactual As DateTime = DateTime.Now




        'CARGA LOS NOMBRES DE LOS AGENTES
        BoxFormPrincipalConsultaAgendaAgente.DataSource = (consultaSQL("select concat(concat(Nombre , (' ')) , Apellido ) from personas where Funcion = 'Agente';").Tables(0))
        BoxFormPrincipalConsultaAgendaAgente.DisplayMember = "concat(concat(Nombre , (' ')) , Apellido )"




        Dim fecha As String = fechaactual.Year & "-" & fechaactual.Month & "-" & fechaactual.Day



        If FormLogin.tipoacceso = "Agente" Then
            DataFormPrincipalConsultaragendaDatos.DataSource = consultaSQL("select EstadoReserva , FehaHraVisita as Fecha , concat(concat(Nombre , (' '))  , Apellido ) as Nombre_Apellido , Tel as Telefono , Cel as Celular , Email , Barrios as Barrio , inmueble.Calle , inmueble.NroPuerta as Numero_Puerta, inmueble.NroApto as Numero_Apartamento , reserva.IdReserva from agenda ,  reserva , personas , inmueble , barrios  where  agenda.IdAgenda = reserva.IdAgenda and reserva.IdInteresado = personas.IdPersona and reserva.IdInmueble = inmueble.IdInmueble and barrios.IdBarrio = inmueble.IdBarrio
                and reserva.FehaHraVisita like '%" & fecha & "%' and  agenda.IdAgente = " & IdPersonal & "
                AND EstadoReserva = 'Activo'
                union all select EstadoReserva ,  FehaHraVisita as Fecha , concat(concat(Nombre , (' '))  , Apellido ) as Nombre_Apellido , Tel as Telefono , Cel as Celular , Email , Barrios as Barrio , inmueble.Calle , inmueble.NroPuerta as Numero_Puerta, inmueble.NroApto as Numero_Apartamento , reserva.IdReserva from agenda ,  reserva , personas , inmueble , barrios  where  agenda.IdAgenda = reserva.IdAgenda and reserva.IdInteresado = personas.IdPersona and reserva.IdInmueble = inmueble.IdInmueble and barrios.IdBarrio = inmueble.IdBarrio
                and EstadoReserva = 'Pendiente' and reserva.IdAgenda = 1
                ;").Tables("consultainmueble")

            DataFormPrincipalConsultaragendaDatos.Columns(10).Visible = False

        Else

            Dim IdPersonal As String = consultaSQL("SELECT IdPersona FROM enigma.personas where  concat(concat(Nombre , (' '))  , Apellido ) = '" & BoxFormPrincipalConsultaAgendaAgente.Text & "';").Tables("personas").Rows(0)("IdPersona").ToString()



            'MUESTRA LAS AGENDAS PARA EL DIA DE HOY
            DataFormPrincipalConsultaragendaDatos.DataSource = consultaSQL("select EstadoReserva , FehaHraVisita as Fecha , concat(concat(Nombre , (' '))  , Apellido ) as Nombre_Apellido , Tel as Telefono , Cel as Celular , Email , Barrios as Barrio , inmueble.Calle , inmueble.NroPuerta as Numero_Puerta, inmueble.NroApto as Numero_Apartamento , reserva.IdReserva from agenda ,  reserva , personas , inmueble , barrios  where  agenda.IdAgenda = reserva.IdAgenda and reserva.IdInteresado = personas.IdPersona and reserva.IdInmueble = inmueble.IdInmueble and barrios.IdBarrio = inmueble.IdBarrio
                and reserva.FehaHraVisita like '%" & fecha & "%' and  agenda.IdAgente = " & IdPersonal & "
                AND EstadoReserva = 'Activo'
                union all select EstadoReserva ,  FehaHraVisita as Fecha , concat(concat(Nombre , (' '))  , Apellido ) as Nombre_Apellido , Tel as Telefono , Cel as Celular , Email , Barrios as Barrio , inmueble.Calle , inmueble.NroPuerta as Numero_Puerta, inmueble.NroApto as Numero_Apartamento , reserva.IdReserva from agenda ,  reserva , personas , inmueble , barrios  where  agenda.IdAgenda = reserva.IdAgenda and reserva.IdInteresado = personas.IdPersona and reserva.IdInmueble = inmueble.IdInmueble and barrios.IdBarrio = inmueble.IdBarrio
                and EstadoReserva = 'Pendiente' and reserva.IdAgenda = 1
                ;").Tables("consultainmueble")

            DataFormPrincipalConsultaragendaDatos.Columns(10).Visible = False

        End If



        inicialagenda = Me.DateFormPrincipalConsultaragendaInicio.Value

        Fechaagendaini = Format(inicialagenda, "yyy-MM-dd")


        finalagenda = Me.DateFormPrincipalConsultaragendaFin.Value
        Fechaagendafin = Format(finalagenda, "yyy-MM-dd")






    End Sub


    'BOTON CONSULTAR CANCELACIONES DE AGENDA DE VISITAS
    Private Sub BtnConsultarCancelaciones_Click(sender As Object, e As EventArgs)

        With Me
            .PanelFormPrincipalBuscarInmuebles.Visible = False
            .BtnFormPrincipalBuscar.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
            .PanelFormPrincipalConsultaragenda.Visible = False
            .BtnFormPrincipalConsultarAgenda.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
            .PanelFormPrincipalCrearreserva.Visible = False
            .BtnPrincipalCrearreserva.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
            .FormPrincipalRegistroFechanacimiento.Visible = False
            .BtnFormPrincipalRegistrarInteresado.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
            .PanelFormPrincipalHistorialreservas.Visible = False
            .BtnFormPrincipalHistorialVisitas.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
            .PanelIngresoPropiedades.Visible = False
            .BtnFormPrincipalIngresoPropiedades.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))


        End With


    End Sub


    'BOTON CONSULTAR MODIFICACIONES DE LAS AGENDAS DE VISITAS
    Private Sub BtnConsultarModificaciones_Click(sender As Object, e As EventArgs)
        With Me
            .PanelFormPrincipalBuscarInmuebles.Visible = False
            .BtnFormPrincipalBuscar.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
            .PanelFormPrincipalCrearreserva.Visible = False
            .BtnPrincipalCrearreserva.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
            .PanelFormPrincipalConsultaragenda.Visible = False
            .BtnFormPrincipalConsultarAgenda.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
            .FormPrincipalRegistroFechanacimiento.Visible = False
            .BtnFormPrincipalRegistrarInteresado.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
            .PanelFormPrincipalHistorialreservas.Visible = False
            .BtnFormPrincipalHistorialVisitas.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
            .PanelIngresoPropiedades.Visible = False
            .BtnFormPrincipalIngresoPropiedades.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))

        End With


    End Sub


    'BOTON REGISTRAR INTERESADO
    Private Sub BtnRegistrarInteresado_Click(sender As Object, e As EventArgs) Handles BtnFormPrincipalRegistrarInteresado.Click

        'CARGA INICIALMENTE LOS DEPARTAMENTOS EN EL COMBO BOX
        ComboFormPrincipalRegistroUsuarioDepartamento.Text = ""
        ComboFormPrincipalRegistroUsuarioDepartamento.DataSource = (consultaSQL("SELECT distinct Departamento FROM ubicacion").Tables(0))
        ComboFormPrincipalRegistroUsuarioDepartamento.DisplayMember = "Departamento"



        With Me

            .PanelFormPrincipalConsultaragenda.Visible = False
            .BtnFormPrincipalConsultarAgenda.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
            .PanelFormPrincipalHistorialreservas.Visible = False
            .BtnFormPrincipalHistorialVisitas.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
            .PanelIngresoPropiedades.Visible = False
            .BtnFormPrincipalIngresoPropiedades.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
            .PanelFormPrincipalHorarios.Visible = False
            .BtnFormPrincipalCrearHorarios.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
            .PanelFormPrincipalBuscarInmuebles.Visible = False
            .BtnFormPrincipalBuscar.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
            .PanelFormPrincipalCrearreserva.Visible = False
            .BtnPrincipalCrearreserva.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))


            .FormPrincipalRegistroFechanacimiento.Visible = True
            .BtnFormPrincipalRegistrarInteresado.ForeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))



        End With

        'CARGA LAS CIUDADES EN EL COMBOBOX SEGUN EL DEPARTAMENTO QUE SE SELECCIONO
        ComboFormPrincipalRegistroUsuarioCiudad.Text = ""
        ComboFormPrincipalRegistroUsuarioCiudad.DataSource = (consultaSQL("SELECT distinct NomCiudad FROM barrios , ubicacion where barrios.IdUbicacion = ubicacion.IdUbicacion and ubicacion.Departamento = '" & ComboFormPrincipalRegistroUsuarioDepartamento.Text & "';").Tables(0))
        ComboFormPrincipalRegistroUsuarioCiudad.DisplayMember = "NomCiudad"






    End Sub


    Private Sub BtnHistorialVisitas_Click(sender As Object, e As EventArgs) Handles BtnFormPrincipalHistorialVisitas.Click
        With Me

            .PanelFormPrincipalHistorialreservas.Visible = False
            .BtnFormPrincipalHistorialVisitas.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
            .PanelFormPrincipalHorarios.Visible = False
            .BtnFormPrincipalCrearHorarios.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
            .PanelFormPrincipalBuscarInmuebles.Visible = False
            .BtnFormPrincipalBuscar.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
            .PanelFormPrincipalCrearreserva.Visible = False
            .BtnPrincipalCrearreserva.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
            .FormPrincipalRegistroFechanacimiento.Visible = False
            .BtnFormPrincipalRegistrarInteresado.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
            .PanelIngresoPropiedades.Visible = False
            .BtnFormPrincipalIngresoPropiedades.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
            .PanelFormPrincipalConsultaragenda.Visible = False
            .BtnFormPrincipalConsultarAgenda.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))


            .PanelFormPrincipalHistorialreservas.Visible = True
            .BtnFormPrincipalHistorialVisitas.ForeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))


        End With



        DataFormPrincipalHistorialreservas.DataSource = consultaSQL("SELECT inmueble.IdInmueble , sucursales.NomSucursal as Sucursal , concat(agent.Apellido , (' ')  , agent.Nombre ) as 'Agente' ,  concat(cli.Apellido , (' ') , cli.Nombre )  as 'Cliente' , ubicacion.Departamento , barrios.NomCiudad as 'Nombre ciudad' ,  notifica.FechaNotifi as 'Fecha notificacion' , notifica.TipoNotif as 'Tipo notificacion' , notifica.FechaNotifi as 'Fecha notificacion' , notifica.Texto as 'Texto notificacion'

FROM personas agent , personas cli , notifica , sucursales , reserva , inmueble , barrios , ubicacion

where notifica.IdEnvia = agent.IdPersona and notifica.IdRecibe = cli.IdPersona and agent.IdSucursal = sucursales.IdSucursal
and reserva.IdReserva = notifica.IdReserva and inmueble.IdInmueble = reserva.IdInmueble and inmueble.IdBarrio = barrios.IdBarrio and barrios.IdUbicacion = ubicacion.IdUbicacion
order by agent.Apellido , FechaNotifi
;").Tables("consultainmueble")

        DataFormPrincipalHistorialreservas.Columns(0).Visible = False


    End Sub



    Private Sub BtnIngresoPropiedades_Click(sender As Object, e As EventArgs) Handles BtnFormPrincipalIngresoPropiedades.Click





        With Me

            .PanelFormPrincipalConsultaragenda.Visible = False
            .BtnFormPrincipalConsultarAgenda.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
            .PanelFormPrincipalHistorialreservas.Visible = False
            .BtnFormPrincipalHistorialVisitas.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
            .PanelFormPrincipalHorarios.Visible = False
            .BtnFormPrincipalCrearHorarios.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
            .PanelFormPrincipalBuscarInmuebles.Visible = False
            .BtnFormPrincipalBuscar.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
            .PanelFormPrincipalCrearreserva.Visible = False
            .BtnPrincipalCrearreserva.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
            .FormPrincipalRegistroFechanacimiento.Visible = False
            .BtnFormPrincipalRegistrarInteresado.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))


            .PanelIngresoPropiedades.Visible = True
            .BtnFormPrincipalIngresoPropiedades.ForeColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))



        End With




        'Carga los departamentos que hay en la bd
        ComboFormPrincipalIngresoPropiedadesDepartamento.DataSource = (consultaSQL("SELECT Departamento FROM ubicacion").Tables(0))
            ComboFormPrincipalIngresoPropiedadesDepartamento.DisplayMember = "Departamento"

            If CheckFormPrincipalIngresoPropiedadesJardin.Checked = False Then
                jardin = "No"
            Else
                jardin = "Si"
            End If

            If CheckFormPrincipalIngresoPropiedadesPiscina.Checked = False Then
                piscina = "No"
            Else
                piscina = "Si"
            End If

            If CheckFormPrincipalIngresoPropiedadesGarage.Checked = False Then
                garage = "No"
            Else
                garage = "Si"
            End If

            If CheckFormPrincipalIngresoPropiedadesAmoblada.Checked = False Then
                amoblada = "No"
            Else
                amoblada = "Si"
            End If

            If CheckFormPrincipalIngresoPropiedadesBarbacoa.Checked = False Then
                barbacoa = "No"
            Else
                barbacoa = "Si"
            End If

            If CheckFormPrincipalIngresoPropiedadesMascotas.Checked = False Then
                mascotas = "No"
            Else
                mascotas = "Si"
            End If



            'CARGA LAS CIUDADES EN EL COMBOBOX SEGUN EL DEPARTAMENTO QUE SE SELECCIONO
            ComboFormPrincipalIngresoPropiedadesCiudad.Text = ""
            ComboFormPrincipalIngresoPropiedadesCiudad.DataSource = (consultaSQL("SELECT distinct NomCiudad FROM barrios , ubicacion where barrios.IdUbicacion = ubicacion.IdUbicacion and ubicacion.Departamento = '" & ComboFormPrincipalIngresoPropiedadesDepartamento.Text & "';").Tables(0))
            ComboFormPrincipalIngresoPropiedadesCiudad.DisplayMember = "NomCiudad"


            ComboFormPrincipalIngresoPropiedadesBarrios.Text = ""
            ComboFormPrincipalIngresoPropiedadesBarrios.DataSource = (consultaSQL("select distinct Barrios from barrios where NomCiudad = '" & ComboFormPrincipalIngresoPropiedadesCiudad.Text & "';").Tables(0))
            ComboFormPrincipalIngresoPropiedadesBarrios.DisplayMember = "Barrios"






        DataFormPrincipalIngresoImagenes.Rows.Clear()
        ComboFormPrincipalIngresoPropiedadesTipodoc.Text = ""
        TxtFormPrincipalIngresoPropiedadesDocumentopropietario.Text = ""
            TxtFormPrincipalIngresoPropiedadesCalle.Text = ""
            TxtFormPrincipalIngresoPropiedadesEsquina.Text = ""
            NumericFormPrincipalIngresoPropiedadesNumeopuerta.Value = 0
            NumericFormPrincipalIngresoPropiedadesNumeroAp.Value = 0
            CheckFormPrincipalIngresoPropiedadesBarbacoa.Checked = False
            CheckFormPrincipalIngresoPropiedadesJardin.Checked = False
            CheckFormPrincipalIngresoPropiedadesGarage.Checked = False
            CheckFormPrincipalIngresoPropiedadesPiscina.Checked = False
            CheckFormPrincipalIngresoPropiedadesAmoblada.Checked = False
            CheckFormPrincipalIngresoPropiedadesMascotas.Checked = False
            TxtFormPrincipalIngresoPropiedadesDescripcion.Text = ""
            TxtFormPrincipalIngresoPropiedadesNombreInmueble.Text = ""
            NumericTxtFormPrincipalIngresoPropiedadesSuperficieTotal.Value = 0
            NumericTxtFormPrincipalIngresoPropiedadesSuperficieEdificada.Value = 0
            NumericFormPrincipalIngresoPropiedadesCantbaños.Value = 0
            NumericFormPrincipalIngresoPropiedadesCantcuartos.Value = 0
            ComboFormPrincipalIngresoPropiedadesTipoInmueble.Text = ""
            ComboFormPrincipalIngresoPropiedadesGarantia.Text = ""
            ComboFormPrincipalIngresoPropiedadesTipopubli.Text = ""
            NumericComboFormPrincipalIngresoPropiedadesPrecio.Value = 0
            ComboFormPrincipalIngresoPropiedadesMoneda.Text = ""

            Idinmueble = ""









    End Sub







    '------------------------------------------------------------------------------------------------------------------------------

    'BOTON CERRAR , Y MUESTRA EL LOGIN
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PickFormPrincipalCerrar.Click
        Application.Exit()

    End Sub
    'BOTON MINIMIZAR
    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PickFormPrincipalMinimizar.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    'BOTON MAXIMIZAR
    Private Sub PickMaximizar_Click(sender As Object, e As EventArgs) Handles PickFormPrincipalMaximizar.Click


        Me.WindowState = FormWindowState.Maximized
        PickFormPrincipalMaximizar.Visible = False
        PickFormPrincipalRestaurar.Visible = True

        'REDIMENCIONA LOS OBJETOS AL MAXIMIZAR LA VENTANA
        With PanelFormPrincipalBuscarInmuebles

            LabelBuscarInmueblesTitulo.Location = New System.Drawing.Point(PanelFormPrincipalBuscarInmuebles.Size.Width / 3, PanelFormPrincipalBuscarInmuebles.Size.Height / 10)
            TabFormPrincipalBuscarInmueblesTipoInteres.Location = New System.Drawing.Point(PanelFormPrincipalBuscarInmuebles.Size.Width / 6, PanelFormPrincipalBuscarInmuebles.Size.Height / 6)
            PanelFormPrincipalBuscarInmueblesCaracteristicas.Height = PanelFormPrincipalBuscarInmuebles.Size.Height - TabFormPrincipalBuscarInmueblesTipoInteres.Height * 2

        End With




    End Sub

    'BOTON RESTAURAR
    Private Sub PickRestaurar1_Click(sender As Object, e As EventArgs) Handles PickFormPrincipalRestaurar.Click

        Me.WindowState = FormWindowState.Normal
        PickFormPrincipalMaximizar.Visible = True
        PickFormPrincipalRestaurar.Visible = False

        'REDIMENCIONA LOS OBJETOS AL RESTAURAR LA VENTANA
        With PanelFormPrincipalBuscarInmuebles
            LabelBuscarInmueblesTitulo.Location = New System.Drawing.Point(PanelFormPrincipalBuscarInmuebles.Size.Width / 3, PanelFormPrincipalBuscarInmuebles.Size.Height / 10)
            TabFormPrincipalBuscarInmueblesTipoInteres.Location = New System.Drawing.Point(PanelFormPrincipalBuscarInmuebles.Size.Width / 6, PanelFormPrincipalBuscarInmuebles.Size.Height / 6)
            PanelFormPrincipalBuscarInmueblesCaracteristicas.Height = PanelFormPrincipalBuscarInmuebles.Size.Height - TabFormPrincipalBuscarInmueblesTipoInteres.Height * 2
        End With

    End Sub




    'CERRAR SESION
    Private Sub PicCerrarSesion_Click(sender As Object, e As EventArgs) Handles PickCerrarSesion.Click
        FormLogin.TxtFormLoginUsuario.Text = "Usuario"
        FormLogin.TxtFormLoginContraseña.Text = "Contraseña"
        Me.Close()

        FormLogin.Show()
    End Sub

    Private Sub LabCerrarSesion_Click(sender As Object, e As EventArgs) Handles LabelCerrarSesion.Click
        Me.Close()
        FormLogin.TxtFormLoginUsuario.Text = "Usuario"
        FormLogin.TxtFormLoginContraseña.Text = "Contraseña"
        FormLogin.Show()
    End Sub



    'CODIGOS PARA MOVER EL FORMULARIO AL HACER CLICK EN FORMULARIO
    <DllImport("user32.DLL", EntryPoint:="ReleaseCapture")>
    Private Shared Sub ReleaseCapture()
    End Sub

    <DllImport("user32.DLL", EntryPoint:="SendMessage")>
    Private Shared Sub SendMessage(ByVal hWnd As System.IntPtr, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer)
    End Sub

    Private Sub PanelFunciones_MouseMove(sender As Object, e As MouseEventArgs) Handles PanelFormPrincipalFunciones.MouseMove
        ReleaseCapture()
        SendMessage(Me.Handle, &H112&, &HF012&, 0)
    End Sub



    Private Sub BtnBuscarInmueblesAlquilerBuscar_Click(sender As Object, e As EventArgs) Handles BtnFormPrincipalBuscarInmueblesAlquilerBuscar.Click



        Me.DataFormPrincipalBuscarInmuebles.DataSource = consultaSQL("select IdInmueble , TipoInm, Departamento, NumDormitorios , TipoPubli , Precio , Moneda , Jardin , Amoblada , Garage , Piscina , Barbacoa
from ubicacion, caracteristicas_inm, inmueble, barrios, publicacion 	
where caracteristicas_inm.IdInm = inmueble.IdInm and ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio and publicacion.IdPubli = inmueble.IdPubli
and TipoInm = '" & BoxFormPrincipalBuscarInmueblesAlquilerTipoPropiedad.SelectedItem & "' and Departamento = '" & BoxFormPrincipalBuscarInmueblesAlquilerUbicacion.SelectedItem & "' and NumDormitorios = '" & BoxFormPrincipalBuscarInmueblesAlquilerHabitaciones.SelectedItem & "'
and TipoPubli = 'Alquiler'

;").Tables("consultainmueble")





    End Sub

    Private Sub ButtonFormPrincipalIngresoPropiedadesAgregarimagen_Click(sender As Object, e As EventArgs) Handles ButtonFormPrincipalIngresoPropiedadesAgregarimagen.Click


        Try


            If cantidadimagenes = 0 Then
                'Abre la busqueda de imagenes que se encuentran en la computadora
                If OpenFileImagenes.ShowDialog <> Windows.Forms.DialogResult.Cancel Then

                    'Agrega esa imagen al datagridview de imagenes (se agrega toda las veces que el usuario desee)
                    DataFormPrincipalIngresoImagenes.Rows.Add()
                    DataFormPrincipalIngresoImagenes.Rows(cantidadimagenes).Cells(0).Value = Image.FromFile(OpenFileImagenes.FileName)
                    DataFormPrincipalIngresoImagenes.Rows(cantidadimagenes).Height = 80
                End If

                cantidadimagenes = cantidadimagenes + 1

            Else


                'Abre la busqueda de imagenes que se encuentran en la computadora
                If OpenFileImagenes.ShowDialog <> Windows.Forms.DialogResult.Cancel Then

                    'Agrega esa imagen al datagridview de imagenes (se agrega toda las veces que el usuario desee)
                    DataFormPrincipalIngresoImagenes.Rows.Add()
                    DataFormPrincipalIngresoImagenes.Rows(cantidadimagenesedit).Cells(0).Value = Image.FromFile(OpenFileImagenes.FileName)
                    DataFormPrincipalIngresoImagenes.Rows(cantidadimagenesedit).Height = 80
                End If

                cantidadimagenesedit = cantidadimagenesedit + 1




            End If




        Catch ex As Exception

        End Try





    End Sub


    Private Sub ButtonFormPrincipalIngresoPropiedadesIngresar_Click(sender As Object, e As EventArgs) Handles ButtonFormPrincipalIngresoPropiedadesIngresar.Click

        'Convierte en integer el id de la persona que quiere publicar el inmueble
        Try
            Dim IdPersona As String = consultaSQL("SELECT IdPersona FROM personas where TipoDoc = '" & ComboFormPrincipalIngresoPropiedadesTipodoc.Text & "' and NumDoc = '" & TxtFormPrincipalIngresoPropiedadesDocumentopropietario.Text & "';").Tables("personas").Rows(0)("IdPersona").ToString()
            IdPersonaInt = CInt(IdPersona)
        Catch ex As Exception
            MsgBox("No hay persona registrada con ese numero de documento")
        End Try



        'Si existe persona registrada con el tipo de documento y el documento registra el inmueble
        If IdPersonaInt <> 0 Then

            'Convierte en integer el id de barrio que obtiene de la bd que se selecciono
            Try
                Dim IdBarrio As String = consultaSQL("select IdBarrio from barrios where barrios.Barrios = '" & ComboFormPrincipalIngresoPropiedadesBarrios.Text & "';").Tables("quebarrioshay").Rows(0)("IdBarrio").ToString()
                IdBarrioinmuebleInt = CInt(IdBarrio)
            Catch ex As Exception
                MsgBox("No se selecciono ningun barrio")
            End Try


            Try
                'devuelve el id del ultimo inmueble ingresado
                Dim IdInmueble As String = consultaSQL("Select IdInmueble from inmueble order by IdInmueble desc limit 1;").Tables("inmueble").Rows(0)("IdInmueble").ToString()
                'Convierte en integer el id del ultimo inmueble
                IdInmuebleInt = CInt(IdInmueble)



            Catch ex As Exception
                IdInmuebleInt = 1
            End Try







            If Idinmueble = "" Then





                'Insertar precio de publicacion
                updateSQL("insert into publicacion ( IdPubli , TipoPubli  , Moneda , Precio , EstadoInm  ) values ( " & IdInmuebleInt + 1 & " , '" & ComboFormPrincipalIngresoPropiedadesTipopubli.Text & "' , '" & ComboFormPrincipalIngresoPropiedadesMoneda.Text & "' , " & NumericComboFormPrincipalIngresoPropiedadesPrecio.Value & " , '" & ComboBox3.Text & "' );")
                'Insertar caracteristicas de inmueble


                updateSQL("insert into caracteristicas_inm ( IdInm , TipoInm , SupParcela , SupConstruida , NumDormitorios , NumBaños , Jardin , Piscina , Garage , Amoblada , Barbacoa , AceptaMascota , Garantia , Descripción) values ( " & IdInmuebleInt + 1 & " , '" & ComboFormPrincipalIngresoPropiedadesTipoInmueble.Text & "' , " & NumericTxtFormPrincipalIngresoPropiedadesSuperficieTotal.Value & " , " & NumericTxtFormPrincipalIngresoPropiedadesSuperficieEdificada.Value & " , " & NumericFormPrincipalIngresoPropiedadesCantcuartos.Value & " , " & NumericFormPrincipalIngresoPropiedadesCantbaños.Value & " , '" & jardin & "' , '" & piscina & "' , '" & garage & "' , '" & amoblada & "' , '" & barbacoa & "' , '" & mascotas & "' , '" & ComboFormPrincipalIngresoPropiedadesGarantia.Text & "' , '" & TxtFormPrincipalIngresoPropiedadesDescripcion.Text & "');")


                'Insertar datos de inmueble
                updateSQL("insert into inmueble ( IdInmueble , IdInm , IdPubli, IdBarrio ,  Calle , Esquina , NroPuerta , NroApto , NomInmueble  ) values ( " & IdInmuebleInt + 1 & " ,  " & IdInmuebleInt + 1 & " , " & IdInmuebleInt + 1 & " , " & IdBarrioinmuebleInt & " , '" & TxtFormPrincipalIngresoPropiedadesCalle.Text & "' , '" & TxtFormPrincipalIngresoPropiedadesEsquina.Text & "' , " & NumericFormPrincipalIngresoPropiedadesNumeopuerta.Value & " , " & NumericFormPrincipalIngresoPropiedadesNumeroAp.Value & " , '" & TxtFormPrincipalIngresoPropiedadesNombreInmueble.Text & "');")


                'Asigna el inmueble a la persona asociada al documento ingresado
                updateSQL("insert into tiene ( IdPersona ,  IdInmueble ) values ( " & IdPersonaInt & " , " & IdInmuebleInt + 1 & " )")




                'Agrega todas las imagenes que se encuentran en el datagridview a la bd
                For index As Integer = 0 To cantidadimagenes - 1
                    Dim ms As New MemoryStream
                    Dim img As Bitmap

                    Dim sql As String = "Insert Into img_inmuebles (IdInmueble , Imagen) Values ( " & IdInmuebleInt + 1 & " , @imagen)"


                    PtbImagen.Image = Nothing
                    PtbImagen.Refresh()

                    img = DataFormPrincipalIngresoImagenes.Rows(index).Cells(0).Value

                    img.Save(ms, ImageFormat.Jpeg)
                    PtbImagen.Image = Image.FromStream(ms)
                    agregarmodificarImagenes(sql, PtbImagen)

                Next
                cantidadimagenes = 0





                'Envia mail al dueño del inmueble informando sobre la correcta publicacion del inmueble
                Dim Email As String = consultaSQL("SELECT Email from personas where IdPersona = " & IdPersonaInt & ";").Tables("personas").Rows(0)("Email").ToString()
                If EnvioMail.enviomail(Email, "RegistroInmueble") = True Then



                    Dim fechahora = fechaactual.Year & "-" & fechaactual.Month & "-" & fechaactual.Day & " " & horaactual

                    'Guarda la notificacion en la bd
                    updateSQL(" insert into notifica ( IdEnvia, IdRecibe, FechaNotifi, TipoNotif, Texto) values ( " & IdPersonal & " , " & IdPersonaInt & " , '" & fechahora & "' , 'Email' , '" & TxtFormPrincipalIngresoPropiedadesNotificacion.Text & "');")

                    With Me

                        .ComboFormPrincipalIngresoPropiedadesTipodoc.Text = ""
                        .TxtFormPrincipalIngresoPropiedadesDocumentopropietario.Text = ""
                        .ComboFormPrincipalIngresoPropiedadesDepartamento.Text = ""
                        .ComboFormPrincipalIngresoPropiedadesCiudad.Text = ""
                        .ComboFormPrincipalIngresoPropiedadesBarrios.Text = ""
                        .TxtFormPrincipalIngresoPropiedadesCalle.Text = ""
                        .TxtFormPrincipalIngresoPropiedadesEsquina.Text = ""
                        .NumericFormPrincipalIngresoPropiedadesNumeopuerta.Value = 0
                        .NumericFormPrincipalIngresoPropiedadesNumeroAp.Value = 0
                        .CheckFormPrincipalIngresoPropiedadesBarbacoa.Checked = False
                        .CheckFormPrincipalIngresoPropiedadesJardin.Checked = False
                        .CheckFormPrincipalIngresoPropiedadesGarage.Checked = False
                        .CheckFormPrincipalIngresoPropiedadesPiscina.Checked = False
                        .CheckFormPrincipalIngresoPropiedadesAmoblada.Checked = False
                        .CheckFormPrincipalIngresoPropiedadesMascotas.Checked = False
                        .TxtFormPrincipalIngresoPropiedadesNombreInmueble.Text = ""
                        .NumericTxtFormPrincipalIngresoPropiedadesSuperficieTotal.Value = 0
                        .NumericTxtFormPrincipalIngresoPropiedadesSuperficieEdificada.Value = 0
                        .NumericFormPrincipalIngresoPropiedadesCantbaños.Value = 0
                        .NumericFormPrincipalIngresoPropiedadesCantcuartos.Value = 0
                        .ComboFormPrincipalIngresoPropiedadesTipoInmueble.Text = ""
                        .ComboFormPrincipalIngresoPropiedadesGarantia.Text = ""
                        .ComboFormPrincipalIngresoPropiedadesTipopubli.Text = ""
                        .NumericComboFormPrincipalIngresoPropiedadesPrecio.Value = 0
                        .ComboFormPrincipalIngresoPropiedadesMoneda.Text = ""
                        .TxtFormPrincipalIngresoPropiedadesNotificacion.Text = "Le informamos que ya ah quedado publicado el inmueble."
                        .TxtFormPrincipalIngresoPropiedadesDescripcion.Text = ""
                        .PtbImagen.Image = Nothing
                        .DataFormPrincipalIngresoImagenes.Rows.Clear()
                        BtnFormPrincipalIngresoPropiedades.PerformClick()



                    End With

                Else


                    With Me

                        .ComboFormPrincipalIngresoPropiedadesTipodoc.Text = ""
                        .TxtFormPrincipalIngresoPropiedadesDocumentopropietario.Text = ""
                        .ComboFormPrincipalIngresoPropiedadesDepartamento.Text = ""
                        .ComboFormPrincipalIngresoPropiedadesCiudad.Text = ""
                        .ComboFormPrincipalIngresoPropiedadesBarrios.Text = ""
                        .TxtFormPrincipalIngresoPropiedadesCalle.Text = ""
                        .TxtFormPrincipalIngresoPropiedadesEsquina.Text = ""
                        .NumericFormPrincipalIngresoPropiedadesNumeopuerta.Value = 0
                        .NumericFormPrincipalIngresoPropiedadesNumeroAp.Value = 0
                        .CheckFormPrincipalIngresoPropiedadesBarbacoa.Checked = False
                        .CheckFormPrincipalIngresoPropiedadesJardin.Checked = False
                        .CheckFormPrincipalIngresoPropiedadesGarage.Checked = False
                        .CheckFormPrincipalIngresoPropiedadesPiscina.Checked = False
                        .CheckFormPrincipalIngresoPropiedadesAmoblada.Checked = False
                        .CheckFormPrincipalIngresoPropiedadesMascotas.Checked = False
                        .TxtFormPrincipalIngresoPropiedadesNombreInmueble.Text = ""
                        .NumericTxtFormPrincipalIngresoPropiedadesSuperficieTotal.Value = 0
                        .NumericTxtFormPrincipalIngresoPropiedadesSuperficieEdificada.Value = 0
                        .NumericFormPrincipalIngresoPropiedadesCantbaños.Value = 0
                        .NumericFormPrincipalIngresoPropiedadesCantcuartos.Value = 0
                        .ComboFormPrincipalIngresoPropiedadesTipoInmueble.Text = ""
                        .ComboFormPrincipalIngresoPropiedadesGarantia.Text = ""
                        .ComboFormPrincipalIngresoPropiedadesTipopubli.Text = ""
                        .NumericComboFormPrincipalIngresoPropiedadesPrecio.Value = 0
                        .ComboFormPrincipalIngresoPropiedadesMoneda.Text = ""
                        .TxtFormPrincipalIngresoPropiedadesNotificacion.Text = "Le informamos que ya ah quedado publicado el inmueble"
                        .TxtFormPrincipalIngresoPropiedadesDescripcion.Text = ""
                        .PtbImagen.Image = Nothing
                        .DataFormPrincipalIngresoImagenes.Rows.Clear()
                        BtnFormPrincipalIngresoPropiedades.PerformClick()


                    End With

                    Dim fechahora = fechaactual.Year & "-" & fechaactual.Month & "-" & fechaactual.Day & " " & horaactual
                    'Guarda la notificacion en la bd con error de envio
                    updateSQL(" insert into notifica ( IdEnvia, IdRecibe, FechaNotifi, TipoNotif, Texto) values ( " & IdPersonal & " , " & IdPersonaInt & " , '" & fechahora & "' , 'Email' , 'Error notificacion Email - Registro propiedad');")



                End If








            Else

                '-------------------------------------------------------------------------------------------------------------------------------------------------------------+




                'Insertar precio de publicacion
                updateSQL("update publicacion set TipoPubli = '" & ComboFormPrincipalIngresoPropiedadesTipopubli.Text & "' , Moneda = '" & ComboFormPrincipalIngresoPropiedadesMoneda.Text & "' , Precio = " & NumericComboFormPrincipalIngresoPropiedadesPrecio.Value & " , EstadoInm = '" & ComboBox3.Text & "' where IdPubli = " & Idinmueble & ";")



                'Insertar caracteristicas de inmueble
                updateSQL("update caracteristicas_inm set TipoInm = '" & ComboFormPrincipalIngresoPropiedadesTipoInmueble.Text & "' , SupParcela = " & NumericTxtFormPrincipalIngresoPropiedadesSuperficieTotal.Value & " , SupConstruida = " & NumericTxtFormPrincipalIngresoPropiedadesSuperficieEdificada.Value & " , NumDormitorios = " & NumericFormPrincipalIngresoPropiedadesCantcuartos.Value & " , NumBaños = " & NumericFormPrincipalIngresoPropiedadesCantbaños.Value & " , Jardin = '" & jardin & "' , Piscina = '" & piscina & "' , Garage = '" & garage & "' , Amoblada = '" & amoblada & "' , Barbacoa = '" & barbacoa & "' , AceptaMascota = '" & mascotas & "' , Garantia = '" & ComboFormPrincipalIngresoPropiedadesGarantia.Text & "' , Descripción = '" & TxtFormPrincipalIngresoPropiedadesDescripcion.Text & "' where IdInm = " & Idinmueble & ";")



                'Insertar datos de inmueble 
                updateSQL("update inmueble set IdInm = " & Idinmueble & " , IdPubli = " & Idinmueble & " , IdBarrio = " & IdBarrioinmuebleInt & " , Calle = '" & TxtFormPrincipalIngresoPropiedadesCalle.Text & "' , Esquina = '" & TxtFormPrincipalIngresoPropiedadesEsquina.Text & "' , NroPuerta = " & NumericFormPrincipalIngresoPropiedadesNumeopuerta.Value & " , NroApto = " & NumericFormPrincipalIngresoPropiedadesNumeroAp.Value & " , NomInmueble = '" & TxtFormPrincipalIngresoPropiedadesNombreInmueble.Text & "' where IdInmueble = " & Idinmueble & " ;")



                'Asigna el inmueble a la persona asociada al documento ingresado
                updateSQL("update tiene set IdPersona = " & IdPersonaInt & " where IdInmueble = " & Idinmueble & ";")




                updateSQL("delete from img_inmuebles where IdInmueble = " & Idinmueble & ";")



                'Agrega todas las imagenes que se encuentran en el datagridview a la bd
                For index As Integer = 0 To cantidadimagenesedit - 1
                    Dim ms As New MemoryStream
                    Dim img As Bitmap

                    Dim sql As String = "Insert Into img_inmuebles (IdInmueble , Imagen) Values ( " & Idinmueble & " , @imagen)"


                    PtbImagen.Image = Nothing
                    PtbImagen.Refresh()

                    img = DataFormPrincipalIngresoImagenes.Rows(index).Cells(0).Value

                    img.Save(ms, ImageFormat.Jpeg)
                    PtbImagen.Image = Image.FromStream(ms)
                    agregarmodificarImagenes(sql, PtbImagen)

                Next
                cantidadimagenesedit = 0


                With Me

                    .ComboFormPrincipalIngresoPropiedadesTipodoc.Text = ""
                    .TxtFormPrincipalIngresoPropiedadesDocumentopropietario.Text = ""
                    .ComboFormPrincipalIngresoPropiedadesDepartamento.Text = ""
                    .ComboFormPrincipalIngresoPropiedadesCiudad.Text = ""
                    .ComboFormPrincipalIngresoPropiedadesBarrios.Text = ""
                    .TxtFormPrincipalIngresoPropiedadesCalle.Text = ""
                    .TxtFormPrincipalIngresoPropiedadesEsquina.Text = ""
                    .NumericFormPrincipalIngresoPropiedadesNumeopuerta.Value = 0
                    .NumericFormPrincipalIngresoPropiedadesNumeroAp.Value = 0
                    .CheckFormPrincipalIngresoPropiedadesBarbacoa.Checked = False
                    .CheckFormPrincipalIngresoPropiedadesJardin.Checked = False
                    .CheckFormPrincipalIngresoPropiedadesGarage.Checked = False
                    .CheckFormPrincipalIngresoPropiedadesPiscina.Checked = False
                    .CheckFormPrincipalIngresoPropiedadesAmoblada.Checked = False
                    .CheckFormPrincipalIngresoPropiedadesMascotas.Checked = False
                    .TxtFormPrincipalIngresoPropiedadesNombreInmueble.Text = ""
                    .NumericTxtFormPrincipalIngresoPropiedadesSuperficieTotal.Value = 0
                    .NumericTxtFormPrincipalIngresoPropiedadesSuperficieEdificada.Value = 0
                    .NumericFormPrincipalIngresoPropiedadesCantbaños.Value = 0
                    .NumericFormPrincipalIngresoPropiedadesCantcuartos.Value = 0
                    .ComboFormPrincipalIngresoPropiedadesTipoInmueble.Text = ""
                    .ComboFormPrincipalIngresoPropiedadesGarantia.Text = ""
                    .ComboFormPrincipalIngresoPropiedadesTipopubli.Text = ""
                    .NumericComboFormPrincipalIngresoPropiedadesPrecio.Value = 0
                    .ComboFormPrincipalIngresoPropiedadesMoneda.Text = ""
                    .TxtFormPrincipalIngresoPropiedadesNotificacion.Text = "Le informamos que ya ah quedado publicado el inmueble."
                    .TxtFormPrincipalIngresoPropiedadesDescripcion.Text = ""
                    .PtbImagen.Image = Nothing
                    .DataFormPrincipalIngresoImagenes.Rows.Clear()
                    BtnFormPrincipalIngresoPropiedades.PerformClick()



                End With





                Idinmueble = ""

            End If




















        End If







    End Sub




    'falta boton
    Private Sub Button1_Click(sender As Object, e As EventArgs)

        DataFormPrincipalIngresoImagenes.Rows.Remove(DataFormPrincipalIngresoImagenes.CurrentRow)

    End Sub

    'Agranda la imagen que se selecciono, para poder identificarla y luego eliminarla
    Private Sub DataFormPrincipalIngresoImagenes_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataFormPrincipalIngresoImagenes.CellClick




        DataFormPrincipalIngresoImagenes.CurrentRow.Height = 150

            Dim result As DialogResult = MessageBox.Show("Desea eliminar la imagen?",
                              "Eliminar imagen",
                              MessageBoxButtons.YesNo)

            If result = DialogResult.Yes Then



            DataFormPrincipalIngresoImagenes.Rows.Remove(DataFormPrincipalIngresoImagenes.CurrentRow)

            cantidadimagenesedit = cantidadimagenesedit - 1

            If cantidadimagenes <> 0 Then

                cantidadimagenes = cantidadimagenes - 1

            End If



        End If

            If result = DialogResult.No Then

            DataFormPrincipalIngresoImagenes.CurrentRow.Height = 80




        End If





    End Sub


    'REGISTRO DE USUARIO EN BD
    Private Sub ButtonFormPrincipalRegistrarUsuarioIngresar_Click(sender As Object, e As EventArgs) Handles ButtonFormPrincipalRegistrarUsuarioIngresar.Click





        'Convierte en integer el id de barrio que obtiene de la bd que se selecciono
        Try
            Dim IdBarrio As String = consultaSQL("select IdBarrio from barrios where barrios.Barrios = '" & ComboFormPrincipalRegistroUsuarioBarrio.Text & "';").Tables("quebarrioshay").Rows(0)("IdBarrio").ToString()
            IdBarriopersonaInt = CInt(IdBarrio)

            Try

                'averigua el id de la sucursal para asignar a la persona segun donde se encuentre la sucursal
                Dim idsucusral As String = consultaSQL("SELECT IdSucursal FROM sucursales , barrios , ubicacion where sucursales.IdBarrio = barrios.IdBarrio and ubicacion.IdUbicacion = barrios.IdUbicacion and ubicacion.Departamento = '" & ComboFormPrincipalRegistroUsuarioDepartamento.Text & "' ;").Tables("sucursal").Rows(0)("IdSucursal").ToString()
                idsucusralInt = CInt(idsucusral)

                'convierte la fecha de nacimiento en un formato que la bd entiende y deja almacenar (separado por guiones)
                Dim fechanacimiento As String = DateFormPrincipalRegistroUsuarioFechaNacimiento.Value.Date.Year & "-" & DateFormPrincipalRegistroUsuarioFechaNacimiento.Value.Date.Month & "-" & DateFormPrincipalRegistroUsuarioFechaNacimiento.Value.Date.Day


                'Primero valida si el mail es correcto para lueo ingresar en bd
                Dim expresionmail As String = "^[A-Za-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?$"
                Dim r As New Regex(expresionmail)
                If r.IsMatch(TxtFormPrincipalRegistrarUsuarioEmail.Text) = False Then
                    MsgBox("Mail incorrecto")
                Else


                    'Asigna la funcion, si es agente solamente va a poder ingresar clientes, si es gerente o admin puede agregar agentes
                    Dim funcion As String
                    If FormLogin.tipoacceso = "Agente" Then
                        funcion = ComboFormPrincipalRegistroUsuarioTipoPersona.Text
                    Else
                        funcion = ComboFormPrincipalRegistroUsuarioFuncion.Text
                    End If


                    'Validaciones de ingreso de datos
                    If ComboFormPrincipalRegistrarUsuarioTipoDocumento.Text = "" Then
                        MsgBox("No se ingreso el tipo de documento")
                    Else
                        If TxtFormPrincipalRegistrarUsuarioDocumento.Text = "" Then
                            MsgBox("No se ingreso documento")
                        Else
                            If TxtFormPrincipalRegistrarUsuarioNombre.Text = "" Then
                                MsgBox("No se ingreso nombre de usuario")
                            Else
                                If TxtFormPrincipalRegistrarUsuarioApellido.Text = "" Then
                                    MsgBox("No se ingreso apellido de usuario")
                                Else
                                    If NumericFormPrincipalRegistrarUsuarioCelular.Value = 0 Then
                                        MsgBox("No se ingreso telefono")
                                    Else
                                        If NumericFormPrincipalRegistrarUsuarioMontosueldo.Value = 0 Then
                                            MsgBox("No se ingreso monto de recibo de sueldo")
                                        Else
                                            '--------------------------------INSERTA EN LA BD UN NUEVO USUARIO-----------------------------------------
                                            updateSQL("insert into personas ( IdSucursal, IdBarrio , TipoDoc, NumDoc, Nombre, Apellido, Fechanac, Tel, Cel, Calle, Esquina, NroApto, NroPuerta, NomInmueble, Email, Funcion, ReciboSueldo ) values ( " & idsucusralInt & " ," & IdBarriopersonaInt & " , '" & ComboFormPrincipalRegistrarUsuarioTipoDocumento.Text & "' , '" & TxtFormPrincipalRegistrarUsuarioDocumento.Text & "' , '" & TxtFormPrincipalRegistrarUsuarioNombre.Text & "' , '" & TxtFormPrincipalRegistrarUsuarioApellido.Text & "' , '" & fechanacimiento & "' , " & NumericFormPrincipalRegistrarUsuarioTelefono.Value & " , " & NumericFormPrincipalRegistrarUsuarioCelular.Value & " , '" & TxtFormPrincipalRegistrarUsuarioCalle.Text & "' , '" & TxtFormPrincipalRegistrarUsuarioEsquina.Text & "' , " & NumericFormPrincipalRegistrarUsuarioNumeroApartamento.Value & " , " & NumericFormPrincipalRegistrarUsuarioNumeroPuerta.Value & " , '" & TxtFormPrincipalRegistrarUsuarioNombreInmueble.Text & "' , '" & TxtFormPrincipalRegistrarUsuarioEmail.Text & "' , '" & funcion & "' , " & NumericFormPrincipalRegistrarUsuarioMontosueldo.Value & " ) ; ")



                                            'devuelve el id de la ultima persona ingresada
                                            Dim IdPersona As String = consultaSQL("Select IdPersona from personas order by IdPersona desc limit 1;").Tables("personas").Rows(0)("IdPersona").ToString()
                                            'Convierte en integer el id de la ultima persona ingresada
                                            IdPersonaInt = CInt(IdPersona)



                                            'CREA EL USUARIO Y CONTRASEÑA PARA LA PERSONA INGRESADA ANTERIORMENTE
                                            updateSQL("insert into usuarios ( IdPersona , Usuario , Contraseña) values ( " & IdPersonaInt & " , '" & TxtFormPrincipalRegistrarUsuarioNombre.Text & "" & TxtFormPrincipalRegistrarUsuarioApellido.Text & "' , '" & TxtFormPrincipalRegistrarUsuarioDocumento.Text & "');")






                                            'ENVIA MAIL AL CLIENTE
                                            If EnvioMail.enviomail(TxtFormPrincipalRegistrarUsuarioEmail.Text, "RegistroCliente") = True Then

                                                Dim fechahora = fechaactual.Year & "-" & fechaactual.Month & "-" & fechaactual.Day & " " & horaactual
                                                'Guarda la notificacion en la bd
                                                updateSQL(" insert into notifica ( IdEnvia, IdRecibe, FechaNotifi, TipoNotif, Texto) values ( " & IdPersonal & " , " & IdPersonaInt & " , '" & fechahora & "' , 'Email' , 'Registro de usuario');")


                                                With Me
                                                    .ComboFormPrincipalRegistrarUsuarioTipoDocumento.Text = ""
                                                    .TxtFormPrincipalRegistrarUsuarioDocumento.Text = ""
                                                    .TxtFormPrincipalRegistrarUsuarioNombre.Text = ""
                                                    .TxtFormPrincipalRegistrarUsuarioApellido.Text = ""
                                                    .NumericFormPrincipalRegistrarUsuarioTelefono.Value = 0
                                                    .NumericFormPrincipalRegistrarUsuarioCelular.Value = 0

                                                    'CARGA INICIALMENTE LOS DEPARTAMENTOS EN EL COMBO BOX
                                                    ComboFormPrincipalRegistroUsuarioDepartamento.DataSource = (consultaSQL("SELECT distinct Departamento FROM ubicacion").Tables(0))
                                                    ComboFormPrincipalRegistroUsuarioDepartamento.DisplayMember = "Departamento"
                                                    .TxtFormPrincipalRegistrarUsuarioCalle.Text = ""
                                                    .TxtFormPrincipalRegistrarUsuarioEsquina.Text = ""
                                                    .NumericFormPrincipalRegistrarUsuarioNumeroPuerta.Value = 0
                                                    .NumericFormPrincipalRegistrarUsuarioNumeroApartamento.Value = 0
                                                    .TxtFormPrincipalRegistrarUsuarioNombreInmueble.Text = ""
                                                    .NumericFormPrincipalRegistrarUsuarioMontosueldo.Value = 0
                                                    .TxtFormPrincipalRegistrarUsuarioEmail.Text = ""


                                                End With
                                                '----------------------------------------Se logro insertar en la bd-----------------------------------------

                                                testingvalidaringresousuario = True


                                            Else


                                                Dim fechahora = fechaactual.Year & "-" & fechaactual.Month & "-" & fechaactual.Day & " " & horaactual
                                                'Guarda la notificacion en la bd
                                                updateSQL(" insert into notifica ( IdEnvia, IdRecibe, FechaNotifi, TipoNotif, Texto) values ( " & IdPersonal & " , " & IdPersonaInt & " , '" & fechahora & "' , 'Email' , 'Registro de usuario , error envio email');")




                                                With Me
                                                    .ComboFormPrincipalRegistrarUsuarioTipoDocumento.Text = ""
                                                    .TxtFormPrincipalRegistrarUsuarioDocumento.Text = ""
                                                    .TxtFormPrincipalRegistrarUsuarioNombre.Text = ""
                                                    .TxtFormPrincipalRegistrarUsuarioApellido.Text = ""
                                                    .NumericFormPrincipalRegistrarUsuarioTelefono.Value = 0
                                                    .NumericFormPrincipalRegistrarUsuarioCelular.Value = 0

                                                    'CARGA INICIALMENTE LOS DEPARTAMENTOS EN EL COMBO BOX
                                                    ComboFormPrincipalRegistroUsuarioDepartamento.DataSource = (consultaSQL("SELECT distinct Departamento FROM ubicacion").Tables(0))
                                                    ComboFormPrincipalRegistroUsuarioDepartamento.DisplayMember = "Departamento"
                                                    .TxtFormPrincipalRegistrarUsuarioCalle.Text = ""
                                                    .TxtFormPrincipalRegistrarUsuarioEsquina.Text = ""
                                                    .NumericFormPrincipalRegistrarUsuarioNumeroPuerta.Value = 0
                                                    .NumericFormPrincipalRegistrarUsuarioNumeroApartamento.Value = 0
                                                    .TxtFormPrincipalRegistrarUsuarioNombreInmueble.Text = ""
                                                    .NumericFormPrincipalRegistrarUsuarioMontosueldo.Value = 0
                                                    .TxtFormPrincipalRegistrarUsuarioEmail.Text = ""

                                                End With


                                            End If


                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If




            Catch ex As Exception
                MsgBox("No existe una sucursal en el barrio seleccionado")
            End Try
        Catch ex As Exception
            MsgBox("No se selecciono ningun barrio")
        End Try






    End Sub



    Private Sub ComboFormPrincipalRegistroUsuarioDepartamento_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboFormPrincipalRegistroUsuarioDepartamento.SelectedIndexChanged
        'CARGA LAS CIUDADES EN EL COMBOBOX SEGUN EL DEPARTAMENTO QUE SE SELECCIONO
        ComboFormPrincipalRegistroUsuarioCiudad.Text = ""
        ComboFormPrincipalRegistroUsuarioBarrio.Text = ""
        ComboFormPrincipalRegistroUsuarioCiudad.DataSource = (consultaSQL("SELECT distinct NomCiudad FROM barrios , ubicacion where barrios.IdUbicacion = ubicacion.IdUbicacion and ubicacion.Departamento = '" & ComboFormPrincipalRegistroUsuarioDepartamento.Text & "';").Tables(0))
        ComboFormPrincipalRegistroUsuarioCiudad.DisplayMember = "Nom_Ciudad"
    End Sub

    Private Sub ComboFormPrincipalRegistroUsuarioCiudad_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboFormPrincipalRegistroUsuarioCiudad.SelectedIndexChanged
        ComboFormPrincipalRegistroUsuarioBarrio.Text = ""
        ComboFormPrincipalRegistroUsuarioBarrio.DataSource = (consultaSQL("select distinct Barrios from barrios where NomCiudad = '" & ComboFormPrincipalRegistroUsuarioCiudad.Text & "';").Tables(0))
        ComboFormPrincipalRegistroUsuarioBarrio.DisplayMember = "barrios"
    End Sub



    Private Sub ComboFormPrincipalIngresoPropiedadesDepartamento_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboFormPrincipalIngresoPropiedadesDepartamento.SelectedIndexChanged
        'CARGA LAS CIUDADES EN EL COMBOBOX SEGUN EL DEPARTAMENTO QUE SE SELECCIONO
        ComboFormPrincipalIngresoPropiedadesCiudad.Text = ""
        ComboFormPrincipalIngresoPropiedadesBarrios.Text = ""
        ComboFormPrincipalIngresoPropiedadesCiudad.DataSource = (consultaSQL("SELECT distinct NomCiudad FROM barrios , ubicacion where barrios.IdUbicacion = ubicacion.IdUbicacion and ubicacion.Departamento = '" & ComboFormPrincipalIngresoPropiedadesDepartamento.Text & "';").Tables(0))
        ComboFormPrincipalIngresoPropiedadesCiudad.DisplayMember = "NomCiudad"
    End Sub
    Private Sub ComboFormPrincipalIngresoPropiedadesCiudad_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboFormPrincipalIngresoPropiedadesCiudad.SelectedIndexChanged

        ComboFormPrincipalIngresoPropiedadesBarrios.Text = ""
        ComboFormPrincipalIngresoPropiedadesBarrios.DataSource = (consultaSQL("select distinct Barrios from barrios where NomCiudad = '" & ComboFormPrincipalIngresoPropiedadesCiudad.Text & "';").Tables(0))
        ComboFormPrincipalIngresoPropiedadesBarrios.DisplayMember = "Barrios"
    End Sub

    Private Sub PanelFormPrincipalConsultaragenda_Paint(sender As Object, e As PaintEventArgs) Handles PanelFormPrincipalConsultaragenda.Paint

    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs)


        EnvioMail.enviomail("mdasilvafigueras@gmail.com", "RegistroCliente")

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)













    End Sub



    Private Sub ButtonFormPrincipalCrearreservaBuscar_Click(sender As Object, e As EventArgs) Handles ButtonFormPrincipalCrearreservaBuscar.Click


        If NumericFormPrincipalCrearreservaUltimos.Value <> 0 Then

            DataFormPrincipalCrearreserva.DataSource = consultaSQL("select IdInmueble , TipoInm, Departamento,  TipoPubli , NumDormitorios , Precio , Moneda , Jardin , Amoblada , Garage , Piscina , Barbacoa
from ubicacion, caracteristicas_inm, inmueble, barrios, publicacion 	
where caracteristicas_inm.IdInm = inmueble.IdInm and ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio and publicacion.IdPubli = inmueble.IdPubli
and TipoInm = '" & ComboFormPrincipalCrearreservaTipopropiedad.SelectedItem & "' and NumDormitorios = '" & ComboFormPrincipalCrearreservaCuartos.SelectedItem & "'
and TipoPubli = '" & ComboFormPrincipalCrearreservaTipoPubli.Text & "' and publicacion.EstadoInm = 'Activo'
limit " & NumericFormPrincipalCrearreservaUltimos.Value & "

;").Tables("consultainmueble")





        Else

            DataFormPrincipalCrearreserva.DataSource = consultaSQL("select IdInmueble , TipoInm, Departamento,  TipoPubli ,  NumDormitorios , Precio , Moneda , Jardin , Amoblada , Garage , Piscina , Barbacoa
from ubicacion, caracteristicas_inm, inmueble, barrios, publicacion 	
where caracteristicas_inm.IdInm = inmueble.IdInm and ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio and publicacion.IdPubli = inmueble.IdPubli
and TipoInm = '" & ComboFormPrincipalCrearreservaTipopropiedad.SelectedItem & "' and  NumDormitorios = '" & ComboFormPrincipalCrearreservaCuartos.SelectedItem & "'

and TipoPubli = '" & ComboFormPrincipalCrearreservaTipoPubli.Text & "'

;").Tables("consultainmueble")





        End If






    End Sub

    Private Sub BtnFormPrincipalCrearreservaAgendar_Click(sender As Object, e As EventArgs) Handles BtnFormPrincipalCrearreservaReservar.Click


        If IdInmuebleInt = 0 Then
            MsgBox("No se selecciono ningun inmueble")

        Else
            Try
                BarrioInmueble = consultaSQL("select concat( calle , (' ') , NroPuerta ) as dir from inmueble where IdInmueble = " & IdInmuebleInt & ";").Tables("inmueble").Rows(0)("dir").ToString()

                Departamentoinmueble = consultaSQL("select Departamento from ubicacion , inmueble , barrios where ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio and IdInmueble = " & IdInmuebleInt & ";").Tables("inmueble").Rows(0)("Departamento").ToString()




                BarrioInmueble = "" & BarrioInmueble & ", " & Departamentoinmueble & ", Departamento+de+" & Departamentoinmueble & ", Uruguay"

            Catch ex As Exception

            End Try


            While BarrioInmueble.Contains(" ")

                BarrioInmueble = BarrioInmueble.Replace(" ", "+")

            End While








            If TxtFormPrincipalCrearreservaDocumento.Text = "Documento" Or TxtFormPrincipalCrearreservaDocumento.Text = "" Then
                MsgBox("No se ingreso documento de identidad")
            Else


                Try
                    'devuelve el id de la persona ingresada
                    Dim IdPersona As String = consultaSQL("select IdPersona from personas where NumDoc = '" & TxtFormPrincipalCrearreservaDocumento.Text & "' and NumDoc is not null;").Tables("personas").Rows(0)("IdPersona").ToString()
                    'Convierte en integer el id de la ultima persona ingresada
                    IdPersonaInt = CInt(IdPersona)
                Catch ex As Exception
                    MsgBox("Documento de identidad no valido")
                End Try


                'SE ENCONTRO UNA PERSONA... Sigue
                If IdPersonaInt <> 0 Then


                    If Fechareserva = "" Then
                        MsgBox("No se ingreso fecha")
                    Else


                        'EL AGENTE TRABAJA ESTE DIA?
                        count = consultaSQL("select count(a.IdAgenda) as IdAgenda from agenda a where a.IdAgente = " & IdPersonal & "
                        and '" & Fechareserva & "' = a.FechaAgente
                        ;").Tables("consultainmueble").Rows(0)("IdAgenda").ToString()

                        If count = 0 Then
                            MsgBox("El agente no trabaja ese dia")
                        Else
                            count = 0
                            agenda = ""

                            inicioreserva = NumericFormPrincipalReservaHorainicio.Value
                            finalhorareserva = NumericFormPrincipalReservaHorainicio.Value + NumericFormPrincipalReservaHoraFinal.Value




                            'consultar si el agente trabaja ese dia 
                            agenda = consultaSQL("
                            select count(a.IdAgenda) as countagenda from agenda a
                            where  a.IdAgente = " & IdPersonal & "
                            and '" & Fechareserva & "' = a.FechaAgente
                            and '" & inicioreserva & ":00:00'  between DATE_ADD(a.HoraEntrada, INTERVAL - 1 HOUR) and  DATE_ADD(a.HoraSalida , INTERVAL -1 HOUR)
                            and '" & finalhorareserva & ":00:00'  between DATE_ADD(a.HoraEntrada, INTERVAL - 1 HOUR) and  DATE_ADD(a.HoraSalida , INTERVAL -1 HOUR) 
                            ;").Tables("consultainmueble").Rows(0)("countagenda").ToString()








                            If agenda = 0 Then
                                MsgBox("Fuera de rango horario del agente")

                            Else



                                agenda = ""




                                Try
                                    'consultar si hay reservas hechas en la hora que el cliente desea realizar una visita
                                    agenda = consultaSQL("
                                     select count(a.IdAgenda) as countagenda from agenda a, personas p, reserva r, Sucursales s
                                     where a.IdAgente = p.IdPersona and a.IdAgenda = r.IdAgenda and s.IdSucursal = p.IdSucursal 
                                     and   a.IdAgente = " & IdPersonal & "
                                     and '" & Fechareserva & "' = a.FechaAgente
                                     and '" & inicioreserva & ":00'  between a.HoraEntrada and  DATE_ADD(a.HoraSalida , INTERVAL -1 HOUR)
                                     and '" & finalhorareserva & ":00'  between  DATE_ADD(a.HoraEntrada, INTERVAL +1 HOUR) and  DATE_ADD(a.HoraSalida , INTERVAL -1 HOUR) 
                            
                                    /*Horario de otras reservas - busca dias siguientes disponibles*/
                                    and date(r.FehahraVisita) = '" & Fechareserva & "' 
                            
                                    /*Si el horario de reserva esta dentro de una reserva ya hecha*/
                                    and ( '" & inicioreserva & ":00' between  time(r.FehahraVisita)  and time(r.FinalVisita)
                                    OR '" & finalhorareserva & ":00' between  time(r.FehahraVisita) and time(r.FinalVisita)
                            
                                    /*Si la reserva ya ingresada esta dentro de la hora de reserva que quiere el cliente*/
                                    or (time(r.FehahraVisita) between '" & inicioreserva & ":00' and '" & finalhorareserva & ":00'
                                    and time(r.FinalVisita) between '" & inicioreserva & ":00' and '" & finalhorareserva & ":00' ))
                                    and EstadoReserva = 'Activo'
                                    ;").Tables("consultainmueble").Rows(0)("countagenda").ToString()

                                Catch ex As Exception

                                End Try





                                If agenda <> 0 Then
                                    MsgBox("Hora no disponible")
                                Else
                                    agenda = ""
                                    Dim idagendaInt As Integer = 0

                                    inicioreserva = NumericFormPrincipalReservaHorainicio.Value
                                    finalhorareserva = NumericFormPrincipalReservaHorainicio.Value + NumericFormPrincipalReservaHoraFinal.Value

                                    Try

                                        Dim idagenda As String = consultaSQL("
                                    select a.IdAgenda as IdAgenda from agenda a
                                    where a.IdAgente = " & IdPersonal & "
                                    and '" & Fechareserva & "' = a.FechaAgente
                                    and '" & inicioreserva & ":00' between a.HoraEntrada and  DATE_ADD(a.HoraSalida , INTERVAL -1 HOUR)
                                    and '" & finalhorareserva & ":00'  between  DATE_ADD(a.HoraEntrada, INTERVAL +1 HOUR) and  DATE_ADD(a.HoraSalida , INTERVAL -1 HOUR) 
                                    limit 1
                                    ;").Tables("consultainmueble").Rows(0)("IdAgenda").ToString()

                                        idagendaInt = CInt(idagenda)

                                    Catch ex As Exception

                                    End Try




                                    If idagendaInt <> 0 Then

                                        Dim result As DialogResult = MessageBox.Show("          Confirma el registro ?", "          Reserva visita", MessageBoxButtons.YesNo)


                                        If result = DialogResult.Yes Then




                                            Dim fechahoraactual = fechaactual.Year & "-" & fechaactual.Month & "-" & fechaactual.Day & " " & horaactual
                                            Dim fechahorainicialreserva = Fechareserva & " " & inicioreserva & ":00:00"
                                            Dim fechahorafinalreserva = Fechareserva & " " & finalhorareserva & ":00:00"


                                            updateSQL("insert into reserva ( IdInmueble , IdAgenda , IdInteresado , FechaHraReserva , FehaHraVisita , FinalVisita , EstadoReserva ) values
                                             ( " & IdInmuebleInt & " , " & idagendaInt & " , " & IdPersonaInt & " , '" & fechahoraactual & "' , '" & fechahorainicialreserva & "' , '" & fechahorafinalreserva & "' , 'Activo');")



                                            Dim idreserva As String = consultaSQL("select IdReserva from reserva where IdInmueble = " & IdInmuebleInt & " and IdAgenda = " & idagendaInt & " and 
                                            IdInteresado = " & IdPersonaInt & " and FehaHraVisita = '" & fechahorainicialreserva & "' and FinalVisita = '" & fechahorafinalreserva & "';").Tables("consultainmueble").Rows(0)("IdReserva").ToString()


                                            updateSQL("insert into notifica ( IdEnvia , IdRecibe , FechaNotifi , TipoNotif , Texto , IdReserva ) values 
                                              ( " & IdPersonal & " , " & IdPersonaInt & " , '" & fechahoraactual & "' , 'QR' , 'Fecha hora y ubicacion del inmueble' , " & idreserva & " ) ;")



                                            'GENERAR CODIGO QR
                                            Dim mensaje As String = "https://calendar.google.com/calendar/render?action=TEMPLATE&text=Visita+inmueble&dates=" & añoreserva & mesreserva & diareserva & "T" & inicioreserva & "0000/" & añoreserva & mesreserva & diareserva & "T" & finalhorareserva & "0000&details=0s+Agenda+visita+a+inmueble&location=" & BarrioInmueble & "&trp=false#eventpage_6"
                                            Dim gen As New QRCodeGenerator
                                            Dim data = gen.CreateQrCode(mensaje, QRCodeGenerator.ECCLevel.Q)
                                            Dim code As New QRCode(data)
                                            PicQRAgenda.Image = code.GetGraphic(8)

                                            PanelQRReserva.Visible = True
                                            PicQRAgenda.Visible = True
                                            PicCerrarqr.Visible = True


                                            NumericFormPrincipalReservaHoraFinal.Value = 1
                                            NumericFormPrincipalReservaHorainicio.Value = 8
                                            TxtFormPrincipalCrearreservaDocumento.Text = "Documento"


                                        Else
                                            idagendaInt = 0

                                        End If

                                    Else

                                        MsgBox("Agenda no encontrada")

                                    End If








                                End If











                            End If




                        End If


                    End If



                End If






            End If












        End If







    End Sub





















    'BOTON BUSCAR AGENDAS
    Private Sub BtnFormPrincipalConsultaragendaBuscar_Click(sender As Object, e As EventArgs) Handles BtnFormPrincipalConsultaragendaBuscar.Click

        If Fechaagendaini = "" Or Fechaagendafin = "" Then
            MsgBox("No se eligio rango horario")
        Else



            If FormLogin.tipoacceso = "Agente" Then
            DataFormPrincipalConsultaragendaDatos.DataSource = consultaSQL("select EstadoReserva , FehaHraVisita as Fecha , concat(concat(Nombre , (' '))  , Apellido ) as Nombre_Apellido , Tel as Telefono , Cel as Celular , Email , Barrios as Barrio , inmueble.Calle , inmueble.NroPuerta as Numero_Puerta, inmueble.NroApto as Numero_Apartamento , reserva.IdReserva from agenda ,  reserva , personas , inmueble , barrios  where  agenda.IdAgenda = reserva.IdAgenda and reserva.IdInteresado = personas.IdPersona and reserva.IdInmueble = inmueble.IdInmueble and barrios.IdBarrio = inmueble.IdBarrio
                and  agenda.IdAgente = " & IdPersonal & "
                AND EstadoReserva = 'Activo' and FehaHraVisita >= '" & Format(DateFormPrincipalConsultaragendaInicio.Value, "yyy-MM-dd") & "' and FehaHraVisita <= '" & Format(DateFormPrincipalConsultaragendaFin.Value, "yyy-MM-dd") & "'
                
                union all select EstadoReserva ,  FehaHraVisita as Fecha , concat(concat(Nombre , (' '))  , Apellido ) as Nombre_Apellido , Tel as Telefono , Cel as Celular , Email , Barrios as Barrio , inmueble.Calle , inmueble.NroPuerta as Numero_Puerta, inmueble.NroApto as Numero_Apartamento , reserva.IdReserva from agenda ,  reserva , personas , inmueble , barrios  where  agenda.IdAgenda = reserva.IdAgenda and reserva.IdInteresado = personas.IdPersona and reserva.IdInmueble = inmueble.IdInmueble and barrios.IdBarrio = inmueble.IdBarrio
                and EstadoReserva = 'Pendiente' and reserva.IdAgenda = 1
                ;").Tables("consultainmueble")



            DataFormPrincipalConsultaragendaDatos.Columns(10).Visible = False

        Else

            Dim IdPersonal As String = consultaSQL("SELECT IdPersona FROM enigma.personas where  concat(concat(Nombre , (' '))  , Apellido ) = '" & BoxFormPrincipalConsultaAgendaAgente.Text & "';").Tables("personas").Rows(0)("IdPersona").ToString()



            'MUESTRA LAS AGENDAS PARA EL DIA DE HOY
            DataFormPrincipalConsultaragendaDatos.DataSource = consultaSQL("select EstadoReserva , FehaHraVisita as Fecha , concat(concat(Nombre , (' '))  , Apellido ) as Nombre_Apellido , Tel as Telefono , Cel as Celular , Email , Barrios as Barrio , inmueble.Calle , inmueble.NroPuerta as Numero_Puerta, inmueble.NroApto as Numero_Apartamento , reserva.IdReserva from agenda ,  reserva , personas , inmueble , barrios  where  agenda.IdAgenda = reserva.IdAgenda and reserva.IdInteresado = personas.IdPersona and reserva.IdInmueble = inmueble.IdInmueble and barrios.IdBarrio = inmueble.IdBarrio
                and  agenda.IdAgente = " & IdPersonal & "
                AND EstadoReserva = 'Activo' and FehaHraVisita >= '" & Format(DateFormPrincipalConsultaragendaInicio.Value, "yyy-MM-dd") & "' and FehaHraVisita <= '" & Format(DateFormPrincipalConsultaragendaFin.Value, "yyy-MM-dd") & "'
                union all select EstadoReserva ,  FehaHraVisita as Fecha , concat(concat(Nombre , (' '))  , Apellido ) as Nombre_Apellido , Tel as Telefono , Cel as Celular , Email , Barrios as Barrio , inmueble.Calle , inmueble.NroPuerta as Numero_Puerta, inmueble.NroApto as Numero_Apartamento , reserva.IdReserva from agenda ,  reserva , personas , inmueble , barrios  where  agenda.IdAgenda = reserva.IdAgenda and reserva.IdInteresado = personas.IdPersona and reserva.IdInmueble = inmueble.IdInmueble and barrios.IdBarrio = inmueble.IdBarrio
                and EstadoReserva = 'Pendiente' and reserva.IdAgenda = 1
                ;").Tables("consultainmueble")

                DataFormPrincipalConsultaragendaDatos.Columns(10).Visible = False

        End If



        End If




    End Sub

    Private Sub DataFormPrincipalCrearreserva_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataFormPrincipalCrearreserva.CellClick


        Try
            IdInmuebleInt = CInt(DataFormPrincipalCrearreserva.CurrentRow.Cells("IdInmueble").Value.ToString)
        Catch ex As Exception

        End Try







    End Sub

    Private Sub DateFormPrincipalCrearreservaFechaReserva_ValueChanged(sender As Object, e As EventArgs) Handles DateFormPrincipalCrearreservaFechaReserva.ValueChanged

        inicial = Me.DateFormPrincipalCrearreservaFechaReserva.Value
        Fechareserva = Format(inicial, "yyy-MM-dd")
        añoreserva = Format(inicial, "yyy")
        mesreserva = Format(inicial, "MM")
        diareserva = Format(inicial, "dd")

    End Sub

    Private Sub NumericFormPrincipalReservaHorainicio_ValueChanged(sender As Object, e As EventArgs) Handles NumericFormPrincipalReservaHorainicio.ValueChanged

    End Sub

    Private Sub DateFormPrincipalConsultaragendaInicio_ValueChanged(sender As Object, e As EventArgs) Handles DateFormPrincipalConsultaragendaInicio.ValueChanged

        inicialagenda = Me.DateFormPrincipalConsultaragendaInicio.Value

        Fechaagendaini = Format(inicialagenda, "yyy-MM-dd")









    End Sub

    Private Sub DateFormPrincipalConsultaragendaFin_ValueChanged(sender As Object, e As EventArgs) Handles DateFormPrincipalConsultaragendaFin.ValueChanged

        finalagenda = Me.DateFormPrincipalConsultaragendaFin.Value
        Fechaagendafin = Format(finalagenda, "yyy-MM-dd")




    End Sub

    Private Sub BtnFormPrincipalConsultaragendaTotal_Click(sender As Object, e As EventArgs) Handles BtnFormPrincipalConsultaragendaTotal.Click


        If FormLogin.tipoacceso = "Agente" Then
            DataFormPrincipalConsultaragendaDatos.DataSource = consultaSQL("select EstadoReserva , FehaHraVisita as Fecha , concat(concat(Nombre , (' '))  , Apellido ) as Nombre_Apellido , Tel as Telefono , Cel as Celular , Email , Barrios as Barrio , inmueble.Calle , inmueble.NroPuerta as Numero_Puerta, inmueble.NroApto as Numero_Apartamento , reserva.IdReserva from agenda ,  reserva , personas , inmueble , barrios  where  agenda.IdAgenda = reserva.IdAgenda and reserva.IdInteresado = personas.IdPersona and reserva.IdInmueble = inmueble.IdInmueble and barrios.IdBarrio = inmueble.IdBarrio
                and  agenda.IdAgente = " & IdPersonal & "
                AND EstadoReserva = 'Activo'
                union all select EstadoReserva ,  FehaHraVisita as Fecha , concat(concat(Nombre , (' '))  , Apellido ) as Nombre_Apellido , Tel as Telefono , Cel as Celular , Email , Barrios as Barrio , inmueble.Calle , inmueble.NroPuerta as Numero_Puerta, inmueble.NroApto as Numero_Apartamento , reserva.IdReserva from agenda ,  reserva , personas , inmueble , barrios  where  agenda.IdAgenda = reserva.IdAgenda and reserva.IdInteresado = personas.IdPersona and reserva.IdInmueble = inmueble.IdInmueble and barrios.IdBarrio = inmueble.IdBarrio
                and EstadoReserva = 'Pendiente' and reserva.IdAgenda = 1
                ;").Tables("consultainmueble")

            DataFormPrincipalConsultaragendaDatos.Columns(10).Visible = False

        Else

            Dim IdPersonal As String = consultaSQL("SELECT IdPersona FROM enigma.personas where  concat(concat(Nombre , (' '))  , Apellido ) = '" & BoxFormPrincipalConsultaAgendaAgente.Text & "';").Tables("personas").Rows(0)("IdPersona").ToString()



            'MUESTRA LAS AGENDAS PARA EL DIA DE HOY
            DataFormPrincipalConsultaragendaDatos.DataSource = consultaSQL("select EstadoReserva , FehaHraVisita as Fecha , concat(concat(Nombre , (' '))  , Apellido ) as Nombre_Apellido , Tel as Telefono , Cel as Celular , Email , Barrios as Barrio , inmueble.Calle , inmueble.NroPuerta as Numero_Puerta, inmueble.NroApto as Numero_Apartamento , reserva.IdReserva from agenda ,  reserva , personas , inmueble , barrios  where  agenda.IdAgenda = reserva.IdAgenda and reserva.IdInteresado = personas.IdPersona and reserva.IdInmueble = inmueble.IdInmueble and barrios.IdBarrio = inmueble.IdBarrio
                and  agenda.IdAgente = " & IdPersonal & "
                AND EstadoReserva = 'Activo'
                union all select EstadoReserva ,  FehaHraVisita as Fecha , concat(concat(Nombre , (' '))  , Apellido ) as Nombre_Apellido , Tel as Telefono , Cel as Celular , Email , Barrios as Barrio , inmueble.Calle , inmueble.NroPuerta as Numero_Puerta, inmueble.NroApto as Numero_Apartamento , reserva.IdReserva from agenda ,  reserva , personas , inmueble , barrios  where  agenda.IdAgenda = reserva.IdAgenda and reserva.IdInteresado = personas.IdPersona and reserva.IdInmueble = inmueble.IdInmueble and barrios.IdBarrio = inmueble.IdBarrio
                and EstadoReserva = 'Pendiente' and reserva.IdAgenda = 1
                ;").Tables("consultainmueble")

            DataFormPrincipalConsultaragendaDatos.Columns(10).Visible = False

        End If




    End Sub

    Private Sub BtnFormPrincipalConsultaragendaModidificar_Click(sender As Object, e As EventArgs) Handles BtnFormPrincipalConsultaragendaModidificar.Click

        If IdReserva = 0 Then

            MsgBox("No se selecciono agenda")

        Else

            'IdReserva

            PanelFormPrincipalConsultaragendaEditarAgenda.Visible = True


        End If



    End Sub

    Private Sub PictureBox1_Click_1(sender As Object, e As EventArgs) Handles PictureFormPrincipalConsultaragendamodificarCerrar.Click

        PanelFormPrincipalConsultaragendaEditarAgenda.Visible = False


    End Sub

    Private Sub DataFormPrincipalConsultaragendaDatos_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataFormPrincipalConsultaragendaDatos.CellClick
        IdReserva = CInt(DataFormPrincipalConsultaragendaDatos.CurrentRow.Cells("IdReserva").Value.ToString)


    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click


        If IdReserva = 0 Then

            MsgBox("No se selecciono agenda")

        Else



            Dim result As DialogResult = MessageBox.Show("Esta seguro de cancelar la agenda?",
                              "Cancelacion de agenda",
                              MessageBoxButtons.YesNo)

            If result = DialogResult.Yes Then

                Dim fechahoraactual = fechaactual.Year & "-" & fechaactual.Month & "-" & fechaactual.Day & " " & horaactual

                updateSQL(" UPDATE reserva SET FechaCancelacion = '" & fechahoraactual & "' , EstadoReserva='Cancelada' WHERE IdReserva = " & IdReserva & " ; ")



                BtnFormPrincipalConsultaragendaBuscar.PerformClick()








            End If



        End If



    End Sub

    Private Sub DataFormPrincipalHistorialreservas_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataFormPrincipalHistorialreservas.CellClick

        IdInmuebleInt = CInt(DataFormPrincipalHistorialreservas.CurrentRow.Cells("IdInmueble").Value.ToString)




    End Sub

    Private Sub DateFormPrincipalHistorialreservasInicio_ValueChanged(sender As Object, e As EventArgs) Handles DateFormPrincipalHistorialreservasInicio.ValueChanged

        inicialhistorico = Me.DateFormPrincipalHistorialreservasInicio.Value

        Fechahistoricoini = Format(inicialhistorico, "yyy-MM-dd")


    End Sub

    Private Sub DateFormPrincipalHistorialreservasFin_ValueChanged(sender As Object, e As EventArgs) Handles DateFormPrincipalHistorialreservasFin.ValueChanged

        finalhistorico = Me.DateFormPrincipalHistorialreservasFin.Value
        Fechahistoricofin = Format(finalhistorico, "yyy-MM-dd")


    End Sub

    Private Sub ButtonFormPrincipalHistorialreservasBuscarreserva_Click(sender As Object, e As EventArgs) Handles ButtonFormPrincipalHistorialreservasBuscarreserva.Click


        If NumericFormPrincipalHistorialreservasUltimas.Value <> 0 And Fechahistoricoini <> "" And Fechahistoricofin <> "" Then



            DataFormPrincipalHistorialreservas.DataSource = consultaSQL("SELECT inmueble.IdInmueble , sucursales.NomSucursal as Sucursal , concat(agent.Apellido , (' ')  , agent.Nombre ) as 'Agente' ,  concat(cli.Apellido , (' ') , cli.Nombre )  as 'Cliente' , ubicacion.Departamento , barrios.NomCiudad as 'Nombre ciudad' ,  notifica.FechaNotifi as 'Fecha notificacion' , notifica.TipoNotif as 'Tipo notificacion' , notifica.FechaNotifi as 'Fecha notificacion' , notifica.Texto as 'Texto notificacion'

FROM personas agent , personas cli , notifica , sucursales , reserva , inmueble , barrios , ubicacion

where notifica.IdEnvia = agent.IdPersona and notifica.IdRecibe = cli.IdPersona and agent.IdSucursal = sucursales.IdSucursal
and reserva.IdReserva = notifica.IdReserva and inmueble.IdInmueble = reserva.IdInmueble and inmueble.IdBarrio = barrios.IdBarrio and barrios.IdUbicacion = ubicacion.IdUbicacion
and FehaHraVisita BETWEEN '" & Fechahistoricoini & " 00:00:00' and   DATE_ADD(date('" & Fechahistoricofin & " 00:00:00'), INTERVAL +1 day) 
limit " & NumericFormPrincipalHistorialreservasUltimas.Value & "
;").Tables("consultainmueble")


            'DataFormPrincipalHistorialreservas.Columns(0).Visible = False


        End If


        If ComboFormPrincipalHistorialreservasDepartamento.Text <> "" And NumericFormPrincipalHistorialreservasUltimas.Value = 0 And Fechahistoricofin = "" And Fechahistoricoini = "" Then

            DataFormPrincipalHistorialreservas.DataSource = consultaSQL("SELECT inmueble.IdInmueble , sucursales.NomSucursal as Sucursal , concat(agent.Apellido , (' ')  , agent.Nombre ) as 'Agente' ,  concat(cli.Apellido , (' ') , cli.Nombre )  as 'Cliente' , ubicacion.Departamento , barrios.NomCiudad as 'Nombre ciudad' ,  notifica.FechaNotifi as 'Fecha notificacion' , notifica.TipoNotif as 'Tipo notificacion' , notifica.FechaNotifi as 'Fecha notificacion' , notifica.Texto as 'Texto notificacion'

FROM personas agent , personas cli , notifica , sucursales , reserva , inmueble , barrios , ubicacion

where notifica.IdEnvia = agent.IdPersona and notifica.IdRecibe = cli.IdPersona and agent.IdSucursal = sucursales.IdSucursal
and reserva.IdReserva = notifica.IdReserva and inmueble.IdInmueble = reserva.IdInmueble and inmueble.IdBarrio = barrios.IdBarrio and barrios.IdUbicacion = ubicacion.IdUbicacion
and ubicacion.Departamento = '" & ComboFormPrincipalHistorialreservasDepartamento.Text & "'
;").Tables("consultainmueble")


        End If




        If Fechahistoricofin = "" And Fechahistoricoini = "" And NumericFormPrincipalHistorialreservasUltimas.Value <> 0 And ComboFormPrincipalHistorialreservasDepartamento.Text <> "" Then


            DataFormPrincipalHistorialreservas.DataSource = consultaSQL("SELECT inmueble.IdInmueble , sucursales.NomSucursal as Sucursal , concat(agent.Apellido , (' ')  , agent.Nombre ) as 'Agente' ,  concat(cli.Apellido , (' ') , cli.Nombre )  as 'Cliente' , ubicacion.Departamento , barrios.NomCiudad as 'Nombre ciudad' ,  notifica.FechaNotifi as 'Fecha notificacion' , notifica.TipoNotif as 'Tipo notificacion' , notifica.FechaNotifi as 'Fecha notificacion' , notifica.Texto as 'Texto notificacion'

FROM personas agent , personas cli , notifica , sucursales , reserva , inmueble , barrios , ubicacion

where notifica.IdEnvia = agent.IdPersona and notifica.IdRecibe = cli.IdPersona and agent.IdSucursal = sucursales.IdSucursal
and reserva.IdReserva = notifica.IdReserva and inmueble.IdInmueble = reserva.IdInmueble and inmueble.IdBarrio = barrios.IdBarrio and barrios.IdUbicacion = ubicacion.IdUbicacion
and ubicacion.Departamento = '" & ComboFormPrincipalHistorialreservasDepartamento.Text & "' limit " & NumericFormPrincipalHistorialreservasUltimas.Value & "
;").Tables("consultainmueble")




        End If




        If Fechahistoricofin <> "" And Fechahistoricoini <> "" And NumericFormPrincipalHistorialreservasUltimas.Value = 0 And ComboFormPrincipalHistorialreservasDepartamento.Text = "" Then


            DataFormPrincipalHistorialreservas.DataSource = consultaSQL("SELECT inmueble.IdInmueble , sucursales.NomSucursal as Sucursal , concat(agent.Apellido , (' ')  , agent.Nombre ) as 'Agente' ,  concat(cli.Apellido , (' ') , cli.Nombre )  as 'Cliente' , ubicacion.Departamento , barrios.NomCiudad as 'Nombre ciudad' ,  notifica.FechaNotifi as 'Fecha notificacion' , notifica.TipoNotif as 'Tipo notificacion' , notifica.FechaNotifi as 'Fecha notificacion' , notifica.Texto as 'Texto notificacion'

FROM personas agent , personas cli , notifica , sucursales , reserva , inmueble , barrios , ubicacion

where notifica.IdEnvia = agent.IdPersona and notifica.IdRecibe = cli.IdPersona and agent.IdSucursal = sucursales.IdSucursal
and reserva.IdReserva = notifica.IdReserva and inmueble.IdInmueble = reserva.IdInmueble and inmueble.IdBarrio = barrios.IdBarrio and barrios.IdUbicacion = ubicacion.IdUbicacion
and FehaHraVisita BETWEEN '" & Fechahistoricoini & " 00:00:00' and   DATE_ADD(date('" & Fechahistoricofin & " 00:00:00'), INTERVAL +1 day) 

;").Tables("consultainmueble")


        End If



        If Fechahistoricofin = "" And Fechahistoricoini = "" And ComboFormPrincipalHistorialreservasDepartamento.Text = "" And NumericFormPrincipalHistorialreservasUltimas.Value <> 0 Then



            DataFormPrincipalHistorialreservas.DataSource = consultaSQL("SELECT inmueble.IdInmueble , sucursales.NomSucursal as Sucursal , concat(agent.Apellido , (' ')  , agent.Nombre ) as 'Agente' ,  concat(cli.Apellido , (' ') , cli.Nombre )  as 'Cliente' , ubicacion.Departamento , barrios.NomCiudad as 'Nombre ciudad' ,  notifica.FechaNotifi as 'Fecha notificacion' , notifica.TipoNotif as 'Tipo notificacion' , notifica.FechaNotifi as 'Fecha notificacion' , notifica.Texto as 'Texto notificacion'

FROM personas agent , personas cli , notifica , sucursales , reserva , inmueble , barrios , ubicacion

where notifica.IdEnvia = agent.IdPersona and notifica.IdRecibe = cli.IdPersona and agent.IdSucursal = sucursales.IdSucursal
and reserva.IdReserva = notifica.IdReserva and inmueble.IdInmueble = reserva.IdInmueble and inmueble.IdBarrio = barrios.IdBarrio and barrios.IdUbicacion = ubicacion.IdUbicacion

limit " & NumericFormPrincipalHistorialreservasUltimas.Value & "
;").Tables("consultainmueble")


        End If



        If Fechahistoricofin <> "" And Fechahistoricoini <> "" And ComboFormPrincipalHistorialreservasDepartamento.Text <> "" And NumericFormPrincipalHistorialreservasUltimas.Value <> 0 Then


            DataFormPrincipalHistorialreservas.DataSource = consultaSQL("SELECT inmueble.IdInmueble , sucursales.NomSucursal as Sucursal , concat(agent.Apellido , (' ')  , agent.Nombre ) as 'Agente' ,  concat(cli.Apellido , (' ') , cli.Nombre )  as 'Cliente' , ubicacion.Departamento , barrios.NomCiudad as 'Nombre ciudad' ,  notifica.FechaNotifi as 'Fecha notificacion' , notifica.TipoNotif as 'Tipo notificacion' , notifica.FechaNotifi as 'Fecha notificacion' , notifica.Texto as 'Texto notificacion'

FROM personas agent , personas cli , notifica , sucursales , reserva , inmueble , barrios , ubicacion

where notifica.IdEnvia = agent.IdPersona and notifica.IdRecibe = cli.IdPersona and agent.IdSucursal = sucursales.IdSucursal
and reserva.IdReserva = notifica.IdReserva and inmueble.IdInmueble = reserva.IdInmueble and inmueble.IdBarrio = barrios.IdBarrio and barrios.IdUbicacion = ubicacion.IdUbicacion
and FehaHraVisita BETWEEN '" & Fechahistoricoini & " 00:00:00' and   DATE_ADD(date('" & Fechahistoricofin & " 00:00:00'), INTERVAL +1 day) 
and ubicacion.Departamento = '" & ComboFormPrincipalHistorialreservasDepartamento.Text & "'
limit " & NumericFormPrincipalHistorialreservasUltimas.Value & "
;").Tables("consultainmueble")


        End If






    End Sub

    Private Sub BtnFormPrincipalHistorialreservasTotal_Click(sender As Object, e As EventArgs) Handles BtnFormPrincipalHistorialreservasTotal.Click



        DataFormPrincipalHistorialreservas.DataSource = consultaSQL("SELECT inmueble.IdInmueble , sucursales.NomSucursal as Sucursal , concat(agent.Apellido , (' ')  , agent.Nombre ) as 'Agente' ,  concat(cli.Apellido , (' ') , cli.Nombre )  as 'Cliente' , ubicacion.Departamento , barrios.NomCiudad as 'Nombre ciudad' ,  notifica.FechaNotifi as 'Fecha notificacion' , notifica.TipoNotif as 'Tipo notificacion' , notifica.FechaNotifi as 'Fecha notificacion' , notifica.Texto as 'Texto notificacion'

FROM personas agent , personas cli , notifica , sucursales , reserva , inmueble , barrios , ubicacion

where notifica.IdEnvia = agent.IdPersona and notifica.IdRecibe = cli.IdPersona and agent.IdSucursal = sucursales.IdSucursal
and reserva.IdReserva = notifica.IdReserva and inmueble.IdInmueble = reserva.IdInmueble and inmueble.IdBarrio = barrios.IdBarrio and barrios.IdUbicacion = ubicacion.IdUbicacion
order by agent.Apellido , FechaNotifi
;").Tables("consultainmueble")

        DataFormPrincipalHistorialreservas.Columns(0).Visible = False

        Fechahistoricofin = ""
        Fechahistoricoini = ""
        NumericFormPrincipalHistorialreservasUltimas.Value = 0
        ComboFormPrincipalHistorialreservasDepartamento.Text = ""

    End Sub

    Private Sub TxtFormPrincipalCrearreservaDocumento_MouseEnter(sender As Object, e As EventArgs) Handles TxtFormPrincipalCrearreservaDocumento.MouseEnter

        If TxtFormPrincipalCrearreservaDocumento.Text = "Documento" Then
            TxtFormPrincipalCrearreservaDocumento.Text = ""
        End If


    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        DataFormPrincipalCrearreserva.DataSource = consultaSQL("select IdInmueble , TipoInm, Departamento,  TipoPubli , NumDormitorios , Precio , Moneda , Jardin , Amoblada , Garage , Piscina , Barbacoa
                                                                    from ubicacion, caracteristicas_inm, inmueble, barrios, publicacion 	
                                                                    where caracteristicas_inm.IdInm = inmueble.IdInm and ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio and publicacion.IdPubli = inmueble.IdPubli
                                                                    and ubicacion.Departamento = '" & departamentoagente & "'  and publicacion.EstadoInm = 'Activo'                                                            
                                                                    limit 50;").Tables("consultainmueble")

        DataFormPrincipalCrearreserva.Columns(0).Visible = False
    End Sub

    Private Sub BtnFormPrincipalBuscarInmueblesCompraBuscar_Click(sender As Object, e As EventArgs) Handles BtnFormPrincipalBuscarInmueblesCompraBuscar.Click


        Me.DataFormPrincipalBuscarInmuebles.DataSource = consultaSQL("select TipoInm, Departamento, NumDormitorios , TipoPubli , Precio , Moneda , Jardin , Amoblada , Garage , Piscina , Barbacoa
from ubicacion, caracteristicas_inm, inmueble, barrios, publicacion 	
where caracteristicas_inm.IdInm = inmueble.IdInm and ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio and publicacion.IdPubli = inmueble.IdPubli
and TipoInm = '" & BoxFormPrincipalBuscarInmueblesCompraTipoPropiedad.SelectedItem & "' and Departamento = '" & BoxFormPrincipalBuscarInmueblesCompraUbicacion.SelectedItem & "' and NumDormitorios = '" & BoxFormPrincipalBuscarInmueblesCompraHabitaciones.SelectedItem & "'

and TipoPubli = 'Venta'
;").Tables("consultainmueble")





    End Sub

    Private Sub CheckFormPrincipalBuscarInmueblesPatio_CheckedChanged(sender As Object, e As EventArgs) Handles CheckFormPrincipalBuscarInmueblesPatio.CheckedChanged



        If buscarcaract = True Then


            Dim tipopubli As Integer = TabFormPrincipalBuscarInmueblesTipoInteres.SelectedIndex



            If CheckFormPrincipalBuscarInmueblesPatio.Checked = True Then
                jardin = "Si"

            Else
                jardin = "No"
            End If




            If tipopubli = 0 Then
                '  MsgBox("alquiler " & " Jardin: " & jardin)

                Me.DataFormPrincipalBuscarInmuebles.DataSource = consultaSQL("select TipoInm, Departamento, NumDormitorios , TipoPubli , Precio , Moneda , Jardin , Amoblada , Garage , Piscina , Barbacoa
from ubicacion, caracteristicas_inm, inmueble, barrios, publicacion 	
where caracteristicas_inm.IdInm = inmueble.IdInm and ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio and publicacion.IdPubli = inmueble.IdPubli
and TipoInm = '" & BoxFormPrincipalBuscarInmueblesAlquilerTipoPropiedad.Text & "' and Departamento = '" & BoxFormPrincipalBuscarInmueblesAlquilerUbicacion.Text & "' and NumDormitorios = '" & BoxFormPrincipalBuscarInmueblesAlquilerHabitaciones.Text & "'

and TipoPubli = 'Alquiler'

and   Jardin = '" & jardin & "' and ( Amoblada = '" & amoblada & "' or Garage = '" & garage & "' or Piscina = '" & piscina & "' or Barbacoa = '" & barbacoa & "' )

 ;").Tables("consultainmueble")



            End If

            If tipopubli = 1 Then
                'MsgBox("compara " & " Jardin: " & jardin)

                Me.DataFormPrincipalBuscarInmuebles.DataSource = consultaSQL("select TipoInm, Departamento, NumDormitorios , TipoPubli , Precio , Moneda , Jardin , Amoblada , Garage , Piscina , Barbacoa
from ubicacion, caracteristicas_inm, inmueble, barrios, publicacion 	
where caracteristicas_inm.IdInm = inmueble.IdInm and ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio and publicacion.IdPubli = inmueble.IdPubli
and TipoInm = '" & BoxFormPrincipalBuscarInmueblesAlquilerTipoPropiedad.Text & "' and Departamento = '" & BoxFormPrincipalBuscarInmueblesAlquilerUbicacion.Text & "' and NumDormitorios = '" & BoxFormPrincipalBuscarInmueblesAlquilerHabitaciones.Text & "'

and TipoPubli = 'Venta'

and Jardin = '" & jardin & "'  and ( Amoblada = '" & amoblada & "' or Garage = '" & garage & "' or Piscina = '" & piscina & "' or Barbacoa = '" & barbacoa & "' ) 



;").Tables("consultainmueble")




            End If



        End If





    End Sub

    Private Sub CheckFormPrincipalBuscarInmueblesParrillero_CheckedChanged(sender As Object, e As EventArgs) Handles CheckFormPrincipalBuscarInmueblesAmoblada.CheckedChanged


        If buscarcaract = True Then



            Dim tipopubli As Integer = TabFormPrincipalBuscarInmueblesTipoInteres.SelectedIndex



            If CheckFormPrincipalBuscarInmueblesAmoblada.Checked = True Then
                amoblada = "Si"

            Else
                amoblada = "No"
            End If


            If tipopubli = 0 Then
                ' MsgBox("alquiler " & " Amoblada: " & amoblada)

                Me.DataFormPrincipalBuscarInmuebles.DataSource = consultaSQL("select TipoInm, Departamento, NumDormitorios , TipoPubli , Precio , Moneda , Jardin , Amoblada , Garage , Piscina , Barbacoa
from ubicacion, caracteristicas_inm, inmueble, barrios, publicacion 	
where caracteristicas_inm.IdInm = inmueble.IdInm and ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio and publicacion.IdPubli = inmueble.IdPubli
and TipoInm = '" & BoxFormPrincipalBuscarInmueblesAlquilerTipoPropiedad.Text & "' and Departamento = '" & BoxFormPrincipalBuscarInmueblesAlquilerUbicacion.Text & "' and NumDormitorios = '" & BoxFormPrincipalBuscarInmueblesAlquilerHabitaciones.Text & "'

and TipoPubli = 'Alquiler'

and   Amoblada = '" & amoblada & "' and ( Jardin = '" & jardin & "'  or Garage = '" & garage & "' or Piscina = '" & piscina & "' or Barbacoa = '" & barbacoa & "' )



;").Tables("consultainmueble")


            End If

            If tipopubli = 1 Then
                'MsgBox("compara " & " Amoblada: " & amoblada)

                Me.DataFormPrincipalBuscarInmuebles.DataSource = consultaSQL("select TipoInm, Departamento, NumDormitorios , TipoPubli , Precio , Moneda , Jardin , Amoblada , Garage , Piscina , Barbacoa
from ubicacion, caracteristicas_inm, inmueble, barrios, publicacion 	
where caracteristicas_inm.IdInm = inmueble.IdInm and ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio and publicacion.IdPubli = inmueble.IdPubli
and TipoInm = '" & BoxFormPrincipalBuscarInmueblesAlquilerTipoPropiedad.Text & "' and Departamento = '" & BoxFormPrincipalBuscarInmueblesAlquilerUbicacion.Text & "' and NumDormitorios = '" & BoxFormPrincipalBuscarInmueblesAlquilerHabitaciones.Text & "'

and TipoPubli = 'Venta'

and  Amoblada = '" & amoblada & "' and ( Jardin = '" & jardin & "'  or  and Garage = '" & garage & "' or Piscina = '" & piscina & "' or Barbacoa = '" & barbacoa & "' ) 


;").Tables("consultainmueble")

            End If


        End If





    End Sub

    Private Sub CheckFormPrincipalBuscarInmueblesGarage_CheckedChanged(sender As Object, e As EventArgs) Handles CheckFormPrincipalBuscarInmueblesGarage.CheckedChanged



        If buscarcaract = True Then



            Dim tipopubli As Integer = TabFormPrincipalBuscarInmueblesTipoInteres.SelectedIndex



            If CheckFormPrincipalBuscarInmueblesGarage.Checked = True Then
                garage = "Si"

            Else
                garage = "No"
            End If


            If tipopubli = 0 Then
                ' MsgBox("alquiler " & " Garage: " & garage)

                Me.DataFormPrincipalBuscarInmuebles.DataSource = consultaSQL("select TipoInm, Departamento, NumDormitorios , TipoPubli , Precio , Moneda , Jardin , Amoblada , Garage , Piscina , Barbacoa
from ubicacion, caracteristicas_inm, inmueble, barrios, publicacion 	
where caracteristicas_inm.IdInm = inmueble.IdInm and ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio and publicacion.IdPubli = inmueble.IdPubli
and TipoInm = '" & BoxFormPrincipalBuscarInmueblesAlquilerTipoPropiedad.Text & "' and Departamento = '" & BoxFormPrincipalBuscarInmueblesAlquilerUbicacion.Text & "' and NumDormitorios = '" & BoxFormPrincipalBuscarInmueblesAlquilerHabitaciones.Text & "'

and TipoPubli = 'Alquiler'

and Garage = '" & garage & "' and ( Jardin = '" & jardin & "'  or Amoblada = '" & amoblada & "'  or Piscina = '" & piscina & "' or Barbacoa = '" & barbacoa & "' ) 



;").Tables("consultainmueble")


            End If

            If tipopubli = 1 Then
                'MsgBox("compara " & " Garage: " & garage)

                Me.DataFormPrincipalBuscarInmuebles.DataSource = consultaSQL("select TipoInm, Departamento, NumDormitorios , TipoPubli , Precio , Moneda , Jardin , Amoblada , Garage , Piscina , Barbacoa
from ubicacion, caracteristicas_inm, inmueble, barrios, publicacion 	
where caracteristicas_inm.IdInm = inmueble.IdInm and ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio and publicacion.IdPubli = inmueble.IdPubli
and TipoInm = '" & BoxFormPrincipalBuscarInmueblesAlquilerTipoPropiedad.Text & "' and Departamento = '" & BoxFormPrincipalBuscarInmueblesAlquilerUbicacion.Text & "' and NumDormitorios = '" & BoxFormPrincipalBuscarInmueblesAlquilerHabitaciones.Text & "'

and TipoPubli = 'Venta'


and Garage = '" & garage & "' and ( Jardin = '" & jardin & "' or Amoblada = '" & amoblada & "' or  and Piscina = '" & piscina & "' or Barbacoa = '" & barbacoa & "' )


;").Tables("consultainmueble")

            End If


        End If



    End Sub

    Private Sub CheckFormPrincipalBuscarInmueblesPiscina_CheckedChanged(sender As Object, e As EventArgs) Handles CheckFormPrincipalBuscarInmueblesPiscina.CheckedChanged



        If buscarcaract = True Then



            Dim tipopubli As Integer = TabFormPrincipalBuscarInmueblesTipoInteres.SelectedIndex



            If CheckFormPrincipalBuscarInmueblesPiscina.Checked = True Then
                piscina = "Si"

            Else
                piscina = "No"
            End If


            If tipopubli = 0 Then
                ' MsgBox("alquiler " & " Piscina: " & piscina)

                Me.DataFormPrincipalBuscarInmuebles.DataSource = consultaSQL("select TipoInm, Departamento, NumDormitorios , TipoPubli , Precio , Moneda , Jardin , Amoblada , Garage , Piscina , Barbacoa
from ubicacion, caracteristicas_inm, inmueble, barrios, publicacion 	
where caracteristicas_inm.IdInm = inmueble.IdInm and ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio and publicacion.IdPubli = inmueble.IdPubli
and TipoInm = '" & BoxFormPrincipalBuscarInmueblesAlquilerTipoPropiedad.Text & "' and Departamento = '" & BoxFormPrincipalBuscarInmueblesAlquilerUbicacion.Text & "' and NumDormitorios = '" & BoxFormPrincipalBuscarInmueblesAlquilerHabitaciones.Text & "'

and TipoPubli = 'Alquiler'

and Piscina = '" & piscina & "' and ( Jardin = '" & jardin & "'  or Amoblada = '" & amoblada & "' or Garage = '" & garage & "' or Barbacoa = '" & barbacoa & "') 



;").Tables("consultainmueble")


            End If

            If tipopubli = 1 Then
                'MsgBox("compara " & " Piscina: " & piscina)


                Me.DataFormPrincipalBuscarInmuebles.DataSource = consultaSQL("select TipoInm, Departamento, NumDormitorios , TipoPubli , Precio , Moneda , Jardin , Amoblada , Garage , Piscina , Barbacoa
from ubicacion, caracteristicas_inm, inmueble, barrios, publicacion 	
where caracteristicas_inm.IdInm = inmueble.IdInm and ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio and publicacion.IdPubli = inmueble.IdPubli
and TipoInm = '" & BoxFormPrincipalBuscarInmueblesAlquilerTipoPropiedad.Text & "' and Departamento = '" & BoxFormPrincipalBuscarInmueblesAlquilerUbicacion.Text & "' and NumDormitorios = '" & BoxFormPrincipalBuscarInmueblesAlquilerHabitaciones.Text & "'


and TipoPubli = 'Venta'

and Piscina = '" & piscina & "' and ( Jardin = '" & jardin & "' or Amoblada = '" & amoblada & "' or Garage = '" & garage & "' or Barbacoa = '" & barbacoa & "' ) 



;").Tables("consultainmueble")

            End If


        End If


    End Sub

    Private Sub CheckFormPrincipalBuscarInmueblesPorche_CheckedChanged(sender As Object, e As EventArgs) Handles CheckFormPrincipalBuscarInmueblesBarbacoa.CheckedChanged


        If buscarcaract = True Then



            Dim tipopubli As Integer = TabFormPrincipalBuscarInmueblesTipoInteres.SelectedIndex



            If CheckFormPrincipalBuscarInmueblesBarbacoa.Checked = True Then
                barbacoa = "Si"

            Else
                barbacoa = "No"
            End If


            If tipopubli = 0 Then
                ' MsgBox("alquiler " & " Barbacoa: " & barbacoa)

                Me.DataFormPrincipalBuscarInmuebles.DataSource = consultaSQL("select TipoInm, Departamento, NumDormitorios , TipoPubli , Precio , Moneda , Jardin , Amoblada , Garage , Piscina , Barbacoa
from ubicacion, caracteristicas_inm, inmueble, barrios, publicacion 	
where caracteristicas_inm.IdInm = inmueble.IdInm and ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio and publicacion.IdPubli = inmueble.IdPubli
and TipoInm = '" & BoxFormPrincipalBuscarInmueblesAlquilerTipoPropiedad.Text & "' and Departamento = '" & BoxFormPrincipalBuscarInmueblesAlquilerUbicacion.Text & "' and NumDormitorios = '" & BoxFormPrincipalBuscarInmueblesAlquilerHabitaciones.Text & "'

and TipoPubli = 'Alquiler'

and  Barbacoa = '" & barbacoa & "'  and ( Jardin = '" & jardin & "' or Amoblada = '" & amoblada & "' or Garage = '" & garage & "' or Piscina = '" & piscina & "' ) 



;").Tables("consultainmueble")


            End If

            If tipopubli = 1 Then
                'MsgBox("compara " & " Barbacoa: " & barbacoa)


                Me.DataFormPrincipalBuscarInmuebles.DataSource = consultaSQL("select TipoInm, Departamento, NumDormitorios , TipoPubli , Precio , Moneda , Jardin , Amoblada , Garage , Piscina , Barbacoa
from ubicacion, caracteristicas_inm, inmueble, barrios, publicacion 	
where caracteristicas_inm.IdInm = inmueble.IdInm and ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio and publicacion.IdPubli = inmueble.IdPubli
and TipoInm = '" & BoxFormPrincipalBuscarInmueblesAlquilerTipoPropiedad.SelectedItem & "' and Departamento = '" & BoxFormPrincipalBuscarInmueblesAlquilerUbicacion.SelectedItem & "' and NumDormitorios = '" & BoxFormPrincipalBuscarInmueblesAlquilerHabitaciones.SelectedItem & "'

and TipoPubli = 'Venta'

and Barbacoa = '" & barbacoa & "' and ( Jardin = '" & jardin & "'  or Amoblada = '" & amoblada & "' or Garage = '" & garage & "' or Piscina = '" & piscina & "' )



;").Tables("consultainmueble")


            End If


        End If



    End Sub





    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateFormPrincipalConsultaragendamodificada.ValueChanged

        inicial = Me.DateFormPrincipalConsultaragendamodificada.Value
        FechareservaModificada = Format(inicial, "yyy-MM-dd")

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles BtnFormPrincipalConsultaragendamodificarGuardar.Click

        Dim IdInmueble As String = consultaSQL("SELECT IdInmueble FROM reserva
                        where IdReserva = " & IdReserva & "
                        ;").Tables("personas").Rows(0)("IdInmueble").ToString()
        IdInmuebleInt = CInt(IdInmueble)



        Dim IdPersona As String = consultaSQL("SELECT IdInteresado FROM reserva
                        where IdReserva = " & IdReserva & "
                        ;").Tables("personas").Rows(0)("IdInteresado").ToString()
        IdPersonaInt = CInt(IdPersona)


        'seleccion de fecha
        If FechareservaModificada = "" Then
            MsgBox("No se ingreso fecha")
        Else



            'EL AGENTE TRABAJA ESTE DIA?
            count = consultaSQL("select count(a.IdAgenda) as IdAgenda from agenda a where a.IdAgente = " & IdPersonal & "
                        and '" & FechareservaModificada & "' = a.FechaAgente
                        ;").Tables("consultainmueble").Rows(0)("IdAgenda").ToString()

            If count = 0 Then
                MsgBox("El agente no trabaja ese dia")
            Else
                count = 0
                agenda = "0"


                Try

                    inicioreservamodificada = NumericFormPrincipalReservaHorainicioModificada.Value
                    finreservamodificada = NumericFormPrincipalReservaHorainicioModificada.Value + NumericFormPrincipalReservaHoraFinalModificada.Value

                    'consultar si el agente trabaja ese dia 
                    agenda = consultaSQL("
                            select count(a.IdAgenda) as countagenda from agenda a
                            where  a.IdAgente = " & IdPersonal & "
                            and '" & FechareservaModificada & "' = a.FechaAgente
                            and '" & inicioreservamodificada & ":00:00'  between a.HoraEntrada and  DATE_ADD(a.HoraSalida , INTERVAL -1 HOUR)
                            and '" & finreservamodificada & ":00:00'  between  DATE_ADD(a.HoraEntrada, INTERVAL +1 HOUR) and  DATE_ADD(a.HoraSalida , INTERVAL -1 HOUR) 
                            ;").Tables("consultainmueble").Rows(0)("countagenda").ToString()
                Catch ex As Exception

                End Try





                If agenda = 0 Then

                    MsgBox("Fuera de rango horario del agente")

                Else
                    agenda = "0"

                    Try
                        'consultar si hay reservas hechas en la hora que el cliente desea realizar una visita
                        agenda = consultaSQL("
                            select count(a.IdAgenda) as countagenda from agenda a, personas p, reserva r, Sucursales s
                            where a.IdAgente = p.IdPersona and a.IdAgenda = r.IdAgenda and s.IdSucursal = p.IdSucursal 
                            and   a.IdAgente = " & IdPersonal & "
                            and '" & FechareservaModificada & "' = a.FechaAgente
                            and '" & inicioreservamodificada & ":00'  between a.HoraEntrada and  DATE_ADD(a.HoraSalida , INTERVAL -1 HOUR)
                            and '" & finreservamodificada & ":00'  between  DATE_ADD(a.HoraEntrada, INTERVAL +1 HOUR) and  DATE_ADD(a.HoraSalida , INTERVAL -1 HOUR) 
                            
                            /*Horario de otras reservas - busca dias siguientes disponibles*/
                            and date(r.FehahraVisita) = '" & FechareservaModificada & "' 
                            
                            /*Si el horario de reserva esta dentro de una reserva ya hecha*/
                            and ( '" & inicioreservamodificada & ":00' between  time(r.FehahraVisita)  and time(r.FinalVisita)
                            OR '" & finreservamodificada & ":00' between  time(r.FehahraVisita) and time(r.FinalVisita)
                            
                            /*Si la reserva ya ingresada esta dentro de la hora de reserva que quiere el cliente*/
                            or (time(r.FehahraVisita) between '" & inicioreservamodificada & ":00' and '" & finreservamodificada & ":00'
                            and time(r.FinalVisita) between '" & inicioreservamodificada & ":00' and '" & finreservamodificada & ":00' ))
                            and EstadoReserva = 'Activo'
                            ;").Tables("consultainmueble").Rows(0)("countagenda").ToString()
                    Catch ex As Exception

                    End Try





                    If agenda <> 0 Then
                        MsgBox("Hora no disponible")
                    Else



                        inicioreservamodificada = NumericFormPrincipalReservaHorainicioModificada.Value
                        finreservamodificada = NumericFormPrincipalReservaHorainicioModificada.Value + NumericFormPrincipalReservaHoraFinalModificada.Value



                        'OBTENGO EL ID DE ESA AGENDA
                        Dim idagenda As String = consultaSQL("
                        select a.IdAgenda as IdAgenda from agenda a, personas p, reserva r, Sucursales s
                        where a.IdAgente = p.IdPersona 
                        and a.IdAgente = " & IdPersonal & "
                        and '" & FechareservaModificada & "' = a.FechaAgente
                        and '" & inicioreservamodificada & ":00'  between  DATE_ADD(a.HoraEntrada, INTERVAL +1 HOUR) and  DATE_ADD(a.HoraSalida , INTERVAL -1 HOUR)
                        and '" & inicioreservamodificada & ":00'  between  DATE_ADD(a.HoraEntrada, INTERVAL +1 HOUR) and  DATE_ADD(a.HoraSalida , INTERVAL -1 HOUR) 
                        limit 1
                         ;").Tables("consultainmueble").Rows(0)("IdAgenda").ToString()

                        Dim idagendaInt As Integer = CInt(idagenda)


                        'GUARDAR EL REGISTRO
                        Dim result As DialogResult = MessageBox.Show("          Confirma el registro ?", "          Reserva visita", MessageBoxButtons.YesNo)


                        If result = DialogResult.Yes Then

                            Dim fechahoraactual = fechaactual.Year & "-" & fechaactual.Month & "-" & fechaactual.Day & " " & horaactual
                            Dim fechahorainicialreserva = FechareservaModificada & " " & inicioreservamodificada & ":00:00"
                            Dim fechahorafinalreserva = FechareservaModificada & " " & finreservamodificada & ":00:00"

                            updateSQL("insert into reserva ( IdInmueble , IdAgenda , IdInteresado , FechaHraReserva , FehaHraVisita , FinalVisita , EstadoReserva ) values
                                    ( " & IdInmuebleInt & " , " & idagendaInt & " , " & IdPersonaInt & " , '" & fechahoraactual & "' , '" & fechahorainicialreserva & "' , '" & fechahorafinalreserva & "' , 'Activo');")


                            updateSQL(" UPDATE reserva SET FechaModificacion = '" & fechahoraactual & "' , EstadoReserva='Modificada' WHERE IdReserva = " & IdReserva & " ; ")
                            MsgBox("SE ACTUALIZO LA RESERVA CON LA AGENDA CORRECTAMENTE")


                            'ENVIAR SMS AL CLIENTE



















                            NumericFormPrincipalReservaHoraFinal.Value = 1
                            NumericFormPrincipalReservaHorainicio.Value = 8
                            idagendaInt = 0
                            count = 0
                            agenda = ""
                            IdReserva = 0
                            TxtFormPrincipalCrearreservaDocumento.Text = "Documento"
                            PanelFormPrincipalConsultaragendaEditarAgenda.Visible = False

                            'Actualiza con los datos del dia
                            DataFormPrincipalConsultaragendaDatos.DataSource = consultaSQL("select FehaHraVisita as Fecha , concat(concat(Nombre , (' '))  , Apellido ) as Nombre_Apellido , 
                            Tel as Telefono , Cel as Celular , Email , Barrios as Barrio , inmueble.Calle , inmueble.NroPuerta as Numero_Puerta, inmueble.NroApto as Numero_Apartamento , reserva.IdReserva 
                            from agenda ,  reserva , personas , inmueble , barrios  where agenda.IdAgenda = reserva.IdAgenda and reserva.IdInteresado = personas.IdPersona and reserva.IdInmueble = inmueble.IdInmueble and barrios.IdBarrio = inmueble.IdBarrio
                            and reserva.FehaHraVisita like '%" & fechaactual.Year & "-%" & fechaactual.Month & "-%" & fechaactual.Day & "%' and  agenda.IdAgente = " & IdPersonal & " 
                            AND (EstadoReserva = 'Activo' OR EstadoReserva = 'Pendiente' )
                            ;").Tables("consultainmueble")




                        End If









                    End If







                End If






            End If

        End If








    End Sub



    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click


        Dim idinm As Integer = NumericNumInmueble.Value
        idinm = idinm - 100




        DataFormPrincipalCrearreserva.DataSource = consultaSQL("select IdInmueble , TipoInm, Departamento,  TipoPubli , NumDormitorios , Precio , Moneda , Jardin , Amoblada , Garage , Piscina , Barbacoa
                                                                    from ubicacion, caracteristicas_inm, inmueble, barrios, publicacion 	
                                                                    where caracteristicas_inm.IdInm = inmueble.IdInm and ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio and publicacion.IdPubli = inmueble.IdPubli
                                                                    and ubicacion.Departamento = '" & departamentoagente & "'  
                                                                    and inmueble.IdInmueble = " & idinm & " and publicacion.EstadoInm = 'Activo'
                                                                    ;").Tables("consultainmueble")

        DataFormPrincipalCrearreserva.Columns(0).Visible = False




    End Sub

    Private Sub Button1_Click_2(sender As Object, e As EventArgs) Handles Button1.Click

        Dim idinm As Integer = NumericIdInmCli.Value
        idinm = idinm - 100


        Me.DataFormPrincipalBuscarInmuebles.DataSource = consultaSQL("select TipoInm, Departamento, NumDormitorios , TipoPubli , Precio , Moneda , Jardin , Amoblada , Garage , Piscina , Barbacoa
from ubicacion, caracteristicas_inm, inmueble, barrios, publicacion 	
where caracteristicas_inm.IdInm = inmueble.IdInm and ubicacion.IdUbicacion = barrios.IdUbicacion and barrios.IdBarrio = inmueble.IdBarrio and publicacion.IdPubli = inmueble.IdPubli
and inmueble.IdInmueble = " & idinm & " and publicacion.EstadoInm = 'Activo'

;").Tables("consultainmueble")

    End Sub




    Private Sub PictureBox1_Click_3(sender As Object, e As EventArgs) Handles PicCerrarqr.Click
        PanelQRReserva.Visible = False
        PicQRAgenda.Visible = False
        PicCerrarqr.Visible = False

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        If Idinmueble = "" Then
            MsgBox("No se selecciono ningun inmueble")
        Else

            'TIPO DE DOCUMENTO
            ComboFormPrincipalIngresoPropiedadesTipodoc.Text = consultaSQL("SELECT TipoDoc FROM personas , tiene , inmueble where personas.IdPersona = tiene.IdPersona and inmueble.IdInmueble = tiene.IdInmueble
             and inmueble.IdInmueble = " & Idinmueble & ";").Tables("consultainmueble").Rows(0)("TipoDoc").ToString()

            'NUMERO DOCUMENTO
            TxtFormPrincipalIngresoPropiedadesDocumentopropietario.Text = consultaSQL("SELECT NumDoc FROM personas , tiene , inmueble where personas.IdPersona = tiene.IdPersona and inmueble.IdInmueble = tiene.IdInmueble
             and inmueble.IdInmueble = " & Idinmueble & ";").Tables("consultainmueble").Rows(0)("NumDoc").ToString()


            'DEPARTAMENTO DEL INMUEBLE
            ComboFormPrincipalIngresoPropiedadesDepartamento.Text = consultaSQL("select Departamento from ubicacion , barrios , inmueble where inmueble.IdBarrio = barrios.IdBarrio and barrios.IdUbicacion = ubicacion.IdUbicacion
            and inmueble.IdInmueble = " & Idinmueble & ";").Tables("consultainmueble").Rows(0)("Departamento").ToString()


            'CIUDAD DEL INMUEBLE
            ComboFormPrincipalIngresoPropiedadesCiudad.Text = consultaSQL("select NomCiudad from ubicacion , barrios , inmueble where inmueble.IdBarrio = barrios.IdBarrio and barrios.IdUbicacion = ubicacion.IdUbicacion
            and inmueble.IdInmueble = " & Idinmueble & ";").Tables("consultainmueble").Rows(0)("NomCiudad").ToString()


            'BARRIO DEL INMUEBLE
            ComboFormPrincipalIngresoPropiedadesBarrios.Text = consultaSQL("select Barrios from ubicacion , barrios , inmueble where inmueble.IdBarrio = barrios.IdBarrio and barrios.IdUbicacion = ubicacion.IdUbicacion
            and inmueble.IdInmueble = " & Idinmueble & ";").Tables("consultainmueble").Rows(0)("Barrios").ToString()

            'CALLE DEL INMUEBLE
            TxtFormPrincipalIngresoPropiedadesCalle.Text = consultaSQL("SELECT Calle FROM inmueble where IdInmueble = " & Idinmueble & "
            ;").Tables("consultainmueble").Rows(0)("Calle").ToString()


            'ESQUINA DEL INMUEBLE
            TxtFormPrincipalIngresoPropiedadesEsquina.Text = consultaSQL("SELECT Esquina FROM inmueble where IdInmueble = " & Idinmueble & "
            ;").Tables("consultainmueble").Rows(0)("Esquina").ToString()

            'NUMERO DE PUERTA DEL INMUEBLE
            NumericFormPrincipalIngresoPropiedadesNumeopuerta.Value = consultaSQL("SELECT NroPuerta FROM inmueble where IdInmueble = " & Idinmueble & "
            ;").Tables("consultainmueble").Rows(0)("NroPuerta")


            'NUMERO DE APARTAMENTO
            NumericFormPrincipalIngresoPropiedadesNumeroAp.Value = consultaSQL("SELECT NroApto FROM inmueble where IdInmueble = " & Idinmueble & "
            ;").Tables("consultainmueble").Rows(0)("NroApto")


            'BARBACOA
            If consultaSQL("SELECT Barbacoa FROM caracteristicas_inm , inmueble where caracteristicas_inm.IdInm = inmueble.IdInmueble and IdInmueble  = " & Idinmueble & "
            ;").Tables("consultainmueble").Rows(0)("Barbacoa").ToString() = "Si" Then

                CheckFormPrincipalIngresoPropiedadesBarbacoa.Checked = True

            End If


            'JARDIN
            If consultaSQL("SELECT Jardin FROM caracteristicas_inm , inmueble where caracteristicas_inm.IdInm = inmueble.IdInmueble and IdInmueble  = " & Idinmueble & "
            ;").Tables("consultainmueble").Rows(0)("Jardin").ToString() = "Si" Then

                CheckFormPrincipalIngresoPropiedadesJardin.Checked = True

            End If



            'GARAGE
            If consultaSQL("SELECT Garage FROM caracteristicas_inm , inmueble where caracteristicas_inm.IdInm = inmueble.IdInmueble and IdInmueble  = " & Idinmueble & "
            ;").Tables("consultainmueble").Rows(0)("Garage").ToString() = "Si" Then

                CheckFormPrincipalIngresoPropiedadesGarage.Checked = True

            End If


            'PISCINA
            If consultaSQL("SELECT Piscina FROM caracteristicas_inm , inmueble where caracteristicas_inm.IdInm = inmueble.IdInmueble and IdInmueble  = " & Idinmueble & "
            ;").Tables("consultainmueble").Rows(0)("Piscina").ToString() = "Si" Then

                CheckFormPrincipalIngresoPropiedadesPiscina.Checked = True

            End If


            'AMOBLADA

            If consultaSQL("SELECT amoblada FROM caracteristicas_inm , inmueble where caracteristicas_inm.IdInm = inmueble.IdInmueble and IdInmueble  = " & Idinmueble & "
            ;").Tables("consultainmueble").Rows(0)("amoblada").ToString() = "Si" Then

                CheckFormPrincipalIngresoPropiedadesAmoblada.Checked = True

            End If


            'MASCOTAS

            If consultaSQL("SELECT AceptaMascota FROM caracteristicas_inm , inmueble where caracteristicas_inm.IdInm = inmueble.IdInmueble and IdInmueble  = " & Idinmueble & "
            ;").Tables("consultainmueble").Rows(0)("AceptaMascota").ToString() = "Si" Then

                CheckFormPrincipalIngresoPropiedadesMascotas.Checked = True

            End If



            'NOMBRE INMUEBLE
            TxtFormPrincipalIngresoPropiedadesNombreInmueble.Text = consultaSQL("SELECT NomInmueble FROM inmueble where IdInmueble = " & Idinmueble & "
            ;").Tables("consultainmueble").Rows(0)("NomInmueble").ToString




            'SUPERFICIE TOTAL CONSTRUIDA
            NumericTxtFormPrincipalIngresoPropiedadesSuperficieTotal.Value = consultaSQL("SELECT SupParcela FROM caracteristicas_inm , inmueble where caracteristicas_inm.IdInm = inmueble.IdInmueble 
            and IdInmueble = " & Idinmueble & ";").Tables("consultainmueble").Rows(0)("SupParcela")


            'SUPERFICIE CONSTRUIDA
            NumericTxtFormPrincipalIngresoPropiedadesSuperficieEdificada.Value = consultaSQL("SELECT SupConstruida FROM caracteristicas_inm , inmueble where caracteristicas_inm.IdInm = inmueble.IdInmueble 
            and IdInmueble = " & Idinmueble & ";").Tables("consultainmueble").Rows(0)("SupConstruida")


            'BAÑOS
            NumericFormPrincipalIngresoPropiedadesCantbaños.Value = consultaSQL("SELECT NumBaños FROM caracteristicas_inm , inmueble where caracteristicas_inm.IdInm = inmueble.IdInmueble 
            and IdInmueble = " & Idinmueble & ";").Tables("consultainmueble").Rows(0)("NumBaños")


            'CUARTOS
            NumericFormPrincipalIngresoPropiedadesCantcuartos.Value = consultaSQL("SELECT NumDormitorios FROM caracteristicas_inm , inmueble where caracteristicas_inm.IdInm = inmueble.IdInmueble 
            and IdInmueble = " & Idinmueble & ";").Tables("consultainmueble").Rows(0)("NumDormitorios")


            'TIPO INMUEBLE
            ComboFormPrincipalIngresoPropiedadesTipoInmueble.Text = consultaSQL("SELECT TipoInm FROM caracteristicas_inm , inmueble where caracteristicas_inm.IdInm = inmueble.IdInmueble and 
           IdInmueble = " & Idinmueble & ";").Tables("consultainmueble").Rows(0)("TipoInm").ToString()


            'GARANTIA
            ComboFormPrincipalIngresoPropiedadesGarantia.Text = consultaSQL("SELECT Garantia FROM caracteristicas_inm , inmueble where caracteristicas_inm.IdInm = inmueble.IdInmueble and 
           IdInmueble = " & Idinmueble & ";").Tables("consultainmueble").Rows(0)("Garantia").ToString()


            'TIPO PUBLICACION
            ComboFormPrincipalIngresoPropiedadesTipopubli.Text = consultaSQL("SELECT TipoPubli FROM publicacion , inmueble where publicacion.IdPubli = inmueble.IdPubli and
            IdInmueble = " & Idinmueble & ";").Tables("consultainmueble").Rows(0)("TipoPubli").ToString()

            'PRECIO
            NumericComboFormPrincipalIngresoPropiedadesPrecio.Value = consultaSQL("SELECT Precio FROM publicacion , inmueble where publicacion.IdPubli = inmueble.IdPubli and
            IdInmueble = " & Idinmueble & ";").Tables("consultainmueble").Rows(0)("Precio")

            'MONEDA
            ComboFormPrincipalIngresoPropiedadesMoneda.Text = consultaSQL("SELECT Moneda FROM publicacion , inmueble where publicacion.IdPubli = inmueble.IdPubli and
            IdInmueble = " & Idinmueble & ";").Tables("consultainmueble").Rows(0)("Moneda").ToString()


            'DESCRIPCION
            TxtFormPrincipalIngresoPropiedadesDescripcion.Text = consultaSQL("SELECT Descripción FROM caracteristicas_inm , inmueble where caracteristicas_inm.IdInm = inmueble.IdInmueble and 
           IdInmueble = " & Idinmueble & ";").Tables("consultainmueble").Rows(0)("Descripción").ToString()



            ' BOTON SIGUIENTE
            'Calcula la cantidad de imagenes que tiene el inmueble
            Dim cantidadimagenes As String = consultaSQL("Select count(Imagen) from img_inmuebles where IdInmueble=" & DataFormPrincipalBuscarInmuebles.CurrentRow.Cells("IdInmueble").Value & " ").Tables("img_inmuebles").Rows(0)("count(Imagen)").ToString()
            Dim cantidadimagenesInt As Integer = CInt(cantidadimagenes)

            DataFormPrincipalIngresoImagenes.Rows.Clear()


            For index As Integer = 0 To cantidadimagenesInt - 1


                DataFormPrincipalIngresoImagenes.Rows.Add()

                mostrarimagenes(index)


                DataFormPrincipalIngresoImagenes.Rows(index).Height = 80





            Next



            'Calcula la cantidad de imagenes que tiene el inmueble
            cantidadimagenesedit = consultaSQL("Select count(Imagen) from img_inmuebles where IdInmueble=" & DataFormPrincipalBuscarInmuebles.CurrentRow.Cells("IdInmueble").Value & " ").Tables("img_inmuebles").Rows(0)("count(Imagen)").ToString()



            'cantidadimagenesedit = cantidadimagenesedit + 1



            With Me

                .PanelFormPrincipalBuscarInmuebles.Visible = False
                .BtnFormPrincipalBuscar.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
                .PanelFormPrincipalCrearreserva.Visible = False
                .BtnPrincipalCrearreserva.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
                .PanelFormPrincipalConsultaragenda.Visible = False
                .BtnFormPrincipalConsultarAgenda.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
                .FormPrincipalRegistroFechanacimiento.Visible = False
                .BtnFormPrincipalRegistrarInteresado.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
                .PanelFormPrincipalHistorialreservas.Visible = False
                .BtnFormPrincipalHistorialVisitas.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))

                .PanelIngresoPropiedades.Visible = True
                .BtnFormPrincipalIngresoPropiedades.ForeColor = System.Drawing.Color.FromArgb(CType(CType(139, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))



            End With










        End If






    End Sub

    Private Sub DataFormPrincipalBuscarInmuebles_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataFormPrincipalBuscarInmuebles.CellClick

        Try
            Idinmueble = DataFormPrincipalBuscarInmuebles.CurrentRow.Cells("IdInmueble").Value

        Catch ex As Exception

        End Try




    End Sub



    'MUESTRA LAS IMAGENES QUE TIENE UN INMUEBLE
    Public Sub mostrarimagen(ByVal numimagen As Integer)


        Try


            Dim sql As String = "Select Imagen from img_inmuebles where IdInmueble=" & DataFormPrincipalBuscarInmuebles.CurrentRow.Cells("IdInmueble").Value & " "
            conectar()
            Dim adaptador As New MySqlDataAdapter(sql, conexion)
            Dim tabla As New DataTable
            adaptador.Fill(tabla)
            Dim imgByte() As Byte

            imgByte = tabla(numimagen)(0)
            Dim ms As New MemoryStream(imgByte)

            PicBuscarInmuebleImaenes.Image = Image.FromStream(ms)
        Catch ex As Exception

            MsgBox("No hay imagenes")

        End Try





    End Sub



    'MUESTRA LAS IMAGENES QUE TIENE UN INMUEBLE
    Public Sub mostrarimagenfull(ByVal numimagen As Integer)


        Try


            Dim sql As String = "Select Imagen from img_inmuebles where IdInmueble=" & DataFormPrincipalBuscarInmuebles.CurrentRow.Cells("IdInmueble").Value & " "
            conectar()
            Dim adaptador As New MySqlDataAdapter(sql, conexion)
            Dim tabla As New DataTable
            adaptador.Fill(tabla)
            Dim imgByte() As Byte

            imgByte = tabla(numimagen)(0)
            Dim ms As New MemoryStream(imgByte)

            imagenesFullPantalla.PicImagenesFullPantalla.Image = Image.FromStream(ms)
        Catch ex As Exception


            MsgBox("No hay imagenes")

        End Try


    End Sub





    'MUESTRA LAS IMAGENES QUE TIENE UN INMUEBLE
    Public Sub mostrarimagenes(ByVal numimagen As Integer)




        Try


            Dim sql As String = "Select Imagen from img_inmuebles where IdInmueble=" & DataFormPrincipalBuscarInmuebles.CurrentRow.Cells("IdInmueble").Value & " "

            conectar()
            Dim adaptador As New MySqlDataAdapter(sql, conexion)
            Dim tabla As New DataTable
            adaptador.Fill(tabla)
            Dim imgByte() As Byte

            imgByte = tabla(numimagen)(0)
            Dim ms As New MemoryStream(imgByte)



            DataFormPrincipalIngresoImagenes.Rows(numimagen).Cells(0).Value = Image.FromStream(ms)











        Catch ex As Exception


            MsgBox(ex.Message)

        End Try


    End Sub










    Private Sub BtnFormPrincipalBuscarInmueblesAmpliar_Click(sender As Object, e As EventArgs) Handles BtnFormPrincipalBuscarInmueblesAmpliar.Click

        PicBuscarInmuebleImaenes.Image = Nothing



        If Idinmueble = "" Then
            MsgBox("No se selecciono inmueble")
        Else




            'llama al metodo para mostrar la primera imagen
            mostrarimagen(cantidadimg)
            PanelBuscarInmueblesImagenes.Visible = True





        End If


    End Sub

    Private Sub PictureBox1_Click_2(sender As Object, e As EventArgs) Handles PictureBox1.Click
        PanelBuscarInmueblesImagenes.Visible = False
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles BtnBuscarinmuebleimagenessiguiente.Click


        ' BOTON SIGUIENTE
        'Calcula la cantidad de imagenes que tiene el inmueble
        Dim cantidadimagenes As String = consultaSQL("Select count(Imagen) from img_inmuebles where IdInmueble=" & DataFormPrincipalBuscarInmuebles.CurrentRow.Cells("IdInmueble").Value & " ").Tables("img_inmuebles").Rows(0)("count(Imagen)").ToString()
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

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles BtnBuscarinmuebleimagenesanterior.Click


        Dim cantidadimagenes As String = consultaSQL("Select count(Imagen) from img_inmuebles where IdInmueble=" & DataFormPrincipalBuscarInmuebles.CurrentRow.Cells("IdInmueble").Value & " ").Tables("img_inmuebles").Rows(0)("count(Imagen)").ToString()
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

    Private Sub PicBuscarInmuebleImaenes_Click(sender As Object, e As EventArgs) Handles PicBuscarInmuebleImaenes.Click

        imagenesFullPantalla.BackColor = System.Drawing.Color.FromArgb(CType(CType(38, Byte), Integer), CType(CType(55, Byte), Integer), CType(CType(52, Byte), Integer))

        imagenesFullPantalla.PicImagenesFullPantalla.Image = PicBuscarInmuebleImaenes.Image


        imagenesFullPantalla.Show()






    End Sub

    Private Sub PanelFormPrincipalBuscarInmuebles_Paint(sender As Object, e As PaintEventArgs) Handles PanelFormPrincipalBuscarInmuebles.Paint

    End Sub

    Private Sub Panel1_Paint(sender As Object, e As PaintEventArgs)

    End Sub

    Private Sub Button8_Click_1(sender As Object, e As EventArgs) Handles BtnFormPrincipalCrearHorarios.Click

        With Me

            .PanelFormPrincipalBuscarInmuebles.Visible = False
            .BtnFormPrincipalBuscar.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
            .PanelFormPrincipalCrearreserva.Visible = False
            .BtnPrincipalCrearreserva.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
            .PanelFormPrincipalConsultaragenda.Visible = False
            .BtnFormPrincipalConsultarAgenda.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
            .PanelFormPrincipalHistorialreservas.Visible = False
            .BtnFormPrincipalHistorialVisitas.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
            .PanelIngresoPropiedades.Visible = False
            .BtnFormPrincipalIngresoPropiedades.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
            .FormPrincipalRegistroFechanacimiento.Visible = False
            .BtnFormPrincipalRegistrarInteresado.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))


            .PanelFormPrincipalHorarios.Visible = True
            .BtnFormPrincipalCrearHorarios.ForeColor = System.Drawing.Color.FromArgb(CType(CType(139, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))

        End With



        'CARGA LOS NOMBRES DE LOS AGENTES
        ComboBox1.DataSource = (consultaSQL("select concat(concat(Nombre , (' ')) , Apellido ) from personas where Funcion = 'Agente';").Tables(0))
        ComboBox1.DisplayMember = "concat(concat(Nombre , (' ')) , Apellido )"



        Try



            Dim IdPersona As String = consultaSQL("SELECT IdPersona FROM enigma.personas where  concat(concat(Nombre , (' '))  , Apellido ) = '" & ComboBox1.Text & "';").Tables("personas").Rows(0)("IdPersona").ToString()

            Dim ultimodiames As String = Format(UltimoDiaDelMes(mifecha.Now), "dd")

            DataGridView1.Rows.Clear()
            DataGridView1.Rows.Add()
            DataGridView1.Rows(0).Height = 30

            DataGridView1.Rows.Add()
            DataGridView1.Rows(1).Height = 30
            DataGridView1.Rows.Add()
            DataGridView1.Rows(2).Height = 30
            DataGridView1.Rows.Add()
            DataGridView1.Rows(3).Height = 30
            DataGridView1.Rows(4).Height = 30



            'PRIMER SEMANA
            For i As Integer = 0 To 6

                primerdiames = Format(PrimerDiaDelMes(mifecha.Now).AddDays(i), "yyy-MM-dd")
                DataGridView1.Rows.Item(0).Cells(i).Value = primerdiames

                'CAMBIA EL COLOR SI TRABAJA ESE DIA
                Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                If encontroagenda <> 0 Then
                    DataGridView1.Rows.Item(0).Cells(i).Style.BackColor = System.Drawing.Color.Green
                End If
            Next






            For i As Integer = 0 To 6



                primerdiames = Format(PrimerDiaDelMes(mifecha.Now).AddDays(7 + i), "yyy-MM-dd")

                'LLENA LAS CELDAS CON LAS FECHAS
                DataGridView1.Rows.Item(1).Cells(i).Value = primerdiames




                Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                If encontroagenda <> 0 Then

                    'CAMBIA EL COLOR
                    DataGridView1.Rows.Item(1).Cells(i).Style.BackColor = System.Drawing.Color.Green


                End If





            Next



            For j As Integer = 0 To 6


                primerdiames = Format(PrimerDiaDelMes(mifecha.Now).AddDays(14 + j), "yyy-MM-dd")


                DataGridView1.Rows.Item(2).Cells(j).Value = primerdiames



                Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                If encontroagenda <> 0 Then


                    'CAMBIA EL COLOR
                    DataGridView1.Rows.Item(2).Cells(j).Style.BackColor = System.Drawing.Color.Green


                End If

            Next






            For k As Integer = 0 To 6

                primerdiames = Format(PrimerDiaDelMes(mifecha.Now).AddDays(21 + k), "yyy-MM-dd")

                DataGridView1.Rows.Item(3).Cells(k).Value = primerdiames


                Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                If encontroagenda <> 0 Then


                    'CAMBIA EL COLOR
                    DataGridView1.Rows.Item(3).Cells(k).Style.BackColor = System.Drawing.Color.Green


                End If




            Next





            If ultimodiames = 28 Then


            Else



                'QUINTA SEMANA
                If ultimodiames = 31 Then

                    For k As Integer = 0 To 2
                        primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(28 + k), "yyy-MM-dd")
                        DataGridView1.Rows.Item(4).Cells(k).Value = primerdiames
                        'CAMBIA EL COLOR SI TRABAJA ESE DIA
                        Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                        If encontroagenda <> 0 Then
                            DataGridView1.Rows.Item(4).Cells(k).Style.BackColor = System.Drawing.Color.Green
                        End If

                    Next
                End If



                'QUINTA SEMANA
                If ultimodiames = 30 Then

                    For k As Integer = 0 To 1
                        primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(28 + k), "yyy-MM-dd")
                        DataGridView1.Rows.Item(4).Cells(k).Value = primerdiames
                        'CAMBIA EL COLOR SI TRABAJA ESE DIA
                        Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                        If encontroagenda <> 0 Then
                            DataGridView1.Rows.Item(4).Cells(k).Style.BackColor = System.Drawing.Color.Green
                        End If

                    Next
                End If




                'QUINTA SEMANA
                If ultimodiames = 29 Then

                    For k As Integer = 0 To 0
                        primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(28 + k), "yyy-MM-dd")
                        DataGridView1.Rows.Item(4).Cells(k).Value = primerdiames

                        'CAMBIA EL COLOR SI TRABAJA ESE DIA
                        Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                        If encontroagenda <> 0 Then
                            DataGridView1.Rows.Item(4).Cells(k).Style.BackColor = System.Drawing.Color.Green
                        End If

                    Next
                End If

            End If












        Catch ex As Exception

        End Try


        DataGridView1.ClearSelection()


    End Sub














    'para saber el primer dia del mes 
    Function PrimerDiaDelMes(ByVal dtmFecha As Date) As Date
        PrimerDiaDelMes = DateSerial(Year(dtmFecha), Month(dtmFecha), 1)
End Function
 
'para saber el ultimo dia del mes
Function UltimoDiaDelMes(ByVal dtmFecha As Date) As Date
     UltimoDiaDelMes = DateSerial(Year(dtmFecha), Month(dtmFecha) + 1, 0)
End Function

    Private Sub Button8_Click_2(sender As Object, e As EventArgs)


        restameses = 0
        sumameses = sumameses + 1

        DataGridView1.Rows.Clear()
        DataGridView1.Rows.Add()
        DataGridView1.Rows(0).Height = 30

        DataGridView1.Rows.Add()
        DataGridView1.Rows(1).Height = 30
        DataGridView1.Rows.Add()
        DataGridView1.Rows(2).Height = 30
        DataGridView1.Rows.Add()
        DataGridView1.Rows(3).Height = 30
        DataGridView1.Rows(4).Height = 30


        Try



            Dim IdPersona As String = consultaSQL("SELECT IdPersona FROM enigma.personas where  concat(concat(Nombre , (' '))  , Apellido ) = '" & ComboBox1.Text & "';").Tables("personas").Rows(0)("IdPersona").ToString()


            primerdiames = Format(PrimerDiaDelMes(mifecha.Now).AddMonths(sumameses), "yyy-MM-dd")

            Dim ultimodiames As String = Format(UltimoDiaDelMes(primerdiames), "dd")





            'PRIMER SEMANA
            For i As Integer = 0 To 6
                primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(i), "yyy-MM-dd")
                DataGridView1.Rows.Item(0).Cells(i).Value = primerdiames


                'CAMBIA EL COLOR SI TRABAJA ESE DIA
                Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                If encontroagenda <> 0 Then
                    DataGridView1.Rows.Item(0).Cells(i).Style.BackColor = System.Drawing.Color.Green
                End If

            Next


            'SEGUNDA SEMANA
            For i As Integer = 0 To 6
                primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(7 + i), "yyy-MM-dd")
                DataGridView1.Rows.Item(1).Cells(i).Value = primerdiames
                'CAMBIA EL COLOR SI TRABAJA ESE DIA
                Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                If encontroagenda <> 0 Then
                    DataGridView1.Rows.Item(1).Cells(i).Style.BackColor = System.Drawing.Color.Green
                End If

            Next


            'TERCER SEMANA
            For j As Integer = 0 To 6
                primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(14 + j), "yyy-MM-dd")
                DataGridView1.Rows.Item(2).Cells(j).Value = primerdiames
                'CAMBIA EL COLOR SI TRABAJA ESE DIA
                Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                If encontroagenda <> 0 Then
                    DataGridView1.Rows.Item(2).Cells(j).Style.BackColor = System.Drawing.Color.Green
                End If

            Next



            'CUARTA SEMANA
            For k As Integer = 0 To 6
                primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(21 + k), "yyy-MM-dd")
                DataGridView1.Rows.Item(3).Cells(k).Value = primerdiames
                'CAMBIA EL COLOR SI TRABAJA ESE DIA
                Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                If encontroagenda <> 0 Then
                    DataGridView1.Rows.Item(3).Cells(k).Style.BackColor = System.Drawing.Color.Green
                End If

            Next






            If ultimodiames = 28 Then


            Else



                'QUINTA SEMANA
                If ultimodiames = 31 Then

                    For k As Integer = 0 To 2
                        primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(28 + k), "yyy-MM-dd")
                        DataGridView1.Rows.Item(4).Cells(k).Value = primerdiames
                        'CAMBIA EL COLOR SI TRABAJA ESE DIA
                        Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                        If encontroagenda <> 0 Then
                            DataGridView1.Rows.Item(4).Cells(k).Style.BackColor = System.Drawing.Color.Green
                        End If

                    Next
                End If



                'QUINTA SEMANA
                If ultimodiames = 30 Then

                    For k As Integer = 0 To 1
                        primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(28 + k), "yyy-MM-dd")
                        DataGridView1.Rows.Item(4).Cells(k).Value = primerdiames
                        'CAMBIA EL COLOR SI TRABAJA ESE DIA
                        Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                        If encontroagenda <> 0 Then
                            DataGridView1.Rows.Item(4).Cells(k).Style.BackColor = System.Drawing.Color.Green
                        End If

                    Next
                End If




                'QUINTA SEMANA
                If ultimodiames = 29 Then

                    For k As Integer = 0 To 0
                        primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(28 + k), "yyy-MM-dd")
                        DataGridView1.Rows.Item(4).Cells(k).Value = primerdiames

                        'CAMBIA EL COLOR SI TRABAJA ESE DIA
                        Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                        If encontroagenda <> 0 Then
                            DataGridView1.Rows.Item(4).Cells(k).Style.BackColor = System.Drawing.Color.Green
                        End If

                    Next
                End If

            End If





            DataGridView1.ClearSelection()



        Catch ex As Exception

        End Try


    End Sub

    Private Sub Button9_Click_1(sender As Object, e As EventArgs)
        sumameses = 0

        restameses = restameses + 1


        DataGridView1.Rows.Clear()
        DataGridView1.Rows.Add()
        DataGridView1.Rows(0).Height = 30

        DataGridView1.Rows.Add()
        DataGridView1.Rows(1).Height = 30
        DataGridView1.Rows.Add()
        DataGridView1.Rows(2).Height = 30
        DataGridView1.Rows.Add()
        DataGridView1.Rows(3).Height = 30
        DataGridView1.Rows(4).Height = 30



        Try



            Dim IdPersona As String = consultaSQL("SELECT IdPersona FROM enigma.personas where  concat(concat(Nombre , (' '))  , Apellido ) = '" & ComboBox1.Text & "';").Tables("personas").Rows(0)("IdPersona").ToString()



            primerdiames = Format(PrimerDiaDelMes(mifecha.Now).AddMonths(-restameses), "yyy-MM-dd")

            Dim ultimodiames As String = Format(UltimoDiaDelMes(primerdiames), "dd")





            'PRIMER SEMANA
            For i As Integer = 0 To 6
                primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(i), "yyy-MM-dd")
                DataGridView1.Rows.Item(0).Cells(i).Value = primerdiames

                'CAMBIA EL COLOR SI TRABAJA ESE DIA
                Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                If encontroagenda <> 0 Then
                    DataGridView1.Rows.Item(0).Cells(i).Style.BackColor = System.Drawing.Color.Green
                End If

            Next


            'SEGUNDA SEMANA
            For i As Integer = 0 To 6
                primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(7 + i), "yyy-MM-dd")
                DataGridView1.Rows.Item(1).Cells(i).Value = primerdiames

                'CAMBIA EL COLOR SI TRABAJA ESE DIA
                Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                If encontroagenda <> 0 Then
                    DataGridView1.Rows.Item(1).Cells(i).Style.BackColor = System.Drawing.Color.Green
                End If

            Next


            'TERCER SEMANA
            For j As Integer = 0 To 6
                primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(14 + j), "yyy-MM-dd")
                DataGridView1.Rows.Item(2).Cells(j).Value = primerdiames

                'CAMBIA EL COLOR SI TRABAJA ESE DIA
                Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                If encontroagenda <> 0 Then
                    DataGridView1.Rows.Item(2).Cells(j).Style.BackColor = System.Drawing.Color.Green
                End If

            Next



            'CUARTA SEMANA
            For k As Integer = 0 To 6
                primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(21 + k), "yyy-MM-dd")
                DataGridView1.Rows.Item(3).Cells(k).Value = primerdiames

                'CAMBIA EL COLOR SI TRABAJA ESE DIA
                Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                If encontroagenda <> 0 Then
                    DataGridView1.Rows.Item(3).Cells(k).Style.BackColor = System.Drawing.Color.Green
                End If

            Next







            If ultimodiames = 28 Then


            Else



                'QUINTA SEMANA
                If ultimodiames = 31 Then

                    For k As Integer = 0 To 2
                        primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(28 + k), "yyy-MM-dd")
                        DataGridView1.Rows.Item(4).Cells(k).Value = primerdiames
                        'CAMBIA EL COLOR SI TRABAJA ESE DIA
                        Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                        If encontroagenda <> 0 Then
                            DataGridView1.Rows.Item(4).Cells(k).Style.BackColor = System.Drawing.Color.Green
                        End If

                    Next
                End If



                'QUINTA SEMANA
                If ultimodiames = 30 Then

                    For k As Integer = 0 To 1
                        primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(28 + k), "yyy-MM-dd")
                        DataGridView1.Rows.Item(4).Cells(k).Value = primerdiames
                        'CAMBIA EL COLOR SI TRABAJA ESE DIA
                        Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                        If encontroagenda <> 0 Then
                            DataGridView1.Rows.Item(4).Cells(k).Style.BackColor = System.Drawing.Color.Green
                        End If

                    Next
                End If




                'QUINTA SEMANA
                If ultimodiames = 29 Then

                    For k As Integer = 0 To 0
                        primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(28 + k), "yyy-MM-dd")
                        DataGridView1.Rows.Item(4).Cells(k).Value = primerdiames

                        'CAMBIA EL COLOR SI TRABAJA ESE DIA
                        Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                        If encontroagenda <> 0 Then
                            DataGridView1.Rows.Item(4).Cells(k).Style.BackColor = System.Drawing.Color.Green
                        End If

                    Next
                End If

            End If





            DataGridView1.ClearSelection()


        Catch ex As Exception

        End Try


    End Sub



    Private Sub DateTimePicker1_ValueChanged_1(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged


        inicialfechaagente = Me.DateTimePicker1.Value

        Fechahinicialagente = Format(inicialfechaagente, "yyy-MM-dd")






    End Sub

    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged

        finalfechaagente = Me.DateTimePicker2.Value

        Fechafinalagente = Format(finalfechaagente, "yyy-MM-dd")





    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

        If ComboBox2.Text = "Mañana" Then


            horainicialagente = "08:00:00"
            horafinalagente = "14:00:00"

        End If


        If ComboBox2.Text = "Tarde" Then

            horainicialagente = "14:00:000"
            horafinalagente = "21:00:00"

        End If






    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click


        Try



            If Fechahinicialagente > finalfechaagente Then

                MsgBox("Fecha inicial mayor a la final")

            Else



                If ComboBox2.Text = "" Then
                    MsgBox("No se eligio turno")

                Else


                    Dim fechacont As Date = Format(inicialfechaagente.AddDays(-1), "yyy-MM-dd")



                    While fechacont <= Format(finalfechaagente.AddDays(-1), "yyy-MM-dd")


                        fechacont = Format(fechacont.AddDays(1), "yyy-MM-dd")

                        'Fechahinicialagente = Format(inicialfechaagente, "yyy-MM-dd")



                        Dim IdPersona As String = consultaSQL("SELECT IdPersona FROM enigma.personas where  concat(concat(Nombre , (' '))  , Apellido ) = '" & ComboBox1.Text & "';").Tables("personas").Rows(0)("IdPersona").ToString()






                        Dim encontroagenda As String = consultaSQL("select count(IdAgenda) from agenda where IdAgente = " & IdPersona & " and Turno = '" & ComboBox2.Text & "' and FechaAgente = '" & Format(fechacont, "yyy-MM-dd") & "' and HoraEntrada = '" & horainicialagente & "'
                            and HoraSalida = '" & horafinalagente & "' and Asiste = 'si';").Tables("quebarrioshay").Rows(0)("count(IdAgenda)").ToString()





                        If encontroagenda = 0 Then

                            'MsgBox(IdPersona & "  " & Format(fechacont, "yyy-MM-dd") & " " & horainicialagente & " " & horafinalagente)


                            updateSQL("insert into agenda ( IdAgente ,  Turno , FechaAgente , HoraEntrada , HoraSalida , Asiste  ) values

                             ( " & IdPersona & " , '" & ComboBox2.Text & "' , '" & Format(fechacont, "yyy-MM-dd") & "' , '" & horainicialagente & "' , '" & horafinalagente & "' , 'si');")


                        End If










                    End While









                End If








            End If





        Catch ex As Exception


            MsgBox("No se eligio rango de fecha")

        End Try



        'CARGA LOS NOMBRES DE LOS AGENTES
        ComboBox1.DataSource = (consultaSQL("select concat(concat(Nombre , (' ')) , Apellido ) from personas where Funcion = 'Agente';").Tables(0))
        ComboBox1.DisplayMember = "concat(concat(Nombre , (' ')) , Apellido )"

        DateTimePicker1.Value = DateTime.Now
        DateTimePicker2.Value = DateTime.Now


        ComboBox2.Text = ""





    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

        Try



            Dim IdPersona As String = consultaSQL("SELECT IdPersona FROM enigma.personas where  concat(concat(Nombre , (' '))  , Apellido ) = '" & ComboBox1.Text & "';").Tables("personas").Rows(0)("IdPersona").ToString()

            Dim ultimodiames As String = Format(UltimoDiaDelMes(mifecha.Now), "dd")

            DataGridView1.Rows.Clear()
            DataGridView1.Rows.Add()
            DataGridView1.Rows(0).Height = 30

            DataGridView1.Rows.Add()
            DataGridView1.Rows(1).Height = 30
            DataGridView1.Rows.Add()
            DataGridView1.Rows(2).Height = 30
            DataGridView1.Rows.Add()
            DataGridView1.Rows(3).Height = 30
            DataGridView1.Rows(4).Height = 30



            'PRIMER SEMANA
            For i As Integer = 0 To 6

                primerdiames = Format(PrimerDiaDelMes(mifecha.Now).AddDays(i), "yyy-MM-dd")
                DataGridView1.Rows.Item(0).Cells(i).Value = primerdiames

                'CAMBIA EL COLOR SI TRABAJA ESE DIA
                Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                If encontroagenda <> 0 Then
                    DataGridView1.Rows.Item(0).Cells(i).Style.BackColor = System.Drawing.Color.Green
                End If
            Next






            For i As Integer = 0 To 6



                primerdiames = Format(PrimerDiaDelMes(mifecha.Now).AddDays(7 + i), "yyy-MM-dd")

                'LLENA LAS CELDAS CON LAS FECHAS
                DataGridView1.Rows.Item(1).Cells(i).Value = primerdiames




                Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                If encontroagenda <> 0 Then

                    'CAMBIA EL COLOR
                    DataGridView1.Rows.Item(1).Cells(i).Style.BackColor = System.Drawing.Color.Green


                End If





            Next



            For j As Integer = 0 To 6


                primerdiames = Format(PrimerDiaDelMes(mifecha.Now).AddDays(14 + j), "yyy-MM-dd")


                DataGridView1.Rows.Item(2).Cells(j).Value = primerdiames



                Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                If encontroagenda <> 0 Then


                    'CAMBIA EL COLOR
                    DataGridView1.Rows.Item(2).Cells(j).Style.BackColor = System.Drawing.Color.Green


                End If

            Next






            For k As Integer = 0 To 6

                primerdiames = Format(PrimerDiaDelMes(mifecha.Now).AddDays(21 + k), "yyy-MM-dd")

                DataGridView1.Rows.Item(3).Cells(k).Value = primerdiames


                Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                If encontroagenda <> 0 Then


                    'CAMBIA EL COLOR
                    DataGridView1.Rows.Item(3).Cells(k).Style.BackColor = System.Drawing.Color.Green


                End If




            Next





            If ultimodiames = 28 Then


            Else



                'QUINTA SEMANA
                If ultimodiames = 31 Then

                    For k As Integer = 0 To 2
                        primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(28 + k), "yyy-MM-dd")
                        DataGridView1.Rows.Item(4).Cells(k).Value = primerdiames
                        'CAMBIA EL COLOR SI TRABAJA ESE DIA
                        Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                        If encontroagenda <> 0 Then
                            DataGridView1.Rows.Item(4).Cells(k).Style.BackColor = System.Drawing.Color.Green
                        End If

                    Next
                End If



                'QUINTA SEMANA
                If ultimodiames = 30 Then

                    For k As Integer = 0 To 1
                        primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(28 + k), "yyy-MM-dd")
                        DataGridView1.Rows.Item(4).Cells(k).Value = primerdiames
                        'CAMBIA EL COLOR SI TRABAJA ESE DIA
                        Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                        If encontroagenda <> 0 Then
                            DataGridView1.Rows.Item(4).Cells(k).Style.BackColor = System.Drawing.Color.Green
                        End If

                    Next
                End If




                'QUINTA SEMANA
                If ultimodiames = 29 Then

                    For k As Integer = 0 To 0
                        primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(28 + k), "yyy-MM-dd")
                        DataGridView1.Rows.Item(4).Cells(k).Value = primerdiames

                        'CAMBIA EL COLOR SI TRABAJA ESE DIA
                        Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                        If encontroagenda <> 0 Then
                            DataGridView1.Rows.Item(4).Cells(k).Style.BackColor = System.Drawing.Color.Green
                        End If

                    Next
                End If

            End If



            DataGridView1.ClearSelection()


        Catch ex As Exception

        End Try


    End Sub



    Private Sub Button8_Click_3(sender As Object, e As EventArgs)
        restameses = 0
        sumameses = 0

        Try

            'CARGA LOS NOMBRES DE LOS AGENTES
            ComboBox1.DataSource = (consultaSQL("select concat(concat(Nombre , (' ')) , Apellido ) from personas where Funcion = 'Agente';").Tables(0))
            ComboBox1.DisplayMember = "concat(concat(Nombre , (' ')) , Apellido )"


            Dim IdPersona As String = consultaSQL("SELECT IdPersona FROM enigma.personas where  concat(concat(Nombre , (' '))  , Apellido ) = '" & ComboBox1.Text & "';").Tables("personas").Rows(0)("IdPersona").ToString()

            Dim ultimodiames As String = Format(UltimoDiaDelMes(mifecha.Now), "dd")

            DataGridView1.Rows.Clear()
            DataGridView1.Rows.Add()
            DataGridView1.Rows(0).Height = 30

            DataGridView1.Rows.Add()
            DataGridView1.Rows(1).Height = 30
            DataGridView1.Rows.Add()
            DataGridView1.Rows(2).Height = 30
            DataGridView1.Rows.Add()
            DataGridView1.Rows(3).Height = 30
            DataGridView1.Rows(4).Height = 30



            'PRIMER SEMANA
            For i As Integer = 0 To 6

                primerdiames = Format(PrimerDiaDelMes(mifecha.Now).AddDays(i), "yyy-MM-dd")
                DataGridView1.Rows.Item(0).Cells(i).Value = primerdiames

                'CAMBIA EL COLOR SI TRABAJA ESE DIA
                Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                If encontroagenda <> 0 Then
                    DataGridView1.Rows.Item(0).Cells(i).Style.BackColor = System.Drawing.Color.Green
                End If
            Next






            For i As Integer = 0 To 6



                primerdiames = Format(PrimerDiaDelMes(mifecha.Now).AddDays(7 + i), "yyy-MM-dd")

                'LLENA LAS CELDAS CON LAS FECHAS
                DataGridView1.Rows.Item(1).Cells(i).Value = primerdiames




                Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                If encontroagenda <> 0 Then

                    'CAMBIA EL COLOR
                    DataGridView1.Rows.Item(1).Cells(i).Style.BackColor = System.Drawing.Color.Green


                End If





            Next



            For j As Integer = 0 To 6


                primerdiames = Format(PrimerDiaDelMes(mifecha.Now).AddDays(14 + j), "yyy-MM-dd")


                DataGridView1.Rows.Item(2).Cells(j).Value = primerdiames



                Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                If encontroagenda <> 0 Then


                    'CAMBIA EL COLOR
                    DataGridView1.Rows.Item(2).Cells(j).Style.BackColor = System.Drawing.Color.Green


                End If

            Next






            For k As Integer = 0 To 6

                primerdiames = Format(PrimerDiaDelMes(mifecha.Now).AddDays(21 + k), "yyy-MM-dd")

                DataGridView1.Rows.Item(3).Cells(k).Value = primerdiames


                Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                If encontroagenda <> 0 Then


                    'CAMBIA EL COLOR
                    DataGridView1.Rows.Item(3).Cells(k).Style.BackColor = System.Drawing.Color.Green


                End If




            Next






            If ultimodiames = 28 Then


            Else



                'QUINTA SEMANA
                If ultimodiames = 31 Then

                    For k As Integer = 0 To 2
                        primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(28 + k), "yyy-MM-dd")
                        DataGridView1.Rows.Item(4).Cells(k).Value = primerdiames
                        'CAMBIA EL COLOR SI TRABAJA ESE DIA
                        Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                        If encontroagenda <> 0 Then
                            DataGridView1.Rows.Item(4).Cells(k).Style.BackColor = System.Drawing.Color.Green
                        End If

                    Next
                End If



                'QUINTA SEMANA
                If ultimodiames = 30 Then

                    For k As Integer = 0 To 1
                        primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(28 + k), "yyy-MM-dd")
                        DataGridView1.Rows.Item(4).Cells(k).Value = primerdiames
                        'CAMBIA EL COLOR SI TRABAJA ESE DIA
                        Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                        If encontroagenda <> 0 Then
                            DataGridView1.Rows.Item(4).Cells(k).Style.BackColor = System.Drawing.Color.Green
                        End If

                    Next
                End If




                'QUINTA SEMANA
                If ultimodiames = 29 Then

                    For k As Integer = 0 To 0
                        primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(28 + k), "yyy-MM-dd")
                        DataGridView1.Rows.Item(4).Cells(k).Value = primerdiames

                        'CAMBIA EL COLOR SI TRABAJA ESE DIA
                        Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                        If encontroagenda <> 0 Then
                            DataGridView1.Rows.Item(4).Cells(k).Style.BackColor = System.Drawing.Color.Green
                        End If

                    Next
                End If

            End If


            DataGridView1.ClearSelection()


        Catch ex As Exception

        End Try




    End Sub

    Private Sub PictureBox2_Click_1(sender As Object, e As EventArgs) Handles PictureBox2.Click
        PanelDisponibilidadMensual.Visible = False
    End Sub

    Private Sub Button9_Click_2(sender As Object, e As EventArgs) Handles Button9.Click
        PanelDisponibilidadMensual.Visible = True


        If FormLogin.tipoacceso = "Gerente" Or FormLogin.tipoacceso = "Administrador" Then

            IdPersonal = consultaSQL("SELECT IdPersona FROM enigma.personas where  concat(concat(Nombre , (' '))  , Apellido ) = '" & ComboFormPrincipalCrearreservaAgentes.Text & "';").Tables("personas").Rows(0)("IdPersona").ToString()

        End If


        Try





            Dim ultimodiames As String = Format(UltimoDiaDelMes(mifecha.Now), "dd")

            DataGridView2.Rows.Clear()
            DataGridView2.Rows.Add()
            DataGridView2.Rows(0).Height = 30

            DataGridView2.Rows.Add()
            DataGridView2.Rows(1).Height = 30
            DataGridView2.Rows.Add()
            DataGridView2.Rows(2).Height = 30
            DataGridView2.Rows.Add()
            DataGridView2.Rows(3).Height = 30
            DataGridView2.Rows(4).Height = 30



            'PRIMER SEMANA
            For i As Integer = 0 To 6

                primerdiames = Format(PrimerDiaDelMes(mifecha.Now).AddDays(i), "yyy-MM-dd")
                DataGridView2.Rows.Item(0).Cells(i).Value = primerdiames

                'CAMBIA EL COLOR SI TRABAJA ESE DIA
                Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersonal & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                If encontroagenda <> 0 Then
                    DataGridView2.Rows.Item(0).Cells(i).Style.BackColor = System.Drawing.Color.Green
                End If
            Next






            For i As Integer = 0 To 6



                primerdiames = Format(PrimerDiaDelMes(mifecha.Now).AddDays(7 + i), "yyy-MM-dd")

                'LLENA LAS CELDAS CON LAS FECHAS
                DataGridView2.Rows.Item(1).Cells(i).Value = primerdiames




                Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersonal & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                If encontroagenda <> 0 Then

                    'CAMBIA EL COLOR
                    DataGridView2.Rows.Item(1).Cells(i).Style.BackColor = System.Drawing.Color.Green


                End If





            Next



            For j As Integer = 0 To 6


                primerdiames = Format(PrimerDiaDelMes(mifecha.Now).AddDays(14 + j), "yyy-MM-dd")


                DataGridView2.Rows.Item(2).Cells(j).Value = primerdiames



                Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersonal & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                If encontroagenda <> 0 Then


                    'CAMBIA EL COLOR
                    DataGridView2.Rows.Item(2).Cells(j).Style.BackColor = System.Drawing.Color.Green


                End If

            Next






            For k As Integer = 0 To 6

                primerdiames = Format(PrimerDiaDelMes(mifecha.Now).AddDays(21 + k), "yyy-MM-dd")

                DataGridView2.Rows.Item(3).Cells(k).Value = primerdiames


                Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersonal & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                If encontroagenda <> 0 Then


                    'CAMBIA EL COLOR
                    DataGridView2.Rows.Item(3).Cells(k).Style.BackColor = System.Drawing.Color.Green


                End If




            Next





            If ultimodiames = 28 Then


            Else



                'QUINTA SEMANA
                If ultimodiames = 31 Then

                    For k As Integer = 0 To 2
                        primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(28 + k), "yyy-MM-dd")
                        DataGridView2.Rows.Item(4).Cells(k).Value = primerdiames
                        'CAMBIA EL COLOR SI TRABAJA ESE DIA
                        Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersonal & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                        If encontroagenda <> 0 Then
                            DataGridView2.Rows.Item(4).Cells(k).Style.BackColor = System.Drawing.Color.Green
                        End If

                    Next
                End If



                'QUINTA SEMANA
                If ultimodiames = 30 Then

                    For k As Integer = 0 To 1
                        primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(28 + k), "yyy-MM-dd")
                        DataGridView2.Rows.Item(4).Cells(k).Value = primerdiames
                        'CAMBIA EL COLOR SI TRABAJA ESE DIA
                        Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersonal & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                        If encontroagenda <> 0 Then
                            DataGridView2.Rows.Item(4).Cells(k).Style.BackColor = System.Drawing.Color.Green
                        End If

                    Next
                End If




                'QUINTA SEMANA
                If ultimodiames = 29 Then

                    For k As Integer = 0 To 0
                        primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(28 + k), "yyy-MM-dd")
                        DataGridView2.Rows.Item(4).Cells(k).Value = primerdiames

                        'CAMBIA EL COLOR SI TRABAJA ESE DIA
                        Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersonal & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                        If encontroagenda <> 0 Then
                            DataGridView2.Rows.Item(4).Cells(k).Style.BackColor = System.Drawing.Color.Green
                        End If

                    Next
                End If

            End If









        Catch ex As Exception

        End Try




        DataGridView2.ClearSelection()


    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        restameses = 0
        sumameses = 0

        Try



            Dim ultimodiames As String = Format(UltimoDiaDelMes(mifecha.Now), "dd")

            DataGridView2.Rows.Clear()
            DataGridView2.Rows.Add()
            DataGridView2.Rows(0).Height = 30

            DataGridView2.Rows.Add()
            DataGridView2.Rows(1).Height = 30
            DataGridView2.Rows.Add()
            DataGridView2.Rows(2).Height = 30
            DataGridView2.Rows.Add()
            DataGridView2.Rows(3).Height = 30
            DataGridView2.Rows(4).Height = 30



            'PRIMER SEMANA
            For i As Integer = 0 To 6

                primerdiames = Format(PrimerDiaDelMes(mifecha.Now).AddDays(i), "yyy-MM-dd")
                DataGridView2.Rows.Item(0).Cells(i).Value = primerdiames

                'CAMBIA EL COLOR SI TRABAJA ESE DIA
                Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersonal & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                If encontroagenda <> 0 Then
                    DataGridView2.Rows.Item(0).Cells(i).Style.BackColor = System.Drawing.Color.Green
                End If
            Next






            For i As Integer = 0 To 6



                primerdiames = Format(PrimerDiaDelMes(mifecha.Now).AddDays(7 + i), "yyy-MM-dd")

                'LLENA LAS CELDAS CON LAS FECHAS
                DataGridView2.Rows.Item(1).Cells(i).Value = primerdiames




                Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersonal & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                If encontroagenda <> 0 Then

                    'CAMBIA EL COLOR
                    DataGridView2.Rows.Item(1).Cells(i).Style.BackColor = System.Drawing.Color.Green


                End If





            Next



            For j As Integer = 0 To 6


                primerdiames = Format(PrimerDiaDelMes(mifecha.Now).AddDays(14 + j), "yyy-MM-dd")


                DataGridView2.Rows.Item(2).Cells(j).Value = primerdiames



                Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersonal & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                If encontroagenda <> 0 Then


                    'CAMBIA EL COLOR
                    DataGridView2.Rows.Item(2).Cells(j).Style.BackColor = System.Drawing.Color.Green


                End If

            Next






            For k As Integer = 0 To 6

                primerdiames = Format(PrimerDiaDelMes(mifecha.Now).AddDays(21 + k), "yyy-MM-dd")

                DataGridView2.Rows.Item(3).Cells(k).Value = primerdiames


                Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersonal & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                If encontroagenda <> 0 Then


                    'CAMBIA EL COLOR
                    DataGridView2.Rows.Item(3).Cells(k).Style.BackColor = System.Drawing.Color.Green


                End If




            Next






            If ultimodiames = 28 Then


            Else



                'QUINTA SEMANA
                If ultimodiames = 31 Then

                    For k As Integer = 0 To 2
                        primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(28 + k), "yyy-MM-dd")
                        DataGridView2.Rows.Item(4).Cells(k).Value = primerdiames
                        'CAMBIA EL COLOR SI TRABAJA ESE DIA
                        Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersonal & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                        If encontroagenda <> 0 Then
                            DataGridView2.Rows.Item(4).Cells(k).Style.BackColor = System.Drawing.Color.Green
                        End If

                    Next
                End If



                'QUINTA SEMANA
                If ultimodiames = 30 Then

                    For k As Integer = 0 To 1
                        primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(28 + k), "yyy-MM-dd")
                        DataGridView2.Rows.Item(4).Cells(k).Value = primerdiames
                        'CAMBIA EL COLOR SI TRABAJA ESE DIA
                        Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersonal & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                        If encontroagenda <> 0 Then
                            DataGridView2.Rows.Item(4).Cells(k).Style.BackColor = System.Drawing.Color.Green
                        End If

                    Next
                End If




                'QUINTA SEMANA
                If ultimodiames = 29 Then

                    For k As Integer = 0 To 0
                        primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(28 + k), "yyy-MM-dd")
                        DataGridView2.Rows.Item(4).Cells(k).Value = primerdiames

                        'CAMBIA EL COLOR SI TRABAJA ESE DIA
                        Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersonal & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                        If encontroagenda <> 0 Then
                            DataGridView2.Rows.Item(4).Cells(k).Style.BackColor = System.Drawing.Color.Green
                        End If

                    Next
                End If

            End If









        Catch ex As Exception

        End Try




        DataGridView2.ClearSelection()
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click

        sumameses = 0

        restameses = restameses + 1


        DataGridView2.Rows.Clear()
        DataGridView2.Rows.Add()
        DataGridView2.Rows(0).Height = 30

        DataGridView2.Rows.Add()
        DataGridView2.Rows(1).Height = 30
        DataGridView2.Rows.Add()
        DataGridView2.Rows(2).Height = 30
        DataGridView2.Rows.Add()
        DataGridView2.Rows(3).Height = 30
        DataGridView2.Rows(4).Height = 30



        Try






            primerdiames = Format(PrimerDiaDelMes(mifecha.Now).AddMonths(-restameses), "yyy-MM-dd")

            Dim ultimodiames As String = Format(UltimoDiaDelMes(primerdiames), "dd")





            'PRIMER SEMANA
            For i As Integer = 0 To 6
                primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(i), "yyy-MM-dd")
                DataGridView2.Rows.Item(0).Cells(i).Value = primerdiames

                'CAMBIA EL COLOR SI TRABAJA ESE DIA
                Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersonal & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                If encontroagenda <> 0 Then
                    DataGridView2.Rows.Item(0).Cells(i).Style.BackColor = System.Drawing.Color.Green
                End If

            Next


            'SEGUNDA SEMANA
            For i As Integer = 0 To 6
                primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(7 + i), "yyy-MM-dd")
                DataGridView2.Rows.Item(1).Cells(i).Value = primerdiames

                'CAMBIA EL COLOR SI TRABAJA ESE DIA
                Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersonal & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                If encontroagenda <> 0 Then
                    DataGridView2.Rows.Item(1).Cells(i).Style.BackColor = System.Drawing.Color.Green
                End If

            Next


            'TERCER SEMANA
            For j As Integer = 0 To 6
                primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(14 + j), "yyy-MM-dd")
                DataGridView2.Rows.Item(2).Cells(j).Value = primerdiames

                'CAMBIA EL COLOR SI TRABAJA ESE DIA
                Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersonal & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                If encontroagenda <> 0 Then
                    DataGridView2.Rows.Item(2).Cells(j).Style.BackColor = System.Drawing.Color.Green
                End If

            Next



            'CUARTA SEMANA
            For k As Integer = 0 To 6
                primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(21 + k), "yyy-MM-dd")
                DataGridView2.Rows.Item(3).Cells(k).Value = primerdiames

                'CAMBIA EL COLOR SI TRABAJA ESE DIA
                Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersonal & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                If encontroagenda <> 0 Then
                    DataGridView2.Rows.Item(3).Cells(k).Style.BackColor = System.Drawing.Color.Green
                End If

            Next








            If ultimodiames = 28 Then


            Else



                'QUINTA SEMANA
                If ultimodiames = 31 Then

                    For k As Integer = 0 To 2
                        primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(28 + k), "yyy-MM-dd")
                        DataGridView2.Rows.Item(4).Cells(k).Value = primerdiames
                        'CAMBIA EL COLOR SI TRABAJA ESE DIA
                        Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersonal & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                        If encontroagenda <> 0 Then
                            DataGridView2.Rows.Item(4).Cells(k).Style.BackColor = System.Drawing.Color.Green
                        End If

                    Next
                End If



                'QUINTA SEMANA
                If ultimodiames = 30 Then

                    For k As Integer = 0 To 1
                        primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(28 + k), "yyy-MM-dd")
                        DataGridView2.Rows.Item(4).Cells(k).Value = primerdiames
                        'CAMBIA EL COLOR SI TRABAJA ESE DIA
                        Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersonal & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                        If encontroagenda <> 0 Then
                            DataGridView2.Rows.Item(4).Cells(k).Style.BackColor = System.Drawing.Color.Green
                        End If

                    Next
                End If




                'QUINTA SEMANA
                If ultimodiames = 29 Then

                    For k As Integer = 0 To 0
                        primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(28 + k), "yyy-MM-dd")
                        DataGridView2.Rows.Item(4).Cells(k).Value = primerdiames

                        'CAMBIA EL COLOR SI TRABAJA ESE DIA
                        Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersonal & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                        If encontroagenda <> 0 Then
                            DataGridView2.Rows.Item(4).Cells(k).Style.BackColor = System.Drawing.Color.Green
                        End If

                    Next
                End If

            End If









            DataGridView2.ClearSelection()


        Catch ex As Exception

        End Try


    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click

        restameses = 0
        sumameses = sumameses + 1

        DataGridView2.Rows.Clear()
        DataGridView2.Rows.Add()
        DataGridView2.Rows(0).Height = 30

        DataGridView2.Rows.Add()
        DataGridView2.Rows(1).Height = 30
        DataGridView2.Rows.Add()
        DataGridView2.Rows(2).Height = 30
        DataGridView2.Rows.Add()
        DataGridView2.Rows(3).Height = 30
        DataGridView2.Rows(4).Height = 30


        Try




            primerdiames = Format(PrimerDiaDelMes(mifecha.Now).AddMonths(sumameses), "yyy-MM-dd")

            Dim ultimodiames As String = Format(UltimoDiaDelMes(primerdiames), "dd")





            'PRIMER SEMANA
            For i As Integer = 0 To 6
                primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(i), "yyy-MM-dd")
                DataGridView2.Rows.Item(0).Cells(i).Value = primerdiames


                'CAMBIA EL COLOR SI TRABAJA ESE DIA
                Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersonal & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                If encontroagenda <> 0 Then
                    DataGridView2.Rows.Item(0).Cells(i).Style.BackColor = System.Drawing.Color.Green
                End If

            Next


            'SEGUNDA SEMANA
            For i As Integer = 0 To 6
                primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(7 + i), "yyy-MM-dd")
                DataGridView2.Rows.Item(1).Cells(i).Value = primerdiames
                'CAMBIA EL COLOR SI TRABAJA ESE DIA
                Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersonal & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                If encontroagenda <> 0 Then
                    DataGridView2.Rows.Item(1).Cells(i).Style.BackColor = System.Drawing.Color.Green
                End If

            Next


            'TERCER SEMANA
            For j As Integer = 0 To 6
                primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(14 + j), "yyy-MM-dd")
                DataGridView2.Rows.Item(2).Cells(j).Value = primerdiames
                'CAMBIA EL COLOR SI TRABAJA ESE DIA
                Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersonal & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                If encontroagenda <> 0 Then
                    DataGridView2.Rows.Item(2).Cells(j).Style.BackColor = System.Drawing.Color.Green
                End If

            Next



            'CUARTA SEMANA
            For k As Integer = 0 To 6
                primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(21 + k), "yyy-MM-dd")
                DataGridView2.Rows.Item(3).Cells(k).Value = primerdiames
                'CAMBIA EL COLOR SI TRABAJA ESE DIA
                Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersonal & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                If encontroagenda <> 0 Then
                    DataGridView2.Rows.Item(3).Cells(k).Style.BackColor = System.Drawing.Color.Green
                End If

            Next







            If ultimodiames = 28 Then


            Else



                'QUINTA SEMANA
                If ultimodiames = 31 Then

                    For k As Integer = 0 To 2
                        primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(28 + k), "yyy-MM-dd")
                        DataGridView2.Rows.Item(4).Cells(k).Value = primerdiames
                        'CAMBIA EL COLOR SI TRABAJA ESE DIA
                        Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersonal & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                        If encontroagenda <> 0 Then
                            DataGridView2.Rows.Item(4).Cells(k).Style.BackColor = System.Drawing.Color.Green
                        End If

                    Next
                End If



                'QUINTA SEMANA
                If ultimodiames = 30 Then

                    For k As Integer = 0 To 1
                        primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(28 + k), "yyy-MM-dd")
                        DataGridView2.Rows.Item(4).Cells(k).Value = primerdiames
                        'CAMBIA EL COLOR SI TRABAJA ESE DIA
                        Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersonal & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                        If encontroagenda <> 0 Then
                            DataGridView2.Rows.Item(4).Cells(k).Style.BackColor = System.Drawing.Color.Green
                        End If

                    Next
                End If




                'QUINTA SEMANA
                If ultimodiames = 29 Then

                    For k As Integer = 0 To 0
                        primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(28 + k), "yyy-MM-dd")
                        DataGridView2.Rows.Item(4).Cells(k).Value = primerdiames

                        'CAMBIA EL COLOR SI TRABAJA ESE DIA
                        Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersonal & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                        If encontroagenda <> 0 Then
                            DataGridView2.Rows.Item(4).Cells(k).Style.BackColor = System.Drawing.Color.Green
                        End If

                    Next
                End If

            End If













            DataGridView2.ClearSelection()



        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click



        If IdReserva = 0 Then

            MsgBox("No se selecciono agenda")

        Else



            Dim result As DialogResult = MessageBox.Show("Esta seguro de marcar pendiente la agenda?",
                              "Modificacion de agenda",
                              MessageBoxButtons.YesNo)

            If result = DialogResult.Yes Then

                Dim fechahoraactual = fechaactual.Year & "-" & fechaactual.Month & "-" & fechaactual.Day & " " & horaactual

                updateSQL(" UPDATE reserva set FehaHraVisita = NULL , IdAgenda = 1  , FechaModificacion = '" & fechahoraactual & "' , EstadoReserva='Pendiente' WHERE IdReserva = " & IdReserva & " ; ")



                Dim fecha As String = fechaactual.Year & "-" & fechaactual.Month & "-" & fechaactual.Day



                If FormLogin.tipoacceso = "Agente" Then
                    DataFormPrincipalConsultaragendaDatos.DataSource = consultaSQL("select EstadoReserva , FehaHraVisita as Fecha , concat(concat(Nombre , (' '))  , Apellido ) as Nombre_Apellido , Tel as Telefono , Cel as Celular , Email , Barrios as Barrio , inmueble.Calle , inmueble.NroPuerta as Numero_Puerta, inmueble.NroApto as Numero_Apartamento , reserva.IdReserva from agenda ,  reserva , personas , inmueble , barrios  where  agenda.IdAgenda = reserva.IdAgenda and reserva.IdInteresado = personas.IdPersona and reserva.IdInmueble = inmueble.IdInmueble and barrios.IdBarrio = inmueble.IdBarrio
                and reserva.FehaHraVisita like '%" & fecha & "%' and  agenda.IdAgente = " & IdPersonal & "
                AND EstadoReserva = 'Activo'
                union all select EstadoReserva ,  FehaHraVisita as Fecha , concat(concat(Nombre , (' '))  , Apellido ) as Nombre_Apellido , Tel as Telefono , Cel as Celular , Email , Barrios as Barrio , inmueble.Calle , inmueble.NroPuerta as Numero_Puerta, inmueble.NroApto as Numero_Apartamento , reserva.IdReserva from agenda ,  reserva , personas , inmueble , barrios  where  agenda.IdAgenda = reserva.IdAgenda and reserva.IdInteresado = personas.IdPersona and reserva.IdInmueble = inmueble.IdInmueble and barrios.IdBarrio = inmueble.IdBarrio
                and EstadoReserva = 'Pendiente' and reserva.IdAgenda = 1
                ;").Tables("consultainmueble")


                Else

                    Dim IdPersonal As String = consultaSQL("SELECT IdPersona FROM enigma.personas where  concat(concat(Nombre , (' '))  , Apellido ) = '" & BoxFormPrincipalConsultaAgendaAgente.Text & "';").Tables("personas").Rows(0)("IdPersona").ToString()



                    'MUESTRA LAS AGENDAS PARA EL DIA DE HOY
                    DataFormPrincipalConsultaragendaDatos.DataSource = consultaSQL("select EstadoReserva , FehaHraVisita as Fecha , concat(concat(Nombre , (' '))  , Apellido ) as Nombre_Apellido , Tel as Telefono , Cel as Celular , Email , Barrios as Barrio , inmueble.Calle , inmueble.NroPuerta as Numero_Puerta, inmueble.NroApto as Numero_Apartamento , reserva.IdReserva from agenda ,  reserva , personas , inmueble , barrios  where  agenda.IdAgenda = reserva.IdAgenda and reserva.IdInteresado = personas.IdPersona and reserva.IdInmueble = inmueble.IdInmueble and barrios.IdBarrio = inmueble.IdBarrio
                and reserva.FehaHraVisita like '%" & fecha & "%' and  agenda.IdAgente = " & IdPersonal & "
                AND EstadoReserva = 'Activo'
                union all select EstadoReserva ,  FehaHraVisita as Fecha , concat(concat(Nombre , (' '))  , Apellido ) as Nombre_Apellido , Tel as Telefono , Cel as Celular , Email , Barrios as Barrio , inmueble.Calle , inmueble.NroPuerta as Numero_Puerta, inmueble.NroApto as Numero_Apartamento , reserva.IdReserva from agenda ,  reserva , personas , inmueble , barrios  where  agenda.IdAgenda = reserva.IdAgenda and reserva.IdInteresado = personas.IdPersona and reserva.IdInmueble = inmueble.IdInmueble and barrios.IdBarrio = inmueble.IdBarrio
                and EstadoReserva = 'Pendiente' and reserva.IdAgenda = 1
                ;").Tables("consultainmueble")



                End If






            End If

            DataFormPrincipalConsultaragendaDatos.Columns(10).Visible = False


        End If


    End Sub

    Private Sub Button8_Click_4(sender As Object, e As EventArgs) Handles Button8.Click


        restameses = 0
        sumameses = 0

        'CARGA LOS NOMBRES DE LOS AGENTES
        ComboBox1.DataSource = (consultaSQL("select concat(concat(Nombre , (' ')) , Apellido ) from personas where Funcion = 'Agente';").Tables(0))
        ComboBox1.DisplayMember = "concat(concat(Nombre , (' ')) , Apellido )"



        Try



            Dim IdPersona As String = consultaSQL("SELECT IdPersona FROM enigma.personas where  concat(concat(Nombre , (' '))  , Apellido ) = '" & ComboBox1.Text & "';").Tables("personas").Rows(0)("IdPersona").ToString()

            Dim ultimodiames As String = Format(UltimoDiaDelMes(mifecha.Now), "dd")

            DataGridView1.Rows.Clear()
            DataGridView1.Rows.Add()
            DataGridView1.Rows(0).Height = 30

            DataGridView1.Rows.Add()
            DataGridView1.Rows(1).Height = 30
            DataGridView1.Rows.Add()
            DataGridView1.Rows(2).Height = 30
            DataGridView1.Rows.Add()
            DataGridView1.Rows(3).Height = 30
            DataGridView1.Rows(4).Height = 30



            'PRIMER SEMANA
            For i As Integer = 0 To 6

                primerdiames = Format(PrimerDiaDelMes(mifecha.Now).AddDays(i), "yyy-MM-dd")
                DataGridView1.Rows.Item(0).Cells(i).Value = primerdiames

                'CAMBIA EL COLOR SI TRABAJA ESE DIA
                Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                If encontroagenda <> 0 Then
                    DataGridView1.Rows.Item(0).Cells(i).Style.BackColor = System.Drawing.Color.Green
                End If
            Next






            For i As Integer = 0 To 6



                primerdiames = Format(PrimerDiaDelMes(mifecha.Now).AddDays(7 + i), "yyy-MM-dd")

                'LLENA LAS CELDAS CON LAS FECHAS
                DataGridView1.Rows.Item(1).Cells(i).Value = primerdiames




                Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                If encontroagenda <> 0 Then

                    'CAMBIA EL COLOR
                    DataGridView1.Rows.Item(1).Cells(i).Style.BackColor = System.Drawing.Color.Green


                End If





            Next



            For j As Integer = 0 To 6


                primerdiames = Format(PrimerDiaDelMes(mifecha.Now).AddDays(14 + j), "yyy-MM-dd")


                DataGridView1.Rows.Item(2).Cells(j).Value = primerdiames



                Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                If encontroagenda <> 0 Then


                    'CAMBIA EL COLOR
                    DataGridView1.Rows.Item(2).Cells(j).Style.BackColor = System.Drawing.Color.Green


                End If

            Next






            For k As Integer = 0 To 6

                primerdiames = Format(PrimerDiaDelMes(mifecha.Now).AddDays(21 + k), "yyy-MM-dd")

                DataGridView1.Rows.Item(3).Cells(k).Value = primerdiames


                Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                If encontroagenda <> 0 Then


                    'CAMBIA EL COLOR
                    DataGridView1.Rows.Item(3).Cells(k).Style.BackColor = System.Drawing.Color.Green


                End If




            Next





            If ultimodiames = 28 Then


            Else



                'QUINTA SEMANA
                If ultimodiames = 31 Then

                    For k As Integer = 0 To 2
                        primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(28 + k), "yyy-MM-dd")
                        DataGridView1.Rows.Item(4).Cells(k).Value = primerdiames
                        'CAMBIA EL COLOR SI TRABAJA ESE DIA
                        Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                        If encontroagenda <> 0 Then
                            DataGridView1.Rows.Item(4).Cells(k).Style.BackColor = System.Drawing.Color.Green
                        End If

                    Next
                End If



                'QUINTA SEMANA
                If ultimodiames = 30 Then

                    For k As Integer = 0 To 1
                        primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(28 + k), "yyy-MM-dd")
                        DataGridView1.Rows.Item(4).Cells(k).Value = primerdiames
                        'CAMBIA EL COLOR SI TRABAJA ESE DIA
                        Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                        If encontroagenda <> 0 Then
                            DataGridView1.Rows.Item(4).Cells(k).Style.BackColor = System.Drawing.Color.Green
                        End If

                    Next
                End If




                'QUINTA SEMANA
                If ultimodiames = 29 Then

                    For k As Integer = 0 To 0
                        primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(28 + k), "yyy-MM-dd")
                        DataGridView1.Rows.Item(4).Cells(k).Value = primerdiames

                        'CAMBIA EL COLOR SI TRABAJA ESE DIA
                        Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                        If encontroagenda <> 0 Then
                            DataGridView1.Rows.Item(4).Cells(k).Style.BackColor = System.Drawing.Color.Green
                        End If

                    Next
                End If

            End If












        Catch ex As Exception

        End Try


        DataGridView1.ClearSelection()
    End Sub

    Private Sub BtnHorariosMesSiguiente_Click(sender As Object, e As EventArgs) Handles BtnHorariosMesSiguiente.Click
        restameses = 0
        sumameses = sumameses + 1

        DataGridView1.Rows.Clear()
        DataGridView1.Rows.Add()
        DataGridView1.Rows(0).Height = 30

        DataGridView1.Rows.Add()
        DataGridView1.Rows(1).Height = 30
        DataGridView1.Rows.Add()
        DataGridView1.Rows(2).Height = 30
        DataGridView1.Rows.Add()
        DataGridView1.Rows(3).Height = 30
        DataGridView1.Rows(4).Height = 30


        Try



            Dim IdPersona As String = consultaSQL("SELECT IdPersona FROM enigma.personas where  concat(concat(Nombre , (' '))  , Apellido ) = '" & ComboBox1.Text & "';").Tables("personas").Rows(0)("IdPersona").ToString()


            primerdiames = Format(PrimerDiaDelMes(mifecha.Now).AddMonths(sumameses), "yyy-MM-dd")

            Dim ultimodiames As String = Format(UltimoDiaDelMes(primerdiames), "dd")





            'PRIMER SEMANA
            For i As Integer = 0 To 6
                primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(i), "yyy-MM-dd")
                DataGridView1.Rows.Item(0).Cells(i).Value = primerdiames


                'CAMBIA EL COLOR SI TRABAJA ESE DIA
                Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                If encontroagenda <> 0 Then
                    DataGridView1.Rows.Item(0).Cells(i).Style.BackColor = System.Drawing.Color.Green
                End If

            Next


            'SEGUNDA SEMANA
            For i As Integer = 0 To 6
                primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(7 + i), "yyy-MM-dd")
                DataGridView1.Rows.Item(1).Cells(i).Value = primerdiames
                'CAMBIA EL COLOR SI TRABAJA ESE DIA
                Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                If encontroagenda <> 0 Then
                    DataGridView1.Rows.Item(1).Cells(i).Style.BackColor = System.Drawing.Color.Green
                End If

            Next


            'TERCER SEMANA
            For j As Integer = 0 To 6
                primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(14 + j), "yyy-MM-dd")
                DataGridView1.Rows.Item(2).Cells(j).Value = primerdiames
                'CAMBIA EL COLOR SI TRABAJA ESE DIA
                Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                If encontroagenda <> 0 Then
                    DataGridView1.Rows.Item(2).Cells(j).Style.BackColor = System.Drawing.Color.Green
                End If

            Next



            'CUARTA SEMANA
            For k As Integer = 0 To 6
                primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(21 + k), "yyy-MM-dd")
                DataGridView1.Rows.Item(3).Cells(k).Value = primerdiames
                'CAMBIA EL COLOR SI TRABAJA ESE DIA
                Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                If encontroagenda <> 0 Then
                    DataGridView1.Rows.Item(3).Cells(k).Style.BackColor = System.Drawing.Color.Green
                End If

            Next






            If ultimodiames = 28 Then


            Else



                'QUINTA SEMANA
                If ultimodiames = 31 Then

                    For k As Integer = 0 To 2
                        primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(28 + k), "yyy-MM-dd")
                        DataGridView1.Rows.Item(4).Cells(k).Value = primerdiames
                        'CAMBIA EL COLOR SI TRABAJA ESE DIA
                        Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                        If encontroagenda <> 0 Then
                            DataGridView1.Rows.Item(4).Cells(k).Style.BackColor = System.Drawing.Color.Green
                        End If

                    Next
                End If



                'QUINTA SEMANA
                If ultimodiames = 30 Then

                    For k As Integer = 0 To 1
                        primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(28 + k), "yyy-MM-dd")
                        DataGridView1.Rows.Item(4).Cells(k).Value = primerdiames
                        'CAMBIA EL COLOR SI TRABAJA ESE DIA
                        Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                        If encontroagenda <> 0 Then
                            DataGridView1.Rows.Item(4).Cells(k).Style.BackColor = System.Drawing.Color.Green
                        End If

                    Next
                End If




                'QUINTA SEMANA
                If ultimodiames = 29 Then

                    For k As Integer = 0 To 0
                        primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(28 + k), "yyy-MM-dd")
                        DataGridView1.Rows.Item(4).Cells(k).Value = primerdiames

                        'CAMBIA EL COLOR SI TRABAJA ESE DIA
                        Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                        If encontroagenda <> 0 Then
                            DataGridView1.Rows.Item(4).Cells(k).Style.BackColor = System.Drawing.Color.Green
                        End If

                    Next
                End If

            End If





            DataGridView1.ClearSelection()



        Catch ex As Exception

        End Try


    End Sub

    Private Sub BtnHorariosMesAnterior_Click(sender As Object, e As EventArgs) Handles BtnHorariosMesAnterior.Click
        sumameses = 0

        restameses = restameses + 1


        DataGridView1.Rows.Clear()
        DataGridView1.Rows.Add()
        DataGridView1.Rows(0).Height = 30

        DataGridView1.Rows.Add()
        DataGridView1.Rows(1).Height = 30
        DataGridView1.Rows.Add()
        DataGridView1.Rows(2).Height = 30
        DataGridView1.Rows.Add()
        DataGridView1.Rows(3).Height = 30
        DataGridView1.Rows(4).Height = 30



        Try



            Dim IdPersona As String = consultaSQL("SELECT IdPersona FROM enigma.personas where  concat(concat(Nombre , (' '))  , Apellido ) = '" & ComboBox1.Text & "';").Tables("personas").Rows(0)("IdPersona").ToString()



            primerdiames = Format(PrimerDiaDelMes(mifecha.Now).AddMonths(-restameses), "yyy-MM-dd")

            Dim ultimodiames As String = Format(UltimoDiaDelMes(primerdiames), "dd")





            'PRIMER SEMANA
            For i As Integer = 0 To 6
                primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(i), "yyy-MM-dd")
                DataGridView1.Rows.Item(0).Cells(i).Value = primerdiames

                'CAMBIA EL COLOR SI TRABAJA ESE DIA
                Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                If encontroagenda <> 0 Then
                    DataGridView1.Rows.Item(0).Cells(i).Style.BackColor = System.Drawing.Color.Green
                End If

            Next


            'SEGUNDA SEMANA
            For i As Integer = 0 To 6
                primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(7 + i), "yyy-MM-dd")
                DataGridView1.Rows.Item(1).Cells(i).Value = primerdiames

                'CAMBIA EL COLOR SI TRABAJA ESE DIA
                Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                If encontroagenda <> 0 Then
                    DataGridView1.Rows.Item(1).Cells(i).Style.BackColor = System.Drawing.Color.Green
                End If

            Next


            'TERCER SEMANA
            For j As Integer = 0 To 6
                primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(14 + j), "yyy-MM-dd")
                DataGridView1.Rows.Item(2).Cells(j).Value = primerdiames

                'CAMBIA EL COLOR SI TRABAJA ESE DIA
                Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                If encontroagenda <> 0 Then
                    DataGridView1.Rows.Item(2).Cells(j).Style.BackColor = System.Drawing.Color.Green
                End If

            Next



            'CUARTA SEMANA
            For k As Integer = 0 To 6
                primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(21 + k), "yyy-MM-dd")
                DataGridView1.Rows.Item(3).Cells(k).Value = primerdiames

                'CAMBIA EL COLOR SI TRABAJA ESE DIA
                Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                If encontroagenda <> 0 Then
                    DataGridView1.Rows.Item(3).Cells(k).Style.BackColor = System.Drawing.Color.Green
                End If

            Next







            If ultimodiames = 28 Then


            Else



                'QUINTA SEMANA
                If ultimodiames = 31 Then

                    For k As Integer = 0 To 2
                        primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(28 + k), "yyy-MM-dd")
                        DataGridView1.Rows.Item(4).Cells(k).Value = primerdiames
                        'CAMBIA EL COLOR SI TRABAJA ESE DIA
                        Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                        If encontroagenda <> 0 Then
                            DataGridView1.Rows.Item(4).Cells(k).Style.BackColor = System.Drawing.Color.Green
                        End If

                    Next
                End If



                'QUINTA SEMANA
                If ultimodiames = 30 Then

                    For k As Integer = 0 To 1
                        primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(28 + k), "yyy-MM-dd")
                        DataGridView1.Rows.Item(4).Cells(k).Value = primerdiames
                        'CAMBIA EL COLOR SI TRABAJA ESE DIA
                        Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                        If encontroagenda <> 0 Then
                            DataGridView1.Rows.Item(4).Cells(k).Style.BackColor = System.Drawing.Color.Green
                        End If

                    Next
                End If




                'QUINTA SEMANA
                If ultimodiames = 29 Then

                    For k As Integer = 0 To 0
                        primerdiames = Format(PrimerDiaDelMes(primerdiames).AddDays(28 + k), "yyy-MM-dd")
                        DataGridView1.Rows.Item(4).Cells(k).Value = primerdiames

                        'CAMBIA EL COLOR SI TRABAJA ESE DIA
                        Dim encontroagenda As String = consultaSQL("SELECT count(FechaAgente) FROM agenda where FechaAgente like '" & primerdiames & "' and IdAgente = " & IdPersona & ";").Tables("quebarrioshay").Rows(0)("count(FechaAgente)").ToString()
                        If encontroagenda <> 0 Then
                            DataGridView1.Rows.Item(4).Cells(k).Style.BackColor = System.Drawing.Color.Green
                        End If

                    Next
                End If

            End If





            DataGridView1.ClearSelection()


        Catch ex As Exception

        End Try


    End Sub

End Class