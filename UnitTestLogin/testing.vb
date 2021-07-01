Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports inmobiliaria
<TestClass()> Public Class Login
    Dim FormLogin As New FormLogin


    'TESTING CONEXION USUARIO GENERICO
    <TestMethod()> Public Sub TestLoginBdGenerico()
        FormLogin.tipoacceso = "Generico"
        Assert.IsTrue(FormLogin.acceso("Generico", "Generico"))
    End Sub

    'TESTING CONEXION USUARIO ADMINISTRADOR
    <TestMethod()> Public Sub TestLoginBdAdministrador()
        FormLogin.tipoacceso = "Administrador"
        Assert.IsTrue(FormLogin.acceso("MartinHegui", "48573734"))
    End Sub

    'TESTING CONEXION USUARIO GERENTE
    <TestMethod()> Public Sub TestLoginBdGerente()
        FormLogin.tipoacceso = "Gerente"
        Assert.IsTrue(FormLogin.acceso("PedroRodriguez", "33638294"))
    End Sub

    'TESTING CONEXION USUARIO AGENTE
    <TestMethod()> Public Sub TestLoginBdAgente()
        FormLogin.tipoacceso = "Agente"
        Assert.IsTrue(FormLogin.acceso("RobertoFagundez", "47384493"))
    End Sub



End Class

<TestClass()> Public Class AltaUsuario
    Dim FormLogin As New FormLogin
    Dim FormPrincipal As New FormPrincipal

    <TestMethod()> Public Sub TestRegistroClienteConUsuarioAgente()


        FormPrincipal.Show()
        FormPrincipal.BtnFormPrincipalRegistrarInteresado.PerformClick()

        FormPrincipal.ComboFormPrincipalRegistrarUsuarioTipoDocumento.Text = "CI"
        FormPrincipal.TxtFormPrincipalRegistrarUsuarioDocumento.Text = "999999999"
        FormPrincipal.TxtFormPrincipalRegistrarUsuarioNombre.Text = "NOMBRETESTING"
        FormPrincipal.TxtFormPrincipalRegistrarUsuarioApellido.Text = "APELLIDOTESTING"
        FormPrincipal.DateFormPrincipalRegistroUsuarioFechaNacimiento.Value = "20/7/2019"
        FormPrincipal.NumericFormPrincipalRegistrarUsuarioTelefono.Value = 999
        FormPrincipal.NumericFormPrincipalRegistrarUsuarioCelular.Value = 999
        FormPrincipal.TxtFormPrincipalRegistrarUsuarioEmail.Text = "testing@gmail.com"
        FormPrincipal.ComboFormPrincipalRegistroUsuarioDepartamento.Text = "Montevideo"
        FormPrincipal.ComboFormPrincipalRegistroUsuarioCiudad.Text = "Montevideo"
        FormPrincipal.ComboFormPrincipalRegistroUsuarioBarrio.Text = "Centro"
        FormPrincipal.TxtFormPrincipalRegistrarUsuarioCalle.Text = "CALLETESTING"
        FormPrincipal.TxtFormPrincipalRegistrarUsuarioEsquina.Text = "ESQUINATESTING"
        FormPrincipal.NumericFormPrincipalRegistrarUsuarioNumeroPuerta.Value = 999
        FormPrincipal.NumericFormPrincipalRegistrarUsuarioNumeroApartamento.Value = 999
        FormPrincipal.TxtFormPrincipalRegistrarUsuarioNombreInmueble.Text = "NOMBREINMUEBLETESTING"
        FormPrincipal.NumericFormPrincipalRegistrarUsuarioMontosueldo.Value = 999
        FormPrincipal.ComboFormPrincipalRegistroUsuarioTipoPersona.Text = "Vendedor"

        FormPrincipal.ButtonFormPrincipalRegistrarUsuarioIngresar.PerformClick()


        Assert.IsTrue(FormPrincipal.testingvalidaringresousuario)


    End Sub





End Class
