'       \          ENIGMA           /           
'        \                         /           
'         \    This page does     /             
'          ]   not exist yet.    [             
'          ]                     [             
'          ]___               ___[    
'          ]  ]\             /[  [          
'          ]  ] \           / [ -[    
'          ]  ]  ]         [  [ -[    
'          ]  ]  ]__     __[  [ -[    
'          ]  ]  ] ]\   /[-[  [ -[    
'          ]  ]  ]-]     [ [  [  [ 
'          ]  ]  ]_]     [_[  [  [
'          ]  ]  ]         [  [  [
'          ]  ] /           \ [ -[
'          ]__]/             \[__[
'          ]                     [
'          ]                     [
'          ]                     [
'         /                       \
'        /                         \
'       /                           \
'      /                             \
'     /                               \
'    /                                 \
'   /                                   \
'  /                                     \
' /                                       \  

Imports System.Runtime.InteropServices
Imports System.Text.RegularExpressions

Public Class FormLogin
    Public tipoacceso As String = ""
    Dim log As Boolean


    Public Function acceso(ByVal usuario As String, ByVal contraseña As String)

        If existeusuario(usuario, contraseña, "Personal") = True Then


            If tipoacceso = "Administrador" Or tipoacceso = "Gerente" Or tipoacceso = "Agente" Or tipoacceso = "AgenteFijo" Then
                log = True

                Me.Hide()
                FormPrincipal.Show()

            ElseIf tipoacceso = "Generico" Then
                log = True

                BusquedasPropiedadesInteresados.Show()
                Me.Hide()







            End If
        Else

            log = False

        End If

        Return log
    End Function


    'BOTON INICIAR
    Private Sub BtnIniciar_Click(sender As Object, e As EventArgs) Handles BtnFormLoginIniciar.Click

        'Primero valida si el mail es correcto para lueo ingresar en bd
        Dim expresionmail As String = "^[^=|*|$|>|;|'|""|%|-]"

        Dim r As New Regex(expresionmail)
        If r.IsMatch(TxtFormLoginUsuario.Text) = False Or r.IsMatch(TxtFormLoginContraseña.Text) = False Then

            MsgBox("Caracteres no validos")

        Else


            acceso(TxtFormLoginUsuario.Text, TxtFormLoginContraseña.Text)

        End If





    End Sub










    'BONTONNNNNNNNNNNNNNNNNNNN CAMBIAR CONTRASEÑA
    Private Sub FormLoginBtnCambiar_Click(sender As Object, e As EventArgs) Handles FormLoginBtnCambiar.Click


        cambiarcontraseña("cambiarcontraseña")


    End Sub








    'LABEL CAMBIAR CONTRASEÑA LABELLLLLLLL
    Private Sub Cambiarcontraseña_Click(sender As Object, e As EventArgs) Handles LabelFormLoginCambiarcontraseña.Click

        If existeusuario(TxtFormLoginUsuario.Text, TxtFormLoginContraseña.Text, "Personal") = True Then
            cambiarcontraseña("abrir")
        End If


    End Sub







    'CAMBIAR CONTRASEÑA
    Private Sub cambiarcontraseña(ByVal opcion As String)


        If opcion = "abrir" Then
            Dim cambiarpaswLogin As New FormLogin
            With cambiarpaswLogin
                .LabelFormLoginUsuario.Text = "Usuario"
                .LabelFormLoginContraseña.Text = "Nueva contraseña"
                .TxtFormLoginUsuario.Text = ""
                .TxtFormLoginContraseña.Text = ""
                .LabelFormLoginCambiarcontraseña.Visible = False
                .BtnFormLoginIniciar.Visible = False
                .FormLoginBtnCambiar.Visible = True
                .PickFormLoginCerrar.Visible = False
                .PickFormLoginMinimizar.Visible = False
            End With
            cambiarpaswLogin.Show()
        End If





        If opcion = "cambiarcontraseña" Then

            updateSQL("UPDATE usuarios SET Contraseña = '" & generarClaveSHA1(TxtFormLoginContraseña.Text) & "' WHERE Usuario = '" & TxtFormLoginUsuario.Text & "'")
            MsgBox("Contraseña actualizada")
            Application.Restart()


        End If









    End Sub


















    '----------------------------------COMPORTAMIENTOS OBJETOS--------------------------------------------


    'PARA MOSTRAR CONTRASEÑA Y OCULTARLA 
    Private Sub CheckMostrarContraseña_CheckedChanged(sender As Object, e As EventArgs) Handles CheckFormLoginMostrarContraseña.CheckedChanged
        If TxtFormLoginContraseña.UseSystemPasswordChar = True Then
            TxtFormLoginContraseña.UseSystemPasswordChar = False

        Else
            TxtFormLoginContraseña.UseSystemPasswordChar = False
            TxtFormLoginContraseña.UseSystemPasswordChar = True
        End If
    End Sub

    'SI EL RATON ENTRA EN EL TEXTBOX SE BORRA EL CONTENIDO
    Private Sub TxtUsuario_MouseEnter(sender As Object, e As EventArgs) Handles TxtFormLoginUsuario.MouseEnter

        If TxtFormLoginUsuario.Text = "Usuario" Then
            TxtFormLoginUsuario.Text = ""
        End If

    End Sub

    'SI EL RATON ENTRA EN EL TEXTBOX SE BORRA EL CONTENIDO
    Private Sub TxtContraseña_MouseEnter(sender As Object, e As EventArgs) Handles TxtFormLoginContraseña.MouseEnter
        If TxtFormLoginContraseña.Text = "Contraseña" Then
            TxtFormLoginContraseña.Text = ""
        End If
    End Sub

    'SI EL RATON SALE DEL TEXBOX MUESTRA EL CONTENIDO
    Private Sub TxtUsuario_MouseLeave(sender As Object, e As EventArgs) Handles TxtFormLoginUsuario.MouseLeave
        If TxtFormLoginUsuario.Text = "" Then
            TxtFormLoginUsuario.Text = "Usuario"
        Else
            If TxtFormLoginContraseña.Text <> "" And TxtFormLoginContraseña.Text <> "Contraseña" Then
                LabelFormLoginCambiarcontraseña.ForeColor = System.Drawing.Color.Orange
                LabelFormLoginCambiarcontraseña.Enabled = True
            Else
                LabelFormLoginCambiarcontraseña.ForeColor = System.Drawing.Color.White
                LabelFormLoginCambiarcontraseña.Enabled = False
            End If
        End If
    End Sub


    'SI EL RATON SALE DEL TEXBOX MUESTRA EL CONTENIDO
    Private Sub TxtContraseña_MouseLeave(sender As Object, e As EventArgs) Handles TxtFormLoginContraseña.MouseLeave
        If TxtFormLoginContraseña.Text = "" Then
            TxtFormLoginContraseña.Text = "Contraseña"
        Else
            If TxtFormLoginUsuario.Text <> "" And TxtFormLoginUsuario.Text <> "Usuario" Then
                LabelFormLoginCambiarcontraseña.ForeColor = System.Drawing.Color.Orange
                LabelFormLoginCambiarcontraseña.Enabled = True
            Else
                LabelFormLoginCambiarcontraseña.ForeColor = System.Drawing.Color.White
                LabelFormLoginCambiarcontraseña.Enabled = False
            End If
        End If
    End Sub


    'BOTON CERRAR
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PickFormLoginCerrar.Click
        Application.Exit()
    End Sub
    'BOTON MINIMIZAR
    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PickFormLoginMinimizar.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub


    'CODIGOS PARA MOVER EL FORMULARIO AL HACER CLICK EN FORMULARIO
    <DllImport("user32.DLL", EntryPoint:="ReleaseCapture")>
    Private Shared Sub ReleaseCapture()
    End Sub

    <DllImport("user32.DLL", EntryPoint:="SendMessage")>
    Private Shared Sub SendMessage(ByVal hWnd As System.IntPtr, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer)
    End Sub


    Private Sub Login_MouseMove(sender As Object, e As MouseEventArgs) Handles MyBase.MouseMove
        ReleaseCapture()
        SendMessage(Me.Handle, &H112&, &HF012&, 0)
    End Sub

    Private Sub TxtFormLoginUsuario_TextChanged(sender As Object, e As EventArgs) Handles TxtFormLoginUsuario.TextChanged

    End Sub

    Public Function login() As Boolean
        Throw New NotImplementedException()
    End Function
End Class
