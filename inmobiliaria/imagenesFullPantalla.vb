Public Class imagenesFullPantalla


    Private Sub PickFormPrincipalCerrar_Click(sender As Object, e As EventArgs) Handles PickFormPrincipalCerrar.Click

        Me.Close()


    End Sub

    Private Sub PicImagenesFullPantalla_MouseClick(sender As Object, e As MouseEventArgs) Handles PicImagenesFullPantalla.MouseClick

        Dim xtotal = PicImagenesFullPantalla.Size.Width
        Dim xtotalint As Integer = CInt(xtotal)
        Dim x As String = e.X
        Dim xint As Integer = CInt(x)


        If xint >= xtotal / 2 Then



            If Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(38, Byte), Integer), CType(CType(55, Byte), Integer), CType(CType(52, Byte), Integer)) Then


                FormPrincipal.BtnBuscarinmuebleimagenessiguiente.PerformClick()

                LabelSiguiente.Parent = PicImagenesFullPantalla

            Else


                BusquedasPropiedadesInteresados.ButtonlFormBusquedasPropiedadesPanelImagenesSiguiente.PerformClick()
                LabelSiguiente.Parent = PicImagenesFullPantalla

            End If





        Else



            If Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(38, Byte), Integer), CType(CType(55, Byte), Integer), CType(CType(52, Byte), Integer)) Then




                FormPrincipal.BtnBuscarinmuebleimagenesanterior.PerformClick()

                LabelSiguiente.Parent = PicImagenesFullPantalla

            Else




                BusquedasPropiedadesInteresados.ButtonlFormBusquedasPropiedadesPanelImagenesAnterior.PerformClick()
                LabelAnterior.Parent = PicImagenesFullPantalla

            End If






        End If


    End Sub

    Private Sub imagenesFullPantalla_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(38, Byte), Integer), CType(CType(55, Byte), Integer), CType(CType(52, Byte), Integer)) Then

            LabelSiguiente.Parent = PicImagenesFullPantalla
            LabelAnterior.Parent = PicImagenesFullPantalla


        Else

            LabelSiguiente.Parent = PicImagenesFullPantalla
            LabelAnterior.Parent = PicImagenesFullPantalla


            Me.BackgroundImage = Global.inmobiliaria.My.Resources.Resources.living


        End If



    End Sub

    Private Sub LabelSiguiente_Click(sender As Object, e As EventArgs) Handles LabelSiguiente.Click
        BusquedasPropiedadesInteresados.ButtonlFormBusquedasPropiedadesPanelImagenesSiguiente.PerformClick()
    End Sub

    Private Sub LabelAnterior_Click(sender As Object, e As EventArgs) Handles LabelAnterior.Click
        BusquedasPropiedadesInteresados.ButtonlFormBusquedasPropiedadesPanelImagenesAnterior.PerformClick()
    End Sub


End Class