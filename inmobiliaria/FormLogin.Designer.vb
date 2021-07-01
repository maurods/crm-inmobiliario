<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FormLogin
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.PanelFormLoginLateral = New System.Windows.Forms.Panel()
        Me.PickFormLoginLogin = New System.Windows.Forms.PictureBox()
        Me.TxtFormLoginUsuario = New System.Windows.Forms.TextBox()
        Me.TxtFormLoginContraseña = New System.Windows.Forms.TextBox()
        Me.LabelFormLoginUsuario = New System.Windows.Forms.Label()
        Me.LabelFormLoginContraseña = New System.Windows.Forms.Label()
        Me.BtnFormLoginIniciar = New System.Windows.Forms.Button()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.CheckFormLoginMostrarContraseña = New System.Windows.Forms.CheckBox()
        Me.PickFormLoginMinimizar = New System.Windows.Forms.PictureBox()
        Me.PickFormLoginCerrar = New System.Windows.Forms.PictureBox()
        Me.LabelFormLoginCambiarcontraseña = New System.Windows.Forms.Label()
        Me.FormLoginBtnCambiar = New System.Windows.Forms.Button()
        Me.PanelFormLoginLateral.SuspendLayout()
        CType(Me.PickFormLoginLogin, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PickFormLoginMinimizar, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PickFormLoginCerrar, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PanelFormLoginLateral
        '
        Me.PanelFormLoginLateral.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(122, Byte), Integer), CType(CType(204, Byte), Integer))
        Me.PanelFormLoginLateral.Controls.Add(Me.PickFormLoginLogin)
        Me.PanelFormLoginLateral.Dock = System.Windows.Forms.DockStyle.Left
        Me.PanelFormLoginLateral.Location = New System.Drawing.Point(0, 0)
        Me.PanelFormLoginLateral.Name = "PanelFormLoginLateral"
        Me.PanelFormLoginLateral.Size = New System.Drawing.Size(210, 312)
        Me.PanelFormLoginLateral.TabIndex = 0
        '
        'PickFormLoginLogin
        '
        Me.PickFormLoginLogin.Image = Global.inmobiliaria.My.Resources.Resources.imagenlogin1
        Me.PickFormLoginLogin.Location = New System.Drawing.Point(54, 87)
        Me.PickFormLoginLogin.Name = "PickFormLoginLogin"
        Me.PickFormLoginLogin.Size = New System.Drawing.Size(119, 107)
        Me.PickFormLoginLogin.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PickFormLoginLogin.TabIndex = 0
        Me.PickFormLoginLogin.TabStop = False
        '
        'TxtFormLoginUsuario
        '
        Me.TxtFormLoginUsuario.BackColor = System.Drawing.Color.FromArgb(CType(CType(38, Byte), Integer), CType(CType(55, Byte), Integer), CType(CType(52, Byte), Integer))
        Me.TxtFormLoginUsuario.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TxtFormLoginUsuario.ForeColor = System.Drawing.SystemColors.Info
        Me.TxtFormLoginUsuario.Location = New System.Drawing.Point(280, 76)
        Me.TxtFormLoginUsuario.Name = "TxtFormLoginUsuario"
        Me.TxtFormLoginUsuario.Size = New System.Drawing.Size(344, 13)
        Me.TxtFormLoginUsuario.TabIndex = 1
        Me.TxtFormLoginUsuario.Text = "Usuario"
        '
        'TxtFormLoginContraseña
        '
        Me.TxtFormLoginContraseña.BackColor = System.Drawing.Color.FromArgb(CType(CType(38, Byte), Integer), CType(CType(55, Byte), Integer), CType(CType(52, Byte), Integer))
        Me.TxtFormLoginContraseña.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TxtFormLoginContraseña.ForeColor = System.Drawing.SystemColors.Info
        Me.TxtFormLoginContraseña.Location = New System.Drawing.Point(280, 181)
        Me.TxtFormLoginContraseña.Name = "TxtFormLoginContraseña"
        Me.TxtFormLoginContraseña.Size = New System.Drawing.Size(344, 13)
        Me.TxtFormLoginContraseña.TabIndex = 2
        Me.TxtFormLoginContraseña.Text = "Contraseña"
        Me.TxtFormLoginContraseña.UseSystemPasswordChar = True
        '
        'LabelFormLoginUsuario
        '
        Me.LabelFormLoginUsuario.AutoSize = True
        Me.LabelFormLoginUsuario.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelFormLoginUsuario.ForeColor = System.Drawing.Color.White
        Me.LabelFormLoginUsuario.Location = New System.Drawing.Point(278, 41)
        Me.LabelFormLoginUsuario.Name = "LabelFormLoginUsuario"
        Me.LabelFormLoginUsuario.Size = New System.Drawing.Size(84, 20)
        Me.LabelFormLoginUsuario.TabIndex = 3
        Me.LabelFormLoginUsuario.Text = "USUARIO"
        '
        'LabelFormLoginContraseña
        '
        Me.LabelFormLoginContraseña.AutoSize = True
        Me.LabelFormLoginContraseña.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelFormLoginContraseña.ForeColor = System.Drawing.Color.White
        Me.LabelFormLoginContraseña.Location = New System.Drawing.Point(280, 142)
        Me.LabelFormLoginContraseña.Name = "LabelFormLoginContraseña"
        Me.LabelFormLoginContraseña.Size = New System.Drawing.Size(119, 20)
        Me.LabelFormLoginContraseña.TabIndex = 4
        Me.LabelFormLoginContraseña.Text = "CONTRASEÑA"
        '
        'BtnFormLoginIniciar
        '
        Me.BtnFormLoginIniciar.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(122, Byte), Integer), CType(CType(204, Byte), Integer))
        Me.BtnFormLoginIniciar.FlatAppearance.BorderSize = 0
        Me.BtnFormLoginIniciar.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnFormLoginIniciar.ForeColor = System.Drawing.Color.White
        Me.BtnFormLoginIniciar.Location = New System.Drawing.Point(294, 264)
        Me.BtnFormLoginIniciar.Name = "BtnFormLoginIniciar"
        Me.BtnFormLoginIniciar.Size = New System.Drawing.Size(344, 36)
        Me.BtnFormLoginIniciar.TabIndex = 7
        Me.BtnFormLoginIniciar.Text = "Iniciar"
        Me.BtnFormLoginIniciar.UseVisualStyleBackColor = False
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.Panel2.Location = New System.Drawing.Point(280, 98)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(344, 1)
        Me.Panel2.TabIndex = 10
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.Panel3.Location = New System.Drawing.Point(280, 207)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(344, 1)
        Me.Panel3.TabIndex = 11
        '
        'CheckFormLoginMostrarContraseña
        '
        Me.CheckFormLoginMostrarContraseña.AutoSize = True
        Me.CheckFormLoginMostrarContraseña.ForeColor = System.Drawing.Color.White
        Me.CheckFormLoginMostrarContraseña.Location = New System.Drawing.Point(280, 222)
        Me.CheckFormLoginMostrarContraseña.Name = "CheckFormLoginMostrarContraseña"
        Me.CheckFormLoginMostrarContraseña.Size = New System.Drawing.Size(118, 17)
        Me.CheckFormLoginMostrarContraseña.TabIndex = 12
        Me.CheckFormLoginMostrarContraseña.Text = "Mostrar Contraseña"
        Me.CheckFormLoginMostrarContraseña.UseVisualStyleBackColor = True
        '
        'PickFormLoginMinimizar
        '
        Me.PickFormLoginMinimizar.Image = Global.inmobiliaria.My.Resources.Resources.minimizar
        Me.PickFormLoginMinimizar.Location = New System.Drawing.Point(652, 5)
        Me.PickFormLoginMinimizar.Name = "PickFormLoginMinimizar"
        Me.PickFormLoginMinimizar.Size = New System.Drawing.Size(31, 20)
        Me.PickFormLoginMinimizar.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PickFormLoginMinimizar.TabIndex = 9
        Me.PickFormLoginMinimizar.TabStop = False
        '
        'PickFormLoginCerrar
        '
        Me.PickFormLoginCerrar.Image = Global.inmobiliaria.My.Resources.Resources.cerrar
        Me.PickFormLoginCerrar.Location = New System.Drawing.Point(684, 5)
        Me.PickFormLoginCerrar.Name = "PickFormLoginCerrar"
        Me.PickFormLoginCerrar.Size = New System.Drawing.Size(29, 20)
        Me.PickFormLoginCerrar.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PickFormLoginCerrar.TabIndex = 8
        Me.PickFormLoginCerrar.TabStop = False
        '
        'LabelFormLoginCambiarcontraseña
        '
        Me.LabelFormLoginCambiarcontraseña.AutoSize = True
        Me.LabelFormLoginCambiarcontraseña.Enabled = False
        Me.LabelFormLoginCambiarcontraseña.ForeColor = System.Drawing.Color.White
        Me.LabelFormLoginCambiarcontraseña.Location = New System.Drawing.Point(523, 224)
        Me.LabelFormLoginCambiarcontraseña.Name = "LabelFormLoginCambiarcontraseña"
        Me.LabelFormLoginCambiarcontraseña.Size = New System.Drawing.Size(101, 13)
        Me.LabelFormLoginCambiarcontraseña.TabIndex = 13
        Me.LabelFormLoginCambiarcontraseña.Text = "Cambiar contraseña"
        '
        'FormLoginBtnCambiar
        '
        Me.FormLoginBtnCambiar.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(122, Byte), Integer), CType(CType(204, Byte), Integer))
        Me.FormLoginBtnCambiar.FlatAppearance.BorderSize = 0
        Me.FormLoginBtnCambiar.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.FormLoginBtnCambiar.ForeColor = System.Drawing.Color.White
        Me.FormLoginBtnCambiar.Location = New System.Drawing.Point(294, 264)
        Me.FormLoginBtnCambiar.Name = "FormLoginBtnCambiar"
        Me.FormLoginBtnCambiar.Size = New System.Drawing.Size(344, 36)
        Me.FormLoginBtnCambiar.TabIndex = 14
        Me.FormLoginBtnCambiar.Text = "Cambiar"
        Me.FormLoginBtnCambiar.UseVisualStyleBackColor = False
        Me.FormLoginBtnCambiar.Visible = False
        '
        'FormLogin
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(38, Byte), Integer), CType(CType(55, Byte), Integer), CType(CType(52, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(726, 312)
        Me.Controls.Add(Me.LabelFormLoginCambiarcontraseña)
        Me.Controls.Add(Me.CheckFormLoginMostrarContraseña)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.PickFormLoginMinimizar)
        Me.Controls.Add(Me.PickFormLoginCerrar)
        Me.Controls.Add(Me.LabelFormLoginContraseña)
        Me.Controls.Add(Me.LabelFormLoginUsuario)
        Me.Controls.Add(Me.TxtFormLoginContraseña)
        Me.Controls.Add(Me.TxtFormLoginUsuario)
        Me.Controls.Add(Me.PanelFormLoginLateral)
        Me.Controls.Add(Me.BtnFormLoginIniciar)
        Me.Controls.Add(Me.FormLoginBtnCambiar)
        Me.ForeColor = System.Drawing.SystemColors.ControlText
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "FormLogin"
        Me.Opacity = 0.96R
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.PanelFormLoginLateral.ResumeLayout(False)
        CType(Me.PickFormLoginLogin, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PickFormLoginMinimizar, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PickFormLoginCerrar, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents PanelFormLoginLateral As Panel
    Friend WithEvents TxtFormLoginUsuario As TextBox
    Friend WithEvents TxtFormLoginContraseña As TextBox
    Friend WithEvents LabelFormLoginUsuario As Label
    Friend WithEvents LabelFormLoginContraseña As Label
    Friend WithEvents BtnFormLoginIniciar As Button
    Friend WithEvents PickFormLoginCerrar As PictureBox
    Friend WithEvents PickFormLoginMinimizar As PictureBox
    Friend WithEvents PickFormLoginLogin As PictureBox
    Friend WithEvents Panel2 As Panel
    Friend WithEvents Panel3 As Panel
    Friend WithEvents CheckFormLoginMostrarContraseña As CheckBox
    Friend WithEvents LabelFormLoginCambiarcontraseña As Label
    Friend WithEvents FormLoginBtnCambiar As Button
End Class
