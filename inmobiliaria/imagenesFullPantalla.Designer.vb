<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class imagenesFullPantalla
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
        Me.PickFormPrincipalCerrar = New System.Windows.Forms.PictureBox()
        Me.PicImagenesFullPantalla = New System.Windows.Forms.PictureBox()
        Me.LabelSiguiente = New System.Windows.Forms.Label()
        Me.LabelAnterior = New System.Windows.Forms.Label()
        CType(Me.PickFormPrincipalCerrar, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PicImagenesFullPantalla, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PickFormPrincipalCerrar
        '
        Me.PickFormPrincipalCerrar.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PickFormPrincipalCerrar.BackColor = System.Drawing.Color.FromArgb(CType(CType(38, Byte), Integer), CType(CType(55, Byte), Integer), CType(CType(52, Byte), Integer))
        Me.PickFormPrincipalCerrar.Cursor = System.Windows.Forms.Cursors.Default
        Me.PickFormPrincipalCerrar.Image = Global.inmobiliaria.My.Resources.Resources.cerrar
        Me.PickFormPrincipalCerrar.Location = New System.Drawing.Point(734, 12)
        Me.PickFormPrincipalCerrar.Name = "PickFormPrincipalCerrar"
        Me.PickFormPrincipalCerrar.Size = New System.Drawing.Size(54, 37)
        Me.PickFormPrincipalCerrar.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PickFormPrincipalCerrar.TabIndex = 12
        Me.PickFormPrincipalCerrar.TabStop = False
        '
        'PicImagenesFullPantalla
        '
        Me.PicImagenesFullPantalla.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PicImagenesFullPantalla.Location = New System.Drawing.Point(12, 12)
        Me.PicImagenesFullPantalla.Name = "PicImagenesFullPantalla"
        Me.PicImagenesFullPantalla.Size = New System.Drawing.Size(776, 423)
        Me.PicImagenesFullPantalla.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PicImagenesFullPantalla.TabIndex = 0
        Me.PicImagenesFullPantalla.TabStop = False
        '
        'LabelSiguiente
        '
        Me.LabelSiguiente.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.LabelSiguiente.AutoSize = True
        Me.LabelSiguiente.BackColor = System.Drawing.Color.Transparent
        Me.LabelSiguiente.Font = New System.Drawing.Font("Impact", 50.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelSiguiente.ForeColor = System.Drawing.Color.FromArgb(CType(CType(38, Byte), Integer), CType(CType(55, Byte), Integer), CType(CType(52, Byte), Integer))
        Me.LabelSiguiente.Location = New System.Drawing.Point(703, 185)
        Me.LabelSiguiente.Name = "LabelSiguiente"
        Me.LabelSiguiente.Size = New System.Drawing.Size(72, 82)
        Me.LabelSiguiente.TabIndex = 13
        Me.LabelSiguiente.Text = ">"
        '
        'LabelAnterior
        '
        Me.LabelAnterior.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.LabelAnterior.AutoSize = True
        Me.LabelAnterior.BackColor = System.Drawing.Color.Transparent
        Me.LabelAnterior.Font = New System.Drawing.Font("Impact", 50.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelAnterior.ForeColor = System.Drawing.Color.FromArgb(CType(CType(38, Byte), Integer), CType(CType(55, Byte), Integer), CType(CType(52, Byte), Integer))
        Me.LabelAnterior.Location = New System.Drawing.Point(12, 185)
        Me.LabelAnterior.Name = "LabelAnterior"
        Me.LabelAnterior.Size = New System.Drawing.Size(72, 82)
        Me.LabelAnterior.TabIndex = 14
        Me.LabelAnterior.Text = "<"
        '
        'imagenesFullPantalla
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.ControlBox = False
        Me.Controls.Add(Me.LabelAnterior)
        Me.Controls.Add(Me.LabelSiguiente)
        Me.Controls.Add(Me.PickFormPrincipalCerrar)
        Me.Controls.Add(Me.PicImagenesFullPantalla)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "imagenesFullPantalla"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "imagenes"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.PickFormPrincipalCerrar, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PicImagenesFullPantalla, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents PickFormPrincipalCerrar As PictureBox
    Public WithEvents PicImagenesFullPantalla As PictureBox
    Friend WithEvents LabelSiguiente As Label
    Friend WithEvents LabelAnterior As Label
End Class
