<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Info
    Inherits System.Windows.Forms.Form

    'Form esegue l'override del metodo Dispose per pulire l'elenco dei componenti.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Richiesto da Progettazione Windows Form
    Private components As System.ComponentModel.IContainer

    'NOTA: la procedura che segue è richiesta da Progettazione Windows Form
    'Può essere modificata in Progettazione Windows Form.  
    'Non modificarla nell'editor del codice.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Info))
        Me.Label_Info = New System.Windows.Forms.Label
        Me.LinkLabel_Website = New System.Windows.Forms.LinkLabel
        Me.RichTextBox_Info = New System.Windows.Forms.RichTextBox
        Me.SuspendLayout()
        '
        'Label_Info
        '
        Me.Label_Info.AutoSize = True
        Me.Label_Info.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_Info.Location = New System.Drawing.Point(12, 9)
        Me.Label_Info.Name = "Label_Info"
        Me.Label_Info.Size = New System.Drawing.Size(85, 20)
        Me.Label_Info.TabIndex = 0
        Me.Label_Info.Text = "Label_Info"
        '
        'LinkLabel_Website
        '
        Me.LinkLabel_Website.AutoSize = True
        Me.LinkLabel_Website.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LinkLabel_Website.Location = New System.Drawing.Point(12, 38)
        Me.LinkLabel_Website.Name = "LinkLabel_Website"
        Me.LinkLabel_Website.Size = New System.Drawing.Size(144, 20)
        Me.LinkLabel_Website.TabIndex = 2
        Me.LinkLabel_Website.TabStop = True
        Me.LinkLabel_Website.Text = "LinkLabel_Website"
        '
        'RichTextBox_Info
        '
        Me.RichTextBox_Info.Location = New System.Drawing.Point(12, 167)
        Me.RichTextBox_Info.Name = "RichTextBox_Info"
        Me.RichTextBox_Info.Size = New System.Drawing.Size(368, 283)
        Me.RichTextBox_Info.TabIndex = 7
        Me.RichTextBox_Info.Text = resources.GetString("RichTextBox_Info.Text")
        '
        'Info
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(392, 462)
        Me.Controls.Add(Me.RichTextBox_Info)
        Me.Controls.Add(Me.LinkLabel_Website)
        Me.Controls.Add(Me.Label_Info)
        Me.Name = "Info"
        Me.Text = "About"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label_Info As System.Windows.Forms.Label
    Friend WithEvents LinkLabel_Website As System.Windows.Forms.LinkLabel
    Friend WithEvents RichTextBox_Info As System.Windows.Forms.RichTextBox
End Class
