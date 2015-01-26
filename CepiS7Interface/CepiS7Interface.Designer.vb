<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CepiS7Interface
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
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(CepiS7Interface))
        Me.Button_Connect = New System.Windows.Forms.Button
        Me.DataGridView_OPC = New System.Windows.Forms.DataGridView
        Me.Timer_OPC = New System.Windows.Forms.Timer(Me.components)
        Me.Button_Disconnect = New System.Windows.Forms.Button
        Me.ProgressBar_OPC = New System.Windows.Forms.ProgressBar
        Me.Button_Write = New System.Windows.Forms.Button
        Me.About = New System.Windows.Forms.Button
        Me.TextBox_Info = New System.Windows.Forms.TextBox
        Me.TextBox_Ip_Addr_1 = New System.Windows.Forms.TextBox
        Me.TextBox_Port_1 = New System.Windows.Forms.TextBox
        Me.TextBox_Port_2 = New System.Windows.Forms.TextBox
        Me.Label_Ip_Addr_1 = New System.Windows.Forms.Label
        Me.Label_Port_1 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.TextBox_Sample = New System.Windows.Forms.TextBox
        Me.Button_Salva = New System.Windows.Forms.Button
        Me.Button_Carica = New System.Windows.Forms.Button
        CType(Me.DataGridView_OPC, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Button_Connect
        '
        Me.Button_Connect.Location = New System.Drawing.Point(15, 120)
        Me.Button_Connect.Name = "Button_Connect"
        Me.Button_Connect.Size = New System.Drawing.Size(75, 23)
        Me.Button_Connect.TabIndex = 1
        Me.Button_Connect.Text = "Connect"
        Me.Button_Connect.UseVisualStyleBackColor = True
        '
        'DataGridView_OPC
        '
        Me.DataGridView_OPC.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView_OPC.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.DataGridView_OPC.Location = New System.Drawing.Point(0, 207)
        Me.DataGridView_OPC.Name = "DataGridView_OPC"
        Me.DataGridView_OPC.Size = New System.Drawing.Size(992, 555)
        Me.DataGridView_OPC.TabIndex = 3
        '
        'Timer_OPC
        '
        '
        'Button_Disconnect
        '
        Me.Button_Disconnect.Location = New System.Drawing.Point(206, 120)
        Me.Button_Disconnect.Name = "Button_Disconnect"
        Me.Button_Disconnect.Size = New System.Drawing.Size(75, 23)
        Me.Button_Disconnect.TabIndex = 4
        Me.Button_Disconnect.Text = "Disconnect"
        Me.Button_Disconnect.UseVisualStyleBackColor = True
        '
        'ProgressBar_OPC
        '
        Me.ProgressBar_OPC.Location = New System.Drawing.Point(15, 149)
        Me.ProgressBar_OPC.Name = "ProgressBar_OPC"
        Me.ProgressBar_OPC.Size = New System.Drawing.Size(266, 23)
        Me.ProgressBar_OPC.TabIndex = 5
        '
        'Button_Write
        '
        Me.Button_Write.Location = New System.Drawing.Point(15, 178)
        Me.Button_Write.Name = "Button_Write"
        Me.Button_Write.Size = New System.Drawing.Size(120, 23)
        Me.Button_Write.TabIndex = 6
        Me.Button_Write.Text = "Write SetValue (F9)"
        Me.Button_Write.UseVisualStyleBackColor = True
        '
        'About
        '
        Me.About.Location = New System.Drawing.Point(206, 177)
        Me.About.Name = "About"
        Me.About.Size = New System.Drawing.Size(75, 23)
        Me.About.TabIndex = 7
        Me.About.Text = "About"
        Me.About.UseVisualStyleBackColor = True
        '
        'TextBox_Info
        '
        Me.TextBox_Info.Location = New System.Drawing.Point(644, 12)
        Me.TextBox_Info.Multiline = True
        Me.TextBox_Info.Name = "TextBox_Info"
        Me.TextBox_Info.Size = New System.Drawing.Size(336, 189)
        Me.TextBox_Info.TabIndex = 8
        Me.TextBox_Info.Text = resources.GetString("TextBox_Info.Text")
        '
        'TextBox_Ip_Addr_1
        '
        Me.TextBox_Ip_Addr_1.Location = New System.Drawing.Point(131, 42)
        Me.TextBox_Ip_Addr_1.Name = "TextBox_Ip_Addr_1"
        Me.TextBox_Ip_Addr_1.Size = New System.Drawing.Size(150, 20)
        Me.TextBox_Ip_Addr_1.TabIndex = 10
        '
        'TextBox_Port_1
        '
        Me.TextBox_Port_1.Location = New System.Drawing.Point(131, 68)
        Me.TextBox_Port_1.Name = "TextBox_Port_1"
        Me.TextBox_Port_1.Size = New System.Drawing.Size(150, 20)
        Me.TextBox_Port_1.TabIndex = 11
        '
        'TextBox_Port_2
        '
        Me.TextBox_Port_2.Location = New System.Drawing.Point(131, 94)
        Me.TextBox_Port_2.Name = "TextBox_Port_2"
        Me.TextBox_Port_2.Size = New System.Drawing.Size(150, 20)
        Me.TextBox_Port_2.TabIndex = 12
        '
        'Label_Ip_Addr_1
        '
        Me.Label_Ip_Addr_1.AutoSize = True
        Me.Label_Ip_Addr_1.Location = New System.Drawing.Point(16, 45)
        Me.Label_Ip_Addr_1.Name = "Label_Ip_Addr_1"
        Me.Label_Ip_Addr_1.Size = New System.Drawing.Size(84, 13)
        Me.Label_Ip_Addr_1.TabIndex = 13
        Me.Label_Ip_Addr_1.Text = "PLC IP Address:"
        '
        'Label_Port_1
        '
        Me.Label_Port_1.AutoSize = True
        Me.Label_Port_1.Location = New System.Drawing.Point(16, 71)
        Me.Label_Port_1.Name = "Label_Port_1"
        Me.Label_Port_1.Size = New System.Drawing.Size(81, 13)
        Me.Label_Port_1.TabIndex = 14
        Me.Label_Port_1.Text = "PLC Port Read:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(16, 97)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(80, 13)
        Me.Label1.TabIndex = 15
        Me.Label1.Text = "PLC Port Write:"
        '
        'TextBox_Sample
        '
        Me.TextBox_Sample.Location = New System.Drawing.Point(287, 12)
        Me.TextBox_Sample.Multiline = True
        Me.TextBox_Sample.Name = "TextBox_Sample"
        Me.TextBox_Sample.Size = New System.Drawing.Size(351, 189)
        Me.TextBox_Sample.TabIndex = 16
        Me.TextBox_Sample.Text = resources.GetString("TextBox_Sample.Text")
        '
        'Button_Salva
        '
        Me.Button_Salva.Location = New System.Drawing.Point(206, 10)
        Me.Button_Salva.Name = "Button_Salva"
        Me.Button_Salva.Size = New System.Drawing.Size(75, 23)
        Me.Button_Salva.TabIndex = 17
        Me.Button_Salva.Text = "Salva"
        Me.Button_Salva.UseVisualStyleBackColor = True
        '
        'Button_Carica
        '
        Me.Button_Carica.Location = New System.Drawing.Point(15, 10)
        Me.Button_Carica.Name = "Button_Carica"
        Me.Button_Carica.Size = New System.Drawing.Size(75, 23)
        Me.Button_Carica.TabIndex = 18
        Me.Button_Carica.Text = "Carica"
        Me.Button_Carica.UseVisualStyleBackColor = True
        '
        'CepiS7Interface
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(992, 762)
        Me.Controls.Add(Me.Button_Carica)
        Me.Controls.Add(Me.Button_Salva)
        Me.Controls.Add(Me.TextBox_Sample)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label_Port_1)
        Me.Controls.Add(Me.Label_Ip_Addr_1)
        Me.Controls.Add(Me.TextBox_Port_2)
        Me.Controls.Add(Me.TextBox_Port_1)
        Me.Controls.Add(Me.TextBox_Ip_Addr_1)
        Me.Controls.Add(Me.TextBox_Info)
        Me.Controls.Add(Me.About)
        Me.Controls.Add(Me.Button_Write)
        Me.Controls.Add(Me.ProgressBar_OPC)
        Me.Controls.Add(Me.Button_Disconnect)
        Me.Controls.Add(Me.DataGridView_OPC)
        Me.Controls.Add(Me.Button_Connect)
        Me.KeyPreview = True
        Me.Name = "CepiS7Interface"
        Me.Text = "Cepi S7Interface"
        CType(Me.DataGridView_OPC, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Button_Connect As System.Windows.Forms.Button
    Friend WithEvents DataGridView_OPC As System.Windows.Forms.DataGridView
    Friend WithEvents Timer_OPC As System.Windows.Forms.Timer
    Friend WithEvents Button_Disconnect As System.Windows.Forms.Button
    Friend WithEvents ProgressBar_OPC As System.Windows.Forms.ProgressBar
    Friend WithEvents Button_Write As System.Windows.Forms.Button
    Friend WithEvents About As System.Windows.Forms.Button
    Friend WithEvents TextBox_Info As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_Ip_Addr_1 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_Port_1 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_Port_2 As System.Windows.Forms.TextBox
    Friend WithEvents Label_Ip_Addr_1 As System.Windows.Forms.Label
    Friend WithEvents Label_Port_1 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox_Sample As System.Windows.Forms.TextBox
    Friend WithEvents Button_Salva As System.Windows.Forms.Button
    Friend WithEvents Button_Carica As System.Windows.Forms.Button

End Class
