Friend Class Info
    Private Sub About_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Label_Info.Text = "S7Interface. Release: " & My.Application.Info.Version.ToString()
        LinkLabel_Website.Text = "www.cepisilos.com"
    End Sub

    Private Sub LinkLabel_Website_LinkClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel_Website.LinkClicked
        System.Diagnostics.Process.Start("http://www.cepisilos.com")
    End Sub
End Class