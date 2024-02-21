Public Class Google
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        WebBrowserurl.Navigate(txtURL.ToString)
    End Sub
End Class