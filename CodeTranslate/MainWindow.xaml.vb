Imports System.Windows.Forms

Class MainWindow

    Private Sub cbSelectPOFile_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles cbSelectPOFile.Click
        Dim frmSelectPOFile As New Microsoft.Win32.OpenFileDialog

        If Not String.IsNullOrEmpty(tbPOFile.Text) Then frmSelectPOFile.FileName = tbPOFile.Text

        With frmSelectPOFile
            .Filter = "PO Files (*.po)|*.po"
            .Multiselect = False
        End With

        If frmSelectPOFile.ShowDialog() = True Then tbPOFile.Text = frmSelectPOFile.FileName

    End Sub

    Private Sub cbSelectSourceCodeFolder_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles cbSelectSourceCodeFolder.Click
        Dim frmSelectSourceCodeFolder As New System.Windows.Forms.FolderBrowserDialog

        If Not String.IsNullOrEmpty(tbSourceCodeFolder.Text) Then frmSelectSourceCodeFolder.SelectedPath = tbSourceCodeFolder.Text

        If frmSelectSourceCodeFolder.ShowDialog() = Forms.DialogResult.OK Then tbSourceCodeFolder.Text = frmSelectSourceCodeFolder.SelectedPath

    End Sub

    Private Sub cbCancel_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles cbCancel.Click
        MyBase.Close()
    End Sub

    Private Sub tbPOFile_TextChanged(sender As Object, e As System.Windows.Controls.TextChangedEventArgs) Handles tbPOFile.TextChanged
        cbStartReplacement.IsEnabled = (Not String.IsNullOrEmpty(tbPOFile.Text) And Not String.IsNullOrEmpty(tbSourceCodeFolder.Text))
    End Sub

    Private Sub tbSourceCodeFolder_TextChanged(sender As Object, e As System.Windows.Controls.TextChangedEventArgs) Handles tbSourceCodeFolder.TextChanged
        cbStartReplacement.IsEnabled = (Not String.IsNullOrEmpty(tbPOFile.Text) And Not String.IsNullOrEmpty(tbSourceCodeFolder.Text))
    End Sub

    Private Sub cbStartReplacement_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles cbStartReplacement.Click
        Dim POFile As New POFileClass(tbPOFile.Text, tbSourceCodeFolder.Text, pbProgress, lblProgress)

        POFile.Process()
    End Sub
End Class
