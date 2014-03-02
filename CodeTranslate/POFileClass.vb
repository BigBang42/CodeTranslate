Public Class POFileClass

    Private _POFilePath As String
    Private _SourceFilesFolderPath As String
    Private _FilesToProcess As System.Collections.ObjectModel.ReadOnlyCollection(Of String)
    Private _POEntries As List(Of POEntry)
    Private _Files As List(Of FileToTranslate)

    Private _ProgressBar As ProgressBar
    Private _ProgressLabel As Label

    Public ReadOnly Property POEntries As List(Of POEntry)
        Get
            Return _POEntries
        End Get
    End Property

    ReadOnly Property Files As List(Of FileToTranslate)
        Get
            Return _Files
        End Get
    End Property

    Private Sub IntegratePOEntryIntoFileList(Entry As POEntry)
        Dim TempFile As FileToTranslate

        For Each AffectedFile As String In Entry.Files
            TempFile = (From F As FileToTranslate In _Files Where F.Name = AffectedFile).FirstOrDefault

            If TempFile IsNot Nothing Then
                TempFile = New FileToTranslate
                TempFile.Name = AffectedFile
                _Files.Add(TempFile)
            End If

            TempFile.Entries.Add(Entry.BaseEntry)
        Next
    End Sub

    Public Function CreateEntryList() As Boolean
        Dim Line As String
        Dim POEntry As POEntry

        Using StreamReader As New System.IO.StreamReader(_POFilePath)
            Do
                POEntry = New POEntry

                Do
                    Line = StreamReader.ReadLine
                    If Not POEntry.ImportPOLine(Line) Then Return False
                Loop Until POEntry.IsComplete

                _POEntries.Add(POEntry)
                IntegratePOEntryIntoFileList(POEntry)

            Loop Until StreamReader.EndOfStream
        End Using

        Return True
    End Function

    Public Shared Function Process(POFilePath As String, SourceCodeFolderPath As String, ProgressBar As ProgressBar) As Boolean
        Dim POFile As New System.IO.FileInfo(POFilePath)
        Dim FilesToProcess As New System.Collections.ObjectModel.ReadOnlyCollection(Of String)(My.Computer.FileSystem.GetFiles(SourceCodeFolderPath,
                                            Microsoft.VisualBasic.FileIO.SearchOption.SearchAllSubDirectories, "*.*"))

        ProgressBar.Maximum = FilesToProcess.Count

        For Each foundFile As String In FilesToProcess
            ProgressBar.Value += 1
            Application.Current.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Background, New System.Action(Function() ProgressBar.Visibility = True))

            Dim File As New System.IO.FileInfo(foundFile)

        Next

    End Function

    Public Sub New(POFilePath As String, SourceFilesFolderPath As String, ProgressBar As ProgressBar, ProgressLabel As Label)
        _POFilePath = POFilePath
        _SourceFilesFolderPath = SourceFilesFolderPath
        _FilesToProcess = New System.Collections.ObjectModel.ReadOnlyCollection(Of String)(My.Computer.FileSystem.GetFiles(SourceFilesFolderPath,
                                            Microsoft.VisualBasic.FileIO.SearchOption.SearchAllSubDirectories, "*.*"))

        _ProgressBar = ProgressBar
        _ProgressLabel = ProgressLabel
    End Sub


End Class
