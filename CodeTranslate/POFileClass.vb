Imports System.IO

Public Class POFileClass

    Private _POFilePath As String
    Private _SourceFilesFolderPath As String
    Private _FilesToProcess As System.Collections.ObjectModel.ReadOnlyCollection(Of String)
    Private _POEntries As New List(Of POEntry)
    Private _Files As New List(Of FileToTranslate)

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

    Private Function SetCompleteFilePathInFilesAndRemoveIrrelevantOnes() As Boolean
        Dim Result As Boolean = True
        Dim TempFileList As New List(Of FileToTranslate)

        For Each File As FileToTranslate In _Files
            File.Name = (From FilePath As String In _FilesToProcess Where FilePath.Contains(File.Name)).FirstOrDefault
            If File.Name IsNot Nothing Then TempFileList.Add(File)
        Next

        Result = _Files.Count = TempFileList.Count
        _Files = TempFileList

        Return Result
    End Function

    Private Sub IntegratePOEntryIntoFileList(Entry As POEntry)
        Dim TempFile As FileToTranslate

        For Each AffectedFile As String In Entry.Files
            TempFile = (From F As FileToTranslate In _Files Where F.Name = AffectedFile).FirstOrDefault

            If TempFile Is Nothing Then
                TempFile = New FileToTranslate
                TempFile.Name = AffectedFile
                _Files.Add(TempFile)
            End If

            TempFile.Entries.Add(Entry.BaseEntry)
        Next
    End Sub

    Private Function CreateEntryList() As Integer
        Dim Line As String
        Dim LineNumber As Integer
        Dim POEntry As POEntry

        Using StreamReader As New System.IO.StreamReader(_POFilePath)
            Do
                POEntry = New POEntry

                Do
                    Line = StreamReader.ReadLine
                    LineNumber += 1

                    If StreamReader.EndOfStream Then Return True 'Quick and dirty ohne Fehlerprüfung, ob evtl. unvollständiger letzter Eintrag vorhanden, da Programm für nur einen Lauf programmiert ist
                    If Not POEntry.ImportPOLine(Line) Then
                        Return LineNumber
                    End If
                Loop Until POEntry.IsComplete

                _POEntries.Add(POEntry)
                IntegratePOEntryIntoFileList(POEntry)

            Loop Until StreamReader.EndOfStream
        End Using

        Return 0
    End Function

    Public Function Process() As Boolean
        Dim FileContent As String
        Dim NewFileContent As String
        Dim Result As Integer

        Using Log As New StreamWriter(Path.Combine(Path.GetDirectoryName(_POFilePath), "CodeTranslate.log"), False)
            Log.AutoFlush = True
            Log.WriteLine("{0}: Starting Process.", Now)

            _ProgressLabel.Content = "Creating entry list..."
            Log.WriteLine()
            Log.WriteLine("Creating entry list...")
            Result = CreateEntryList()
            If Result > 0 Then
                Log.WriteLine("Invalid .po-File Entry at Line {0} - Process canceled.", Result.ToString)
                MessageBox.Show("Invalid .po-File Entry at Line " & Result.ToString & " - Process canceled.", "Error parsing po.-File", MessageBoxButton.OK, MessageBoxImage.Error)
                Return False
            End If

            _ProgressLabel.Content = "Creating full file list..."
            Log.WriteLine()
            Log.WriteLine("Creating full file list...")
            If Not SetCompleteFilePathInFilesAndRemoveIrrelevantOnes() AndAlso _
                MessageBox.Show("Some files referred in the po.-file were not found in the specified folder. Maybe they have been deleted or moved since the po.-file was created or you have chosen a subfolder of the specified root folder when the po.-file was created. - Check Log for details. Continue anyway?", _
                                "Error creating full file list", MessageBoxButton.YesNo, MessageBoxImage.Error) = MsgBoxResult.No Then
                Return False
            End If

            _ProgressLabel.Content = "Translating files..."
            Log.WriteLine()
            Log.WriteLine("Translating files...")
            Log.WriteLine()
            _ProgressBar.Maximum = _Files.Count

            For Each File In _Files
                _ProgressBar.Value += 1
                Application.Current.Dispatcher.Invoke(System.Windows.Threading.DispatcherPriority.Background, New System.Action(Function() _ProgressBar.Visibility = True))

                FileContent = My.Computer.FileSystem.ReadAllText(File.Name)

                Log.WriteLine()
                Log.WriteLine("File: {0}", File.Name)
                Log.WriteLine()

                For Each Entry As BaseEntry In File.Entries
                    Log.WriteLine("MsgID: {0}", Entry.MsgID)
                    Log.WriteLine("MsgStr: {0}", Entry.MsgStr)

                    If Not String.IsNullOrEmpty(Entry.MsgStr) Then  'Nur schon übersetzte Strings ersetzen
                        NewFileContent = FileContent.Replace(String.Concat("__('", Entry.MsgID, "'"), Entry.MsgStr) 'MsgId kapseln, damit nicht evtl. Teilstrings ersetzt werden
                        If FileContent.Equals(NewFileContent) Then
                            Log.WriteLine("Error: MsgID not Found.")
                        Else
                            FileContent = NewFileContent
                            Log.WriteLine("Replaced.")
                        End If
                    Else
                        Log.WriteLine("Skipped.")
                    End If

                    Log.WriteLine()
                Next

                My.Computer.FileSystem.WriteAllText(File.Name, FileContent, False)
            Next

            _ProgressLabel.Content = "Done."
            Log.WriteLine()
            Log.WriteLine("{0}: Completed.", Now)
        End Using
        Return True
    End Function

    Public Sub New(POFilePath As String, SourceFilesFolderPath As String, ProgressBar As ProgressBar, ProgressLabel As Label)
        _POFilePath = POFilePath
        _SourceFilesFolderPath = SourceFilesFolderPath
        _FilesToProcess = New System.Collections.ObjectModel.ReadOnlyCollection(Of String)(My.Computer.FileSystem.GetFiles(SourceFilesFolderPath,
                                            Microsoft.VisualBasic.FileIO.SearchOption.SearchAllSubDirectories, "*.*"))

        _ProgressBar = ProgressBar
        _ProgressLabel = ProgressLabel

        POEntry.Reset()
    End Sub

End Class
