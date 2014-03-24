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

        Result = (_Files.Count = TempFileList.Count)
        _Files = TempFileList

        Return Result
    End Function

    Private Sub IntegratePOEntryIntoFileList(Entry As POEntry)
        Dim TempFile As FileToTranslate

        For Each AffectedFile As String In Entry.Files
            TempFile = (From F As FileToTranslate In _Files Where F.Name = AffectedFile).FirstOrDefault

            If TempFile Is Nothing Then     'Falls für dieses File bis jetzt noch kein POEntry-Eintrag besteht, es also noch nicht in der Liste ist, anlegen
                TempFile = New FileToTranslate
                TempFile.Name = AffectedFile
                _Files.Add(TempFile)
            End If

            TempFile.Entries.Add(Entry)     'POEntry zu diesem File hinzufügen
        Next
    End Sub

    Private Sub AddPOEntryToPOEntryList(POEntry)
        POEntry.Finish()
        _POEntries.Add(POEntry)
        IntegratePOEntryIntoFileList(POEntry)
    End Sub

    Private Function CreateEntryList() As Integer
        Dim Line As String
        Dim LineNumber As Integer
        Dim POEntry As New POEntry

        Using StreamReader As New System.IO.StreamReader(_POFilePath)
            Do While StreamReader.Peek > -1
                Line = StreamReader.ReadLine
                LineNumber += 1

                If Line.StartsWith("#:") AndAlso POEntry.IsValid Then   'Wenn neuer PO-Eintrag und aktuelle POEntry gültig dann Entry in Liste aufnehmen und neue leere erzeugen
                    AddPOEntryToPOEntryList(POEntry)
                    POEntry = New POEntry
                ElseIf POEntry.ImportPOLine(Line) Then  'Wenn die Struktur des POFiles stimmt, PO-Zeile in POEntry aufnehmen
                Else        'Strukturfehler im POFile
                    Return LineNumber
                End If
            Loop
        End Using

        If POEntry.IsEmpty Then
            Return 0
        ElseIf POEntry.IsValid And Not _POEntries.Contains(POEntry) Then    'Falls letzter Eintrag gültig und noch nicht in der Liste aufgenommen wurde
            AddPOEntryToPOEntryList(POEntry)
            Return 0
        Else    'Letzte POEntry ist ungültig
            Return LineNumber
        End If
    End Function


    Private Function ReplaceEntry(FileContent As String, POEntry As POEntry) As String
        Dim SearchString As String = String.Empty, ReplaceString As String = String.Empty

        'Ist jetzt erstmal schnell programmiert für genaue Übereinstimmung mit dem Searchstring - ansonsten wird variabel im FileContent unabhängig von Whitespaces u. ä. gesucht

        Select Case POEntry.Properties.Type
            Case EntryType.SingularText
                SearchString = String.Concat(POEntry.Properties.Prefix, "('", POEntry.MsgID, "'")
                ReplaceString = String.Concat(POEntry.Properties.Prefix, "('", POEntry.MsgStr, "'")
            Case EntryType.PluralText
                SearchString = String.Concat(POEntry.Properties.Prefix, "('", POEntry.MsgID, "', '", POEntry.MsgID_Plural, "'")
                ReplaceString = String.Concat(POEntry.Properties.Prefix, "('", POEntry.MsgStr, "', '", POEntry.MsgStr_Plural, "'")
        End Select

        Return FileContent.Replace(SearchString, ReplaceString) 'MsgId kapseln, damit nicht evtl. Teilstrings ersetzt werden
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

                For Each Entry As POEntry In File.Entries
                    Log.WriteLine("MsgID: {0}", Entry.MsgID)
                    Log.WriteLine("MsgStr: {0}", Entry.MsgStr)

                    If Entry.Properties.Type = EntryType.PluralText Then
                        Log.WriteLine("MsgID_Plural: {0}", Entry.MsgID_Plural)
                        Log.WriteLine("MsgStr_Plural: {0}", Entry.MsgStr_Plural)
                    End If

                    If Not String.IsNullOrEmpty(Entry.MsgStr) Then  'Nur schon übersetzte Strings ersetzen
                        'NewFileContent = FileContent.Replace(String.Concat("__('", Entry.MsgID, "'"), Entry.MsgStr) 'MsgId kapseln, damit nicht evtl. Teilstrings ersetzt werden
                        NewFileContent = ReplaceEntry(FileContent, Entry)
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

    End Sub

End Class
