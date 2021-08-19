Imports System.IO
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Word

Public Class frm_Main
    Private _header As String = String.Empty
    Private _footer As String = String.Empty
    Private _intStop As Boolean = False

    'TO DO: Preserve original file margin sizes.
    '       Test how Page Break orientations are handled
    '       Suppress "detail" messages in log
    '       Protected documents not converting
    '       Preserve original file properties (date stamps, etc.)
    '       Preserve original folder structure

    ''' <summary>
    ''' Select the source directory
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub txt_SourceDirectory_DoubleClick(sender As Object, e As EventArgs) Handles txt_SourceDirectory.DoubleClick
        ChooseSourceDirectory()
    End Sub

    Private Sub ChooseSourceDirectory()
        Dim fbd As New FolderBrowserDialog()
        fbd.ShowNewFolderButton = False
        fbd.Description = "Please select source directory..."

        If fbd.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            txt_SourceDirectory.Text = fbd.SelectedPath.ToString
        End If
    End Sub

    Private Sub txt_TemplateFile_DoubleClick(sender As Object, e As EventArgs) Handles txt_TemplateFile.DoubleClick
        ChooseTemplateFile()
    End Sub

    Private Sub ChooseTemplateFile()
        Dim ofd As New OpenFileDialog()
        ofd.Title = "Please select document template file..."
        ofd.Filter = "Word Files (*.doc, *.docx, *.docm)|*.doc;*.docx;*.docm"
        ofd.Multiselect = False

        If ofd.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            txt_TemplateFile.Text = ofd.FileName.ToString
        End If
    End Sub

    Private Sub txt_OutputDirectory_DoubleClick(sender As Object, e As EventArgs) Handles txt_OutputDirectory.DoubleClick
        ChooseOutputDirectory()
    End Sub

    Private Sub ChooseOutputDirectory()
        Dim fbd As New FolderBrowserDialog()
        fbd.ShowNewFolderButton = True
        fbd.Description = "Please select output directory..."

        If fbd.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            txt_OutputDirectory.Text = fbd.SelectedPath.ToString
        End If
    End Sub

    Private Sub btn_Convert_Click(sender As Object, e As EventArgs) Handles btn_Convert.Click
        Dim strSourceDirectory As String = String.Empty
        Dim strTemplateFile As String = String.Empty
        Dim strOutputDirectory As String = String.Empty

        'ensure that a "source" directory has been specified
        If txt_SourceDirectory.Text = "" Then
            MessageBox.Show("Please specify a source directory.", "No source directory specified", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            UpdateStatus("No source directory specified.")
            Exit Sub
        Else
            strSourceDirectory = txt_SourceDirectory.Text
        End If

        'ensure that a "template" file has been specified
        If txt_TemplateFile.Text = "" Then
            MessageBox.Show("Please specify a template file.", "No template file specified", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            UpdateStatus("No template file specified.")
            Exit Sub
        Else
            strTemplateFile = txt_TemplateFile.Text
        End If

        'ensure that an "output" directory has been specified
        If txt_OutputDirectory.Text = "" Then
            MessageBox.Show("Please specify an output directory.", "No output directory specified", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            UpdateStatus("No output directory specified.")
            Exit Sub
        Else
            strOutputDirectory = txt_OutputDirectory.Text
        End If

        ConvertFiles(strSourceDirectory, strTemplateFile, strOutputDirectory)
    End Sub

    Private Sub ConvertFiles(SourceDirectory As String, TemplateFile As String, OutputDirectory As String)
        Try
            Dim AllFiles As New List(Of String)
            AllFiles = GetFilesRecursive(SourceDirectory)

            'If Preserve Original Folder Structure is checked, create all subfolders
            If chk_PreserveOriginalFolderStructure.Checked Then
                CreateSubDirectory(SourceDirectory, OutputDirectory)
            End If

            Dim i As Integer = 0
            prog_Progress.Value = i
            prog_Progress.Maximum = AllFiles.Count
            prog_Progress.Visible = True

            Dim dtStart As Date = Date.Now

            Dim strFilePrefix As String = Me.txt_ConvertedFilePrefix.Text

            For Each FilePath In AllFiles
                'Stop processing if the user clicks the "Stop" button
                If _intStop = True Then
                    Exit For
                End If

                i += 1
                UpdateStatus(String.Format("Processing file {0} of {1}...", i, AllFiles.Count))
                prog_Progress.Value = i

                Dim fileAttribs = File.GetAttributes(FilePath)
                If (fileAttribs And FileAttributes.System) Or (fileAttribs And FileAttributes.Hidden) Then
                    UpdateStatus(String.Format("    Skipping file ({0}) because it is either a Hidden or System file.", Path.GetFileName(FilePath)))
                    Continue For
                End If

                'only convert word documents
                Dim ext As String = Path.GetExtension(FilePath).ToLower
                If ext = ".doc" Or ext = ".dot" Or ext = ".docx" Or ext = ".docm" Or ext = ".dotx" Or ext = "dotm" Or ext = "docb" Then
                    ConvertToTemplate(FilePath, TemplateFile, OutputDirectory, strFilePrefix)
                Else
                    CopyFile(FilePath, OutputDirectory)
                End If
            Next

            prog_Progress.Value = 0
            prog_Progress.Visible = False

            'calculate the total time and avg time per file
            Dim dur As TimeSpan = Now.Subtract(dtStart)
            Dim dblAvgTimePerFile As Double = dur.TotalSeconds / i

            UpdateStatus(String.Format("File conversion(s) complete. {0} files processed in {1} seconds. Avg conversion took {2} seconds per file.", i, Math.Round(dur.TotalSeconds, 0), Math.Round(dblAvgTimePerFile, 2)))

        Catch ex As Exception
            UpdateStatus("ERROR: An error occurred while converting files. " & ex.Message)
        End Try
    End Sub

    Private Sub CopyFile(SourceFile As String, OutputDirectory As String)
        Try
            'Make sure the output directory ends in a backslash
            If Right(OutputDirectory, 1) <> "\" Then
                OutputDirectory = String.Format("{0}\", OutputDirectory)
            End If

            Dim OutputFile As String = String.Format("{0}{1}", OutputDirectory, Path.GetFileName(SourceFile))

            File.Copy(SourceFile, OutputFile, True)
            UpdateStatus(String.Format("  Copied file '{0}' to the output directory unchanged.", Path.GetFileName(SourceFile)))

        Catch ex As Exception
            'MessageBox.Show("An error occurred!" & Environment.NewLine & ex.Message, "ERROR!", MessageBoxButtons.OK, MessageBoxIcon.Error)
            UpdateStatus("ERROR: " & ex.Message)
        End Try
    End Sub

    Private Sub ConvertToTemplate(SourceDoc As String, TemplateDoc As String, OutputDirectory As String, FilePrefix As String)
        Dim WordApp As Word.Application = New Word.Application

        Try
            'Overwrite the Template Styles with the Original file's Styles
            'If chk_PreserveOriginalStyles.Checked Then
            '    PreserveOriginalStyles(TemplateDoc, SourceDoc)
            'End If

            Dim TempDoc As Word.Document
            'Open the document
            UpdateStatus("  Opening template document...")
            TempDoc = WordApp.Documents.Open(TemplateDoc)
            TempDoc.Activate()
            UpdateStatus("  Template document activated...")

            'Define the Template header and footer
            Dim templateHeader As HeaderFooter = TempDoc.Sections(1).Headers(WdHeaderFooterIndex.wdHeaderFooterPrimary)
            Dim templateFooter As HeaderFooter = TempDoc.Sections(1).Footers(WdHeaderFooterIndex.wdHeaderFooterPrimary)

            'Copy the template header to the clipboard
            templateHeader.Range.Copy()

            UpdateStatus(String.Format("   Converting '{0}' into new Template...", Path.GetFileName(SourceDoc)))
            'Insert the source file into the template, keeping the template Header/Footer
            Dim rng As Range = TempDoc.Range(0, 0)
            rng.InsertFile(SourceDoc)
            UpdateStatus("   Inserted source document into template...")

            'Overwrite the Template margins with the Original file's margins
            If chk_PreserveOriginalMargins.Checked Then
                PreserveOriginalMargins(TempDoc, SourceDoc)
            End If

            'Overwrite the template header in case it was erased during the "InsertFile" command
            TempDoc.Sections(1).Headers(WdHeaderFooterIndex.wdHeaderFooterPrimary).Range.Paste()

            'set the headers and footers on all pages to link to the previous (template file) header
            Dim sec As Word.Section
            For Each sec In TempDoc.Sections
                Dim hdr As Word.HeaderFooter
                'UpdateStatus(String.Format("    Currently in section {0}/{1}...", sec.Index, TempDoc.Sections.Count))
                For Each hdr In sec.Headers
                    Try
                        hdr.LinkToPrevious = True
                        hdr.PageNumbers.RestartNumberingAtSection = False

                    Catch ex As Exception
                        'UpdateStatus("     Error linking previous section. " & ex.Message)
                    End Try
                Next
            Next

            'prefix the converted file
            Dim newFileName As String = String.Format("{0}{1}", FilePrefix, Path.GetFileName(SourceDoc))

            'If Preserve Original Folder Structure is checked
            If chk_PreserveOriginalFolderStructure.Checked Then
                OutputDirectory = Path.GetDirectoryName(SourceDoc).Replace(txt_SourceDirectory.Text, txt_OutputDirectory.Text)
            End If

            'Make sure the output directory ends in a backslash
            If Right(OutputDirectory, 1) <> "\" Then
                OutputDirectory = String.Format("{0}\", OutputDirectory)
            End If

            Dim newFilePath As String = String.Format("{0}{1}", OutputDirectory, newFileName)

            'save using the .docx extension since that's what the template file was created as...
            newFilePath = Path.ChangeExtension(newFilePath, ".docx")

            TempDoc.SaveAs2(newFilePath)
            UpdateStatus("   New document saved.")

            TempDoc.Close()
            WordApp.Application.Quit(False)
            releaseObject(TempDoc)
            releaseObject(WordApp)

            UpdateStatus("   Word application closed.")
            UpdateStatus("   File converted and saved successfully.")

            If chk_PreserveOriginalFileDates.Checked Then
                File.SetCreationTime(newFilePath, File.GetCreationTime(SourceDoc))
                File.SetLastAccessTime(newFilePath, File.GetLastAccessTime(SourceDoc))
                File.SetLastAccessTimeUtc(newFilePath, File.GetLastAccessTimeUtc(SourceDoc))
                File.SetLastWriteTime(newFilePath, File.GetLastWriteTime(SourceDoc))
            End If

        Catch ex As Exception
            'dispose of the WordApp object to close the WINWORD.exe that was spawned
            WordApp.Quit()
            releaseObject(WordApp)

            'MessageBox.Show("An error occurred!" & Environment.NewLine & ex.Message, "ERROR!", MessageBoxButtons.OK, MessageBoxIcon.Error)
            UpdateStatus(" ERROR: " & ex.Message)
        End Try
    End Sub

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
            GC.Collect()
            GC.WaitForPendingFinalizers()
        Catch ex As Exception
            obj = Nothing
            GC.Collect()
            GC.WaitForPendingFinalizers()
        Finally
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Sub
    Private Sub PreserveOriginalMargins(ByRef TempDoc As Word.Document, SourceDoc As String)
        UpdateStatus("   Setting Template file margins to match Original document...")
        Dim WordApp As Word.Application = New Word.Application
        Dim OrigDoc As New Word.Document

        Try
            OrigDoc = WordApp.Documents.Open(SourceDoc)
            'OrigWordApp.Options.AllowReadingMode = True
            Dim OriginalProtection As WdProtectionType
            OriginalProtection = OrigDoc.ProtectionType
            If OriginalProtection <> WdProtectionType.wdNoProtection Then
                UpdateStatus("   Removing Protected state from original file...")
                OrigDoc.Unprotect()
                UpdateStatus("   Protected state removed successfully.")
            End If

            TempDoc.Sections.First.PageSetup.PageWidth = OrigDoc.Sections.First.PageSetup.PageWidth
            TempDoc.Sections.First.PageSetup.PageHeight = OrigDoc.Sections.First.PageSetup.PageHeight
            TempDoc.Sections.First.PageSetup.LeftMargin = OrigDoc.Sections.First.PageSetup.LeftMargin
            TempDoc.Sections.First.PageSetup.RightMargin = OrigDoc.Sections.First.PageSetup.RightMargin
            TempDoc.Sections.First.PageSetup.TopMargin = OrigDoc.Sections.First.PageSetup.TopMargin
            TempDoc.Sections.First.PageSetup.BottomMargin = OrigDoc.Sections.First.PageSetup.BottomMargin
            TempDoc.Sections.First.PageSetup.HeaderDistance = OrigDoc.Sections.First.PageSetup.HeaderDistance
            TempDoc.Sections.First.PageSetup.FooterDistance = OrigDoc.Sections.First.PageSetup.FooterDistance
            TempDoc.Sections.First.PageSetup.Gutter = OrigDoc.Sections.First.PageSetup.Gutter
            TempDoc.Sections.First.PageSetup.GutterPos = OrigDoc.Sections.First.PageSetup.GutterPos
            TempDoc.Sections.First.PageSetup.GutterStyle = OrigDoc.Sections.First.PageSetup.GutterStyle

            UpdateStatus("   Template margins set successfully.")

            ' Perform the business logic on the word document
            If OriginalProtection <> WdProtectionType.wdNoProtection Then
                OrigDoc.Protect(Type:=OriginalProtection, NoReset:=True)
                UpdateStatus("   Protected state reinstated for the original file.")
            End If

        Catch ex As Exception
            UpdateStatus("    ERROR: An error occurred while setting the Template file margins to match the Original document. " & ex.Message)

        Finally
            OrigDoc.Close(WdSaveOptions.wdDoNotSaveChanges)
            WordApp.Quit()
            releaseObject(WordApp)
        End Try
    End Sub

    Private Sub PreserveOriginalStyles(ByVal TemplateDoc As String, ByVal SourceDoc As String)
        UpdateStatus("   Setting Template Styles to match Original document...")
        Dim WordApp As Word.Application = New Word.Application
        Dim OrigDoc As New Word.Document

        Try
            OrigDoc = WordApp.Documents.Open(SourceDoc)
            Dim OriginalProtection As WdProtectionType
            OriginalProtection = OrigDoc.ProtectionType
            If OriginalProtection <> WdProtectionType.wdNoProtection Then
                UpdateStatus("   Removing Protected state from original file...")
                OrigDoc.Unprotect()
                UpdateStatus("   Protected state removed successfully.")
            End If

            For Each stl As Word.Style In OrigDoc.Styles
                WordApp.OrganizerCopy(SourceDoc, TemplateDoc, stl.NameLocal, WdOrganizerObject.wdOrganizerObjectStyles)
            Next

            UpdateStatus("   Template Styles set successfully.")

            ' Perform the business logic on the word document
            If OriginalProtection <> WdProtectionType.wdNoProtection Then
                OrigDoc.Protect(Type:=OriginalProtection, NoReset:=True)
                UpdateStatus("   Protected state reinstated for the original file.")
            End If

        Catch ex As Exception
            UpdateStatus("    ERROR: An error occurred while setting the Template Styles to match the Original document. " & ex.Message)

        Finally
            OrigDoc.Close(WdSaveOptions.wdDoNotSaveChanges)
            WordApp.Quit()
            releaseObject(WordApp)
        End Try
    End Sub

    ''' <summary>
    ''' This method starts at the specified directory and traverses through all subdirectores and returns a list of files
    ''' </summary>
    ''' <param name="InitialDirectory">Parent directory to begin scanning down from</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetFilesRecursive(ByVal InitialDirectory As String) As List(Of String)
        ' This list stores the results.
        Dim files As New List(Of String)

        ' This stack stores the directories to process.
        Dim directories As New Stack(Of String)

        ' Add the initial directory
        directories.Push(InitialDirectory)

        ' Continue processing for each stacked directory
        Do While (directories.Count > 0)
            ' Get top directory string
            Dim dir As String = directories.Pop
            Try
                ' Add all immediate file paths
                files.AddRange(Directory.GetFiles(dir, "*.*"))

                ' Loop through all subdirectories and add them to the stack.
                Dim directoryName As String
                For Each directoryName In Directory.GetDirectories(dir)
                    directories.Push(directoryName)
                Next

            Catch ex As Exception
            End Try
        Loop

        ' Return the list
        Return files
    End Function

    Private Sub CreateSubDirectory(ByVal SourcePath As String, DestPath As String)
        If Not Directory.Exists(DestPath) Then
            Directory.CreateDirectory(DestPath)
        End If

        For Each folder As String In Directory.GetDirectories(SourcePath)
            Dim dest As String = Path.Combine(DestPath, Path.GetFileName(folder))
            CreateSubDirectory(folder, dest)
        Next
    End Sub

    Private Shadows Function Right(Value As String, Length As Integer) As String
        Try
            Return Value.Substring(Value.Length - Length, Length)

        Catch ex As exception
            'do nothing
            Return String.Empty
        End Try
    End Function

    Private Sub UpdateStatus(Message As String)
        Dim msg As String = String.Format("{0} - {1}", Date.Now.ToString, Message)
        lbl_Status.Text = msg
        txt_Log.Text += String.Format("{0}{1}", msg, Environment.NewLine)
        System.Windows.Forms.Application.DoEvents()
    End Sub

    Private Sub frm_Main_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        SaveSettings()
    End Sub

    Private Sub SaveSettings()
        If txt_SourceDirectory.Text <> "" Then
            My.Settings.SourceDirectory = txt_SourceDirectory.Text
        End If

        If txt_TemplateFile.Text <> "" Then
            My.Settings.TemplateFile = txt_TemplateFile.Text
        End If

        If txt_OutputDirectory.Text <> "" Then
            My.Settings.OutputDirectory = txt_OutputDirectory.Text
        End If

        If txt_ConvertedFilePrefix.Text <> "" Then
            My.Settings.ConvertedFilePrefix = txt_ConvertedFilePrefix.Text
        End If

        My.Settings.PreserveOriginalMarginSizes = chk_PreserveOriginalMargins.Checked
        My.Settings.PreserveOriginalFileDates = chk_PreserveOriginalFileDates.Checked
        My.Settings.PreserveOriginalFolderStructure = chk_PreserveOriginalFolderStructure.Checked

        My.Settings.Save()
    End Sub

    Private Sub frm_Main_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadSettings()
    End Sub

    Private Sub LoadSettings()
        chk_PreserveOriginalMargins.Checked = My.Settings.PreserveOriginalMarginSizes
        chk_PreserveOriginalFileDates.Checked = My.Settings.PreserveOriginalFileDates
        chk_PreserveOriginalFolderStructure.Checked = My.Settings.PreserveOriginalFolderStructure
        txt_SourceDirectory.Text = My.Settings.SourceDirectory
        txt_TemplateFile.Text = My.Settings.TemplateFile
        txt_OutputDirectory.Text = My.Settings.OutputDirectory
        txt_ConvertedFilePrefix.Text = My.Settings.ConvertedFilePrefix
    End Sub

    Private Sub txt_SourceDirectory_TextChanged(sender As Object, e As EventArgs) Handles txt_SourceDirectory.TextChanged
        SaveSettings()
    End Sub

    Private Sub txt_TemplateFile_TextChanged(sender As Object, e As EventArgs) Handles txt_TemplateFile.TextChanged
        SaveSettings()
    End Sub

    Private Sub txt_OutputDirectory_TextChanged(sender As Object, e As EventArgs) Handles txt_OutputDirectory.TextChanged
        SaveSettings()
    End Sub

    Private Sub txt_ConvertedFilePrefix_TextChanged(sender As Object, e As EventArgs) Handles txt_ConvertedFilePrefix.TextChanged
        SaveSettings()
    End Sub

    Private Sub btn_Stop_Click(sender As Object, e As EventArgs) Handles btn_Stop.Click
        _intStop = True
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()  'SaveSettings is called automatically by the FormClosing event handler
    End Sub

    Private Sub SaveLogToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveLogToolStripMenuItem.Click
        SaveLog()
    End Sub

    Private Sub SaveLog()
        sfd_SaveFileDialog = New SaveFileDialog
        sfd_SaveFileDialog.FileName = "DocumentTemplateConverter.log"
        sfd_SaveFileDialog.Filter = "Log Files | *.log"

        If sfd_SaveFileDialog.ShowDialog() = DialogResult.OK Then
            IO.File.WriteAllText(sfd_SaveFileDialog.FileName, txt_Log.Text)
        End If

        UpdateStatus("Log contents saved to file.")
    End Sub

    Private Sub btn_ChooseSourceDirectory_Click(sender As Object, e As EventArgs) Handles btn_ChooseSourceDirectory.Click
        ChooseSourceDirectory()
    End Sub

    Private Sub btn_SelectTemplateFile_Click(sender As Object, e As EventArgs) Handles btn_SelectTemplateFile.Click
        ChooseTemplateFile()
    End Sub

    Private Sub btn_ChooseOutputDirectory_Click(sender As Object, e As EventArgs) Handles btn_ChooseOutputDirectory.Click
        ChooseOutputDirectory()
    End Sub

    Private Sub ClearLogToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClearLogToolStripMenuItem.Click
        ClearLog()
    End Sub

    Private Sub ClearLog()
        txt_Log.Text = ""
    End Sub

    Private Sub SaveApplicationSettingsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveApplicationSettingsToolStripMenuItem.Click
        SaveSettings()
    End Sub
End Class
