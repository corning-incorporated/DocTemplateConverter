' **************************************************
'   Copyright 2021 Corning Incorporated
'   Author: Tom Gee
' **************************************************

Imports System.IO
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Vbe.Interop
Imports System.Text

Public Class frm_Main
    Private _header As String = String.Empty
    Private _footer As String = String.Empty
    Private _intStop As Boolean = False
    Private _xlFileProperties As Hashtable
    Private _xlFilePropertyTypes As Hashtable
    Private _logPath As String = String.Empty
    'Private _oDocument As DSOFile.OleDocumentPropertiesClass
    Private _docFileProperties As Hashtable

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

    Private Sub txt_TemplateFile_DoubleClick(sender As Object, e As EventArgs) Handles txt_WordTemplateFile.DoubleClick
        ChooseWordTemplateFile()
    End Sub

    Private Sub ChooseWordTemplateFile()
        Dim ofd As New OpenFileDialog()
        ofd.Title = "Please select Word template file..."
        ofd.Filter = "Word Files (*.doc, *.docx, *.docm)|*.doc;*.docx;*.docm"
        ofd.Multiselect = False

        If ofd.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            txt_WordTemplateFile.Text = ofd.FileName.ToString
        End If
    End Sub

    Private Sub ChooseExcelTemplateFile()
        Dim ofd As New OpenFileDialog()
        ofd.Title = "Please select Excel template file..."
        ofd.Filter = "Excel Files (*.xls, *.xlsx, *.xlsm)|*.xls;*.xlsx;*.xlsm"
        ofd.Multiselect = False

        If ofd.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            txt_ExcelTemplateFile.Text = ofd.FileName.ToString
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
        Dim strWordTemplateFile As String = String.Empty
        Dim strExcelTemplateFile As String = String.Empty
        Dim strOutputDirectory As String = String.Empty

        SetupLogger()

        'ensure that a "source" directory has been specified
        If txt_SourceDirectory.Text = "" Then
            MessageBox.Show("Please specify a source directory.", "No source directory specified", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            UpdateStatus("No source directory specified.")
            Exit Sub
        Else
            strSourceDirectory = txt_SourceDirectory.Text
        End If

        'ensure that a Word "template" file has been specified
        If chk_ConvertWordFiles.Checked Then
            If txt_WordTemplateFile.Text = "" Then
                MessageBox.Show("Please specify a Word template file.", "No template file specified", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                UpdateStatus("No Word template file specified.")
                Exit Sub
            Else
                strWordTemplateFile = txt_WordTemplateFile.Text
            End If
        End If

        'ensure that a Excel "template" file has been specified
        If chk_ConvertExcelFiles.Checked Then
            If txt_ExcelTemplateFile.Text = "" Then
                MessageBox.Show("Please specify an Excel template file.", "No template file specified", MessageBoxButtons.OK, MessageBoxIcon.Stop)
                UpdateStatus("No Excel template file specified.")
                Exit Sub
            Else
                strExcelTemplateFile = txt_ExcelTemplateFile.Text
            End If
        End If

        'ensure that an "output" directory has been specified
        If txt_OutputDirectory.Text = "" Then
            MessageBox.Show("Please specify an output directory.", "No output directory specified", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            UpdateStatus("No output directory specified.")
            Exit Sub
        Else
            strOutputDirectory = txt_OutputDirectory.Text
        End If

        ConvertFiles(strSourceDirectory, strWordTemplateFile, strExcelTemplateFile, strOutputDirectory)
    End Sub

    Private Sub ConvertFiles(SourceDirectory As String, WordTemplateFile As String, ExcelTemplateFile As String, OutputDirectory As String)
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
            UpdateStatus(String.Format("Starting processing of {0} file(s)...", AllFiles.Count), True)

            Dim strFilePrefix As String = Me.txt_ConvertedFilePrefix.Text

            'Create a new dictionary object to hold the Excel templat file properties
            _xlFileProperties = New Hashtable
            _xlFilePropertyTypes = New Hashtable
            _docFileProperties = New Hashtable

            'Export the Excel VBA .cls file to import to each Excel file later. This function will also populate the xlFileProperties dictionary
            Dim ExcelClassPath As String = String.Empty

            If chk_ConvertExcelFiles.Checked Then
                ExcelClassPath = ExportExcelTemplateClass(ExcelTemplateFile)
            End If

            For Each FilePath In AllFiles
                'Stop processing if the user clicks the "Stop" button
                If _intStop = True Then
                    Exit For
                End If

                i += 1

                If txt_StartAtRecordNumber.Text <> "" Then
                    Try
                        Dim StartAtRecord As Integer = txt_StartAtRecordNumber.Text
                        If i < StartAtRecord Then
                            Continue For
                        End If

                    Catch ex As Exception
                        ' do nothing
                    End Try
                End If

                UpdateStatus(String.Format(" Processing file {0} of {1} [{2}]...", i, AllFiles.Count, Path.GetFileName(FilePath)), True)
                prog_Progress.Value = i
                System.Windows.Forms.Application.DoEvents()

                Dim fileAttribs = File.GetAttributes(FilePath)
                If (fileAttribs And FileAttributes.System) Or (fileAttribs And FileAttributes.Hidden) Then
                    UpdateStatus(String.Format("    Skipping file ({0}) because it is either a Hidden or System file.", Path.GetFileName(FilePath)))
                    Continue For
                End If

                'only convert word documents
                Dim ext As String = Path.GetExtension(FilePath).ToLower
                If ext = ".doc" Or ext = ".dot" Or ext = ".docx" Or ext = ".docm" Or ext = ".dotx" Or ext = "dotm" Or ext = "docb" Then
                    'CONVERT WORD FILE
                    If chk_ConvertWordFiles.Checked Then
                        ConvertToWordTemplate(FilePath, WordTemplateFile, OutputDirectory, strFilePrefix)
                    Else
                        If Not chk_OnlyCopyTemplateFilesTypes.Checked Then
                            CopyFile(FilePath, OutputDirectory)
                        End If
                    End If

                ElseIf ext = ".xls" Or ext = ".xlt" Or ext = ".xlsx" Or ext = ".xlsm" Or ext = ".xltx" Or ext = "xltm" Or ext = "xlsb" Or ext = "xlw" Then
                    'CONVERT EXCEL FILE
                    If chk_ConvertExcelFiles.Checked Then
                        ConvertToExcelTemplate(FilePath, ExcelClassPath, OutputDirectory, strFilePrefix)
                    Else
                        If Not chk_OnlyCopyTemplateFilesTypes.Checked Then
                            CopyFile(FilePath, OutputDirectory)
                        End If
                    End If

                Else
                    If Not chk_OnlyCopyTemplateFilesTypes.Checked Then
                        CopyFile(FilePath, OutputDirectory)
                    End If

                    If chk_OnlyCopyNONTemplateFilesTypes.Checked Then
                        CopyFile(FilePath, OutputDirectory)
                    End If
                End If
            Next

            prog_Progress.Value = 0
            prog_Progress.Visible = False

            'calculate the total time and avg time per file
            Dim dur As TimeSpan = Now.Subtract(dtStart)
            Dim dblAvgTimePerFile As Double = dur.TotalSeconds / i

            UpdateStatus(String.Format("File conversion(s) complete. {0} files processed in {1} seconds. Avg conversion took {2} seconds per file.", i, Math.Round(dur.TotalSeconds, 0), Math.Round(dblAvgTimePerFile, 2)), True)

        Catch ex As Exception
            UpdateStatus("ERROR: An error occurred while converting files. " & ex.Message)
        End Try
    End Sub

    Private Sub CopyFile(ByVal SourceFile As String, ByVal OutputDirectory As String)
        Try
            'Make sure the output directory ends in a backslash
            If Right(OutputDirectory, 1) <> "\" Then
                OutputDirectory = String.Format("{0}\", OutputDirectory)
            End If

            Dim OutputFile As String = String.Format("{0}{1}", OutputDirectory, Path.GetFileName(SourceFile))

            File.Copy(SourceFile, OutputFile, True)
            UpdateStatus(String.Format("   Copied file '{0}' to the output directory unchanged.", Path.GetFileName(SourceFile)))

        Catch ex As Exception
            'MessageBox.Show("An error occurred!" & Environment.NewLine & ex.Message, "ERROR!", MessageBoxButtons.OK, MessageBoxIcon.Error)
            UpdateStatus("   ERROR: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' This method performs the actual conversion of the original Word file into the new Template format
    ''' </summary>
    ''' <param name="SourceDoc">The path to the original document you wish to convert</param>
    ''' <param name="TemplateDoc">The path to the template file you want the original converted into</param>
    ''' <param name="OutputDirectory">The output directory where the converted file should be saved</param>
    ''' <param name="FilePrefix">The string that should be prefixed to the new file name (optional)</param>
    Private Sub ConvertToWordTemplate(SourceDoc As String, TemplateDoc As String, OutputDirectory As String, Optional FilePrefix As String = "")
        Dim WordApp As Word.Application = New Word.Application

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

        Try
            'Overwrite the Template Styles with the Original file's Styles
            'If chk_PreserveOriginalStyles.Checked Then
            '    PreserveOriginalStyles(TemplateDoc, SourceDoc)
            'End If

            Dim TempDoc As Word.Document
            'Open the document
            UpdateStatus("  Opening template document...")      '#DEBUG
            TempDoc = WordApp.Documents.Open(TemplateDoc)
            TempDoc.Activate()
            UpdateStatus("  Template document activated...")

            'Define the Template header and footer
            Dim templateHeader As Microsoft.Office.Interop.Word.HeaderFooter = TempDoc.Sections(1).Headers(WdHeaderFooterIndex.wdHeaderFooterPrimary)
            Dim templateFooter As Microsoft.Office.Interop.Word.HeaderFooter = TempDoc.Sections(1).Footers(WdHeaderFooterIndex.wdHeaderFooterPrimary)

            'Copy the template header to the clipboard
            templateHeader.Range.Copy()
            UpdateStatus("  Template header copied to clipboard...")

            UpdateStatus(String.Format("   Converting '{0}' into new Template...", Path.GetFileName(SourceDoc)))
            'Insert the source file into the template, keeping the template Header/Footer
            Dim rng As Microsoft.Office.Interop.Word.Range = TempDoc.Range(0, 0)
            rng.InsertFile(SourceDoc)
            UpdateStatus("   Inserted source document into template...")    '#DEBUG

            'Overwrite the Template margins with the Original file's margins
            If chk_PreserveOriginalMargins.Checked Then
                UpdateStatus(String.Format("   Preserve Margins - TempDoc: {0}, SourceDoc: {1}", TempDoc.FullName, SourceDoc))    '#DEBUG
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

                'Don't allow Word to change the header and footer for subsequent and odd/even pages
                Try
                    sec.PageSetup.OddAndEvenPagesHeaderFooter = False
                Catch ex As Exception
                    UpdateStatus("    ERROR setting section OddAndEvenPagesHeaderFooter to FALSE: " & ex.Message)
                End Try

                Try
                    sec.PageSetup.DifferentFirstPageHeaderFooter = False
                Catch ex As Exception
                    UpdateStatus("    ERROR setting section DifferentFirstPageHeaderFooter to FALSE: " & ex.Message)
                End Try
            Next


            'Don't allow Word to change the header and footer for subsequent and odd/even pages
            Try
                TempDoc.PageSetup.OddAndEvenPagesHeaderFooter = False
            Catch ex As Exception
                UpdateStatus("    WARNING setting Template Document OddAndEvenPagesHeaderFooter to FALSE: " & ex.Message)
            End Try

            Try
                TempDoc.PageSetup.DifferentFirstPageHeaderFooter = False
            Catch ex As Exception
                UpdateStatus("    WARNING setting Template Document DifferentFirstPageHeaderFooter to FALSE: " & ex.Message)
            End Try

            If chk_PreserveOriginalCustomProperties.Checked Then
                'Write the the source Word file's custom properties to the template
                Dim iAdded As Integer = 0
                Dim iUpdated As Integer = 0
                '_docFileProperties is set in the PreserveOriginalMargins method
                For Each prop In _docFileProperties
                    Try
                        Dim match As Boolean = False

                        For Each tempCustProp In TempDoc.CustomDocumentProperties
                            If tempCustProp.Name = prop.key Then
                                UpdateStatus(String.Format("    Custom Property [{0}] found in both files, updating template to match original value of [{1}]...", prop.Key, prop.Value))
                                tempCustProp.Value = prop.value
                                match = True
                                iUpdated += 1
                            End If
                        Next

                        If match = False Then
                            Try
                                TempDoc.CustomDocumentProperties.Add(prop.Key, False, Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeString, prop.Value)
                                UpdateStatus(String.Format("    Added Custom Property [{0}] to template with value of [{1}]...", prop.Key, prop.Value))
                                iAdded += 1
                            Catch ex As Exception
                                UpdateStatus(" ERROR Adding " & prop.Key & " Custom Property to template file: " & ex.Message)
                            End Try
                        End If

                    Catch ex As Exception
                        UpdateStatus(" ERROR setting Original Word File Properties: " & ex.Message)
                    End Try
                Next

                UpdateStatus(String.Format("   {0} file properties added and {1} file properties updated from the original Word file into new file.", iAdded, iUpdated))
            End If


            'save using the .docx extension since that's what the template file was created as...
            newFilePath = Path.ChangeExtension(newFilePath, ".docx")
                Try
                    If File.Exists(newFilePath) Then
                        Try
                            File.Delete(newFilePath)

                        Catch ex As Exception
                            UpdateStatus("    ERROR deleting existing Word file. " & ex.Message)
                        End Try

                    End If

                    TempDoc.SaveAs2(newFilePath)
                    UpdateStatus("   New document saved.")      '#DEBUG
                Catch ex As Exception
                    UpdateStatus("    ERROR saving Word file. " & ex.Message)
                End Try

                TempDoc.Close(False)
                UpdateStatus("   Template Doc closed.")     '#DEBUG

                WordApp.Application.Quit(False)
                UpdateStatus("   Template Word application closed.")    '#DEBUG
                releaseObject(TempDoc)
                'releaseObject(WordApp)

                UpdateStatus("   File converted and saved successfully.")

                If chk_PreserveOriginalFileDates.Checked Then
                    UpdateStatus("   Setting file date/times...")
                    File.SetCreationTime(newFilePath, File.GetCreationTime(SourceDoc))
                    File.SetLastAccessTime(newFilePath, File.GetLastAccessTime(SourceDoc))
                    File.SetLastAccessTimeUtc(newFilePath, File.GetLastAccessTimeUtc(SourceDoc))
                    File.SetLastWriteTime(newFilePath, File.GetLastWriteTime(SourceDoc))
                    UpdateStatus("   Date/times set successfully...")
                End If

            Catch ex As Exception
                'dispose of the WordApp object to close the WINWORD.exe that was spawned
                WordApp.NormalTemplate.Saved = True
            WordApp.Quit(False)
            UpdateStatus("   ERROR: Template Word app quit")   '#DEBUG
            releaseObject(WordApp)

            'MessageBox.Show("An error occurred!" & Environment.NewLine & ex.Message, "ERROR!", MessageBoxButtons.OK, MessageBoxIcon.Error)
            UpdateStatus("    ERROR: " & ex.Message)

            CopyFile(SourceDoc, OutputDirectory)
        End Try
    End Sub

#Region "Excel"
    Private Function ExportExcelTemplateClass(ByVal TemplateFile As String) As String
        Dim ExcelApp As Excel.Application = New Excel.Application

        Try
            Dim ExportPath As String = Path.GetDirectoryName(TemplateFile)
            If Right(ExportPath, 1) <> "\" Then
                ExportPath = String.Format("{0}\", ExportPath)
            End If

            ExportPath += "ThisWorkbook.cls"

            ExcelApp.Workbooks.Open(TemplateFile)

            Try
                ExcelApp.ActiveWorkbook.VBProject.VBComponents.Item("ThisWorkbook").Export(ExportPath)
                UpdateStatus(String.Format("  Exported template file VBA class to {0}", ExportPath))

                Dim arrText() As String
                Dim sLine As String = String.Empty
                'read the .cls file into memory
                arrText = File.ReadAllLines(ExportPath)
                'delete the original file because it will be recreated in the sw object below
                File.Delete(ExportPath)

                're-write the .cls file with the first 9 lines removed
                Using sw As StreamWriter = New StreamWriter(ExportPath)
                    'skip the first 9 lines
                    For i As Integer = 9 To arrText.Length() - 1
                        sw.WriteLine(arrText(i))
                    Next
                End Using
                UpdateStatus("   Cleansed VBA class file.")

                'Try to read the template file properties into memory       *********************************************************************
                Try
                    For Each p In ExcelApp.ActiveWorkbook.CustomDocumentProperties
                        _xlFileProperties.Add(p.Name, p.Value)
                        _xlFilePropertyTypes.Add(p.name, p.Type)
                    Next
                    UpdateStatus(String.Format("  {0} custom file properties read from the template file into memory.", _xlFileProperties.Count))

                Catch ex As Exception
                    UpdateStatus(" ERROR Reading Excel Template File Properties: " & ex.Message)
                    Return String.Empty
                End Try

                Return ExportPath

            Catch ex As Exception
                'MessageBox.Show("An error occurred!" & Environment.NewLine & ex.Message, "ERROR!", MessageBoxButtons.OK, MessageBoxIcon.Error)
                UpdateStatus(" ERROR: " & ex.Message)
                Return String.Empty

            Finally
                ExcelApp.ActiveWorkbook.Close(False)
                releaseObject(ExcelApp.ActiveWorkbook)
                ExcelApp.Quit()
                releaseObject(ExcelApp)
            End Try

        Catch ex As Exception
            'dispose of the WordApp object to close the WINWORD.exe that was spawned
            ExcelApp.Quit()
            releaseObject(ExcelApp)

            'MessageBox.Show("An error occurred!" & Environment.NewLine & ex.Message, "ERROR!", MessageBoxButtons.OK, MessageBoxIcon.Error)
            UpdateStatus(" ERROR: " & ex.Message)
            Return String.Empty
        End Try
    End Function

    Private Sub ConvertToExcelTemplate(SourceFile As String, ModulePath As String, OutputDirectory As String, FilePrefix As String)
        'prefix the converted file
        Dim newFileName As String = String.Format("{0}{1}", FilePrefix, Path.GetFileName(SourceFile))

        'If Preserve Original Folder Structure is checked
        If chk_PreserveOriginalFolderStructure.Checked Then
            OutputDirectory = Path.GetDirectoryName(SourceFile).Replace(txt_SourceDirectory.Text, txt_OutputDirectory.Text)
        End If

        'Make sure the output directory ends in a backslash
        If Right(OutputDirectory, 1) <> "\" Then
            OutputDirectory = String.Format("{0}\", OutputDirectory)
        End If

        Dim newFilePath As String = String.Format("{0}{1}", OutputDirectory, newFileName)

        Dim ExcelApp As Excel.Application = New Excel.Application

        Try
            'ExcelApp.DisplayAlerts = False
            With ExcelApp
                .Visible = False
                .DisplayAlerts = False
            End With

            Try
                ExcelApp.Workbooks.Open(SourceFile)

            Catch ex As Exception
                UpdateStatus(String.Format("  ERROR opening workbook ({0}) to Excel file: {1}", SourceFile, ex.Message))

                'dispose of the WordApp object to close the WINWORD.exe that was spawned
                ExcelApp.Quit()
                releaseObject(ExcelApp)

                'move the file, unmodified, to the output directory
                CopyFile(SourceFile, OutputDirectory)

                Exit Sub
            End Try

            UpdateStatus(String.Format("   Importing '{0}' into '{1}'...", ModulePath, Path.GetFileName(SourceFile)))
            'Import the template file's VBA class into the source file
            Dim xlModule = ExcelApp.ActiveWorkbook.VBProject.VBComponents.Item("ThisWorkbook")
            xlModule.CodeModule.AddFromFile(ModulePath)
            UpdateStatus("    VBA successfully inserted into original file...")    '#DEBUG

            'save using the .xlsm extension since that's what the template file was created as...
            newFilePath = Path.ChangeExtension(newFilePath, ".xlsm")

            'Copy the template file properties to the original Excel file

            Dim i As Integer = 0
            For Each prop In _xlFileProperties
                Try
                    Dim xlProperties = ExcelApp.ActiveWorkbook.CustomDocumentProperties
                    xlProperties.Add(prop.Key, False, _xlFilePropertyTypes(prop.key), prop.Value)

                    'UpdateStatus("     Added property '" & prop.Key & "', value '" & prop.Value & "', type '" & _xlFilePropertyTypes(prop.key) & "'...")
                    i += 1

                Catch ex As Exception
                    If (prop.Key.Contains("ETQ")) Then
                        'only log issues for EtQ-specific properties
                        UpdateStatus(String.Format("     ERROR copying workbook property ({0}) to Excel file: {1}", prop.key, ex.Message))
                    End If
                End Try
            Next

            UpdateStatus(String.Format("   {0} Custom Properties added to the file.", i))

            Try
                If File.Exists(newFilePath) Then
                    File.Delete(newFilePath)
                End If

                ExcelApp.ActiveWorkbook.SaveAs(newFilePath, Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled)

            Catch ex As Exception
                UpdateStatus(" ERROR saving new file: " & ex.Message)
            End Try

            ExcelApp.ActiveWorkbook.Close(False)
            releaseObject(ExcelApp.ActiveWorkbook)
            ExcelApp.Quit()
            releaseObject(ExcelApp)

            UpdateStatus("   File update complete.")

            If chk_PreserveOriginalFileDates.Checked Then
                File.SetCreationTime(newFilePath, File.GetCreationTime(SourceFile))
                File.SetLastAccessTime(newFilePath, File.GetLastAccessTime(SourceFile))
                File.SetLastAccessTimeUtc(newFilePath, File.GetLastAccessTimeUtc(SourceFile))
                File.SetLastWriteTime(newFilePath, File.GetLastWriteTime(SourceFile))
            End If

        Catch ex As Exception
            'MessageBox.Show("An error occurred!" & Environment.NewLine & ex.Message, "ERROR!", MessageBoxButtons.OK, MessageBoxIcon.Error)
            UpdateStatus("    ERROR: " & ex.Message)

            'dispose of the ExcelApp object to close the WINWORD.exe that was spawned
            ExcelApp.Quit()
            releaseObject(ExcelApp)

            'move the file, unmodified, to the output directory
            CopyFile(SourceFile, OutputDirectory)
        End Try
    End Sub

    Private Function ReadDocumentProperty(ByVal PropertyName As String, ByRef xlWorkbook As Excel.Workbook) As String
        Dim properties As Microsoft.Office.Core.DocumentProperties
        properties = CType(xlWorkbook.CustomDocumentProperties, Microsoft.Office.Core.DocumentProperties)

        For Each prop As Microsoft.Office.Core.DocumentProperty In properties
            If prop.Name = PropertyName Then
                Return prop.Value.ToString()
            End If
        Next

        Return Nothing
    End Function
#End Region

    Private Sub releaseObject(ByVal obj As Object)
        Try
            'System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
            UpdateStatus("   Object released.") '#DEBUG
        Catch ex As Exception
            'obj = Nothing
            UpdateStatus("   Error releasing Object." + ex.Message + " - " + ex.InnerException.ToString) '#DEBUG
        Finally
            System.Threading.Thread.Sleep(1000)
            'GC.Collect()
            'GC.WaitForPendingFinalizers()
        End Try
    End Sub
    Private Sub PreserveOriginalMargins(ByRef TempDoc As Word.Document, SourceDoc As String)
        UpdateStatus("   Setting Template file margins to match Original document...")      '#DEBUG
        Dim WordApp As Word.Application = New Word.Application
        Dim OrigDoc As Word.Document

        Try
            OrigDoc = WordApp.Documents.Open(SourceDoc)
            OrigDoc.Activate()
            UpdateStatus("    Original document opened to fetch Margin settings...")
            'OrigWordApp.Options.AllowReadingMode = True

            Dim OriginalProtection As WdProtectionType
            OriginalProtection = OrigDoc.ProtectionType
            If OriginalProtection <> WdProtectionType.wdNoProtection Then
                UpdateStatus("    Attempting to remove 'Protected' status from original file...")
                OrigDoc.Unprotect()
                UpdateStatus("    Protected state removed successfully.")
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
            TempDoc.Sections.First.PageSetup.GutterPos = OrigDoc.Sections.First.PageSetup.GutterPos
            TempDoc.Sections.First.PageSetup.GutterPos = OrigDoc.Sections.First.PageSetup.GutterPos
            TempDoc.Sections.First.PageSetup.GutterStyle = OrigDoc.Sections.First.PageSetup.GutterStyle

            UpdateStatus("    Template margins set successfully.")

            ' Perform the business logic on the word document
            If OriginalProtection <> WdProtectionType.wdNoProtection Then
                OrigDoc.Protect(Type:=OriginalProtection, NoReset:=True)
                UpdateStatus("    Protected state reinstated for the original file.")
            End If

            If chk_PreserveOriginalCustomProperties.Checked Then
                'Read the source Word file's custom properties into memory
                Try
                    _docFileProperties = New Hashtable
                    UpdateStatus(String.Format("   {0} custom file properties found...", OrigDoc.CustomDocumentProperties.Count))
                    For Each oCustProp In OrigDoc.CustomDocumentProperties
                        _docFileProperties.Add(oCustProp.Name, oCustProp.Value)
                    Next

                Catch ex As Exception
                    UpdateStatus(" ERROR Reading Original Word File Properties: " & ex.Message)
                End Try
            End If

        Catch ex As Exception
            UpdateStatus("    ERROR: An error occurred while setting the Template file margins to match the Original document. " & ex.Message)

        Finally
            Try
                OrigDoc.Close(WdSaveOptions.wdDoNotSaveChanges)
                UpdateStatus("    Original document closed.")    '#DEBUG
                UpdateStatus("    Original document closed.")    '#DEBUG
                UpdateStatus("    Original document closed.")    '#DEBUG
                UpdateStatus("    Original document closed.")    '#DEBUG

            Catch ex As Exception
                UpdateStatus("    ERROR closing Original document. " + ex.Message)    '#DEBUG
            End Try

            WordApp.NormalTemplate.Saved = True
            WordApp.Quit()
            UpdateStatus("    Original Word App closed.")    '#DEBUG
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

        Catch ex As Exception
            'do nothing
            Return String.Empty
        End Try
    End Function

    Private Sub UpdateStatus(Message As String, Optional ForceDisplay As Boolean = False)
        Dim msg As String = String.Format("{0} - {1}", Date.Now.ToString, Message)

        lbl_Status.Text = msg

        If chk_SaveLogToFile.Checked Then
            Using w As StreamWriter = File.AppendText(_logPath)
                w.WriteLine(String.Format("{0}: {1}", Date.Now, Message))
            End Using
        Else
            txt_Log.Text += String.Format("{0}{1}", msg, Environment.NewLine)

            Exit Sub '  Exit the Sub after displaying the message in the log textbox to prevent it from being displayed 2X if ForceDisplay = True
        End If

        'if the message is forced to display, show it
        If ForceDisplay Then
            txt_Log.Text += String.Format("{0}{1}", msg, Environment.NewLine)
        End If

        System.Windows.Forms.Application.DoEvents()
    End Sub

    Private Sub frm_Main_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        SaveSettings()
    End Sub

    Private Sub SaveSettings()
        If txt_SourceDirectory.Text <> "" Then
            My.Settings.SourceDirectory = txt_SourceDirectory.Text
        End If

        If txt_WordTemplateFile.Text <> "" Then
            My.Settings.WordTemplateFile = txt_WordTemplateFile.Text
        End If

        If txt_ExcelTemplateFile.Text <> "" Then
            My.Settings.ExcelTemplateFile = txt_ExcelTemplateFile.Text
        End If

        If txt_OutputDirectory.Text <> "" Then
            My.Settings.OutputDirectory = txt_OutputDirectory.Text
        End If

        If txt_ConvertedFilePrefix.Text <> "" Then
            My.Settings.ConvertedFilePrefix = txt_StartAtRecordNumber.Text
        End If

        My.Settings.PreserveOriginalMarginSizes = chk_PreserveOriginalMargins.Checked
        My.Settings.PreserveOriginalFileDates = chk_PreserveOriginalFileDates.Checked
        My.Settings.PreserveOriginalFolderStructure = chk_PreserveOriginalFolderStructure.Checked
        My.Settings.PreserveOriginalCustomProperties = chk_PreserveOriginalCustomProperties.Checked

        My.Settings.Save()
    End Sub

    Private Sub frm_Main_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadSettings()
    End Sub

    Private Sub LoadSettings()
        chk_PreserveOriginalMargins.Checked = My.Settings.PreserveOriginalMarginSizes
        chk_PreserveOriginalFileDates.Checked = My.Settings.PreserveOriginalFileDates
        chk_PreserveOriginalFolderStructure.Checked = My.Settings.PreserveOriginalFolderStructure
        chk_PreserveOriginalCustomProperties.Checked = My.Settings.PreserveOriginalCustomProperties
        txt_SourceDirectory.Text = My.Settings.SourceDirectory
        txt_WordTemplateFile.Text = My.Settings.WordTemplateFile
        txt_ExcelTemplateFile.Text = My.Settings.ExcelTemplateFile
        txt_OutputDirectory.Text = My.Settings.OutputDirectory
        txt_ConvertedFilePrefix.Text = My.Settings.ConvertedFilePrefix

        SetupLogger()
    End Sub

    Private Sub SetupLogger()
        If chk_SaveLogToFile.Checked Then
            If txt_OutputDirectory.Text <> "" Then
                _logPath = String.Format("{0}\{1}_DocTemplateConverter.log", txt_OutputDirectory.Text, Date.Today.ToString("yyyyMMdd"))
                'check to see if the Log path exists
                If Not File.Exists(_logPath) Then
                    'create the Log file if it doesn't already exist
                    File.CreateText(_logPath).Close()
                End If
            End If
        End If
    End Sub

    Private Sub txt_SourceDirectory_TextChanged(sender As Object, e As EventArgs) Handles txt_SourceDirectory.TextChanged
        SaveSettings()
    End Sub

    Private Sub txt_TemplateFile_TextChanged(sender As Object, e As EventArgs) Handles txt_WordTemplateFile.TextChanged
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

    Private Sub btn_SelectTemplateFile_Click(sender As Object, e As EventArgs) Handles btn_WordTemplateFile.Click
        ChooseWordTemplateFile()
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

    Private Sub btn_ExcelTemplateFile_Click(sender As Object, e As EventArgs) Handles btn_ExcelTemplateFile.Click
        ChooseExcelTemplateFile()

    End Sub

    Private Sub chk_ConvertWordFiles_CheckedChanged(sender As Object, e As EventArgs) Handles chk_ConvertWordFiles.CheckedChanged
        txt_WordTemplateFile.Enabled = chk_ConvertWordFiles.Checked
    End Sub

    Private Sub chk_ConvertExcelFiles_CheckedChanged(sender As Object, e As EventArgs) Handles chk_ConvertExcelFiles.CheckedChanged
        txt_ExcelTemplateFile.Enabled = chk_ConvertExcelFiles.Checked
    End Sub
End Class
