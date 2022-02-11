' **************************************************
'   Copyright 2021 Corning Incorporated
'   Author: Tom Gee
' **************************************************

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frm_Main
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frm_Main))
        Me.lbl_SourceDirectory = New System.Windows.Forms.Label()
        Me.txt_SourceDirectory = New System.Windows.Forms.TextBox()
        Me.txt_WordTemplateFile = New System.Windows.Forms.TextBox()
        Me.lbl_WordTemplateFile = New System.Windows.Forms.Label()
        Me.txt_OutputDirectory = New System.Windows.Forms.TextBox()
        Me.lbl_OutputDirectory = New System.Windows.Forms.Label()
        Me.btn_Convert = New System.Windows.Forms.Button()
        Me.txt_Log = New System.Windows.Forms.TextBox()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.prog_Progress = New System.Windows.Forms.ToolStripProgressBar()
        Me.lbl_Status = New System.Windows.Forms.ToolStripStatusLabel()
        Me.txt_StartAtRecordNumber = New System.Windows.Forms.TextBox()
        Me.lbl_StartAtRecordNumber = New System.Windows.Forms.Label()
        Me.btn_Stop = New System.Windows.Forms.Button()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.btn_WordTemplateFile = New System.Windows.Forms.Button()
        Me.btn_ChooseOutputDirectory = New System.Windows.Forms.Button()
        Me.btn_ChooseSourceDirectory = New System.Windows.Forms.Button()
        Me.btn_ExcelTemplateFile = New System.Windows.Forms.Button()
        Me.txt_ExcelTemplateFile = New System.Windows.Forms.TextBox()
        Me.chk_PreserveOriginalMargins = New System.Windows.Forms.CheckBox()
        Me.chk_PreserveOriginalFileDates = New System.Windows.Forms.CheckBox()
        Me.chk_PreserveOriginalFolderStructure = New System.Windows.Forms.CheckBox()
        Me.chk_SaveLogToFile = New System.Windows.Forms.CheckBox()
        Me.chk_ConvertWordFiles = New System.Windows.Forms.CheckBox()
        Me.chk_ConvertExcelFiles = New System.Windows.Forms.CheckBox()
        Me.chk_OnlyCopyTemplateFilesTypes = New System.Windows.Forms.CheckBox()
        Me.txt_ConvertedFilePrefix = New System.Windows.Forms.TextBox()
        Me.chk_OnlyCopyNONTemplateFilesTypes = New System.Windows.Forms.CheckBox()
        Me.chk_PreserveOriginalCustomProperties = New System.Windows.Forms.CheckBox()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SaveApplicationSettingsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.ClearLogToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SaveLogToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CheckForUpdateToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HelpToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ActivateProductLicenseToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.FrequentlyAskedQuestionsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.sfd_SaveFileDialog = New System.Windows.Forms.SaveFileDialog()
        Me.lbl_ExcelTemplateFile = New System.Windows.Forms.Label()
        Me.lbl_ConvertedFilePrefix = New System.Windows.Forms.Label()
        Me.StatusStrip1.SuspendLayout()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'lbl_SourceDirectory
        '
        Me.lbl_SourceDirectory.AutoSize = True
        Me.lbl_SourceDirectory.BackColor = System.Drawing.Color.Transparent
        Me.lbl_SourceDirectory.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_SourceDirectory.ForeColor = System.Drawing.Color.Gainsboro
        Me.lbl_SourceDirectory.Location = New System.Drawing.Point(56, 35)
        Me.lbl_SourceDirectory.Name = "lbl_SourceDirectory"
        Me.lbl_SourceDirectory.Size = New System.Drawing.Size(106, 13)
        Me.lbl_SourceDirectory.TabIndex = 0
        Me.lbl_SourceDirectory.Text = "Source Directory:"
        '
        'txt_SourceDirectory
        '
        Me.txt_SourceDirectory.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txt_SourceDirectory.Location = New System.Drawing.Point(168, 32)
        Me.txt_SourceDirectory.Name = "txt_SourceDirectory"
        Me.txt_SourceDirectory.Size = New System.Drawing.Size(576, 20)
        Me.txt_SourceDirectory.TabIndex = 0
        Me.ToolTip1.SetToolTip(Me.txt_SourceDirectory, "Double-click to select a source directory (original files)")
        '
        'txt_WordTemplateFile
        '
        Me.txt_WordTemplateFile.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txt_WordTemplateFile.Location = New System.Drawing.Point(168, 58)
        Me.txt_WordTemplateFile.Name = "txt_WordTemplateFile"
        Me.txt_WordTemplateFile.Size = New System.Drawing.Size(576, 20)
        Me.txt_WordTemplateFile.TabIndex = 2
        Me.ToolTip1.SetToolTip(Me.txt_WordTemplateFile, "Double-click to select a Word template file")
        '
        'lbl_WordTemplateFile
        '
        Me.lbl_WordTemplateFile.AutoSize = True
        Me.lbl_WordTemplateFile.BackColor = System.Drawing.Color.Transparent
        Me.lbl_WordTemplateFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_WordTemplateFile.ForeColor = System.Drawing.Color.Gainsboro
        Me.lbl_WordTemplateFile.Location = New System.Drawing.Point(41, 61)
        Me.lbl_WordTemplateFile.Name = "lbl_WordTemplateFile"
        Me.lbl_WordTemplateFile.Size = New System.Drawing.Size(121, 13)
        Me.lbl_WordTemplateFile.TabIndex = 2
        Me.lbl_WordTemplateFile.Text = "Word Template File:"
        '
        'txt_OutputDirectory
        '
        Me.txt_OutputDirectory.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txt_OutputDirectory.Location = New System.Drawing.Point(168, 110)
        Me.txt_OutputDirectory.Name = "txt_OutputDirectory"
        Me.txt_OutputDirectory.Size = New System.Drawing.Size(576, 20)
        Me.txt_OutputDirectory.TabIndex = 6
        Me.ToolTip1.SetToolTip(Me.txt_OutputDirectory, "Double-click to select an output directory where you want the new files saved")
        '
        'lbl_OutputDirectory
        '
        Me.lbl_OutputDirectory.AutoSize = True
        Me.lbl_OutputDirectory.BackColor = System.Drawing.Color.Transparent
        Me.lbl_OutputDirectory.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_OutputDirectory.ForeColor = System.Drawing.Color.Gainsboro
        Me.lbl_OutputDirectory.Location = New System.Drawing.Point(58, 113)
        Me.lbl_OutputDirectory.Name = "lbl_OutputDirectory"
        Me.lbl_OutputDirectory.Size = New System.Drawing.Size(104, 13)
        Me.lbl_OutputDirectory.TabIndex = 4
        Me.lbl_OutputDirectory.Text = "Output Directory:"
        '
        'btn_Convert
        '
        Me.btn_Convert.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_Convert.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_Convert.Location = New System.Drawing.Point(12, 210)
        Me.btn_Convert.Name = "btn_Convert"
        Me.btn_Convert.Size = New System.Drawing.Size(732, 46)
        Me.btn_Convert.TabIndex = 13
        Me.btn_Convert.Text = "CONVERT!"
        Me.ToolTip1.SetToolTip(Me.btn_Convert, "Convert all documents in the Source Directory and all subfolders")
        Me.btn_Convert.UseVisualStyleBackColor = True
        '
        'txt_Log
        '
        Me.txt_Log.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txt_Log.BackColor = System.Drawing.Color.DimGray
        Me.txt_Log.ForeColor = System.Drawing.Color.WhiteSmoke
        Me.txt_Log.Location = New System.Drawing.Point(0, 262)
        Me.txt_Log.Multiline = True
        Me.txt_Log.Name = "txt_Log"
        Me.txt_Log.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txt_Log.Size = New System.Drawing.Size(808, 195)
        Me.txt_Log.TabIndex = 15
        '
        'StatusStrip1
        '
        Me.StatusStrip1.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.prog_Progress, Me.lbl_Status})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 457)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(808, 22)
        Me.StatusStrip1.TabIndex = 8
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'prog_Progress
        '
        Me.prog_Progress.Name = "prog_Progress"
        Me.prog_Progress.Size = New System.Drawing.Size(100, 16)
        Me.prog_Progress.Visible = False
        '
        'lbl_Status
        '
        Me.lbl_Status.Name = "lbl_Status"
        Me.lbl_Status.Size = New System.Drawing.Size(0, 17)
        '
        'txt_StartAtRecordNumber
        '
        Me.txt_StartAtRecordNumber.Location = New System.Drawing.Point(168, 136)
        Me.txt_StartAtRecordNumber.Name = "txt_StartAtRecordNumber"
        Me.txt_StartAtRecordNumber.Size = New System.Drawing.Size(89, 20)
        Me.txt_StartAtRecordNumber.TabIndex = 8
        Me.ToolTip1.SetToolTip(Me.txt_StartAtRecordNumber, "Specify a prefix that all converted files should include (optional)")
        '
        'lbl_StartAtRecordNumber
        '
        Me.lbl_StartAtRecordNumber.AutoSize = True
        Me.lbl_StartAtRecordNumber.BackColor = System.Drawing.Color.Transparent
        Me.lbl_StartAtRecordNumber.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_StartAtRecordNumber.ForeColor = System.Drawing.Color.Gainsboro
        Me.lbl_StartAtRecordNumber.Location = New System.Drawing.Point(52, 139)
        Me.lbl_StartAtRecordNumber.Name = "lbl_StartAtRecordNumber"
        Me.lbl_StartAtRecordNumber.Size = New System.Drawing.Size(110, 13)
        Me.lbl_StartAtRecordNumber.TabIndex = 9
        Me.lbl_StartAtRecordNumber.Text = "Start at Record #:"
        '
        'btn_Stop
        '
        Me.btn_Stop.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_Stop.Image = Global.Document_Converter.My.Resources.Resources.red_square
        Me.btn_Stop.Location = New System.Drawing.Point(751, 211)
        Me.btn_Stop.Name = "btn_Stop"
        Me.btn_Stop.Size = New System.Drawing.Size(45, 45)
        Me.btn_Stop.TabIndex = 14
        Me.ToolTip1.SetToolTip(Me.btn_Stop, "Stop Converter")
        Me.btn_Stop.UseVisualStyleBackColor = True
        '
        'btn_WordTemplateFile
        '
        Me.btn_WordTemplateFile.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_WordTemplateFile.Image = Global.Document_Converter.My.Resources.Resources.word
        Me.btn_WordTemplateFile.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btn_WordTemplateFile.Location = New System.Drawing.Point(751, 55)
        Me.btn_WordTemplateFile.Name = "btn_WordTemplateFile"
        Me.btn_WordTemplateFile.Size = New System.Drawing.Size(45, 24)
        Me.btn_WordTemplateFile.TabIndex = 3
        Me.ToolTip1.SetToolTip(Me.btn_WordTemplateFile, "Choose template file")
        Me.btn_WordTemplateFile.UseVisualStyleBackColor = True
        '
        'btn_ChooseOutputDirectory
        '
        Me.btn_ChooseOutputDirectory.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_ChooseOutputDirectory.Image = Global.Document_Converter.My.Resources.Resources.folder_next
        Me.btn_ChooseOutputDirectory.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btn_ChooseOutputDirectory.Location = New System.Drawing.Point(751, 107)
        Me.btn_ChooseOutputDirectory.Name = "btn_ChooseOutputDirectory"
        Me.btn_ChooseOutputDirectory.Size = New System.Drawing.Size(45, 24)
        Me.btn_ChooseOutputDirectory.TabIndex = 7
        Me.ToolTip1.SetToolTip(Me.btn_ChooseOutputDirectory, "Choose output directory (location of converted files)")
        Me.btn_ChooseOutputDirectory.UseVisualStyleBackColor = True
        '
        'btn_ChooseSourceDirectory
        '
        Me.btn_ChooseSourceDirectory.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_ChooseSourceDirectory.Image = Global.Document_Converter.My.Resources.Resources.folder_process
        Me.btn_ChooseSourceDirectory.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btn_ChooseSourceDirectory.Location = New System.Drawing.Point(750, 29)
        Me.btn_ChooseSourceDirectory.Name = "btn_ChooseSourceDirectory"
        Me.btn_ChooseSourceDirectory.Size = New System.Drawing.Size(45, 24)
        Me.btn_ChooseSourceDirectory.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.btn_ChooseSourceDirectory, "Choose ""source"" directory (location of original files)")
        Me.btn_ChooseSourceDirectory.UseVisualStyleBackColor = True
        '
        'btn_ExcelTemplateFile
        '
        Me.btn_ExcelTemplateFile.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_ExcelTemplateFile.Image = Global.Document_Converter.My.Resources.Resources.excel24
        Me.btn_ExcelTemplateFile.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btn_ExcelTemplateFile.Location = New System.Drawing.Point(751, 81)
        Me.btn_ExcelTemplateFile.Name = "btn_ExcelTemplateFile"
        Me.btn_ExcelTemplateFile.Size = New System.Drawing.Size(45, 24)
        Me.btn_ExcelTemplateFile.TabIndex = 5
        Me.ToolTip1.SetToolTip(Me.btn_ExcelTemplateFile, "Choose template file")
        Me.btn_ExcelTemplateFile.UseVisualStyleBackColor = True
        '
        'txt_ExcelTemplateFile
        '
        Me.txt_ExcelTemplateFile.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txt_ExcelTemplateFile.Location = New System.Drawing.Point(168, 84)
        Me.txt_ExcelTemplateFile.Name = "txt_ExcelTemplateFile"
        Me.txt_ExcelTemplateFile.Size = New System.Drawing.Size(576, 20)
        Me.txt_ExcelTemplateFile.TabIndex = 4
        Me.ToolTip1.SetToolTip(Me.txt_ExcelTemplateFile, "Double-click to select an Excel template file")
        '
        'chk_PreserveOriginalMargins
        '
        Me.chk_PreserveOriginalMargins.AutoSize = True
        Me.chk_PreserveOriginalMargins.BackColor = System.Drawing.Color.Transparent
        Me.chk_PreserveOriginalMargins.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chk_PreserveOriginalMargins.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chk_PreserveOriginalMargins.ForeColor = System.Drawing.Color.Gainsboro
        Me.chk_PreserveOriginalMargins.Location = New System.Drawing.Point(16, 187)
        Me.chk_PreserveOriginalMargins.Name = "chk_PreserveOriginalMargins"
        Me.chk_PreserveOriginalMargins.Size = New System.Drawing.Size(199, 17)
        Me.chk_PreserveOriginalMargins.TabIndex = 9
        Me.chk_PreserveOriginalMargins.Text = "Preserve Original Margin Sizes"
        Me.ToolTip1.SetToolTip(Me.chk_PreserveOriginalMargins, "Check this box to have the margin sizes in the new Word file match the original")
        Me.chk_PreserveOriginalMargins.UseVisualStyleBackColor = False
        '
        'chk_PreserveOriginalFileDates
        '
        Me.chk_PreserveOriginalFileDates.AutoSize = True
        Me.chk_PreserveOriginalFileDates.BackColor = System.Drawing.Color.Transparent
        Me.chk_PreserveOriginalFileDates.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chk_PreserveOriginalFileDates.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chk_PreserveOriginalFileDates.ForeColor = System.Drawing.Color.Gainsboro
        Me.chk_PreserveOriginalFileDates.Location = New System.Drawing.Point(236, 187)
        Me.chk_PreserveOriginalFileDates.Name = "chk_PreserveOriginalFileDates"
        Me.chk_PreserveOriginalFileDates.Size = New System.Drawing.Size(184, 17)
        Me.chk_PreserveOriginalFileDates.TabIndex = 10
        Me.chk_PreserveOriginalFileDates.Text = "Preserve Original File Dates"
        Me.chk_PreserveOriginalFileDates.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.chk_PreserveOriginalFileDates, "Check this box to have the file dates for the newly converted file set to the ori" &
        "ginal file's values")
        Me.chk_PreserveOriginalFileDates.UseVisualStyleBackColor = False
        '
        'chk_PreserveOriginalFolderStructure
        '
        Me.chk_PreserveOriginalFolderStructure.AutoSize = True
        Me.chk_PreserveOriginalFolderStructure.BackColor = System.Drawing.Color.Transparent
        Me.chk_PreserveOriginalFolderStructure.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chk_PreserveOriginalFolderStructure.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chk_PreserveOriginalFolderStructure.ForeColor = System.Drawing.Color.Gainsboro
        Me.chk_PreserveOriginalFolderStructure.Location = New System.Drawing.Point(441, 187)
        Me.chk_PreserveOriginalFolderStructure.Name = "chk_PreserveOriginalFolderStructure"
        Me.chk_PreserveOriginalFolderStructure.Size = New System.Drawing.Size(218, 17)
        Me.chk_PreserveOriginalFolderStructure.TabIndex = 11
        Me.chk_PreserveOriginalFolderStructure.Text = "Preserve Original Folder Structure"
        Me.chk_PreserveOriginalFolderStructure.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.chk_PreserveOriginalFolderStructure, "Check this box to have all directories and subdirectories recreated in the Output" &
        " Directory. Converted files will be saved into their respective folders.")
        Me.chk_PreserveOriginalFolderStructure.UseVisualStyleBackColor = False
        '
        'chk_SaveLogToFile
        '
        Me.chk_SaveLogToFile.AutoSize = True
        Me.chk_SaveLogToFile.BackColor = System.Drawing.Color.Transparent
        Me.chk_SaveLogToFile.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chk_SaveLogToFile.Checked = True
        Me.chk_SaveLogToFile.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chk_SaveLogToFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chk_SaveLogToFile.ForeColor = System.Drawing.Color.Gainsboro
        Me.chk_SaveLogToFile.Location = New System.Drawing.Point(676, 187)
        Me.chk_SaveLogToFile.Name = "chk_SaveLogToFile"
        Me.chk_SaveLogToFile.Size = New System.Drawing.Size(119, 17)
        Me.chk_SaveLogToFile.TabIndex = 12
        Me.chk_SaveLogToFile.Text = "Save Log to File"
        Me.chk_SaveLogToFile.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.chk_SaveLogToFile, "Check this box to minimize the logging displayed in the application, and, instead" &
        ", write the full log to the Output Directory instead.")
        Me.chk_SaveLogToFile.UseVisualStyleBackColor = False
        '
        'chk_ConvertWordFiles
        '
        Me.chk_ConvertWordFiles.AutoSize = True
        Me.chk_ConvertWordFiles.Checked = True
        Me.chk_ConvertWordFiles.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chk_ConvertWordFiles.Location = New System.Drawing.Point(19, 58)
        Me.chk_ConvertWordFiles.Name = "chk_ConvertWordFiles"
        Me.chk_ConvertWordFiles.Size = New System.Drawing.Size(15, 14)
        Me.chk_ConvertWordFiles.TabIndex = 21
        Me.ToolTip1.SetToolTip(Me.chk_ConvertWordFiles, "Convert Word files")
        Me.chk_ConvertWordFiles.UseVisualStyleBackColor = True
        '
        'chk_ConvertExcelFiles
        '
        Me.chk_ConvertExcelFiles.AutoSize = True
        Me.chk_ConvertExcelFiles.Checked = True
        Me.chk_ConvertExcelFiles.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chk_ConvertExcelFiles.Location = New System.Drawing.Point(19, 84)
        Me.chk_ConvertExcelFiles.Name = "chk_ConvertExcelFiles"
        Me.chk_ConvertExcelFiles.Size = New System.Drawing.Size(15, 14)
        Me.chk_ConvertExcelFiles.TabIndex = 22
        Me.ToolTip1.SetToolTip(Me.chk_ConvertExcelFiles, "Convert Excel files")
        Me.chk_ConvertExcelFiles.UseVisualStyleBackColor = True
        '
        'chk_OnlyCopyTemplateFilesTypes
        '
        Me.chk_OnlyCopyTemplateFilesTypes.AutoSize = True
        Me.chk_OnlyCopyTemplateFilesTypes.BackColor = System.Drawing.Color.Transparent
        Me.chk_OnlyCopyTemplateFilesTypes.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chk_OnlyCopyTemplateFilesTypes.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chk_OnlyCopyTemplateFilesTypes.ForeColor = System.Drawing.Color.Gainsboro
        Me.chk_OnlyCopyTemplateFilesTypes.Location = New System.Drawing.Point(14, 164)
        Me.chk_OnlyCopyTemplateFilesTypes.Name = "chk_OnlyCopyTemplateFilesTypes"
        Me.chk_OnlyCopyTemplateFilesTypes.Size = New System.Drawing.Size(201, 17)
        Me.chk_OnlyCopyTemplateFilesTypes.TabIndex = 23
        Me.chk_OnlyCopyTemplateFilesTypes.Text = "Only Copy Template File Types"
        Me.ToolTip1.SetToolTip(Me.chk_OnlyCopyTemplateFilesTypes, "Check this box to only convert/copy files of the selected template type(s) above " &
        "(Word and/or Excel). All other files will remain, unmodified, in the Source Dire" &
        "ctory.")
        Me.chk_OnlyCopyTemplateFilesTypes.UseVisualStyleBackColor = False
        '
        'txt_ConvertedFilePrefix
        '
        Me.txt_ConvertedFilePrefix.Location = New System.Drawing.Point(595, 137)
        Me.txt_ConvertedFilePrefix.Name = "txt_ConvertedFilePrefix"
        Me.txt_ConvertedFilePrefix.Size = New System.Drawing.Size(149, 20)
        Me.txt_ConvertedFilePrefix.TabIndex = 24
        Me.ToolTip1.SetToolTip(Me.txt_ConvertedFilePrefix, "Specify a prefix that all converted files should include (optional)")
        '
        'chk_OnlyCopyNONTemplateFilesTypes
        '
        Me.chk_OnlyCopyNONTemplateFilesTypes.AutoSize = True
        Me.chk_OnlyCopyNONTemplateFilesTypes.BackColor = System.Drawing.Color.Transparent
        Me.chk_OnlyCopyNONTemplateFilesTypes.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chk_OnlyCopyNONTemplateFilesTypes.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chk_OnlyCopyNONTemplateFilesTypes.ForeColor = System.Drawing.Color.Gainsboro
        Me.chk_OnlyCopyNONTemplateFilesTypes.Location = New System.Drawing.Point(236, 164)
        Me.chk_OnlyCopyNONTemplateFilesTypes.Name = "chk_OnlyCopyNONTemplateFilesTypes"
        Me.chk_OnlyCopyNONTemplateFilesTypes.Size = New System.Drawing.Size(232, 17)
        Me.chk_OnlyCopyNONTemplateFilesTypes.TabIndex = 26
        Me.chk_OnlyCopyNONTemplateFilesTypes.Text = "Only Copy NON-Template File Types"
        Me.ToolTip1.SetToolTip(Me.chk_OnlyCopyNONTemplateFilesTypes, "Check this box to only copy files to the Output directory if they are not a selec" &
        "ted Template option (Word/Excel)")
        Me.chk_OnlyCopyNONTemplateFilesTypes.UseVisualStyleBackColor = False
        '
        'chk_PreserveOriginalCustomProperties
        '
        Me.chk_PreserveOriginalCustomProperties.AutoSize = True
        Me.chk_PreserveOriginalCustomProperties.BackColor = System.Drawing.Color.Transparent
        Me.chk_PreserveOriginalCustomProperties.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chk_PreserveOriginalCustomProperties.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chk_PreserveOriginalCustomProperties.ForeColor = System.Drawing.Color.Gainsboro
        Me.chk_PreserveOriginalCustomProperties.Location = New System.Drawing.Point(489, 164)
        Me.chk_PreserveOriginalCustomProperties.Name = "chk_PreserveOriginalCustomProperties"
        Me.chk_PreserveOriginalCustomProperties.Size = New System.Drawing.Size(229, 17)
        Me.chk_PreserveOriginalCustomProperties.TabIndex = 27
        Me.chk_PreserveOriginalCustomProperties.Text = "Preserve Original Custom Properties"
        Me.ToolTip1.SetToolTip(Me.chk_PreserveOriginalCustomProperties, "Check this box to have the margin sizes in the new Word file match the original")
        Me.chk_PreserveOriginalCustomProperties.UseVisualStyleBackColor = False
        '
        'MenuStrip1
        '
        Me.MenuStrip1.BackColor = System.Drawing.Color.DimGray
        Me.MenuStrip1.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.HelpToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(808, 24)
        Me.MenuStrip1.TabIndex = 16
        Me.MenuStrip1.Text = "mnuMenuStrip"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SaveApplicationSettingsToolStripMenuItem, Me.ToolStripSeparator1, Me.ClearLogToolStripMenuItem, Me.SaveLogToolStripMenuItem, Me.CheckForUpdateToolStripMenuItem, Me.ToolStripSeparator2, Me.ExitToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(37, 20)
        Me.FileToolStripMenuItem.Text = "File"
        '
        'SaveApplicationSettingsToolStripMenuItem
        '
        Me.SaveApplicationSettingsToolStripMenuItem.Name = "SaveApplicationSettingsToolStripMenuItem"
        Me.SaveApplicationSettingsToolStripMenuItem.Size = New System.Drawing.Size(207, 22)
        Me.SaveApplicationSettingsToolStripMenuItem.Text = "Save Application Settings"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(204, 6)
        '
        'ClearLogToolStripMenuItem
        '
        Me.ClearLogToolStripMenuItem.Name = "ClearLogToolStripMenuItem"
        Me.ClearLogToolStripMenuItem.Size = New System.Drawing.Size(207, 22)
        Me.ClearLogToolStripMenuItem.Text = "Clear log contents"
        '
        'SaveLogToolStripMenuItem
        '
        Me.SaveLogToolStripMenuItem.Name = "SaveLogToolStripMenuItem"
        Me.SaveLogToolStripMenuItem.Size = New System.Drawing.Size(207, 22)
        Me.SaveLogToolStripMenuItem.Text = "Save log contents"
        '
        'CheckForUpdateToolStripMenuItem
        '
        Me.CheckForUpdateToolStripMenuItem.Name = "CheckForUpdateToolStripMenuItem"
        Me.CheckForUpdateToolStripMenuItem.Size = New System.Drawing.Size(207, 22)
        Me.CheckForUpdateToolStripMenuItem.Text = "Check for update"
        Me.CheckForUpdateToolStripMenuItem.Visible = False
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(204, 6)
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(207, 22)
        Me.ExitToolStripMenuItem.Text = "Exit"
        '
        'HelpToolStripMenuItem
        '
        Me.HelpToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ActivateProductLicenseToolStripMenuItem, Me.FrequentlyAskedQuestionsToolStripMenuItem})
        Me.HelpToolStripMenuItem.Enabled = False
        Me.HelpToolStripMenuItem.Name = "HelpToolStripMenuItem"
        Me.HelpToolStripMenuItem.Size = New System.Drawing.Size(44, 20)
        Me.HelpToolStripMenuItem.Text = "Help"
        Me.HelpToolStripMenuItem.Visible = False
        '
        'ActivateProductLicenseToolStripMenuItem
        '
        Me.ActivateProductLicenseToolStripMenuItem.Name = "ActivateProductLicenseToolStripMenuItem"
        Me.ActivateProductLicenseToolStripMenuItem.Size = New System.Drawing.Size(221, 22)
        Me.ActivateProductLicenseToolStripMenuItem.Text = "Activate Product License"
        '
        'FrequentlyAskedQuestionsToolStripMenuItem
        '
        Me.FrequentlyAskedQuestionsToolStripMenuItem.Name = "FrequentlyAskedQuestionsToolStripMenuItem"
        Me.FrequentlyAskedQuestionsToolStripMenuItem.Size = New System.Drawing.Size(221, 22)
        Me.FrequentlyAskedQuestionsToolStripMenuItem.Text = "Frequently Asked Questions"
        '
        'lbl_ExcelTemplateFile
        '
        Me.lbl_ExcelTemplateFile.AutoSize = True
        Me.lbl_ExcelTemplateFile.BackColor = System.Drawing.Color.Transparent
        Me.lbl_ExcelTemplateFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_ExcelTemplateFile.ForeColor = System.Drawing.Color.Gainsboro
        Me.lbl_ExcelTemplateFile.Location = New System.Drawing.Point(40, 87)
        Me.lbl_ExcelTemplateFile.Name = "lbl_ExcelTemplateFile"
        Me.lbl_ExcelTemplateFile.Size = New System.Drawing.Size(122, 13)
        Me.lbl_ExcelTemplateFile.TabIndex = 20
        Me.lbl_ExcelTemplateFile.Text = "Excel Template File:"
        '
        'lbl_ConvertedFilePrefix
        '
        Me.lbl_ConvertedFilePrefix.AutoSize = True
        Me.lbl_ConvertedFilePrefix.BackColor = System.Drawing.Color.Transparent
        Me.lbl_ConvertedFilePrefix.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_ConvertedFilePrefix.ForeColor = System.Drawing.Color.Gainsboro
        Me.lbl_ConvertedFilePrefix.Location = New System.Drawing.Point(460, 140)
        Me.lbl_ConvertedFilePrefix.Name = "lbl_ConvertedFilePrefix"
        Me.lbl_ConvertedFilePrefix.Size = New System.Drawing.Size(129, 13)
        Me.lbl_ConvertedFilePrefix.TabIndex = 25
        Me.lbl_ConvertedFilePrefix.Text = "Converted File Prefix:"
        '
        'frm_Main
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImage = Global.Document_Converter.My.Resources.Resources.Gradient_DarkTop
        Me.ClientSize = New System.Drawing.Size(808, 479)
        Me.Controls.Add(Me.chk_PreserveOriginalCustomProperties)
        Me.Controls.Add(Me.chk_OnlyCopyNONTemplateFilesTypes)
        Me.Controls.Add(Me.txt_ConvertedFilePrefix)
        Me.Controls.Add(Me.lbl_ConvertedFilePrefix)
        Me.Controls.Add(Me.chk_OnlyCopyTemplateFilesTypes)
        Me.Controls.Add(Me.chk_ConvertExcelFiles)
        Me.Controls.Add(Me.chk_ConvertWordFiles)
        Me.Controls.Add(Me.btn_ExcelTemplateFile)
        Me.Controls.Add(Me.txt_ExcelTemplateFile)
        Me.Controls.Add(Me.lbl_ExcelTemplateFile)
        Me.Controls.Add(Me.chk_SaveLogToFile)
        Me.Controls.Add(Me.chk_PreserveOriginalFolderStructure)
        Me.Controls.Add(Me.chk_PreserveOriginalFileDates)
        Me.Controls.Add(Me.chk_PreserveOriginalMargins)
        Me.Controls.Add(Me.btn_ChooseSourceDirectory)
        Me.Controls.Add(Me.btn_ChooseOutputDirectory)
        Me.Controls.Add(Me.btn_WordTemplateFile)
        Me.Controls.Add(Me.btn_Stop)
        Me.Controls.Add(Me.txt_StartAtRecordNumber)
        Me.Controls.Add(Me.lbl_StartAtRecordNumber)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Controls.Add(Me.txt_Log)
        Me.Controls.Add(Me.btn_Convert)
        Me.Controls.Add(Me.txt_OutputDirectory)
        Me.Controls.Add(Me.lbl_OutputDirectory)
        Me.Controls.Add(Me.txt_WordTemplateFile)
        Me.Controls.Add(Me.lbl_WordTemplateFile)
        Me.Controls.Add(Me.txt_SourceDirectory)
        Me.Controls.Add(Me.lbl_SourceDirectory)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "frm_Main"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Document Template Converter"
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lbl_SourceDirectory As System.Windows.Forms.Label
    Friend WithEvents txt_SourceDirectory As System.Windows.Forms.TextBox
    Friend WithEvents txt_WordTemplateFile As System.Windows.Forms.TextBox
    Friend WithEvents lbl_WordTemplateFile As System.Windows.Forms.Label
    Friend WithEvents txt_OutputDirectory As System.Windows.Forms.TextBox
    Friend WithEvents lbl_OutputDirectory As System.Windows.Forms.Label
    Friend WithEvents btn_Convert As System.Windows.Forms.Button
    Friend WithEvents txt_Log As System.Windows.Forms.TextBox
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents txt_StartAtRecordNumber As System.Windows.Forms.TextBox
    Friend WithEvents lbl_StartAtRecordNumber As System.Windows.Forms.Label
    Friend WithEvents prog_Progress As System.Windows.Forms.ToolStripProgressBar
    Friend WithEvents lbl_Status As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents btn_Stop As System.Windows.Forms.Button
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents FileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CheckForUpdateToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SaveLogToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents sfd_SaveFileDialog As System.Windows.Forms.SaveFileDialog
    Friend WithEvents btn_WordTemplateFile As System.Windows.Forms.Button
    Friend WithEvents btn_ChooseOutputDirectory As System.Windows.Forms.Button
    Friend WithEvents btn_ChooseSourceDirectory As System.Windows.Forms.Button
    Friend WithEvents ClearLogToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents chk_PreserveOriginalMargins As System.Windows.Forms.CheckBox
    Friend WithEvents chk_PreserveOriginalFileDates As System.Windows.Forms.CheckBox
    Friend WithEvents chk_PreserveOriginalFolderStructure As System.Windows.Forms.CheckBox
    Friend WithEvents SaveApplicationSettingsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As ToolStripSeparator
    Friend WithEvents ToolStripSeparator2 As ToolStripSeparator
    Friend WithEvents HelpToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ActivateProductLicenseToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents FrequentlyAskedQuestionsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents chk_SaveLogToFile As CheckBox
    Friend WithEvents btn_ExcelTemplateFile As Button
    Friend WithEvents txt_ExcelTemplateFile As TextBox
    Friend WithEvents lbl_ExcelTemplateFile As Label
    Friend WithEvents chk_ConvertWordFiles As CheckBox
    Friend WithEvents chk_ConvertExcelFiles As CheckBox
    Friend WithEvents chk_OnlyCopyTemplateFilesTypes As CheckBox
    Friend WithEvents txt_ConvertedFilePrefix As TextBox
    Friend WithEvents lbl_ConvertedFilePrefix As Label
    Friend WithEvents chk_OnlyCopyNONTemplateFilesTypes As CheckBox
    Friend WithEvents chk_PreserveOriginalCustomProperties As CheckBox
End Class
