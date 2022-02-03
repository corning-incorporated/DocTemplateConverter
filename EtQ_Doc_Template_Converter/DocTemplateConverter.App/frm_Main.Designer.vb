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
        Me.txt_TemplateFile = New System.Windows.Forms.TextBox()
        Me.lbl_TemplateFile = New System.Windows.Forms.Label()
        Me.txt_OutputDirectory = New System.Windows.Forms.TextBox()
        Me.lbl_OutputDirectory = New System.Windows.Forms.Label()
        Me.btn_Convert = New System.Windows.Forms.Button()
        Me.txt_Log = New System.Windows.Forms.TextBox()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.prog_Progress = New System.Windows.Forms.ToolStripProgressBar()
        Me.lbl_Status = New System.Windows.Forms.ToolStripStatusLabel()
        Me.txt_ConvertedFilePrefix = New System.Windows.Forms.TextBox()
        Me.lbl_ConvertedFilePrefix = New System.Windows.Forms.Label()
        Me.btn_Stop = New System.Windows.Forms.Button()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.btn_SelectTemplateFile = New System.Windows.Forms.Button()
        Me.btn_ChooseOutputDirectory = New System.Windows.Forms.Button()
        Me.btn_ChooseSourceDirectory = New System.Windows.Forms.Button()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ClearLogToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SaveLogToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CheckForUpdateToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.sfd_SaveFileDialog = New System.Windows.Forms.SaveFileDialog()
        Me.chk_PreserveOriginalMargins = New System.Windows.Forms.CheckBox()
        Me.chk_PreserveOriginalFileDates = New System.Windows.Forms.CheckBox()
        Me.chk_PreserveOriginalFolderStructure = New System.Windows.Forms.CheckBox()
        Me.SaveApplicationSettingsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.StatusStrip1.SuspendLayout()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'lbl_SourceDirectory
        '
        Me.lbl_SourceDirectory.AutoSize = True
        Me.lbl_SourceDirectory.BackColor = System.Drawing.Color.Transparent
        Me.lbl_SourceDirectory.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_SourceDirectory.Location = New System.Drawing.Point(35, 32)
        Me.lbl_SourceDirectory.Name = "lbl_SourceDirectory"
        Me.lbl_SourceDirectory.Size = New System.Drawing.Size(106, 13)
        Me.lbl_SourceDirectory.TabIndex = 0
        Me.lbl_SourceDirectory.Text = "Source Directory:"
        '
        'txt_SourceDirectory
        '
        Me.txt_SourceDirectory.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txt_SourceDirectory.Location = New System.Drawing.Point(143, 29)
        Me.txt_SourceDirectory.Name = "txt_SourceDirectory"
        Me.txt_SourceDirectory.Size = New System.Drawing.Size(511, 20)
        Me.txt_SourceDirectory.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.txt_SourceDirectory, "Double-click to select a source directory")
        '
        'txt_TemplateFile
        '
        Me.txt_TemplateFile.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txt_TemplateFile.Location = New System.Drawing.Point(143, 55)
        Me.txt_TemplateFile.Name = "txt_TemplateFile"
        Me.txt_TemplateFile.Size = New System.Drawing.Size(511, 20)
        Me.txt_TemplateFile.TabIndex = 3
        Me.ToolTip1.SetToolTip(Me.txt_TemplateFile, "Double-click to select a template file")
        '
        'lbl_TemplateFile
        '
        Me.lbl_TemplateFile.AutoSize = True
        Me.lbl_TemplateFile.BackColor = System.Drawing.Color.Transparent
        Me.lbl_TemplateFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_TemplateFile.Location = New System.Drawing.Point(54, 58)
        Me.lbl_TemplateFile.Name = "lbl_TemplateFile"
        Me.lbl_TemplateFile.Size = New System.Drawing.Size(87, 13)
        Me.lbl_TemplateFile.TabIndex = 2
        Me.lbl_TemplateFile.Text = "Template File:"
        '
        'txt_OutputDirectory
        '
        Me.txt_OutputDirectory.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txt_OutputDirectory.Location = New System.Drawing.Point(143, 81)
        Me.txt_OutputDirectory.Name = "txt_OutputDirectory"
        Me.txt_OutputDirectory.Size = New System.Drawing.Size(511, 20)
        Me.txt_OutputDirectory.TabIndex = 5
        Me.ToolTip1.SetToolTip(Me.txt_OutputDirectory, "Double-click to select an output directory")
        '
        'lbl_OutputDirectory
        '
        Me.lbl_OutputDirectory.AutoSize = True
        Me.lbl_OutputDirectory.BackColor = System.Drawing.Color.Transparent
        Me.lbl_OutputDirectory.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_OutputDirectory.Location = New System.Drawing.Point(37, 84)
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
        Me.btn_Convert.Location = New System.Drawing.Point(12, 154)
        Me.btn_Convert.Name = "btn_Convert"
        Me.btn_Convert.Size = New System.Drawing.Size(642, 46)
        Me.btn_Convert.TabIndex = 6
        Me.btn_Convert.Text = "CONVERT!"
        Me.ToolTip1.SetToolTip(Me.btn_Convert, "Convert all documents in the Source Directory and all subfolders")
        Me.btn_Convert.UseVisualStyleBackColor = True
        '
        'txt_Log
        '
        Me.txt_Log.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txt_Log.Location = New System.Drawing.Point(0, 206)
        Me.txt_Log.Multiline = True
        Me.txt_Log.Name = "txt_Log"
        Me.txt_Log.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txt_Log.Size = New System.Drawing.Size(718, 251)
        Me.txt_Log.TabIndex = 7
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.prog_Progress, Me.lbl_Status})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 457)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(718, 22)
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
        'txt_ConvertedFilePrefix
        '
        Me.txt_ConvertedFilePrefix.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txt_ConvertedFilePrefix.Location = New System.Drawing.Point(143, 107)
        Me.txt_ConvertedFilePrefix.Name = "txt_ConvertedFilePrefix"
        Me.txt_ConvertedFilePrefix.Size = New System.Drawing.Size(511, 20)
        Me.txt_ConvertedFilePrefix.TabIndex = 10
        Me.ToolTip1.SetToolTip(Me.txt_ConvertedFilePrefix, "Specify a prefix that all converted files should include (optional)")
        '
        'lbl_ConvertedFilePrefix
        '
        Me.lbl_ConvertedFilePrefix.AutoSize = True
        Me.lbl_ConvertedFilePrefix.BackColor = System.Drawing.Color.Transparent
        Me.lbl_ConvertedFilePrefix.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_ConvertedFilePrefix.Location = New System.Drawing.Point(12, 110)
        Me.lbl_ConvertedFilePrefix.Name = "lbl_ConvertedFilePrefix"
        Me.lbl_ConvertedFilePrefix.Size = New System.Drawing.Size(129, 13)
        Me.lbl_ConvertedFilePrefix.TabIndex = 9
        Me.lbl_ConvertedFilePrefix.Text = "Converted File Prefix:"
        '
        'btn_Stop
        '
        Me.btn_Stop.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_Stop.Image = Global.Document_Converter.My.Resources.Resources.red_square
        Me.btn_Stop.Location = New System.Drawing.Point(661, 155)
        Me.btn_Stop.Name = "btn_Stop"
        Me.btn_Stop.Size = New System.Drawing.Size(45, 45)
        Me.btn_Stop.TabIndex = 11
        Me.ToolTip1.SetToolTip(Me.btn_Stop, "Stop Converter")
        Me.btn_Stop.UseVisualStyleBackColor = True
        '
        'btn_SelectTemplateFile
        '
        Me.btn_SelectTemplateFile.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_SelectTemplateFile.Image = Global.Document_Converter.My.Resources.Resources.word
        Me.btn_SelectTemplateFile.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btn_SelectTemplateFile.Location = New System.Drawing.Point(661, 52)
        Me.btn_SelectTemplateFile.Name = "btn_SelectTemplateFile"
        Me.btn_SelectTemplateFile.Size = New System.Drawing.Size(45, 24)
        Me.btn_SelectTemplateFile.TabIndex = 13
        Me.ToolTip1.SetToolTip(Me.btn_SelectTemplateFile, "Choose template file")
        Me.btn_SelectTemplateFile.UseVisualStyleBackColor = True
        '
        'btn_ChooseOutputDirectory
        '
        Me.btn_ChooseOutputDirectory.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_ChooseOutputDirectory.Image = Global.Document_Converter.My.Resources.Resources.folder_next
        Me.btn_ChooseOutputDirectory.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btn_ChooseOutputDirectory.Location = New System.Drawing.Point(661, 78)
        Me.btn_ChooseOutputDirectory.Name = "btn_ChooseOutputDirectory"
        Me.btn_ChooseOutputDirectory.Size = New System.Drawing.Size(45, 24)
        Me.btn_ChooseOutputDirectory.TabIndex = 14
        Me.ToolTip1.SetToolTip(Me.btn_ChooseOutputDirectory, "Choose output directory (location of converted files)")
        Me.btn_ChooseOutputDirectory.UseVisualStyleBackColor = True
        '
        'btn_ChooseSourceDirectory
        '
        Me.btn_ChooseSourceDirectory.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_ChooseSourceDirectory.Image = Global.Document_Converter.My.Resources.Resources.folder_process
        Me.btn_ChooseSourceDirectory.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btn_ChooseSourceDirectory.Location = New System.Drawing.Point(660, 26)
        Me.btn_ChooseSourceDirectory.Name = "btn_ChooseSourceDirectory"
        Me.btn_ChooseSourceDirectory.Size = New System.Drawing.Size(45, 24)
        Me.btn_ChooseSourceDirectory.TabIndex = 15
        Me.ToolTip1.SetToolTip(Me.btn_ChooseSourceDirectory, "Choose ""source"" directory (location of original files)")
        Me.btn_ChooseSourceDirectory.UseVisualStyleBackColor = True
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(718, 24)
        Me.MenuStrip1.TabIndex = 12
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SaveApplicationSettingsToolStripMenuItem, Me.ToolStripSeparator1, Me.ClearLogToolStripMenuItem, Me.SaveLogToolStripMenuItem, Me.CheckForUpdateToolStripMenuItem, Me.ToolStripSeparator2, Me.ExitToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(37, 20)
        Me.FileToolStripMenuItem.Text = "File"
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
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(207, 22)
        Me.ExitToolStripMenuItem.Text = "Exit"
        '
        'chk_PreserveOriginalMargins
        '
        Me.chk_PreserveOriginalMargins.AutoSize = True
        Me.chk_PreserveOriginalMargins.BackColor = System.Drawing.Color.Transparent
        Me.chk_PreserveOriginalMargins.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chk_PreserveOriginalMargins.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chk_PreserveOriginalMargins.Location = New System.Drawing.Point(12, 133)
        Me.chk_PreserveOriginalMargins.Name = "chk_PreserveOriginalMargins"
        Me.chk_PreserveOriginalMargins.Size = New System.Drawing.Size(199, 17)
        Me.chk_PreserveOriginalMargins.TabIndex = 16
        Me.chk_PreserveOriginalMargins.Text = "Preserve Original Margin Sizes"
        Me.chk_PreserveOriginalMargins.UseVisualStyleBackColor = False
        '
        'chk_PreserveOriginalFileDates
        '
        Me.chk_PreserveOriginalFileDates.AutoSize = True
        Me.chk_PreserveOriginalFileDates.BackColor = System.Drawing.Color.Transparent
        Me.chk_PreserveOriginalFileDates.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chk_PreserveOriginalFileDates.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chk_PreserveOriginalFileDates.Location = New System.Drawing.Point(232, 133)
        Me.chk_PreserveOriginalFileDates.Name = "chk_PreserveOriginalFileDates"
        Me.chk_PreserveOriginalFileDates.Size = New System.Drawing.Size(184, 17)
        Me.chk_PreserveOriginalFileDates.TabIndex = 17
        Me.chk_PreserveOriginalFileDates.Text = "Preserve Original File Dates"
        Me.chk_PreserveOriginalFileDates.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chk_PreserveOriginalFileDates.UseVisualStyleBackColor = False
        '
        'chk_PreserveOriginalFolderStructure
        '
        Me.chk_PreserveOriginalFolderStructure.AutoSize = True
        Me.chk_PreserveOriginalFolderStructure.BackColor = System.Drawing.Color.Transparent
        Me.chk_PreserveOriginalFolderStructure.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chk_PreserveOriginalFolderStructure.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chk_PreserveOriginalFolderStructure.Location = New System.Drawing.Point(437, 133)
        Me.chk_PreserveOriginalFolderStructure.Name = "chk_PreserveOriginalFolderStructure"
        Me.chk_PreserveOriginalFolderStructure.Size = New System.Drawing.Size(218, 17)
        Me.chk_PreserveOriginalFolderStructure.TabIndex = 18
        Me.chk_PreserveOriginalFolderStructure.Text = "Preserve Original Folder Structure"
        Me.chk_PreserveOriginalFolderStructure.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chk_PreserveOriginalFolderStructure.UseVisualStyleBackColor = False
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
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(204, 6)
        '
        'frm_Main
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImage = Global.Document_Converter.My.Resources.Resources.Gradient
        Me.ClientSize = New System.Drawing.Size(718, 479)
        Me.Controls.Add(Me.chk_PreserveOriginalFolderStructure)
        Me.Controls.Add(Me.chk_PreserveOriginalFileDates)
        Me.Controls.Add(Me.chk_PreserveOriginalMargins)
        Me.Controls.Add(Me.btn_ChooseSourceDirectory)
        Me.Controls.Add(Me.btn_ChooseOutputDirectory)
        Me.Controls.Add(Me.btn_SelectTemplateFile)
        Me.Controls.Add(Me.btn_Stop)
        Me.Controls.Add(Me.txt_ConvertedFilePrefix)
        Me.Controls.Add(Me.lbl_ConvertedFilePrefix)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Controls.Add(Me.txt_Log)
        Me.Controls.Add(Me.btn_Convert)
        Me.Controls.Add(Me.txt_OutputDirectory)
        Me.Controls.Add(Me.lbl_OutputDirectory)
        Me.Controls.Add(Me.txt_TemplateFile)
        Me.Controls.Add(Me.lbl_TemplateFile)
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
    Friend WithEvents txt_TemplateFile As System.Windows.Forms.TextBox
    Friend WithEvents lbl_TemplateFile As System.Windows.Forms.Label
    Friend WithEvents txt_OutputDirectory As System.Windows.Forms.TextBox
    Friend WithEvents lbl_OutputDirectory As System.Windows.Forms.Label
    Friend WithEvents btn_Convert As System.Windows.Forms.Button
    Friend WithEvents txt_Log As System.Windows.Forms.TextBox
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents txt_ConvertedFilePrefix As System.Windows.Forms.TextBox
    Friend WithEvents lbl_ConvertedFilePrefix As System.Windows.Forms.Label
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
    Friend WithEvents btn_SelectTemplateFile As System.Windows.Forms.Button
    Friend WithEvents btn_ChooseOutputDirectory As System.Windows.Forms.Button
    Friend WithEvents btn_ChooseSourceDirectory As System.Windows.Forms.Button
    Friend WithEvents ClearLogToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents chk_PreserveOriginalMargins As System.Windows.Forms.CheckBox
    Friend WithEvents chk_PreserveOriginalFileDates As System.Windows.Forms.CheckBox
    Friend WithEvents chk_PreserveOriginalFolderStructure As System.Windows.Forms.CheckBox
    Friend WithEvents SaveApplicationSettingsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As ToolStripSeparator
    Friend WithEvents ToolStripSeparator2 As ToolStripSeparator
End Class
