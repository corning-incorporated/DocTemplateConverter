<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_Activate
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frm_Activate))
        Me.btn_Activate = New System.Windows.Forms.Button()
        Me.btn_Cancel = New System.Windows.Forms.Button()
        Me.grp_LicenseCode = New System.Windows.Forms.GroupBox()
        Me.txt_LicenseCode = New System.Windows.Forms.MaskedTextBox()
        Me.lbl_LicenseCode = New System.Windows.Forms.Label()
        Me.lbl_Instructions = New System.Windows.Forms.Label()
        Me.grp_LicenseCode.SuspendLayout()
        Me.SuspendLayout()
        '
        'btn_Activate
        '
        Me.btn_Activate.Location = New System.Drawing.Point(58, 103)
        Me.btn_Activate.Name = "btn_Activate"
        Me.btn_Activate.Size = New System.Drawing.Size(60, 23)
        Me.btn_Activate.TabIndex = 3
        Me.btn_Activate.Text = "Activate"
        Me.btn_Activate.UseVisualStyleBackColor = True
        '
        'btn_Cancel
        '
        Me.btn_Cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btn_Cancel.Location = New System.Drawing.Point(301, 103)
        Me.btn_Cancel.Name = "btn_Cancel"
        Me.btn_Cancel.Size = New System.Drawing.Size(60, 23)
        Me.btn_Cancel.TabIndex = 4
        Me.btn_Cancel.Text = "Cancel"
        Me.btn_Cancel.UseVisualStyleBackColor = True
        '
        'grp_LicenseCode
        '
        Me.grp_LicenseCode.BackColor = System.Drawing.Color.Transparent
        Me.grp_LicenseCode.Controls.Add(Me.txt_LicenseCode)
        Me.grp_LicenseCode.Controls.Add(Me.lbl_LicenseCode)
        Me.grp_LicenseCode.Location = New System.Drawing.Point(58, 33)
        Me.grp_LicenseCode.Name = "grp_LicenseCode"
        Me.grp_LicenseCode.Size = New System.Drawing.Size(303, 64)
        Me.grp_LicenseCode.TabIndex = 5
        Me.grp_LicenseCode.TabStop = False
        '
        'txt_LicenseCode
        '
        Me.txt_LicenseCode.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txt_LicenseCode.Location = New System.Drawing.Point(127, 19)
        Me.txt_LicenseCode.Mask = "AAAA-000-AAAA"
        Me.txt_LicenseCode.Name = "txt_LicenseCode"
        Me.txt_LicenseCode.Size = New System.Drawing.Size(123, 26)
        Me.txt_LicenseCode.TabIndex = 3
        '
        'lbl_LicenseCode
        '
        Me.lbl_LicenseCode.AutoSize = True
        Me.lbl_LicenseCode.BackColor = System.Drawing.Color.Transparent
        Me.lbl_LicenseCode.ForeColor = System.Drawing.Color.LightGray
        Me.lbl_LicenseCode.Location = New System.Drawing.Point(46, 27)
        Me.lbl_LicenseCode.Name = "lbl_LicenseCode"
        Me.lbl_LicenseCode.Size = New System.Drawing.Size(75, 13)
        Me.lbl_LicenseCode.TabIndex = 2
        Me.lbl_LicenseCode.Text = "License Code:"
        '
        'lbl_Instructions
        '
        Me.lbl_Instructions.AutoSize = True
        Me.lbl_Instructions.BackColor = System.Drawing.Color.Transparent
        Me.lbl_Instructions.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_Instructions.ForeColor = System.Drawing.Color.LightGray
        Me.lbl_Instructions.Location = New System.Drawing.Point(54, 14)
        Me.lbl_Instructions.Name = "lbl_Instructions"
        Me.lbl_Instructions.Size = New System.Drawing.Size(315, 13)
        Me.lbl_Instructions.TabIndex = 6
        Me.lbl_Instructions.Text = "Please enter your license key to activate your license."
        '
        'frm_Activate
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImage = Global.Document_Converter.My.Resources.Resources.Gradient_DarkTop
        Me.CancelButton = Me.btn_Cancel
        Me.ClientSize = New System.Drawing.Size(425, 138)
        Me.Controls.Add(Me.lbl_Instructions)
        Me.Controls.Add(Me.grp_LicenseCode)
        Me.Controls.Add(Me.btn_Cancel)
        Me.Controls.Add(Me.btn_Activate)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frm_Activate"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Activate Product License..."
        Me.TopMost = True
        Me.grp_LicenseCode.ResumeLayout(False)
        Me.grp_LicenseCode.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btn_Activate As Button
    Friend WithEvents btn_Cancel As Button
    Friend WithEvents grp_LicenseCode As GroupBox
    Friend WithEvents txt_LicenseCode As MaskedTextBox
    Friend WithEvents lbl_LicenseCode As Label
    Friend WithEvents lbl_Instructions As Label
End Class
