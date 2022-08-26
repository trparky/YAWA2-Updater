<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.lblWebSiteVersion = New System.Windows.Forms.Label()
        Me.lblYourVersion = New System.Windows.Forms.Label()
        Me.lblYourCustomEntries = New System.Windows.Forms.Label()
        Me.TxtCustomEntries = New System.Windows.Forms.TextBox()
        Me.chkTrim = New System.Windows.Forms.CheckBox()
        Me.chkLoadAtUserStartup = New System.Windows.Forms.CheckBox()
        Me.chkNotifyAfterUpdateatLogon = New System.Windows.Forms.CheckBox()
        Me.lblUpdateNeededOrNot = New System.Windows.Forms.Label()
        Me.chkMobileMode = New System.Windows.Forms.CheckBox()
        Me.lblAdmin = New System.Windows.Forms.Label()
        Me.btnAbout = New System.Windows.Forms.Button()
        Me.btnCheckForUpdates = New System.Windows.Forms.Button()
        Me.btnReDownload = New System.Windows.Forms.Button()
        Me.btnTrim = New System.Windows.Forms.Button()
        Me.btnApplyNewINIFile = New System.Windows.Forms.Button()
        Me.ChkSleepOnSilentStartup = New System.Windows.Forms.CheckBox()
        Me.lblSeconds = New System.Windows.Forms.Label()
        Me.txtSeconds = New System.Windows.Forms.TextBox()
        Me.btnSaveSeconds = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'lblWebSiteVersion
        '
        Me.lblWebSiteVersion.AutoSize = True
        Me.lblWebSiteVersion.Location = New System.Drawing.Point(4, 25)
        Me.lblWebSiteVersion.Name = "lblWebSiteVersion"
        Me.lblWebSiteVersion.Size = New System.Drawing.Size(271, 13)
        Me.lblWebSiteVersion.TabIndex = 0
        Me.lblWebSiteVersion.Text = "Web Site WinApp2.ini Version: (Loading... Please Wait.)"
        '
        'lblYourVersion
        '
        Me.lblYourVersion.AutoSize = True
        Me.lblYourVersion.Location = New System.Drawing.Point(4, 9)
        Me.lblYourVersion.Name = "lblYourVersion"
        Me.lblYourVersion.Size = New System.Drawing.Size(130, 13)
        Me.lblYourVersion.TabIndex = 1
        Me.lblYourVersion.Text = "Your WinApp2.ini Version:"
        '
        'lblYourCustomEntries
        '
        Me.lblYourCustomEntries.AutoSize = True
        Me.lblYourCustomEntries.Location = New System.Drawing.Point(4, 73)
        Me.lblYourCustomEntries.Name = "lblYourCustomEntries"
        Me.lblYourCustomEntries.Size = New System.Drawing.Size(102, 13)
        Me.lblYourCustomEntries.TabIndex = 2
        Me.lblYourCustomEntries.Text = "Your Custom Entries"
        '
        'TxtCustomEntries
        '
        Me.TxtCustomEntries.Location = New System.Drawing.Point(7, 89)
        Me.TxtCustomEntries.Multiline = True
        Me.TxtCustomEntries.Name = "TxtCustomEntries"
        Me.TxtCustomEntries.Size = New System.Drawing.Size(671, 291)
        Me.TxtCustomEntries.TabIndex = 1
        '
        'chkTrim
        '
        Me.chkTrim.AutoSize = True
        Me.chkTrim.Location = New System.Drawing.Point(283, 8)
        Me.chkTrim.Name = "chkTrim"
        Me.chkTrim.Size = New System.Drawing.Size(158, 17)
        Me.chkTrim.TabIndex = 5
        Me.chkTrim.Text = "Trim INI File After Download"
        Me.chkTrim.UseVisualStyleBackColor = True
        '
        'chkLoadAtUserStartup
        '
        Me.chkLoadAtUserStartup.AutoSize = True
        Me.chkLoadAtUserStartup.Location = New System.Drawing.Point(267, 383)
        Me.chkLoadAtUserStartup.Name = "chkLoadAtUserStartup"
        Me.chkLoadAtUserStartup.Size = New System.Drawing.Size(258, 30)
        Me.chkLoadAtUserStartup.TabIndex = 6
        Me.chkLoadAtUserStartup.Text = "Check for new WinApp2.ini File at User Logon" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(UAC Friendly, you will not receive" &
    " a UAC prompt)"
        Me.chkLoadAtUserStartup.UseVisualStyleBackColor = True
        '
        'chkNotifyAfterUpdateatLogon
        '
        Me.chkNotifyAfterUpdateatLogon.AutoSize = True
        Me.chkNotifyAfterUpdateatLogon.Location = New System.Drawing.Point(100, 389)
        Me.chkNotifyAfterUpdateatLogon.Name = "chkNotifyAfterUpdateatLogon"
        Me.chkNotifyAfterUpdateatLogon.Size = New System.Drawing.Size(161, 17)
        Me.chkNotifyAfterUpdateatLogon.TabIndex = 9
        Me.chkNotifyAfterUpdateatLogon.Text = "Notify After Update at Logon"
        Me.chkNotifyAfterUpdateatLogon.UseVisualStyleBackColor = True
        '
        'lblUpdateNeededOrNot
        '
        Me.lblUpdateNeededOrNot.AutoSize = True
        Me.lblUpdateNeededOrNot.Location = New System.Drawing.Point(498, 9)
        Me.lblUpdateNeededOrNot.Name = "lblUpdateNeededOrNot"
        Me.lblUpdateNeededOrNot.Size = New System.Drawing.Size(109, 13)
        Me.lblUpdateNeededOrNot.TabIndex = 10
        Me.lblUpdateNeededOrNot.Text = "Update NOT Needed"
        '
        'chkMobileMode
        '
        Me.chkMobileMode.AutoSize = True
        Me.chkMobileMode.Location = New System.Drawing.Point(7, 390)
        Me.chkMobileMode.Name = "chkMobileMode"
        Me.chkMobileMode.Size = New System.Drawing.Size(87, 17)
        Me.chkMobileMode.TabIndex = 11
        Me.chkMobileMode.Text = "Mobile Mode"
        Me.chkMobileMode.UseVisualStyleBackColor = True
        '
        'lblAdmin
        '
        Me.lblAdmin.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAdmin.Image = Global.YAWA2_Updater.My.Resources.Resources.UAC
        Me.lblAdmin.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblAdmin.Location = New System.Drawing.Point(4, 41)
        Me.lblAdmin.Name = "lblAdmin"
        Me.lblAdmin.Size = New System.Drawing.Size(200, 20)
        Me.lblAdmin.TabIndex = 12
        Me.lblAdmin.Text = "WARNING! Running as Administrator"
        Me.lblAdmin.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblAdmin.Visible = False
        '
        'btnAbout
        '
        Me.btnAbout.Image = Global.YAWA2_Updater.My.Resources.Resources.info_blue
        Me.btnAbout.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnAbout.Location = New System.Drawing.Point(131, 415)
        Me.btnAbout.Name = "btnAbout"
        Me.btnAbout.Size = New System.Drawing.Size(58, 23)
        Me.btnAbout.TabIndex = 8
        Me.btnAbout.Text = "About"
        Me.btnAbout.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnAbout.UseVisualStyleBackColor = True
        '
        'btnCheckForUpdates
        '
        Me.btnCheckForUpdates.Image = Global.YAWA2_Updater.My.Resources.Resources.refresh
        Me.btnCheckForUpdates.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnCheckForUpdates.Location = New System.Drawing.Point(7, 415)
        Me.btnCheckForUpdates.Name = "btnCheckForUpdates"
        Me.btnCheckForUpdates.Size = New System.Drawing.Size(118, 23)
        Me.btnCheckForUpdates.TabIndex = 7
        Me.btnCheckForUpdates.Text = "Check for Updates"
        Me.btnCheckForUpdates.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnCheckForUpdates.UseVisualStyleBackColor = True
        '
        'btnReDownload
        '
        Me.btnReDownload.Image = Global.YAWA2_Updater.My.Resources.Resources.arrow_down
        Me.btnReDownload.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnReDownload.Location = New System.Drawing.Point(492, 31)
        Me.btnReDownload.Name = "btnReDownload"
        Me.btnReDownload.Size = New System.Drawing.Size(186, 23)
        Me.btnReDownload.TabIndex = 4
        Me.btnReDownload.Text = "Force Re-Download INI File"
        Me.btnReDownload.UseVisualStyleBackColor = True
        '
        'btnTrim
        '
        Me.btnTrim.Image = Global.YAWA2_Updater.My.Resources.Resources.scissors_blue
        Me.btnTrim.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnTrim.Location = New System.Drawing.Point(283, 60)
        Me.btnTrim.Name = "btnTrim"
        Me.btnTrim.Size = New System.Drawing.Size(395, 23)
        Me.btnTrim.TabIndex = 3
        Me.btnTrim.Text = "Trim INI File to your system's software configuration"
        Me.btnTrim.UseVisualStyleBackColor = True
        '
        'btnApplyNewINIFile
        '
        Me.btnApplyNewINIFile.Image = Global.YAWA2_Updater.My.Resources.Resources.save
        Me.btnApplyNewINIFile.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnApplyNewINIFile.Location = New System.Drawing.Point(283, 31)
        Me.btnApplyNewINIFile.Name = "btnApplyNewINIFile"
        Me.btnApplyNewINIFile.Size = New System.Drawing.Size(203, 23)
        Me.btnApplyNewINIFile.TabIndex = 0
        Me.btnApplyNewINIFile.Text = "Apply New WinApp2.ini File Version"
        Me.btnApplyNewINIFile.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnApplyNewINIFile.UseVisualStyleBackColor = True
        '
        'ChkSleepOnSilentStartup
        '
        Me.ChkSleepOnSilentStartup.AutoSize = True
        Me.ChkSleepOnSilentStartup.Location = New System.Drawing.Point(531, 382)
        Me.ChkSleepOnSilentStartup.Name = "ChkSleepOnSilentStartup"
        Me.ChkSleepOnSilentStartup.Size = New System.Drawing.Size(140, 30)
        Me.ChkSleepOnSilentStartup.TabIndex = 15
        Me.ChkSleepOnSilentStartup.Text = "Sleep before processing" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "at user logon"
        Me.ChkSleepOnSilentStartup.UseVisualStyleBackColor = True
        '
        'lblSeconds
        '
        Me.lblSeconds.AutoSize = True
        Me.lblSeconds.Location = New System.Drawing.Point(528, 415)
        Me.lblSeconds.Name = "lblSeconds"
        Me.lblSeconds.Size = New System.Drawing.Size(55, 13)
        Me.lblSeconds.TabIndex = 16
        Me.lblSeconds.Text = "Seconds?"
        '
        'txtSeconds
        '
        Me.txtSeconds.Location = New System.Drawing.Point(589, 412)
        Me.txtSeconds.Name = "txtSeconds"
        Me.txtSeconds.Size = New System.Drawing.Size(44, 20)
        Me.txtSeconds.TabIndex = 17
        '
        'btnSaveSeconds
        '
        Me.btnSaveSeconds.Location = New System.Drawing.Point(639, 410)
        Me.btnSaveSeconds.Name = "btnSaveSeconds"
        Me.btnSaveSeconds.Size = New System.Drawing.Size(38, 23)
        Me.btnSaveSeconds.TabIndex = 18
        Me.btnSaveSeconds.Text = "Set"
        Me.btnSaveSeconds.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(690, 445)
        Me.Controls.Add(Me.btnSaveSeconds)
        Me.Controls.Add(Me.txtSeconds)
        Me.Controls.Add(Me.lblSeconds)
        Me.Controls.Add(Me.ChkSleepOnSilentStartup)
        Me.Controls.Add(Me.lblAdmin)
        Me.Controls.Add(Me.chkMobileMode)
        Me.Controls.Add(Me.lblUpdateNeededOrNot)
        Me.Controls.Add(Me.chkNotifyAfterUpdateatLogon)
        Me.Controls.Add(Me.btnAbout)
        Me.Controls.Add(Me.btnCheckForUpdates)
        Me.Controls.Add(Me.chkLoadAtUserStartup)
        Me.Controls.Add(Me.chkTrim)
        Me.Controls.Add(Me.btnReDownload)
        Me.Controls.Add(Me.btnTrim)
        Me.Controls.Add(Me.btnApplyNewINIFile)
        Me.Controls.Add(Me.TxtCustomEntries)
        Me.Controls.Add(Me.lblYourCustomEntries)
        Me.Controls.Add(Me.lblYourVersion)
        Me.Controls.Add(Me.lblWebSiteVersion)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "Form1"
        Me.Text = "YAWA2 (Yet Another WinApp2.ini) Updater"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblWebSiteVersion As System.Windows.Forms.Label
    Friend WithEvents lblYourVersion As System.Windows.Forms.Label
    Friend WithEvents lblYourCustomEntries As System.Windows.Forms.Label
    Friend WithEvents TxtCustomEntries As System.Windows.Forms.TextBox
    Friend WithEvents btnApplyNewINIFile As System.Windows.Forms.Button
    Friend WithEvents btnTrim As System.Windows.Forms.Button
    Friend WithEvents btnReDownload As System.Windows.Forms.Button
    Friend WithEvents chkTrim As System.Windows.Forms.CheckBox
    Friend WithEvents chkLoadAtUserStartup As System.Windows.Forms.CheckBox
    Friend WithEvents btnCheckForUpdates As System.Windows.Forms.Button
    Friend WithEvents btnAbout As System.Windows.Forms.Button
    Friend WithEvents chkNotifyAfterUpdateatLogon As System.Windows.Forms.CheckBox
    Friend WithEvents lblUpdateNeededOrNot As System.Windows.Forms.Label
    Friend WithEvents chkMobileMode As System.Windows.Forms.CheckBox
    Friend WithEvents lblAdmin As System.Windows.Forms.Label
    Friend WithEvents ChkSleepOnSilentStartup As CheckBox
    Friend WithEvents lblSeconds As Label
    Friend WithEvents txtSeconds As TextBox
    Friend WithEvents btnSaveSeconds As Button
End Class
