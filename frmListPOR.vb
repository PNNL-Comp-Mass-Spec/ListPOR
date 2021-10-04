Option Strict On

Imports System.ComponentModel
Imports System.IO
Imports PRISM

Public Class frmListPOR
    Inherits Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        mListPOR = New clsListPOR()

        txtStatus.Text = String.Empty
    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    Friend WithEvents fraOptions As System.Windows.Forms.GroupBox
    Friend WithEvents cboConfidenceLevel As System.Windows.Forms.ComboBox
    Friend WithEvents chkIterateRemoval As System.Windows.Forms.CheckBox
    Friend WithEvents lblMinimumFinalDataPointCount As System.Windows.Forms.Label
    Friend WithEvents txtMinimumFinalDataPointCount As System.Windows.Forms.TextBox
    Friend WithEvents fraFilePaths As System.Windows.Forms.GroupBox
    Friend WithEvents cmdBrowseForOutputFile As System.Windows.Forms.Button
    Friend WithEvents lblOutputFilePath As System.Windows.Forms.Label
    Friend WithEvents txtOutputFilePath As System.Windows.Forms.TextBox
    Friend WithEvents cmdBrowseForInputFile As System.Windows.Forms.Button
    Friend WithEvents lblInputFilePath As System.Windows.Forms.Label
    Friend WithEvents txtInputFilePath As System.Windows.Forms.TextBox
    Friend WithEvents fraControls As System.Windows.Forms.GroupBox
    Friend WithEvents pbarProgress As System.Windows.Forms.ProgressBar
    Friend WithEvents chkUseSymmetricValues As System.Windows.Forms.CheckBox
    Friend WithEvents chkUseNaturalLogValues As System.Windows.Forms.CheckBox
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents mnuFile As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFileExit As System.Windows.Forms.MenuItem
    Friend WithEvents mnuEdit As System.Windows.Forms.MenuItem
    Friend WithEvents mnuEditResetToDefaults As System.Windows.Forms.MenuItem
    Friend WithEvents mnuHelp As System.Windows.Forms.MenuItem
    Friend WithEvents mnuHelpAbout As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFileSelectInputFile As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFileSelectOutputFile As System.Windows.Forms.MenuItem
    Friend WithEvents mnuFileSep1 As System.Windows.Forms.MenuItem
    Friend WithEvents cmdExit As System.Windows.Forms.Button
    Friend WithEvents cmdStart As System.Windows.Forms.Button
    Friend WithEvents mnuHelpOverview As System.Windows.Forms.MenuItem
    Friend WithEvents txtStatus As TextBox
    Friend WithEvents chkAssumeSorted As CheckBox
    Friend WithEvents mnuHelpSep1 As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.fraOptions = New System.Windows.Forms.GroupBox()
        Me.chkUseNaturalLogValues = New System.Windows.Forms.CheckBox()
        Me.chkUseSymmetricValues = New System.Windows.Forms.CheckBox()
        Me.txtMinimumFinalDataPointCount = New System.Windows.Forms.TextBox()
        Me.lblMinimumFinalDataPointCount = New System.Windows.Forms.Label()
        Me.cboConfidenceLevel = New System.Windows.Forms.ComboBox()
        Me.chkIterateRemoval = New System.Windows.Forms.CheckBox()
        Me.fraFilePaths = New System.Windows.Forms.GroupBox()
        Me.cmdBrowseForOutputFile = New System.Windows.Forms.Button()
        Me.lblOutputFilePath = New System.Windows.Forms.Label()
        Me.txtOutputFilePath = New System.Windows.Forms.TextBox()
        Me.cmdBrowseForInputFile = New System.Windows.Forms.Button()
        Me.lblInputFilePath = New System.Windows.Forms.Label()
        Me.txtInputFilePath = New System.Windows.Forms.TextBox()
        Me.fraControls = New System.Windows.Forms.GroupBox()
        Me.pbarProgress = New System.Windows.Forms.ProgressBar()
        Me.MainMenu1 = New System.Windows.Forms.MainMenu(Me.components)
        Me.mnuFile = New System.Windows.Forms.MenuItem()
        Me.mnuFileSelectInputFile = New System.Windows.Forms.MenuItem()
        Me.mnuFileSelectOutputFile = New System.Windows.Forms.MenuItem()
        Me.mnuFileSep1 = New System.Windows.Forms.MenuItem()
        Me.mnuFileExit = New System.Windows.Forms.MenuItem()
        Me.mnuEdit = New System.Windows.Forms.MenuItem()
        Me.mnuEditResetToDefaults = New System.Windows.Forms.MenuItem()
        Me.mnuHelp = New System.Windows.Forms.MenuItem()
        Me.mnuHelpOverview = New System.Windows.Forms.MenuItem()
        Me.mnuHelpSep1 = New System.Windows.Forms.MenuItem()
        Me.mnuHelpAbout = New System.Windows.Forms.MenuItem()
        Me.cmdExit = New System.Windows.Forms.Button()
        Me.cmdStart = New System.Windows.Forms.Button()
        Me.txtStatus = New System.Windows.Forms.TextBox()
        Me.chkAssumeSorted = New System.Windows.Forms.CheckBox()
        Me.fraOptions.SuspendLayout()
        Me.fraFilePaths.SuspendLayout()
        Me.fraControls.SuspendLayout()
        Me.SuspendLayout()
        '
        'fraOptions
        '
        Me.fraOptions.Controls.Add(Me.chkAssumeSorted)
        Me.fraOptions.Controls.Add(Me.chkUseNaturalLogValues)
        Me.fraOptions.Controls.Add(Me.chkUseSymmetricValues)
        Me.fraOptions.Controls.Add(Me.txtMinimumFinalDataPointCount)
        Me.fraOptions.Controls.Add(Me.lblMinimumFinalDataPointCount)
        Me.fraOptions.Controls.Add(Me.cboConfidenceLevel)
        Me.fraOptions.Controls.Add(Me.chkIterateRemoval)
        Me.fraOptions.Location = New System.Drawing.Point(10, 157)
        Me.fraOptions.Name = "fraOptions"
        Me.fraOptions.Size = New System.Drawing.Size(297, 192)
        Me.fraOptions.TabIndex = 2
        Me.fraOptions.TabStop = False
        Me.fraOptions.Text = "Options"
        '
        'chkUseNaturalLogValues
        '
        Me.chkUseNaturalLogValues.Location = New System.Drawing.Point(154, 111)
        Me.chkUseNaturalLogValues.Name = "chkUseNaturalLogValues"
        Me.chkUseNaturalLogValues.Size = New System.Drawing.Size(134, 48)
        Me.chkUseNaturalLogValues.TabIndex = 5
        Me.chkUseNaturalLogValues.Text = "Use Natural Log Values (all data should be > 0)"
        '
        'chkUseSymmetricValues
        '
        Me.chkUseSymmetricValues.Checked = True
        Me.chkUseSymmetricValues.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkUseSymmetricValues.Location = New System.Drawing.Point(19, 111)
        Me.chkUseSymmetricValues.Name = "chkUseSymmetricValues"
        Me.chkUseSymmetricValues.Size = New System.Drawing.Size(135, 48)
        Me.chkUseSymmetricValues.TabIndex = 4
        Me.chkUseSymmetricValues.Text = "Use Symmetric Values (all data should be > 0)"
        '
        'txtMinimumFinalDataPointCount
        '
        Me.txtMinimumFinalDataPointCount.Location = New System.Drawing.Point(211, 83)
        Me.txtMinimumFinalDataPointCount.Name = "txtMinimumFinalDataPointCount"
        Me.txtMinimumFinalDataPointCount.Size = New System.Drawing.Size(58, 22)
        Me.txtMinimumFinalDataPointCount.TabIndex = 3
        Me.txtMinimumFinalDataPointCount.Text = "3"
        '
        'lblMinimumFinalDataPointCount
        '
        Me.lblMinimumFinalDataPointCount.Location = New System.Drawing.Point(10, 85)
        Me.lblMinimumFinalDataPointCount.Name = "lblMinimumFinalDataPointCount"
        Me.lblMinimumFinalDataPointCount.Size = New System.Drawing.Size(201, 19)
        Me.lblMinimumFinalDataPointCount.TabIndex = 2
        Me.lblMinimumFinalDataPointCount.Text = "Minimum final data point count"
        '
        'cboConfidenceLevel
        '
        Me.cboConfidenceLevel.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboConfidenceLevel.Location = New System.Drawing.Point(19, 51)
        Me.cboConfidenceLevel.Name = "cboConfidenceLevel"
        Me.cboConfidenceLevel.Size = New System.Drawing.Size(135, 24)
        Me.cboConfidenceLevel.TabIndex = 1
        '
        'chkIterateRemoval
        '
        Me.chkIterateRemoval.Location = New System.Drawing.Point(19, 18)
        Me.chkIterateRemoval.Name = "chkIterateRemoval"
        Me.chkIterateRemoval.Size = New System.Drawing.Size(221, 28)
        Me.chkIterateRemoval.TabIndex = 0
        Me.chkIterateRemoval.Text = "Repeatedly remove outliers"
        '
        'fraFilePaths
        '
        Me.fraFilePaths.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.fraFilePaths.Controls.Add(Me.cmdBrowseForOutputFile)
        Me.fraFilePaths.Controls.Add(Me.lblOutputFilePath)
        Me.fraFilePaths.Controls.Add(Me.txtOutputFilePath)
        Me.fraFilePaths.Controls.Add(Me.cmdBrowseForInputFile)
        Me.fraFilePaths.Controls.Add(Me.lblInputFilePath)
        Me.fraFilePaths.Controls.Add(Me.txtInputFilePath)
        Me.fraFilePaths.Location = New System.Drawing.Point(10, 18)
        Me.fraFilePaths.Name = "fraFilePaths"
        Me.fraFilePaths.Size = New System.Drawing.Size(562, 130)
        Me.fraFilePaths.TabIndex = 1
        Me.fraFilePaths.TabStop = False
        Me.fraFilePaths.Text = "File Paths"
        '
        'cmdBrowseForOutputFile
        '
        Me.cmdBrowseForOutputFile.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdBrowseForOutputFile.Location = New System.Drawing.Point(486, 92)
        Me.cmdBrowseForOutputFile.Name = "cmdBrowseForOutputFile"
        Me.cmdBrowseForOutputFile.Size = New System.Drawing.Size(67, 23)
        Me.cmdBrowseForOutputFile.TabIndex = 5
        Me.cmdBrowseForOutputFile.Text = "Br&owse"
        '
        'lblOutputFilePath
        '
        Me.lblOutputFilePath.Location = New System.Drawing.Point(19, 74)
        Me.lblOutputFilePath.Name = "lblOutputFilePath"
        Me.lblOutputFilePath.Size = New System.Drawing.Size(120, 18)
        Me.lblOutputFilePath.TabIndex = 3
        Me.lblOutputFilePath.Text = "Output file path"
        '
        'txtOutputFilePath
        '
        Me.txtOutputFilePath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtOutputFilePath.Location = New System.Drawing.Point(19, 92)
        Me.txtOutputFilePath.Name = "txtOutputFilePath"
        Me.txtOutputFilePath.Size = New System.Drawing.Size(457, 22)
        Me.txtOutputFilePath.TabIndex = 4
        Me.txtOutputFilePath.Text = "Output file"
        '
        'cmdBrowseForInputFile
        '
        Me.cmdBrowseForInputFile.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdBrowseForInputFile.Location = New System.Drawing.Point(486, 37)
        Me.cmdBrowseForInputFile.Name = "cmdBrowseForInputFile"
        Me.cmdBrowseForInputFile.Size = New System.Drawing.Size(67, 23)
        Me.cmdBrowseForInputFile.TabIndex = 2
        Me.cmdBrowseForInputFile.Text = "&Browse"
        '
        'lblInputFilePath
        '
        Me.lblInputFilePath.Location = New System.Drawing.Point(19, 18)
        Me.lblInputFilePath.Name = "lblInputFilePath"
        Me.lblInputFilePath.Size = New System.Drawing.Size(120, 19)
        Me.lblInputFilePath.TabIndex = 0
        Me.lblInputFilePath.Text = "Input file path"
        '
        'txtInputFilePath
        '
        Me.txtInputFilePath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtInputFilePath.Location = New System.Drawing.Point(19, 37)
        Me.txtInputFilePath.Name = "txtInputFilePath"
        Me.txtInputFilePath.Size = New System.Drawing.Size(457, 22)
        Me.txtInputFilePath.TabIndex = 1
        Me.txtInputFilePath.Text = "Input file"
        '
        'fraControls
        '
        Me.fraControls.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.fraControls.Controls.Add(Me.pbarProgress)
        Me.fraControls.Location = New System.Drawing.Point(317, 157)
        Me.fraControls.Name = "fraControls"
        Me.fraControls.Size = New System.Drawing.Size(254, 55)
        Me.fraControls.TabIndex = 3
        Me.fraControls.TabStop = False
        Me.fraControls.Text = "Progress"
        '
        'pbarProgress
        '
        Me.pbarProgress.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pbarProgress.Location = New System.Drawing.Point(10, 18)
        Me.pbarProgress.Name = "pbarProgress"
        Me.pbarProgress.Size = New System.Drawing.Size(226, 28)
        Me.pbarProgress.Style = System.Windows.Forms.ProgressBarStyle.Continuous
        Me.pbarProgress.TabIndex = 2
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuFile, Me.mnuEdit, Me.mnuHelp})
        '
        'mnuFile
        '
        Me.mnuFile.Index = 0
        Me.mnuFile.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuFileSelectInputFile, Me.mnuFileSelectOutputFile, Me.mnuFileSep1, Me.mnuFileExit})
        Me.mnuFile.Text = "&File"
        '
        'mnuFileSelectInputFile
        '
        Me.mnuFileSelectInputFile.Index = 0
        Me.mnuFileSelectInputFile.Text = "Select &Input File"
        '
        'mnuFileSelectOutputFile
        '
        Me.mnuFileSelectOutputFile.Index = 1
        Me.mnuFileSelectOutputFile.Text = "Select &Output File"
        '
        'mnuFileSep1
        '
        Me.mnuFileSep1.Index = 2
        Me.mnuFileSep1.Text = "-"
        '
        'mnuFileExit
        '
        Me.mnuFileExit.Index = 3
        Me.mnuFileExit.Text = "E&xit"
        '
        'mnuEdit
        '
        Me.mnuEdit.Index = 1
        Me.mnuEdit.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuEditResetToDefaults})
        Me.mnuEdit.Text = "&Edit"
        '
        'mnuEditResetToDefaults
        '
        Me.mnuEditResetToDefaults.Index = 0
        Me.mnuEditResetToDefaults.Text = "&Reset to Defaults"
        '
        'mnuHelp
        '
        Me.mnuHelp.Index = 2
        Me.mnuHelp.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.mnuHelpOverview, Me.mnuHelpSep1, Me.mnuHelpAbout})
        Me.mnuHelp.Text = "&Help"
        '
        'mnuHelpOverview
        '
        Me.mnuHelpOverview.Index = 0
        Me.mnuHelpOverview.Text = "&Overview"
        '
        'mnuHelpSep1
        '
        Me.mnuHelpSep1.Index = 1
        Me.mnuHelpSep1.Text = "-"
        '
        'mnuHelpAbout
        '
        Me.mnuHelpAbout.Index = 2
        Me.mnuHelpAbout.Text = "&About"
        '
        'cmdExit
        '
        Me.cmdExit.Location = New System.Drawing.Point(432, 231)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(67, 23)
        Me.cmdExit.TabIndex = 7
        Me.cmdExit.Text = "E&xit"
        '
        'cmdStart
        '
        Me.cmdStart.Location = New System.Drawing.Point(326, 231)
        Me.cmdStart.Name = "cmdStart"
        Me.cmdStart.Size = New System.Drawing.Size(68, 23)
        Me.cmdStart.TabIndex = 6
        Me.cmdStart.Text = "&Start"
        '
        'txtStatus
        '
        Me.txtStatus.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtStatus.Location = New System.Drawing.Point(327, 280)
        Me.txtStatus.Name = "txtStatus"
        Me.txtStatus.ReadOnly = True
        Me.txtStatus.Size = New System.Drawing.Size(226, 22)
        Me.txtStatus.TabIndex = 8
        '
        'chkAssumeSorted
        '
        Me.chkAssumeSorted.Location = New System.Drawing.Point(19, 162)
        Me.chkAssumeSorted.Name = "chkAssumeSorted"
        Me.chkAssumeSorted.Size = New System.Drawing.Size(192, 24)
        Me.chkAssumeSorted.TabIndex = 6
        Me.chkAssumeSorted.Text = "Assume sorted"
        '
        'frmListPOR
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.ClientSize = New System.Drawing.Size(582, 354)
        Me.Controls.Add(Me.txtStatus)
        Me.Controls.Add(Me.cmdExit)
        Me.Controls.Add(Me.cmdStart)
        Me.Controls.Add(Me.fraControls)
        Me.Controls.Add(Me.fraFilePaths)
        Me.Controls.Add(Me.fraOptions)
        Me.Menu = Me.MainMenu1
        Me.MinimumSize = New System.Drawing.Size(600, 425)
        Me.Name = "frmListPOR"
        Me.Text = "ListPOR - List Parser for Outlier Removal"
        Me.fraOptions.ResumeLayout(False)
        Me.fraOptions.PerformLayout()
        Me.fraFilePaths.ResumeLayout(False)
        Me.fraFilePaths.PerformLayout()
        Me.fraControls.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

#Region "Module-level Variables"
    Private Const XML_SETTINGS_FILE_NAME As String = "ListPORSettings.xml"
    Private Const XML_FILE_OPTIONS_SECTION As String = "Options"

    Private mProcessing As Boolean
    Private mAbortProcessing As Boolean

    Private WithEvents mListPOR As clsListPOR

#End Region

    Private Function BrowseForFile(strCurrentPath As String, Optional ByVal strDialogTitle As String = "Select file") As String

        Dim Dialog As SaveFileDialog
        Dim strCandidateParentFolderName As String

        Dim Result As Integer

        Dialog = New SaveFileDialog
        With Dialog
            .Filter = "Text (*.txt)|*.txt|Comma separated (*.csv)|*.csv|All Files (*.*)|*.*"

            If strCurrentPath.Length > 0 Then
                strCandidateParentFolderName = Path.GetDirectoryName(strCurrentPath)

                If File.Exists(strCurrentPath) Then
                    .FileName = strCurrentPath
                    .InitialDirectory = strCandidateParentFolderName
                ElseIf Directory.Exists(strCurrentPath) Then
                    .InitialDirectory = strCurrentPath
                ElseIf Directory.Exists(strCandidateParentFolderName) Then
                    .InitialDirectory = strCandidateParentFolderName
                    .FileName = Path.GetFileName(strCurrentPath)
                End If
            End If

            .CreatePrompt = False
            .OverwritePrompt = False
            .CheckPathExists = True

            .AddExtension = True
            .ValidateNames = True
            .DefaultExt = ".txt"
            .Title = strDialogTitle
        End With

        Try
            Result = Dialog.ShowDialog
        Catch ex As Exception
            ' Bad file name
            Dialog.FileName = String.Empty
            Result = DialogResult.Cancel
            MessageBox.Show("Invalid filename: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try

        If Result = DialogResult.OK Or Result = DialogResult.Yes Then
            Return Dialog.FileName
        Else
            Return String.Empty
        End If

    End Function

    Private Sub ExitProgram()
        If mProcessing Then
            mAbortProcessing = True
        End If
        Me.Close()
    End Sub

    Private Function GetSettingsFilePath() As String

        Dim strSettingsFilePathLocal As String
        strSettingsFilePathLocal = PRISM.FileProcessor.ProcessFilesBase.GetSettingsFilePathLocal("ListPOR", XML_SETTINGS_FILE_NAME)

        PRISM.FileProcessor.ProcessFilesBase.CreateSettingsFileIfMissing(strSettingsFilePathLocal)

        Return strSettingsFilePathLocal

    End Function

    Private Sub InitializeControls(Optional ByVal blnRetainFilePaths As Boolean = False)

        Try
            If Not blnRetainFilePaths Then
                txtInputFilePath.Text = String.Empty
                txtOutputFilePath.Text = String.Empty
            End If

            chkIterateRemoval.Checked = True
            txtMinimumFinalDataPointCount.Text = "3"

            cboConfidenceLevel.SelectedIndex = clsGrubbsTestOutlierFilter.eclConfidenceLevelConstants.e95Pct

            chkUseSymmetricValues.Checked = False
            chkUseNaturalLogValues.Checked = False

        Catch ex As Exception
            ' Ignore errors here
        End Try

    End Sub

    Private Sub PopulateComboBoxes()
        With cboConfidenceLevel
            .Sorted = False
            .Items.Clear()
            .Items.Add("95% confidence")
            .Items.Add("97% confidence")
            .Items.Add("99% confidence")
            .SelectedIndex = clsGrubbsTestOutlierFilter.eclConfidenceLevelConstants.e95Pct
        End With
    End Sub

    Private Sub SelectInputFile()
        Try
            Dim strNewFilePath As String
            strNewFilePath = BrowseForFile(txtInputFilePath.Text)

            If strNewFilePath.Length > 0 Then
                txtInputFilePath.Text = strNewFilePath
            End If
        Catch ex As Exception
            ' Ignore errors here
        End Try
    End Sub

    Private Sub SelectOutputFile()
        Try
            Dim strNewFilePath As String
            strNewFilePath = BrowseForFile(txtOutputFilePath.Text)

            If strNewFilePath.Length > 0 Then
                txtOutputFilePath.Text = strNewFilePath
            End If
        Catch ex As Exception
            ' Ignore errors here
        End Try
    End Sub

    Private Sub ShowAboutBox()
        Dim strMessage As String

        strMessage = String.Empty
        strMessage &= "Program written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA) in 2004" & ControlChars.NewLine & ControlChars.NewLine

        strMessage &= "This is version " & Application.ProductVersion & " (" & PROGRAM_DATE & ")" & ControlChars.NewLine & ControlChars.NewLine

        strMessage &= "E-mail: matthew.monroe@pnnl.gov or proteomics@pnnl.gov" & ControlChars.NewLine
        strMessage &= "Website: https://github.com/PNNL-Comp-Mass-Spec/ or https://panomics.pnnl.gov/ or https://www.pnnl.gov/integrative-omics" & ControlChars.NewLine & ControlChars.NewLine

        strMessage &= "Licensed under the Apache License, Version 2.0; you may not use this file except in compliance with the License.  "
        strMessage &= "You may obtain a copy of the License at http://www.apache.org/licenses/LICENSE-2.0" & ControlChars.NewLine & ControlChars.NewLine

        MessageBox.Show(strMessage, "About", MessageBoxButtons.OK, MessageBoxIcon.Information)

    End Sub

    Private Sub ShowOverview()
        Dim strMessage As String

        strMessage = String.Empty
        strMessage &= "This program will read the input file and examine the values to remove outliers. Typically, the first column will contain key information (text or numbers) on which to group the data. "
        strMessage &= "The second column will contain the values to examine for outliers, examining the data by groups.  If additional columns are present, they will be written to the output file along with the key and value columns. "
        strMessage &= "Alternatively, if you provide a file with only one column of data, then all of the values will be examined en masse and the outliers removed.  If you do not provide an output file path, then the input file's path will be used, but with _filtered appeneded to the name"

        MessageBox.Show(strMessage, "Overview", MessageBoxButtons.OK, MessageBoxIcon.Information)

    End Sub

    Private Sub ValidateFilesAndProcess()
        Dim strInputFilePath As String
        Dim strOutputFilePath As String

        Dim intErrorCode As Integer

        txtStatus.Text = String.Empty

        Try
            strInputFilePath = txtInputFilePath.Text.Trim()
            strOutputFilePath = txtOutputFilePath.Text.Trim()

            If String.IsNullOrWhiteSpace(strInputFilePath) Then
                MessageBox.Show("Please enter a valid input file path", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Else
                If Not File.Exists(strInputFilePath) Then
                    MessageBox.Show("File not found: " & strInputFilePath, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Else
                    If String.Equals(strInputFilePath, strOutputFilePath, StringComparison.OrdinalIgnoreCase) Then
                        MessageBox.Show("The input and output file paths cannot be identical", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Else
                        Try
                            mProcessing = True
                            mAbortProcessing = False

                            Me.Cursor = Cursors.WaitCursor
                            pbarProgress.Value = 0
                            cmdStart.Text = "&Abort"

                            With mListPOR
                                .AssumeSortedInputFile = chkAssumeSorted.Checked
                                .RemoveMultipleValues = chkIterateRemoval.Checked
                                .ConfidenceLevel = CType(cboConfidenceLevel.SelectedIndex, clsGrubbsTestOutlierFilter.eclConfidenceLevelConstants)
                                If IsNumeric(txtMinimumFinalDataPointCount.Text) Then
                                    .MinFinalValueCount = CType(txtMinimumFinalDataPointCount.Text, Integer)
                                End If
                                .UseSymmetricValues = chkUseSymmetricValues.Checked
                                .UseNaturalLogValues = chkUseNaturalLogValues.Checked
                            End With
                            intErrorCode = mListPOR.RemoveOutliersFromListInFile(strInputFilePath, strOutputFilePath)

                            mProcessing = False
                            mAbortProcessing = False

                        Catch ex As Exception
                            intErrorCode = -100
                        Finally
                            Me.Cursor = Cursors.Default
                            cmdStart.Text = "&Start"
                        End Try

                        If intErrorCode = 0 Then
                            txtStatus.Text = mListPOR.MostRecentOutputFile
                            MessageBox.Show("Processing complete", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Else
                            MessageBox.Show("Error while processing: " & mListPOR.GetErrorMessage(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                        End If
                    End If
                End If
            End If

        Catch ex As Exception
            MessageBox.Show("Error validating the input and output file paths", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try

    End Sub

    Private Sub ValidateDataTransformOptions(blnFavorNaturalLog As Boolean)
        If chkUseNaturalLogValues.Checked And chkUseSymmetricValues.Checked Then
            If blnFavorNaturalLog Then
                chkUseSymmetricValues.Checked = False
            Else
                chkUseNaturalLogValues.Checked = False
            End If
        End If
    End Sub

    Private Sub XMLReadSettings()
        ' Read the settings from the XML file
        Dim strFilePath As String
        Dim intConfidenceValue As Integer

        Try
            InitializeControls(True)

            strFilePath = GetSettingsFilePath()

            Dim objXmlFile As New XmlSettingsFileAccessor()

            If objXmlFile.LoadSettings(strFilePath) Then
                txtInputFilePath.Text = objXmlFile.GetParam(XML_FILE_OPTIONS_SECTION, "InputFilePath", txtInputFilePath.Text)
                txtOutputFilePath.Text = objXmlFile.GetParam(XML_FILE_OPTIONS_SECTION, "OutputFilePath", txtOutputFilePath.Text)

                chkAssumeSorted.Checked = objXmlFile.GetParam(XML_FILE_OPTIONS_SECTION, "AssumeSorted", chkAssumeSorted.Checked)

                intConfidenceValue = clsGrubbsTestOutlierFilter.eclConfidenceLevelConstants.e95Pct
                intConfidenceValue = objXmlFile.GetParam(XML_FILE_OPTIONS_SECTION, "ConfidenceValueIndex", intConfidenceValue)

                Try
                    cboConfidenceLevel.SelectedIndex = intConfidenceValue
                Catch ex As Exception
                    cboConfidenceLevel.SelectedIndex = clsGrubbsTestOutlierFilter.eclConfidenceLevelConstants.e95Pct
                End Try

                chkIterateRemoval.Checked = objXmlFile.GetParam(XML_FILE_OPTIONS_SECTION, "IterateRemoval", chkIterateRemoval.Checked)
                txtMinimumFinalDataPointCount.Text = objXmlFile.GetParam(XML_FILE_OPTIONS_SECTION, "MinimumFinalDataPointCount", txtMinimumFinalDataPointCount.Text)

                chkUseSymmetricValues.Checked = objXmlFile.GetParam(XML_FILE_OPTIONS_SECTION, "UseSymmetricValues", chkUseSymmetricValues.Checked)
                chkUseNaturalLogValues.Checked = objXmlFile.GetParam(XML_FILE_OPTIONS_SECTION, "UseNaturalLogVAlues", chkUseNaturalLogValues.Checked)
            End If

        Catch ex As Exception
            MessageBox.Show("Error loading settings from XML Settings file: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try

    End Sub

    Private Sub XMLSaveSettings()
        Dim blnSuccess As Boolean
        Dim strFilePath As String

        Try
            strFilePath = GetSettingsFilePath()

            Dim objXmlFile As New XmlSettingsFileAccessor()

            If objXmlFile.LoadSettings(strFilePath) Then

                blnSuccess = objXmlFile.SetParam(XML_FILE_OPTIONS_SECTION, "InputFilePath", txtInputFilePath.Text)

                If Not blnSuccess Then
                    Throw New Exception("Unknown error while setting a value.")
                Else
                    ' Continue saving settings

                    objXmlFile.SetParam(XML_FILE_OPTIONS_SECTION, "OutputFilePath", txtOutputFilePath.Text)

                    objXmlFile.SetParam(XML_FILE_OPTIONS_SECTION, "AssumeSorted", chkAssumeSorted.Checked)

                    If cboConfidenceLevel.SelectedIndex >= 0 Then
                        objXmlFile.SetParam(XML_FILE_OPTIONS_SECTION, "ConfidenceValueIndex", cboConfidenceLevel.SelectedIndex)
                    Else
                        objXmlFile.SetParam(XML_FILE_OPTIONS_SECTION, "ConfidenceValueIndex", clsGrubbsTestOutlierFilter.eclConfidenceLevelConstants.e95Pct)
                    End If

                    objXmlFile.SetParam(XML_FILE_OPTIONS_SECTION, "IterateRemoval", chkIterateRemoval.Checked)
                    objXmlFile.SetParam(XML_FILE_OPTIONS_SECTION, "MinimumFinalDataPointCount", txtMinimumFinalDataPointCount.Text)

                    objXmlFile.SetParam(XML_FILE_OPTIONS_SECTION, "UseSymmetricValues", chkUseSymmetricValues.Checked)
                    objXmlFile.SetParam(XML_FILE_OPTIONS_SECTION, "UseNaturalLogVAlues", chkUseNaturalLogValues.Checked)

                    If Not objXmlFile.SaveSettings() Then
                        Throw New Exception("Unknown error while saving the file.")
                    End If
                End If

            End If

        Catch ex As Exception
            MessageBox.Show("Error saving settings to XML Settings file: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try

    End Sub

    Private Sub chkUseSymmetricValues_CheckedChanged(sender As Object, e As EventArgs) Handles chkUseSymmetricValues.CheckedChanged
        ValidateDataTransformOptions(False)
    End Sub

    Private Sub chkUseNaturalLogValues_CheckedChanged(sender As Object, e As EventArgs) Handles chkUseNaturalLogValues.CheckedChanged
        ValidateDataTransformOptions(True)
    End Sub

    Private Sub cmdBrowseForInputFile_Click(sender As Object, e As EventArgs) Handles cmdBrowseForInputFile.Click
        SelectInputFile()
    End Sub

    Private Sub cmdBrowseForOutputFile_Click(sender As Object, e As EventArgs) Handles cmdBrowseForOutputFile.Click
        SelectOutputFile()
    End Sub

    Private Sub cmdExit_Click(sender As Object, e As EventArgs) Handles cmdExit.Click
        ExitProgram()
    End Sub

    Private Sub cmdStart_Click(sender As Object, e As EventArgs) Handles cmdStart.Click
        If mProcessing Then
            mAbortProcessing = True
        Else
            ValidateFilesAndProcess()
        End If
    End Sub

    Private Sub frmListPOR_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        PopulateComboBoxes()
        InitializeControls(False)

        XMLReadSettings()
    End Sub

    Private Sub txtMinimumFinalDataPointCount_TextChanged(sender As Object, e As EventArgs) Handles txtMinimumFinalDataPointCount.TextChanged
        If txtMinimumFinalDataPointCount.Text.Length > 0 Then
            If Not IsNumeric(txtMinimumFinalDataPointCount.Text) Then
                txtMinimumFinalDataPointCount.Text = "3"
            End If
        End If
    End Sub

    Private Sub frmListPOR_Closing(sender As Object, e As CancelEventArgs) Handles MyBase.Closing
        If mProcessing Then
            mAbortProcessing = True
        End If

        XMLSaveSettings()
    End Sub

    Private Sub mListPOR_ProgressUpdate(progressMessage As String, percentComplete As Single) Handles mListPOR.ProgressUpdate
        Try
            pbarProgress.Value = CInt(Math.Round(percentComplete, 0))
            Application.DoEvents()

            If mAbortProcessing Then
                mListPOR.AbortProcessing()
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub txtMinimumFinalDataPointCount_LostFocus(sender As Object, e As EventArgs) Handles txtMinimumFinalDataPointCount.LostFocus
        Try
            If CType(txtMinimumFinalDataPointCount.Text, Integer) < 2 Then
                txtMinimumFinalDataPointCount.Text = "3"
            End If
        Catch ex As Exception
            txtMinimumFinalDataPointCount.Text = "3"
        End Try
    End Sub

    Private Sub mnuEditResetToDefaults_Click(sender As Object, e As EventArgs) Handles mnuEditResetToDefaults.Click
        If Not mProcessing Then
            InitializeControls(True)
        End If
    End Sub

    Private Sub mnuFileSelectInputFile_Click(sender As Object, e As EventArgs) Handles mnuFileSelectInputFile.Click
        SelectInputFile()
    End Sub

    Private Sub mnuFileSelectOutputFile_Click(sender As Object, e As EventArgs) Handles mnuFileSelectOutputFile.Click
        SelectOutputFile()
    End Sub

    Private Sub mnuFileExit_Click(sender As Object, e As EventArgs) Handles mnuFileExit.Click
        ExitProgram()
    End Sub

    Private Sub mnuHelpAbout_Click(sender As Object, e As EventArgs) Handles mnuHelpAbout.Click
        ShowAboutBox()
    End Sub

    Private Sub mnuHelpOverview_Click(sender As Object, e As EventArgs) Handles mnuHelpOverview.Click
        ShowOverview()
    End Sub

    Private Sub mListPOR_ErrorEvent(message As String, ex As Exception) Handles mListPOR.ErrorEvent
        MessageBox.Show(message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
    End Sub

    Private Sub mListPOR_WarningEvent(message As String) Handles mListPOR.WarningEvent
        MessageBox.Show(message, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
    End Sub

    Private Sub mListPOR_StatusEvent(message As String) Handles mListPOR.StatusEvent
        txtStatus.Text = message
        Application.DoEvents()
    End Sub
End Class
