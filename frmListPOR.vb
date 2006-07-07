Option Strict On

Public Class frmListPOR
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        mListPOR = New clsListPOR

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
    Friend WithEvents pbarProgress As SmoothProgressBar.SmoothProgressBar
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
    Friend WithEvents mnuHelpSep1 As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.fraOptions = New System.Windows.Forms.GroupBox
        Me.chkUseNaturalLogValues = New System.Windows.Forms.CheckBox
        Me.chkUseSymmetricValues = New System.Windows.Forms.CheckBox
        Me.txtMinimumFinalDataPointCount = New System.Windows.Forms.TextBox
        Me.lblMinimumFinalDataPointCount = New System.Windows.Forms.Label
        Me.cboConfidenceLevel = New System.Windows.Forms.ComboBox
        Me.chkIterateRemoval = New System.Windows.Forms.CheckBox
        Me.fraFilePaths = New System.Windows.Forms.GroupBox
        Me.cmdBrowseForOutputFile = New System.Windows.Forms.Button
        Me.lblOutputFilePath = New System.Windows.Forms.Label
        Me.txtOutputFilePath = New System.Windows.Forms.TextBox
        Me.cmdBrowseForInputFile = New System.Windows.Forms.Button
        Me.lblInputFilePath = New System.Windows.Forms.Label
        Me.txtInputFilePath = New System.Windows.Forms.TextBox
        Me.fraControls = New System.Windows.Forms.GroupBox
        Me.pbarProgress = New SmoothProgressBar.SmoothProgressBar
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.mnuFile = New System.Windows.Forms.MenuItem
        Me.mnuFileSelectInputFile = New System.Windows.Forms.MenuItem
        Me.mnuFileSelectOutputFile = New System.Windows.Forms.MenuItem
        Me.mnuFileSep1 = New System.Windows.Forms.MenuItem
        Me.mnuFileExit = New System.Windows.Forms.MenuItem
        Me.mnuEdit = New System.Windows.Forms.MenuItem
        Me.mnuEditResetToDefaults = New System.Windows.Forms.MenuItem
        Me.mnuHelp = New System.Windows.Forms.MenuItem
        Me.mnuHelpAbout = New System.Windows.Forms.MenuItem
        Me.cmdExit = New System.Windows.Forms.Button
        Me.cmdStart = New System.Windows.Forms.Button
        Me.mnuHelpOverview = New System.Windows.Forms.MenuItem
        Me.mnuHelpSep1 = New System.Windows.Forms.MenuItem
        Me.fraOptions.SuspendLayout()
        Me.fraFilePaths.SuspendLayout()
        Me.fraControls.SuspendLayout()
        Me.SuspendLayout()
        '
        'fraOptions
        '
        Me.fraOptions.Controls.Add(Me.chkUseNaturalLogValues)
        Me.fraOptions.Controls.Add(Me.chkUseSymmetricValues)
        Me.fraOptions.Controls.Add(Me.txtMinimumFinalDataPointCount)
        Me.fraOptions.Controls.Add(Me.lblMinimumFinalDataPointCount)
        Me.fraOptions.Controls.Add(Me.cboConfidenceLevel)
        Me.fraOptions.Controls.Add(Me.chkIterateRemoval)
        Me.fraOptions.Location = New System.Drawing.Point(8, 136)
        Me.fraOptions.Name = "fraOptions"
        Me.fraOptions.Size = New System.Drawing.Size(248, 144)
        Me.fraOptions.TabIndex = 2
        Me.fraOptions.TabStop = False
        Me.fraOptions.Text = "Options"
        '
        'chkUseNaturalLogValues
        '
        Me.chkUseNaturalLogValues.Location = New System.Drawing.Point(128, 96)
        Me.chkUseNaturalLogValues.Name = "chkUseNaturalLogValues"
        Me.chkUseNaturalLogValues.Size = New System.Drawing.Size(112, 42)
        Me.chkUseNaturalLogValues.TabIndex = 5
        Me.chkUseNaturalLogValues.Text = "Use Natural Log Values (all data should be > 0)"
        '
        'chkUseSymmetricValues
        '
        Me.chkUseSymmetricValues.Checked = True
        Me.chkUseSymmetricValues.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkUseSymmetricValues.Location = New System.Drawing.Point(16, 96)
        Me.chkUseSymmetricValues.Name = "chkUseSymmetricValues"
        Me.chkUseSymmetricValues.Size = New System.Drawing.Size(112, 42)
        Me.chkUseSymmetricValues.TabIndex = 4
        Me.chkUseSymmetricValues.Text = "Use Symmetric Values (all data should be > 0)"
        '
        'txtMinimumFinalDataPointCount
        '
        Me.txtMinimumFinalDataPointCount.Location = New System.Drawing.Point(176, 72)
        Me.txtMinimumFinalDataPointCount.Name = "txtMinimumFinalDataPointCount"
        Me.txtMinimumFinalDataPointCount.Size = New System.Drawing.Size(48, 20)
        Me.txtMinimumFinalDataPointCount.TabIndex = 3
        Me.txtMinimumFinalDataPointCount.Text = "3"
        '
        'lblMinimumFinalDataPointCount
        '
        Me.lblMinimumFinalDataPointCount.Location = New System.Drawing.Point(8, 74)
        Me.lblMinimumFinalDataPointCount.Name = "lblMinimumFinalDataPointCount"
        Me.lblMinimumFinalDataPointCount.Size = New System.Drawing.Size(168, 16)
        Me.lblMinimumFinalDataPointCount.TabIndex = 2
        Me.lblMinimumFinalDataPointCount.Text = "Minimum final data point count"
        '
        'cboConfidenceLevel
        '
        Me.cboConfidenceLevel.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboConfidenceLevel.Location = New System.Drawing.Point(16, 44)
        Me.cboConfidenceLevel.Name = "cboConfidenceLevel"
        Me.cboConfidenceLevel.Size = New System.Drawing.Size(112, 21)
        Me.cboConfidenceLevel.TabIndex = 1
        '
        'chkIterateRemoval
        '
        Me.chkIterateRemoval.Location = New System.Drawing.Point(16, 16)
        Me.chkIterateRemoval.Name = "chkIterateRemoval"
        Me.chkIterateRemoval.Size = New System.Drawing.Size(184, 24)
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
        Me.fraFilePaths.Location = New System.Drawing.Point(8, 16)
        Me.fraFilePaths.Name = "fraFilePaths"
        Me.fraFilePaths.Size = New System.Drawing.Size(504, 112)
        Me.fraFilePaths.TabIndex = 1
        Me.fraFilePaths.TabStop = False
        Me.fraFilePaths.Text = "File Paths"
        '
        'cmdBrowseForOutputFile
        '
        Me.cmdBrowseForOutputFile.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdBrowseForOutputFile.Location = New System.Drawing.Point(440, 80)
        Me.cmdBrowseForOutputFile.Name = "cmdBrowseForOutputFile"
        Me.cmdBrowseForOutputFile.Size = New System.Drawing.Size(56, 20)
        Me.cmdBrowseForOutputFile.TabIndex = 5
        Me.cmdBrowseForOutputFile.Text = "Br&owse"
        '
        'lblOutputFilePath
        '
        Me.lblOutputFilePath.Location = New System.Drawing.Point(16, 64)
        Me.lblOutputFilePath.Name = "lblOutputFilePath"
        Me.lblOutputFilePath.Size = New System.Drawing.Size(72, 16)
        Me.lblOutputFilePath.TabIndex = 3
        Me.lblOutputFilePath.Text = "Output file path"
        '
        'txtOutputFilePath
        '
        Me.txtOutputFilePath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtOutputFilePath.Location = New System.Drawing.Point(16, 80)
        Me.txtOutputFilePath.Name = "txtOutputFilePath"
        Me.txtOutputFilePath.Size = New System.Drawing.Size(416, 20)
        Me.txtOutputFilePath.TabIndex = 4
        Me.txtOutputFilePath.Text = "Output file"
        '
        'cmdBrowseForInputFile
        '
        Me.cmdBrowseForInputFile.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdBrowseForInputFile.Location = New System.Drawing.Point(440, 32)
        Me.cmdBrowseForInputFile.Name = "cmdBrowseForInputFile"
        Me.cmdBrowseForInputFile.Size = New System.Drawing.Size(56, 20)
        Me.cmdBrowseForInputFile.TabIndex = 2
        Me.cmdBrowseForInputFile.Text = "&Browse"
        '
        'lblInputFilePath
        '
        Me.lblInputFilePath.Location = New System.Drawing.Point(16, 16)
        Me.lblInputFilePath.Name = "lblInputFilePath"
        Me.lblInputFilePath.Size = New System.Drawing.Size(80, 16)
        Me.lblInputFilePath.TabIndex = 0
        Me.lblInputFilePath.Text = "Input file path"
        '
        'txtInputFilePath
        '
        Me.txtInputFilePath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtInputFilePath.Location = New System.Drawing.Point(16, 32)
        Me.txtInputFilePath.Name = "txtInputFilePath"
        Me.txtInputFilePath.Size = New System.Drawing.Size(416, 20)
        Me.txtInputFilePath.TabIndex = 1
        Me.txtInputFilePath.Text = "Input file"
        '
        'fraControls
        '
        Me.fraControls.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.fraControls.Controls.Add(Me.pbarProgress)
        Me.fraControls.Location = New System.Drawing.Point(264, 136)
        Me.fraControls.Name = "fraControls"
        Me.fraControls.Size = New System.Drawing.Size(247, 48)
        Me.fraControls.TabIndex = 3
        Me.fraControls.TabStop = False
        Me.fraControls.Text = "Progress"
        '
        'pbarProgress
        '
        Me.pbarProgress.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pbarProgress.Location = New System.Drawing.Point(8, 16)
        Me.pbarProgress.Maximum = 100
        Me.pbarProgress.Minimum = 0
        Me.pbarProgress.Name = "pbarProgress"
        Me.pbarProgress.Size = New System.Drawing.Size(224, 24)
        Me.pbarProgress.TabIndex = 2
        Me.pbarProgress.Value = 0
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
        'mnuHelpAbout
        '
        Me.mnuHelpAbout.Index = 2
        Me.mnuHelpAbout.Text = "&About"
        '
        'cmdExit
        '
        Me.cmdExit.Location = New System.Drawing.Point(360, 200)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(56, 20)
        Me.cmdExit.TabIndex = 7
        Me.cmdExit.Text = "E&xit"
        '
        'cmdStart
        '
        Me.cmdStart.Location = New System.Drawing.Point(272, 200)
        Me.cmdStart.Name = "cmdStart"
        Me.cmdStart.Size = New System.Drawing.Size(56, 20)
        Me.cmdStart.TabIndex = 6
        Me.cmdStart.Text = "&Start"
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
        'frmListPOR
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(520, 290)
        Me.Controls.Add(Me.cmdExit)
        Me.Controls.Add(Me.cmdStart)
        Me.Controls.Add(Me.fraControls)
        Me.Controls.Add(Me.fraFilePaths)
        Me.Controls.Add(Me.fraOptions)
        Me.Menu = Me.MainMenu1
        Me.MinimumSize = New System.Drawing.Size(475, 0)
        Me.Name = "frmListPOR"
        Me.Text = "ListPOR - List Parser for Outlier Removal"
        Me.fraOptions.ResumeLayout(False)
        Me.fraFilePaths.ResumeLayout(False)
        Me.fraControls.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Module-level Variables"
    Private Const XML_SETTINGS_FILENAME As String = "ListPORSettings.xml"
    Private Const XML_FILE_OPTIONS_SECTION As String = "Options"

    Private mProcessing As Boolean
    Private mAbortProcessing As Boolean

    Private WithEvents mListPOR As clsListPOR

#End Region

    Private Function BrowseForFile(ByVal strCurrentPath As String, Optional ByVal strDialogTitle As String = "Select file") As String

        Dim Dialog As SaveFileDialog
        Dim strCandidateParentFolderName As String

        Dim Result As Integer

        Dialog = New SaveFileDialog
        With Dialog
            .Filter = "Text (*.txt)|*.txt|Comma separated (*.csv)|*.csv|All Files (*.*)|*.*"

            If strCurrentPath.Length > 0 Then
                strCandidateParentFolderName = System.IO.Path.GetDirectoryName(strCurrentPath)

                If System.IO.File.Exists(strCurrentPath) Then
                    .FileName = strCurrentPath
                    .InitialDirectory = strCandidateParentFolderName
                ElseIf System.IO.Directory.Exists(strCurrentPath) Then
                    .InitialDirectory = strCurrentPath
                ElseIf System.IO.Directory.Exists(strCandidateParentFolderName) Then
                    .InitialDirectory = strCandidateParentFolderName
                    .FileName = System.IO.Path.GetFileName(strCurrentPath)
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
            MsgBox("Invalid filename: " & ex.Message, MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Error")
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

    Private Sub PositionControls()
        'Dim intDesiredLocation As Integer

        'Try
        '    ' The following could be used to Center the controls
        '    With cmdStart
        '        intDesiredLocation = pnlFiles.Width / 2 - .Width - 10
        '        If intDesiredLocation < 1 Then intDesiredLocation = 1
        '        cmdStart.Left = intDesiredLocation
        '    End With

        '    With cmdExit
        '        .Left = cmdStart.Left + cmdStart.Width + 20
        '    End With
        'Catch ex As Exception
        '    Debug.Assert(False, "Error positioning controls")
        'End Try

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

        strMessage &= "This is version " & System.Windows.Forms.Application.ProductVersion & " (" & PROGRAM_DATE & ")" & ControlChars.NewLine & ControlChars.NewLine

        strMessage &= "E-mail: matthew.monroe@pnl.gov or matt@alchemistmatt.com" & ControlChars.NewLine
        strMessage &= "Website: http://ncrr.pnl.gov/ or http://www.sysbio.org/resources/staff/" & ControlChars.NewLine & ControlChars.NewLine

        strMessage &= "Licensed under the Apache License, Version 2.0; you may not use this file except in compliance with the License.  "
        strMessage &= "You may obtain a copy of the License at http://www.apache.org/licenses/LICENSE-2.0" & ControlChars.NewLine & ControlChars.NewLine

        strMessage &= "Notice: This computer software was prepared by Battelle Memorial Institute, "
        strMessage &= "hereinafter the Contractor, under Contract No. DE-AC05-76RL0 1830 with the "
        strMessage &= "Department of Energy (DOE).  All rights in the computer software are reserved "
        strMessage &= "by DOE on behalf of the United States Government and the Contractor as "
        strMessage &= "provided in the Contract.  NEITHER THE GOVERNMENT NOR THE CONTRACTOR MAKES ANY "
        strMessage &= "WARRANTY, EXPRESS OR IMPLIED, OR ASSUMES ANY LIABILITY FOR THE USE OF THIS "
        strMessage &= "SOFTWARE.  This notice including this sentence must appear on any copies of "
        strMessage &= "this computer software." & ControlChars.NewLine

        Windows.Forms.MessageBox.Show(strMessage, "About", MessageBoxButtons.OK, MessageBoxIcon.Information)

    End Sub

    Private Sub ShowOverview()
        Dim strMessage As String

        strMessage = String.Empty
        strMessage &= "This program will read the input file and examine the values to remove outliers. Typically, the first column will contain key information (text or numbers) on which to group the data. "
        strMessage &= "The second column will contain the values to examine for outliers, examining the data by groups.  If additional columns are present, they will be written to the output file along with the key and value columns. "
        strMessage &= "Alternatively, if you provide a file with only one column of data, then all of the values will be examined en masse and the outliers removed.  If you do not provide an output file path, then the input file's path will be used, but with .filtered appended."

        Windows.Forms.MessageBox.Show(strMessage, "Overview", MessageBoxButtons.OK, MessageBoxIcon.Information)

    End Sub
    Private Sub ValidateFilesAndProcess()
        Dim strInputFilePath As String
        Dim strOutputFilePath As String

        Dim intErrorCode As Integer

        Try
            strInputFilePath = txtInputFilePath.Text.Trim
            strOutputFilePath = txtOutputFilePath.Text.Trim

            If strInputFilePath Is Nothing OrElse strInputFilePath.Length = 0 Then
                MsgBox("Please enter a valid input file path", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Missing Input File Path")
            Else
                If Not System.IO.File.Exists(strInputFilePath) Then
                    MsgBox("File not found: " & strInputFilePath, MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Missing Input File Path")
                Else
                    If strInputFilePath Is Nothing OrElse strInputFilePath.Length = 0 Then
                        MsgBox("Please enter a valid output file path", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Missing Output File Path")
                    ElseIf strInputFilePath = strOutputFilePath Then
                        MsgBox("The input and output file paths cannot be identical", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Identical Paths")
                    Else
                        Try
                            mProcessing = True
                            mAbortProcessing = False

                            Me.Cursor = Windows.Forms.Cursors.WaitCursor
                            pbarProgress.InitializeProgressBar(0, 100, True)
                            cmdStart.Text = "&Abort"

                            With mListPOR
                                .AssumeSortedInputFile = False
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
                            Me.Cursor = Windows.Forms.Cursors.Default
                            cmdStart.Text = "&Start"
                        End Try

                        If intErrorCode = 0 Then
                            MsgBox("Processing complete", MsgBoxStyle.Information Or MsgBoxStyle.OKOnly, "Done")
                        Else
                            MsgBox("Error while processing: " & mListPOR.GetErrorMessage(), MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Error")
                        End If
                    End If
                End If
            End If

        Catch ex As Exception
            MsgBox("Error validating the input and output file paths", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Path Error")
        End Try

    End Sub

    Private Sub ValidateDataTransformOptions(ByVal blnFavorNaturalLog As Boolean)
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

            strFilePath = System.IO.Path.GetDirectoryName(Application.ExecutablePath)
            strFilePath = System.IO.Path.Combine(strFilePath, XML_SETTINGS_FILENAME)

            Dim objXmlFile As New PRISM.Files.XmlSettingsFileAccessor

            If objXmlFile.LoadSettings(strFilePath) Then
                txtInputFilePath.Text = objXmlFile.GetParam(XML_FILE_OPTIONS_SECTION, "InputFilePath", txtInputFilePath.Text)
                txtOutputFilePath.Text = objXmlFile.GetParam(XML_FILE_OPTIONS_SECTION, "OutputFilePath", txtOutputFilePath.Text)

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
            MsgBox("Error loading settings from XML Settings file:" & ControlChars.NewLine & ex.Message, MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Error")
        End Try

    End Sub

    Private Sub XMLSaveSettings()
        Dim blnSuccess As Boolean
        Dim strFilePath As String

        Try
            strFilePath = System.IO.Path.GetDirectoryName(Application.ExecutablePath)
            strFilePath = System.IO.Path.Combine(strFilePath, XML_SETTINGS_FILENAME)

            Dim objXmlFile As New PRISM.Files.XmlSettingsFileAccessor

            If objXmlFile.LoadSettings(strFilePath) Then

                blnSuccess = objXmlFile.SetParam(XML_FILE_OPTIONS_SECTION, "InputFilePath", txtInputFilePath.Text)

                If Not blnSuccess Then
                    Throw New System.Exception("Unknown error while setting a value.")
                Else
                    ' Continue saving settings

                    objXmlFile.SetParam(XML_FILE_OPTIONS_SECTION, "OutputFilePath", txtOutputFilePath.Text)

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
                        Throw New System.Exception("Unknown error while saving the file.")
                    End If
                End If

            End If

        Catch ex As Exception
            MsgBox("Error saving settings to XML Settings file:" & ControlChars.NewLine & ex.Message, MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Error")
        End Try

    End Sub

    Private Sub chkUseSymmetricValues_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkUseSymmetricValues.CheckedChanged
        ValidateDataTransformOptions(False)
    End Sub

    Private Sub chkUseNaturalLogValues_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkUseNaturalLogValues.CheckedChanged
        ValidateDataTransformOptions(True)
    End Sub

    Private Sub cmdBrowseForInputFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowseForInputFile.Click
        SelectInputFile()
    End Sub

    Private Sub cmdBrowseForOutputFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowseForOutputFile.Click
        SelectOutputFile()
    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        ExitProgram()
    End Sub

    Private Sub cmdStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStart.Click
        If mProcessing Then
            mAbortProcessing = True
        Else
            ValidateFilesAndProcess()
        End If
    End Sub

    Private Sub frmListPOR_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        PopulateComboBoxes()
        InitializeControls(False)

        XMLReadSettings()
        PositionControls()
    End Sub

    Private Sub frmListPOR_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        PositionControls()
    End Sub

    Private Sub txtMinimumFinalDataPointCount_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMinimumFinalDataPointCount.TextChanged
        If txtMinimumFinalDataPointCount.Text.Length > 0 Then
            If Not IsNumeric(txtMinimumFinalDataPointCount.Text) Then
                txtMinimumFinalDataPointCount.Text = "3"
            End If
        End If
    End Sub

    Private Sub frmListPOR_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If mProcessing Then
            mAbortProcessing = True
        End If

        XMLSaveSettings()
    End Sub

    Private Sub mListPOR_ProgressUpdate() Handles mListPOR.ProgressUpdate
        Try
            pbarProgress.Value = mListPOR.PercentComplete
            Application.DoEvents()

            If mAbortProcessing Then
                mListPOR.AbortProcessing()
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub txtMinimumFinalDataPointCount_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMinimumFinalDataPointCount.LostFocus
        Try
            If CType(txtMinimumFinalDataPointCount.Text, Integer) < 2 Then
                txtMinimumFinalDataPointCount.Text = "3"
            End If
        Catch ex As Exception
            txtMinimumFinalDataPointCount.Text = "3"
        End Try
    End Sub

    Private Sub mnuEditResetToDefaults_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuEditResetToDefaults.Click
        If Not mProcessing Then
            InitializeControls(True)
        End If
    End Sub

    Private Sub mnuFileSelectInputFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileSelectInputFile.Click
        SelectInputFile()
    End Sub

    Private Sub mnuFileSelectOutputFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileSelectOutputFile.Click
        SelectOutputFile()
    End Sub

    Private Sub mnuFileExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuFileExit.Click
        ExitProgram()
    End Sub

    Private Sub mnuHelpAbout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuHelpAbout.Click
        ShowAboutBox()
    End Sub

    Private Sub mnuHelpOverview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuHelpOverview.Click
        ShowOverview()
    End Sub
End Class
