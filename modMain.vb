Option Strict On

' Wrapper functions to call clsListPOR or show frmListPOR
'
' -------------------------------------------------------------------------------
' Written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA)
' Program started August 13, 2004

' E-mail: matthew.monroe@pnl.gov or matt@alchemistmatt.com
' Website: http://ncrr.pnl.gov/ or http://www.sysbio.org/resources/staff/
' -------------------------------------------------------------------------------
' 
' Licensed under the Apache License, Version 2.0; you may not use this file except
' in compliance with the License.  You may obtain a copy of the License at 
' http://www.apache.org/licenses/LICENSE-2.0
'
' Notice: This computer software was prepared by Battelle Memorial Institute, 
' hereinafter the Contractor, under Contract No. DE-AC05-76RL0 1830 with the 
' Department of Energy (DOE).  All rights in the computer software are reserved 
' by DOE on behalf of the United States Government and the Contractor as 
' provided in the Contract.  NEITHER THE GOVERNMENT NOR THE CONTRACTOR MAKES ANY 
' WARRANTY, EXPRESS OR IMPLIED, OR ASSUMES ANY LIABILITY FOR THE USE OF THIS 
' SOFTWARE.  This notice including this sentence must appear on any copies of 
' this computer software.

Module modMain

    Public Const PROGRAM_DATE As String = "October 5, 2005"

    Private mMainWindow As New frmListPOR

    Private mInputFilePath As String
    Private mOutputFilePath As String

    Private mListPORClass As clsListPOR

    Private mQuietMode As Boolean

    Public Function Main() As Integer
        ' Returns 0 if no error, error code if an error

        Dim intReturnCode As Integer
        Dim objParseCommandLine As New SharedVBNetRoutines.clsParseCommandLine
        Dim blnProceed As Boolean
        Dim blnShowForm As Boolean

        intReturnCode = 0
        mInputFilePath = ""
        mOutputFilePath = ""

        mListPORClass = New clsListPOR

        mListPORClass.AssumeSortedInputFile = False
        mListPORClass.RemoveMultipleValues = True
        mListPORClass.ConfidenceLevel = clsGrubbsTestOutlierFilter.eclConfidenceLevelConstants.e95Pct
        mListPORClass.MinFinalValueCount = 3
        mListPORClass.ColumnCountOverride = 0

        mQuietMode = False

        Try
            blnProceed = False
            blnShowForm = False
            If objParseCommandLine.ParseCommandLine Then
                If SetOptionsUsingCommandLineParameters(objParseCommandLine) Then blnProceed = True
            Else
                ' Show form
                blnShowForm = True
            End If

            If blnShowForm And Not objParseCommandLine.NeedToShowHelp Then
                mMainWindow.ShowDialog()
            ElseIf Not blnProceed OrElse objParseCommandLine.NeedToShowHelp OrElse objParseCommandLine.ParameterCount = 0 OrElse mInputFilePath.Length = 0 Then
                ShowProgramHelp()
                intReturnCode = -1
            Else

                With mListPORClass

                    If mOutputFilePath.Length = 0 Then
                        mOutputFilePath = mInputFilePath & ".filtered"
                    End If

                    intReturnCode = .RemoveOutliersFromListInFile(mInputFilePath, mOutputFilePath)
                End With

                If intReturnCode <> 0 AndAlso Not mQuietMode Then
                    MsgBox("Error while processing: ReturnCode = " & intReturnCode.ToString, MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Error")
                End If
            End If

        Catch ex As Exception
            If mQuietMode Then
                Throw ex
            Else
                MsgBox("Error occurred: " & ControlChars.NewLine & ex.Message, MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Error")
            End If
            intReturnCode = -1
        End Try

        Return intReturnCode

    End Function

     Private Function SetOptionsUsingCommandLineParameters(ByVal objParseCommandLine As SharedVBNetRoutines.clsParseCommandLine) As Boolean
        ' Returns True if no problems; otherwise, returns false

        Dim strValue As String
        Dim strValidParameters() As String = New String() {"I", "O", "S", "L", "M", "C", "Q"}

        Try
            ' Make sure no invalid parameters are present
            If objParseCommandLine.InvalidParametersPresent(strValidParameters) Then
                Return False
            Else

                ' Query objParseCommandLine to see if various parameters are present
                With objParseCommandLine
                    If .RetrieveValueForParameter("I", strValue) Then mInputFilePath = strValue
                    If .RetrieveValueForParameter("O", strValue) Then mOutputFilePath = strValue
                    If .RetrieveValueForParameter("S", strValue) Then mListPORClass.AssumeSortedInputFile = True
                    If .RetrieveValueForParameter("L", strValue) Then mListPORClass.UseSymmetricValues = True
                    If .RetrieveValueForParameter("M", strValue) Then
                        If IsNumeric(strValue) Then
                            mListPORClass.MinFinalValueCount = CType(strValue, Integer)
                        End If
                    End If
                    If .RetrieveValueForParameter("C", strValue) Then
                        If IsNumeric(strValue) Then
                            mListPORClass.ColumnCountOverride = CType(strValue, Integer)
                        End If
                    End If
                    If .RetrieveValueForParameter("Q", strValue) Then mQuietMode = True
                End With

                Return True
            End If

        Catch ex As Exception
            If mQuietMode Then
                Throw New System.Exception("Error parsing the command line parameters", ex)
            Else
                MsgBox("Error parsing the command line parameters: " & ControlChars.NewLine & ex.Message, MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Error")
            End If
        End Try

    End Function

    Private Sub ShowProgramHelp()

        Dim strSyntax As String
        Dim ioPath As System.IO.Path

        Try

            strSyntax = "ListPOR (List Parser for Outlier Removal)" & ControlChars.NewLine
            strSyntax &= "Reads a tab-delimeted file of groups of data and removes outliers from each group of data points." & ControlChars.NewLine
            strSyntax &= "Program syntax:" & ControlChars.NewLine & ioPath.GetFileName(System.Reflection.Assembly.GetExecutingAssembly().Location)
            strSyntax &= " /I:InputFilePath [/O:OutputFilePath] [/S] [/L] [/M:MinimumFinalDataPointCount] [/C:ColumnCount] [/Q]" & ControlChars.NewLine & ControlChars.NewLine
            strSyntax &= "The output file path is optional.  If omitted, the output file will be created in the same folder as the input file, but with the extension '.filtered' added." & ControlChars.NewLine
            strSyntax &= " /S indicates that the input file is already sorted by group.  This allows very large files to be processed, since the entire file does not need to be cached in memory." & ControlChars.NewLine
            strSyntax &= " /L will cause the program to shifted symmetric values, prior to looking for outliers.  This is appropriate for data where a value of 1 is unchanged, >1 is an increase, and <1 is a decrease.  This is not appropriate for data with values of 0 or less than 0." & ControlChars.NewLine
            strSyntax &= " /M is the minimum number of data points that must remain in the group.  It cannot be less than 3" & ControlChars.NewLine
            strSyntax &= " /C:1 can be used to indicate that there is only 1 column of data to analyze; if other columns of text are present after the first column, they will be written to the output file, but will not be considered for outlier removal.  Use /C:2 to specify that there are two columns to be examined: a Key column and a Value column.  Again, additional columns will be written to disk, but not utilized for comparison purposes. " & ControlChars.NewLine
            strSyntax &= "The optional /Q switch will suppress all error messages." & ControlChars.NewLine  & ControlChars.NewLine

            strSyntax &= "Program written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA) in 2004" & ControlChars.NewLine & ControlChars.NewLine

            strSyntax &= "This is version " & System.Windows.Forms.Application.ProductVersion & " (" & PROGRAM_DATE & ")" & ControlChars.NewLine & ControlChars.NewLine

            strSyntax &= "E-mail: matthew.monroe@pnl.gov or matt@alchemistmatt.com" & ControlChars.NewLine
            strSyntax &= "Website: http://ncrr.pnl.gov/ or http://www.sysbio.org/resources/staff/" & ControlChars.NewLine & ControlChars.NewLine

            strSyntax &= "Licensed under the Apache License, Version 2.0; you may not use this file except in compliance with the License.  "
            strSyntax &= "You may obtain a copy of the License at http://www.apache.org/licenses/LICENSE-2.0" & ControlChars.NewLine & ControlChars.NewLine

            strSyntax &= "Notice: This computer software was prepared by Battelle Memorial Institute, "
            strSyntax &= "hereinafter the Contractor, under Contract No. DE-AC05-76RL0 1830 with the "
            strSyntax &= "Department of Energy (DOE).  All rights in the computer software are reserved "
            strSyntax &= "by DOE on behalf of the United States Government and the Contractor as "
            strSyntax &= "provided in the Contract.  NEITHER THE GOVERNMENT NOR THE CONTRACTOR MAKES ANY "
            strSyntax &= "WARRANTY, EXPRESS OR IMPLIED, OR ASSUMES ANY LIABILITY FOR THE USE OF THIS "
            strSyntax &= "SOFTWARE.  This notice including this sentence must appear on any copies of "
            strSyntax &= "this computer software." & ControlChars.NewLine


            If Not mQuietMode Then
                MsgBox(strSyntax, MsgBoxStyle.Information Or MsgBoxStyle.OKOnly, "Syntax")
            End If

        Catch ex As Exception
            If mQuietMode Then
                Throw New System.Exception("Error displaying the program syntax", ex)
            Else
                MsgBox("Error displaying the program syntax: " & ControlChars.NewLine & ex.Message, MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly, "Error")
            End If
        End Try

    End Sub

End Module