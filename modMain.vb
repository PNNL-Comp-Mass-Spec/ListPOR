Option Strict On

Imports System.IO
Imports System.Reflection
Imports System.Threading
Imports PRISM

' -------------------------------------------------------------------------------
' Written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA)
' Program started August 13, 2004

' E-mail: matthew.monroe@pnnl.gov or proteomics@pnnl.gov
' Website: https://github.com/PNNL-Comp-Mass-Spec/ or https://panomics.pnnl.gov/ or https://www.pnnl.gov/integrative-omics
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

''' <summary>
'''  Calls clsListPOR at the console or shows frmListPOR
''' </summary>
Module modMain

    Private Declare Auto Function ShowWindow Lib "user32.dll" (hWnd As IntPtr, nCmdShow As Integer) As Boolean
    Private Declare Auto Function GetConsoleWindow Lib "kernel32.dll" () As IntPtr
    Private Const SW_HIDE As Integer = 0
    Private Const SW_SHOW As Integer = 5

    Public Const PROGRAM_DATE As String = "October 4, 2021"

    Private WithEvents mListPORClass As clsListPOR
    Private mInputFilePath As String
    Private mOutputFilePath As String

    ''' <summary>
    ''' Entry method
    ''' </summary>
    ''' <returns>0 if no error, error code if an error</returns>
    Public Function Main() As Integer

        mInputFilePath = String.Empty
        mOutputFilePath = String.Empty


        mListPORClass = New clsListPOR With {
                .AssumeSortedInputFile = False,
                .RemoveMultipleValues = True,
                .ConfidenceLevel = clsGrubbsTestOutlierFilter.eclConfidenceLevelConstants.e95Pct,
                .MinFinalValueCount = 3,
                .ColumnCountOverride = 0
                }
        Try
            Dim blnProceed = False
            Dim blnShowForm = False

            Dim objParseCommandLine As New clsParseCommandLine()

            If objParseCommandLine.ParseCommandLine Then
                If SetOptionsUsingCommandLineParameters(objParseCommandLine) Then blnProceed = True
            Else
                ' Show form
                blnShowForm = True
            End If

            If blnShowForm And Not objParseCommandLine.NeedToShowHelp Then
                ' Hide the console
                Dim hWndConsole As IntPtr
                hWndConsole = GetConsoleWindow()
                ShowWindow(hWndConsole, SW_HIDE)

                Dim mainWindow = New frmListPOR()
                mainWindow.ShowDialog()

                ShowWindow(hWndConsole, SW_SHOW)

            ElseIf Not blnProceed OrElse objParseCommandLine.NeedToShowHelp OrElse String.IsNullOrWhiteSpace(mInputFilePath) Then
                ShowProgramHelp()
                Return -1
            Else

                If String.IsNullOrWhiteSpace(mOutputFilePath) Then
                    mOutputFilePath = clsListPOR.AutoGenerateOutputFileName(mInputFilePath)
                End If

                Dim returnCode = mListPORClass.RemoveOutliersFromListInFile(mInputFilePath, mOutputFilePath)

                If returnCode <> 0 Then
                    ConsoleMsgUtils.ShowError("Error while processing: ReturnCode = " & returnCode)
                    Thread.Sleep(1500)
                    Return returnCode
                End If
            End If

            Thread.Sleep(1500)
            Return 0
        Catch ex As Exception
            ConsoleMsgUtils.ShowError("Error while processing", ex)
            Thread.Sleep(1500)
            Return -1
        End Try

    End Function

    Private Function SetOptionsUsingCommandLineParameters(objParseCommandLine As clsParseCommandLine) As Boolean
        ' Returns True if no problems; otherwise, returns false

        Dim strValue = String.Empty
        Dim strValidParameters = New String() {"I", "O", "Sorted", "L", "M", "C", "Conf"}

        Try
            ' Make sure no invalid parameters are present
            If objParseCommandLine.InvalidParametersPresent(strValidParameters) Then
                Return False
            Else

                ' Query objParseCommandLine to see if various parameters are present
                With objParseCommandLine

                    If .NonSwitchParameterCount > 0 Then
                        mInputFilePath = .RetrieveNonSwitchParameter(0)
                    Else
                        If .RetrieveValueForParameter("I", strValue) Then mInputFilePath = strValue
                    End If

                    If .NonSwitchParameterCount > 1 Then
                        mOutputFilePath = .RetrieveNonSwitchParameter(1)
                    Else
                        If .RetrieveValueForParameter("O", strValue) Then mOutputFilePath = strValue
                    End If

                    If .RetrieveValueForParameter("Sorted", strValue) Then mListPORClass.AssumeSortedInputFile = True

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

                    If .RetrieveValueForParameter("Conf", strValue) Then
                        If IsNumeric(strValue) Then
                            Dim confidenceInterval = CType(strValue, Integer)
                            Select Case confidenceInterval
                                Case 95
                                    mListPORClass.ConfidenceLevel = clsGrubbsTestOutlierFilter.eclConfidenceLevelConstants.e95Pct
                                Case 97
                                    mListPORClass.ConfidenceLevel = clsGrubbsTestOutlierFilter.eclConfidenceLevelConstants.e97Pct
                                Case 99
                                    mListPORClass.ConfidenceLevel = clsGrubbsTestOutlierFilter.eclConfidenceLevelConstants.e99Pct
                            End Select

                        End If
                    End If

                End With

                Return True
            End If

        Catch ex As Exception
            ConsoleMsgUtils.ShowError("Error parsing the command line parameters", ex)
            Thread.Sleep(1500)
            Return False
        End Try

    End Function

    Private Sub ShowProgramHelp()

        Try

            Console.WriteLine("ListPOR (List Parser for Outlier Removal)")
            Console.WriteLine("Reads a tab-delimited file of groups of data and removes outliers from each group of data points.")
            Console.WriteLine()
            Console.WriteLine("Program syntax:" & Path.GetFileName(Assembly.GetExecutingAssembly().Location))
            Console.WriteLine("/I:InputFilePath [/O:OutputFilePath] [/Sorted] [/L] [/M:MinimumFinalDataPointCount] [/C:ColumnCount] [/Conf:#]")
            Console.WriteLine()
            Console.WriteLine("The input file is a tab-delimited data file with one or more columns of data; specify it using /I or by simply using the file path (use double quotes if spaces in the path)")
            Console.WriteLine()
            Console.WriteLine("The output file path is optional. If omitted, the output file will be created in the same folder as the input file, but with _filtered appended to the name")
            Console.WriteLine()
            Console.WriteLine("/Sorted indicates that the input file is already sorted by group. This allows very large files to be processed, since the entire file does not need to be cached in memory. It is also useful for obtaining a filtered file with data in the exact same order as the input file.")
            Console.WriteLine()
            Console.WriteLine("/L will cause the program to convert the data to symmetric values, prior to looking for outliers. This is appropriate for data where a value of 1 is unchanged, >1 is an increase, and <1 is a decrease. This is not appropriate for data with values of 0 or less than 0.")
            Console.WriteLine()
            Console.WriteLine("/M is the minimum number of data points that must remain in the group. It cannot be less than 3")
            Console.WriteLine()
            Console.WriteLine("/C:1 can be used to indicate that there is only 1 column of data to analyze; if other columns of text are present after the first column, they will be written to the output file, but will not be considered for outlier removal. Use /C:2 to specify that there are two columns to be examined: a Key column and a Value column. Again, additional columns will be written to disk, but not utilized for comparison purposes. ")
            Console.WriteLine()
            Console.WriteLine("/Conf can be used to specify the confidence level; options are")
            Console.WriteLine("  /Conf:95")
            Console.WriteLine("  /Conf:97")
            Console.WriteLine("  /Conf:99")
            Console.WriteLine()
            Console.WriteLine()
            Console.WriteLine("Program written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA) in 2004")

            Console.WriteLine("This is version " & Application.ProductVersion & " (" & PROGRAM_DATE & ")")

            Console.WriteLine("E-mail: matthew.monroe@pnnl.gov or proteomics@pnnl.gov")
            Console.WriteLine("Website: https://github.com/PNNL-Comp-Mass-Spec/ or https://panomics.pnnl.gov/ or https://www.pnnl.gov/integrative-omics")
            Console.WriteLine()

            Console.WriteLine("Licensed under the Apache License, Version 2.0; you may not use this file except in compliance with the License.")
            Console.WriteLine("You may obtain a copy of the License at http://www.apache.org/licenses/LICENSE-2.0")

        Catch ex As Exception
            ConsoleMsgUtils.ShowError("Error displaying the program syntax", ex)
        End Try

        Thread.Sleep(1500)

    End Sub

    Private Sub mListPORClass_DebugEvent(message As String) Handles mListPORClass.DebugEvent
        ConsoleMsgUtils.ShowDebug(message)
    End Sub

    Private Sub mListPORClass_ErrorEvent(message As String, ex As Exception) Handles mListPORClass.ErrorEvent
        ConsoleMsgUtils.ShowError(message)
    End Sub

    Private Sub mListPORClass_StatusEvent(message As String) Handles mListPORClass.StatusEvent
        Console.WriteLine(message)
    End Sub

    Private Sub mListPORClass_WarningEvent(message As String) Handles mListPORClass.WarningEvent
        ConsoleMsgUtils.ShowWarning(message)
    End Sub
End Module