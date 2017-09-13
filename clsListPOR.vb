Option Strict On

Imports System.Collections.Generic
Imports System.IO
Imports System.Linq

''' <summary>
''' ListPOR (List Parser for Outlier Removal)
''' Parses an input file looking for blocks of related data
''' For each block, filters the data using clsGrubbsTestOutlierFilter
''' </summary>
Public Class clsListPOR
    Inherits PRISM.clsEventNotifier

#Region "Enums"

    Public Enum eListPORErrorCodeCodes
        NoError = 0
        ErrorReadingInputFile = 1
        ErrorWritingOutputFile = 2
        ErrorWithGrubbsFilterClass = 4
        ProcessingAborted = 8
        UnspecifiedError = -1
    End Enum
#End Region

#Region "Module-level Variables"
    Private mAssumeSortedInputFile As Boolean
    Private mConfidenceLevel As clsGrubbsTestOutlierFilter.eclConfidenceLevelConstants
    Private mMinFinalValueCount As Integer
    Private mMostRecentOutputFile As String
    Private mRemoveMultipleValues As Boolean
    Private mUseSymmetricValues As Boolean
    Private mUseNaturalLogValues As Boolean

    Private mAppendGroupAverage As Boolean
    Private mColumnCountOverride As Integer
    Private mLocalErrorCode As eListPORErrorCodeCodes

    Private mFileLengthBytes As Long
    Private mFileBytesRead As Long
    Private mPercentComplete As Single
    Private mAbortProcessing As Boolean

#End Region

#Region "Interface functions"

    Public Property AppendGroupAverage As Boolean
        Get
            Return mAppendGroupAverage
        End Get
        Set
            mAppendGroupAverage = Value
        End Set
    End Property

    Public Property AssumeSortedInputFile As Boolean
        Get
            Return mAssumeSortedInputFile
        End Get
        Set
            mAssumeSortedInputFile = Value
        End Set
    End Property

    Public Property ConfidenceLevel As clsGrubbsTestOutlierFilter.eclConfidenceLevelConstants
        Get
            Return mConfidenceLevel
        End Get
        Set
            mConfidenceLevel = Value
        End Set
    End Property

    Public Property ColumnCountOverride As Integer
        Get
            Return mColumnCountOverride
        End Get
        Set
            ' Value can only be 0, 1 or 2
            If Value = 1 Or Value = 2 Then
                mColumnCountOverride = Value
            Else
                mColumnCountOverride = 0
            End If
        End Set
    End Property

    Public ReadOnly Property ErrorCode As eListPORErrorCodeCodes
        Get
            Return mLocalErrorCode
        End Get
    End Property

    Public Property MinFinalValueCount As Integer
        Get
            Return mMinFinalValueCount
        End Get
        Set
            If Value < 2 Then Value = 2
            mMinFinalValueCount = Value
        End Set
    End Property

    Public ReadOnly Property MostRecentOutputFile As String
        Get
            Return mMostRecentOutputFile
        End Get
    End Property

    Public ReadOnly Property PercentComplete As Single
        Get
            Return mPercentComplete
        End Get
    End Property

    Public Property RemoveMultipleValues As Boolean
        Get
            Return mRemoveMultipleValues
        End Get
        Set
            mRemoveMultipleValues = Value
        End Set
    End Property

    Public Property UseNaturalLogValues As Boolean
        Get
            Return mUseNaturalLogValues
        End Get
        Set
            mUseNaturalLogValues = Value
        End Set
    End Property

    Public Property UseSymmetricValues As Boolean
        Get
            Return mUseSymmetricValues
        End Get
        Set
            mUseSymmetricValues = Value
        End Set
    End Property
#End Region

    ''' <summary>
    ''' Constructor
    ''' </summary>
    Public Sub New()
        mAssumeSortedInputFile = False
        mRemoveMultipleValues = True
        mConfidenceLevel = clsGrubbsTestOutlierFilter.eclConfidenceLevelConstants.e95Pct
        MinFinalValueCount = 3
        ColumnCountOverride = 0
        mUseSymmetricValues = False
        UseNaturalLogValues = False
        mAppendGroupAverage = True

        mLocalErrorCode = eListPORErrorCodeCodes.NoError
    End Sub

    Public Sub AbortProcessing()
        mAbortProcessing = True
    End Sub

    Public Shared Function AutoGenerateOutputFileName(inputFilePath As String) As String
        Dim inputFile = New FileInfo(inputFilePath)
        Dim outputFilePath = Path.Combine(inputFile.DirectoryName, Path.GetFileNameWithoutExtension(inputFile.Name) & "_filtered" & inputFile.Extension)
        Return outputFilePath
    End Function

    Public Function GetErrorMessage() As String
        ' Returns "" if no error

        If mLocalErrorCode = eListPORErrorCodeCodes.NoError Then
            Return String.Empty
        ElseIf (mLocalErrorCode And eListPORErrorCodeCodes.ErrorReadingInputFile) = eListPORErrorCodeCodes.ErrorReadingInputFile Then
            Return "Error reading input file"
        ElseIf (mLocalErrorCode And eListPORErrorCodeCodes.ErrorWritingOutputFile) = eListPORErrorCodeCodes.ErrorWritingOutputFile Then
            Return "Error writing output file"
        ElseIf (mLocalErrorCode And eListPORErrorCodeCodes.ErrorWithGrubbsFilterClass) = eListPORErrorCodeCodes.ErrorWithGrubbsFilterClass Then
            Return "Error with Grubbs filter class"
        ElseIf (mLocalErrorCode And eListPORErrorCodeCodes.ProcessingAborted) = eListPORErrorCodeCodes.ProcessingAborted Then
            Return "Processing aborted"
        Else
            Return "Unknown error state"
        End If

    End Function

    Private Function NaturalLogValueToNormalValue(dblNaturalLogValue As Double) As Double
        Try
            Return Math.Exp(dblNaturalLogValue)
        Catch
            Return 0
        End Try

    End Function

    ''' <summary>
    ''' Converts ratio-based value to natural log
    ''' </summary>
    ''' <param name="dblValue"></param>
    ''' <returns></returns>
    Private Function NormalValueToNaturalLog(dblValue As Double) As Double
        Try
            Return Math.Log(dblValue)
        Catch
            Return 0
        End Try
    End Function

    ''' <summary>
    ''' Converts ratio-based value to shifted symmetric value
    ''' </summary>
    ''' <param name="dblValue"></param>
    ''' <returns></returns>
    Private Function NormalValueToSymmetricValue(dblValue As Double) As Double
        Try
            If dblValue >= 1 Then
                Return dblValue - 1
            Else
                Return 1 - (1 / dblValue)
            End If
        Catch
            Return 0
        End Try
    End Function

    Private Sub ParseBlock(
      objOutlierFilter As clsGrubbsTestOutlierFilter,
      lstData As List(Of clsFileData),
      columnCount As Integer,
      swOutFile As TextWriter)

        If lstData.Count = 0 Then
            ' No data; nothing to do
            Exit Sub
        End If

        Try
            Dim lstDataValues = New List(Of Double)

            ' Copy the values to examine into dblValues
            For Each dataPoint In lstData

                Dim dblValue = dataPoint.ValueDbl

                If mUseSymmetricValues OrElse mUseNaturalLogValues Then
                    If dblValue > 0 Then
                        If mUseNaturalLogValues Then
                            dblValue = NormalValueToNaturalLog(dblValue)
                        Else
                            dblValue = NormalValueToSymmetricValue(dblValue)
                        End If
                    Else
                        ' Invalid value for computing the symmetric value or logarithm
                        dblValue = 0
                    End If
                End If

                lstDataValues.Add(dblValue)
            Next

            Dim lstFilteredData = New List(Of clsFileData)
            Dim dblAverage As Double = 0

            If lstDataValues.Count <= mMinFinalValueCount Then
                ' Not enough data to remove outliers
                lstFilteredData = lstData
                dblAverage = MathNet.Numerics.Statistics.ArrayStatistics.Mean(lstDataValues.ToArray())
            Else
                ' Find outliers
                Dim outlierIndices As SortedSet(Of Integer) = Nothing

                If objOutlierFilter.FindOutliers(lstDataValues, outlierIndices) Then
                    ' Successful call to .FindOutliers

                    ' Remove the outlier data from lstFileData
                    ' (outlierIndices might be empty; that's OK, we need to compute the average)

                    Dim dblSum As Double = 0
                    For i = 0 To lstData.Count - 1
                        If outlierIndices.Contains(i) Then Continue For

                        lstFilteredData.Add(lstData(i))
                        dblSum += lstDataValues(i)
                    Next

                    If lstFilteredData.Count > 0 Then
                        ' Compute the average value for the values in lstFilteredData since we may append it as an additional column below
                        dblAverage = dblSum / lstFilteredData.Count
                    End If

                Else
                    ' Error removing outliers; simply re-write the input text to the output file
                    lstFilteredData = lstData
                    dblAverage = MathNet.Numerics.Statistics.ArrayStatistics.Mean(lstDataValues.ToArray())
                End If

            End If

            Dim blnAppendGroupAverage = mAppendGroupAverage And columnCount > 1

            If blnAppendGroupAverage AndAlso (mUseSymmetricValues OrElse mUseNaturalLogValues) Then
                If mUseNaturalLogValues Then
                    dblAverage = NaturalLogValueToNormalValue(dblAverage)
                Else
                    dblAverage = SymmetricValueToNormalValue(dblAverage)
                End If
            End If

            For Each item In lstFilteredData
                WriteBlockLine(swOutFile, columnCount, item, blnAppendGroupAverage, dblAverage)
            Next

        Catch ex As Exception
            Throw New Exception("Error parsing block", ex)
        End Try

    End Sub

    ''' <summary>
    ''' Reads a tab-delimeted text file (strSourcePath) and removes outliers from lists of numbers in column
    ''' </summary>
    ''' <param name="strSourcePath"></param>
    ''' <param name="strDestPath"></param>
    ''' <returns></returns>
    Public Function RemoveOutliersFromListInFile(strSourcePath As String, strDestPath As String) As Integer

        ' Modes of operation:
        ' Mode 1
        '   The text file has two columns
        '   The first column will be a Group Key column (text or numbers), with group key values repeated numerous times
        '   The second column will contain the values for each Group
        '   In this mode, the values for each group will be examined, and outliers for each group will be removed
        '
        ' Mode 2
        '   The text file has one column
        '   The data will be treated as one large Group, outliers will be found and removed
        '
        ' For both modes, the output will be saved in strDestPath (auto-generated name if empty)
        ' If mAssumeSortedInputFile is False, the entire file will be read into memory,
        '  the data will be sorted on the first column, then the outliers will be found

        ' If mAssumeSortedInputFile is True, only the data for the most recent group will be retained
        '  This is useful for parsing files will too much data to reside in memory
        '  It is also useful for retaining the order of the input data
        '

        Dim objOutlierFilter As clsGrubbsTestOutlierFilter

        Dim strColumnDelimeter As Char
        Dim strDelimList() As Char
        Dim strSplitLine() As String

        Dim strLineIn As String

        Dim columnCount As Integer
        Dim lstData = New List(Of clsFileData)

        Dim strHeaderLine As String

        Dim intIndexBlockStart As Integer

        Dim blnDataBlockReached As Boolean
        Dim blnAssumeFileIsSorted As Boolean

        Dim strDestPathToUse As String

        blnAssumeFileIsSorted = mAssumeSortedInputFile
        SetLocalErrorCode(eListPORErrorCodeCodes.NoError)

        mFileBytesRead = 0
        mFileLengthBytes = 1
        mAbortProcessing = False

        Try
            objOutlierFilter = New clsGrubbsTestOutlierFilter
            With objOutlierFilter
                .ConfidenceLevel() = mConfidenceLevel
                .MinFinalValueCount() = mMinFinalValueCount
                .RemoveMultipleValues = mRemoveMultipleValues
            End With
        Catch ex As Exception
            SetLocalErrorCode(eListPORErrorCodeCodes.ErrorWithGrubbsFilterClass)
            OnErrorEvent(GetErrorMessage())
            Return mLocalErrorCode
        End Try

        Try
            If String.IsNullOrWhiteSpace(strDestPath) OrElse String.Equals(strSourcePath, strDestPath, StringComparison.OrdinalIgnoreCase) Then
                strDestPathToUse = AutoGenerateOutputFileName(strSourcePath)
            Else
                strDestPathToUse = String.Copy(strDestPath)
            End If

            Dim sourceFile = New FileInfo(strSourcePath)
            Dim destFile = New FileInfo(strDestPathToUse)

            If String.Equals(sourceFile.FullName, destFile.FullName, StringComparison.OrdinalIgnoreCase) Then
                SetLocalErrorCode(eListPORErrorCodeCodes.ErrorWritingOutputFile)
                OnErrorEvent("Input and output file cannot be the same path: " & sourceFile.FullName)
                Return mLocalErrorCode
            End If

            mMostRecentOutputFile = destFile.FullName

            Using swOutFile = New StreamWriter(New FileStream(destFile.FullName, FileMode.Create, FileAccess.Write, FileShare.Read))

                ' Assume the column delimeter is a tab, unless the input file ends in .csv
                If String.Equals(sourceFile.Extension, ".csv", StringComparison.OrdinalIgnoreCase) Then
                    strColumnDelimeter = ","c
                Else
                    strColumnDelimeter = ControlChars.Tab
                End If
                strDelimList = New Char() {strColumnDelimeter}

                Dim percentCompleteAtStart As Integer
                Dim nextPercentComplete As Integer

                If blnAssumeFileIsSorted Then
                    percentCompleteAtStart = 0
                    nextPercentComplete = 100
                Else
                    percentCompleteAtStart = 0
                    nextPercentComplete = 95
                End If

                ' Open the input file
                Using srInFile = New StreamReader(New FileStream(sourceFile.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))

                    mFileLengthBytes = srInFile.BaseStream.Length

                    If mColumnCountOverride = 1 Or mColumnCountOverride = 2 Then
                        columnCount = mColumnCountOverride
                    Else
                        ' Set the column count to 0 for now, unless
                        ' We'll update it when we reach a header line or a data block line
                        columnCount = 0
                    End If

                    strHeaderLine = String.Empty
                    Do While Not srInFile.EndOfStream

                        strLineIn = srInFile.ReadLine()
                        mFileBytesRead += strLineIn.Length + 2

                        UpdatePercentComplete(mFileBytesRead, mFileLengthBytes, percentCompleteAtStart, nextPercentComplete)
                        OnProgressUpdate("Processing", mPercentComplete)

                        If mAbortProcessing Then Exit Do

                        If Not strLineIn Is Nothing AndAlso strLineIn.Length > 0 Then
                            ' Split the line, allowing at most 3 columns (if more than 3, then all data is lumped into the third column)
                            strSplitLine = strLineIn.Split(strDelimList, 3)

                            If Not blnDataBlockReached Then
                                ' Haven't found the data block yet, is this a header line?
                                If columnCount = 1 Then
                                    If strSplitLine.Length > 0 AndAlso IsNumeric(strSplitLine(0)) Then
                                        blnDataBlockReached = True
                                    Else
                                        strHeaderLine = strLineIn
                                    End If
                                ElseIf columnCount = 2 Then
                                    If strSplitLine.Length > 1 AndAlso IsNumeric(strSplitLine(1)) Then
                                        blnDataBlockReached = True
                                    Else
                                        strHeaderLine = strLineIn
                                    End If
                                ElseIf strSplitLine.Length >= 2 Then
                                    If Not IsNumeric(strSplitLine(0)) And IsNumeric(strSplitLine(1)) Then
                                        blnDataBlockReached = True
                                        columnCount = 2
                                    Else
                                        strHeaderLine = strLineIn
                                    End If
                                ElseIf strSplitLine.Length = 1 Then
                                    If IsNumeric(strSplitLine(0)) Then
                                        blnDataBlockReached = True
                                        columnCount = 1
                                    Else
                                        strHeaderLine = strLineIn
                                    End If
                                Else
                                    strHeaderLine = strLineIn
                                End If

                                If Not blnDataBlockReached Then
                                    ' Write out the header lines as we read them
                                    swOutFile.WriteLine(strHeaderLine)
                                End If
                            End If

                            If blnDataBlockReached Then
                                If columnCount = 0 Then
                                    ' Determine the number of columns of data

                                    If strSplitLine.Length >= 2 AndAlso IsNumeric(strSplitLine(1)) Then
                                        columnCount = 2
                                    ElseIf strSplitLine.Length = 1 And IsNumeric(strSplitLine(0)) Then
                                        columnCount = 1
                                    Else
                                        ' Assume columnCount = 1, even though the input file probably isn't formatted correctly
                                        Debug.Assert(False, "Input file does not have the expected number of columns of data")
                                        columnCount = 1
                                    End If
                                End If

                                Dim remainingCols As String
                                If strSplitLine.Length > 2 Then
                                    remainingCols = strSplitLine(2)
                                Else
                                    remainingCols = String.Empty
                                End If

                                If columnCount = 2 AndAlso strSplitLine.Length >= 2 AndAlso IsNumeric(strSplitLine(1)) Then
                                    Try
                                        Dim newDataPoint = New clsFileData(strSplitLine(0), strSplitLine(1), remainingCols)
                                        lstData.Add(newDataPoint)

                                    Catch ex As Exception
                                        ' Line parsing error; skip this line
                                    End Try
                                ElseIf columnCount = 2 AndAlso strSplitLine.Length = 1 AndAlso IsNumeric(strSplitLine(0)) Then
                                    Try
                                        Dim newDataPoint = New clsFileData(strSplitLine(0), "0", String.Empty)
                                        lstData.Add(newDataPoint)

                                    Catch ex As Exception
                                        ' Line parsing error; skip this line
                                    End Try
                                ElseIf columnCount = 1 AndAlso strSplitLine.Length >= 1 AndAlso IsNumeric(strSplitLine(0)) Then
                                    Try
                                        Dim newDataPoint = New clsFileData("A", strSplitLine(0), remainingCols)
                                        lstData.Add(newDataPoint)

                                    Catch ex As Exception
                                        ' Line parsing error; skip this line
                                    End Try
                                Else
                                    ' Ignore the line
                                End If

                                If blnAssumeFileIsSorted Then
                                    If lstData.Count > 1 Then
                                        If lstData(lstData.Count - 1).Key <> lstData(lstData.Count - 2).Key Then
                                            ' Parse this block
                                            ' Need to save the latest value in udtDataNextLine since it will be removed from lstData

                                            Dim nextDataPoint = lstData(lstData.Count - 1)

                                            ParseBlock(objOutlierFilter, lstData, columnCount, swOutFile)

                                            ' Reset lstData
                                            lstData.Clear()
                                            lstData.Add(nextDataPoint)
                                        End If
                                    End If
                                End If
                            End If

                        End If
                    Loop

                End Using

                If mAbortProcessing Then
                    SetLocalErrorCode(eListPORErrorCodeCodes.ProcessingAborted)
                    Return mLocalErrorCode
                End If

                Dim finalDataBlock = New List(Of clsFileData)

                If Not blnAssumeFileIsSorted Then

                    If lstData.Count > 0 Then


                        ' Sort udtData by Key, then step through the list and process each block, writing to disk as we go
                        Dim sortedData = (From item In lstData Order By item.Key, item.ValueDbl Select item).ToList()

                        intIndexBlockStart = 0
                        percentCompleteAtStart = nextPercentComplete
                        nextPercentComplete = 100

                        Dim sortedDataCount = sortedData.Count
                        For i = 1 To sortedDataCount - 1
                            If sortedData(i).Key <> sortedData(i - 1).Key Then
                                ' Parse this block
                                ' Copy data from sortedData to dataForBlock

                                Dim dataForBlock = New List(Of clsFileData)

                                For sourceIndex = intIndexBlockStart To i - 1
                                    dataForBlock.Add(sortedData(sourceIndex))
                                Next

                                ParseBlock(objOutlierFilter, dataForBlock, columnCount, swOutFile)

                                intIndexBlockStart = i
                            End If

                            UpdatePercentComplete(i, sortedDataCount, percentCompleteAtStart, nextPercentComplete)
                            OnProgressUpdate("Writing results", mPercentComplete)
                        Next

                        ' Copy the final elements into finalDataBlock, so that the last block will be written
                        For sourceIndex = intIndexBlockStart To lstData.Count - 1
                            finalDataBlock.Add(lstData(sourceIndex))
                        Next

                    End If

                End If

                If finalDataBlock.Count > 0 And Not mAbortProcessing Then
                    ' Parse the last block, then close the output file
                    ParseBlock(objOutlierFilter, finalDataBlock, columnCount, swOutFile)
                End If

            End Using

        Catch ex As Exception
            SetLocalErrorCode(eListPORErrorCodeCodes.ErrorWritingOutputFile)
            OnErrorEvent(GetErrorMessage())
            Return mLocalErrorCode
        End Try

        Return mLocalErrorCode
    End Function

    Private Sub SetLocalErrorCode(eNewErrorCode As eListPORErrorCodeCodes)
        SetLocalErrorCode(eNewErrorCode, False)
    End Sub

    Private Sub SetLocalErrorCode(eNewErrorCode As eListPORErrorCodeCodes, blnLeaveExistingErrorCodeUnchanged As Boolean)

        If blnLeaveExistingErrorCodeUnchanged AndAlso mLocalErrorCode <> eListPORErrorCodeCodes.NoError Then
            ' An error code is already defined; do not change it
        Else
            If eNewErrorCode = eListPORErrorCodeCodes.NoError Then
                mLocalErrorCode = eListPORErrorCodeCodes.NoError
            Else
                mLocalErrorCode = mLocalErrorCode Or eNewErrorCode
            End If

        End If

    End Sub

    ''' <summary>
    ''' Converts shifted symmetric value to normal value
    ''' </summary>
    ''' <param name="dblSymmetricValue"></param>
    ''' <returns></returns>
    Private Function SymmetricValueToNormalValue(dblSymmetricValue As Double) As Double

        Try
            If dblSymmetricValue >= 0 Then
                Return dblSymmetricValue + 1
            Else
                Return 1 / (1 - dblSymmetricValue)
            End If
        Catch
            Return 0
        End Try

    End Function

    Private Sub UpdatePercentComplete(stepsComplete As Long, totalSteps As Long, percentCompleteAtStart As Integer, nextPercentComplete As Integer)

        Try
            If totalSteps > 0 Then
                Dim subtaskPercentComplete = CType(stepsComplete / (totalSteps / 100.0), Single)

                Dim deltaPercentComplete = nextPercentComplete - percentCompleteAtStart

                If deltaPercentComplete > 0 Then
                    mPercentComplete = percentCompleteAtStart + subtaskPercentComplete * deltaPercentComplete / 100
                Else
                    mPercentComplete = subtaskPercentComplete
                End If

                If mPercentComplete > 100 Then mPercentComplete = 100
            Else
                mPercentComplete = 0
            End If

        Catch ex As Exception
            OnErrorEvent("Error in UpdatePercentComplete")
        End Try

    End Sub

    Private Sub WriteBlockLine(swOutFile As TextWriter, columnCount As Integer, dataPoint As clsFileData, blnAppendGroupAverage As Boolean, dblAverage As Double)

        Dim strOutputLine As String

        If columnCount = 2 Then
            strOutputLine = dataPoint.Key & ControlChars.Tab & dataPoint.Value
        Else
            strOutputLine = dataPoint.Value
        End If

        If Not String.IsNullOrWhiteSpace(dataPoint.RemainingCols) Then
            strOutputLine &= ControlChars.Tab & dataPoint.RemainingCols
        End If

        If blnAppendGroupAverage Then
            strOutputLine &= ControlChars.Tab & PRISM.StringUtilities.ValueToString(dblAverage, 5)
        End If

        swOutFile.WriteLine(strOutputLine)

    End Sub

End Class
