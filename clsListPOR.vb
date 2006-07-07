Option Strict On

' ListPOR stands for List Parser for Outlier Removal
' 
' Parses an input file looking for blocks of related data
' For each block, filters the data using clsGrubbsTestOutlierFilter
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

Public Class clsListPOR

#Region "Public Structures and Constants"
    Public Structure udtFileDataType
        Public Key As String
        Public Value As String
        Public RemainingCols As String
    End Structure

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
    Private mRemoveMultipleValues As Boolean
    Private mUseSymmetricValues As Boolean
    Private mUseNaturalLogValues As Boolean

    Private mAppendGroupAverage As Boolean
    Private mColumnCountOverride As Integer
    Private mLocalErrorCode As eListPORErrorCodeCodes

    Private mFileLengthBytes As Long
    Private mFileBytesRead As Long
    Private mAbortProcessing As Boolean
#End Region

#Region "Interface functions"

    Public Property AppendGroupAverage() As Boolean
        Get
            Return mAppendGroupAverage
        End Get
        Set(ByVal Value As Boolean)
            mAppendGroupAverage = Value
        End Set
    End Property

    Public Property AssumeSortedInputFile() As Boolean
        Get
            Return mAssumeSortedInputFile
        End Get
        Set(ByVal Value As Boolean)
            mAssumeSortedInputFile = Value
        End Set
    End Property

    Public Property ConfidenceLevel() As clsGrubbsTestOutlierFilter.eclConfidenceLevelConstants
        Get
            Return mConfidenceLevel
        End Get
        Set(ByVal Value As clsGrubbsTestOutlierFilter.eclConfidenceLevelConstants)
            mConfidenceLevel = Value
        End Set
    End Property

    Public Property ColumnCountOverride() As Integer
        Get
            Return mColumnCountOverride
        End Get
        Set(ByVal Value As Integer)
            ' Value can only be 1 or 2
            If Value = 1 Or Value = 2 Then
                mColumnCountOverride = Value
            Else
                mColumnCountOverride = 0
            End If
        End Set
    End Property

    Public ReadOnly Property ErrorCode() As eListPORErrorCodeCodes
        Get
            Return mLocalErrorCode
        End Get
    End Property

    Public Property MinFinalValueCount() As Integer
        Get
            Return mMinFinalValueCount
        End Get
        Set(ByVal Value As Integer)
            If Value < 2 Then Value = 2
            mMinFinalValueCount = Value
        End Set
    End Property

    Public ReadOnly Property PercentComplete() As Single
        Get
            Return ComputePercentComplete()
        End Get
    End Property

    Public Property RemoveMultipleValues() As Boolean
        Get
            Return mRemoveMultipleValues
        End Get
        Set(ByVal Value As Boolean)
            mRemoveMultipleValues = Value
        End Set
    End Property

    Public Property UseNaturalLogValues() As Boolean
        Get
            Return mUseNaturalLogValues
        End Get
        Set(ByVal Value As Boolean)
            mUseNaturalLogValues = Value
        End Set
    End Property

    Public Property UseSymmetricValues() As Boolean
        Get
            Return mUseSymmetricValues
        End Get
        Set(ByVal Value As Boolean)
            mUseSymmetricValues = Value
        End Set
    End Property
#End Region

    Public Event ProgressUpdate()

    Public Sub AbortProcessing()
        mAbortProcessing = True
    End Sub

    Private Function ComputePercentComplete() As Single

        Dim sngPercentComplete As Single

        Try
            If mFileLengthBytes > 0 Then
                sngPercentComplete = CType(mFileBytesRead / CType(mFileLengthBytes, Single) * 100, Single)
                If sngPercentComplete > 100 Then sngPercentComplete = 100
                Return sngPercentComplete
            Else
                Return 0
            End If

        Catch ex As Exception
            Return 0
        End Try

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

    Private Function NaturalLogValueToNormalValue(ByVal dblNaturalLogValue As Double) As Double
        Try
            Return Math.Exp(dblNaturalLogValue)
        Catch
            Return 0
        End Try

    End Function

    Private Function NormalValueToNaturalLog(ByVal dblValue As Double) As Double
        '------------------------------------------------------
        'Converts ratio-based value to natural log
        '------------------------------------------------------
        Try
            Return Math.Log(dblValue)
        Catch
            Return 0
        End Try
    End Function

    Private Function NormalValueToSymmetricValue(ByVal dblValue As Double) As Double
        '------------------------------------------------------
        'Converts ratio-based value to shifted symmetric value
        '------------------------------------------------------
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


    Private Sub ParseBlock(ByRef objOutlierFilter As clsGrubbsTestOutlierFilter, ByRef udtData() As udtFileDataType, ByVal intMaxIndexToUse As Integer, ByVal intColCount As Integer, ByVal blnWriteToDisk As Boolean, ByRef srOutFile As System.IO.StreamWriter)

        Dim intIndexPointers() As Integer
        Dim dblValues() As Double

        Dim blnAppendGroupAverage As Boolean
        Dim dblSum As Double
        Dim dblAverage As Double

        Dim udtDataFiltered() As udtFileDataType

        Dim intIndex As Integer
        Dim intValueCountRemoved As Integer

        If intMaxIndexToUse < 0 Then
            Debug.Assert(False, "Invalid call to ParseBlock")
            Exit Sub
        End If

        Try
            ReDim intIndexPointers(intMaxIndexToUse)
            ReDim dblValues(intMaxIndexToUse)

            ' Copy the values to examine into dblValues
            For intIndex = 0 To intMaxIndexToUse
                Try
                    dblValues(intIndex) = CType(udtData(intIndex).Value, Double)
                    If mUseSymmetricValues Or mUseNaturalLogValues Then
                        If dblValues(intIndex) > 0 Then
                            If mUseNaturalLogValues Then
                                dblValues(intIndex) = NormalValueToNaturalLog(dblValues(intIndex))
                            Else
                                dblValues(intIndex) = NormalValueToSymmetricValue(dblValues(intIndex))
                            End If
                        Else
                            ' Invalid value for computing the symmetric value or logarithm
                            dblValues(intIndex) = 0
                        End If
                    End If
                Catch ex As Exception
                    dblValues(intIndex) = 0
                End Try
                intIndexPointers(intIndex) = intIndex
            Next intIndex

            If intMaxIndexToUse < mMinFinalValueCount Then
                ' Not enough data to remove outliers
            Else
                ' Enough data to remove outliers
                If objOutlierFilter.RemoveOutliers(dblValues, intIndexPointers, intValueCountRemoved) Then
                    ' Successful call to .RemoveOutliers

                    If blnWriteToDisk Then
                        ' We'll write the block to the output file below
                        ' No need to update udtData since we're going to discard it anyway
                    Else
                        ' Remove the outlier data from udtData
                        If intValueCountRemoved > 0 Then
                            ' Remove the outlier data from udtData

                            ReDim udtDataFiltered(intMaxIndexToUse)

                            For intIndex = 0 To intIndexPointers.Length - 1
                                udtDataFiltered(intIndex) = udtData(intIndexPointers(intIndex))
                            Next

                            ' Copy the data from udtDataFiltered back to udtData
                            udtData = udtDataFiltered
                        End If
                    End If
                Else
                    ' Error removing outliers; simply re-write the input text to the output file
                End If

                'If blnUseCenteredMedian Then

                'End If
            End If

            If blnWriteToDisk Then
                blnAppendGroupAverage = mAppendGroupAverage And intColCount > 1
                If blnAppendGroupAverage Then
                    ' Compute the average value for the values in dblValues and append as an additional column

                    dblSum = 0
                    For intIndex = 0 To dblValues.Length - 1
                        dblSum += dblValues(intIndex)
                    Next

                    If dblValues.Length > 0 Then
                        dblAverage = dblSum / dblValues.Length
                        If mUseSymmetricValues Or mUseNaturalLogValues Then
                            If mUseNaturalLogValues Then
                                dblAverage = NaturalLogValueToNormalValue(dblAverage)
                            Else
                                dblAverage = SymmetricValueToNormalValue(dblAverage)
                            End If
                        End If
                    Else
                        dblAverage = 0
                    End If
                End If

                For intIndex = 0 To intIndexPointers.Length - 1
                    WriteBlockLine(srOutFile, intColCount, udtData(intIndexPointers(intIndex)), blnAppendGroupAverage, dblAverage)
                Next intIndex
            End If

        Catch ex As Exception
            Throw New Exception("Error parsing block", ex)
        End Try

    End Sub

    Public Function RemoveOutliersFromListInMemory(ByRef strValueLabels() As String, ByRef dblValues() As Double, ByRef strAdditionalLabelsOrData() As String) As Boolean
        ' Examines the labels in strValueLabels to find groups of data
        ' For each group, examines the corresponding data in dblValues and removes outliers for the group
        ' If strValueLabels is blank, then simply examines dblValues for outliers
        ' Carries along strAdditionalLabelsOrData (if defined) and removes data from that array if removing data from dblValues

        ' ToDo: Code this someday
        SetLocalErrorCode(eListPORErrorCodeCodes.NoError)

        Return False
    End Function

    Public Function RemoveOutliersFromListInFile(ByVal strSourcePath As String, ByVal strDestPath As String) As Integer
        ' Reads a tab-delimeted text file (strSourcePath) and removes outliers from lists of numbers in column
        '
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
        ' For both modes, the output will be saved in strDestPath if defined, otherwise, 
        '  strSourcePath will be overwritten
        ' If blnAssumeFileIsSorted is False, then the entire file will be read into memory, 
        '  the data will be sorted on the first column, then the outliers will be found
        ' If blnAssumeFileIsSorted is True, then only the data for the most recent group will be retained
        '  This is useful for parsing files will too much data to reside in memory
        '
        ' Auth: mem
        ' Date: 08/13/2004


        Const ARRAY_DIM_SIZE As Integer = 10000

        Dim objOutlierFilter As clsGrubbsTestOutlierFilter

        Dim srInFile As System.IO.StreamReader
        Dim srOutFile As System.IO.StreamWriter

        Dim strColumnDelimeter As Char
        Dim strDelimList() As Char
        Dim strSplitLine() As String

        Dim strLineIn As String, strLineOut As String

        Dim intLineCount As Integer, intColCount As Integer
        Dim udtData() As udtFileDataType
        Dim udtDataForBlock() As udtFileDataType

        Dim udtDataNextLine As udtFileDataType
        Dim objComparer As System.Collections.IComparer

        Dim strHeaderLine As String

        Dim intIndex As Integer
        Dim intIndexBlockStart As Integer
        Dim intIndexCopy As Integer

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
            Console.WriteLine(GetErrorMessage)
            Return mLocalErrorCode
        End Try

        Try

            ' Assume the column delimeter is a tab, unless the input file ends in .csv
            If Right(strSourcePath, 4).ToLower = ".csv" Then
                strColumnDelimeter = ","c
            Else
                strColumnDelimeter = ControlChars.Tab
            End If
            strDelimList = New Char() {strColumnDelimeter}

            If strDestPath Is Nothing OrElse strDestPath.Length = 0 OrElse strDestPath = strSourcePath Then
                strDestPathToUse = strSourcePath & ".filtered"
                strDestPath = String.Copy(strSourcePath)
            Else
                strDestPathToUse = String.Copy(strDestPath)
            End If

            ' Open the input file
            srInFile = New System.IO.StreamReader(strSourcePath)

            mFileLengthBytes = srInFile.BaseStream.Length
            RaiseEvent ProgressUpdate()

            Try
                ' Open the output file
                srOutFile = New System.IO.StreamWriter(strDestPathToUse)
            Catch ex As Exception
                SetLocalErrorCode(eListPORErrorCodeCodes.ErrorWritingOutputFile)
                Console.WriteLine(GetErrorMessage)
                Return mLocalErrorCode
            End Try

            intLineCount = 0
            ReDim udtData(ARRAY_DIM_SIZE)

            If mColumnCountOverride = 1 Or mColumnCountOverride = 2 Then
                intColCount = mColumnCountOverride
            Else
                ' Set the column count to 0 for now, unless 
                ' We'll update it when we reach a header line or a data block line
                intColCount = 0
            End If

            strLineIn = String.Empty
            strHeaderLine = String.Empty
            Do While srInFile.Peek() >= 0

                strLineIn = srInFile.ReadLine()
                mFileBytesRead += strLineIn.Length + 2
                RaiseEvent ProgressUpdate()
                If mAbortProcessing Then Exit Do

                If Not strLineIn Is Nothing AndAlso strLineIn.Length > 0 Then
                    ' Split the line, allowing at most 3 columns (if more than 3, then all data is lumped into the third column)
                    strSplitLine = strLineIn.Split(strDelimList, 3)

                    If Not blnDataBlockReached Then
                        ' Haven't found the data block yet, is this a header line?
                        If intColCount = 1 Then
                            If strSplitLine.Length > 0 AndAlso IsNumeric(strSplitLine(0)) Then
                                blnDataBlockReached = True
                            Else
                                strHeaderLine = strLineIn
                            End If
                        ElseIf intColCount = 2 Then
                            If strSplitLine.Length > 1 AndAlso IsNumeric(strSplitLine(1)) Then
                                blnDataBlockReached = True
                            Else
                                strHeaderLine = strLineIn
                            End If
                        ElseIf strSplitLine.Length >= 2 Then
                            If Not IsNumeric(strSplitLine(0)) And IsNumeric(strSplitLine(1)) Then
                                blnDataBlockReached = True
                                intColCount = 2
                            Else
                                strHeaderLine = strLineIn
                            End If
                        ElseIf strSplitLine.Length = 1 Then
                            If IsNumeric(strSplitLine(0)) Then
                                blnDataBlockReached = True
                                intColCount = 1
                            Else
                                strHeaderLine = strLineIn
                            End If
                        Else
                            strHeaderLine = strLineIn
                        End If

                        If Not blnDataBlockReached Then
                            ' Write out the header lines as we read them
                            srOutFile.WriteLine(strHeaderLine)
                        End If
                    End If

                    If blnDataBlockReached Then
                        If intColCount = 0 Then
                            ' Determine the number of columns of data

                            If strSplitLine.Length >= 2 AndAlso IsNumeric(strSplitLine(1)) Then
                                intColCount = 2
                            ElseIf strSplitLine.Length = 1 And IsNumeric(strSplitLine(0)) Then
                                intColCount = 1
                            Else
                                ' Assume intColCount = 1, even though the input file probably isn't formatted correctly
                                Debug.Assert(False, "Input file does not have the expected number of columns of data")
                                intColCount = 1
                            End If
                        End If

                        If intColCount = 2 AndAlso strSplitLine.Length >= 2 AndAlso IsNumeric(strSplitLine(1)) Then
                            Try
                                With udtData(intLineCount)
                                    .Key = strSplitLine(0)
                                    .Value = strSplitLine(1)
                                    If strSplitLine.Length > 2 Then
                                        .RemainingCols = strSplitLine(2)
                                    Else
                                        .RemainingCols = String.Empty
                                    End If
                                End With
                                intLineCount += 1
                            Catch ex As Exception
                                ' Line parsing error; skip this line
                            End Try
                        ElseIf intColCount = 2 AndAlso strSplitLine.Length = 1 AndAlso IsNumeric(strSplitLine(0)) Then
                            Try
                                With udtData(intLineCount)
                                    .Key = strSplitLine(0)
                                    .Value = "0"
                                    .RemainingCols = String.Empty
                                End With
                                intLineCount += 1
                            Catch ex As Exception
                                ' Line parsing error; skip this line
                            End Try
                        ElseIf intColCount = 1 AndAlso strSplitLine.Length >= 1 AndAlso IsNumeric(strSplitLine(0)) Then
                            Try
                                With udtData(intLineCount)
                                    .Key = "A"
                                    .Value = strSplitLine(0)
                                    If strSplitLine.Length > 1 Then
                                        .RemainingCols = strSplitLine(1)
                                        If strSplitLine.Length > 2 Then .RemainingCols &= strSplitLine(2)
                                    Else
                                        .RemainingCols = String.Empty
                                    End If

                                End With
                                intLineCount += 1
                            Catch ex As Exception
                                ' Line parsing error; skip this line
                            End Try
                        Else
                            ' Ignore the line
                        End If

                        If intLineCount >= udtData.Length Then
                            ReDim Preserve udtData(udtData.Length + ARRAY_DIM_SIZE)
                        End If

                        If blnAssumeFileIsSorted Then
                            If intLineCount > 1 Then
                                If udtData(intLineCount - 1).Key <> udtData(intLineCount - 2).Key Then
                                    ' Parse this block
                                    ' Need to save the latest value in udtDataNextLine since it will be removed from udtData

                                    udtDataNextLine = udtData(intLineCount - 1)

                                    ParseBlock(objOutlierFilter, udtData, intLineCount - 2, intColCount, True, srOutFile)

                                    ' Reset udtData
                                    intLineCount = 1
                                    udtData(0) = udtDataNextLine
                                End If
                            End If
                        End If
                    End If

                End If
            Loop

        Catch ex As Exception
            SetLocalErrorCode(eListPORErrorCodeCodes.ErrorReadingInputFile)
            Console.WriteLine(GetErrorMessage)
            Return mLocalErrorCode
        Finally
            If Not srInFile Is Nothing Then
                srInFile.Close()
            End If
        End Try

        Try
            If mAbortProcessing Then
                SetLocalErrorCode(eListPORErrorCodeCodes.ProcessingAborted)
            ElseIf Not blnAssumeFileIsSorted Then
                ' Shrink udtData to the correct size

                If intLineCount > 0 Then
                    ReDim Preserve udtData(intLineCount - 1)

                    ' Sort udtData by Key, then step through the list and process each block, writing to disk as we go
                    Array.Sort(udtData, New ComparerFileDataLine)

                    intIndexBlockStart = 0
                    For intIndex = 1 To udtData.Length - 1
                        If udtData(intIndex).Key <> udtData(intIndex - 1).Key Then
                            ' Parse this block
                            ' Copy data from udtData to udtDataForBlock

                            ReDim udtDataForBlock(intIndex - intIndexBlockStart - 1)

                            For intIndexCopy = intIndexBlockStart To intIndex - 1
                                udtDataForBlock(intIndexCopy - intIndexBlockStart) = udtData(intIndexCopy)
                            Next

                            ParseBlock(objOutlierFilter, udtDataForBlock, udtDataForBlock.Length - 1, intColCount, True, srOutFile)

                            intIndexBlockStart = intIndex
                        End If
                    Next

                    ' Copy the final elements into the start of udtData, so that the last block will be written
                    For intIndexCopy = intIndexBlockStart To udtData.Length - 1
                        udtData(intIndexCopy - intIndexBlockStart) = udtData(intIndexCopy)
                    Next

                    ' Bump down intLineCount
                    intLineCount = udtData.Length - intIndexBlockStart
                End If

            End If

            If intLineCount > 0 And Not mAbortProcessing Then
                ' Parse the last block, then close the output file
                ParseBlock(objOutlierFilter, udtData, intLineCount - 1, intColCount, True, srOutFile)
            End If


        Catch ex As Exception
            SetLocalErrorCode(eListPORErrorCodeCodes.ErrorWritingOutputFile)
            Console.WriteLine(GetErrorMessage)
            Return mLocalErrorCode

        Finally
            If Not srOutFile Is Nothing Then
                srOutFile.Close()
            End If
        End Try

        Return mLocalErrorCode
    End Function

    Private Sub SetLocalErrorCode(ByVal eNewErrorCode As eListPORErrorCodeCodes)
        SetLocalErrorCode(eNewErrorCode, False)
    End Sub

    Private Sub SetLocalErrorCode(ByVal eNewErrorCode As eListPORErrorCodeCodes, ByVal blnLeaveExistingErrorCodeUnchanged As Boolean)

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

    Private Function SymmetricValueToNormalValue(ByVal dblSymmetricValue As Double) As Double
        '------------------------------------------------------
        'Converts shifted symmetric value to normal value
        '------------------------------------------------------

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

    Private Sub WriteBlockLine(ByRef srOutFile As System.IO.StreamWriter, ByVal intColCount As Integer, ByRef udtData As udtFileDataType, ByVal blnAppendGroupAverage As Boolean, ByVal dblaverage As Double)

        Dim strOutputLine As String

        With udtData

            If intColCount = 2 Then
                strOutputLine = .Key & ControlChars.Tab & .Value
            Else
                strOutputLine = .Value
            End If

            If Not .RemainingCols Is Nothing AndAlso .RemainingCols.Length > 0 Then
                strOutputLine &= ControlChars.Tab & .RemainingCols
            End If

            If blnAppendGroupAverage Then
                strOutputLine &= ControlChars.Tab & Math.Round(dblaverage, 8).ToString
            End If

            srOutFile.WriteLine(strOutputLine)
        End With

    End Sub

    Public Sub New()
        mAssumeSortedInputFile = False
        mRemoveMultipleValues = True
        mConfidenceLevel = clsGrubbsTestOutlierFilter.eclConfidenceLevelConstants.e95Pct
        mMinFinalValueCount = 3
        mColumnCountOverride = 0
        mUseSymmetricValues = False
        UseNaturalLogValues = False
        mAppendGroupAverage = True

        mLocalErrorCode = eListPORErrorCodeCodes.NoError
    End Sub
End Class

Public Class ComparerFileDataLine
    Implements System.Collections.IComparer

    Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare
        Dim xData As clsListPOR.udtFileDataType = CType(x, clsListPOR.udtFileDataType)
        Dim yData As clsListPOR.udtFileDataType = CType(y, clsListPOR.udtFileDataType)

        If (xData.Key > yData.Key) Then
            Return 1
        ElseIf xData.Key < yData.Key Then
            Return -1
        Else
            If (xData.Value > yData.Value) Then
                Return 1
            ElseIf xData.Value < yData.Value Then
                Return -1
            Else
                Return 0
            End If

        End If
    End Function

End Class
