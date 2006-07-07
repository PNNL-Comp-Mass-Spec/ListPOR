Option Strict On

' This class can be used to remove outliers from a list of numbers (doubles)
' It uses Grubb's test to determine whether or not each number in the list
'  is far enough away from the mean to be thrown out
'
' Utilizes classes QSDouble and StatDoubles
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

Public Class clsGrubbsTestOutlierFilter

#Region "Module-wide Variables"
    Public Enum eclConfidenceLevelConstants
        e95Pct = 0
        e97Pct = 1
        e99Pct = 2
    End Enum

    Private mConfidenceLevel As eclConfidenceLevelConstants
    Private mMinFinalValueCount As Integer
    Private mIterate As Boolean
#End Region

#Region "Interface functions"

    Public Property ConfidenceLevel() As eclConfidenceLevelConstants
        Get
            Return mConfidenceLevel
        End Get
        Set(ByVal Value As eclConfidenceLevelConstants)
            mConfidenceLevel = Value
        End Set
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

    Public Property RemoveMultipleValues() As Boolean
        Get
            Return mIterate
        End Get
        Set(ByVal Value As Boolean)
            mIterate = Value
        End Set
    End Property
#End Region

    Public Function RemoveOutliers(ByRef dblValues() As Double, ByRef intIndexPointers() As Integer, ByRef intValueCountRemovedOut As Integer) As Boolean
        ' Removes outliers from dblValues() using Grubb's test and the given confidence level
        ' intIndexPointers() is an array of integers that is parallel to dblValues(), and will be
        '  kept in sync with any changes made to dblValues

        ' If intMaxIterations > 1, then will repeatedly remove the outliers, until no outliers
        '  remain or the number of values falls below intMinValues
        '
        ' Returns True if success (even if no values removed) and false if an error or dblValuesSorted doesn't contain any data
        ' Returns the number of values removed in intValueCountRemovedOut

        Dim dblValuesSorted() As Double         ' Sorted array of doubles            ' 0-based array
        Dim intIndexPointersSorted() As Integer

        Dim blnSuccess As Boolean
        Dim blnValuesRemoved As Boolean

        If dblValues.Length <= mMinFinalValueCount Then
            ' Cannot remove outliers since not enough members
            intValueCountRemovedOut = 0
            Return True
        End If

        Try
            dblValuesSorted = dblValues
            intIndexPointersSorted = intIndexPointers

            Array.Sort(dblValuesSorted, intIndexPointersSorted)

            If dblValuesSorted.Length > 0 Then
                blnSuccess = True
            Else
                blnSuccess = False
            End If

        Catch ex As Exception
            blnSuccess = False
        End Try

        If dblValuesSorted.Length <= 0 Or Not blnSuccess Then
            blnSuccess = False
        Else
            Try
                ' Copy the data from dblValuesSorted back to dblValues
                dblValues = dblValuesSorted
                intIndexPointers = intIndexPointersSorted

                intValueCountRemovedOut = 0
                Do
                    blnValuesRemoved = RemoveOutliersWork(dblValues, intIndexPointers, mConfidenceLevel)

                    If blnValuesRemoved Then
                        intValueCountRemovedOut += 1
                    End If

                Loop While blnValuesRemoved And mIterate And dblValues.Length > mMinFinalValueCount
                blnSuccess = True

            Catch ex As Exception
                blnSuccess = False
            End Try
        End If

        RemoveOutliers = blnSuccess

    End Function

    Private Function RemoveOutliersWork(ByRef dblValues() As Double, ByRef intIndexPointers() As Integer, ByVal eclConfidenceLevel As eclConfidenceLevelConstants) As Boolean
        ' Removes, at most, one outlier from dblValues (and from the corresponding position in intIndexPointers)
        ' Returns True if an outlier is removed, and false if not
        ' Returns false if an error occurs
        '
        ' NOTE: This function assumes that dblValues() is sorted ascending, and thus only tests
        '       the first and last value in the list as an outlier (removing only one of them, if appropriate)

        Dim objStatDoubles As New clsStatDoubles
        Dim blnValueRemoved As Boolean

        Dim intCount As Integer

        Dim intIndex As Integer, intIndex2 As Integer
        Dim intTargetIndex As Integer
        Dim dblTargetDistance As Double
        Dim dblCompareDistance As Double

        Dim dblMean As Double
        Dim dblStDev As Double
        Dim dblZScore As Double
        Dim dblPValue As Double

        Try
            intCount = dblValues.Length
            If intCount < 3 Then
                ' Cannot remove an outlier from fewer than 3 values
                RemoveOutliersWork = False
                Exit Function
            End If

        Catch ex As Exception
            RemoveOutliersWork = False
            Exit Function
        End Try

        Try
            blnValueRemoved = False
            If objStatDoubles.Fill(dblValues) Then
                dblMean = objStatDoubles.Mean
                dblStDev = objStatDoubles.StDev

                If dblStDev > 0 Then
                    ' Find the value furthest away from the mean
                    ' Since dblValues() is sorted, it can only be the first or last value

                    intTargetIndex = 0
                    dblTargetDistance = Math.Abs(dblValues(0) - dblMean)

                    dblCompareDistance = Math.Abs(dblValues(intCount - 1) - dblMean)
                    If dblCompareDistance > dblTargetDistance Then
                        dblTargetDistance = dblCompareDistance
                        intTargetIndex = intCount - 1
                    End If

                    ' Compute the z-score for intTargetIndex
                    dblZScore = dblTargetDistance / dblStDev

                    If intCount = 3 Then
                        ' When there are 3 values in the list, the p-value is always 1.15, regardless of the confidence level
                        dblPValue = 1.15
                    Else
                        ' Compute the p-value, based on eclConfidenceLevel
                        Select Case eclConfidenceLevel
                            Case eclConfidenceLevelConstants.e95Pct
                                ' Estimate the P value at the 95%'ile using a formula provided by
                                '  Robin Edwards <robin.edwards@argonet.co.uk>
                                dblPValue = (3.6996 * intCount + 145.9 - 186.7 / intCount) / (intCount + 59.5 + 58.5 / intCount)
                            Case eclConfidenceLevelConstants.e97Pct
                                dblPValue = Lookup97PctPValue(intCount)
                            Case Else
                                ' Includes eclConfidenceLevelConstants.e99pct
                                ' Estimate the P value at the 99%'ile using a formula provided by
                                '  Robin Edwards <robin.edwards@argonet.co.uk>
                                dblPValue = (4.1068 * intCount + 273.6 - 328.5 / intCount) / (intCount + 88.7 + 185 / intCount)
                        End Select
                    End If

                    If dblZScore > dblPValue Then
                        ' Remove the value
                        ' Copy the data in place, skipping the outlier value

                        intIndex2 = 0
                        For intIndex = 0 To intCount - 1
                            If intIndex <> intTargetIndex Then
                                dblValues(intIndex2) = dblValues(intIndex)
                                intIndexPointers(intIndex2) = intIndexPointers(intIndex)
                                intIndex2 += 1
                            End If
                        Next intIndex

                        ReDim Preserve dblValues(intCount - 2)
                        ReDim Preserve intIndexPointers(intCount - 2)
                        blnValueRemoved = True
                    End If
                End If
            End If

        Catch ex As Exception
            blnValueRemoved = False
        End Try

        RemoveOutliersWork = blnValueRemoved

    End Function

    Private Function Lookup97PctPValue(ByVal intCount As Integer) As Double

        If intCount <= 3 Then
            Return 1.15
        ElseIf intCount = 4 Then
            Return 1.48
        ElseIf intCount = 5 Then
            Return 1.71
        ElseIf intCount = 6 Then
            Return 1.89
        ElseIf intCount = 7 Then
            Return 2.02
        ElseIf intCount = 8 Then
            Return 2.13
        ElseIf intCount = 9 Then
            Return 2.21
        ElseIf intCount = 10 Then
            Return 2.29
        ElseIf intCount = 11 Then
            Return 2.34
        ElseIf intCount = 12 Then
            Return 2.41
        ElseIf intCount = 13 Then
            Return 2.46
        ElseIf intCount = 14 Then
            Return 2.51
        ElseIf intCount = 15 Then
            Return 2.55
        ElseIf intCount = 16 Then
            Return 2.59
        ElseIf intCount = 17 Then
            Return 2.62
        ElseIf intCount = 18 Then
            Return 2.65
        ElseIf intCount = 19 Then
            Return 2.68
        ElseIf intCount = 20 Then
            Return 2.71
        ElseIf intCount = 21 Then
            Return 2.73
        ElseIf intCount = 22 Then
            Return 2.76
        ElseIf intCount = 23 Then
            Return 2.78
        ElseIf intCount = 24 Then
            Return 2.8
        ElseIf intCount = 25 Then
            Return 2.82
        ElseIf intCount = 26 Then
            Return 2.84
        ElseIf intCount = 27 Then
            Return 2.86
        ElseIf intCount = 28 Then
            Return 2.88
        ElseIf intCount = 29 Then
            Return 2.89
        ElseIf intCount = 30 Then
            Return 2.91
        ElseIf intCount = 31 Then
            Return 2.92
        ElseIf intCount = 32 Then
            Return 2.94
        ElseIf intCount = 33 Then
            Return 2.95
        ElseIf intCount = 34 Then
            Return 2.97
        ElseIf intCount = 35 Then
            Return 2.98
        ElseIf intCount = 36 Then
            Return 2.99
        ElseIf intCount = 37 Then
            Return 3
        ElseIf intCount = 38 Then
            Return 3.01
        ElseIf intCount = 39 Then
            Return 3.03
        ElseIf intCount = 40 Then
            Return 3.04
        ElseIf intCount <= 50 Then
            Return 3.13
        ElseIf intCount <= 60 Then
            Return 3.2
        ElseIf intCount <= 70 Then
            Return 3.26
        ElseIf intCount <= 80 Then
            Return 3.31
        ElseIf intCount <= 90 Then
            Return 3.35
        ElseIf intCount <= 100 Then
            Return 3.38
        ElseIf intCount <= 110 Then
            Return 3.42
        ElseIf intCount <= 120 Then
            Return 3.44
        ElseIf intCount <= 130 Then
            Return 3.47
        ElseIf intCount <= 140 Then
            Return 3.49
        Else
            Return 3.5
        End If

    End Function

    Public Sub New()
        mConfidenceLevel = eclConfidenceLevelConstants.e95Pct
        mMinFinalValueCount = 3
        mIterate = False
    End Sub
End Class
