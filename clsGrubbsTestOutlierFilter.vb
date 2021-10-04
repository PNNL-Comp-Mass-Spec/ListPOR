Option Strict On

Imports System.Collections.Generic
Imports System.Runtime.InteropServices

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

''' <summary>
'''  This class can be used to identify outliers in a list of numbers.
''' It uses Grubb's test to determine whether or not each number in the list
'''  is far enough away from the mean to be thrown out
''' </summary>
Public Class clsGrubbsTestOutlierFilter

#Region "Enums and Structs"
    Public Enum eclConfidenceLevelConstants
        e95Pct = 0
        e97Pct = 1
        e99Pct = 2
    End Enum

    Private Structure udtThresholdInfoType
        Public DataCount As Integer
        Public Threshold As Single
    End Structure

#End Region

#Region "Module-wide Variables"

    Private mConfidenceLevel As eclConfidenceLevelConstants
    Private mMinFinalValueCount As Integer
    Private mIterate As Boolean

    Private ReadOnly m97PctThresholds As List(Of udtThresholdInfoType)

#End Region

#Region "Interface functions"

    ''' <summary>
    ''' Confidence level
    ''' </summary>
    ''' <returns></returns>
    Public Property ConfidenceLevel As eclConfidenceLevelConstants
        Get
            Return mConfidenceLevel
        End Get
        Set
            mConfidenceLevel = Value
        End Set
    End Property

    ''' <summary>
    ''' Minimum number of values that must be kept in the list
    ''' </summary>
    ''' <returns></returns>
    Public Property MinFinalValueCount As Integer
        Get
            Return mMinFinalValueCount
        End Get
        Set
            If Value < 2 Then Value = 2
            mMinFinalValueCount = Value
        End Set
    End Property

    ''' <summary>
    ''' If true, find multiple outliers
    ''' If false, only find the most extreme outlier
    ''' </summary>
    ''' <returns></returns>
    Public Property RemoveMultipleValues As Boolean
        Get
            Return mIterate
        End Get
        Set
            mIterate = Value
        End Set
    End Property

#End Region

    ''' <summary>
    ''' Constructor
    ''' </summary>
    Public Sub New()
        mConfidenceLevel = eclConfidenceLevelConstants.e95Pct
        mMinFinalValueCount = 3
        mIterate = False

        m97PctThresholds = New List(Of udtThresholdInfoType)
    End Sub

    ''' <summary>
    ''' Use Grubb's test to identify outliers in lstData (at a given confidence level)
    ''' </summary>
    ''' <param name="lstData"></param>
    ''' <param name="outlierIndices">Output: indices of outlier values</param>
    ''' <returns>True if success (even if no values removed) and false if an error or sortedValues doesn't contain any data</returns>
    ''' <remarks>
    ''' If RemoveMultipleValues is true, will repeatedly remove the outliers, until no outliers remain
    ''' or the number of values falls below MinFinalValueCount
    ''' </remarks>
    Public Function FindOutliers(lstData As List(Of Double), <Out> ByRef outlierIndices As SortedSet(Of Integer)) As Boolean

        outlierIndices = New SortedSet(Of Integer)

        If lstData.Count <= mMinFinalValueCount Then
            ' Cannot remove outliers since not enough members
            Return True
        End If

        Try

            Dim sortedValues = lstData.ToArray()
            Dim indexPointers(sortedValues.Length - 1) As Integer

            For i = 0 To sortedValues.Length - 1
                indexPointers(i) = i
            Next

            Array.Sort(sortedValues, indexPointers)

            Dim sortedDataCount = sortedValues.Count

            Do
                Dim candidateOutlierIndex As Integer
                Dim outlierFound = FindOutlierWork(sortedValues, sortedDataCount, mConfidenceLevel, candidateOutlierIndex)

                If Not outlierFound Then
                    Exit Do
                End If

                outlierIndices.Add(indexPointers(candidateOutlierIndex))

                ' Remove the outlier by copying data in place
                Dim targetIndex = 0
                For i = 0 To sortedDataCount - 1
                    If i = candidateOutlierIndex Then Continue For

                    sortedValues(targetIndex) = sortedValues(i)
                    indexPointers(targetIndex) = indexPointers(i)
                    targetIndex += 1
                Next

                sortedDataCount -= 1

            Loop While mIterate And sortedDataCount > mMinFinalValueCount

            Return True

        Catch ex As Exception
            outlierIndices = New SortedSet(Of Integer)
            Return False
        End Try

    End Function

    Private Function FindOutlierWork(
      sortedValues As IList(Of Double),
      dataCount As Integer,
      eclConfidenceLevel As eclConfidenceLevelConstants,
      <Out> ByRef candidateOutlierIndex As Integer) As Boolean

        ' Removes, at most, one outlier from dblValues (and from the corresponding position in intIndexPointers)
        ' Returns True if an outlier is removed, and false if not
        ' Returns false if an error occurs
        '
        ' NOTE: This function assumes that sortedValues() is sorted ascending, and thus only tests
        '       the first and last value in the list as an outlier (removing only one of them, if appropriate)
        ' NOTE 2: This function does not use sortedValues.Count; it instead trusts that dataCount lists the number of items in the array

        candidateOutlierIndex = -1

        If dataCount < 3 Then
            ' Cannot remove an outlier from fewer than 3 values
            Return False
        End If

        Try
            Dim dblMean As Double = MathNet.Numerics.Statistics.ArrayStatistics.Mean(sortedValues.Take(dataCount).ToArray())
            Dim dblStDev As Double = MathNet.Numerics.Statistics.ArrayStatistics.StandardDeviation(sortedValues.Take(dataCount).ToArray())

            If dblStDev <= 0 Then
                Return False
            End If

            candidateOutlierIndex = 0

            ' Find the value furthest away from the mean
            ' Since dblValues() is sorted, it can only be the first or last value

            Dim dblTargetDistance As Double = Math.Abs(sortedValues(0) - dblMean)

            Dim dblCompareDistance As Double = Math.Abs(sortedValues(dataCount - 1) - dblMean)
            If dblCompareDistance > dblTargetDistance Then
                dblTargetDistance = dblCompareDistance
                candidateOutlierIndex = dataCount - 1
            End If

            ' Compute the z-score for candidateOutlierIndex
            Dim dblZScore As Double = dblTargetDistance / dblStDev
            Dim dblPValue As Double

            If dataCount = 3 Then
                ' When there are 3 values in the list, the p-value is always 1.15, regardless of the confidence level
                dblPValue = 1.15
            Else
                ' Compute the p-value, based on eclConfidenceLevel
                Select Case eclConfidenceLevel
                    Case eclConfidenceLevelConstants.e95Pct
                        ' Estimate the P value at the 95%'ile using a formula provided by
                        '  Robin Edwards <robin.edwards@argonet.co.uk>
                        dblPValue = (3.6996 * dataCount + 145.9 - 186.7 / dataCount) /
                                    (dataCount + 59.5 + 58.5 / dataCount)
                    Case eclConfidenceLevelConstants.e97Pct
                        dblPValue = Lookup97PctPValue(dataCount)
                    Case Else
                        ' Includes eclConfidenceLevelConstants.e99pct
                        ' Estimate the P value at the 99%'ile using a formula provided by
                        '  Robin Edwards <robin.edwards@argonet.co.uk>
                        dblPValue = (4.1068 * dataCount + 273.6 - 328.5 / dataCount) /
                                    (dataCount + 88.7 + 185 / dataCount)
                End Select
            End If

            If dblZScore > dblPValue Then
                ' Outlier found
                Return True
            End If

            ' The value furthest from the mean is not an outlier
            candidateOutlierIndex = -1
            Return False

        Catch ex As Exception
            candidateOutlierIndex = -1
            Return False
        End Try

    End Function

    Private Function Lookup97PctPValue(dataCount As Integer) As Double

        If m97PctThresholds.Count = 0 Then
            AddThreshold(m97PctThresholds, 3, 1.15)
            AddThreshold(m97PctThresholds, 4, 1.48)
            AddThreshold(m97PctThresholds, 5, 1.71)
            AddThreshold(m97PctThresholds, 6, 1.89)
            AddThreshold(m97PctThresholds, 7, 2.02)
            AddThreshold(m97PctThresholds, 8, 2.13)
            AddThreshold(m97PctThresholds, 9, 2.21)
            AddThreshold(m97PctThresholds, 10, 2.29)
            AddThreshold(m97PctThresholds, 11, 2.34)
            AddThreshold(m97PctThresholds, 12, 2.41)
            AddThreshold(m97PctThresholds, 13, 2.46)
            AddThreshold(m97PctThresholds, 14, 2.51)
            AddThreshold(m97PctThresholds, 15, 2.55)
            AddThreshold(m97PctThresholds, 16, 2.59)
            AddThreshold(m97PctThresholds, 17, 2.62)
            AddThreshold(m97PctThresholds, 18, 2.65)
            AddThreshold(m97PctThresholds, 19, 2.68)
            AddThreshold(m97PctThresholds, 20, 2.71)
            AddThreshold(m97PctThresholds, 21, 2.73)
            AddThreshold(m97PctThresholds, 22, 2.76)
            AddThreshold(m97PctThresholds, 23, 2.78)
            AddThreshold(m97PctThresholds, 24, 2.8)
            AddThreshold(m97PctThresholds, 25, 2.82)
            AddThreshold(m97PctThresholds, 26, 2.84)
            AddThreshold(m97PctThresholds, 27, 2.86)
            AddThreshold(m97PctThresholds, 28, 2.88)
            AddThreshold(m97PctThresholds, 29, 2.89)
            AddThreshold(m97PctThresholds, 30, 2.91)
            AddThreshold(m97PctThresholds, 31, 2.92)
            AddThreshold(m97PctThresholds, 32, 2.94)
            AddThreshold(m97PctThresholds, 33, 2.95)
            AddThreshold(m97PctThresholds, 34, 2.97)
            AddThreshold(m97PctThresholds, 35, 2.98)
            AddThreshold(m97PctThresholds, 36, 2.99)
            AddThreshold(m97PctThresholds, 37, 3)
            AddThreshold(m97PctThresholds, 38, 3.01)
            AddThreshold(m97PctThresholds, 39, 3.03)
            AddThreshold(m97PctThresholds, 40, 3.04)
            AddThreshold(m97PctThresholds, 50, 3.13)
            AddThreshold(m97PctThresholds, 60, 3.2)
            AddThreshold(m97PctThresholds, 70, 3.26)
            AddThreshold(m97PctThresholds, 80, 3.31)
            AddThreshold(m97PctThresholds, 90, 3.35)
            AddThreshold(m97PctThresholds, 100, 3.38)
            AddThreshold(m97PctThresholds, 110, 3.42)
            AddThreshold(m97PctThresholds, 120, 3.44)
            AddThreshold(m97PctThresholds, 130, 3.47)
            AddThreshold(m97PctThresholds, 140, 3.49)
            AddThreshold(m97PctThresholds, 150, 3.51)
            AddThreshold(m97PctThresholds, 170, 3.55)
            AddThreshold(m97PctThresholds, 225, 3.62)
            AddThreshold(m97PctThresholds, 300, 3.68)
            AddThreshold(m97PctThresholds, 500, 3.75)
            AddThreshold(m97PctThresholds, 720, 3.8)
            AddThreshold(m97PctThresholds, 795, 3.81)
            AddThreshold(m97PctThresholds, 960, 3.82)
            AddThreshold(m97PctThresholds, 1050, 3.83)
            AddThreshold(m97PctThresholds, 1280, 3.84)
            AddThreshold(m97PctThresholds, 1550, 3.85)
            AddThreshold(m97PctThresholds, 2070, 3.86)
            AddThreshold(m97PctThresholds, 2750, 3.87)
            AddThreshold(m97PctThresholds, 4400, 3.88)
            AddThreshold(m97PctThresholds, 9500, 3.89)
            AddThreshold(m97PctThresholds, Integer.MaxValue, 3.9)
        End If

        Dim threshold As Single = 1.15

        For Each item In m97PctThresholds
            If dataCount > item.DataCount Then
                threshold = item.Threshold
            Else
                Exit For
            End If
        Next

        Return threshold

    End Function

    Private Sub AddThreshold(thresholds As ICollection(Of udtThresholdInfoType), dataCount As Integer, threshold As Single)

        Dim thresholdInfo = New udtThresholdInfoType With {
            .DataCount = dataCount,
            .Threshold = threshold
        }

        thresholds.Add(thresholdInfo)
    End Sub
End Class
