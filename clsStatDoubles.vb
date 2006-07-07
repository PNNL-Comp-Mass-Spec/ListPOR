Option Strict On

' -------------------------------------------------------------------------------
' Written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA)
' E-mail: matthew.monroe@pnl.gov or matt@alchemistmatt.com
' Website: http://ncrr.pnl.gov/ or http://www.sysbio.org/resources/staff/
' -------------------------------------------------------------------------------
' 
' Licensed under the Apache License, Version 2.0; you may not use this file except
' in compliance with the License.  You may obtain a copy of the License at 
' http://www.apache.org/licenses/LICENSE-2.0

Public Class clsStatDoubles
    ' ---------------------------------------------------------------------
    ' This function is used to calculate various statistics on an array
    ' of doubles; this is an overkill for some functions that do not need sorting
    ' Functions: Mean, Median, Quartile, Percentile, St.Dev., Minimum, Maximum
    ' ---------------------------------------------------------------------
    ' created: 11/18/2002 nt
    ' Ported to VB.NET 08/13/2004 by mem
    ' Last updated 7/30/2005 by mem
    '---------------------------------------------------------------------

    Public Enum sdQuarts
        sdQuart0 = 0            'minimum element
        sdQuart1 = 1            '25 percentile
        sdQuart2 = 2            'median
        sdQuart3 = 3            '75 percentile
        sdQuart4 = 4            'maximum element
    End Enum

    Private mCnt As Integer                 'count of array members
    Private mQSDbl() As Double              'sorted array of doubles            ' 0-based array

    Public ReadOnly Property Count() As Integer
        Get
            Return mCnt
        End Get
    End Property

    Public Function Median() As Double
        '-----------------------------------------------------------------------
        ' NOTE:if the number of members is even, then median is the average
        ' of the two members in the middle
        '-----------------------------------------------------------------------
        Dim HalfInd As Integer

        If mCnt > 0 Then
            HalfInd = CType(Math.Floor(mCnt / 2.0), Integer)
            If mCnt Mod 2 > 0 Then               'odd membership
                Return mQSDbl(HalfInd)
            Else                                 'even membership
                Return (mQSDbl(HalfInd - 1) + mQSDbl(HalfInd)) / 2
            End If
        Else
            Return 0
        End If
    End Function

    Public Function Mean() As Double
        If mCnt > 0 Then
            Return Sum() / mCnt
        Else
            Return 0
        End If
    End Function

    Public Function Maximum() As Double
        Dim i As Integer
        Dim dblMaximum As Double

        If mCnt > 0 Then
            dblMaximum = mQSDbl(0)
            For i = 1 To mCnt - 1
                If mQSDbl(i) > dblMaximum Then
                    dblMaximum = mQSDbl(i)
                End If
            Next i
            Return dblMaximum
        Else
            Return 0
        End If
    End Function

    Public Function Minimum() As Double
        Dim i As Integer
        Dim dblMinimum As Double

        If mCnt > 0 Then
            dblMinimum = mQSDbl(0)
            For i = 1 To mCnt - 1
                If mQSDbl(i) < dblMinimum Then
                    dblMinimum = mQSDbl(i)
                End If
            Next i
            Return dblMinimum
        Else
            Return 0
        End If
    End Function

    'Public Function Mode() As Double
    '    If mCnt > 0 Then
    '        ' For lists, the mode is the most common (frequent) value. A list can have more than one mode.
    '        ' This has not been coded
    '        Debug.Assert(False)
    '    Else
    '        Throw New System.Exception("clsStatDoubles.Mode: No data in memory; call Fill function first")
    '    End If
    'End Function

    Public Function Sum() As Double
        Dim i As Integer
        Dim dblSum As Double

        If mCnt > 0 Then
            dblSum = 0
            For i = 0 To mCnt - 1
                dblSum += mQSDbl(i)
            Next i
            Return dblSum
        Else
            Return 0
        End If
    End Function

    Public Function SumSquared() As Double
        Dim i As Integer
        Dim dblSumSq As Double

        If mCnt > 0 Then
            For i = 0 To mCnt - 1
                dblSumSq += mQSDbl(i) ^ 2
            Next i
            Return dblSumSq
        Else
            Return 0
        End If
    End Function

    Public Function Fill(ByVal DblArr() As Double) As Boolean
        '-------------------------------------------------------------------------
        'returns True if array is successfully sorted and has at least one element
        '-------------------------------------------------------------------------

        Try
            mCnt = -1

            ReDim mQSDbl(DblArr.Length - 1)
            Array.Copy(DblArr, mQSDbl, DblArr.Length)

            Array.Sort(mQSDbl)

            mCnt = mQSDbl.Length
        Catch ex As Exception
            mCnt = -1
        End Try

        Return (mCnt > 0)

    End Function

    Public Function Quartile(ByVal Quart As sdQuarts) As Double
        Select Case Quart
            Case sdQuarts.sdQuart0
                If mCnt > 0 Then
                    Return mQSDbl(0)
                Else
                    Return 0
                End If
            Case sdQuarts.sdQuart1
                Return Percentile(0.25)
            Case sdQuarts.sdQuart2
                Return Median()
            Case sdQuarts.sdQuart3
                Return Percentile(0.75)
            Case sdQuarts.sdQuart4
                If mCnt > 0 Then
                    Return mQSDbl(mCnt - 1)
                Else
                    Return 0
                End If
        End Select
    End Function

    Public Function Percentile(ByVal Pct As Double) As Double
        '-----------------------------------------------------------------------
        ' NOTE: we can probably do better interpolation but for practical purposes
        ' this is good enough
        '-----------------------------------------------------------------------
        Dim PctInd As Integer

        If Pct > 0 And Pct < 1 Then
            If mCnt > 0 Then
                PctInd = CType(Math.Round(mCnt * Pct, 0), Integer)
                If PctInd < mCnt * Pct Then
                    If PctInd < mCnt - 1 Then
                        Return (mQSDbl(PctInd) + mQSDbl(PctInd + 1)) / 2
                    Else
                        Return mQSDbl(PctInd)
                    End If
                Else
                    Return mQSDbl(PctInd)
                End If
            Else
                Return 0
            End If
        Else
            Throw New System.Exception("clsStatDoubles.Percentile: Invalid percentile value; should be greater than 0 and less than 1")
        End If
    End Function

    Public Function StDev() As Double
        '-----------------------------------------------------------------------
        ' returns standard deviation (nonbiased) of array of numbers (doubles);
        ' -1 if not applicable or any error
        '-----------------------------------------------------------------------

        Const USE_RUNNING_SUMS As Boolean = False

        Dim i As Integer

        Dim dblMean As Double
        Dim dblSum As Double
        Dim dblSumSq As Double
        Dim dblValue As Double

        Try
            If mCnt > 1 Then
                If Not USE_RUNNING_SUMS Then
                    ' The following method is the most robust for computing standard deviation
                    ' This is the method used in Microsoft Excel for the StDev() function, and also
                    ' listed at Wikipedia: http://en.wikipedia.org/wiki/Standard_deviation

                    dblMean = Mean()

                    ' dblValue holds the sum of (x - dblMean)^2
                    dblValue = 0
                    For i = 0 To mCnt - 1
                        dblValue += (mQSDbl(i) - dblMean) ^ 2
                    Next i

                    ' The standard deviation is the square root of dblValue divided by n-1
                    Return Math.Sqrt(dblValue / (mCnt - 1))

                Else

                    ' The following method computes the unbiased standard deviation using running sums
                    ' It can give odd results due to round-off error
                    dblSum = Sum()
                    dblSumSq = SumSquared()
                    dblValue = dblSumSq / (mCnt - 1) - (dblSum ^ 2) / (mCnt * (mCnt - 1))
                    If dblValue > 0 Then
                        Return Math.Sqrt(dblValue)
                    Else
                        Return 0
                    End If
                End If

            ElseIf mCnt = 1 Then
                Return 0
            Else
                Return -1
            End If
        Catch ex As Exception
            Return -1
        End Try
    End Function

End Class
