Imports System.Text.RegularExpressions

''' <summary>
''' Converts feet and inches to a decimal format and vice versa within a 16th of an inch accuracy.
''' Also includes formatting functions to keep compatibility with other functions.
''' 
''' COPYRIGHT © 2018 By Alliance Steel All Rights Reserved
''' </summary>
''' <remarks>Author Austin and Ian derived from library written by Stephanie Chesser</remarks>
Public Class FeetAndInches
    Implements IDisposable

    Private QUOTE As String = """"
    ''' <summary>
    ''' Converts a decimal in inches to feet and inches.
    ''' </summary>
    ''' <param name="numInInches">The number to be passed in inches.</param>
    ''' <returns>A string of the feet and inches with a fraction.</returns>
    Public Function decimal_FeetInches(ByVal numInInches As Double) As Object
        Dim Ft As Long
        Dim Inc As Long
        Dim Sx As Long
        Dim Denom As Long
        Dim Neg As Boolean

        Dim toReturn As String = ""
        If (numInInches = 0) Then Return "0"
        If numInInches < 0 Then Neg = True Else Neg = False
        numInInches = Math.Abs(numInInches)
        Denom = 16
        Ft = Fix(numInInches / 12)
        numInInches = numInInches - (Ft * 12)
        Inc = Fix(numInInches)
        numInInches = numInInches - Inc
        Sx = Int((numInInches * 16) + 0.5)

        If Ft <> 0 Then toReturn = Ft & "' "
        If Sx = 0 Then
            toReturn = toReturn & Inc & QUOTE
        ElseIf Sx = 16 Then
            Inc = Inc + 1
            Sx = 0
            If Inc = 12 Then
                Inc = 0
                Ft = Ft + 1
                toReturn = Ft & "' "
            End If
            toReturn = toReturn & Inc & QUOTE
        Else
            If Sx Mod 8 = 0 Then
                Sx = Sx / 8
                Denom = Denom / 8
            ElseIf Sx Mod 4 = 0 Then
                Sx = Sx / 4
                Denom = Denom / 4
            ElseIf Sx Mod 2 = 0 Then
                Sx = Sx / 2
                Denom = Denom / 2
            End If
            toReturn = toReturn & Inc & " " & Sx & "/" & Denom & QUOTE
        End If
        If Ft = 0 And Inc = 0 Then toReturn = Sx & "/" & Denom & QUOTE
        If Neg Then toReturn = "-" & toReturn
        Ft = Nothing
        Inc = Nothing
        Sx = Nothing
        Denom = Nothing
        Neg = Nothing

        Return toReturn
    End Function

    Public Const STANDARD_ARCH_INPUT = "(\d+\'( |-)(\d+|\d+\/\d+)(( |-)\d+\/\d+""|"")|\d+""$|\d+\'$)"
    ''' <summary>
    ''' Takes in a string and matches it with a regular expression to ensure that it meets these requirements.
    ''' <para>
    ''' Number_A' Number_B Number_C/Number_D<para/>
    ''' Number_A'-Number_B Number_C/Number_D<para/>
    ''' Number_A'-Number_B Number_C/Number_D"<para/>
    ''' Number_A'-Number_B"<para/>
    ''' Number_A'Number_B Number_C/Number_D<para/>
    ''' Number_A-Number_B Number_C/Number_D<para/>
    ''' Number_A.Number_B
    ''' </para>
    ''' </summary>
    ''' <param name="input">The string to be checked to the regular expression.</param>
    ''' <returns>If input matches the regular expression returns true. Otherwise returns false.</returns>
    ''' <remarks>
    ''' u0027 is '
    ''' u002D is -
    ''' u002E is .
    ''' u002F is /
    ''' </remarks>
    Public Function GoodInput(ByVal input As Object) As Boolean
        Try
            'input = input.ToString().Replace("\", "")
            Dim rx As New Regex("^\d+(\u002E\d*|(([\u0027]|[\s]*)?\s*[\u002D]?\s*\d*\u0022?(\s*\d*(\u002F[1-9]+\d?|\u005c[1-9]+\d?))?\u0022?))")
            input = input.Trim()
            Dim m As Match = rx.Match(input)
            GoodInput = If(m.ToString.Length = input.Length And m.Success, True, False)
            m = Nothing
            rx = Nothing
            Return GoodInput
        Catch ex As Exception
            Throw New ApplicationException("Invalid input!", ex)
        End Try
    End Function

    ''' <summary>
    ''' Converts a given string of <see cref="GoodInput">good input</see> to decimal value in inches. <para/>
    ''' The string must fit the format listed in <see cref="GoodInput">GoodInput</see>
    ''' </summary>
    ''' <param name="input">Input as String in ft/in/sixteenths</param>
    ''' <returns>Returns decimal value in inches.</returns>
    ''' <remarks>
    ''' u0027 is '
    ''' u002D is -
    ''' u002E is .
    ''' u002F is /
    ''' u005C is \
    ''' </remarks>
    Public Function InputNum(ByVal input As Object) As Double
        'Verify input is good
        If Not GoodInput(input) Then Return -1
        'Return variable
        Dim toReturn As Double = 0

        'Regex for feet
        'Dim ftRx = New Regex("^\d+(\u0027|\u002D|$|\s)")
        Dim ftRx = New Regex("^\d+(\'|\-|\'\-|$|\s)\s*(?=$|(\d+\s*))(?!\d+([\/]|[\\])\d+\s*)")
        'Regex for inches
        'Dim inRx = New Regex("(\u002D\d+)|(\d+\s*\u0022)")
        '(?!^\d+\s*$)(?!^\d+\s+\d+(\s+|$|[""]|\u002D).*$)(^|\s)\d+([""]|\u002D|$|\s)
        '(?!^)(\d+)(?!\")(?!\/)((?=\-)|(?=\s))(?!$)
        Dim inRx = New Regex("((?<!\/|\/\d\d|\/\d)\d+(?!\/|\')([""]|[ ])(?![Gg][Aa]))")
        'Regex for sixteenths
        Dim sxRx = New Regex("(?=.*)\d+([/]|[\\])[1-9]+(?=\s*)")
        'Dim sxRx = New Regex("\d+(\u002F|\u005C)[1-9]+[\d]*")
        'Regex for decimal
        Dim dcRx = New Regex("^\d*\u002E\d*$")

        'Match variable to keep track of regex matches.
        'Starts off matching against Feet, then inches, then sixteenths.
        'Finally, checks if input is in decimal. If it is, return it as is.
        Dim m As Match = ftRx.Match(input)
        toReturn += If(m.Success, CInt(m.ToString().Replace("'", "").Replace("-", "")) * 12, 0)
        m = inRx.Match(input)
        toReturn += If(m.Success, CInt(m.ToString().Replace("""", "").Replace("-", "")), 0)
        m = sxRx.Match(input)
        If (m.Success) Then
            toReturn += If(m.ToString.Contains("/"),
                        CInt(m.ToString().Split("/")(0)) / CInt(m.ToString().Split("/")(1)),
                        CInt(m.ToString().Split("\")(0)) / CInt(m.ToString().Split("\")(1)))
        End If
        m = dcRx.Match(input)
        toReturn = If(m.Success, input, toReturn)
        m = Nothing
        ftRx = Nothing
        inRx = Nothing
        sxRx = Nothing
        dcRx = Nothing

        Return toReturn
    End Function

    'End InputNum
    ''' <summary>
    ''' Converts a decimal in inches to feet and inches. Into a readable format for an external text file for the shop to cut the sized plates.
    ''' </summary>
    ''' <param name="Num">The decimal to be changed into the string format.</param>
    ''' <returns>The string to be on the shop text file.</returns>
    Public Function decimal_Imperial(ByVal Num As Double) As String
        'converts decimal number in inches to imperial feet and inches format,
        'rounded to the nearest 1/16", for use with FMI Flange/Plate software.
        'returns formatted string containing feet,inches,sixteenth inches
        Dim Ft As Long
        Dim Inc As Long
        Dim Sx As Long
        Dim Denom As Long
        Dim Neg As Boolean
        Dim Imperial As String = ""

        If Num = 0 Then
            Imperial = "0"""
            GoTo ExitHere
        End If
        If Num < 0 Then Neg = True Else Neg = False
        Num = Math.Abs(Num)
        Denom = 16
        Ft = Fix(Num / 12)
        Num = Num - (Ft * 12)
        Inc = Fix(Num)
        Num = Num - Inc
        Sx = Int((Num * 16) + 0.5) 'rounds to nearest 1/16"
        'the following code reduces the resulting fraction to the least common denominator
        'and handles the cases where sx = 0 and sx = 16
        'also handles the case where inc = 11 and sx = 16
        If Ft <> 0 Then Imperial = Ft & "'"
        If Sx = 0 Then
            Imperial = Imperial & Inc & Chr(34)
        ElseIf Sx = 16 Then
            Inc = Inc + 1
            Sx = 0
            If Inc = 12 Then
                Inc = 0
                Ft = Ft + 1
                Imperial = Ft & "'"
            End If
            Imperial = Imperial & Inc & Chr(34)
        Else
            If Sx Mod 8 = 0 Then
                Sx = Sx / 8
                Denom = Denom / 8
            ElseIf Sx Mod 4 = 0 Then
                Sx = Sx / 4
                Denom = Denom / 4
            ElseIf Sx Mod 2 = 0 Then
                Sx = Sx / 2
                Denom = Denom / 2
            End If
            Imperial = Imperial & Inc & """" & Sx & "/" & Denom
        End If
        If Ft = 0 And Inc = 0 Then Imperial = "0""" & Sx & "/" & Denom
        If Neg Then Imperial = "-" & Imperial

        Ft = Nothing
        Inc = Nothing
        Sx = Nothing
        Denom = Nothing
        Neg = Nothing
        Imperial = Nothing

ExitHere:
        Return Imperial
    End Function
    Public Function String_Double(ByVal input As String) As Double
        Try
            Return Convert.ToDouble(input)
        Catch ex As Exception
            Return -1
        End Try
    End Function
    ''' <summary>
    ''' Rounds to sixteenths.
    ''' </summary>
    ''' <param name="Num">Decimal to be rounded.</param>
    ''' <returns>Rounded decimal</returns>
    ''' <remarks>We should find where this is called in the program.</remarks>
    Public Function RndSix(ByVal Num As Double) As Double
        RndSix = Int((Num * 16) + 0.5) / 16
    End Function

    Public Function GetGaugeInDecimal(ByVal Gauge As String) As Double
        Select Case Gauge
            Case "10"
                GetGaugeInDecimal = 0.1345
            Case "11"
                GetGaugeInDecimal = 0.1196
            Case "12"
                GetGaugeInDecimal = 0.1046
            Case "13"
                GetGaugeInDecimal = 0.0897
            Case "14"
                GetGaugeInDecimal = 0.0747
            Case "15"
                GetGaugeInDecimal = 0.0673
            Case "16"
                GetGaugeInDecimal = 0.0598
            Case "17"
                GetGaugeInDecimal = 0.0538
            Case "18"
                GetGaugeInDecimal = 0.0478
            Case "19"
                GetGaugeInDecimal = 0.0418
            Case "20"
                GetGaugeInDecimal = 0.0359
            Case "21"
                GetGaugeInDecimal = 0.0329
            Case "22"
                GetGaugeInDecimal = 0.0299
            Case "23"
                GetGaugeInDecimal = 0.0269
            Case "24"
                GetGaugeInDecimal = 0.0239
            Case "25"
                GetGaugeInDecimal = 0.0209
            Case "26"
                GetGaugeInDecimal = 0.0179
            Case Else
                GetGaugeInDecimal = -1
        End Select
        If (GetGaugeInDecimal < 0) Then
            GetGaugeInDecimal = InputNum(Gauge)
        End If
    End Function

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls
    ' IDisposable
   Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                'All refrences and values are sub and function level.

            End If
            QUOTE = Nothing
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class

