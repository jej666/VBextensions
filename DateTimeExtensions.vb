' ***********************************************************************
' Assembly         : Infrastructure
' Author           : jerome
' Created          : 02-26-2013
'
' Last Modified By : jerome
' Last Modified On : 06-04-2013
' ***********************************************************************
' <copyright file="DateTimeExtensions.vb" company="Info Santé">
'     Info Santé. All rights reserved.
' </copyright>
' <summary></summary>
' ***********************************************************************
Imports System.Globalization
Imports System.Runtime.CompilerServices

''' <summary>
'''
''' </summary>
<HideModuleName(), System.ComponentModel.ImmutableObject(True)>
Public Module DateTimeExtensions

    ''' <summary>
    ''' Enumerates the sequence and invokes the given action for each value in the sequence.
    ''' </summary>
    ''' <param name="source">Source sequence.</param>
    ''' <param name="onNext">Action to invoke for each element.</param>
    <Extension()>
    Public Sub ForEachDay(source As Date, until As Date, onNext As Action(Of Date))
        onNext.AssertNotNull(Function() onNext)
        If source.IsAfter(until) Then
            Throw New InvalidOperationException("Source date must be anterior to until date")
        End If
        Dim d = source
        Do Until d > until
            onNext(d)
            d = d.AddDays(1)
        Loop
    End Sub

    <Extension()>
    Public Function FirstDayOfMonth(source As Date) As Date
        Return source.AddDays(-(source.Day - 1))
    End Function

    ''' <summary>
    ''' YYYYMMDD in DD/MM/YYYY
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function ToShortDate(value As String) As String
        Return String.Format(CultureInfo.InvariantCulture,
                             "{0}/{1}/{2}",
                             value.Substring(6, 2),
                             value.Substring(4, 2),
                             value.Substring(0, 4))
    End Function

    ''' <summary>
    ''' date in DD/MM/YYYY
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function ToShortDate(value As Date) As String
        Return value.ToString("dd/MM/yyyy", CultureInfo.CurrentCulture)
    End Function

    ''' <summary>in DD/MM/YYYY
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function ToShortDate(value As Integer) As String
        Return String.Format(CultureInfo.InvariantCulture,
                             "{0}/{1}/{2}",
                             value.ToString.Substring(6, 2),
                             value.ToString.Substring(4, 2),
                             value.ToString.Substring(0, 4))
    End Function

    ''' <summary>
    ''' DDMMYYYY in YYYYMMDD
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function ToDateInDBformat(value As String) As Integer
        Return CInt(String.Format(CultureInfo.CurrentCulture,
                                  "{0}{1}{2}",
                                  value.Substring(6, 4),
                                  value.Substring(3, 2),
                                  value.Substring(0, 2)))
    End Function

    ''' <summary>
    ''' DDMMYYY in YYYYMMDD
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function ToDateInDBformat(value As Integer) As Integer
        Return CInt(String.Format(CultureInfo.CurrentCulture,
                                  "{0}{1}{2}",
                                  value.ToString.Substring(6, 4),
                                  value.ToString.Substring(3, 2),
                                  value.ToString.Substring(0, 2)))
    End Function

    ''' <summary>
    ''' DD/MM/YYYY in YYYYMMDD
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function ToDateInDBformat(value As DateTime) As Integer
        Return CInt(value.Year &
                    value.Month.ToString("00") &
                    value.Day.ToString("00"))
    End Function

    ''' <summary>
    ''' YYYYMMDD converti en DD/MM/YYYY
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function ToDate(value As Integer) As Date
        Return New Date(CInt(value.ToString(CultureInfo.CurrentCulture).Substring(0, 4)),
                        CInt(value.ToString(CultureInfo.CurrentCulture).Substring(4, 2)),
                        CInt(value.ToString(CultureInfo.CurrentCulture).Substring(6, 2)))
    End Function

    ''' <summary>
    ''' YYYYMMDD converti en DD/MM/YYYY
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function ToDate(value As String) As Date
        Return New Date(CInt(value.Substring(0, 4)),
                        CInt(value.Substring(4, 2)),
                        CInt(value.Substring(6, 2)))
    End Function

    ''' <summary>
    ''' HHMM converti en DateTime
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function ToTime(value As Integer) As DateTime
        If IsNothing(value) OrElse
            value.ToString.Length > 4 Then
            Return Nothing
        End If
        Select Case value.ToString.Length
            Case 3
                Return New DateTime(1, 1, 1,
                                    CInt(value.ToString(CultureInfo.CurrentCulture).Substring(0, 1)),
                                    CInt(value.ToString(CultureInfo.CurrentCulture).Substring(1, 2)), 0)
            Case 4
                Return New DateTime(1, 1, 1,
                                    CInt(value.ToString(CultureInfo.CurrentCulture).Substring(0, 2)),
                                    CInt(value.ToString(CultureInfo.CurrentCulture).Substring(2, 2)), 0)
            Case Else
                Return Nothing
        End Select
    End Function

    ''' <summary>
    ''' HHMM converti en DateTime
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function ToTime(value As String) As DateTime
        If String.IsNullOrEmpty(value) OrElse
            value.Length > 4 Then
            Return Nothing
        End If
        Select Case value.ToString.Length
            Case 3
                Return CDate(String.Format(CultureInfo.CurrentCulture,
                                            "{0}{1}",
                                            value.Substring(0, 1),
                                            value.Substring(1, 2)))
            Case 4
                Return CDate(String.Format(CultureInfo.CurrentCulture,
                                            "{0}{1}",
                                            value.Substring(0, 2),
                                            value.Substring(2, 2)))
            Case Else
                Return Nothing
        End Select
    End Function

    ''' <summary>
    ''' DateTime converti en HH:MM
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function ToTimeString(value As DateTime) As String
        Return String.Format(CultureInfo.CurrentCulture,
                             "{0}:{1}",
                             value.TimeOfDay.Hours.ToString(CultureInfo.CurrentCulture).PadLeft(2, "0"c),
                             value.TimeOfDay.Minutes.ToString(CultureInfo.CurrentCulture).PadLeft(2, "0"c))
    End Function

    ''' <summary>
    ''' HHMM converti en HH:MM
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function ToShortTime(value As Integer) As String
        If IsNothing(value) OrElse
            value.ToString.Length > 4 Then
            Return String.Empty
        End If
        Select Case value.ToString.Length
            Case 3 : Return String.Format(CultureInfo.CurrentCulture,
                                          "0{0}:{1}",
                                          value.ToString(CultureInfo.CurrentCulture).Substring(0, 1),
                                          value.ToString(CultureInfo.CurrentCulture).Substring(1, 2))
            Case 4 : Return String.Format(CultureInfo.CurrentCulture,
                                          "{0}:{1}",
                                          value.ToString(CultureInfo.CurrentCulture).Substring(0, 2),
                                          value.ToString(CultureInfo.CurrentCulture).Substring(2, 2))
            Case Else
                Return String.Empty
        End Select
    End Function

    ''' <summary>
    ''' HHMM converti en HH:MM
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function ToShortTime(value As String) As String
        If IsNothing(value) OrElse
            value.ToString.Length > 4 Then
            Return String.Empty
        End If
        Select Case value.Length
            Case 3
                Return String.Format(CultureInfo.CurrentCulture,
                                    "0{0}:{1}",
                                    value.Substring(0, 1),
                                    value.Substring(1, 2))
            Case 4
                Return String.Format(CultureInfo.CurrentCulture,
                                    "{0}:{1}",
                                    value.Substring(0, 2),
                                    value.Substring(2, 2))
            Case Else
                Return String.Empty
        End Select
    End Function

    ''' <summary>
    ''' HH:MM converti en HHMM
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function ToTimeInDBformat(value As String) As Integer
        If String.IsNullOrEmpty(value) OrElse
            value.Length > 5 Then
            Return 0
        End If

        Select Case value.ToString.Length
            Case 4
                Return CInt(String.Format(CultureInfo.CurrentCulture,
                                          "{0}{1}",
                                          value.Substring(0, 1),
                                          value.Substring(2, 2)))
            Case 5
                Return CInt(String.Format(CultureInfo.CurrentCulture,
                                          "{0}{1}",
                                          value.Substring(0, 2),
                                          value.Substring(3, 2)))
            Case Else
                Return 0
        End Select
    End Function

    ''' <summary>
    ''' HH:MM converti en HHMM
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function ToTimeInDBformat(value As DateTime) As Integer
        Return CInt(String.Format(CultureInfo.CurrentCulture,
                                  "{0}{1}",
                                  value.TimeOfDay.Hours,
                                  value.TimeOfDay.Minutes))
    End Function

    ''' <summary>
    ''' Numéro de jour dans la semaine converti en lun, mar, ...
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function ToShortDayName(value As Integer) As String
        If value >= 0 And value < 8 Then
            Return DateTimeFormatInfo.CurrentInfo.DayNames(value).Substring(0, 3)
        End If
        Return String.Empty
    End Function

    ''' <summary>
    ''' Numéro de jour dans la semaine converti en lun, mar, ...
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function ToShortDayName(value As String) As String
        If Not value.IsNumeric() Then
            Return String.Empty
        End If
        Dim i = CInt(value)
        If i >= 0 And i < 8 Then
            Return DateTimeFormatInfo.CurrentInfo.DayNames(i).Substring(0, 3)
        End If
        Return String.Empty
    End Function

    ''' <summary>
    ''' Numéro de jour dans la semaine converti en lun, mar, ...
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function ToShortDayName(value As DayOfWeek) As String
        Return DateTimeFormatInfo.CurrentInfo.DayNames(value).Substring(0, 3)
    End Function

    ''' <summary>
    ''' LUN 25/06 ...
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function ToShortDayAndDate(value As Date) As String
        Return String.Format(CultureInfo.CurrentCulture,
                             "{0} {1}",
                             value.ToShortDayName,
                             value.ToString("dd/MM"))
    End Function

    ''' <summary>
    ''' Numéro de jour dans la semaine converti en lun, mar, ...
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function ToShortDayName(value As Date) As String
        Return DateTimeFormatInfo.CurrentInfo.DayNames(value.DayOfWeek).Substring(0, 3)
    End Function

    ''' <summary>
    ''' Numéro de jour dans la semaine converti en lundi, mardi, ...
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function ToDayName(value As Integer) As String
        If value >= 0 And value < 8 Then
            Return DateTimeFormatInfo.CurrentInfo.DayNames(value)
        End If
        Return String.Empty
    End Function

    ''' <summary>
    ''' Numéro de jour dans la semaine converti en lundi, mardi, ...
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function ToDayName(value As DayOfWeek) As String
        Return DateTimeFormatInfo.CurrentInfo.DayNames(value)
    End Function

    ''' <summary>
    ''' Numéro de jour dans la semaine converti en lundi, mardi, ...
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function ToDayName(value As Date) As String
        Return DateTimeFormatInfo.CurrentInfo.DayNames(value.DayOfWeek)
    End Function

    ''' <summary>
    ''' Retourne le premier jour de la semaine indiqué par le paramètre
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function FirstWeekDay(value As Date) As Date
        Dim firstDayWeek = CultureInfo.CurrentCulture.DateTimeFormat.FirstDayOfWeek
        While value.DayOfWeek <> firstDayWeek
            value = value.AddDays(-1)
        End While
        Return value
    End Function

    ''' <summary>
    ''' Retourne le premier jour de la semaine indiqué par le paramètre
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function FirstWeekDay(value As String) As Date
        Return value.ToDate.FirstWeekDay
    End Function

    ''' <summary>
    ''' Retourne le premier jour de la semaine indiqué par le paramètre
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function FirstWeekDay(value As Integer) As Date
        Return value.ToDate.FirstWeekDay
    End Function

    ''' <summary>
    ''' Retourne le dernier jour de la semaine indiqué par le paramètre
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function LastWeekDay(value As Date) As Date
        Return value.FirstWeekDay.AddDays(6)
    End Function

    ''' <summary>
    ''' Retourne le premier jour de la semaine indiqué par le paramètre
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function LastWeekDay(value As String) As Date
        Return value.ToDate.LastWeekDay
    End Function

    ''' <summary>
    ''' Retourne le premier jour de la semaine indiqué par le paramètre
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()>
    Public Function LastWeekDay(value As Integer) As Date
        Return value.ToDate.LastWeekDay
    End Function

    ''' <summary>
    ''' Returns the next occurrence of the day of the week specified.
    ''' </summary>
    <Extension()> _
    Public Function GetNext(dt As DateTime, dayOfWeek As DayOfWeek) As DateTime
        Dim daysToAdd As Integer = If(dt.DayOfWeek < dayOfWeek,
                                      dayOfWeek - dt.DayOfWeek,
                                      (7 - dt.DayOfWeek) + dayOfWeek)
        Return dt.AddDays(daysToAdd)
    End Function

    ''' <summary>
    ''' Returns the prior occurrence of the day of the week specified.
    ''' </summary>
    <Extension()> _
    Public Function GetLast(dt As DateTime, dayOfWeek As DayOfWeek) As DateTime
        Dim daysToSubtract As Integer = If(dt.DayOfWeek > dayOfWeek,
                                           dt.DayOfWeek - dayOfWeek,
                                           (7 - dayOfWeek) + dt.DayOfWeek)
        Return dt.AddDays(daysToSubtract * -1)
    End Function

    ''' <summary>
    ''' 	Determines whether the date only part of twi DateTime values are equal.
    ''' </summary>
    ''' <param name = "date">The date.</param>
    ''' <param name = "dateToCompare">The date to compare with.</param>
    ''' <returns>
    ''' 	<c>true</c> if both date values are equal; otherwise, <c>false</c>.
    ''' </returns>
    <Extension()> _
    Public Function IsDateEqual([date] As DateTime, dateToCompare As DateTime) As Boolean
        Return ([date].Date = dateToCompare.Date)
    End Function

    ''' <summary>
    ''' Teste si la date est comprise dans l'intervalle
    ''' </summary>
    ''' <param name="inBetweenDate"></param>
    ''' <param name="startDate"></param>
    ''' <param name="endDate"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()> _
    Public Function IsDateBetween(inBetweenDate As DateTime, startDate As DateTime, endDate As DateTime) As Boolean
        Return inBetweenDate >= startDate AndAlso
            inBetweenDate <= endDate
    End Function

    ''' <summary>
    ''' Teste si la valeur entrée est de type date
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Extension()> _
    Public Function IsDate(value As String) As Boolean
        If value.IsNotEmpty Then
            Dim dt As DateTime
            Return DateTime.TryParse(value, dt)
        Else
            Return False
        End If
    End Function

    ''' <summary>
    ''' 	Indicates whether the specified date is a weekend (Saturday or Sunday).
    ''' </summary>
    ''' <param name = "date">The date.</param>
    ''' <returns>
    ''' 	<c>true</c> if the specified date is a weekend; otherwise, <c>false</c>.
    ''' </returns>
    <Extension()> _
    Public Function IsWeekend([date] As DateTime) As Boolean
        Return [date].DayOfWeek = DayOfWeek.Saturday OrElse
                [date].DayOfWeek = DayOfWeek.Sunday
    End Function

    ''' <summary>
    '''     Indicates whether the source DateTime is before the supplied DateTime.
    ''' </summary>
    ''' <param name="source">The source DateTime.</param>
    ''' <param name="other">The compared DateTime.</param>
    ''' <returns>True if the source is before the other DateTime, False otherwise</returns>
    <Extension()> _
    Public Function IsBefore(source As DateTime, other As DateTime) As Boolean
        Return source.CompareTo(other) < 0
    End Function

    ''' <summary>
    '''     Indicates whether the source DateTime is before the supplied DateTime.
    ''' </summary>
    ''' <param name="source">The source DateTime.</param>
    ''' <param name="other">The compared DateTime.</param>
    ''' <returns>True if the source is before the other DateTime, False otherwise</returns>
    <Extension()> _
    Public Function IsAfter(source As DateTime, other As DateTime) As Boolean
        Return source.CompareTo(other) > 0
    End Function

    ''' <summary>
    ''' Gest the elapsed seconds since the input DateTime
    '''</summary>
    ''' <param name="value">Input DateTime</param>
    ''' <returns>Returns a Double value with the elapsed seconds since the input DateTime</returns>
    ''' <example>
    ''' Double elapsed = dtStart.ElapsedSeconds();
    ''' </example>
    <Extension()>
    Public Function ElapsedSeconds(value As DateTime) As Double
        Return DateTime.Now.Subtract(value).TotalSeconds
    End Function

    ''' <summary>
    ''' 	Calculates the age based on today.
    ''' </summary>
    ''' <param name = "dateOfBirth">The date of birth.</param>
    ''' <returns>The calculated age.</returns>
    <Extension()> _
    Public Function CalculateAge(dateOfBirth As DateTime) As Integer
        Return dateOfBirth.CalculateAge(Now.[Date])
    End Function

    ''' <summary>
    ''' 	Calculates the age based on a passed reference date.
    ''' </summary>
    ''' <param name = "dateOfBirth">The date of birth.</param>
    ''' <param name = "referenceDate">The reference date to calculate on.</param>
    ''' <returns>The calculated age.</returns>
    <Extension()> _
    Public Function CalculateAge(dateOfBirth As DateTime, referenceDate As DateTime) As Integer
        Dim years = referenceDate.Year - dateOfBirth.Year
        If referenceDate.Month < dateOfBirth.Month OrElse
            (referenceDate.Month = dateOfBirth.Month AndAlso
             referenceDate.Day < dateOfBirth.Day) Then
            years -= 1
        End If
        Return years
    End Function

    <Extension()> _
    Public Function January(day As Integer, year As Integer) As DateTime
        Return New DateTime(year, 1, day)
    End Function

    <Extension()> _
    Public Function February(day As Integer, year As Integer) As DateTime
        Return New DateTime(year, 2, day)
    End Function

    <Extension()> _
    Public Function March(day As Integer, year As Integer) As DateTime
        Return New DateTime(year, 3, day)
    End Function

    <Extension()> _
    Public Function April(day As Integer, year As Integer) As DateTime
        Return New DateTime(year, 4, day)
    End Function

    <Extension()> _
    Public Function May(day As Integer, year As Integer) As DateTime
        Return New DateTime(year, 5, day)
    End Function

    <Extension()> _
    Public Function June(day As Integer, year As Integer) As DateTime
        Return New DateTime(year, 6, day)
    End Function

    <Extension()> _
    Public Function July(day As Integer, year As Integer) As DateTime
        Return New DateTime(year, 7, day)
    End Function

    <Extension()> _
    Public Function August(day As Integer, year As Integer) As DateTime
        Return New DateTime(year, 8, day)
    End Function

    <Extension()> _
    Public Function September(day As Integer, year As Integer) As DateTime
        Return New DateTime(year, 9, day)
    End Function

    <Extension()> _
    Public Function October(day As Integer, year As Integer) As DateTime
        Return New DateTime(year, 10, day)
    End Function

    <Extension()> _
    Public Function November(day As Integer, year As Integer) As DateTime
        Return New DateTime(year, 11, day)
    End Function

    <Extension()> _
    Public Function December(day As Integer, year As Integer) As DateTime
        Return New DateTime(year, 12, day)
    End Function

    <Extension()> _
    Public Function At(d As DateTime, hour As Integer, Optional minute As Integer = 0, Optional second As Integer = 0) As DateTime
        Return New DateTime(d.Year, d.Month, d.Day, hour, minute, second)
    End Function

End Module
