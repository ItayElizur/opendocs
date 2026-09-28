using System;
using Xunit;
using OfficeAi.Shared;

public class RecurrenceValidatorTests
{
    [Fact]
    public void Daily_Valid_NoOtherFieldsRequired()
    {
        string error;
        RecurrenceSpec spec = RecurrenceValidator.Parse("daily", 1, null, null, null, null, null, null, out error);
        Assert.Null(error);
        Assert.NotNull(spec);
        Assert.Equal("daily", spec.Type);
        Assert.Equal(1, spec.Interval);
    }

    [Fact]
    public void InvalidType_ListsValidValues()
    {
        string error;
        RecurrenceSpec spec = RecurrenceValidator.Parse("hourly", 1, null, null, null, null, null, null, out error);
        Assert.Null(spec);
        Assert.Contains("hourly", error);
        Assert.Contains("daily", error);
        Assert.Contains("weekly", error);
        Assert.Contains("monthlyNth", error);
        Assert.Contains("yearlyNth", error);
    }

    [Fact]
    public void IntervalLessThanOne_Rejected()
    {
        string error;
        RecurrenceSpec spec = RecurrenceValidator.Parse("daily", 0, null, null, null, null, null, null, out error);
        Assert.Null(spec);
        Assert.Contains("interval", error);
    }

    [Fact]
    public void CountAndUntilBothSet_Rejected()
    {
        string error;
        RecurrenceSpec spec = RecurrenceValidator.Parse("daily", 1, null, null, null, null, 5, DateTime.Today, out error);
        Assert.Null(spec);
        Assert.Contains("mutually exclusive", error);
    }

    [Fact]
    public void CountLessThanOne_Rejected()
    {
        string error;
        RecurrenceSpec spec = RecurrenceValidator.Parse("daily", 1, null, null, null, null, 0, null, out error);
        Assert.Null(spec);
        Assert.Contains("count", error);
    }

    [Fact]
    public void Weekly_MissingDaysOfWeek_Rejected()
    {
        string error;
        RecurrenceSpec spec = RecurrenceValidator.Parse("weekly", 1, null, null, null, null, null, null, out error);
        Assert.Null(spec);
        Assert.Contains("days_of_week", error);
        Assert.Contains("weekly", error);
    }

    [Fact]
    public void Weekly_Valid_NormalizesDayCase()
    {
        string error;
        RecurrenceSpec spec = RecurrenceValidator.Parse("weekly", 2, new[] { "Monday", "WEDNESDAY" }, null, null, null, null, null, out error);
        Assert.Null(error);
        Assert.Equal(new[] { "monday", "wednesday" }, spec.DaysOfWeek);
        Assert.Equal(2, spec.Interval);
    }

    [Fact]
    public void UnknownDayName_ListsValidDays()
    {
        string error;
        RecurrenceSpec spec = RecurrenceValidator.Parse("weekly", 1, new[] { "Funday" }, null, null, null, null, null, out error);
        Assert.Null(spec);
        Assert.Contains("Funday", error);
        Assert.Contains("sunday", error);
        Assert.Contains("saturday", error);
    }

    [Fact]
    public void Monthly_MissingDayOfMonth_Rejected()
    {
        string error;
        RecurrenceSpec spec = RecurrenceValidator.Parse("monthly", 1, null, null, null, null, null, null, out error);
        Assert.Null(spec);
        Assert.Contains("day_of_month", error);
    }

    [Fact]
    public void Monthly_DayOfMonthOutOfRange_Rejected()
    {
        string error;
        RecurrenceSpec spec = RecurrenceValidator.Parse("monthly", 1, null, 32, null, null, null, null, out error);
        Assert.Null(spec);
        Assert.Contains("1-31", error);
    }

    [Fact]
    public void Monthly_Valid()
    {
        string error;
        RecurrenceSpec spec = RecurrenceValidator.Parse("monthly", 1, null, 15, null, null, null, null, out error);
        Assert.Null(error);
        Assert.Equal(15, spec.DayOfMonth);
    }

    [Fact]
    public void MonthlyNth_MissingInstanceAndDay_Rejected()
    {
        string error;
        RecurrenceSpec spec = RecurrenceValidator.Parse("monthlyNth", 1, null, null, null, null, null, null, out error);
        Assert.Null(spec);
        Assert.Contains("instance", error);
    }

    [Fact]
    public void MonthlyNth_InstanceOutOfRange_Rejected()
    {
        string error;
        RecurrenceSpec spec = RecurrenceValidator.Parse("monthlyNth", 1, new[] { "tuesday" }, null, 6, null, null, null, out error);
        Assert.Null(spec);
        Assert.Contains("1-4", error);
    }

    [Fact]
    public void MonthlyNth_MultipleDays_Rejected()
    {
        string error;
        RecurrenceSpec spec = RecurrenceValidator.Parse("monthlyNth", 1, new[] { "monday", "tuesday" }, null, 2, null, null, null, out error);
        Assert.Null(spec);
        Assert.Contains("exactly one day", error);
    }

    [Fact]
    public void MonthlyNth_Valid_SecondTuesday()
    {
        string error;
        RecurrenceSpec spec = RecurrenceValidator.Parse("monthlyNth", 1, new[] { "tuesday" }, null, 2, null, null, null, out error);
        Assert.Null(error);
        Assert.Equal(2, spec.Instance);
        Assert.Equal(new[] { "tuesday" }, spec.DaysOfWeek);
    }

    [Fact]
    public void MonthlyNth_LastInstance_Valid()
    {
        string error;
        RecurrenceSpec spec = RecurrenceValidator.Parse("monthlyNth", 1, new[] { "friday" }, null, 5, null, null, null, out error);
        Assert.Null(error);
        Assert.Equal(5, spec.Instance);
    }

    [Fact]
    public void Yearly_MissingMonthAndDay_Rejected()
    {
        string error;
        RecurrenceSpec spec = RecurrenceValidator.Parse("yearly", 1, null, null, null, null, null, null, out error);
        Assert.Null(spec);
        Assert.Contains("month_of_year", error);
    }

    [Fact]
    public void Yearly_Valid()
    {
        string error;
        RecurrenceSpec spec = RecurrenceValidator.Parse("yearly", 1, null, 25, null, 12, null, null, out error);
        Assert.Null(error);
        Assert.Equal(25, spec.DayOfMonth);
        Assert.Equal(12, spec.MonthOfYear);
    }

    [Fact]
    public void YearlyNth_MissingMonthOfYear_Rejected()
    {
        string error;
        RecurrenceSpec spec = RecurrenceValidator.Parse("yearlyNth", 1, new[] { "monday" }, null, 1, null, null, null, out error);
        Assert.Null(spec);
        Assert.Contains("month_of_year", error);
    }

    [Fact]
    public void YearlyNth_Valid()
    {
        string error;
        RecurrenceSpec spec = RecurrenceValidator.Parse("yearlyNth", 1, new[] { "monday" }, null, 1, 9, null, null, out error);
        Assert.Null(error);
        Assert.Equal(1, spec.Instance);
        Assert.Equal(9, spec.MonthOfYear);
    }

    [Fact]
    public void Count_SetsEndCondition()
    {
        string error;
        RecurrenceSpec spec = RecurrenceValidator.Parse("daily", 1, null, null, null, null, 10, null, out error);
        Assert.Null(error);
        Assert.Equal(10, spec.Count);
        Assert.Null(spec.Until);
    }

    [Fact]
    public void Until_SetsEndCondition()
    {
        DateTime until = new DateTime(2027, 1, 1);
        string error;
        RecurrenceSpec spec = RecurrenceValidator.Parse("daily", 1, null, null, null, null, null, until, out error);
        Assert.Null(error);
        Assert.Equal(until, spec.Until);
        Assert.Null(spec.Count);
    }

    [Fact]
    public void NeitherCountNorUntil_NoEndDate()
    {
        string error;
        RecurrenceSpec spec = RecurrenceValidator.Parse("daily", 1, null, null, null, null, null, null, out error);
        Assert.Null(error);
        Assert.Null(spec.Count);
        Assert.Null(spec.Until);
    }
}
