using ExcelMapper.Abstractions;

namespace ExcelMapper.Abstractions.Tests;

public class CellMapperResultTests
{
    [Fact]
    public void Success_Invoke_ReturnsExpected()
    {
        var result = CellMapperResult.Success("value");
        Assert.Equal("value", result.Value);
        Assert.Null(result.Exception);
        Assert.Equal(CellMapperResult.HandleAction.UseResultAndStopMapping, result.Action);
        Assert.True(result.Succeeded);
    }

    [Fact]
    public void SuccessIfNoOtherSuccess_Invoke_ReturnsExpected()
    {
        var result = CellMapperResult.SuccessIfNoOtherSuccess("value");
        Assert.Equal("value", result.Value);
        Assert.Null(result.Exception);
        Assert.Equal(CellMapperResult.HandleAction.UseResultAndContinueMapping, result.Action);
        Assert.True(result.Succeeded);
    }

    [Fact]
    public void Ignore_Invoke_ReturnsExpected()
    {
        var result = CellMapperResult.Ignore();
        Assert.Null(result.Value);
        Assert.Null(result.Exception);
        Assert.Equal(CellMapperResult.HandleAction.IgnoreResultAndContinueMapping, result.Action);
        Assert.False(result.Succeeded);
    }

    [Fact]
    public void Invalid_Exception_ReturnsExpected()
    {
        var exception = new FormatException();
        var result = CellMapperResult.Invalid(exception);
        Assert.Null(result.Value);
        Assert.Same(exception, result.Exception);
        Assert.Equal(CellMapperResult.HandleAction.ErrorAndContinueMapping, result.Action);
        Assert.False(result.Succeeded);
    }

    [Fact]
    public void Invalid_ExceptionFactory_CreatesExceptionWhenRead()
    {
        var calls = 0;
        var result = CellMapperResult.Invalid(() =>
        {
            calls++;
            return new FormatException("message");
        });

        // The factory isn't called until the exception is needed.
        Assert.Equal(0, calls);
        Assert.Null(result.Value);
        Assert.Equal(CellMapperResult.HandleAction.ErrorAndContinueMapping, result.Action);
        Assert.False(result.Succeeded);
        Assert.Equal(0, calls);

        var exception = Assert.IsType<FormatException>(result.Exception);
        Assert.Equal("message", exception.Message);
        Assert.Equal(1, calls);
    }

    [Fact]
    public void Invalid_NullExceptionFactory_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("exceptionFactory", () => CellMapperResult.Invalid((Func<Exception>)null!));
    }

    [Fact]
    public void Default_ReturnsExpected()
    {
        CellMapperResult result = default;
        Assert.Null(result.Value);
        Assert.Null(result.Exception);
        Assert.Equal(CellMapperResult.HandleAction.UseResultAndContinueMapping, result.Action);
        Assert.True(result.Succeeded);
    }
}
