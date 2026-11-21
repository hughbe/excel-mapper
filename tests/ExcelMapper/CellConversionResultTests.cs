namespace ExcelMapper.Tests;

public class CellConversionResultTests
{
    [Fact]
    public void Success_Invoke_ReturnsExpected()
    {
        var result = CellConversionResult<int>.Success(5);
        Assert.Equal(CellConversionStatus.Success, result.Status);
        Assert.Equal(5, result.Value);
        Assert.Null(result.Exception);
    }

    [Fact]
    public void Success_Null_ReturnsExpected()
    {
        var result = CellConversionResult<string?>.Success(null);
        Assert.Equal(CellConversionStatus.Success, result.Status);
        Assert.Null(result.Value);
    }

    [Fact]
    public void Invalid_NoException_ReturnsExpected()
    {
        var result = CellConversionResult<int>.Invalid();
        Assert.Equal(CellConversionStatus.Invalid, result.Status);
        Assert.Equal(0, result.Value);
        Assert.Null(result.Exception);
    }

    [Fact]
    public void Invalid_Exception_ReturnsExpected()
    {
        var exception = new FormatException();
        var result = CellConversionResult<int>.Invalid(exception);
        Assert.Equal(CellConversionStatus.Invalid, result.Status);
        Assert.Same(exception, result.Exception);
    }

    [Fact]
    public void Invalid_NullException_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("exception", () => CellConversionResult<int>.Invalid((Exception)null!));
    }

    [Fact]
    public void Invalid_CreateException_CreatesWhenRead()
    {
        var calls = 0;
        var result = CellConversionResult<int>.Invalid(() =>
        {
            calls++;
            return new FormatException();
        });
        Assert.Equal(CellConversionStatus.Invalid, result.Status);
        Assert.Equal(0, calls);
        Assert.IsType<FormatException>(result.Exception);
        Assert.Equal(1, calls);
    }

    [Fact]
    public void Invalid_NullCreateException_ThrowsArgumentNullException()
    {
        Assert.Throws<ArgumentNullException>("createException", () => CellConversionResult<int>.Invalid((Func<Exception>)null!));
    }

    [Fact]
    public void NotHandled_Get_ReturnsExpected()
    {
        var result = CellConversionResult<int>.NotHandled;
        Assert.Equal(CellConversionStatus.NotHandled, result.Status);
        Assert.Null(result.Exception);
    }

    [Fact]
    public void Default_Get_IsNotHandled()
    {
        Assert.Equal(CellConversionStatus.NotHandled, default(CellConversionResult<int>).Status);
    }
}
