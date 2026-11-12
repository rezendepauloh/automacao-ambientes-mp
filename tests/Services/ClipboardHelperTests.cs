using AmbientesHub.Services;
using Xunit;

namespace AmbientesHub.Tests.Services;

public class ClipboardHelperTests
{
    [Fact]
    public void Copy_ReturnsFalse_WhenTextIsEmpty()
    {
        bool resultNull = ClipboardHelper.Copy(null!);
        bool resultEmpty = ClipboardHelper.Copy("");
        bool resultWhiteSpace = ClipboardHelper.Copy("   ");

        Assert.False(resultNull);
        Assert.False(resultEmpty);
        Assert.False(resultWhiteSpace);
    }
}

