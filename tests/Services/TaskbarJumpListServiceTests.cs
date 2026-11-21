using AmbientesHub.Services;
using Xunit;

namespace AmbientesHub.Tests.Services;

public class TaskbarJumpListServiceTests
{
    [Theory]
    [InlineData(null, null)]
    [InlineData(new string[0], null)]
    [InlineData(new[] { "--other", "foo" }, null)]
    [InlineData(new[] { "--jump-action=" }, null)]
    [InlineData(new[] { "--jump-action=exit" }, "exit")]
    [InlineData(new[] { "--foo", "--JUMP-ACTION=env:Chamados" }, "env:Chamados")]
    public void ExtractJumpAction_ParsesExpectedAction(string[]? args, string? expected)
    {
        Assert.Equal(expected, TaskbarJumpListService.ExtractJumpAction(args));
    }

    [Fact]
    public async Task PipeServerAndClient_DeliversJumpActionEndToEnd()
    {
        string pipeName = $"AmbientesHub_TestPipe_{Guid.NewGuid():N}";
        var tcs = new TaskCompletionSource<string>(TaskCreationOptions.RunContinuationsAsynchronously);

        using var cts = TaskbarJumpListService.StartPipeServer(action => tcs.TrySetResult(action), pipeName);
        await Task.Delay(50);

        bool sent = TaskbarJumpListService.TrySendActionToRunningInstance("exit", pipeName, timeoutMs: 2000);
        Assert.True(sent);

        string received = await tcs.Task.WaitAsync(TimeSpan.FromSeconds(3));
        Assert.Equal("exit", received);

        cts.Cancel();
    }
}
