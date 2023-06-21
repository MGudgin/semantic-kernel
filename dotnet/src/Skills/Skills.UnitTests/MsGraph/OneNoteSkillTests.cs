// Copyright (c) Microsoft. All rights reserved.

using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.SemanticKernel.Orchestration;
using Microsoft.SemanticKernel.Skills.MsGraph;
using Moq;
using SemanticKernel.Skills.UnitTests.XunitHelpers;
using Xunit;
using Xunit.Abstractions;

namespace SemanticKernel.Skills.UnitTests.MsGraph;

public class OneNoteSkillTests : IDisposable
{
    private readonly XunitLogger<SKContext> _logger;
    private bool _disposedValue = false;

    public OneNoteSkillTests(ITestOutputHelper output)
    {
        this._logger = new XunitLogger<SKContext>(output);
    }

    [Fact]
    public async Task ReadTextAsyncAsyncSucceedsAsync()
    {
        // Arrange
        string notebookName = "Mine";
        string pageContent = "Some text content";

        MemoryStream s = new();
        using var writer = new StreamWriter(s, Encoding.UTF8);
        await writer.WriteAsync(pageContent);
        await writer.FlushAsync();
        s.Seek(0, SeekOrigin.Begin);

        Mock<INoteConnector> connectorMock = new();
        connectorMock.Setup(c => c.GetPageContentStreamAsync(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<CancellationToken>())).ReturnsAsync(s);
        NoteSkill target = new(connectorMock.Object);

        // Act
        string actual = await target.GetPageContentAsync(notebookName, "Journal/2022/2022-05/2022-05-05", CancellationToken.None);

        // Assert
        Assert.Equal(pageContent, actual);
        connectorMock.VerifyAll();
    }

    protected virtual void Dispose(bool disposing)
    {
        if (!this._disposedValue)
        {
            if (disposing)
            {
                this._logger.Dispose();
            }

            this._disposedValue = true;
        }
    }

    public void Dispose()
    {
        // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        this.Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }
}
