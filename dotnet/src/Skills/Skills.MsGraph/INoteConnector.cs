// Copyright (c) Microsoft. All rights reserved.

using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Microsoft.SemanticKernel.Skills.MsGraph;

/// <summary>
/// Interface for notes documents (e.g. OneNote).
/// </summary>
public interface INoteConnector
{
    /// <summary>
    /// Get the content stream for a page.
    /// </summary>
    /// <param name="notebookName">Name of the notebook.</param>
    /// <param name="path">Path to the page.</param>
    /// <param name="cancellationToken">The <see cref="CancellationToken"/> to monitor for cancellation requests. The default is <see cref="CancellationToken.None"/>.</param>
    Task<Stream> GetPageContentStreamAsync(string notebookName, string path, CancellationToken cancellationToken = default);

    /// <summary>
    /// Get the content stream for a section.
    /// </summary>
    /// <param name="notebookName">Name of the notebook.</param>
    /// <param name="path">Path to the section.</param>
    /// <param name="cancellationToken">The <see cref="CancellationToken"/> to monitor for cancellation requests. The default is <see cref="CancellationToken.None"/>.</param>
    Task<Stream> GetSectionContentStreamAsync(string notebookName, string path, CancellationToken cancellationToken = default);
}
