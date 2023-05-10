// Copyright (c) Microsoft. All rights reserved.

using System.Threading;
using System.Threading.Tasks;

namespace Microsoft.SemanticKernel.Skills.MsGraph;

/// <summary>
/// Interface for notes documents (e.g. OneNote).
/// </summary>
public interface INoteConnector
{
    /// <summary>
    /// Get the content of a page.
    /// </summary>
    /// <param name="name">Name of the notebook.</param>
    /// <param name="path">Path to the page.</param>
    /// <param name="cancellationToken">The <see cref="CancellationToken"/> to monitor for cancellation requests. The default is <see cref="CancellationToken.None"/>.</param>
    Task<string> GetPageContentAsync(string name, string path, CancellationToken cancellationToken = default);
}
