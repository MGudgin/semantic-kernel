// Copyright (c) Microsoft. All rights reserved.

using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Graph;
using Microsoft.SemanticKernel.Skills.MsGraph.Connectors.Diagnostics;
using Microsoft.SemanticKernel.Skills.MsGraph.Connectors.Exceptions;

namespace Microsoft.SemanticKernel.Skills.MsGraph.Connectors;

/// <summary>
/// Connector for OneDrive API
/// </summary>
public class OneNoteConnector : INoteConnector
{
    private readonly GraphServiceClient _graphServiceClient;

    /// <summary>
    /// Initializes a new instance of the <see cref="OneNoteConnector"/> class.
    /// </summary>
    /// <param name="graphServiceClient">A graph service client.</param>
    public OneNoteConnector(GraphServiceClient graphServiceClient)
    {
        this._graphServiceClient = graphServiceClient;
    }

    /// <inheritdoc/>
    public async Task<string> GetPageContentAsync(string name, string path, CancellationToken cancellationToken = default)
    {
        Ensure.NotNullOrWhitespace(name, nameof(name));
        Ensure.NotNullOrWhitespace(path, nameof(path));

        string[] pathParts = path.Split('/');

        if(pathParts.Length > 3)
        {
            // TODO: throw proper exception
            throw new Exception($"Path parts whould be 3 or less");
        }

        IOnenoteNotebooksCollectionPage notebooks = await this._graphServiceClient.Me.Onenote.Notebooks.Request().GetAsync(cancellationToken).ConfigureAwait(false);

        Notebook notebook = notebooks.FirstOrDefault(x => x.DisplayName.Equals(name, StringComparison.CurrentCultureIgnoreCase));

        if(notebook == null)
        {
            // TODO: throw proper exception
            throw new Exception($"Unable to find notebook {name}");
        }

        // TODO: Parse the path.
        // Possibilities:
        //
        // somepage - find the page(s) with this name
        // somesection/somepage - find the section(s) then page(s) with these names
        // somesectiongroup/somesection/somepage - find the section group(s), then sections(s) then page(s) with these names
        // section groups are nestable, does this matter?

        return string.Empty;;
    }

}
