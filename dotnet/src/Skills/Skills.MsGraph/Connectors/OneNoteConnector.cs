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
using static System.Collections.Specialized.BitVector32;

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
    public async Task<Stream> GetPageContentStreamAsync(string notebookName, string path, CancellationToken cancellationToken = default)
    {
        Ensure.NotNullOrWhitespace(notebookName, nameof(notebookName));
        Ensure.NotNullOrWhitespace(path, nameof(path));

        string[] pathParts = path.Split('/');

        if (pathParts.Length < 2)
        {
            // TODO: throw proper exception
            throw new ArgumentException($"Path should have 2 or more parts, was {pathParts.Length}");
        }

        IOnenoteNotebooksCollectionPage notebooks = await this._graphServiceClient.Me.Onenote.Notebooks.Request().GetAsync(cancellationToken).ConfigureAwait(false);

        Notebook notebook = notebooks.FirstOrDefault(x => x.DisplayName.Equals(notebookName, StringComparison.OrdinalIgnoreCase));

        if (notebook == null)
        {
            // TODO: throw proper exception
            throw new ArgumentException($"Unable to find notebook {notebookName}");
        }

        // Path possibilities:
        //
        // somepage - find the page(s) with this name
        // somesection/somepage - find the section(s) then page(s) with these names
        // somesectiongroup/somesection/somepage - find the section group(s), then sections(s) then page(s) with these names
        // section groups are nestable, does this matter?
        OnenoteSection? section = null;
        SectionGroup? sectionGroup = null;

        string sectionName = pathParts[pathParts.Length - 2];
        string pageName = pathParts[pathParts.Length - 1];

        if (pathParts.Length > 2)
        {
            int numNestedSectionGroups = pathParts.Length - 3;

            INotebookSectionGroupsCollectionPage noteBookSectionGroups = await this._graphServiceClient.Me.Onenote.Notebooks[notebook.Id].SectionGroups.Request().GetAsync(cancellationToken).ConfigureAwait(false);
            string notebookSectionGroupName = pathParts[0];
            sectionGroup = noteBookSectionGroups.FirstOrDefault(x => x.DisplayName.Equals(notebookSectionGroupName, StringComparison.OrdinalIgnoreCase));

            if (sectionGroup == null)
            {
                // TODO: throw proper exception
                throw new ArgumentException($"Unable to find section group {notebookSectionGroupName} for notebook {notebookName} with path {path}");
            }

            for (int i = 0; i < numNestedSectionGroups; i++)
            {
                ISectionGroupSectionGroupsCollectionPage sectionGroupsectionGroups = await this._graphServiceClient.Me.Onenote.SectionGroups[sectionGroup.Id].SectionGroups.Request().GetAsync(cancellationToken).ConfigureAwait(false);
                string sectionGroupName = pathParts[i + 1];

                sectionGroup = sectionGroupsectionGroups.FirstOrDefault(x => x.DisplayName.Equals(sectionGroupName, StringComparison.OrdinalIgnoreCase));

                if (sectionGroup == null)
                {
                    // TODO: throw proper exception
                    throw new ArgumentException($"Unable to find section group {sectionGroupName} for notebook {notebookName} with path {path}");
                }
            }

            ISectionGroupSectionsCollectionPage sectionGroupSections = await this._graphServiceClient.Me.Onenote.SectionGroups[sectionGroup.Id].Sections.Request().GetAsync(cancellationToken).ConfigureAwait(false);
            section = sectionGroupSections.FirstOrDefault(x => x.DisplayName.Equals(sectionName, StringComparison.OrdinalIgnoreCase));
        }
        else
        {
            INotebookSectionsCollectionPage sections = await this._graphServiceClient.Me.Onenote.Notebooks[notebook.Id].Sections.Request().GetAsync(cancellationToken).ConfigureAwait(false);
            section = sections.FirstOrDefault(x => x.DisplayName.Equals(sectionName, StringComparison.OrdinalIgnoreCase));
        }

        if (section == null)
        {
            // TODO: throw proper exception
            throw new ArgumentException($"Unable to find section {sectionName} for notebook {notebookName} with path {path}");
        }

        IOnenoteSectionPagesCollectionPage pages = await this._graphServiceClient.Me.Onenote.Sections[section.Id].Pages.Request().GetAsync(cancellationToken).ConfigureAwait(false);

        OnenotePage page = pages.FirstOrDefault(x => x.Title.Equals(pageName, StringComparison.OrdinalIgnoreCase));

        if (page == null)
        {
            // TODO: throw proper exception
            throw new ArgumentException($"Unable to find page {pageName} for notebook {notebookName} with path {path}");
        }

        return page.Content;
    }
}
