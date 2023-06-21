// Copyright (c) Microsoft. All rights reserved.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Graph;
using Microsoft.SemanticKernel.Skills.MsGraph.Connectors.Diagnostics;
using Microsoft.SemanticKernel.Skills.MsGraph.Connectors.Utilities;

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
        OnenotePage page = await this.GetNotebookPageAsync(notebookName, path, cancellationToken).ConfigureAwait(false);
        return await this.GetPageStreamAsync(page, cancellationToken).ConfigureAwait(false);
    }

    /// <inheritdoc/>
    public async Task<Stream> GetSectionContentStreamAsync(string notebookName, string path, CancellationToken cancellationToken = default)
    {
        Ensure.NotNullOrWhitespace(notebookName, nameof(notebookName));
        Ensure.NotNullOrWhitespace(path, nameof(path));
        OnenoteSection section = await this.GetNotebookSectionAsync(notebookName, path, cancellationToken).ConfigureAwait(false);
        IEnumerable<OnenotePage> pages = await this._graphServiceClient.Me.Onenote.Sections[section.Id].Pages.Request().GetAsync(cancellationToken).ConfigureAwait(false);
        IEnumerable<Stream> streams = await this.GetPageStreamsAsync(pages, cancellationToken).ConfigureAwait(false);
        return new MultiStream(streams);
    }

    /// <inheritdoc/>
    public async Task<string> CreatePageShareLinkAsync(string notebookName, string path, string type = "view", string scope = "anonymous", CancellationToken cancellationToken = default)
    {
        Ensure.NotNullOrWhitespace(notebookName, nameof(notebookName));
        Ensure.NotNullOrWhitespace(path, nameof(path));

        OnenotePage page = await this.GetNotebookPageAsync(notebookName, path, cancellationToken).ConfigureAwait(false);

        // TODO: Honour type and scope
        return page.Links.OneNoteWebUrl.Href;
    }

    /// <inheritdoc/>
    public async Task<string> CreateSectionShareLinkAsync(string notebookName, string path, string type = "view", string scope = "anonymous", CancellationToken cancellationToken = default)
    {
        Ensure.NotNullOrWhitespace(notebookName, nameof(notebookName));
        Ensure.NotNullOrWhitespace(path, nameof(path));

        OnenoteSection section = await GetNotebookSectionAsync(notebookName, path, cancellationToken).ConfigureAwait(false);

        // TODO: Honour type and scope
        return section.Links.OneNoteWebUrl.Href;
    }

    private async Task<IEnumerable<Stream>> GetPageStreamsAsync(IEnumerable<OnenotePage> pages, CancellationToken cancellationToken)
    {
        IList<Stream> streams = new List<Stream>();

        foreach (OnenotePage page in pages)
        {
            Stream s = await GetPageStreamAsync(page, cancellationToken).ConfigureAwait(false);
            streams.Add(s);
        }

        return streams;
    }

    private Task<Stream> GetPageStreamAsync(OnenotePage page, CancellationToken cancellationToken)
    {
        return this._graphServiceClient.Me.Onenote.Pages[page.Id].Content.Request().GetAsync(cancellationToken);
    }

    private async Task<Notebook> GetNotebookAsync(string notebookName, CancellationToken cancellationToken)
    {
        IOnenoteNotebooksCollectionPage notebooks = await this._graphServiceClient.Me.Onenote.Notebooks.Request().GetAsync(cancellationToken).ConfigureAwait(false);
        Notebook notebook = notebooks.FirstOrDefault(x => x.DisplayName.Equals(notebookName, StringComparison.OrdinalIgnoreCase));

        if (notebook == null)
        {
            // TODO: throw proper exception
            throw new ArgumentException($"Unable to find notebook {notebookName}");
        }

        return notebook;
    }

    private async Task<OnenotePage> GetNotebookPageAsync(string notebookName, string path, CancellationToken cancellationToken)
    {
        VerifyPagePath(path);

        Notebook notebook = await this.GetNotebookAsync(notebookName, cancellationToken).ConfigureAwait(false);

        string sectionPath = GetSectionPath(path);
        OnenoteSection section = await this.GetSectionAsync(notebook.Id, sectionPath, cancellationToken).ConfigureAwait(false);

        string pageName = GetPageName(path);
        return await this.GetSectionPageAsync(section.Id, pageName, cancellationToken).ConfigureAwait(false);
    }

    private async Task<OnenotePage> GetSectionPageAsync(string sectionId, string pageName, CancellationToken cancellationToken)
    {
        IOnenoteSectionPagesCollectionPage pages = await this._graphServiceClient.Me.Onenote.Sections[sectionId].Pages.Request().GetAsync(cancellationToken).ConfigureAwait(false);
        OnenotePage page = pages.FirstOrDefault(x => x.Title.Equals(pageName, StringComparison.OrdinalIgnoreCase));

        if (page == null)
        {
            // TODO: throw proper exception
            throw new ArgumentException($"Unable to find page {pageName}");
        }

        return page;
    }

    private async Task<OnenoteSection> GetNotebookSectionAsync(string notebookName, string path, CancellationToken cancellationToken)
    {
        VerifySectionPath(path);

        Notebook notebook = await this.GetNotebookAsync(notebookName, cancellationToken).ConfigureAwait(false);
        return await this.GetSectionAsync(notebook.Id, path, cancellationToken).ConfigureAwait(false);
    }

    private async Task<OnenoteSection> GetSectionAsync(string notebookId, string path, CancellationToken cancellationToken)
    {
        string[] pathParts = path.Split('/');
        string sectionName = GetSectionName(path);
        OnenoteSection? section = null;

        if (pathParts.Length > 1)
        {
            int numNestedSectionGroups = pathParts.Length - 2;
            string notebookSectionGroupName = pathParts[0];

            INotebookSectionGroupsCollectionPage noteBookSectionGroups = await this._graphServiceClient.Me.Onenote.Notebooks[notebookId].SectionGroups.Request().GetAsync(cancellationToken).ConfigureAwait(false);
            SectionGroup? sectionGroup = noteBookSectionGroups.FirstOrDefault(x => x.DisplayName.Equals(notebookSectionGroupName, StringComparison.OrdinalIgnoreCase));

            if (sectionGroup == null)
            {
                // TODO: throw proper exception
                throw new ArgumentException($"Unable to find section group {notebookSectionGroupName} with path {path}");
            }

            for (int i = 0; i < numNestedSectionGroups; i++)
            {
                string sectionGroupName = pathParts[i + 1];

                ISectionGroupSectionGroupsCollectionPage sectionGroupsectionGroups = await this._graphServiceClient.Me.Onenote.SectionGroups[sectionGroup.Id].SectionGroups.Request().GetAsync(cancellationToken).ConfigureAwait(false);
                sectionGroup = sectionGroupsectionGroups.FirstOrDefault(x => x.DisplayName.Equals(sectionGroupName, StringComparison.OrdinalIgnoreCase));

                if (sectionGroup == null)
                {
                    // TODO: throw proper exception
                    throw new ArgumentException($"Unable to find section group {sectionGroupName} with path {path}");
                }
            }

            ISectionGroupSectionsCollectionPage sectionGroupSections = await this._graphServiceClient.Me.Onenote.SectionGroups[sectionGroup.Id].Sections.Request().GetAsync(cancellationToken).ConfigureAwait(false);
            section = sectionGroupSections.FirstOrDefault(x => x.DisplayName.Equals(sectionName, StringComparison.OrdinalIgnoreCase));
        }
        else
        {
            INotebookSectionsCollectionPage sections = await this._graphServiceClient.Me.Onenote.Notebooks[notebookId].Sections.Request().GetAsync(cancellationToken).ConfigureAwait(false);
            section = sections.FirstOrDefault(x => x.DisplayName.Equals(sectionName, StringComparison.OrdinalIgnoreCase));
        }

        if (section == null)
        {
            // TODO: throw proper exception
            throw new ArgumentException($"Unable to find section {sectionName} with path {path}");
        }

        return section;
    }

    private static string GetPageName(string path)
    {
        return path.Substring(path.LastIndexOf('/') + 1);
    }

    private static string GetSectionName(string path)
    {
        return path.Substring(path.LastIndexOf('/') + 1);
    }

    private static string GetSectionPath(string path)
    {
        // Trim page from path to get path to section
        return path.Substring(0, path.LastIndexOf('/'));
    }

    private static string[] VerifyPagePath(string path)
    {
        return VerifyPath(path, 2);
    }

    private static string[] VerifySectionPath(string path)
    {
        return VerifyPath(path, 1);
    }

    private static string[] VerifyPath(string path, int minExpectedParts)
    {
        string[] pathParts = path.Split('/');

        if (pathParts.Length < minExpectedParts)
        {
            // TODO: throw proper exception
            throw new ArgumentException($"Path should have {minExpectedParts} or more parts, was {pathParts.Length}");
        }

        return pathParts;
    }
}
