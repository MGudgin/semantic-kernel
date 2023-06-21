// Copyright (c) Microsoft. All rights reserved.

using System;
using System.ComponentModel;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.SemanticKernel.Orchestration;
using Microsoft.SemanticKernel.SkillDefinition;

namespace Microsoft.SemanticKernel.Skills.MsGraph;

/// <summary>
/// Skill for interacting with notes (e.g. OneNote)
/// </summary>
public class NoteSkill
{
    private const string DefaultLinkType = "view"; // TODO expose this as an SK variable
    private const string DefaultLinkScope = "anonymous"; // TODO expose this as an SK variable

    /// <summary>
    /// <see cref="ContextVariables"/> parameter names.
    /// </summary>
    public static class Parameters
    {
        /// <summary>
        /// Name of OneNote.
        /// </summary>
        public const string Name = "name";

        /// <summary>
        /// Path to page or section.
        /// </summary>
        public const string Path = "path";

        /// <summary>
        /// Type of link to create
        /// </summary>
        public const string LinkType = "linkType";

        /// <summary>
        /// Scope of link to create
        /// </summary>
        public const string LinkScope = "linkScope";
    }

    private readonly INoteConnector _noteConnector;
    private readonly ILogger<NoteSkill> _logger;

    /// <summary>
    /// Initializes a new instance of the <see cref="NoteSkill"/> class.
    /// </summary>
    /// <param name="noteConnector">Document connector</param>
    /// <param name="logger">Optional logger</param>
    public NoteSkill(INoteConnector noteConnector, ILogger<NoteSkill>? logger = null)
    {
        this._noteConnector = noteConnector ?? throw new ArgumentNullException(nameof(noteConnector));
        this._logger = logger ?? new NullLogger<NoteSkill>();
    }

    /// <summary>
    /// Read text from a page in a notebook, using <see cref="ContextVariables.Input"/> as the name of the notebook
    /// </summary>
    [SKFunction, Description("Read text from a page in a notebook.")]
    public async Task<string> GetPageContentAsync(
        [Description("Notebook name"), SKName("input")] string name,
        [Description("Path to page")] string path,
        CancellationToken cancellationToken = default)
    {
        this._logger.LogInformation("Reading text from {0} OneNote", name);

        Stream s = await this._noteConnector.GetPageContentStreamAsync(name, path, cancellationToken).ConfigureAwait(false);

        using var reader = new StreamReader(s);
        return await reader.ReadToEndAsync().ConfigureAwait(false);
    }

    /// <summary>
    /// Read all text from all pages in a section of a notebook, using <see cref="ContextVariables.Input"/> as the name of the notebook
    /// </summary>
    [SKFunction, Description("Read text from all pages in a section of a notebook.")]
    public async Task<string> GetSectionContentAsync(
        [Description("Notebook name"), SKName("input")] string name,
        [Description("Path to section")] string path,
        CancellationToken cancellationToken = default)
    {
        this._logger.LogInformation("Reading text from {0} OneNote", name);

        Stream s = await this._noteConnector.GetSectionContentStreamAsync(name, path, cancellationToken).ConfigureAwait(false);

        using var reader = new StreamReader(s);
        return await reader.ReadToEndAsync().ConfigureAwait(false);
    }

    /// <summary>
    /// Create a sharable link to a page in a notebook, using <see cref="ContextVariables.Input"/> as the name of the notebook
    /// </summary>
    [SKFunction, Description("Create a sharable link to a page in a notebook.")]
    public async Task<string> CreatePageLinkAsync(
        [Description("Notebook name"), SKName("input")] string name,
        [Description("Path to page")] string path,
        CancellationToken cancellationToken = default)
    {
        this._logger.LogDebug("Creating link for page at '{0}' in notebook '{1}'", path, name);

        return await this._noteConnector.CreatePageShareLinkAsync(name, path, DefaultLinkType, DefaultLinkScope, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Create a sharable link to a page in a OneNote
    /// </summary>
    [SKFunction, Description("Create a sharable link to a section in a notebook.")]
    public async Task<string> CreateSectionLinkAsync(
        [Description("Notebook name"), SKName("input")] string name,
        [Description("Path to section")] string path,
        CancellationToken cancellationToken = default)
    {
        this._logger.LogDebug("Creating link for section at '{0}' in notebook '{1}'", path, name);

        return await this._noteConnector.CreateSectionShareLinkAsync(name, path, DefaultLinkType, DefaultLinkScope, cancellationToken).ConfigureAwait(false);
    }
}
