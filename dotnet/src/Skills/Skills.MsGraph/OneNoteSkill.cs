// Copyright (c) Microsoft. All rights reserved.

using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.SemanticKernel.Orchestration;
using Microsoft.SemanticKernel.SkillDefinition;

namespace Microsoft.SemanticKernel.Skills.MsGraph;

//**********************************************************************************************************************
// EXAMPLE USAGE
// Option #1: as a standalone C# function
//
// OneNoteSkill oneNoteSkill = new(new OneNoteConnector());
// string filePath = "PATH_TO_DOCX_FILE.docx";
// string text = await oneNoteSkill.ReadTextAsync(filePath);
// Console.WriteLine(text);
//
//
// Option #2: with the Semantic Kernel
//
// DocumentSkill documentSkill = new(new WordDocumentConnector(), new LocalDriveConnector());
// string filePath = "PATH_TO_DOCX_FILE.docx";
// ISemanticKernel kernel = SemanticKernel.Build();
// var result = await kernel.RunAsync(
//      filePath,
//      documentSkill.ReadTextAsync);
// Console.WriteLine(result);
//**********************************************************************************************************************

/// <summary>
/// Skill for interacting with OneNote
/// </summary>
public class OneNoteSkill
{
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
        /// Path to content.
        /// </summary>
        public const string Path = "path";
    }

    private readonly INoteConnector _noteConnector;
    private readonly ILogger<OneNoteSkill> _logger;

    /// <summary>
    /// Initializes a new instance of the <see cref="OneNoteSkill"/> class.
    /// </summary>
    /// <param name="noteConnector">Document connector</param>
    /// <param name="logger">Optional logger</param>
    public OneNoteSkill(INoteConnector noteConnector, ILogger<OneNoteSkill>? logger = null)
    {
        this._noteConnector = noteConnector ?? throw new ArgumentNullException(nameof(noteConnector));
        this._logger = logger ?? new NullLogger<OneNoteSkill>();
    }

    /// <summary>
    /// Read all text from a OneNote page, using <see cref="ContextVariables.Input"/> as the name of the notebook
    /// </summary>
    [SKFunction("Read text from a OneNote page")]
    [SKFunctionInput(Description = "Name of the OneNote to read")]
    [SKFunctionContextParameter(Name = Parameters.Path, Description = "Path to page")]
    public async Task<string> GetPageContentAsync(string name, SKContext context)
    {
        this._logger.LogInformation("Reading text from {0} OneNote", name);
        if (!context.Variables.Get(Parameters.Path, out string path))
        {
            context.Fail($"Missing variable {Parameters.Path}.");
            return string.Empty;
        }
        
        Stream s = await this._noteConnector.GetPageContentStreamAsync(name, path, context.CancellationToken).ConfigureAwait(false);

        using var reader = new StreamReader(s);
        return await reader.ReadToEndAsync().ConfigureAwait(false);
    }

    /// <summary>
    /// Read all text from all pages in a OneNote section, using <see cref="ContextVariables.Input"/> as the name of the notebook
    /// </summary>
    [SKFunction("Read text from all pages in a OneNote section")]
    [SKFunctionInput(Description = "Name of the OneNote to read")]
    [SKFunctionContextParameter(Name = Parameters.Path, Description = "Path to section")]
    public async Task<string> GetSectionContentAsync(string name, SKContext context)
    {
        this._logger.LogInformation("Reading text from {0} OneNote", name);
        if (!context.Variables.Get(Parameters.Path, out string path))
        {
            context.Fail($"Missing variable {Parameters.Path}.");
            return string.Empty;
        }

        Stream s = await this._noteConnector.GetSectionContentStreamAsync(name, path, context.CancellationToken).ConfigureAwait(false);

        using var reader = new StreamReader(s);
        return await reader.ReadToEndAsync().ConfigureAwait(false);
    }
}
