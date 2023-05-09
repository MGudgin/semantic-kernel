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
// DocumentSkill documentSkill = new(new WordDocumentConnector(), new LocalDriveConnector());
// string filePath = "PATH_TO_DOCX_FILE.docx";
// string text = await documentSkill.ReadTextAsync(filePath);
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
/// Skill for interacting with OneNotes
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

    private readonly IDocumentConnector _documentConnector;
    private readonly IFileSystemConnector _fileSystemConnector;
    private readonly ILogger<OneNoteSkill> _logger;

    /// <summary>
    /// Initializes a new instance of the <see cref="OneNoteSkill"/> class.
    /// </summary>
    /// <param name="documentConnector">Document connector</param>
    /// <param name="fileSystemConnector">File system connector</param>
    /// <param name="logger">Optional logger</param>
    public OneNoteSkill(IDocumentConnector documentConnector, IFileSystemConnector fileSystemConnector, ILogger<OneNoteSkill>? logger = null)
    {
        this._documentConnector = documentConnector ?? throw new ArgumentNullException(nameof(documentConnector));
        this._fileSystemConnector = fileSystemConnector ?? throw new ArgumentNullException(nameof(fileSystemConnector));
        this._logger = logger ?? new NullLogger<DocumentSkill>();
    }

    /// <summary>
    /// Read all text from a OneNote, using <see cref="ContextVariables.Input"/> as the name.
    /// </summary>
    [SKFunction("Read text from a OneNote")]
    [SKFunctionInput(Description = "Name of the OneNote to read")]
    [SKFunctionContextParameter(Name = Parameters.Path, Description = "Path to section group, section or page")]
    public async Task<string> ReadTextAsync(string name, SKContext context)
    {
        this._logger.LogInformation("Reading text from {0} OneNote", name);
        using var stream = await this._fileSystemConnector.GetFileContentStreamAsync(filePath, context.CancellationToken).ConfigureAwait(false);
        return this._documentConnector.ReadText(stream);
    }

    /// <summary>
    /// Add the text in <see cref="ContextVariables.Input"/> to a OneNote page. If the page doesn't exist, it will be created.
    /// </summary>
    [SKFunction("Add text to a OneNote page. If the page doesn't exist, it will be created.")]
    [SKFunctionInput(Description = "Text to add")]
    [SKFunctionContextParameter(Name = Parameters.Path, Description = "Path to destination page")]
    public async Task AppendTextAsync(string text, SKContext context)
    {
        if (!context.Variables.Get(Parameters.Path, out string path))
        {
            context.Fail($"Missing variable {Parameters.Path}.");
            return;
        }

        // If the page already exists, open it. If not, create it.
        if (await this._fileSystemConnector.FileExistsAsync(filePath).ConfigureAwait(false))
        {
            this._logger.LogInformation("Writing text to file {0}", filePath);
            using Stream stream = await this._fileSystemConnector.GetWriteableFileStreamAsync(filePath, context.CancellationToken).ConfigureAwait(false);
            this._documentConnector.AppendText(stream, text);
        }
        else
        {
            this._logger.LogInformation("File does not exist. Creating file at {0}", filePath);
            using Stream stream = await this._fileSystemConnector.CreateFileAsync(filePath).ConfigureAwait(false);
            this._documentConnector.Initialize(stream);

            this._logger.LogInformation("Writing text to {0}", filePath);
            this._documentConnector.AppendText(stream, text);
        }
    }
}
