using System.Text.Json.Serialization;

namespace McpOffice.Models;

/// <summary>
/// A block of document body content. Serialized polymorphically with an explicit `type`
/// discriminator: without it, System.Text.Json writes each element by its declared type —
/// this abstract base, which has no properties — and every block reaches the caller as `{}`.
/// <para>
/// `type` rather than the default `$type` because the caller is an agent, and the values are
/// a closed vocabulary it can branch on: `heading` | `paragraph`.
/// </para>
/// </summary>
[JsonPolymorphic(TypeDiscriminatorPropertyName = "type")]
[JsonDerivedType(typeof(HeadingBlock), "heading")]
[JsonDerivedType(typeof(ParagraphBlock), "paragraph")]
public abstract record Block;

public sealed record HeadingBlock(int Level, string Text) : Block;

public sealed record ParagraphBlock(IReadOnlyList<Run> Runs) : Block;
