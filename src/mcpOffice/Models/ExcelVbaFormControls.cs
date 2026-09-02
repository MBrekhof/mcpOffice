namespace McpOffice.Models;

/// <summary>Result of <c>excel_list_vba_form_controls</c>: the UI spec of each UserForm, inferred from its code-behind.</summary>
public sealed record ExcelVbaFormControlsResult(
    string Path,
    bool HasVbaProject,
    ExcelVbaFormControlsSummary Summary,
    IReadOnlyList<ExcelVbaUserForm> Forms);

public sealed record ExcelVbaFormControlsSummary(int FormCount, int ControlCount, int TypedControlCount);

/// <summary><c>FormEvents</c> are the form's own handlers (UserForm_Initialize, …).</summary>
public sealed record ExcelVbaUserForm(
    string Name,
    IReadOnlyList<ExcelVbaFormControl> Controls,
    IReadOnlyList<string> FormEvents,
    int HandlerCount);

/// <summary>
/// <c>InferredType</c> is an MSForms type name (TextBox, CommandButton, …) or <c>Control</c> when
/// nothing narrowed it; <c>TypeConfidence</c> says what did: declared | prefix | event | member | none.
/// </summary>
public sealed record ExcelVbaFormControl(
    string Name,
    string InferredType,
    string TypeConfidence,
    IReadOnlyList<string> Events,
    IReadOnlyList<string> ReferencedProperties);
