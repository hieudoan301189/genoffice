/** The DVH-Script v0 reference the agent reads (dvh_script_reference) and the editor's help shows. */
export const SCRIPT_REFERENCE = `# DVH-Script v0
One statement per line; ' starts a comment; keywords in any case; text in "double quotes" ("" for a quote).

Header (optional, before the statements):
  Workflow "Name"
  Id "wf_name"
  Description "What it does"
  Param records As List = Json([]) Description "the rows"     ' types: Text Number Boolean Record List Any

Statements:
  x = <expression>                      ' a variable for the whole run
  r = Group.Verb(arg:=value, other:=x)   ' call an action, keep its output
  Group.Verb(arg:=value) On "doc_id"    ' call an action in another open document
  Log <expression>
  If <condition> Then … ElseIf <condition> Then … Else … End If
  For Each row In records [Where row.qty > 0] … Next       ' INDEX is the 1-based position
  Try … Catch [err] … End Try           ' err.message, err.step
  Transaction … End Transaction         ' all or nothing

Values: "text", 12.5, True, False, Null, Json({"a": [1, 2]}) for lists and records.
Expressions (Excel-like): + - * / % ^, & joins text, = <> < <= > >=, AND OR NOT,
  record.field, list.Count, [Column title] of the current record,
  IF LEN UPPER LOWER TRIM LEFT RIGHT MID CONTAINS ISBLANK ROUND ABS MIN MAX SUM CONCAT TEXT PAD.
Scripts run in a sandbox: only the listed actions reach the document; nothing else (files, network) is reachable.
A run is one transaction: if a step fails, every change of the run is undone.`
