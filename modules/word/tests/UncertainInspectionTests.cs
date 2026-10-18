using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.Serialization;
using UnderTree.WordAddin;

internal static class UncertainInspectionTests
{
    private const string Document = "uncertain-inspection-fixture", Revision = "revision", Source = "Alpha beta.", Binding = "native-stories";
    private const string RequestId = "retained-uncertain";
    private static string contract, ownerScope, otherScope;
    private static WordDocumentReview state;
    private static WordEditor editor;
    private static int checks, failures, inspections, effects;
    internal static bool AccessRunning { get { return true; } }

    public static int Main(string[] args)
    {
        try
        {
            contract = File.ReadAllText(args[0]);
            var coordination = new McpCoordination();
            ownerScope = coordination.Resolve(coordination.Create(), null).NativeScope;
            otherScope = coordination.Resolve(coordination.Create(), null).NativeScope;
            foreach (bool highlighting in new[] { false, true }) Run(highlighting);
            Console.WriteLine("CHECKS " + checks + "; FAILURES " + failures + "; no native Word or production file writes.");
            return failures == 0 ? 0 : 1;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 2; }
    }

    internal static WordDocumentReview Resolve(IDictionary<string, object> args, string scope)
    {
        object id;
        if (!(args.TryGetValue("manuscript_id", out id) || args.TryGetValue("documentId", out id)) || !Object.Equals(id, Document))
            throw new Exception("Fixture document identity changed.");
        return state;
    }

    internal static WordDocumentReview Inspect(WordDocumentReview current)
    {
        inspections++;
        throw new McpOperationException("Synthetic mixed native mark exceeds its color verification budget.",
            "native_color_verification_limited", new { additionalProbes = 256, inspectionMilliseconds = 1000 });
    }

    private static void Run(bool highlighting)
    {
        string tool = highlighting ? "highlight_text" : "suggest_text";
        var policy = new PolicyEngine(contract);
        policy.AddMark(new PassageMark { Id = "Wrai_probe", RuleId = "snappier", From = 0, To = Source.Length, Quote = Source, Prompt = "" }, Source);
        policy.Share(Document);
        string token = policy.ReadContext(Document, Revision, Source, Binding, ownerScope);
        state = new WordDocumentReview { Id = Document, Policy = policy };
        editor = (WordEditor)FormatterServices.GetUninitializedObject(typeof(WordEditor));
        typeof(WordEditor).GetField("documents", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(editor,
            new Dictionary<long, WordDocumentReview> { { 1L, state } });
        string nativeId = (string)typeof(WordEditor).GetMethod("ClientRequestId", BindingFlags.Static | BindingFlags.NonPublic).Invoke(null,
            new object[] { ownerScope, RequestId });
        Dictionary<string, object> request;
        Action record;
        if (highlighting)
        {
            var marking = new HighlightRequest { DocumentId = Document, RequestId = nativeId, ContextId = token, Revision = Revision,
                ColorId = "snappier", Ranges = new List<HighlightRange> { new HighlightRange { From = 0, To = Source.Length, Quote = Source } } };
            policy.Highlight(marking, Source, Binding, ownerScope); policy.ReserveHighlight(nativeId);
            record = delegate { policy.FailHighlight(nativeId); };
            request = Values("documentId", Document, "marking", Values("requestId", RequestId, "contextId", token, "revision", Revision,
                "colorId", "snappier", "ranges", new object[] { Values("from", 0, "to", Source.Length, "quote", Source) }));
        }
        else
        {
            policy.Submit(new Proposal { Id = nativeId, DocumentId = Document, Revision = Revision, From = 0, To = Source.Length,
                Quote = Source, Replacement = "A clearer point.", Rationale = "Clarify." }, token, Source, Binding, ownerScope);
            policy.BeginApplication(nativeId);
            record = delegate { policy.RecordOutcomeUncertain(nativeId); };
            request = Values("manuscript_id", Document, "request_id", RequestId, "context_id", token, "revision", Revision,
                "from", 0, "to", Source.Length, "quote", Source, "replacement", "A clearer point.", "rationale", "Clarify.");
        }
        effects = 1;
        McpOperationException first = null;
        try { WordNativeOutcome.Unknown(record, policy.Revoke, delegate { }); }
        catch (McpOperationException error) { first = error; }
        Check(first != null && first.Code == "outcome_unknown" && !policy.IsShared && effects == 1,
            tool + ": initial consumed native uncertainty is retained and sharing is revoked.");
        string before = McpEndpoint.Json().Serialize(policy.ExportState());
        McpOperationException exact = Call(tool, request, ownerScope);
        Check(exact != null && exact.Code == "outcome_unknown" && inspections == 0 && Stable(before),
            tool + ": an exact same-scope uncertain retry remains unknown before blocked native color inspection.");
        foreach (bool contextOnly in new[] { true, false })
        {
            var changed = Clone(request);
            if (highlighting)
            {
                var marking = Map(changed["marking"]);
                marking[contextOnly ? "contextId" : "colorId"] = contextOnly ? "changed-original-context" : "protect";
                changed["marking"] = marking;
            }
            else changed[contextOnly ? "context_id" : "replacement"] = contextOnly ? "changed-original-context" : "Different wording.";
            McpOperationException denied = Call(tool, changed, ownerScope);
            Check(denied != null && denied.Code != "outcome_unknown" && inspections == 0 && Stable(before),
                tool + ": changed " + (contextOnly ? "context" : "arguments") + " refuse before inspection without inheriting unknown or adding an effect.");
        }
        McpOperationException foreign = Call(tool, request, otherScope);
        Check(foreign != null && foreign.Code == "native_color_verification_limited" && inspections == 1 && Stable(before),
            tool + ": another transport scope cannot retrieve the retained unknown receipt or bypass inspection.");
        var fresh = Clone(request);
        if (highlighting) { var marking = Map(fresh["marking"]); marking["requestId"] = "different-id"; fresh["marking"] = marking; }
        else fresh["request_id"] = "different-id";
        McpOperationException newIdentity = Call(tool, fresh, ownerScope);
        Check(newIdentity != null && newIdentity.Code == "native_color_verification_limited" && inspections == 1 && Stable(before),
            tool + ": a new identity still follows normal native inspection and receives no receipt authority.");
    }

    private static bool Stable(string before)
    { return !state.Policy.IsShared && effects == 1 && McpEndpoint.Json().Serialize(state.Policy.ExportState()) == before; }

    private static McpOperationException Call(string tool, Dictionary<string, object> args, string scope)
    {
        inspections = 0;
        var routed = Clone(args); routed[WordEditor.ReviewSessionKey] = scope;
        try { typeof(WordEditor).GetMethod("Review", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(editor, new object[] { tool, routed }); return null; }
        catch (TargetInvocationException error)
        {
            var refusal = error.InnerException as McpOperationException;
            if (refusal == null) throw;
            Console.WriteLine("RESULT " + tool + " " + (refusal.Code ?? "definite refusal") + " inspections=" + inspections);
            return refusal;
        }
    }

    private static Dictionary<string, object> Values(params object[] entries)
    { var result = new Dictionary<string, object>(); for (int index = 0; index < entries.Length; index += 2) result.Add((string)entries[index], entries[index + 1]); return result; }
    private static Dictionary<string, object> Clone(object value)
    { return McpEndpoint.Json().Deserialize<Dictionary<string, object>>(McpEndpoint.Json().Serialize(value)); }
    private static Dictionary<string, object> Map(object value) { return Clone(value); }
    private static void Check(bool condition, string message)
    { checks++; if (!condition) failures++; Console.WriteLine((condition ? "PASS " : "FAIL ") + message); }
}
