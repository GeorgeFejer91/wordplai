using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using UnderTree.WordAddin;

internal static partial class NativeSuggestionWorkflowSmoke
{
    private static void NativeUncertainOutcome(string scenario)
    {
        bool suggestion = scenario.StartsWith("suggestion", StringComparison.Ordinal);
        bool committed = scenario == "suggestion-continuation";
        NewFixture("uncertain-" + scenario, Original);
        Mark(Original, "snappier");
        string address = Text(Local("connection_status"), "localMcpUrl");
        string conversation = "uncertain-native-" + scenario;
        var client = new NativeMcpClient(address, "uncertain-native-probe", conversation);
        var context = NativeTool(client, "read_manuscript", Values("manuscript_id", documentId));
        const string requestId = "one-native-outcome";
        int markingFrom = Text(context, "text").IndexOf(Outside, StringComparison.Ordinal);
        object request = suggestion ? (object)ContextRequest(requestId, context, Original, Revised) :
            Values("documentId", documentId, "marking", Values("requestId", requestId, "contextId", context["context_id"],
                "revision", context["revision"], "color", "#DACBEF", "ranges", new object[] { Values("from", markingFrom, "to", markingFrom + Outside.Length, "quote", Outside) }));
        string tool = suggestion ? "suggest_text" : "highlight_text";
        string before = new WordNativeEditState().NativeIdentity(document);
        Action failingFinalizer = delegate { throw new IOException("Injected finalizer failure after unconfirmed native effect."); };
        client = new NativeMcpClient(address, "uncertain-native-probe", conversation);
        if (scenario == "suggestion-rollback") editor.ReviewChanged += failingFinalizer;
        Dictionary<string, object> failed;
        try { failed = client.Call(tool, request); }
        finally { editor.ReviewChanged -= failingFinalizer; }
        Check((bool)failed["isError"] && failed.ContainsKey("structuredContent") && Text(NativeCoordination(failed), "state") == "outcome_unknown",
            "An unconfirmed or undelivered native effect remains outcome_unknown through rollback, persistence and finalizer failures.");
        var error = Map(Map(failed["structuredContent"])["error"]);
        var details = Map(error["details"]);
        Check(Text(error, "code") == "outcome_unknown" && (bool)details["operation_started"] && (bool)details["outcome_unknown"],
            "The native failure explicitly reports a started operation whose result needs inspection.");
        WordDocumentReview state = NativeDocumentState();
        string nativeId = OwnedNativeRequestId(client, requestId);
        SavedWordReview durable = WordReviewStore.Load((object)document);
        if (suggestion)
        {
            string outcome = committed ? "applied" : "outcome_uncertain";
            Check(state.Policy.Proposals.Single(receipt => receipt.Id == nativeId).Status == outcome &&
                durable.Policy.Proposals.Single(receipt => receipt.Id == nativeId).Status == outcome &&
                state.Policy.IsShared == committed && new WordNativeEditState().NativeIdentity(document) != before,
                "The actual native edit retains its consumed durable outcome; only unconfirmed effects revoke sharing.");
        }
        else
        {
            Check(state.Policy.HighlightReceipts.Single(receipt => receipt.Request.RequestId == nativeId).Outcome == "outcome_uncertain" &&
                durable.Policy.HighlightReceipts.Single(receipt => receipt.Request.RequestId == nativeId).Outcome ==
                    (scenario == "marking-persistence" ? "reserved" : "outcome_uncertain") && !state.Policy.IsShared &&
                new WordNativeEditState().NativeIdentity(document) != before,
                "Native marking stays consumed and revoked even when the uncertain-receipt write fails after its durable reservation.");
        }
        string after = new WordNativeEditState().NativeIdentity(document);
        int comments = (int)document.Comments.Count;
        client = new NativeMcpClient(address, "uncertain-native-probe", conversation);
        var retry = client.Call(tool, request);
        Check((bool)retry["isError"] != committed && new WordNativeEditState().NativeIdentity(document) == after && (int)document.Comments.Count == comments &&
            (!committed || Text(NativeCoordination(retry), "state") == "replayed"),
            "The exact identity retry cannot duplicate a native effect; committed delivery recovery retrieves its retained receipt.");
        if (!committed)
        {
            Check(retry.ContainsKey("structuredContent") && Text(NativeCoordination(retry), "state") == "outcome_unknown" &&
                Text(Map(Map(retry["structuredContent"])["error"]), "code") == "outcome_unknown",
                "An exact retry of an uncertain native receipt retains outcome_unknown instead of becoming a definite rejection or success.");
            foreach (string field in suggestion ? new[] { "context_id", "replacement" } : new[] { "contextId", "color" })
            {
                var changed = Map(request);
                if (suggestion) changed[field] = field == "context_id" ? "changed-original-context" : "Different retried wording.";
                else { var marking = Map(changed["marking"]); marking[field] = field == "contextId" ? "changed-original-context" : "#112233"; changed["marking"] = marking; }
                var denied = client.Call(tool, changed);
                Check((bool)denied["isError"] && Text(NativeCoordination(denied), "state") == "rejected" &&
                    new WordNativeEditState().NativeIdentity(document) == after && (int)document.Comments.Count == comments,
                    "A changed original context or changed arguments cannot inherit the uncertain identity's replay result or create another native effect.");
            }
        }
    }
}
