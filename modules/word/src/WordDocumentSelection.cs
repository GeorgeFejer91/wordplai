using System;
using System.Collections.Generic;
using System.Linq;

namespace UnderTree.WordAddin
{
    internal sealed partial class WordEditor
    {
        internal const string ReviewSessionKey = WordDocumentPicker.SessionKey;
        private readonly Dictionary<string, WordDocumentReview> reviewSelections = new Dictionary<string, WordDocumentReview>();

        private WordDocumentReview[] OpenDocumentReviews()
        {
            object[] native = ((System.Collections.IEnumerable)application.Documents).Cast<object>().ToArray();
            if (native.Length > 32)
                throw new InvalidOperationException("Close some Word documents before reviewing; at most 32 can be listed at once.");
            var identities = new HashSet<long>(native.Select(WordSnapshot.Identity));
            foreach (long closed in documents.Keys.Where(key => !identities.Contains(key)).ToArray())
            {
                if (documents[closed].Policy != null) documents[closed].Policy.Revoke();
                documents.Remove(closed);
            }
            var result = new List<WordDocumentReview>();
            foreach (object document in native)
            {
                long identity = WordSnapshot.Identity(document);
                WordDocumentReview state;
                if (!documents.TryGetValue(identity, out state))
                {
                    state = new WordDocumentReview { Document = document,
                        Id = WordReviewStore.DocumentIdentity(document) ?? Guid.NewGuid().ToString("N") };
                    documents.Add(identity, state);
                }
                result.Add(state);
            }
            return result.ToArray();
        }

        private bool IsOpen(WordDocumentReview state)
        {
            return state != null && documents.Values.Contains(state) &&
                ((System.Collections.IEnumerable)application.Documents).Cast<object>().Any(document =>
                    WordSnapshot.Identity(document) == WordSnapshot.Identity((object)state.Document));
        }

        private object DocumentChoices(WordDocumentReview[] open)
        {
            long active = ActiveDocumentIdentity();
            return open.Select(state => new { manuscript_id = state.Id, documentId = state.Id,
                title = (string)state.Document.Name, active = WordSnapshot.Identity((object)state.Document) == active }).ToArray();
        }

        private long ActiveDocumentIdentity()
        {
            if ((int)application.Documents.Count == 0) return 0;
            try { return WordSnapshot.Identity((object)application.ActiveDocument); }
            catch (System.Runtime.InteropServices.COMException error)
            {
                if (error.ErrorCode != unchecked((int)0x800A1098)) throw;
                return 0;
            }
        }

        private WordDocumentReview ReviewDocument(string tool, IDictionary<string, object> args, string session)
        {
            WordDocumentReview selected;
            reviewSelections.TryGetValue(session, out selected);
            string key = tool == "highlight_text" ? "documentId" : "manuscript_id";
            string requested = args.ContainsKey(key) ? Text(args, key) : null;
            if (requested == null && tool != "read_manuscript") throw new ArgumentException("Missing " + key + ".");
            if (requested == null && selected != null)
            {
                if (!IsOpen(selected)) throw ClosedReviewDocument(selected.Id);
                return selected;
            }
            if (selected != null && selected.Id == requested)
            {
                if (IsOpen(selected)) return selected;
                if (tool != "read_manuscript" || Number(args, "offset", 0) != 0) throw ClosedReviewDocument(selected.Id);
            }
            if (tool != "read_manuscript")
                throw new McpOperationException("Read the chosen Word document completely before editing it.", "document_selection_required",
                    new { request_applied = false, instruction = "Keep the chosen manuscript_id. A Word window change never changes the writing target." });
            WordDocumentReview[] open = OpenDocumentReviews();
            if (requested == null)
            {
                if (open.Length == 0) throw new McpOperationException("Open a Word document first.", "no_document_open");
                if (open.Length != 1) return null;
                return open[0];
            }
            WordDocumentReview found = open.SingleOrDefault(state => state.Id == requested);
            if (found == null) throw ClosedReviewDocument(requested);
            if (Number(args, "offset", 0) != 0)
                throw new McpOperationException("Read the newly chosen document from offset zero.", "document_selection_required");
            return found;
        }

        private static McpOperationException ClosedReviewDocument(string id)
        {
            return new McpOperationException("The chosen Word document is no longer open. Reopen it or choose another document; no edit was made.",
                "document_closed", new { manuscript_id = id, request_applied = false,
                    instruction = "Do not switch to the active window automatically. Ask for another document only if the user has not already named one." });
        }

        internal IEnumerable<object> BackgroundReviewDocuments()
        {
            if ((int)application.Documents.Count == 0) return new object[0];
            long active = ActiveDocumentIdentity();
            var open = new HashSet<long>(((System.Collections.IEnumerable)application.Documents).Cast<object>().Select(WordSnapshot.Identity));
            return documents.Where(entry => entry.Value.AutomaticReview && open.Contains(entry.Key) &&
                entry.Key != active).Select(entry => (object)entry.Value.Document).ToArray();
        }

        internal void RefreshReviewDocument(object target)
        {
            WordDocumentReview state;
            if (target == null || operationActive || !documents.TryGetValue(WordSnapshot.Identity(target), out state) ||
                !IsOpen(state) || !state.AutomaticReview) return;
            operationActive = true;
            try { Current(DocumentReview(target)); }
            finally { operationActive = false; ClearSnapshots(); }
        }
    }
}
