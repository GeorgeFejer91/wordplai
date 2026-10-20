using System;
using System.Collections.Generic;
using System.Linq;

namespace UnderTree.WordAddin
{
    internal sealed partial class WordEditor
    {
        internal static bool AutomationMutates(IDictionary<string, object> args)
        {
            object command;
            return args != null && args.TryGetValue("command", out command) && command is string &&
                new[] { "save_rule", "mark_range", "remove_mark", "accept", "reject", "revoke_context" }.Contains((string)command);
        }

        private object Automation(IDictionary<string, object> args)
        {
            string command = Text(args, "command");
            if (command == "documents") return new { documents = DocumentChoices(OpenDocumentReviews()) };
            if (command != "state" && command != "context_preview" && !AutomationMutates(args))
                throw new ArgumentException("Unsupported automation command.");
            string id = Text(args, "documentId");
            var matches = OpenDocumentReviews().Where(state => state.Id == id).ToArray();
            if (matches.Length == 0) throw new InvalidOperationException("The requested document is no longer open; no active-document fallback is allowed.");
            if (matches.Length != 1) throw new InvalidOperationException("Another open copy has this review identity. Close that copy before continuing.");
            automationTarget = DocumentReview(matches[0].Document);
            try
            {
                WordDocumentReview state = Current();
                if (AutomationMutates(args) && Text(args, "snapshot") != state.Snapshot.Revision)
                    throw new InvalidOperationException("The document changed. Read its current state or complete context before continuing.");
                if (command == "mark_range") return CommitLocalMark(state, args, Number(args, "from"), Number(args, "to"));
                return DispatchCore(command, args);
            }
            finally { automationTarget = null; }
        }
    }
}
