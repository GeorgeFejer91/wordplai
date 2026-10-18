using System;
using System.Linq;

namespace UnderTree.WordAddin
{
    internal sealed class WordClientRead
    {
        internal PolicyEngine Policy;
        internal string Revision;
        internal long Generation;
        internal int Offset;
    }

    internal sealed partial class WordEditor
    {
        private void RevokeClient(string prefix)
        {
            if (operationActive) throw new McpOperationException("Word is finishing a native review operation before clearing retired client context.", "word_busy");
            foreach (string scope in reviewSelections.Keys.Where(value => value.StartsWith(prefix, StringComparison.Ordinal)).ToArray())
                reviewSelections.Remove(scope);
            foreach (WordDocumentReview state in documents.Values)
            {
                foreach (string scope in state.ClientReads.Keys.Where(value => value.StartsWith(prefix, StringComparison.Ordinal)).ToArray())
                    state.ClientReads.Remove(scope);
                if (state.Policy != null) state.Policy.RevokeClientContexts(prefix);
            }
        }

        private static string ClientRequestId(string scope, string requestId)
        {
            if (String.IsNullOrEmpty(requestId) || requestId.Length > 80 || requestId.Any(character =>
                !(character >= 'A' && character <= 'Z' || character >= 'a' && character <= 'z' ||
                  character >= '0' && character <= '9' || character == '_' || character == '-')))
                throw new ArgumentException("Use a stable requestId of at most 80 letters, digits, underscores or hyphens.");
            return scope == "default" ? requestId : PolicyEngine.Hash(scope + "\0" + requestId);
        }
    }
}
