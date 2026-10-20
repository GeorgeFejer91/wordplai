namespace UnderTree.WordAddin
{
    internal sealed class McpResourceDiagnostics
    {
        private readonly object gate = new object();
        private long httpRequests;
        private long httpRefusals;
        private long rpcErrors;
        private long originRefusals;
        private long resourceLists;
        private long resourceReadAttempts;
        private long resourceReadResults;
        private long resourceReadErrors;
        private long resourceResponsesWritten;

        internal void OriginRefused() { lock (gate) originRefusals++; }
        internal void HttpRequested() { lock (gate) httpRequests++; }
        internal void HttpRefused() { lock (gate) httpRefusals++; }
        internal void RpcFailed() { lock (gate) rpcErrors++; }
        internal void ResourceListed() { lock (gate) resourceLists++; }
        internal void ReadStarted() { lock (gate) resourceReadAttempts++; }
        internal void ReadSucceeded() { lock (gate) resourceReadResults++; }
        internal void ReadFailed() { lock (gate) resourceReadErrors++; }
        internal void ResponseWritten() { lock (gate) resourceResponsesWritten++; }

        internal object Snapshot()
        {
            lock (gate) return new {
                httpRequests = httpRequests,
                httpRefusals = httpRefusals,
                rpcErrors = rpcErrors,
                originRefusals = originRefusals,
                resourceLists = resourceLists,
                resourceReadAttempts = resourceReadAttempts,
                resourceReadResults = resourceReadResults,
                resourceReadErrors = resourceReadErrors,
                resourceResponsesWritten = resourceResponsesWritten
            };
        }
    }
}
