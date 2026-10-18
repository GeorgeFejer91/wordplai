using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;

namespace UnderTree.WordAddin
{
    internal sealed class McpClientSession
    {
        internal string Id, ScopePrefix;
        internal long LastActivity;
        internal bool Active = true;
        internal readonly HashSet<McpLogicalConversation> Conversations = new HashSet<McpLogicalConversation>();
    }

    internal sealed class McpLogicalConversation
    {
        internal string HintHash, ScopePrefix;
        internal long LastActivity;
        internal bool Active = true;
    }

    internal sealed class McpCallScope
    {
        internal readonly McpClientSession Session;
        internal readonly McpLogicalConversation Conversation;
        internal readonly string NativeScope;

        internal McpCallScope(McpClientSession session, McpLogicalConversation conversation)
        {
            Session = session; Conversation = conversation;
            NativeScope = (conversation == null ? session.ScopePrefix : conversation.ScopePrefix) + "default";
        }
    }

    internal sealed class McpCoordination
    {
        internal const string Protocol = "undertree-concurrent-review/2";
        internal const int MaxSessions = 512;
        internal const int MaxConversations = 64;
        internal const int IdleTimeoutSeconds = 1800;
        internal const int MaxQueuedEdits = 8;
        private readonly object gate = new object();
        private readonly Dictionary<string, McpClientSession> sessions = new Dictionary<string, McpClientSession>();
        private readonly Dictionary<string, McpLogicalConversation> conversations = new Dictionary<string, McpLogicalConversation>();
        private sealed class RetiredScope
        {
            internal string Prefix;
            internal bool Conversation;
        }
        private readonly Queue<RetiredScope> retiredScopes = new Queue<RetiredScope>();
        private int retiredSessions, retiredConversations;
        private readonly Queue<long> turns = new Queue<long>();
        private readonly Func<long> clock;
        private readonly long idleTicks;
        private long sequence;
        private bool active = true;

        internal McpCoordination() : this(Stopwatch.GetTimestamp, Stopwatch.Frequency * (long)IdleTimeoutSeconds) { }
        internal McpCoordination(Func<long> clock, long idleTicks)
        { this.clock = clock; this.idleTicks = idleTicks; }

        internal McpClientSession Create()
        {
            lock (gate)
            {
                Expire();
                if (!active) throw new McpHttpException(404);
                if (sessions.Count + retiredSessions >= MaxSessions) throw new McpHttpException(429);
                string id = Guid.NewGuid().ToString("N") + Guid.NewGuid().ToString("N");
                var session = new McpClientSession { Id = id, ScopePrefix = "mcp_" + Digest(id) + ":", LastActivity = clock() };
                sessions.Add(id, session);
                return session;
            }
        }

        internal McpClientSession Require(string id)
        {
            lock (gate)
            {
                Expire();
                if (String.IsNullOrEmpty(id)) throw new McpHttpException(400);
                McpClientSession session;
                if (!active || !sessions.TryGetValue(id, out session)) throw new McpHttpException(404);
                session.LastActivity = clock();
                return session;
            }
        }

        internal void Retire(McpClientSession session)
        {
            lock (gate)
            {
                RetireSession(session, true);
            }
        }

        private void RetireSession(McpClientSession session, bool explicitDelete)
        {
            if (!session.Active) return;
            if (explicitDelete)
                foreach (McpLogicalConversation conversation in session.Conversations.ToArray()) RetireConversation(conversation);
            session.Active = false;
            session.Conversations.Clear();
            sessions.Remove(session.Id);
            retiredScopes.Enqueue(new RetiredScope { Prefix = session.ScopePrefix });
            retiredSessions++;
            Monitor.PulseAll(gate);
        }

        private void RetireConversation(McpLogicalConversation conversation)
        {
            if (!conversation.Active) return;
            conversation.Active = false;
            conversations.Remove(conversation.HintHash);
            foreach (McpClientSession session in sessions.Values) session.Conversations.Remove(conversation);
            retiredScopes.Enqueue(new RetiredScope { Prefix = conversation.ScopePrefix, Conversation = true });
            retiredConversations++;
            Monitor.PulseAll(gate);
        }

        private void Expire()
        {
            long now = clock();
            foreach (McpClientSession session in sessions.Values.Where(value => now - value.LastActivity >= idleTicks).ToArray())
                RetireSession(session, false);
            foreach (McpLogicalConversation conversation in conversations.Values.Where(value => now - value.LastActivity >= idleTicks).ToArray())
                RetireConversation(conversation);
        }

        internal bool IsActive(McpClientSession session)
        { lock (gate) { Expire(); return active && session.Active; } }

        internal bool IsActive(McpCallScope scope)
        { lock (gate) { Expire(); return IsLive(scope); } }

        private bool IsLive(McpCallScope scope)
        { return active && scope.Session.Active && (scope.Conversation == null || scope.Conversation.Active); }

        internal bool HasRetired
        { get { lock (gate) { Expire(); return retiredScopes.Count != 0; } } }

        internal McpCallScope Resolve(McpClientSession session, string hint)
        {
            lock (gate)
            {
                Expire();
                if (!active || !session.Active) throw new McpOperationException("The client session was revoked. No operation started.", "session_revoked");
                session.LastActivity = clock();
                McpLogicalConversation conversation = null;
                if (hint != null)
                {
                    string key = Digest(hint);
                    if (!conversations.TryGetValue(key, out conversation))
                    {
                        if (conversations.Count + retiredConversations >= MaxConversations)
                            throw new McpOperationException("Word reached its active conversation limit. No operation started.", "conversation_limit",
                                new { operation_started = false, outcome_unknown = false });
                        conversation = new McpLogicalConversation { HintHash = key,
                            ScopePrefix = "mcp_" + Digest(Guid.NewGuid().ToString("N") + Guid.NewGuid().ToString("N")) + ":" };
                        conversations.Add(key, conversation);
                    }
                    conversation.LastActivity = clock();
                    session.Conversations.Add(conversation);
                }
                return new McpCallScope(session, conversation);
            }
        }

        internal void DrainRetired(Action<string> revoke)
        {
            while (true)
            {
                RetiredScope scope;
                lock (gate) { if (retiredScopes.Count == 0) return; scope = retiredScopes.Peek(); }
                revoke(scope.Prefix);
                lock (gate)
                    if (retiredScopes.Count != 0 && retiredScopes.Peek() == scope)
                    {
                        retiredScopes.Dequeue();
                        if (scope.Conversation) retiredConversations--; else retiredSessions--;
                    }
            }
        }

        internal long Enqueue()
        {
            lock (gate)
            {
                if (!active || turns.Count >= MaxQueuedEdits) throw new McpOperationException("Word's review queue is full. No operation started.", "word_busy",
                    new { operation_started = false, outcome_unknown = false });
                long ticket = checked(++sequence);
                turns.Enqueue(ticket);
                return ticket;
            }
        }

        internal void Await(long ticket, McpCallScope scope, long deadline)
        {
            lock (gate)
            {
                Expire();
                while (IsLive(scope) && turns.Count != 0 && turns.Peek() != ticket)
                {
                    long remaining = deadline - Stopwatch.GetTimestamp();
                    if (remaining <= 0) throw Busy();
                    Monitor.Wait(gate, (int)Math.Min(100, Math.Max(1, remaining * 1000 / Stopwatch.Frequency)));
                    Expire();
                }
                if (!IsLive(scope)) throw new McpOperationException("The client session or conversation was revoked. No operation started.", "session_revoked");
                if (Stopwatch.GetTimestamp() >= deadline) throw Busy();
            }
        }

        private static McpOperationException Busy()
        { return new McpOperationException("Word did not start this queued operation before its deadline.", "word_busy", new { operation_started = false, outcome_unknown = false }); }

        internal void Finish(long ticket)
        {
            lock (gate)
            {
                long[] remaining = turns.Where(value => value != ticket).ToArray();
                turns.Clear();
                foreach (long value in remaining) turns.Enqueue(value);
                Monitor.PulseAll(gate);
            }
        }

        internal void Stop()
        {
            lock (gate)
            {
                active = false;
                foreach (McpClientSession session in sessions.Values) { session.Active = false; session.Conversations.Clear(); }
                foreach (McpLogicalConversation conversation in conversations.Values) conversation.Active = false;
                sessions.Clear();
                conversations.Clear();
                retiredScopes.Clear();
                retiredSessions = retiredConversations = 0;
                Monitor.PulseAll(gate);
            }
        }

        internal static string Digest(string value)
        {
            using (SHA256 hash = SHA256.Create())
                return BitConverter.ToString(hash.ComputeHash(Encoding.UTF8.GetBytes(value))).Replace("-", "").ToLowerInvariant();
        }
    }

    internal sealed class McpHttpException : Exception
    {
        internal readonly int Status;
        internal McpHttpException(int status) { Status = status; }
    }
}
