using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace UnderTree.WordAddin
{
    internal sealed class WordAutoSave : IDisposable
    {
        private readonly dynamic application;
        private readonly Func<bool> available;
        private readonly Action refresh;
        private readonly Action<string> report;
        private readonly Func<object> targetDocument;
        private readonly Func<IEnumerable<object>> targetDocuments;
        private readonly Action<object> refreshDocument;
        private readonly Dictionary<long, WordAutoSave> documentWorkers = new Dictionary<long, WordAutoSave>();
        private readonly List<string> pendingErrors = new List<string>();
        private readonly Timer timer;
        private long documentId;
        private int bodyEnd, revisions, comments;
        private bool saved, pending, running, disposed;
        private DateTime due, retryAfter;
        internal string LastError { get; private set; }

        internal WordAutoSave(object application, Func<bool> available, Action refresh, Action<string> report, Func<object> targetDocument = null)
            : this(application, available, refresh, report, targetDocument, true) { }

        internal WordAutoSave(object application, Func<bool> available, Action<object> refreshDocument, Action<string> report, Func<IEnumerable<object>> targetDocuments)
            : this(application, available, null, report, null, true)
        {
            if (refreshDocument == null) throw new ArgumentNullException("refreshDocument");
            if (targetDocuments == null) throw new ArgumentNullException("targetDocuments");
            this.refreshDocument = refreshDocument;
            this.targetDocuments = targetDocuments;
        }

        private WordAutoSave(object application, Func<bool> available, Action refresh, Action<string> report, Func<object> targetDocument, bool startTimer)
        {
            this.application = application;
            this.available = available;
            this.refresh = refresh;
            this.report = report;
            this.targetDocument = targetDocument;
            if (startTimer)
            {
                timer = new Timer { Interval = 500 };
                timer.Tick += delegate { Pulse(DateTime.UtcNow); };
                timer.Start();
            }
        }

        internal void Changed(DateTime now)
        {
            if (targetDocuments != null)
            {
                foreach (WordAutoSave worker in documentWorkers.Values)
                    if (!worker.pending) worker.Changed(now);
                return;
            }
            pending = true;
            due = now.AddMilliseconds(750);
        }

        internal void Pulse(DateTime now)
        {
            if (disposed || running) return;
            running = true;
            string phase = "refresh the document's writing permissions";
            try
            {
                if (!available() || (int)application.Documents.Count == 0 ||
                    (int)application.UndoRecord.CustomRecordLevel != 0 || (int)application.BackgroundSavingStatus != 0) return;
                if (targetDocuments != null)
                {
                    if (now >= retryAfter) PulseDocuments(now);
                    return;
                }
                object target = TargetDocument();
                if (target == null) return;
                dynamic document = target;
                if ((bool)document.ReadOnly || (int)document.ProtectionType != -1) return;
                long identity = WordSnapshot.Identity((object)document);
                int end = (int)document.Content.End, revisionCount = (int)document.Revisions.Count, commentCount = (int)document.Comments.Count;
                bool isSaved = (bool)document.Saved;
                if (documentId != identity || bodyEnd != end || revisions != revisionCount || comments != commentCount || saved != isSaved)
                {
                    documentId = identity; bodyEnd = end; revisions = revisionCount; comments = commentCount; saved = isSaved;
                    Changed(now);
                }
                if (!pending || now < due || now < retryAfter) return;
                refresh();
                if (disposed || !available() || (int)application.Documents.Count == 0 ||
                    (int)application.UndoRecord.CustomRecordLevel != 0 || (int)application.BackgroundSavingStatus != 0) return;
                object refreshedTarget = TargetDocument();
                if (refreshedTarget == null || WordSnapshot.Identity(refreshedTarget) != identity) return;
                pending = false;
                phase = "automatically save this document";
                if (!(bool)document.Saved && !String.IsNullOrEmpty((string)document.Path) &&
                    !(bool)document.ReadOnly && (int)document.ProtectionType == -1)
                {
                    document.Save();
                    if (!(bool)document.Saved && (int)application.BackgroundSavingStatus == 0)
                        throw new InvalidOperationException("Word did not confirm that the changes were saved.");
                }
                bodyEnd = (int)document.Content.End; revisions = (int)document.Revisions.Count;
                comments = (int)document.Comments.Count; saved = (bool)document.Saved;
                LastError = null;
                retryAfter = DateTime.MinValue;
            }
            catch (Exception error)
            {
                pending = true;
                retryAfter = now.AddSeconds(10);
                string message = AppBranding.Name + " could not " + phase + ". Use Word's Save command or resolve the document issue before continuing.\n\n" + error.Message;
                if (LastError != message) { LastError = message; report(message); }
            }
            finally { running = false; }
        }

        private object TargetDocument()
        { return targetDocument == null ? (object)application.ActiveDocument : targetDocument(); }

        private Dictionary<long, object> Targets()
        {
            var targets = new Dictionary<long, object>();
            int inspected = 0;
            foreach (object target in targetDocuments() ?? new object[0])
            {
                if (++inspected > 32) throw new InvalidOperationException("At most 32 background documents can be refreshed and saved.");
                if (target == null) continue;
                targets[WordSnapshot.Identity(target)] = target;
            }
            return targets;
        }

        private object Target(long identity)
        {
            object target;
            return Targets().TryGetValue(identity, out target) ? target : null;
        }

        private void PulseDocuments(DateTime now)
        {
            Dictionary<long, object> targets = Targets();
            foreach (long closed in documentWorkers.Keys.Where(identity => !targets.ContainsKey(identity)).ToArray())
            {
                documentWorkers[closed].Dispose();
                documentWorkers.Remove(closed);
            }
            foreach (var target in targets)
            {
                WordAutoSave worker;
                if (!documentWorkers.TryGetValue(target.Key, out worker))
                {
                    long identity = target.Key;
                    object document = target.Value;
                    worker = new WordAutoSave((object)application, available, () => refreshDocument(document), pendingErrors.Add, () => Target(identity), false);
                    documentWorkers.Add(identity, worker);
                }
                worker.Pulse(now);
            }
            LastError = documentWorkers.Values.Select(worker => worker.LastError).FirstOrDefault(error => error != null);
            retryAfter = DateTime.MinValue;
            string[] errors = pendingErrors.Distinct().ToArray();
            pendingErrors.Clear();
            foreach (string error in errors)
                if (!disposed) report(error);
        }

        public void Dispose()
        {
            disposed = true;
            if (timer != null) timer.Dispose();
            foreach (WordAutoSave worker in documentWorkers.Values) worker.Dispose();
            documentWorkers.Clear();
            pendingErrors.Clear();
        }
    }
}
