using System;

namespace UnderTree.WordAddin
{
    internal static class WordNativeOutcome
    {
        internal enum Stage
        {
            window_prepare, native_prepare, protected_prepare, begin_application, begin_undo,
            tracking_settings, original_anchors, replace_text, restore_formatting, restore_protected,
            remap_markers, repaint_markers, record_review, record_acceptance, verify_discussion,
            add_comment, add_undo_guard, persist_native, review_visibility, restore_settings,
            capture_review, end_undo, verify_protected, verify_snapshot, commit_receipt,
            persist_commit, capture_commit, settings_cleanup, undo_cleanup, protected_cleanup,
            window_cleanup, continuation_context, review_core, review_snapshot_cleanup,
            review_changed_callback, uncertain_receipt, revoke_authority, persist_uncertainty
        }

        internal sealed class Diagnostic
        {
            internal Stage Phase = Stage.native_prepare;
            internal bool Committed { get; private set; }
            private bool failed, rollbackAttempted, rollbackConfirmed;
            private string firstStage, firstType;
            private int firstHResult;
            private object rollbackError;

            internal void Capture(Stage stage, Exception error, bool committed)
            {
                Committed |= committed;
                if (failed) return;
                failed = true;
                firstStage = stage.ToString(); firstType = error.GetType().FullName; firstHResult = error.HResult;
            }

            internal bool AttemptRollback(Func<bool> rollback)
            {
                rollbackAttempted = true;
                try { rollbackConfirmed = rollback(); return rollbackConfirmed; }
                catch (Exception error)
                {
                    rollbackError = new { exception_type = error.GetType().FullName, hresult = error.HResult };
                    throw;
                }
            }

            internal object Details()
            {
                return new { stage = firstStage, exception_type = firstType, hresult = failed ? (int?)firstHResult : null,
                    committed = Committed, rollback = new { attempted = rollbackAttempted, confirmed = rollbackConfirmed, error = rollbackError } };
            }
        }

        internal static void Unknown(Action record, Action revoke, Action persist, Diagnostic diagnostic = null)
        {
            try { record(); }
            catch (Exception error) { if (diagnostic != null) diagnostic.Capture(Stage.uncertain_receipt, error, diagnostic.Committed); throw; }
            finally
            {
                try { revoke(); }
                catch (Exception error) { if (diagnostic != null) diagnostic.Capture(Stage.revoke_authority, error, diagnostic.Committed); throw; }
                finally
                {
                    try { persist(); }
                    catch (Exception error) { if (diagnostic != null) diagnostic.Capture(Stage.persist_uncertainty, error, diagnostic.Committed); throw; }
                    finally
                    {
                        throw UnknownException(diagnostic);
                    }
                }
            }
        }

        internal static McpOperationException UnknownException(Diagnostic diagnostic = null)
        {
            object details = diagnostic == null ? (object)new { operation_started = true, outcome_unknown = true } :
                new { operation_started = true, outcome_unknown = true, diagnostic = diagnostic.Details() };
            return new McpOperationException("Word's result could not be confirmed. Inspect Word and Undo locally. Retain the exact original request_id (or marking.requestId) and arguments for receipt lookup; never submit this unresolved edit under a new request ID.",
                "outcome_unknown", details);
        }
    }
}
