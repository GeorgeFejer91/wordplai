using System;
using System.Collections.Generic;
using System.Linq;

namespace UnderTree.WordAddin
{
    public sealed partial class PolicyEngine
    {
        private static void ValidateRecolorRequest(HighlightRequest request)
        {
            if (request.RecolorMarkIds == null) return;
            if (request.RecolorMarkIds.Count != request.Ranges.Count || request.RecolorMarkIds.Any(id => id == null || !Identifier.IsMatch(id)) ||
                request.RecolorMarkIds.Distinct(StringComparer.Ordinal).Count() != request.RecolorMarkIds.Count)
                throw Refusal("Recoloring requires one unique current marker ID for each exact range, in the same order.");
        }

        private List<PassageMark> RecolorSelection(HighlightRequest request, string text)
        {
            var selected = new List<PassageMark>();
            if (request.RecolorMarkIds == null) return selected;
            for (int index = 0; index < request.RecolorMarkIds.Count; index++)
            {
                string id = request.RecolorMarkIds[index];
                PassageMark mark = Marks.FirstOrDefault(candidate => candidate.Id == id);
                if (mark == null) throw Refusal("The selected recolor marker " + id + " no longer exists; read the current document again.");
                if (!IsCurrent(mark, text) || mark.RequiresRemarking)
                    throw Refusal("The selected recolor marker " + id + " is stale or unresolved; inspect its current passage in Word.");
                HighlightRange range = request.Ranges[index];
                if (mark.From != range.From || mark.To != range.To || mark.Quote != range.Quote)
                    throw Refusal("Each recolor range and quote must exactly match its paired current marker.");
                if (mark.NativeBookmark != null || Proposals.Any(proposal => proposal.Status != "applied" && proposal.Status != "rejected" &&
                    proposal.NativeOriginalMarks != null && proposal.NativeOriginalMarks.Any(witness => witness.Id == id)))
                    throw Refusal("Resolve the selected passage's native tracked review before recoloring marker " + id + ".");
                selected.Add(mark);
            }
            return selected;
        }

        private static void ValidateRecolorPrompts(IEnumerable<PassageMark> selected, ColorRule color)
        {
            if (color.Action == "neutral" && selected.Any(mark => mark.Prompt.Length != 0))
                throw Refusal("This recolor would discard a nonempty passage prompt. Choose a functional color or clear the passage prompt locally first.");
        }

        private static void ValidateRecolorReceipt(HighlightReceipt receipt)
        {
            if (receipt.Request.RecolorMarkIds == null)
            {
                if (receipt.ReplacedMarkIds != null) throw Refusal("An ordinary highlighting receipt cannot retire marker identities.");
                return;
            }
            if (receipt.ReplacedMarkIds == null || !receipt.ReplacedMarkIds.SequenceEqual(receipt.Request.RecolorMarkIds) ||
                receipt.ReplacedMarkIds.Any(id => receipt.MarkIds.Contains(id)))
                throw Refusal("The recolor receipt must retain the exact replaced identities and distinct new marker identities.");
        }
    }
}
