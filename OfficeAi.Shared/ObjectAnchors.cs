using System.Collections.Generic;

namespace OfficeAi.Shared
{
    /// <summary>A non-body object (table, shape, footnote, ...) pinned to a character offset in the main story.</summary>
    public struct ObjectAnchor
    {
        public int Start;
        public string Label;

        public ObjectAnchor(int start, string label)
        {
            Start = start;
            Label = label;
        }
    }

    /// <summary>
    /// Pure merge step for Word's read_blocks: attaches objects to the paragraph whose character range
    /// contains their anchor. Kept COM-free so it can be unit tested.
    /// </summary>
    public static class ObjectAnchors
    {
        /// <summary>
        /// Returns the "  {label, label}" suffix for the paragraph [paraStart, paraEnd), or "" when no object
        /// is anchored in it. <paramref name="anchors"/> must be sorted ascending by Start and
        /// <paramref name="next"/> is the caller's cursor into it, so paragraphs must be passed in document order.
        /// Anchors before <paramref name="paraStart"/> (outside the requested block range) are skipped.
        /// </summary>
        public static string LabelsFor(IList<ObjectAnchor> anchors, ref int next, int paraStart, int paraEnd, int maxLabels)
        {
            while (next < anchors.Count && anchors[next].Start < paraStart) next++;

            var labels = new List<string>();
            int extra = 0;
            while (next < anchors.Count && anchors[next].Start < paraEnd)
            {
                if (labels.Count < maxLabels) labels.Add(anchors[next].Label);
                else extra++;
                next++;
            }
            if (labels.Count == 0) return "";
            return "  {" + string.Join(", ", labels) + (extra > 0 ? ", +" + extra + " more" : "") + "}";
        }
    }
}
