using System;
using System.Collections.Generic;

namespace OfficeAi.Shared
{
    /// <summary>
    /// Undo/redo bookkeeping for a host with no usable native undo for the
    /// assistant's own actions (Outlook - see OutlookTools.Undo.cs). Pure data
    /// structure; knows nothing about how an entry is reversed, only the
    /// order entries come back in. Full caller protocol: ActionHistory.cs.md.
    /// </summary>
    public sealed class ActionHistory<T> where T : class
    {
        public const int DefaultCapacity = 50;

        private sealed class Slot
        {
            public T Entry;
            public string Barrier; // non-null for a barrier slot
        }

        private readonly int _capacity;
        private readonly LinkedList<Slot> _undo = new LinkedList<Slot>(); // newest at Last
        private readonly Stack<T> _redo = new Stack<T>();

        public ActionHistory(int capacity = DefaultCapacity)
        {
            if (capacity < 1) throw new ArgumentOutOfRangeException(nameof(capacity));
            _capacity = capacity;
        }

        public int UndoCount { get { return _undo.Count; } }
        public int RedoCount { get { return _redo.Count; } }

        /// <summary>Records a new reversible action. Clears redo, like any editor.</summary>
        public void Push(T entry)
        {
            if (entry == null) throw new ArgumentNullException(nameof(entry));
            AddUndo(new Slot { Entry = entry });
            _redo.Clear();
        }

        /// <summary>Records an irreversible action. Clears redo.</summary>
        public void PushBarrier(string description)
        {
            AddUndo(new Slot { Barrier = description ?? "" });
            _redo.Clear();
        }

        /// <summary>
        /// Returns false with both outputs null when there is nothing to undo,
        /// or false with <paramref name="barrier"/> set when the most recent
        /// action is irreversible (the barrier stays in place).
        /// </summary>
        public bool TryTakeUndo(out T entry, out string barrier)
        {
            entry = null;
            barrier = null;
            if (_undo.Count == 0) return false;
            Slot top = _undo.Last.Value;
            if (top.Barrier != null)
            {
                barrier = top.Barrier;
                return false;
            }
            _undo.RemoveLast();
            entry = top.Entry;
            return true;
        }

        public void CompleteUndo(T entry)
        {
            if (entry == null) throw new ArgumentNullException(nameof(entry));
            _redo.Push(entry);
        }

        public bool TryTakeRedo(out T entry)
        {
            entry = _redo.Count > 0 ? _redo.Pop() : null;
            return entry != null;
        }

        /// <summary>Puts a redone entry back on the undo stack without clearing redo.</summary>
        public void CompleteRedo(T entry)
        {
            if (entry == null) throw new ArgumentNullException(nameof(entry));
            AddUndo(new Slot { Entry = entry });
        }

        /// <summary>
        /// Visits every entry on both stacks - used to rewrite item ids after an
        /// action that changes them (an Outlook move assigns a new EntryID).
        /// </summary>
        public void ForEachEntry(Action<T> visit)
        {
            foreach (Slot s in _undo)
                if (s.Entry != null) visit(s.Entry);
            foreach (T e in _redo)
                visit(e);
        }

        private void AddUndo(Slot slot)
        {
            _undo.AddLast(slot);
            while (_undo.Count > _capacity) _undo.RemoveFirst();
        }
    }
}
