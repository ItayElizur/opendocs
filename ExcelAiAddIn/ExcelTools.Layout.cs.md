# ExcelTools.Layout.cs

## `CopyOrMoveRange`

copy_range/move_range share this one method (cut:false/true), same
shape as InsertDeleteRows/InsertDeleteCols above. The 2000-cell cap
matches ReadRange's, but is unmeasured for a Copy/Cut-based op
specifically - ReadRange's cap exists because of Value2 marshaling
cost back across COM, which a native Copy/Cut (Excel-side only)
doesn't incur, so this may be needlessly conservative once tested
against real Excel.

## `CopyOrMoveRange` - destination lookup

Fresh, plain indexer lookup - never a .Resize/.Offset-chained
object passed as another COM method's argument. See
Chart.SetSourceData's E_INVALIDARG scar (STATUS.md's "Live
debugging session") - Destination is documented to genuinely
accept a Range here, so this exact failure isn't expected to
recur, but a plain lookup costs nothing and removes a variable.
