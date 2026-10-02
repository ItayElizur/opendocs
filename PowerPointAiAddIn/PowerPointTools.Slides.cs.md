# PowerPointTools.Slides.cs

## `DeleteSlide`

PP-19 Task 1: delete_slide/move_slide/duplicate_slide. Deliberately no
slideIndexes:number[] batch form - deleting slide 2 shifts every later
slide's index down by one, so a batch would need to either resolve all
targets up front or delete in strict descending order. One slide per
call (with the index-shift warning in the output/description) is the
safer answer; it makes the model re-read the deck between deletes
instead of silently deleting the wrong slides.
