## AttachmentTextExtractor

```
// The single seam for turning a saved email attachment into plain text for
// the model. OutlookTools' get_attachment calls only TryExtract - replacing
// this whole folder with an HTTP call to an external parser API (the user's
// planned Word/PowerPoint/PDF service) is a localized change that never
// touches the tool code.
//
// Deliberately handles text-family types and OpenXML (.docx/.xlsx/.pptx)
// only. .pdf and images return false - the caller then omits extracted_text
// and returns just the saved file path.
```
