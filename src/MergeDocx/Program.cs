using MergeDocx;

var basePath = Path.Combine(Directory.GetCurrentDirectory(), "..\\..\\..\\Templates\\");
Merger.Merge(new[] { $"{basePath}Doc1.docx", $"{basePath}Doc3.docx" }, $"{basePath}result.docx");

//var basePath = Path.Combine(Directory.GetCurrentDirectory(), "..\\..\\..\\Templates\\1\\");
//Merger.Merge(new[] { $"{basePath}Doc1.docx", $"{basePath}ContentPart_doc.docx" }, $"{basePath}result.docx");