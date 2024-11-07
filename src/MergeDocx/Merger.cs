using Codeuctivity.OpenXmlPowerTools;
using Codeuctivity.OpenXmlPowerTools.DocumentBuilder;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;

namespace MergeDocx;

public static class Merger
{
    public static void Merge(string[] files, string output)
    {
        var sources = new List<Source>();
        var headers = new HashSet<string>();
        var headersNames = new HashSet<string>();
        var i = 0;
        foreach (var path in files)
        {
            using var file = new StreamReader(path);
            using var ms = new MemoryStream();
            file.BaseStream.CopyTo(ms);
            
            TryReplaceContentPart(ms);
            ms.Seek(0, SeekOrigin.Begin);
            StyleConflictResolution(i++, headers, ms, headersNames);

            var doc = new WmlDocument(path, ms);

            var source = new Source(doc)
            {
                KeepSections = true,
            };
            sources.Add(source);
        }

        var result = DocumentBuilder.BuildDocument(sources);
        using MemoryStream memoryStream = new MemoryStream(result.DocumentByteArray);
        File.WriteAllBytes(output, memoryStream.ToArray());
    }

    private static void TryReplaceContentPart(MemoryStream ms)
    {
        using var document = WordprocessingDocument.Open(ms, true);

        var contentPartEls = document.MainDocumentPart?.Document?.Body
            ?.Descendants()
            ?.Where(e => e.LocalName == "contentPart");

        if (!contentPartEls.Any())
        {
            return;
        }

        foreach (var el in contentPartEls)
        {
            var parent = el?.Ancestors<AlternateContent>()?.First();

            var choice = parent?.Elements()?.FirstOrDefault(e => e is AlternateContentChoice);
            var fallback = parent?.Elements()?.FirstOrDefault(e => e is AlternateContentFallback);
            var fallbackChild = fallback?.FirstChild?.CloneNode(true);

            fallback?.Remove();
            choice?.FirstChild?.Remove();
            choice?.AppendChild(fallbackChild);
        }

        document.Save();
    }

    private static void StyleConflictResolution(int i,HashSet<string> existsHeaders, MemoryStream ms, HashSet<string> headersNames)
    {
        using WordprocessingDocument doc = WordprocessingDocument.Open(ms, true);
        var mainPart = doc.MainDocumentPart;
        if (mainPart == null) return;

        var stylesPart = mainPart.StyleDefinitionsPart;
        if (stylesPart != null)
        {
            foreach (var style in stylesPart.Styles.Elements<Style>())
            {
                if (existsHeaders.Any(x => x == style.StyleId) || headersNames.Any(x => x == style.StyleName.Val))
                {
                    var newStyleId = $"{style.StyleId}-{$"{Guid.NewGuid()}".Substring(0, 5)}";
                    var oldStyleId = style.StyleId;

                    style.StyleId = newStyleId;

                    if (style.StyleName != null)
                    {
                        style.StyleName.Val = newStyleId;
                    }

                    UpdateStyleReferences(mainPart.Document, oldStyleId, newStyleId);

                    if (mainPart.HeaderParts != null)
                    {
                        foreach (var headerPart in mainPart.HeaderParts)
                        {
                            UpdateStyleReferences(headerPart.Header, oldStyleId, newStyleId);
                        }
                    }

                    if (mainPart.FooterParts != null)
                    {
                        foreach (var footerPart in mainPart.FooterParts)
                        {
                            UpdateStyleReferences(footerPart.Footer, oldStyleId, newStyleId);
                        }
                    }

                    var numberingPart = mainPart.NumberingDefinitionsPart;
                    if (numberingPart != null)
                    {
                        var abstractNums = numberingPart.Numbering.Elements<AbstractNum>();
                        foreach (var abstractNum in abstractNums)
                        {
                            foreach (var lvl in abstractNum.Elements<Level>())
                            {
                                var pStyle = lvl.Elements<ParagraphProperties>()?.FirstOrDefault()?.ParagraphStyleId;
                                if (pStyle?.Val == oldStyleId)
                                {
                                    pStyle.Val = newStyleId;
                                }
                            }
                        }

                        UpdateStyleReferences(numberingPart.Numbering, oldStyleId, newStyleId);
                    }
                }

                headersNames.Add(style.StyleName.Val);
            }
        }

        mainPart.Document.Save();
    }

    private static void UpdateStyleReferences(OpenXmlElement element, string oldStyleId, string newStyleId)
    {
        var paragraphs = element.Descendants<Paragraph>();
        foreach (var paragraph in paragraphs)
        {
            var pStyle = paragraph.ParagraphProperties?.ParagraphStyleId;
            if (pStyle?.Val?.Value == oldStyleId)
            {
                pStyle.Val = newStyleId;
            }
        }

        var runs = element.Descendants<Run>();
        foreach (var run in runs)
        {
            var rStyle = run.RunProperties?.RunStyle;
            if (rStyle?.Val?.Value == oldStyleId)
            {
                rStyle.Val = newStyleId;
            }
        }

        var tables = element.Descendants<Table>();
        foreach (var table in tables)
        {
            foreach (var tblPr in table.Elements<TableProperties>())
            {
                var tblStyle = tblPr.TableStyle;
                if (tblStyle?.Val?.Value == oldStyleId)
                {
                    tblStyle.Val = newStyleId;
                }

            }
        }
    }
}