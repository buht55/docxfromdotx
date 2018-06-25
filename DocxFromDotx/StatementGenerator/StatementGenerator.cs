using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Break = DocumentFormat.OpenXml.Wordprocessing.Break;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using Path = System.IO.Path;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using TableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
using DocxFromDotx.StatementGenerator.Types;


namespace DocxFromDotx.StatementGenerator
{
    public abstract class StatementGenerator
    {
        private readonly ReportItems _marks;
        public string Name { get; set; }

        protected StatementGenerator(string name)
        {
            Name = name;
            _marks = new ReportItems();
        }

        private void Parse(WordprocessingDocument doc, List<OpenXmlElement> childsNodes, ReportItems reportItems)
        {
            if (childsNodes == null || childsNodes.Count == 0)
                return;
            var blocks = GetParagraphs(childsNodes).Where(y => Regex.IsMatch(y.InnerText, @"<[\w-]+>"));
            foreach (var begBlock in blocks)
            {
                var endBlock = GetParagraphs(childsNodes).FirstOrDefault(y => Regex.IsMatch(y.InnerText, begBlock.InnerText.Replace("<", "</")));
                if (endBlock == null || begBlock.Parent != endBlock.Parent)
                    continue;
                var tmpBlock = GetNextElement(childsNodes, begBlock);
                var nodes = new List<OpenXmlElement>();
                while (tmpBlock != endBlock && tmpBlock != null)
                {
                    nodes.Add(tmpBlock);
                    tmpBlock = GetNextElement(childsNodes, tmpBlock);
                }
                RemoveChildNodes(childsNodes, nodes);
                var blockName = begBlock.InnerText.Replace("<", "").Replace(">", "");

                var reportBlock = reportItems.GetReportBlock(blockName);
                if (reportBlock != null)
                {
                    foreach (var row in reportBlock.Rows)
                    {
                        var insertedNodes = InsertChildNodes(childsNodes, nodes, endBlock);
                        Parse(doc, insertedNodes, row);
                    }
                    if (reportBlock.ShowFromNewPage && begBlock.Parent is Body)
                    {
                        var elem = begBlock.NextSibling();
                        if (elem is Paragraph && elem != endBlock)
                            elem.PrependChild(new Run(new Break { Type = BreakValues.Page }));
                    }
                }
                RemoveBegEndNodes(childsNodes, begBlock, endBlock);

            }
            var paragraphsWithMarks = GetParagraphs(childsNodes).Where(y => Regex.IsMatch(y.InnerText, @".*\[[\w-]+\].*")).ToList();
            foreach (var paragraph in paragraphsWithMarks)
            {
                foreach (Match markMatch in Regex.Matches(paragraph.InnerText, @"\[[\w-]+\]", RegexOptions.Compiled))
                {
                    var paragraphMarkValue = markMatch.Value.Trim('[', ']');
                    var markValueFromCollection = reportItems.GetReportRecord(paragraphMarkValue)?.Value ?? "";
                    var rPr = paragraph.Descendants<Run>().FirstOrDefault()?.RunProperties?.CloneNode(true);
                    var editedParagraphText = paragraph.InnerText.Replace(markMatch.Value, markValueFromCollection);
                    paragraph.RemoveAllChildren<Run>();
                    var elems = editedParagraphText.Split("\n");
                    var list = new List<OpenXmlElement> { new Text(elems[0]) };
                    for (var i = 1; i < elems.Length; i++)
                    {
                        list.Add(new Break());
                        list.Add(new Text(elems[i]));
                    }
                    var r = new Run(list);
                    if (rPr != null)
                        r.PrependChild(rPr);
                    paragraph.AppendChild(r);
                }
            }
        }

        private List<OpenXmlElement> GetParagraphs(List<OpenXmlElement> nodes)
        {
            var list = nodes.Where(x =>
            {
                return x is Paragraph;
            }).ToList();
            list.AddRange(nodes.SelectMany(x => x.Descendants<Paragraph>()));
            return list;
        }

        private OpenXmlElement GetNextElement(List<OpenXmlElement> childNodes, OpenXmlElement node)
        {
            if (node.Parent != null)
                return node.NextSibling();
            var index = childNodes.IndexOf(node);
            if (index == -1 || ++index >= childNodes.Count)
                return null;
            return childNodes[index];
        }

        private static void RemoveChildNodes(List<OpenXmlElement> nodes, List<OpenXmlElement> childNodes)
        {
            if (childNodes.Count == 0)
                return;
            var parent = childNodes[0].Parent;
            if (parent != null)
            {
                foreach (var node in childNodes)
                    parent.RemoveChild(node);
                return;
            }
            var begIndex = nodes.IndexOf(childNodes[0]);
            var endIndex = nodes.IndexOf(childNodes[childNodes.Count - 1]);
            for (var i = endIndex; i >= begIndex; i--)
                nodes.RemoveAt(i);
        }

        private void RemoveBegEndNodes(List<OpenXmlElement> nodes, OpenXmlElement begNode, OpenXmlElement endNode)
        {
            var parent = begNode.Parent;
            if (parent != null)
            {
                parent.RemoveChild(begNode);
                parent.RemoveChild(endNode);
                if (parent is TableCell && !parent.Descendants<Paragraph>().Any())
                    parent.AppendChild(new Paragraph());
            }
            else
            {
                nodes.Remove(begNode);
                nodes.Remove(endNode);
            }
        }

        private List<OpenXmlElement> InsertChildNodes(List<OpenXmlElement> nodes, List<OpenXmlElement> childNodes, OpenXmlElement endNode)
        {
            if (nodes == null || nodes.Count == 0 || childNodes == null || childNodes.Count == 0 || endNode == null)
                return childNodes;
            var list = childNodes.Select(x => x.CloneNode(true)).ToList();
            if (endNode.Parent != null)
            {
                foreach (var node in list)
                    endNode.Parent.InsertBefore(node, endNode);
                return list;
            }
            var index = nodes.IndexOf(endNode);
            foreach (var node in list)
                nodes.Insert(index++, node);
            return list;
        }

        protected abstract void FillMarks(ReportItems marks);

        protected void FillDocument(WordprocessingDocument doc)
        {
            FillMarks(_marks);
            foreach (var header in doc.MainDocumentPart.HeaderParts)
                Parse(doc, header.Header.ChildElements.ToList(), _marks);
            Parse(doc, doc.MainDocumentPart.Document.Body.ChildElements.ToList(), _marks);
        }

        public virtual byte[] BuildReport()
        {

            var path = Path.Combine(Path.GetDirectoryName(Assembly.GetEntryAssembly().Location),
                $"StatementGenerator{Path.DirectorySeparatorChar}dotx{Path.DirectorySeparatorChar}{Name}.dotx");
            var mem = new MemoryStream();
            var byteArray = File.ReadAllBytes(path);
            mem.Write(byteArray, 0, byteArray.Length);
            using (var document = WordprocessingDocument.Open(mem, true))
            {
                document.ChangeDocumentType(WordprocessingDocumentType.Document);
                FillDocument(document);
                document.Save();
            }
            mem.Seek(0, SeekOrigin.Begin);
            return mem.ToArray();
        }

    }
}
