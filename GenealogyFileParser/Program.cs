using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using System.Configuration;
using System.Reflection;
using HtmlAgilityPack;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace GenealogyFileParser
{
    class Program
    {
        static void Main(string[] args)
        {
            ReadDocument();
        }

        static void ReadDocument()
        {
            string filepath = ConfigurationManager.AppSettings.Get("filename");

            using (Stream stream = new FileStream(filepath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(stream, false))
                {
                    var xmlString = wordDocument.MainDocumentPart.Document.Body.InnerXml;
                    var xml = XElement.Parse("<root>" + xmlString + "</root>");

                    var doc = BuildHtml(xml, "", 999);
                    SaveHtml(doc);
                }
            }
        }

        static HtmlDocument BuildHtml(XElement xml, string entryPoint = "", int depth = 999)
        {
            bool includeSiblings = true;
            var elements = xml.Elements().ToList();
            if(entryPoint.Trim() != "")
            {
                var entryIndex = elements.FindIndex(e => e.Value.StartsWith(entryPoint));
                elements = elements.Skip(entryIndex).ToList();
                includeSiblings = false;
            }

            HtmlDocument doc = new HtmlDocument();

            if(elements.Count == 0)
            {
                return doc;
            }

            var topIlvl = GetIlvl(elements[0]);
            var currentIlvl = topIlvl;
            var currentNode = doc.DocumentNode;

            for (int i = 0; i < elements.Count; i++)
            {
                var element = elements[i];
                var ilvl = GetIlvl(element);

                bool emptyNode = element.Value.Trim() == "";
                bool outOfScope = ilvl < topIlvl;
                bool beyondDepth = ilvl > topIlvl + depth;

                if (outOfScope) // if this node is above the top level, quit
                {
                    break;
                }
                if (emptyNode || beyondDepth) // skip empty nodes
                {
                    continue;
                }

                var person = ilvl == topIlvl ? // don't add top lvl nodes to a list
                    HtmlNode.CreateNode("<p>" + element.Value + "</p") :
                    HtmlNode.CreateNode("<li>" + element.Value + "</li>");

                if(ilvl > currentIlvl) // start a new list
                {
                    var list = currentNode.AppendChild(HtmlNode.CreateNode("<ol></ol>"));
                    list.AppendChild(person);
                    currentIlvl = ilvl;
                    currentNode = list;
                } else if(ilvl == currentIlvl) // append to this list
                {
                    currentNode.AppendChild(person);
                } else // traverse the tree up to the correct level
                {
                    for (; currentIlvl > ilvl; currentIlvl--)
                    {
                        currentNode = currentNode.ParentNode;
                    }
                    if (currentIlvl == topIlvl && !includeSiblings)
                    {
                        break;
                    }
                    currentNode.AppendChild(person);
                }
            }

            return doc;
        }

        static void SaveHtml(HtmlDocument doc)
        {
            using (StreamWriter writer = File.CreateText("test.html"))
            {
                doc.Save(writer);
            }
        }

        static int GetIlvl(XElement xml)
        {
            var ilvlString = xml.Descendants().Attributes().FirstOrDefault(a => a.Parent.Name.LocalName == "ilvl");
            var ilvl = -1;
            return int.TryParse(ilvlString?.Value, out ilvl) ? ilvl : -1;
        }
    }
}
