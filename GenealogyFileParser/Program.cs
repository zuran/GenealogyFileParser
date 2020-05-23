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

                    BuildTree(xml);




                    //var check = xml.Elements().Select(t => t.Name).ToList();
                    // Console.WriteLine(xml.ToString());
                }
            }


        }

        static XElement BuildTree(XElement xml)
        {
            XElement root = new XElement("root");
            var elements = xml.Elements().ToList();

            HtmlDocument doc = new HtmlDocument();
            var body = doc.CreateElement("body");
            doc.DocumentNode.AppendChild(body);

            var currentIlvl = -2;
            var currentNode = body;

            for(int i = 0; i < elements.Count; i++)
            {
                var element = elements[i];
                var ilvl = GetIlvl(element);
                if(element.Value.Trim() == "")
                {
                    continue;
                }

                if(ilvl > currentIlvl) // start a new list
                {
                    var list = currentNode.AppendChild(HtmlNode.CreateNode("<ol></ol>"));
                    list.AppendChild(HtmlNode.CreateNode("<li>" + element.Value + "</li>"));
                    currentIlvl = ilvl;
                    currentNode = list;
                } else if(ilvl == currentIlvl) // append to this list
                {
                    currentNode.AppendChild(HtmlNode.CreateNode("<li>" + element.Value + "</li>"));
                } else // traverse the tree up to the correct level
                {
                    for (; currentIlvl > ilvl; currentIlvl--)
                    {
                        currentNode = currentNode.ParentNode;
                    }
                    currentNode.AppendChild(HtmlNode.CreateNode("<li>" + element.Value + "</li>"));
                }
            }

            //Console.WriteLine(body.InnerHtml);

            using (StreamWriter writer = File.CreateText("test.html"))
            {
                doc.Save(writer);
            }
                

            //foreach(XElement element in xml.Elements())
            //{
            //    //element.Attributes().Select(n => { Console.WriteLine(n.Name); return ""; }).ToList() ;

            //    //Console.WriteLine(element.ToString());
            //    var ilvl = GetIlvl(element);

            //    for(int i = 0; i < ilvl; i++)
            //    {
            //        Console.Write("  ");
            //    }
            //    Console.WriteLine(element.Value.Substring(0, Math.Min(30, element.Value.Length)));
            //    //break;
            //}

            foreach (XElement element in root.Elements())
            {
                //
            }

            return xml;
        }

        static int GetIlvl(XElement xml)
        {
            var ilvlString = xml.Descendants().Attributes().FirstOrDefault(a => a.Parent.Name.LocalName == "ilvl");
            var ilvl = -1;
            return int.TryParse(ilvlString?.Value, out ilvl) ? ilvl : -1;
        }
    }
}
