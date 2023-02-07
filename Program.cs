using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace PdfSplitter
{
    class Program  
    {
        //** it is for test **
        private static string InputPath => ConfigurationManager.AppSettings.Get("InputPath");
        private static string OutputPath => ConfigurationManager.AppSettings.Get("OutputPath");

        private static string Left => ConfigurationManager.AppSettings.Get("Left");

        private static string Top => ConfigurationManager.AppSettings.Get("Top");

        private static bool GeneratePdf => Convert.ToBoolean(ConfigurationManager.AppSettings.Get("GeneratePdf"));

        private const string BodySuffix = "_Body";

        private const string CosignerSuffix = "_Cosigner";

        static void Main(string[] args)
        {
            if (!Directory.Exists($"{OutputPath}PDFTemplates"))
            {
                Directory.CreateDirectory($"{OutputPath}PDFTemplates");
            }
            
            if (!Directory.Exists($"{OutputPath}TemplateSets"))
            {
                Directory.CreateDirectory($"{OutputPath}TemplateSets");
            }

            var templateFilesList = Directory.GetFiles($"{InputPath}TemplateSets");

            List<Task> listTask = new List<Task>();

            foreach (var xmlTemplate in templateFilesList)
            {
                listTask.Add(Task.Run(async () => await GenerateTemplateXml(Path.GetFileNameWithoutExtension(xmlTemplate))));
            }

            var t = Task.WhenAll(listTask);
            t.Wait();
        }

        static async Task GenerateTemplateXml(string xmlTemplateName)
        {
            Console.WriteLine(xmlTemplateName);

            XmlDocument xml = new XmlDocument();

            xml.Load($@"{InputPath}TemplateSets\{xmlTemplateName}.xml");

            var xRoot = xml.DocumentElement;

            XmlNode applicationBody = null;

            XmlNode applicationCover = null;

            var coverPageAttribute = xml.CreateAttribute("coverPage");

            XmlNode applicationCosignerNote = null;

            foreach (XmlElement xnode in xRoot.ChildNodes)
            {
                if (xnode.Name == "Application")
                {
                    applicationCover = xnode;

                    applicationBody = xnode.CloneNode(true);

                    PdfSplitter(Path.GetFileName(applicationBody.Attributes["pdf"].Value));

                    applicationBody.Attributes["pdf"].Value =
                        applicationBody.Attributes["pdf"].Value.Replace(".pdf", $"{BodySuffix}.pdf");

                    applicationBody.Attributes["template"].Value =
                        applicationBody.Attributes["template"].Value.Replace(".xml", $"{BodySuffix}.xml");

                    applicationCover.Attributes["pdf"].Value = @"PDFTemplates\CoverPage.pdf";

                    applicationCover.Attributes["template"].Value = @"PDFTemplates\CoverPage.xml";

                    applicationCosignerNote = applicationCover.CloneNode(true);

                    applicationCover.Attributes.Append(coverPageAttribute);

                    applicationCover.Attributes["coverPage"].Value = "true";
                    
                }
            }

            xRoot.InsertAfter(applicationBody, applicationCover);

            await Task.Run(() =>  xml.Save($@"{OutputPath}TemplateSets\{xmlTemplateName}.xml"));


            applicationCosignerNote.Attributes["pdf"].Value = @"PDFTemplates\CosignerNote.pdf";

            applicationCosignerNote.Attributes["template"].Value = @"PDFTemplates\CosignerNote.xml";

            xRoot.InsertAfter(applicationCosignerNote, applicationCover);

            await Task.Run(() => xml.Save($@"{OutputPath}TemplateSets\{xmlTemplateName}{CosignerSuffix}.xml"));
        }

        static async Task PdfSplitter(string pdfName)
        {

            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            PdfDocument inputDocument = PdfReader.Open($@"{InputPath}PDFTemplates\{pdfName}", PdfDocumentOpenMode.Import);

            PdfDocument outputDocument = new PdfDocument();
            outputDocument.Info.Creator = inputDocument.Info.Creator;
            outputDocument.Version = inputDocument.Version;
            outputDocument.Info.Title = inputDocument.Info.Title;

            for (int idx = 1; idx < inputDocument.PageCount; idx++)
            {
                outputDocument.AddPage(inputDocument.Pages[idx]);
            }

            if (GeneratePdf)
            {
                var fileName = Path.GetFileNameWithoutExtension(pdfName);
                var fullPath = $"{OutputPath}PDFTemplates\\{fileName}{BodySuffix}.pdf";
                await Task.Run(() => outputDocument.Save(fullPath));
            }

            await Task.Run(() => GeneratePdfXml(Path.GetFileNameWithoutExtension(pdfName), inputDocument.PageCount - 1));
        }

        static async Task GeneratePdfXml(string xmlName, int pages)
        {
            XmlDocument xml = new XmlDocument();

            xml.Load($@"{InputPath}PDFTemplates\{xmlName}.xml");

            var xRoot = xml.DocumentElement;

            var nodeToDelete = new List<XmlNode>();

            foreach (XmlElement xnode in xml.ChildNodes[1].ChildNodes)
            {
                if (xnode.Attributes["Page"].Value == "1")
                {
                    nodeToDelete.Add(xnode);
                }
                else if (xnode.Attributes["Page"].Value != "0")
                {
                    xnode.Attributes["Page"].Value = (Convert.ToInt32(xnode.Attributes["Page"].Value) - 1).ToString();
                }
            }

            foreach (var node in nodeToDelete)
            {
                node.ParentNode.RemoveChild(node);
            }

            XmlNode filed = xml.ChildNodes[1].ChildNodes[0].CloneNode(true);

            filed.Attributes["Name"].Value = $"LVPF:PagesInfo";
            filed.Attributes["Page"].Value = "0";
            filed.Attributes["DefaultValue"].Value = pages.ToString();
            filed.Attributes["SubCount"].Value = "4";
            filed.Attributes["Left"].Value = Left;
            filed.Attributes["Top"].Value = Top;

            xRoot.AppendChild(filed);

            await Task.Run(() => xml.Save($@"{OutputPath}PDFTemplates\{xmlName}{BodySuffix}.xml"));
        }
    }
}