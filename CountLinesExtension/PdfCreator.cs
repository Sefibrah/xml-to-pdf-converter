using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Diagnostics;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf.Content.Objects;
using PdfSharp.Drawing;

namespace CountLinesExtension
{
    public class PdfCreator
    {
        public static string FindPdfContent(string filePath, string tagName)
        {
            // Load the XML file.
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(filePath);

            // Find the tag with the specified name.
            XmlNodeList nodes = xmlDoc.GetElementsByTagName(tagName);
            return nodes[0].InnerText;
        }

        public static void OpenPdf(string fileNameNoExt, string pdfContent)
        {
            byte[] pdfBytes = Convert.FromBase64String(pdfContent);
            using (var stream = new MemoryStream(pdfBytes))
            {
                var document = PdfReader.Open(stream);
                string desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                string filePath = Path.Combine(desktopFolder, fileNameNoExt + ".pdf");
                document.Save(filePath);
                System.Diagnostics.Process.Start(filePath);
            }
        }
    }
}