using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using PdfSharp.Pdf.Advanced;
using SharpShell.Attributes;
using SharpShell.SharpContextMenu;

namespace CountLinesExtension
{
    [ComVisible(true)]
    [Guid("C8440364-6AA7-4A32-A8FA-9E498C24C6C9")]
    [COMServerAssociation(AssociationType.AllFiles)]
    public class CountLinesExtension : SharpContextMenu
    {
        /// <summary>
        /// Determines whether this instance can a shell context show menu, given the specified selected file list.
        /// </summary>
        /// <returns>
        ///   <c>true</c> if this instance should show a shell context menu for the specified file list; otherwise, <c>false</c>.
        /// </returns>
        protected override bool CanShowMenu()
        {
            //  We only show the menu if any selected file ends with .xml.
            return SelectedItemPaths?.Any(path => path.EndsWith(".xml")) ?? false;
        }

        /// <summary>
        /// Creates the context menu. This can be a single menu item or a tree of them.
        /// </summary>
        /// <returns>
        /// The context menu for the shell context menu.
        /// </returns>
        protected override ContextMenuStrip CreateMenu()
        {
            //  Create the menu strip.
            var menu = new ContextMenuStrip();

            //  Create a 'count lines' item.
            var toolStripMenuItem = new ToolStripMenuItem
            {
                Text = "Konvertuj XML u PDF..."
                // Image = Properties.Resources.convert_tiny
            };

            //  When we click, we'll count the lines.
            toolStripMenuItem.Click += (sender, args) => ConvertXMLtoPDF();

            //  Add the item to the context menu.
            menu.Items.Add(toolStripMenuItem);

            //  Return the menu.
            return menu;
        }

        /// <summary>
        /// Converts XML to PDF
        /// </summary>
        private void ConvertXMLtoPDF()
        {
            foreach (var filePath in SelectedItemPaths)
            {
                try
                {
                    var fileNameNoExt = Path.GetFileName(filePath).Split('.')[0];
                    var pdfContent = PdfCreator.FindPdfContent(filePath, "env:DocumentPdf");

                    // TODO: PdfCreator.OpenPdf(fileNameNoExt, pdfContent);

                    Log($"Converted XML to PDF. File: [{filePath}] | Content (Length): [{pdfContent?.Length}]");
                }
                catch (System.Exception exception)
                {
                    LogError($"Error Converting XML to PDF. File: [{filePath}]", exception);
                }
            }
        }
    }
}