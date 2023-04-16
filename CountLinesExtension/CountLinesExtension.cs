using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using SharpShell.Attributes;
using SharpShell.SharpContextMenu;

namespace CountLinesExtension
{
    /// <summary>
    /// The CountLinesExtensions is an example shell context menu extension,
    /// implemented with SharpShell. It adds the command 'Count Lines' to text
    /// files.
    /// </summary>
    [ComVisible(true)]
    [COMServerAssociation(AssociationType.ClassOfExtension, ".txt")]
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
            //  We always show the menu.
            return true;
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
            var itemCountLines = new ToolStripMenuItem
            {
                Text = "Konvertuj XML u PDF..."
                // Image = Properties.Resources.convert_tiny
            };

            //  When we click, we'll count the lines.
            itemCountLines.Click += (sender, args) => ConvertXMLtoPDF();

            //  Add the item to the context menu.
            menu.Items.Add(itemCountLines);

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
                var fileNameNoExt = Path.GetFileName(filePath).Split('.')[0];
                var pdfContent = PdfCreator.FindPdfContent(filePath, "env:DocumentPdf");
                PdfCreator.OpenPdf(fileNameNoExt, pdfContent);
            }
        }
    }
}