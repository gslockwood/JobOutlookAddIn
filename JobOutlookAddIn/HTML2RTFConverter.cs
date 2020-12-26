using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace JobOutlookAddIn
{
    class HTML2RTFConverter
    {
        public string Convert(string html)
        {
            RichTextBox rtbTemp = new RichTextBox();
            WebBrowser wb = new WebBrowser();
            wb.Navigate( "about:blank" );

            wb.Document.Write( html );
            wb.Document.ExecCommand( "SelectAll", false, null );
            wb.Document.ExecCommand( "Copy", false, null );

            rtbTemp.SelectAll();
            rtbTemp.Paste();

            return rtbTemp.Rtf;

        }
    }
}
