using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace Makro_Starter
{
    public partial class WordRibbon
    {
        private void WordRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnOpenVBE_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.Application.ShowVisualBasicEditor = true;
        }
    }
}
