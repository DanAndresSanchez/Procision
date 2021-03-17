using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Procision
{
    public partial class Ribbon1
    {
        DataClense dc = new DataClense();
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            dc.clearSheet();
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            dc.centerAlign();
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            dc.formatNames();
        }
    }
}
