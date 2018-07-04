using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace Shift.Companion
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            ThisAddIn.mainform.Show();
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Form2 myform2 = new Form2();
            myform2.Show();

        }
    }
}
