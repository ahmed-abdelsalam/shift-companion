﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;

namespace Shift.Companion
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }
        
        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            if (ThisAddIn.mainform != null)
            {
                ThisAddIn.mainform.Show();
            }
            

        
        }

        

    }
}
