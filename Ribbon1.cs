using Microsoft.Office.Tools.Ribbon;
using SlideEvents.Properties;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace SlideEvents
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void editBtn_Click(object sender, RibbonControlEventArgs e)
        {
            int index = Globals.ThisAddIn.Application.ActiveWindow.View.Slide.SlideIndex;
            EditEventForm form;

            if (Globals.ThisAddIn.urlPairs.ContainsKey(index))
            {
                form = new EditEventForm(Globals.ThisAddIn.urlPairs[index]);
            }
            else
            {
                form = new EditEventForm();
            }

            if (form.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string url = form.GetUrl();

                Globals.ThisAddIn.urlPairs[index] = url;
            }
            form.Dispose();
        }
    }
}
