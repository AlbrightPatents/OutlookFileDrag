using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

namespace OutlookFileDrag
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void OpenSettingsWindow(object sender, RibbonControlEventArgs e)
        {
            var settings = new SettingsDialog();
            settings.ShowDialog();
        }
    }
}
