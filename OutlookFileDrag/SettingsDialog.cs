using OutlookFileDrag.Properties;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OutlookFileDrag
{
    public partial class SettingsDialog : Form
    {
        List<TextString> lst = new List<TextString>();

        public SettingsDialog()
        {
            InitializeComponent();
        }

        private void Settings_Load(object sender, EventArgs e)
        {
            lst.Clear();
            foreach(string url in Settings.Default.dropUrlAccept)
            {
                lst.Add(new TextString(url));
            }
            BindingSource bs = new BindingSource();
            // bind to the new wrapper class
            bs.DataSource = lst;
            this.dataGridView1.DataSource = bs;
        }

        private void SaveClicked(object sender, EventArgs e)
        {
            System.Collections.Specialized.StringCollection urlStrings = new System.Collections.Specialized.StringCollection();
            foreach(TextString textString in lst)
            {
                if (!String.IsNullOrWhiteSpace(textString.Text))
                {
                    urlStrings.Add(textString.Text);
                }
            }
            Settings.Default.dropUrlAccept = urlStrings;
            Settings.Default.Save();
            this.Dispose();
        }

        private void dataGridView1_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView1.BeginEdit(true);
        }

        private void CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //if click is on new row or header row
            if (e.RowIndex == dataGridView1.NewRowIndex || e.RowIndex < 0)
                return;

            //Check if click is on specific column 
            if (e.ColumnIndex == dataGridView1.Columns["deleteButtonColumn"].Index)
            {
                ((BindingSource)this.dataGridView1.DataSource).RemoveAt(e.RowIndex);
            }
        }

        private void cancelClicked(object sender, EventArgs e)
        {
            this.Dispose();
        }
    }
}
