using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp2
{
    public partial class ConfigDialog : Form
    {
        private string[] listColumn = { "A", "B", "C", "D" };
        public ConfigDialog()
        {
            InitializeComponent();

            btnOk.DialogResult = DialogResult.OK;
            btnCancel.DialogResult = DialogResult.Cancel;

            cbSheet.Items.AddRange(Global.listSheet);
            cbSheet.SelectedIndex = 0;
            cbDate.Items.AddRange(listColumn);
            cbDate.SelectedIndex = 0;
            cbName.Items.AddRange(listColumn);
            cbName.SelectedIndex = 2;
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            Global.sheetIndex = cbSheet.SelectedIndex;
            Global.colDate = cbDate.SelectedIndex + 1;
            Global.colName = cbName.SelectedIndex + 1;
        }
    }
}
