using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace VirtualMode
{
    public partial class SheetSelectorForm: Form
    {
        public SheetSelectorForm()
        {
            InitializeComponent();
        }

        public SheetSelectorForm(string[] SheetNames)
            : this()
        {
            foreach (string s in SheetNames)
            {
                SheetList.Items.Add(s);
            }

            SheetList.SelectedIndex = 0;
        }

        internal bool Execute()
        {
            return ShowDialog() == DialogResult.OK;
        }

        public string SelectedSheet
        {
            get
            {
                return Convert.ToString(SheetList.SelectedItem);
            }
        }

        public int SelectedSheetIndex
        {
            get
            {
                return SheetList.SelectedIndex;
            }
        }

        private void SheetList_DoubleClick(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
        }
    }
}
