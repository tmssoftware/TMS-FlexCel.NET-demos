using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace GenericReports2
{
    /// <summary>
    /// A dialog where you can enter any SQL.
    /// </summary>
    public partial class EnterSQLDialog: System.Windows.Forms.Form
    {

        public EnterSQLDialog()
        {
            InitializeComponent();
        }

        public string SQL
        {
            get
            {
                return edSQL.Text;
            }
        }
    }
}
