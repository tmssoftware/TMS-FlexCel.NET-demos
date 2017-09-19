using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace EncryptedFiles
{
    /// <summary>
    /// Summary description for PasswordDialog.
    /// </summary>
    public partial class PasswordDialog: System.Windows.Forms.Form
    {

        public PasswordDialog()
        {
            //
            // Required for Windows Form Designer support
            //
            InitializeComponent();

            //
            // TODO: Add any constructor code after InitializeComponent call
            //
        }

        public string Password
        {
            get
            {
                return PasswordEdit.Text;
            }
        }
    }
}
