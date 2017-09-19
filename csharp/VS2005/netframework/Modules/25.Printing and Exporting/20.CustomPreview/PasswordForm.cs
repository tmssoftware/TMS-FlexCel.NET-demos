using System;
using System.Drawing;
using System.ComponentModel;
using System.Windows.Forms;

namespace CustomPreview
{
    /// <summary>
    /// Form for asking for a password when the file is password protected.
    /// </summary>
    public partial class PasswordForm: System.Windows.Forms.Form
    {

        public PasswordForm()
        {
            InitializeComponent();
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
