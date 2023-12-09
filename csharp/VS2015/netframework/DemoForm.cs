using System;
using System.Drawing;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;
using System.Reflection;
using System.Globalization;
using System.Resources;
using System.Threading;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace MainDemo
{
    /// <summary>
    /// Demo Browser for FlexCel. This application will run all the other demos available.
    /// </summary>
    public partial class DemoForm: System.Windows.Forms.Form
    {
        private RichTextBox50 descriptionText;

        public DemoForm()
        {
            InitializeComponent();
            CreateBoxDescription();
            ResizeToolbar(mainToolbar);

            CleanSearchbox();
            LoadModules();
            FilterTree(null);
        }

        private void CreateBoxDescription()
        {
            //Until .NET 4.7, the rich text box would show hyperlinks badly.
            this.descriptionText = new RichTextBox50();
            this.descriptionText.BackColor = System.Drawing.SystemColors.Window;
            this.descriptionText.Dock = System.Windows.Forms.DockStyle.Fill;
            this.descriptionText.Location = new System.Drawing.Point(3, 24);
            this.descriptionText.Name = "descriptionText";
            this.descriptionText.ReadOnly = true;
            this.descriptionText.TabIndex = 1;
            this.descriptionText.Text = "";
            this.descriptionText.LinkClicked += new System.Windows.Forms.LinkClickedEventHandler(this.descriptionText_LinkClicked);
            this.descriptionText.Parent = panel3;

        }

        private void ResizeToolbar(ToolStrip toolbar)
        {

            using (Graphics gr = CreateGraphics())
            {
                double xFactor = gr.DpiX / 96.0;
                double yFactor = gr.DpiY / 96.0;
                toolbar.ImageScalingSize = new Size((int)(24 * xFactor), (int)(24 * yFactor));
                toolbar.Width = 0; //force a recalc of the buttons.
            }

        }


        #region Global constants.
        private readonly string PathToExe = Path.Combine("bin", "Debug");
        private readonly string ExtLaunch = ".xls";
        private readonly string ExtTemplate = ".template.xls";
        private readonly string ExtTemplateX = ".template.xlsx";
        private readonly string ExtCsProject = ".csproj";
        private readonly string ExtVbProject = ".vbproj";
        private readonly string ExtPrismProject = ".oxygene";
        private readonly string ExtLocation = ".location.txt";

        private SearchEngine Finder;
        private TTreeNode MainNode;
        #endregion

        private static void LaunchFile(string f)
        {
            if (f != null)
            {
                using (Process p = new Process())
                {               
                    p.StartInfo.FileName = f;
                    p.StartInfo.UseShellExecute = true;
                    p.Start();
                }              
            }            
        }

        private void LoadModules()
        {
            string MainPath = Path.GetFullPath(Path.Combine(Path.Combine(Path.GetDirectoryName(Application.ExecutablePath), ".."), ".."));
            MainNode = new TTreeNode(Text, Path.Combine(MainPath, "MainDemo"));

            string[] subdirectoryEntries = Directory.GetDirectories(Path.Combine(MainPath, "Modules"));
            foreach (string subdirectory in subdirectoryEntries)
                LoadModule(Path.Combine(MainPath, "Modules"), subdirectory, MainNode);
        }

        private void LoadModule(string MainPath, string modulePath, TTreeNode node)
        {
            string LinkFile = Path.Combine(modulePath, "link.txt");
            if (File.Exists(LinkFile))
            {
                using (StreamReader sr = new StreamReader(LinkFile))
                {
                    string RelPath = sr.ReadLine().Replace('\\', Path.DirectorySeparatorChar);
                    modulePath = Path.GetFullPath(Path.Combine(MainPath, RelPath));
                }
            }

            string moduleName = Path.GetFileName(modulePath);
            string shortModule = moduleName.Substring(moduleName.IndexOf(".") + 1);
            if (moduleName.Length < 1 || moduleName[0] == '.') return; //Do not process hidden folders.
            if (moduleName.IndexOf('.') < 1) return; //Do not process folders without the convention xx.name

            string NodePath = null;
            if (File.Exists(Path.Combine(modulePath, "README.rtf")))
            {
                NodePath = Path.Combine(modulePath, shortModule);
            }

            TTreeNode NewNode = new TTreeNode(shortModule, NodePath);
            node.Children.Add(NewNode);


            string[] subdirectoryEntries = Directory.GetDirectories(modulePath);
            foreach (string subdirectory in subdirectoryEntries)
                LoadModule(MainPath, subdirectory, NewNode);
        }

        private void modulesList_AfterSelect(object sender, System.Windows.Forms.TreeViewEventArgs e)
        {
            if (e.Node.Tag == null) descriptionText.Clear();
            else descriptionText.LoadFile(Path.Combine(Path.GetDirectoryName((string)e.Node.Tag), "README.rtf"));

            statusBar1.Text = e.Node.FullPath;

            btnRunSelected.Enabled = (HasModuleForm()) || (HasFileToLaunch(ExtLaunch) != null)
                || (HasFileToLaunch(ExtCsProject) != null) || (HasFileToLaunch(ExtVbProject) != null);

            menuRunSelected.Enabled = btnRunSelected.Enabled;

            btnViewTemplate.Enabled = HasFileToLaunch(ExtTemplate) != null || HasFileToLaunch(ExtTemplateX) != null;
            menuViewTemplate.Enabled = btnViewTemplate.Enabled;

            btnOpenProject.Enabled = HasFileToLaunch(ExtCsProject) != null || HasFileToLaunch(ExtVbProject) != null || HasFileToLaunch(ExtPrismProject) != null;
            menuOpenProject.Enabled = btnOpenProject.Enabled;
        }

        private void Exit_Click(object sender, System.EventArgs e)
        {
            Close();
        }

        private bool HasModuleForm()
        {
            Form Frm = GetModuleForm();
            if (Frm == null) return false;
            Frm.Dispose();
            return true;
        }

        private Form GetModuleForm()
        {
            string mName;
            string moduleName = GetModuleName(out mName);
            if (moduleName == null || !File.Exists(moduleName)) return null;
            Assembly assembly = Assembly.LoadFrom(moduleName);
            return (Form)assembly.CreateInstance(mName + ".mainForm");
        }

        private string GetModuleName(out string mName)
        {
            mName = null;
            if (modulesList.SelectedNode == null || modulesList.SelectedNode.Tag == null) return null;
            string mPath = Path.Combine(Path.GetDirectoryName(((string)modulesList.SelectedNode.Tag)), PathToExe);
            mName = Path.GetFileName((string)modulesList.SelectedNode.Tag);
            mName = mName.Replace(" ", "");
            return Path.GetFullPath(Path.Combine(mPath, mName + ".exe"));
        }

        private string HasFileToLaunch(string extension)
        {
            if (modulesList.SelectedNode == null || modulesList.SelectedNode.Tag == null) return null;
            string mPath = Path.GetDirectoryName(((string)modulesList.SelectedNode.Tag));
            string mName = Path.GetFileName((string)modulesList.SelectedNode.Tag);

            if (File.Exists(Path.Combine(mPath, extension.Substring(1) + ExtLocation)))
            {
                using (StreamReader sr = new StreamReader(Path.Combine(mPath, extension.Substring(1) + ExtLocation)))
                {
                    return mPath + sr.ReadLine();
                }
            }
            if (File.Exists(Path.Combine(mPath, mName + extension))) return Path.Combine(mPath, mName + extension);
            return null;
        }

        private bool IgnoreInMainDemo()
        {
            return IgnoreInMainDemoMessage() != null;
        }

        private string IgnoreInMainDemoMessage()
        {
            string IgnoreFile = HasFileToLaunch(".IgnoreInMainDemo.txt");
            if (String.IsNullOrEmpty(IgnoreFile)) return null;
            return File.ReadAllText(IgnoreFile);
        }


        private void RunSelected_Click(object sender, System.EventArgs e)
        {
            if (IgnoreInMainDemo())
            {
                MessageBox.Show(IgnoreInMainDemoMessage());
                return;
            }

            TryToCompileProject();
            Form frm = GetModuleForm();
            try
            {
                if (frm == null)
                {
                    string f = HasFileToLaunch(ExtLaunch);
                    if (f != null)
                    {
                        LaunchFile(f);
                    }
                    return;
                }
                Type tfrm = frm.GetType();
                MethodInfo autorun = tfrm.GetMethod("AutoRun");
                if (autorun != null)
                {
                    autorun.Invoke(frm, new object[0]);
                    return;
                }

                frm.StartPosition = FormStartPosition.CenterParent;
                frm.ShowInTaskbar = false;
                frm.ShowDialog();
            }
            finally
            {
                if (frm != null) frm.Dispose();
            }
        }

        private void TryToCompileProject()
        {
            string mName;
            string moduleName = GetModuleName(out mName);
            if (moduleName != null && File.Exists(moduleName)) return;


            string CsProj = HasFileToLaunch(ExtCsProject);
            if (CsProj != null)
            {
                Builder.Build(CsProj);
            }

            string VbProj = HasFileToLaunch(ExtVbProject);
            if (VbProj != null)
            {
                Builder.Build(VbProj);
            }

        }

        private void ViewTemplate_Click(object sender, System.EventArgs e)
        {
            string f = HasFileToLaunch(ExtTemplateX);
            if (f != null)
            {
                LaunchFile(f);
                return;
            }

            f = HasFileToLaunch(ExtTemplate);
            if (f != null)
            {
                LaunchFile(f);
            }

        }

        private void About_Click(object sender, EventArgs e)
        {
            using (AboutForm af = new AboutForm())
            {
                af.ShowDialog();
            }
        }

        private void descriptionText_LinkClicked(object sender, System.Windows.Forms.LinkClickedEventArgs e)
        {
            try
            {
                LaunchFile(e.LinkText);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnOpenProject_Click(object sender, System.EventArgs e)
        {
            string f = HasFileToLaunch(ExtCsProject);
            if (f != null)
            {
                LaunchFile(f);
                return;
            }

            f = HasFileToLaunch(ExtVbProject);
            if (f != null)
            {
                LaunchFile(f);
                return;
            }

            f = HasFileToLaunch(ExtPrismProject);
            if (f != null)
            {
                LaunchFile(f);
                return;
            }

        }

        private void btnOpenFolder_Click(object sender, EventArgs e)
        {
            if (modulesList.SelectedNode == null || modulesList.SelectedNode.Tag == null) return;
            string f = Path.GetDirectoryName((string)(modulesList.SelectedNode.Tag));
            LaunchFile(f);

        }

        private void sdSearch_TextChanged(object sender, System.EventArgs e)
        {
            if (sdSearch.Tag != null) return;

            if (Finder == null || !Finder.Initialized)
            {
                Finder = new SearchEngine(Path.GetDirectoryName(Application.ExecutablePath));
                Thread SearchThread = new Thread(new ThreadStart(Finder.Index));
                SearchThread.Start();

                using (ProgressDialog Pg = new ProgressDialog())
                {
                    Pg.ShowProgress(SearchThread);
                    if (Finder != null && Finder.MainException != null)
                    {
                        Exception ex = Finder.MainException;
                        Finder = null;
                        throw ex;
                    }
                }
            }

            if (String.Compare(sdSearch.Text, "why?", true, CultureInfo.InvariantCulture) == 0) Answer();

            Dictionary<string, string> FoundModules = Finder.Search(sdSearch.Text);
            FilterTree(FoundModules);
        }

        private void FilterTree(Dictionary<string, string> FoundModules)
        {
            modulesList.BeginUpdate();
            try
            {
                TreeNode OldSelected = modulesList.SelectedNode;
                string OldSelectedPath = null;
                if (OldSelected != null) OldSelectedPath = Convert.ToString(OldSelected.Tag);

                modulesList.Nodes.Clear();
                TreeNode MainTreeNode = new TreeNode(MainNode.NodeName);
                MainTreeNode.Tag = MainNode.NodePath;
                modulesList.Nodes.Add(MainTreeNode);
                TreeNode NewSelected = null;
                FilterTree(FoundModules, MainNode, MainTreeNode, OldSelectedPath, ref NewSelected);
                modulesList.ExpandAll();
                if (NewSelected == null) NewSelected = MainTreeNode;
                modulesList.SelectedNode = NewSelected;
                NewSelected.EnsureVisible();
            }
            finally
            {
                modulesList.EndUpdate();
            }
        }

        private void FilterTree(Dictionary<string, string> FoundModules, TTreeNode ParentNode, TreeNode ParentTreeNode, string OldSelectedPath, ref TreeNode NewSelected)
        {
            foreach (TTreeNode node in ParentNode.Children)
            {
                if (FoundModules == null || HasKey(FoundModules, Path.GetDirectoryName(node.NodePath)))
                {
                    TreeNode NewNode = new TreeNode(node.NodeName);
                    NewNode.Tag = node.NodePath;
                    ParentTreeNode.Nodes.Add(NewNode);
                    FilterTree(FoundModules, node, NewNode, OldSelectedPath, ref NewSelected);
                    if (node.NodePath == OldSelectedPath) NewSelected = NewNode;
                }

            }
        }

        private bool HasKey(Dictionary<string, string> FoundModules, string pattern)
        {
            if (pattern == null) return false;
            foreach (string s in FoundModules.Keys)
            {
                if (s.StartsWith(pattern)) return true;
            }
            return false;
        }


        private void Answer()
        {
            string[] Answers =  {
                                    "It was not my fault. I was just following your orders.",
                                    "Because that's the way life is. Better go getting used to it.",
                                    "The answer is 42. Sometimes.",
                                    "If I told you then I would have to kill you.",
                                    "It is the user's fault",
                                    "I can only answer you after my NDA expires.",
                                    "Whatever it is, don't worry. Tomorrow we will look at it and we will laugh.",
                                    "Please give me some time to think about it.",
                                    "I could tell you, but then where would be the fun?"
                                };

            Random rnd = new Random();
            MessageBox.Show(Answers[rnd.Next(Answers.Length)]);
        }

        readonly string TxtTypeToSearch = "Type to search...";  //this isn't a nice way to show a hint, but it will work for this simple demo, without using a third party control.

        private void sdSearch_Enter(object sender, EventArgs e)
        {
            sdSearch.ForeColor = Color.Black;
            if (sdSearch.Tag != null)
            {
                sdSearch.Text = "";
                sdSearch.Tag = null;
            }
        }

        private void sdSearch_Leave(object sender, EventArgs e)
        {
            CleanSearchbox();
        }

        private void CleanSearchbox()
        {
            sdSearch.ForeColor = Color.Gray;
            if (string.IsNullOrEmpty(sdSearch.Text))
            {
                sdSearch.Tag = "e";
                sdSearch.Text = TxtTypeToSearch;
            }
            else sdSearch.Tag = null;
        }

    }


    class TTreeNode
    {
        public string NodeName;
        public string NodePath;
        public List<TTreeNode> Children;

        public TTreeNode(string aNodeName, string aNodePath)
        {
            NodeName = aNodeName;
            NodePath = aNodePath;
            Children = new List<TTreeNode>();
        }
    }

    public class RichTextBox50 : RichTextBox
    {
        //This class is not needed after .NET 4.7
        [DllImport("kernel32.dll", EntryPoint = "LoadLibraryW",
            CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr LoadLibraryW(string s_File);
        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int wMsg, IntPtr wParam, IntPtr lParam);

        public RichTextBox50()
        {
            const int EM_SETMARGINS = 211;
            IntPtr EC_LEFTMARGIN = (IntPtr)1;
            SendMessage(Handle, EM_SETMARGINS, EC_LEFTMARGIN, (IntPtr)40);

        }

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                LoadLibraryW("MsftEdit.dll");
                cp.ClassName = "RichEdit50W";
                return cp;
            }
        }


    }
}

