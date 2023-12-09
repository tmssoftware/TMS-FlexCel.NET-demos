using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;

using FlexCel.Core;
using FlexCel.XlsAdapter;
using System.Collections.Generic;
using System.Diagnostics;

namespace ObjectExplorer
{
    /// <summary>
    /// Object explorer.
    /// </summary>
    public partial class mainForm: System.Windows.Forms.Form
    {

        public mainForm()
        {
            InitializeComponent();
            ResizeToolbar(mainToolbar);
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


        #region Global variables
        XlsFile Xls;
        #endregion

        private void btnExit_Click(object sender, System.EventArgs e)
        {
            Close();
        }

        private void btnOpenFile_Click(object sender, System.EventArgs e)
        {
            if (openFileDialog1.ShowDialog() != DialogResult.OK) return;

            Xls = new XlsFile();
            Xls.Open(openFileDialog1.FileName);


            cbSheet.Items.Clear();
            int ActSheet = Xls.ActiveSheet;
            for (int i = 1; i <= Xls.SheetCount; i++)
            {
                Xls.ActiveSheet = i;
                cbSheet.Items.Add(Xls.SheetName);
            }
            Xls.ActiveSheet = ActSheet;
            cbSheet.SelectedIndex = ActSheet - 1;

            FillListBox();
        }

        private void FillListBox()
        {
            lblObjects.Text = openFileDialog1.FileName;
            dataGrid.DataSource = null;

            ObjTree.BeginUpdate();
            try
            {
                ObjTree.Nodes.Clear();

                for (int i = 1; i <= Xls.ObjectCount; i++)
                {
                    TShapeProperties ShapeProps = Xls.GetObjectProperties(i, true);
                    string s = "Object " + i.ToString();
                    if (ShapeProps.ShapeName != null) s = ShapeProps.ShapeName;

                    TreeNode RootNode = new TreeNode(s);
                    FillNodes(ShapeProps, RootNode);


                    ObjTree.Nodes.Add(RootNode);
                }
            }
            finally
            {
                ObjTree.EndUpdate();
            }
        }

        private void FillNodes(TShapeProperties ShapeProps, TreeNode Node)
        {
            Node.Tag = ShapeProps; //In this simple demo we will use the tag property to store the Shape properties. This is not indented for 'real' use.


            for (int i = 1; i <= ShapeProps.ChildrenCount; i++)
            {
                TShapeProperties ChildProps = ShapeProps.Children(i);
                string ShapeName = ChildProps.ShapeName;
                if (ShapeName == null) ShapeName = "Object " + i.ToString();
                TreeNode Child = new TreeNode(ShapeName);
                FillNodes(ChildProps, Child);
                Node.Nodes.Add(Child);
            }
        }

        private void btnOpen_Click(object sender, System.EventArgs e)
        {
            if (Xls == null)
            {
                MessageBox.Show("There is no open file");
                return;
            }
            using (Process p = new Process())
            {               
                p.StartInfo.FileName = Xls.ActiveFileName;
                p.StartInfo.UseShellExecute = true;
                p.Start();
            }              
        }


        private void btnSaveImage_Click(object sender, System.EventArgs e)
        {
            if (PreviewBox.Image == null)
            {
                MessageBox.Show("There is no selected image to save", "Error");
                return;
            }
            if (saveImageDialog.ShowDialog() != DialogResult.OK) return;

            PreviewBox.Image.Save(saveImageDialog.FileName);
        }

        private void btnInfo_Click(object sender, System.EventArgs e)
        {
            MessageBox.Show("Object Explorer allows to explore inside the objects in an Excel file.\n" +
                "Objects in xls files are hierachily distributed, you can have two objects grouped as a third object, " +
                "and this hierarchy is shown in the 'Objects' pane at the left. The properties for the selected object are displayed at the 'Object properties' pane.");
        }

        private void ObjTree_AfterSelect(object sender, System.Windows.Forms.TreeViewEventArgs e)
        {
            RenderNode(e.Node);
            FillProperties(e.Node);
        }


        private void RenderNode(TreeNode Node)
        {
            if (Xls == null)
            {
                PreviewBox.Image = null;
                return;
            }

            TreeNode t = Node;
            if (t == null)
            {
                PreviewBox.Image = null;
                return;
            }

            while (t.Parent != null) t = t.Parent;  //Only root level objects will be rendered.

            if (t.Index + 1 > Xls.ObjectCount)
            {
                PreviewBox.Image = null;
                return;
            }

            PreviewBox.Image = Xls.RenderObject(t.Index + 1);


        }

        private void FillProperties(TreeNode Node)
        {
            lblObjName.Text = "Name:";
            lblObjText.Text = "Text:";
            TShapeProperties Props = (TShapeProperties)Node.Tag;
            if (Props == null)
            {
                dataGrid.DataSource = null;
                return;
            }

            TShapeOptionList ShapeOptions = (Node.Tag as TShapeProperties).ShapeOptions;
            if (ShapeOptions == null)
            {
                dataGrid.DataSource = null;
                return;
            }

            lblObjName.Text = "Name: " + Props.ShapeName;
            lblObjText.Text = "Text: " + Props.Text;

            ArrayList ShapeOpts = new ArrayList();
            foreach (KeyValuePair<TShapeOption, object> opt in ShapeOptions)
            {
                ShapeOpts.Add(new KeyValue(opt.Key, ShapeOptions));
            }
            dataGrid.DataSource = ShapeOpts;
        }

        private void cbSheet_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            Xls.ActiveSheet = cbSheet.SelectedIndex + 1;
            FillListBox();
        }

        private void btnStretchPreview_Click(object sender, EventArgs e)
        {
            if (btnStretchPreview.Checked)
                PreviewBox.SizeMode = PictureBoxSizeMode.StretchImage;
            else
                PreviewBox.SizeMode = PictureBoxSizeMode.Normal;

        }

    }

    class KeyValue
    {
        private string FKey;
        private string FAs1616;
        private string FAsLong;
        private string FAsString;

        public KeyValue(TShapeOption aKey, TShapeOptionList List)
        {
            FKey = Convert.ToString(aKey);
            FAs1616 = Convert.ToString(List.As1616(aKey, 0));
            FAsLong = Convert.ToString(List.AsLong(aKey, 0));
            FAsString = List.AsUnicodeString(aKey, "");
        }

        public string Key { get { return FKey; } set { FKey = value; } }
        public string As1616 { get { return FAs1616; } set { FAs1616 = value; } }
        public string AsLong { get { return FAsLong; } set { FAsLong = value; } }
        public string AsString { get { return FAsString; } set { FAsString = value; } }

        public override string ToString()
        {
            return Key;
        }

    }
}
