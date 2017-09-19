using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;
using System.Diagnostics;
using System.Reflection;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using FlexCel.Report;
using System.Collections.Generic;
using System.Linq;


namespace Linq
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
    public partial class mainForm : System.Windows.Forms.Form
    {

        public mainForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, System.EventArgs e)
        {
            AutoRun();
        }

        public void AutoRun()
        {
            using (FlexCelReport report = new FlexCelReport(true))
            {
                LoadTables(report);

                string DataPath = Path.Combine(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), ".."), "..") + Path.DirectorySeparatorChar;

                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    report.Run(DataPath + "Linq.template.xls", saveFileDialog1.FileName);

                    if (MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        Process.Start(saveFileDialog1.FileName);
                    }
                }
            }
        }

        private void LoadTables(FlexCelReport report)
        {
            List<Categories> Categories = new List<Categories>();
            Categories Animals = new Categories("Animals");
            Animals.Elements.Add(new Elements(1, "Penguin"));
            Animals.Elements.Add(new Elements(2, "Cat"));
            Animals.Elements.Add(new Elements(3, "Unicorn"));
            Categories.Add(Animals);

            Categories Flowers = new Categories("Flowers");
            Flowers.Elements.Add(new Elements(4, "Daisy"));
            Flowers.Elements.Add(new Elements(5, "Rose"));
            Flowers.Elements.Add(new Elements(6, "Orchid"));
            Categories.Add(Flowers);

            report.AddTable("Categories", Categories );
            //We don't need to call AddTable for elements since it is already added when we add Categories.


            List<ElementName> ElementNames = new List<ElementName>();
            ElementNames.Add(new ElementName(1, "Linus"));
            ElementNames.Add(new ElementName(1, "Gerard"));
            ElementNames.Add(new ElementName(2, "Rover"));
            ElementNames.Add(new ElementName(3, "Mike"));
            ElementNames.Add(new ElementName(5, "Rosalyn"));
            ElementNames.Add(new ElementName(5, "Monica"));
            ElementNames.Add(new ElementName(6, "Lisa"));

            report.AddTable("ElementName", ElementNames);
            //ElementName doesn't have an intrinsic relationship with categories, so we will have to manually add a relationship.
            //Non intrinsic relationships should be rare, but we do it here to show how it can be done.
            report.AddRelationship("Elements", "ElementName", "ElementID", "ElementID");
        }

        private void btnCancel_Click(object sender, System.EventArgs e)
        {
            Close();
        }
    }

    public class Categories
    {
        //Public properties can be used in reports.
        public string Name { get; private set; }

        //Elements is in master-detail relationship with this element, even when we don't explicitly add a relationship. 
        //Relationship is inferred because Elements is a property of this object
        public List<Elements> Elements { get; private set; }

        public Categories(string name)
        {
            this.Name = name;
            Elements = new List<Elements>();
        }

    }

    public class Elements
    {
        //We will relate this property with the table of colors by adding a relationship.
        public int ElementID { get; private set; }

        public string Name { get; private set; }
        

        public Elements(int elementID, string name)
        {
            this.Name = name;
            this.ElementID = elementID;
        }

    }

    public class ElementName
    {
        public int ElementID { get; private set; }
        public string Name { get; private set; }

        public ElementName(int elementID, string name)
        {
            this.ElementID = elementID;
            this.Name = name;
        }
    }

}
