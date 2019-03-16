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


namespace AdvancedLinq
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
                    report.Run(DataPath + "Advanced Linq.template.xlsx", saveFileDialog1.FileName);

                    if (MessageBox.Show("Do you want to open the generated file?", "Confirm", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        Process.Start(saveFileDialog1.FileName);
                    }
                }
            }
        }

        private void LoadTables(FlexCelReport report)
        {
            var Countries = new List<Country>();
            Countries.Add(new Country("China",
                          new People(1384688986),
                          new Geography(
                              new Area(270550, 9326410))));

            var country = Countries[Countries.Count - 1];
            country.People.Language.Add(new Language(
                new LanguageName("Md", "Mandarin"),
                new LanguageSpeakers(0, 66.2)));

            country.People.Language.Add(new Language(
                new LanguageName("Yue", "Yue"),
                new LanguageSpeakers(0, 4.9)));

            country.People.Language.Add(new Language(
                new LanguageName("Wu", "Wu"),
                new LanguageSpeakers(0, 6.1)));

            country.People.Language.Add(new Language(
                new LanguageName("Mb", "Minbei"),
                new LanguageSpeakers(0, 6.2)));

            country.People.Language.Add(new Language(
                new LanguageName("Mn", "Minnan"),
                new LanguageSpeakers(0, 5.2)));

            country.People.Language.Add(new Language(
                new LanguageName("Xi", "Xiang"),
                new LanguageSpeakers(0, 3.0)));

            country.People.Language.Add(new Language(
                new LanguageName("Gan", "Gan"),
                new LanguageSpeakers(0, 4.0)));


            Countries.Add(new Country("India",
                          new People(1296834042),
                          new Geography(
                              new Area(314070, 2973193))));

            country = Countries[Countries.Count - 1];
            country.People.Language.Add(new Language(
                new LanguageName("Hi", "Hindi"),
                new LanguageSpeakers(0, 43.6)));

            country.People.Language.Add(new Language(
                new LanguageName("Bg", "Bengali"),
                new LanguageSpeakers(0, 8)));

            country.People.Language.Add(new Language(
                new LanguageName("Ma", "Marath"),
                new LanguageSpeakers(0, 6.9)));

            country.People.Language.Add(new Language(
                new LanguageName("Te", "Telugu"),
                new LanguageSpeakers(0, 6.7)));

            country.People.Language.Add(new Language(
                new LanguageName("Ta", "Tamil"),
                new LanguageSpeakers(0, 5.7)));

            country.People.Language.Add(new Language(
                new LanguageName("Gu", "Gujarati"),
                new LanguageSpeakers(0, 4.6)));

            country.People.Language.Add(new Language(
                new LanguageName("Ur", "Urdu"),
                new LanguageSpeakers(0, 4.2)));

            country.People.Language.Add(new Language(
                new LanguageName("Ka", "Kannada"),
                new LanguageSpeakers(0, 3.6)));

            country.People.Language.Add(new Language(
                new LanguageName("Od", "Odia"),
                new LanguageSpeakers(0, 3.1)));

            country.People.Language.Add(new Language(
                new LanguageName("Ma", "Malayalam"),
                new LanguageSpeakers(0, 2.9)));

            country.People.Language.Add(new Language(
                new LanguageName("Pu", "Punjabi"),
                new LanguageSpeakers(0, 2.7)));

            country.People.Language.Add(new Language(
                new LanguageName("As", "Assamese"),
                new LanguageSpeakers(0, 1.3)));

            country.People.Language.Add(new Language(
                new LanguageName("Mi", "Maithili"),
                new LanguageSpeakers(0, 1.1)));

            country.People.Language.Add(new Language(
                new LanguageName("O", "Other"),
                new LanguageSpeakers(0, 5.6)));


            Countries.Add(new Country("United States",
                          new People(329256465),
                          new Geography(
                              new Area(685924, 9147593))));

            country = Countries[Countries.Count - 1];
            country.People.Language.Add(new Language(
                new LanguageName("En", "English"),
                new LanguageSpeakers(0, 78.2)));

            country.People.Language.Add(new Language(
                new LanguageName("Sp", "Spanish"),
                new LanguageSpeakers(0, 13.4)));

            country.People.Language.Add(new Language(
                new LanguageName("Ch", "Chinese"),
                new LanguageSpeakers(0, 1.1)));

            country.People.Language.Add(new Language(
                new LanguageName("O", "Other"),
                new LanguageSpeakers(0, 7.3)));

            report.AddTable("country", Countries );
        }

        private void btnCancel_Click(object sender, System.EventArgs e)
        {
            Close();
        }
    }

    public class Country
    {
        public string Name { get; private set; }

        public People People { get; set; }
        public Geography Geography { get; set; }

        public Country(string name, People people, Geography geography)
        {
            this.Name = name;
            this.People = people;
            this.Geography = geography;
        }

    }

    public class Geography
    {
        public Area Area { get; private set; } 
        
        public Geography(Area area)
        {
            this.Area = area;
        }
    }

    public class Area
    {
        public int Total { get { return Water + Land; } }
        public int Water { get; private set; }
        public int Land { get; private set; }

        public Area(int water, int land)
        {
            this.Water = water;
            this.Land = land;
        }
    }

    public class People
    {
        public int Population { get; private set; }
        public List<Language> Language { get; private set; }

        public People(int population)
        {
            this.Population = population;
            this.Language = new List<Language>();
        }
    }

    public class Language
    {
        public LanguageName Name { get; private set; }
        public LanguageSpeakers Speakers { get; private set; }

        public Language(LanguageName name, LanguageSpeakers speakers)
        {
            this.Name = name;
            this.Speakers = speakers;
        }

    }

    public class LanguageName
    {
        public string ShortName { get; private set; }
        public string LongName { get; private set; }

        public LanguageName(string shortName, string longName)
        {
            this.ShortName = shortName;
            this.LongName = longName;
        }
    }

    public class LanguageSpeakers
    {
        public int AbsoluteNumber { get; private set; }
        public double Percent { get; private set; }

        public LanguageSpeakers(int absoluteNumber, double percent)
        {
            this.AbsoluteNumber = absoluteNumber;
            this.Percent = percent / 100.0;
        }
    }
}
