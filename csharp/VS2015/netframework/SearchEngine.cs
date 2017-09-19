using System;
using System.IO;
using System.Data;
using System.Collections.Generic;
using System.Text;
using System.Globalization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Diagnostics;
using System.Reflection;

using FlexCel.Core;
using FlexCel.XlsAdapter;

namespace MainDemo
{
    /// <summary>
    /// A small search engine for finding information in the always-growing number of demos.
    /// This is not production-quality code, just a fast and dirty implementation, but feel free to use
    /// this code in your own projects. 
    /// </summary>
    public class SearchEngine
    {
        #region Private data

        private string MainPath;
        private Exception FMainException;
        private bool FInitialized;

        private DataTable WordTable;
        private DataView WordView;

        private const string DataFile = "keywords.dat";
        private const string KeywordTable = "Keywords";
        private const string WordColumn = "word";
        private const string ModuleColumn = "module";
        private readonly char[] WordDelim = { ' ', '.', '/', '\\' };
        #endregion

        #region Constructor And Indexing
        public SearchEngine(string aMainPath)
        {
            MainPath = aMainPath;
            FInitialized = false;
        }

        public void Index()
        {
            try
            {
                string ModulesPath = Path.GetFullPath(Path.Combine(Path.Combine(Path.Combine(MainPath, ".."), ".."), "Modules"));
                string ConfigFile = Path.Combine(MainPath, DataFile);

                try
                {
                    WordTable = new DataTable(KeywordTable);
                    WordTable.PrimaryKey = new DataColumn[] { WordTable.Columns.Add(WordColumn, typeof(string)) };
                    WordTable.Columns.Add(ModuleColumn, typeof(ModuleList));

                    bool Loaded = false;
                    if (File.Exists(ConfigFile))
                    {
                        Loaded = LoadData(ConfigFile);
                    }

                    if (!Loaded)
                    {
                        Crawl(ModulesPath);
                        SaveData(ConfigFile);
                    }

                    WordView = new DataView(WordTable);
                }
                catch
                {
                    File.Delete(ConfigFile);
                    throw;
                }

                FInitialized = true;
            }
            catch (Exception ex) //this method is designed to run in a thread, so we will not pass exceptions.
            {
                FMainException = ex;
            }
        }

        internal Exception MainException
        {
            get
            {
                return FMainException;
            }
        }

        internal bool Initialized { get { return FInitialized; } }

        #endregion

        #region Search interface

        public Dictionary<string, string> Search(string words)
        {
            string[] w = words.Split(WordDelim);

            Dictionary<string, string> Result = null;

            foreach (string s in w)
            {
                string s1 = s.Trim().ToUpper();
                if (s1.Length <= 0) continue;
                s1 = s1.Replace("'", ""); //Avoid escape inside the like expression.
                string filter = WordColumn + " like '%" + s1 + "%'";
                WordView.RowFilter = filter;

                if (WordView.Count > 100) continue; //Do not bother filtering by this keyword, too many entries.


                Dictionary<string, string> WordModules = new Dictionary<string, string>();
                for (int i = 0; i < WordView.Count; i++)
                {
                    DataRowView dv = WordView[i];
                    object value = dv[ModuleColumn];

                    Dictionary<string, string> ht = (Dictionary<string, string>)value;
                    foreach (string module in ht.Keys)
                    {
                        WordModules[module] = module;
                    }
                }

                if (Result == null) Result = WordModules;
                else //"And" words together.
                {
                    string[] keys = new string[Result.Keys.Count];
                    Result.Keys.CopyTo(keys, 0);
                    foreach (string key in keys)
                    {
                        if (!WordModules.ContainsKey(key)) Result.Remove(key);
                    }
                }

                if (Result.Count == 0) return Result; //no need to keep on filtering.
            }
            return Result;

        }
        #endregion

        #region Implementation
        private void Crawl(string RelativePath)
        {
            DirectoryInfo Parent = new DirectoryInfo(RelativePath);
            foreach (DirectoryInfo Child in Parent.GetDirectories())
            {
                Crawl(Child.FullName);
            }

            foreach (FileInfo file in Parent.GetFiles("*.rtf"))
            {
                AddRtfFile(file.FullName);
            }

            foreach (FileInfo file in Parent.GetFiles("*.cs"))
            {
                AddTxtFile(file.FullName);
            }

            foreach (FileInfo file in Parent.GetFiles("*.vb"))
            {
                AddTxtFile(file.FullName);
            }

            foreach (FileInfo file in Parent.GetFiles("*.pas"))
            {
                AddTxtFile(file.FullName);

            }
            foreach (FileInfo file in Parent.GetFiles("*.vb"))
            {
                AddTxtFile(file.FullName);
            }

            foreach (FileInfo file in Parent.GetFiles("*.txt"))
            {
                AddTxtFile(file.FullName);
            }

            foreach (FileInfo file in Parent.GetFiles("*.xls"))
            {
                AddXlsFile(file.FullName);
            }

        }

        private void AddRtfFile(string FileName)
        {
            //This implements a *really* basic parser, but it doesn't matter for this use.
            using (StreamReader sr = new StreamReader(FileName))
            {
                int key;
                while ((key = sr.Read()) > 0)
                {
                    if (key == '}' || key == '{') continue;
                    if (key == '\\')
                    {
                        SkipCommand(sr);
                        continue;
                    }

                    if (Char.IsLetterOrDigit((char)key))
                    {
                        GetWord((char)key, sr, FileName);
                        continue;
                    }
                }
            }
        }

        //This method is too naive, it will ignore parameters. This means that in text like:
        // "\fcharset0 Garamond;" Garamond will be considered a word. Again, not a big problem here.
        // What matters more here is speed, and this is faster than using a RichTextBox
        private void SkipCommand(StreamReader sr)
        {
            int key;
            while ((key = sr.Read()) > 0)
            {
                if (key == ' ') return;
            }
        }

        private void GetWord(char first, StreamReader sr, string FileName)
        {
            int key;
            StringBuilder sb = new StringBuilder();
            sb.Append(first);
            while ((key = sr.Read()) > 0)
            {
                if (Char.IsLetterOrDigit((char)key))
                {
                    sb.Append((char)key);
                }
                else
                {
                    AddWord(sb.ToString(), FileName);
                    return;
                }

            }
        }

        private void AddTxtFile(string FileName)
        {
            using (StreamReader sr = new StreamReader(FileName))
            {
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    string[] words = line.Split(WordDelim);
                    foreach (string word in words)
                    {
                        AddWord(word, FileName);
                    }
                }
            }
        }

        private void AddXlsFile(string FileName)
        {
            XlsFile xls = new XlsFile();
            try
            {
                xls.Open(FileName);
            }
            catch (FlexCelXlsAdapterException ex)
            {
                if (ex.ErrorCode == XlsErr.ErrInvalidPassword) return;
                throw;
            }

            for (int sheet = 1; sheet <= xls.SheetCount; sheet++)
            {
                xls.ActiveSheet = sheet;
                for (int r = 1; r <= xls.RowCount; r++)
                {
                    for (int cindex = 1; cindex <= xls.ColCountInRow(r); cindex++)
                    {
                        int XF = -1;
                        object cell = xls.GetCellValueIndexed(r, cindex, ref XF);
                        AddWord(Convert.ToString(cell), FileName);  //we could use TFlxNumberFormat.FormatValue() here, but we don't care about formatted values for searching.
                    }
                }
            }

        }

        private void AddWord(string word, string module)
        {
            string Trimmed = word.Trim().ToUpper();
            if (Trimmed.Length <= 2) return;  //Filter small words. we need 3, for things like .net or asp, or com.

            DataRow dr = WordTable.Rows.Find(Trimmed);
            if (dr == null)
            {
                ModuleList Mod = new ModuleList();
                Mod.Add(module, module);
                WordTable.Rows.Add(new object[] { Trimmed, Mod });
            }
            else
            {
                ((ModuleList)dr[ModuleColumn])[module] = module;
            }

        }


        #endregion

        #region Save Dataset

        private string FlexCelVersion()
        {
            Assembly asm = Assembly.GetAssembly(typeof(XlsFile));
            return asm.GetName().Version.ToString();
        }

        private void SaveData(string filename)
        {
            using (FileStream fs = new FileStream(filename, FileMode.Create))
            {
                BinaryFormatter bin = new BinaryFormatter();
                bin.Serialize(fs, FlexCelVersion());
                bin.Serialize(fs, WordTable.Rows.Count);
                foreach (DataRow dr in WordTable.Rows)
                {
                    bin.Serialize(fs, dr[WordColumn]);

                    ModuleList list = (ModuleList)dr[ModuleColumn];
                    bin.Serialize(fs, list.Count);
                    foreach (string key in list.Keys)
                    {
                        bin.Serialize(fs, key);
                    }

                }
            }
        }

        private bool LoadData(string filename)
        {
            try
            {
                using (FileStream fs = new FileStream(filename, FileMode.Open))
                {
                    if (fs.Length <= 0) return false;
                    BinaryFormatter bin = new BinaryFormatter();
                    string Version = (string)bin.Deserialize(fs);
                    if (Version != FlexCelVersion()) return false; //if this is a new version, regenerate the index.

                    int Entries = (int)bin.Deserialize(fs);

                    for (int i = 0; i < Entries; i++)
                    {
                        string word = (string)bin.Deserialize(fs);

                        int modulecount = (int)bin.Deserialize(fs);
                        ModuleList list = new ModuleList(modulecount);
                        for (int k = 0; k < modulecount; k++)
                        {
                            string m = (string)bin.Deserialize(fs);
                            list.Add(m, m);
                        }

                        WordTable.Rows.Add(new object[] { word, list });

                    }

                    Debug.Assert(Entries == WordTable.Rows.Count);
                }
            }
            catch
            {
                return false;
            }
            return true;
        }

        #endregion
    }

    class ModuleList: Dictionary<string, string>
    {
        public ModuleList() :
            base(StringComparer.Create(CultureInfo.InvariantCulture, true))
        {
        }

        public ModuleList(int Capacity) :
            base(Capacity, StringComparer.Create(CultureInfo.InvariantCulture, true))
        {
        }

    }
}
