using System;
using Android.App;
using Android.Content;
using Android.Runtime;
using Android.Views;
using Android.Widget;
using Android.OS;
using FlexCel.XlsAdapter;
using FlexCel.Core;
using System.IO;

namespace FlexCalc
{

    [Activity (Label = "FlexCalc", MainLauncher = true)]
    public class MainActivity : Activity
    {
        XlsFile xls;
        string ConfigFile = Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal), "config.xls");

        protected override void OnCreate(Bundle bundle)
        {
            string[] Predefined = new string[]
            {
                "5", "=A1 * 3 + 7", "=Sum(A1, A2)*9", "=Sin(a1) + cos(a2)", "=Average(a1:a4)",
                "", "", "", "", "", "", "", "", "", "", ""
            };


            base.OnCreate(bundle);

            bool Restoring = false;
            xls = new XlsFile(true); 
            if (File.Exists(ConfigFile)) 
            {
                try
                {
                    xls.Open(ConfigFile); 
                    Restoring = true;
                }
                catch 
                {
                    //if the file is corrupt, we'll just ignore it.
                    //Restoring will be false, and we will create a new file.
                }
            }

            if (!Restoring)
            {
                xls.NewFile(1);            

                for (int k = 0; k < Predefined.Length; k++)
                {
                    xls.SetCellFromString(k + 1, 1, Predefined [k]); //Initialize the grid with something so users know what they have to do.
                }
            }
            xls.Recalc();

            TextView[] Results = new TextView[xls.RowCount];

            var Layout = new TableLayout(this);
            for (int i = 0; i < Results.Length; i++)
            {
                var Row = new TableRow(Layout.Context);

                var ColHeading = new TextView(Row.Context);
                ColHeading.Text = new TCellAddress(i + 1, 1).CellRef;
                ColHeading.Gravity = GravityFlags.Left;

                EditText CellValue = new EditText(Row.Context);
                CellValue.Gravity = GravityFlags.Fill;
                CellValue.Text = GetCellOrFormula(i + 1);
                CellValue.Tag = i;
               

                CellValue.AfterTextChanged += (object sender, Android.Text.AfterTextChangedEventArgs e) => 
                {
                    int z = (int)(sender as EditText).Tag;
                    xls.SetCellFromString(z + 1, 1, (sender as EditText).Text);
                    xls.Recalc();
                    for (int k = 0; k < Results.Length; k++) 
                    {
                        Results[k].Text = xls.GetStringFromCell(k + 1, 1); 
                    }
                };

                Results[i] = new TextView(Row.Context);
                Results[i].Gravity = GravityFlags.Right;
                Results[i].Text = xls.GetStringFromCell(i + 1, 1);

                Row.AddView(ColHeading);
                Row.AddView(CellValue);
                Row.AddView(Results[i]);

                Layout.AddView(Row);
                SetContentView(Layout);

            }
        }

        string GetCellOrFormula(int row)
        {
            object cell = xls.GetCellValue(row, 1);
            if (cell == null)
                return "";
            TFormula fmla = (cell as TFormula);
            if (fmla != null)
                return fmla.Text;

            return Convert.ToString(cell);
        }    
    
       protected override void OnPause()
        {
            if (xls != null)
            {
                xls.Save(ConfigFile);
            }
            base.OnPause();
        }
    }
}


