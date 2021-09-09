﻿using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using FlexCel.Core;
using FlexCel.XlsAdapter;

namespace FlexCalc
{
    class DataModel
    {
        ExcelFile xls;
        bool Saving;

        public void LoadSpreadsheet(string FileName)
        {
            if (xls == null) xls = new XlsFile(true);
            try
            {
                xls.Open(FileName);
            }
            catch
            {
                xls.NewFile(1, TExcelFileFormat.v2019);
                xls.SetCellValue(1, 1, "Example");
                xls.SetCellValue(2, 1, 42);
                xls.SetCellValue(3, 1, new TFormula("=Sqrt(A2) * A2^2"));
            }

        }

        public bool Loaded
        {
            get { return xls != null; }
        }

        public string GetCellOrFormula(int row)
        {
            object cell = xls.GetCellValue(row, 1);
            if (cell == null)
                return "";
            TFormula fmla = (cell as TFormula);
            if (fmla != null)
                return fmla.Text;

            return Convert.ToString(cell);
        }

        public string GetStringFromCell(int row, int col)
        {
            return xls.GetStringFromCell(row, col);
        }

        public void SetCellFromString(int row, int col, string value)
        {
            xls.SetCellFromString(row, col, value);
        }

        internal void SaveState(string FileName)
        {
            if (Saving) return; //if 2 or more events try to save, only listen to one.
            Saving = true;
            try
            {
                xls.Save(FileName);
            }
            finally
            {
                Saving = false;
            }
        }

        public void Recalc()
        {
            xls.Recalc();
        }
    }
}
