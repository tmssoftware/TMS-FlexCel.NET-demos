using System;
using System.Collections.Generic;
using System.Text;
using FlexCel.Core;
using System.Drawing;

namespace VirtualMode
{
    //A simple cell reader that will get the values from FlexCel and put them into a grid.
    class CellReader
    {
        private bool Only50Rows;
        private SparseCellArray CellData;
        private bool FormatValues;
        private int SheetToRead;
        public DateTime StartSheetSelect;
        public DateTime EndSheetSelect;

        public CellReader(bool aOnly50Rows, SparseCellArray aCellData, bool aFormatValues)
        {
            Only50Rows = aOnly50Rows;
            CellData = aCellData;
            FormatValues = aFormatValues;
        }

        public void OnStartReading(object sender, VirtualCellStartReadingEventArgs e)
        {
            StartSheetSelect = DateTime.Now;
            using (SheetSelectorForm SheetSelector = new SheetSelectorForm(e.SheetNames))
            {

                if (!SheetSelector.Execute())
                {
                    EndSheetSelect = DateTime.Now;
                    e.NextSheet = null; //stop reading
                    return;
                }

                EndSheetSelect = DateTime.Now;
                e.NextSheet = SheetSelector.SelectedSheet;
                SheetToRead = SheetSelector.SelectedSheetIndex + 1;
            }
        }

        public void OnCellRead(object sender, VirtualCellReadEventArgs e)
        {
            if (Only50Rows && e.Cell.Row > 50)
            {
                e.NextSheet = null; //Stop reading all sheets.
                return;
            }

            if (e.Cell.Sheet != SheetToRead)
            {
                e.NextSheet = null; //Stop reading all sheets.
                return;
            }

            if (FormatValues)
            {
                TUIColor Clr = Color.Empty;
                CellData.AddValue(e.Cell.Row, e.Cell.Col,
                   TFlxNumberFormat.FormatValue(e.Cell.Value,
                   ((ExcelFile)sender).GetFormat(e.Cell.XF).Format, ref Clr, ((ExcelFile)sender)));
            }
            else
            {
                CellData.AddValue(e.Cell.Row, e.Cell.Col, Convert.ToString(e.Cell.Value));
            }
        }
    }
}
