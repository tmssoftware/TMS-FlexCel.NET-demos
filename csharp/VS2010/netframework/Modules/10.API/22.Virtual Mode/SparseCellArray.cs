using System;
using System.Collections.Generic;
using System.Text;

namespace VirtualMode
{
    ///	<summary>
    ///	  This is a simple class that holds cell values. Items are supposed to
    ///	  be entered in sorted order, and it isn't really production-ready, just
    ///	  to be used in a demo.
    ///	</summary>
    class SparseCellArray
    {
        List<SparseRow> Data;
        int FColCount;

        public SparseCellArray()
        {
            FColCount = 0;
        }
        public void AddValue(int Row, int Col, string Value)
        {
            if (Col > FColCount) FColCount = Col;
            if (Data == null) Data = new List<SparseRow>();
            SparseRow SpRow = new SparseRow(Row);
            int Idx = Data.BinarySearch(SpRow);
            if (Idx < 0)
            {
                SpRow.CreateData();
                Data.Insert(~Idx, SpRow);
            }
            else SpRow = Data[Idx];

            SparseCell SpCell = new SparseCell(Col, Value);
            Idx = SpRow.Data.BinarySearch(SpCell);
            if (Idx < 0)
            {
                SpRow.Data.Insert(~Idx, SpCell);
            }
            else
            {
                SpRow.Data[Idx] = SpCell;
            }

        }

        public string GetValue(int Row, int Col)
        {
            if (Data == null) return null;

            SparseRow SpRow = new SparseRow(Row);
            int Idx = Data.BinarySearch(SpRow);
            if (Idx < 0) return null;
            SpRow = Data[Idx];
            SparseCell SpCell = new SparseCell(Col, null);
            Idx = SpRow.Data.BinarySearch(SpCell);
            if (Idx < 0) return null;
            return SpRow.Data[Idx].Value;
        }

        public int ColCount { get { return FColCount; } }
        public int RowCount
        {
            get
            {
                if (Data == null || Data.Count == 0) return 0;
                return Data[Data.Count - 1].Row;
            }
        }
    }

    struct SparseRow: IComparable<SparseRow>
    {
        public int Row;
        public List<SparseCell> Data;

        public SparseRow(int aRow)
        {
            Row = aRow;
            Data = null;
        }

        public void CreateData()
        {
            Data = new List<SparseCell>();
        }

        #region IComparable<SparseRow> Members

        public int CompareTo(SparseRow other)
        {
            return Row.CompareTo(other.Row);
        }

        #endregion
    }

    class SparseCell: IComparable<SparseCell>

    {
        public int Col;
        public string Value;

        public SparseCell(int aCol, string aValue)
        {
            Col = aCol;
            Value = aValue;
        }

        #region IComparable<SparseCell> Members

        public int CompareTo(SparseCell other)
        {
            return Col.CompareTo(other.Col);
        }

        #endregion
    }
}
