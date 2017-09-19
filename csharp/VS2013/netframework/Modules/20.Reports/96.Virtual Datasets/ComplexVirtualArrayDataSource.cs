using System;
using System.Globalization;
using System.Collections.Generic;

using FlexCel.Report;
using System.Collections;


namespace VirtualDatasets
{
    /// <summary>
    /// This class implements the complete functionality needed to run a FlexCelReport from an array of objects.
    /// Sorting/Filtering on the config sheet is allowed for this datasource, but they are not implemented on an efficient way.
    /// If you do not have an efficient way to do those things (using indexes) and you plan to do them, you will probably get better performance 
    /// using Datasets.
    /// </summary>
    public class ComplexVirtualArrayDataSource: VirtualDataTable
    {
        #region Private variables
        object[][] FData;
        string[] FColumnCaptions;
        #endregion

        #region Constructors
        public ComplexVirtualArrayDataSource(VirtualDataTable aCreatedBy, object[][] aData, string[] aColumnCaptions, string aTableName)
            : base(aTableName, aCreatedBy)
        {
            FData = aData;
            FColumnCaptions = aColumnCaptions;
        }
        #endregion

        #region Columns
        public override int ColumnCount
        {
            get
            {
                return FColumnCaptions.Length;
            }
        }

        public override int GetColumn(string columnName)
        {
            //not very optimized method, but this is just a demo.
            if (columnName == null) return -1;
            columnName = columnName.Trim();
            for (int i = 0; i < FColumnCaptions.Length; i++)
            {
                if (string.Compare(FColumnCaptions[i], columnName, true) == 0) return i;
            }
            return -1;
        }

        public override string GetColumnName(int columnIndex)
        {
            return FColumnCaptions[columnIndex];
        }

        public override string GetColumnCaption(int columnIndex)
        {
            return FColumnCaptions[columnIndex];
        }

        #endregion

        #region Settings
        public override System.Globalization.CultureInfo Locale
        {
            get
            {
                return CultureInfo.CurrentCulture;
            }
        }
        #endregion

        #region Create State
        public override VirtualDataTableState CreateState(string sort, TMasterDetailLink[] masterDetailLinks, TSplitLink splitLink)
        {
            return new ComplexVirtualArrayDataSourceState(this, sort, masterDetailLinks, splitLink);
        }

        #endregion

        #region Data
        public object[][] Data { get { return FData; } }

        public override VirtualDataTable FilterData(string newDataName, string rowFilter)
        {
            if (rowFilter == null || rowFilter.Length == 0)
            {
                return new ComplexVirtualArrayDataSource(this, FData, FColumnCaptions, newDataName);  //no need to copy the data since it is invariant.
            }

            RelationshipType Relationship = RelationshipType.Equals;
            //on this demo we will only support filters of the type: "field = value", "field > value" or "field < value"
            string[] filteredData = rowFilter.Split('=');
            if (filteredData == null || filteredData.Length != 2)
            {
                filteredData = rowFilter.Split('<');
                if (filteredData == null || filteredData.Length != 2)
                {
                    filteredData = rowFilter.Split('>');
                    if (filteredData == null || filteredData.Length != 2)
                    {
                        throw new Exception("Filter \"" + rowFilter + "\" is invalid. The dataset \"" + TableName + "\" only supports filters of the type \"field =/>/< value\".");
                    }
                    else
                    {
                        Relationship = RelationshipType.BiggerThan;
                    }
                }
                else
                {
                    Relationship = RelationshipType.LessThan;
                }

            }

            int ColIndex = GetColumn(filteredData[0].Trim());
            if (ColIndex < 0)
            {
                throw new Exception("Filter \"" + rowFilter + "\" is invalid. Can not find column \"" + filteredData[0].Trim() + "\"");
            }

            //Remember, this is only a demo to show what to do in this event. This code is not good code to use on 
            //a real application!. You should use some index here to find the data or this would crawl on large objects.

            List<object[]> Result = new List<object[]>();
            string SearchValue = filteredData[1].Trim();
            double Searchfloat = 0;
            if (Relationship != RelationshipType.Equals) //when relationship is not equals, we will assume columns are numbers. This is because we do not have any type definition for our columns on this simple example.
            {
                Searchfloat = Convert.ToDouble(filteredData[1].Trim());
            }

            foreach (object[] Row in Data)
            {
                switch (Relationship)
                {
                    case RelationshipType.Equals:
                        if (Convert.ToString(Row[ColIndex]) == SearchValue)
                        {
                            Result.Add(Row);  //remember, data is invariant, so we do not need to clone Row.
                        }
                        break;
                    case RelationshipType.LessThan:
                        if (Convert.ToDouble(Row[ColIndex]) < Searchfloat)
                        {
                            Result.Add(Row);  //remember, data is invariant, so we do not need to clone Row.
                        }
                        break;
                    case RelationshipType.BiggerThan:
                        if (Convert.ToDouble(Row[ColIndex]) > Searchfloat)
                        {
                            Result.Add(Row);  //remember, data is invariant, so we do not need to clone Row.
                        }
                        break;
                }
            }

            return new ComplexVirtualArrayDataSource(this, Result.ToArray(), FColumnCaptions, newDataName);

        }

        public override VirtualDataTable GetDistinct(string newDataName, int[] filterFields)
        {
            if (filterFields == null || filterFields.Length == 0)
            {
                return new ComplexVirtualArrayDataSource(this, FData, FColumnCaptions, newDataName);  //no need to copy the data since it is invariant.
            }

            Dictionary<object[], object[]> Result = new Dictionary<object[], object[]>();
            object[] Keys = new object[filterFields.Length];
            foreach (object[] Row in FData)
            {
                for (int i = 0; i < filterFields.Length; i++)
                {
                    Keys[i] = Row[filterFields[i]];
                }
                Result[Keys] = Keys;
            }

            object[][] R = new object[Data.Length][];
            Result.Keys.CopyTo(R, 0);

            string[] NewColumnCaptions = new string[filterFields.Length];
            for (int i = 0; i < filterFields.Length; i++)
            {
                NewColumnCaptions[i] = FColumnCaptions[filterFields[i]];
            }
            return new ComplexVirtualArrayDataSource(this, R, NewColumnCaptions, newDataName);

        }

        #endregion

    }

    public class ComplexVirtualArrayDataSourceState: VirtualDataTableState
    {
        #region Privates
        private object[][] SortedData;
        private List<object[]> FilteredData;
        #endregion

        #region Constructors
        public ComplexVirtualArrayDataSourceState(ComplexVirtualArrayDataSource aTableData, string sort, TMasterDetailLink[] masterDetailLinks, TSplitLink splitLink)
            : base(aTableData)
        {
            if (sort == null || sort.Trim().Length == 0)
            {
                SortedData = aTableData.Data; //no need to clone, this is invariant.
            }
            else
            {
                SortedData = (object[][])aTableData.Data.Clone();
                int sortcolumn = aTableData.GetColumn(sort);
                if (sortcolumn < 0)
                {
                    throw new Exception("Can not find column \"" + sort + "\" in dataset \"" + TableName);
                }
                Array.Sort(SortedData, new ArrayComparer(sortcolumn));
            }


            //here we should use the data in masterdetaillinks and splitlink to create indexes to make the FilteredrowCount and MoveMasterRecord methods faster.
            //on this demo we are not going to do it.
            if ((masterDetailLinks != null && masterDetailLinks.Length > 0) || splitLink != null)
            {
                FilteredData = new List<object[]>();
            }
        }
        #endregion

        #region Data
        private object[][] Data { get { return ((SimpleVirtualArrayDataSource)TableData).Data; } }

        /// <summary>
        /// Remember that this method should be fast!
        /// </summary>
        public override int RowCount
        {
            get
            {
                if (FilteredData == null) return SortedData.Length;
                return FilteredData.Count;
            }
        }

        public override object GetValue(int column)
        {
            if (FilteredData == null) return SortedData[Position][column];
            return ((object[])FilteredData[Position])[column];
        }

        public override object GetValue(int row, int column)
        {
            if (FilteredData == null) return SortedData[row][column];
            return ((object[])FilteredData[row])[column];
        }
        #endregion

        #region Move
        public override void MoveFirst()
        {
            //No need to do anything in an array, since "Position" is moved for us. 
            //If we had an IEnumerator for example as data backend, we would reset it here.
        }

        public override void MoveNext()
        {
            //No need to do anything in an array, since "Position" is moved for us. 
            //If we had an IEnumerator for example as data backend, we would move it here.
        }

        #endregion


        #region Relationships
        public override void MoveMasterRecord(TMasterDetailLink[] masterDetailLinks, TSplitLink splitLink)
        {
            //Here we need to modify FilteredData to contain only those records on FilteredData that are visible on this state.
            //If we created indexes on this method constructor's, this can be done fast.

            //on this example, as always, we will use a very slow method. The idea of this demo is not to demostrate how to do efficient code
            // (you probably have efficient methods on the bussines objects you are wrapping), but how to override the methods.

            if (FilteredData == null) return;  //this dataset is not on master-detail relationship and does not have any split relationship either.


            FilteredData.Clear();
            int[] ChildColumn = new int[masterDetailLinks.Length];
            for (int i = 0; i < masterDetailLinks.Length; i++)
            {
                ChildColumn[i] = TableData.GetColumn(masterDetailLinks[i].ChildFieldName);
            }

            int SplitPos = 0;
            int StartRow = 0;
            if (splitLink != null)
            {
                StartRow = splitLink.SplitCount * splitLink.ParentDataSource.Position;
            }

            for (int r = 0; r < SortedData.Length; r++)
            {
                object[] Row = SortedData[r];
                bool RowApplies = true;
                for (int i = 0; i < masterDetailLinks.Length; i++)
                {
                    object key = masterDetailLinks[i].ParentDataSource.GetValue(masterDetailLinks[i].ParentField);
                    if (Convert.ToString(Row[ChildColumn[i]]) != Convert.ToString(key))
                    {
                        RowApplies = false;
                        break;
                    }
                }

                if (!RowApplies) continue; //the row does not fit this master detail relationship.

                SplitPos++;
                if (SplitPos <= StartRow) continue; //we are not filling the correct split slot.

                FilteredData.Add(Row);
                if (splitLink != null && FilteredData.Count >= splitLink.SplitCount) return;
            }


        }

        public override int FilteredRowCount(TMasterDetailLink[] masterDetailLinks)
        {
            int Result = 0;

            int[] ChildColumn = new int[masterDetailLinks.Length];
            for (int i = 0; i < masterDetailLinks.Length; i++)
            {
                ChildColumn[i] = TableData.GetColumn(masterDetailLinks[i].ChildFieldName);
            }

            for (int r = 0; r < SortedData.Length; r++)
            {
                object[] Row = SortedData[r];
                bool RowApplies = true;
                for (int i = 0; i < masterDetailLinks.Length; i++)
                {
                    object key = masterDetailLinks[i].ParentDataSource.GetValue(masterDetailLinks[i].ParentField);
                    if (Convert.ToString(Row[ChildColumn[i]]) != Convert.ToString(key))
                    {
                        RowApplies = false;
                        break;
                    }
                }

                if (!RowApplies) continue; //the row does not fit this master detail relationship.

                Result++;
            }

            return Result;

        }


        #endregion

    }

    internal enum RelationshipType
    {
        Equals,
        LessThan,
        BiggerThan
    }

    public class ArrayComparer: IComparer<object[]>
    {
        int Column;
        static CaseInsensitiveComparer Comparer = new CaseInsensitiveComparer();

        public ArrayComparer(int aColumn)
        {
            Column = aColumn;
        }

        int IComparer<object[]>.Compare(Object[] x, Object[] y)
        {
            return (Comparer.Compare(x[Column], y[Column]));
        }

    }


}
