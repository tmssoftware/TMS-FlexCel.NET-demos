using System;
using System.Globalization;

using FlexCel.Report;

namespace VirtualDatasets
{
    /// <summary>
    /// This class implements the minimum functionality needed to run a FlexCelReport from an array of objects.
    /// No Sorting/Filtering on the config sheet is allowed for this datasource, and you can not use it on master detail relationships.
    /// </summary>
    public class SimpleVirtualArrayDataSource: VirtualDataTable
    {
        #region Private variables
        object[][] FData;
        string[] FColumnCaptions;
        #endregion

        #region Constructors
        public SimpleVirtualArrayDataSource(VirtualDataTable aCreatedBy, object[][] aData, string[] aColumnCaptions, string aTableName) : base(aTableName, aCreatedBy)
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
            return new SimpleVirtualArrayDataSourceState(this);
        }

        #endregion

        #region Data
        public object[][] Data { get { return FData; } }
        #endregion

        #region Filter
        /// <summary>
        /// Even when we don't implement filter, we need to return a new instance when rowFilter is null.
        /// </summary>
        /// <param name="newDataName"></param>
        /// <param name="rowFilter"></param>
        /// <returns></returns>
        public override VirtualDataTable FilterData(string newDataName, string rowFilter)
        {
            if (string.IsNullOrEmpty(rowFilter))
            {
                return new SimpleVirtualArrayDataSource(this, Data, FColumnCaptions, TableName);
            }
            return base.FilterData(newDataName, rowFilter);
        }
        #endregion

    }

    public class SimpleVirtualArrayDataSourceState: VirtualDataTableState
    {
        #region Constructors
        public SimpleVirtualArrayDataSourceState(SimpleVirtualArrayDataSource aTableData) : base(aTableData)
        {
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
                return Data.Length;
            }
        }

        public override object GetValue(int column)
        {
            return Data[Position][column];
        }

        public override object GetValue(int row, int column)
        {
            return Data[row][column];
        }

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

    }

}
