using System;
using System.Collections.ObjectModel;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Runtime.Serialization;
using System.Xml.Serialization;
using System.Diagnostics;
using System.Reflection;
using Ssepan.Application;
using Ssepan.Application.MVC;
using Ssepan.Collections;
using Ssepan.Utility;

namespace MultiDimensionalSpreadsheetLibrary
{
    /// <summary>
    /// 
    /// </summary>
    //[Serializable()]
    [DataContract(IsReference=true)]
    [KnownType(typeof(Sheet))]
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class Sheet :
        SettingsComponentBase,
        IDisposable,
        IEquatable<Sheet>,
        INotifyPropertyChanged
    {
        #region Declarations
        public const String NewItemName = "(new sheet)";

        private Boolean disposed = false;

        public static Char SheetAddressNameItemDelimiter = ':';
        public static Char SheetAddressNameItemComponentDelimiter = '.';
        
        public enum CellTypes
        { 
            Empty,
            XCategory,
            YCategory,
            Value
        }

        public enum SheetDimensionChange
        { 
            Increase,
            Decrease
        }
        #endregion Declarations

        #region constructors
        public Sheet()
        { 
            try
            {
                if (_Cells != null)
                {
                    _Cells.ListChanged += new ListChangedEventHandler(Cells_ListChanged);
                }
                if (_Categories != null)
                {
                    _Categories.ListChanged += new ListChangedEventHandler(Categories_ListChanged);
                }
                if (_Formulae != null)
                {
                    _Formulae.ListChanged += new ListChangedEventHandler(Formulae_ListChanged);
                    _Formulae.AddingNew += new AddingNewEventHandler(Formulae_AddingNew);
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
        }
        
        public Sheet(String name) :
            this()
        { 
            try
            {
                Name = name;

                //if (Cells != null)
                //{
                //    Cells.ListChanged += new ListChangedEventHandler(Cells_ListChanged);
                //}
                //if (Categories != null)
                //{
                //    Categories.ListChanged += new ListChangedEventHandler(Categories_ListChanged);
                //}
                //if (Formulae != null)
                //{
                //    Formulae.ListChanged += new ListChangedEventHandler(Formulae_ListChanged);
                //    Formulae.AddingNew += new AddingNewEventHandler(Formulae_AddingNew);
                //}
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
        }
        #endregion constructors

        #region IDisposable 
        ~Sheet()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            // dispose of the managed and unmanaged resources
            Dispose(true);

            // tell the GC that the Finalize process no longer needs
            // to be run for this object.
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(Boolean disposeManagedResources)
        {
            // process only if mananged and unmanaged resources have
            // not been disposed of.
            if (!this.disposed)
            {
                //Resources not disposed
                if (disposeManagedResources)
                {
                    // dispose managed resources
                    if (Cells != null)
                    {
                        Cells.ListChanged -= new ListChangedEventHandler(Cells_ListChanged);
                    }
                    if (Categories != null)
                    {
                        Categories.ListChanged -= new ListChangedEventHandler(Categories_ListChanged);
                    }
                    if (Formulae != null)
                    {
                        Formulae.ListChanged -= new ListChangedEventHandler(Formulae_ListChanged);
                        Formulae.AddingNew -= new AddingNewEventHandler(Formulae_AddingNew);
                    }
                }
                // dispose unmanaged resources
                disposed = true;
            }
            else
            {
                //Resources already disposed
            }
        }
        #endregion IDisposable

        #region INotifyPropertyChanged support
        public event PropertyChangedEventHandler PropertyChanged;

        void OnPropertyChanged(String propertyName)
        {
            try
            {
                if (this.PropertyChanged != null)
                {
                    this.PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
#if debug
                    Log.Write(propertyName, EventLogEntryType.Error);

#endif 
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                //throw;
            }
        }
        #endregion INotifyPropertyChanged support

        #region IEquatable<Sheet>
        /// <summary>
        /// Compare property values of this object to another.
        /// </summary>
        /// <param name="anotherSettings"></param>
        /// <returns></returns>
        public Boolean Equals(Sheet other)
        {
            Boolean returnValue = default(Boolean);

            try
            {
                if (this == other)
                {
                    returnValue = true;
                }
                else
                {
                    if (this.Name != other.Name)
                    {
                        returnValue = false;
                    }
                    else if (!this.Cells.Equals(other.Cells))
                    {
                        returnValue = false;
                    }
                    else if (!this.Categories.Equals(other.Categories))
                    {
                        returnValue = false;
                    }
                    else if (!this.Formulae.Equals(other.Formulae))
                    {
                        returnValue = false;
                    }
                    else
                    {
                        returnValue = true;
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }

            return returnValue;
        }
        #endregion IEquatable<Sheet>

        #region ListChanged handlers
        void Cells_ListChanged(object sender, ListChangedEventArgs e)
        {
            try
            {
                this.OnPropertyChanged(String.Format("Sheet.Cells[{0}].{1}", e.NewIndex, (e.PropertyDescriptor == null ? String.Empty : e.PropertyDescriptor.Name)));
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
        }

        void Categories_ListChanged(object sender, ListChangedEventArgs e)
        {
            try
            {
                this.OnPropertyChanged(String.Format("Sheet.Categories[{0}].{1}", e.NewIndex, (e.PropertyDescriptor == null ? String.Empty : e.PropertyDescriptor.Name)));
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
        }

        void Formulae_ListChanged(object sender, ListChangedEventArgs e)
        {
            try
            {
                this.OnPropertyChanged(String.Format("Sheet.Formulae[{0}].{1}", e.NewIndex, (e.PropertyDescriptor == null ? String.Empty : e.PropertyDescriptor.Name)));
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
        }
        #endregion ListChanged handlers

        #region AddingNew handlers
        void Formulae_AddingNew(object sender, AddingNewEventArgs e)
        {
            try
            {
                e.NewObject = new Formula(Formula.NewItemName, this);
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
        }
        #endregion ListChanged handlers

        #region Properties
        #region NonPersisted Properties
        /// <summary>
        /// Grid Column Y (vertical) offset of Y (vertical) category items and values. Offset by number of X (horizontal) categories from top.
        /// </summary>
        [XmlIgnore]
        public Int32 YOffset
        {
            get { return (from sheetCategory in CategoryX() select sheetCategory).Count(); }
        }

        /// <summary>
        /// Grid Column X (horizontal) offset of X (horizontal) category items and values. Offset by number of Y (vertical) categories from left.
        /// </summary>
        [XmlIgnore]
        public Int32 XOffset
        {
            get { return (from sheetCategory in CategoryY() select sheetCategory).Count(); }
        }

        /// <summary>
        /// Sum of X (horizontal) offset and aggregate product of X (horizontal) category item counts.
        /// </summary>
        [XmlIgnore]
        public Int32 ColumnCount
        {
            get 
            {
                if (CategoryX().Count == 0)
                {
                    return XOffset; 
                }
                else
                {
                    return XOffset + (from sheetCategory in CategoryX() select sheetCategory.Items.Count).Aggregate(1, (seed, n) => seed * n); 
                }
            }
        }

        /// <summary>
        /// Sum of Y (vertical) offset and aggregate product of Y (vertical) category item counts.
        /// </summary>
        [XmlIgnore]
        public Int32 RowCount
        {
            get 
            { 
                if (CategoryY().Count == 0)
                {
                    return YOffset; 
                }
                else
                {
                    return YOffset + (from sheetCategory in CategoryY() select sheetCategory.Items.Count).Aggregate(1, (seed, n) => seed * n) ; 
                }
            }
        }

        /// <summary>
        /// Color of Background in empty cells.
        /// </summary>
        [XmlIgnore]
        public Color ColorEmptyBackground
        {
            get { return SystemColors.ControlDark; }
        }

        /// <summary>
        /// Color of foreground in empty cells.
        /// </summary>
        [XmlIgnore]
        public Color ColorEmptyForeground
        {
            get { return SystemColors.ControlText; }
        }

        /// <summary>
        /// Color of Background in category item cells.
        /// </summary>
        [XmlIgnore]
        public Color ColorCategoryItemBackground
        {
            get { return SystemColors.Control; }
        }

        /// <summary>
        /// Color of foreground in category item cells.
        /// </summary>
        [XmlIgnore]
        public Color ColorCategoryItemForeground
        {
            get { return SystemColors.ControlText; }
        }

        /// <summary>
        /// Color of background in value cells.
        /// </summary>
        [XmlIgnore]
        public Color ColorValueBackground
        {
            get { return SystemColors.Window; }
        }

        /// <summary>
        /// Color of foreground in value cells.
        /// </summary>
        [XmlIgnore]
        public Color ColorValueForeground
        {
            get { return SystemColors.ControlText; }
        }
        #endregion NonPersisted Properties

        #region Persisted Properties
        private String _Name = String.Empty;
        [DataMember]
        public String Name
        {
            get { return _Name; }
            set
            {
                _Name = value;
                this.OnPropertyChanged("Name");
            }
        }
        
        private EquatableBindingList<SheetCell> _Cells = new EquatableBindingList<SheetCell>();
        [DataMember]
        public EquatableBindingList<SheetCell> Cells
        {
            get { return _Cells; }
            set
            {
                if (_Cells != null)
                {
                    _Cells.ListChanged -= new ListChangedEventHandler(Cells_ListChanged);
                }
                _Cells = value;
                if (_Cells != null)
                {
                    _Cells.ListChanged += new ListChangedEventHandler(Cells_ListChanged);
                }
                this.OnPropertyChanged("Cells");
            }
        }

        private OrderedEquatableBindingList<Category> _Categories = new OrderedEquatableBindingList<Category>();
        [DataMember]
        public OrderedEquatableBindingList<Category> Categories
        {
            get { return _Categories; }
            set
            {
                if (_Categories != null)
                {
                    _Categories.ListChanged -= new ListChangedEventHandler(Categories_ListChanged);
                }
                _Categories = value;
                if (_Categories != null)
                {
                    _Categories.ListChanged += new ListChangedEventHandler(Categories_ListChanged);
                }
                this.OnPropertyChanged("Categories");
            }
        }

        private OrderedEquatableBindingList<Formula> _Formulae = new OrderedEquatableBindingList<Formula>();
        [DataMember]
        public OrderedEquatableBindingList<Formula> Formulae
        {
            get { return _Formulae; }
            set
            {
                if (_Formulae != null)
                {
                    _Formulae.ListChanged -= new ListChangedEventHandler(Formulae_ListChanged);
                    _Formulae.AddingNew -= new AddingNewEventHandler(Formulae_AddingNew);
                }
                _Formulae = value;
                if (_Formulae != null)
                {
                    _Formulae.ListChanged += new ListChangedEventHandler(Formulae_ListChanged);
                    _Formulae.AddingNew += new AddingNewEventHandler(Formulae_AddingNew);
                }
                this.OnPropertyChanged("Formulae");
            }
        }

        public List<Category> CategoryFilters()
        {
            List<Category> returnValue;
            var result = (from sheetCategory in Categories
                          where sheetCategory.CategoryType == Category.SheetCategoryType.Filter
                          select sheetCategory);
            returnValue = result.ToList<Category>();
            return returnValue;
        }

        public List<Category> CategoryX()
        {
                List<Category> returnValue;
                var result = (from sheetCategory in Categories
                        where sheetCategory.CategoryType == Category.SheetCategoryType.X
                        select sheetCategory) ;
                returnValue = result.ToList<Category>();
                return returnValue ;
        }

        public List<Category> CategoryY()
        {
            List<Category> returnValue;
            var result = (from sheetCategory in Categories
                          where sheetCategory.CategoryType == Category.SheetCategoryType.Y
                          select sheetCategory);
            returnValue = result.ToList<Category>();
            return returnValue;
        }
        #endregion Persisted Properties
        #endregion Properties

        #region Public Methods
        /// <summary>
        /// Takes row, column indices, and returns enum indicating in which range the cell resides
        /// </summary>
        /// <param name="rowIndex"></param>
        /// <param name="columnIndex"></param>
        /// <returns></returns>
        public CellTypes GetCellTypeAtCoordinates
        (
            Int32 rowIndex, 
            Int32 columnIndex
        )
        {
            CellTypes returnValue = CellTypes.Empty;
            try
            {
                if ((rowIndex >= 0) && (rowIndex < this.YOffset))
                {
                    //Empty or X categories
                    if ((columnIndex >= 0) && (columnIndex < this.XOffset))
                    {
                        returnValue = CellTypes.Empty;
                    }
                    else if ((columnIndex >= this.XOffset) && (columnIndex < this.ColumnCount))
                    {
                        returnValue = CellTypes.XCategory;
                    }
                    else
                    {
                        throw new IndexOutOfRangeException(String.Format("Invalid column index ({0}). Valid range is ({1} .. {2}).", columnIndex, 0, this.ColumnCount));
                    }
                }
                else if ((rowIndex >= this.YOffset) && (rowIndex < this.RowCount))
                {
                    //Y categories or Values
                    if ((columnIndex >= 0) && (columnIndex < this.XOffset))
                    {
                        returnValue = CellTypes.YCategory;
                    }
                    else if ((columnIndex >= this.XOffset) && (columnIndex < this.ColumnCount))
                    {
                        returnValue = CellTypes.Value;
                    }
                    else
                    {
                        throw new IndexOutOfRangeException(String.Format("Invalid column index ({0}). Valid range is ({1} .. {2}).", columnIndex, 0, this.ColumnCount));
                    }
                }
                else if ((this.YOffset == 0) && (this.XOffset == 0))
                {
                    //no X or Y categories defined
                    returnValue = CellTypes.Empty;
                }
                else
                {
                    throw new IndexOutOfRangeException(String.Format("Invalid row index ({0}). Valid range is ({1} .. {2}).", rowIndex, 0, this.RowCount));
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
            return returnValue;
        }

        /// <summary>
        /// Takes row / column indices, and returns integer indicating a category-item. 
        /// Also returns an 'out' parameter indicating if the category item is the first instance displayed 
        ///  or a repeated item (for use in drawing an interface).
        /// </summary>
        /// <param name="rowIndex"></param>
        /// <param name="columnIndex"></param>
        /// <param name="isItemFirstIndex"></param>
        /// <returns></returns>
        public Int32 GetCategoryItemIndexAtCoordinates
        (
            Int32 rowIndex, 
            Int32 columnIndex, 
            out Boolean isItemFirstIndex
        )
        {
            Int32 returnValue = -1;
            Int32 itemFirstIndex = -1;
            CellTypes cellType = default(CellTypes);
            Int32 categoryItemsCount =  -1;
            Int32 subCategoriesItemsCount =  -1;
            isItemFirstIndex = default(Boolean);

            try
            {

                //identify cell type from range given by row/column
                switch (this.GetCellTypeAtCoordinates(rowIndex, columnIndex))
                {
                    case Sheet.CellTypes.Empty:
                        {
                            throw new ArgumentOutOfRangeException(String.Format("Cell at coordinates ({0}, {1}) is not a Category Item; it is of type {2}.", rowIndex, columnIndex, cellType.ToString()));

                            //break;
                        }
                    case Sheet.CellTypes.XCategory:
                        {
                            //if x category range, identify category-item @ x, y 
                            categoryItemsCount = (this.CategoryX()).Where((c, index) => index == rowIndex).Select(c => c.Items.Count).Aggregate(1, (seed, n) => seed * n);
                            subCategoriesItemsCount = (this.CategoryX()).Where((c, index) => index > rowIndex).Select(c => c.Items.Count).Aggregate(1, (seed, n) => seed * n);
                            
                            returnValue = ((columnIndex - this.XOffset) / subCategoriesItemsCount) % (categoryItemsCount);
                            
                            itemFirstIndex = ((columnIndex - this.XOffset) % subCategoriesItemsCount);
                            isItemFirstIndex = (itemFirstIndex == 0);

                            break;
                        }
                    case Sheet.CellTypes.YCategory:
                        {
                            categoryItemsCount = (this.CategoryY()).Where((c, index) => index == columnIndex).Select(c => c.Items.Count).Aggregate(1, (seed, n) => seed * n);
                            subCategoriesItemsCount = (this.CategoryY()).Where((c, index) => index > columnIndex).Select(c => c.Items.Count).Aggregate(1, (seed, n) => seed * n);

                            returnValue = ((rowIndex - this.YOffset) / subCategoriesItemsCount) % (categoryItemsCount);

                            itemFirstIndex = ((rowIndex - this.YOffset) % subCategoriesItemsCount);
                            isItemFirstIndex = (itemFirstIndex == 0);

                            break;
                        }
                    case Sheet.CellTypes.Value:
                        {
                            throw new ArgumentOutOfRangeException(String.Format("Cell at coordinates ({0}, {1}) is not a Category Item; it is of type {2}.", rowIndex, columnIndex, cellType.ToString()));

                            //break;
                        }
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
            return returnValue;
        }

        /// <summary>
        /// Takes row / column indices, and returns string representing the name of a category-item. 
        /// Also takes boolean indicating whether to return blank instead of repeating the first instance of the item.
        /// </summary>
        /// <param name="rowIndex"></param>
        /// <param name="columnIndex"></param>
        /// <param name="skipIfNotFirstItemInstance"></param>
        /// <returns></returns>
        public String GetCategoryItemNameAtCoordinates
        (
            Int32 rowIndex, 
            Int32 columnIndex, 
            Boolean skipIfNotFirstItemInstance
        )
        {
            String returnValue = default(String);
            CellTypes cellType = default(CellTypes);

            try
            {
                //identify cell type from range given by row/column
                switch (this.GetCellTypeAtCoordinates(rowIndex, columnIndex))
                {
                    case Sheet.CellTypes.Empty:
                        {
                            throw new ArgumentOutOfRangeException(String.Format("Cell at coordinates ({0}, {1}) is not a Category Item; it is of type {2}.", rowIndex, columnIndex, cellType.ToString()));

                            //break;
                        }
                    case Sheet.CellTypes.XCategory:
                        {
                            //if x category range, identify category-item @ x, y 
                            Int32 itemIndex = -1;
                            Boolean isItemFirstIndex;

                            itemIndex = this.GetCategoryItemIndexAtCoordinates(rowIndex, columnIndex, out isItemFirstIndex);

                            if (isItemFirstIndex || !skipIfNotFirstItemInstance)
                            {
                                returnValue = this.CategoryX()[rowIndex].Items[itemIndex].Name;
                            }
                            else
                            {
                                returnValue = "";
                            }

                            break;
                        }
                    case Sheet.CellTypes.YCategory:
                        {
                            //if y category range, identify category-item @ x, y 
                            Int32 itemIndex = -1;
                            Boolean isItemFirstIndex;

                            itemIndex = this.GetCategoryItemIndexAtCoordinates(rowIndex, columnIndex, out isItemFirstIndex);

                            if (isItemFirstIndex || !skipIfNotFirstItemInstance)
                            {
                                returnValue = this.CategoryY()[columnIndex].Items[itemIndex].Name;
                            }
                            else
                            {
                                returnValue = "";
                            }

                            break;
                        }
                    case Sheet.CellTypes.Value:
                        {
                            throw new ArgumentOutOfRangeException(String.Format("Cell at coordinates ({0}, {1}) is not a Category Item; it is of type {2}.", rowIndex, columnIndex, cellType.ToString()));

                            //break;
                        }
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
            return returnValue;
        }

        /// <summary>
        /// Takes index into Filter category and returns index of Category's selected Item.
        /// </summary>
        /// <param name="filterIndex"></param>
        /// <returns></returns>
        public Int32 GetCategoryItemIndexAtFilterIndex
        (
            Int32 filterIndex
        )
        {
            Int32 returnValue = -1;
            Category category = default(Category);
            List<Category> filters = default(List<Category>);

            try
            {
                filters = this.CategoryFilters();
                if ((filterIndex > -1) && (filterIndex < filters.Count))
                {
                    category = filters[filterIndex];
                    returnValue = category.SelectedItemIndex;
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
            return returnValue;
        }

        /// <summary>
        /// Takes row / column indices of value cell, and returns list of category-items. 
        /// </summary>
        /// <param name="rowIndex"></param>
        /// <param name="columnIndex"></param>
        /// <returns></returns>
        public List<CategoryItem> GetCategoryItemsAtValueCoordinates
        (
            Int32 rowIndex, 
            Int32 columnIndex
        )
        {
            List<CategoryItem> returnValue = new List<CategoryItem>();
            CellTypes cellType = default(CellTypes);
            Int32 categoryItemIndex = -1;
            Boolean isItemFirstIndex = default(Boolean);

            try
            {
                //identify cell type from range given by row/column
                switch (this.GetCellTypeAtCoordinates(rowIndex, columnIndex))
                {
                    case Sheet.CellTypes.Empty:
                        {
                            //if empty range, ignore
                            throw new ArgumentOutOfRangeException(String.Format("Cell at coordinates ({0}, {1}) is not a Value; it is of type {2}.", rowIndex, columnIndex, cellType.ToString()));
                        }
                    case Sheet.CellTypes.XCategory:
                        {
                            //if x category range, ignore
                            throw new ArgumentOutOfRangeException(String.Format("Cell at coordinates ({0}, {1}) is not a Value; it is of type {2}.", rowIndex, columnIndex, cellType.ToString()));
                        }
                    case Sheet.CellTypes.YCategory:
                        {
                            //if y category range, ignore
                            throw new ArgumentOutOfRangeException(String.Format("Cell at coordinates ({0}, {1}) is not a Value; it is of type {2}.", rowIndex, columnIndex, cellType.ToString()));
                        }
                    case Sheet.CellTypes.Value:
                        {
                            //if value range, identify category-item(s) for x, y 
                            
                            //get Filter items
                            for (Int32 filterIndex = 0; filterIndex < this.CategoryFilters().Count; filterIndex++)
                            {
                                categoryItemIndex = GetCategoryItemIndexAtFilterIndex(filterIndex);
                                if (categoryItemIndex != -1)
                                {
                                    returnValue.Add(this.CategoryFilters()[filterIndex].Items[categoryItemIndex]);
                                }
                            }
                            //get X items
                            for (Int32 yIndex = 0; yIndex < this.YOffset; yIndex++)
                            {
                                categoryItemIndex = GetCategoryItemIndexAtCoordinates(yIndex, columnIndex, out isItemFirstIndex);
                                returnValue.Add(this.CategoryX()[yIndex].Items[categoryItemIndex]);
                            }
                            //get Y items
                            for (Int32 xIndex = 0; xIndex < this.XOffset; xIndex++)
                            {
                                categoryItemIndex = GetCategoryItemIndexAtCoordinates(rowIndex, xIndex, out isItemFirstIndex);
                                returnValue.Add(this.CategoryY()[xIndex].Items[categoryItemIndex]);
                            }

                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
            return returnValue;
        }

        /// <summary>
        /// Takes a string representing a named cell (or cell-range) address, and returns a List of CategoryItem (which can be used to search for the cells represented).
        /// </summary>
        /// <param name="addressName"></param>
        /// <param name="exactMatch"></param>
        /// <returns></returns>
        public List<CategoryItem> GetCategoryItemsForAddressName
        (
            String addressName, 
            Boolean exactMatch
        )
        {
            List<CategoryItem> returnValue = default(List<CategoryItem>);
            List<CategoryItem> result = default(List<CategoryItem>);
            String[] addressNameItems = default(String[]);
            String[] addressNameItemComponents = default(String[]);
            String categoryName = default(String);
            String itemName = default(String);

            try
            {
                returnValue = new List<CategoryItem>();

                //validate address
                if ((addressName == null) || (addressName == String.Empty))
                {
                    throw new ArgumentException(String.Format("Address Name was empty: '{0}'", addressName));
                }

                //separate item names
                addressNameItems = addressName.Split(SheetAddressNameItemDelimiter);

                if (addressNameItems == null || addressNameItems.Length == 0)
                {
                    throw new ArgumentException(String.Format("Address Name was empty: '{0}'", addressName));
                }

                foreach (String addressNameItem in addressNameItems)
                {
                    if (addressNameItem == String.Empty)
                    {
                        throw new ArgumentException(String.Format("Address Name component was empty. Address Name: '{0}'", addressName));
                    }

                    //separate item name components
                    addressNameItemComponents = addressNameItem.Split(SheetAddressNameItemComponentDelimiter);

                    if ((addressNameItemComponents == null) || (addressNameItemComponents.Length < 1) || (addressNameItemComponents.Length > 2))
                    {
                        throw new ArgumentException(String.Format("Address Name Item was incorrectly formatted: '{0}'", addressNameItem));
                    }

                    if (addressNameItemComponents.Length == 1)
                    { 
                        categoryName = default(String);
                        itemName = addressNameItemComponents[0];
                    }
                    else if (addressNameItemComponents.Length == 2)
                    { 
                        categoryName = addressNameItemComponents[0];
                        itemName = addressNameItemComponents[1];
                    }

                    //find item(s)
                    //append to List on each iteration; list should have 1 found item each time.
                    result = GetCategoryItemsWhereCategoryAndOrItemsNamesMatch(categoryName, itemName);
                    if (result.Count == 0)
                    {
                        throw new ApplicationException(String.Format("No category items found for category '{0}' and item '{1}'", categoryName, itemName));
                    }
                    if (result.Count == 1)
                    { 
                        returnValue.AddRange(result);
                    }
                    if (result.Count > 1)
                    {
                        if (exactMatch)
                        {
                            throw new ApplicationException(String.Format("{0} category items found for category '{1}' and item '{2}'", result.Count, categoryName, itemName));
                        }
                        else
                        { 
                            returnValue.AddRange(result);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
            return returnValue;
        }

        /// <summary>
        /// Takes row / column indices of value cell, and returns index of a single cell object. 
        /// </summary>
        /// <param name="rowIndex"></param>
        /// <param name="columnIndex"></param>
        /// <returns></returns>
        public Int32 GetValueIndexAtCoordinates
        (
            Int32 rowIndex, 
            Int32 columnIndex
        )
        {
            Int32 returnValue = 0; // -1; 
            CellTypes cellType = default(CellTypes);
            List<CategoryItem> valueCellCategoryItems = default(List<CategoryItem>);
            List<SheetCell> cells = default(List<SheetCell>);
            //return returnValue;
            try
            {
                //identify cell type from range given by row/column
                switch (this.GetCellTypeAtCoordinates(rowIndex, columnIndex))
                {
                    case Sheet.CellTypes.Empty:
                        {
                            //if empty range, ignore
                            throw new ArgumentOutOfRangeException(String.Format("Cell at coordinates ({0}, {1}) is not a Value; it is of type {2}.", rowIndex, columnIndex, cellType.ToString()));
                        }
                    case Sheet.CellTypes.XCategory:
                        {
                            //if x category range, ignore
                            throw new ArgumentOutOfRangeException(String.Format("Cell at coordinates ({0}, {1}) is not a Value; it is of type {2}.", rowIndex, columnIndex, cellType.ToString()));
                        }
                    case Sheet.CellTypes.YCategory:
                        {
                            //if y category range, ignore
                            throw new ArgumentOutOfRangeException(String.Format("Cell at coordinates ({0}, {1}) is not a Value; it is of type {2}.", rowIndex, columnIndex, cellType.ToString()));
                        }
                    case Sheet.CellTypes.Value:
                        {
                            //if value range, identify category-item(s) for x, y 
                            valueCellCategoryItems = GetCategoryItemsAtValueCoordinates(rowIndex, columnIndex);
                            if (valueCellCategoryItems.Count < 1)
                            {
                                returnValue = -1;
                            }
                            else
                            {
                                //perform lookup in cells for matching category and item
                                cells = GetCellsWhereAllCellCategoryItemsMatchAnyInGivenSearch(this.Cells.ToList<SheetCell>(), valueCellCategoryItems);
                                if (cells.Count > 1)
                                {
                                    //too many matches; exact match not found
                                    //This may happen if a dimension  (category) was  recently removed from the sheet,
                                    // but it should have been dealt with at that time.
                                    //TODO:we could try to deal with it here, and try to identify and eliminate cells 
                                    // with categories not present in sheet categories.
                                    returnValue = -1;
                                }
                                else if (cells.Count < 1)
                                {
                                    //no matches found; is missing, need to add new cell
                                    //Add sheet cells with references to sheet-categories
                                    //This may happen if a dimension (category) was  recently added to the sheet,
                                    // but was deferred until cell was needed here.

                                    SheetCell cell = new SheetCell("", valueCellCategoryItems.ToEquatableBindingList<CategoryItem>());
                                    this.Cells.Add(cell);

                                    returnValue = this.Cells.IndexOf(cell);
                                }
                                else //cells.Count == 1
                                {
                                    //identify actual index in Cells collection
                                    returnValue = this.Cells.IndexOf(cells[0]);
                                }
                            }
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
            return returnValue;
        }

        /// <summary>
        /// Set selected Item index of filter Category Item.
        /// </summary>
        /// <param name="categoryItem"></param>
        public void SetFilterCategoryItem
        (
            CategoryItem categoryItem
        )
        {
            Category category = default(Category);

            try
            {
                category = this.Categories.Find(c => c.Equals(categoryItem.Parent));
                //category = this.Categories.Find(c => c.Name == categoryItem.Parent);
                if (category == null)
                {
                    throw new ArgumentException(String.Format("Category not found: '{0}'", categoryItem.Parent));
                }

                if (category.Items.Count > 0)
                {
                    category.SelectedItemIndex = category.Items.IndexOf(categoryItem);
                }
                else
                {
                    throw new ApplicationException(String.Format("Category '{0}' has no items; cannot set a selected Item.", category.Name));
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
        }

        /// <summary>
        /// When assigning a SheetCategoryType to a Category, if the source or destination SheetCategoryType 
        ///  is None, then the assignment will also trigger a change of sheet dimension.
        /// If moving category FROM None, this is an increase in dimension, and we must add one category-item from category to all cells.
        /// If moving category TO None, this is a decrease in dimension, and we must remove one category-item from category from matching cells, and delete all cells containing category's other items.
        /// </summary>
        /// <param name="categoryToBeAssigned">Category that had it's CategoryType assignment changed.</param>
        /// <param name="categoryItemDefault">Default CategoryItem assigned to existing cells if adding a new dimension, 
        /// or used to choose cells to keep (among those with different CategoryItems for the category) if removing a dimension</param>
        /// <param name="sheetDimensionChange">Increase or decrease of sheet dimension, as defined by Sheet.Categories with Category.CategoryType not equal to None.</param>
        public void ChangeSheetDimensionByCategory
        (
            Category categoryToBeAssigned, 
            CategoryItem categoryItemDefault,
            SheetDimensionChange sheetDimensionChange
        )
        {
            try
            {
                switch (sheetDimensionChange)
                {
                    case SheetDimensionChange.Increase:
                        {
                            //Increase dimension of sheet

                            //must add one category-item from category to all cells
                            ProcessMatchingCells
                            (
                                () => this.Cells.ToList<SheetCell>(),
                                c => c.CategoryItems.Add(categoryItemDefault)
                            );

                            break;
                        }
                    case SheetDimensionChange.Decrease:
                        {
                            //Decrease dimension of sheet

                            //remove category-item from matching cells
                            //must remove matching category-item in removed category from containing cells 
                            ProcessMatchingCells
                            (
                                () => GetCellsWhereAllSearchCategoryItemsMatchAnyInGivenCell(this.Cells.ToList<SheetCell>(), new List<CategoryItem> { categoryItemDefault }),
                                c => c.CategoryItems.Remove(categoryItemDefault)
                            );

                            //remove non-matching cells
                            //must delete all cells containing category's other items
                            ProcessMatchingCells
                            (
                                () => GetCellsWhereAllListNotMatchAnyCategoryItem(this.Cells.ToList<SheetCell>(), new List<CategoryItem> { categoryItemDefault }),
                                c => this.Cells.Remove(c)
                            );

                            break;
                        }
                }

            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

            }
        }
        
        /// <summary>
        /// Add or remove a Category to one of the SheetCategoryTypes.
        /// If the source or destination SheetCategoryType is None, then the assignment will also trigger 
        ///  a change of sheet dimension.
        /// </summary>
        /// <param name="categoryNameToBeAssigned">Name of Category to have it's CategoryType assignment changed.</param>
        /// <param name="categoryTypeDestination">CategoryType to be assigned to named Category.</param>
        /// <param name="categoryItemDefault">See <see cref="ChangeSheetDimensionByCategory"/>.
        /// Ignored (and may be null) if there is no change of dimension</param>
        public void AssignCategory
        (
            String categoryNameToBeAssigned, 
            Category.SheetCategoryType categoryTypeDestination, 
            CategoryItem categoryItemDefault
        )
        {
            Category categoryToBeAssigned = default(Category);

            try
            {
                categoryToBeAssigned = this.Categories.Find(c => c.Name == categoryNameToBeAssigned);

                if (categoryToBeAssigned.Items.Count > 0)
                {
                    if (categoryToBeAssigned.CategoryType != categoryTypeDestination)
                    {
                        //move it
                        if (categoryTypeDestination == Category.SheetCategoryType.None)
                        {
                            //move TO None
                            categoryToBeAssigned.CategoryType = categoryTypeDestination;
                            
                             //Decrease dimension of sheet
                            ChangeSheetDimensionByCategory(categoryToBeAssigned, categoryItemDefault, SheetDimensionChange.Decrease);
                        }
                        else if (categoryToBeAssigned.CategoryType == Category.SheetCategoryType.None)
                        {
                            //move FROM None
                            categoryToBeAssigned.CategoryType = categoryTypeDestination;
                            
                            //Increase dimension of sheet
                            ChangeSheetDimensionByCategory(categoryToBeAssigned, categoryItemDefault, SheetDimensionChange.Increase);
                        }
                        else
                        {
                            //simple move among X / Y / Filter
                            categoryToBeAssigned.CategoryType = categoryTypeDestination;

                            //dimension of sheet unchanged
                        }
                    }
                    else
                    { 
                        //source and destination are the same; take no action.
                    }
                }
                else
                {
                    throw new ApplicationException(String.Format("Category '{0}' has no items; it cannot be moved to '{1}'.", categoryToBeAssigned.Name, categoryTypeDestination.ToString()));
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
        }
        #endregion Public Methods

        #region Private Methods
        /// <summary>
        /// Take category name and / or item name, and return list of category items.
        /// </summary>
        /// <param name="categoryName"></param>
        /// <param name="itemName"></param>
        /// <returns></returns>
        public List<CategoryItem> GetCategoryItemsWhereCategoryAndOrItemsNamesMatch
        (
            String categoryName, 
            String itemName
        )
        {
            List<CategoryItem> returnValue = default(List<CategoryItem>);
            String categoryFilter = String.Empty;
            String itemFilter = String.Empty;

            try
            {   
                if ((categoryName == null) || (categoryName == String.Empty))
                {
                    categoryFilter = "*";
                }
                else
                {
                    categoryFilter = categoryName;
                }
                
                if ((itemName == null) || (itemName == String.Empty))
                {
                    itemFilter = "*";
                }
                else
                {
                    itemFilter = itemName;
                }
                
                returnValue =
                    (from c in this.Categories
                     where c.Name.Like(categoryFilter)
                     from i in c.Items
                     where i.Name.Like(itemFilter)
                     select i).ToList<CategoryItem>();
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
            return returnValue;
        }
        #endregion Private Methods

        #region Static Methods
        /// <summary>
        /// Takes a list of Category Items and formats them as string.
        /// </summary>
        /// <param name="criteria"></param>
        /// <returns></returns>
        public static String GetFormattedCellAddress
        (
            List<CategoryItem> criteria
        )
        {
            String returnValue = default(String);
            try
            {
                returnValue = 
                    String.Join
                    (
                        Sheet.SheetAddressNameItemDelimiter.ToString(), 
                        (
                            from categoryItem in criteria
                            select String.Format
                            (
                                "{0}{1}{2}", 
                                categoryItem.Parent, 
                                Sheet.SheetAddressNameItemComponentDelimiter, 
                                categoryItem.Name
                            )
                        ).ToArray<String>()
                    );
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
            return returnValue;
        }

        //re-usable functions to return different cell search results
        #region Cell Searching Components
        /// <summary>
        /// Takes a list of cells and a list of category items, and returns a list of cells. 
        /// It lets the caller decide how to use them, 
        ///  and makes no assumption about how strict a match the caller intended.
        /// Get cells where all cell category items match any entry of a given search's category items.
        /// Typically used to match any??? sheet cells whose category item criteria contain all category item criteria defined at a 
        /// particular set of sheet coordinates.
        /// </summary>
        /// <param name="cells"></param>
        /// <param name="searchCategoryItems"></param>
        /// <returns></returns>
        public static List<SheetCell> GetCellsWhereAllCellCategoryItemsMatchAnyInGivenSearch
        (
            List<SheetCell> cells, 
            List<CategoryItem> searchCategoryItems
        )
        {
            List<SheetCell> returnValue = new List<SheetCell>();
            try
            {
                returnValue = 
                    (from cell in cells
                    where cell.CategoryItems.All(cellCategoryItem => searchCategoryItems.Any(searchCategoryItem => searchCategoryItem == cellCategoryItem))
                    select cell).ToList<SheetCell>();

            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
            return returnValue;
        }

        /// <summary>
        /// Takes a list of cells and a list of category items, and returns a list of cells. 
        /// It lets the caller decide how to use them, 
        ///  and makes no assumption about how strict a match the caller intended.
        /// Get cells where all search criteria category items match any of a given cell's category items.
        /// Typically used by formulas and for finding cells with simlar categories but different category items in one category.
        /// </summary>
        /// <param name="cells"></param>
        /// <param name="searchCategoryItems"></param>
        /// <returns></returns>
        public static List<SheetCell> GetCellsWhereAllSearchCategoryItemsMatchAnyInGivenCell
        (
            List<SheetCell> cells, 
            List<CategoryItem> searchCategoryItems
        )
        {
            List<SheetCell> returnValue = new List<SheetCell>();
            try
            {
                returnValue = 
                    (from cell in cells
                    where searchCategoryItems.All(searchCategoryItem => cell.CategoryItems.Any(cellCategoryItem => cellCategoryItem == searchCategoryItem))
                    select cell).ToList<SheetCell>();

            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
            return returnValue;
        }

        /// <summary>
        /// Takes a list of sheet cells and a list of category items, and returns a list of cell objects. 
        /// It lets the caller decide how to use them, 
        ///  and makes no assumption about how strict a match the caller intended.
        /// Get all cells where *all* entries in the passed category-item list match *any* entry in a given cell's category-item list 
        ///  where the category is a match, but the child items are NOT the matching item.
        /// Used for partial match on a subset of the category-items.
        /// </summary>
        /// <param name="sheetCells"></param>
        /// <param name="sheetCategoryItems"></param>
        /// <returns></returns>
        public static List<SheetCell> GetCellsWhereAllListNotMatchAnyCategoryItem
        (
            List<SheetCell> sheetCells, 
            List<CategoryItem> sheetCategoryItems
        )
        {
            List<SheetCell> returnValue = new List<SheetCell>();
            try
            {
                returnValue = 
                    (from cell in sheetCells
                    where sheetCategoryItems.All(ci1 => cell.CategoryItems.Any(ci2 => ((ci2.Parent == ci1.Parent) && (ci2.Name != ci1.Name))))
                    select cell).ToList<SheetCell>();
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
            return returnValue;
        }
        #endregion Cell Searching Components

        //re-usable functions to apply specified cell processing to different search results
        #region Cell Processing Components
        /// <summary>
        /// Apply a process to all matching cells. Takes a function returning void.
        /// </summary>
        /// <param name="criteriaCategoryItems"></param>
        /// <param name="cellFunctor"></param>
        public static void ProcessMatchingCells
        (
            Func<List<SheetCell>> criteriaFunctor, 
            Action<SheetCell> cellFunctor
        )
        {
            List<SheetCell> matchingCells = default(List<SheetCell>);

            try
            {
                matchingCells = criteriaFunctor();
                foreach (SheetCell cell in matchingCells)
                {
                    cellFunctor(cell);
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
        }

        /// <summary>
        /// Apply a process to all matching cells. Takes a function returning Boolean.
        /// </summary>
        /// <param name="criteriaCategoryItems"></param>
        /// <param name="cellFunctor"></param>
        /// <returns></returns>
        public static Boolean ProcessMatchingCells
        (
            Func<List<SheetCell>> criteriaFunctor, 
            Func<SheetCell, Boolean> cellFunctor
        )
        {
            Boolean returnValue = default(Boolean);
            List<SheetCell> matchingCells = default(List<SheetCell>);

            try
            {
                matchingCells = criteriaFunctor();
                foreach (SheetCell cell in matchingCells)
                {
                    returnValue = cellFunctor(cell);
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
            return returnValue;
        }
        #endregion Cell Processing Components
        #endregion Static Methods

        #region layout examples
        /*
         *   0   1   2   3   
         *  -----------------
         * 0|Q1 |Q2 |Q3 |Q4 |     Quarter
         *  -----------------
         *  (null)
        */
        /*
         *   0     1   2   3   4   
         *  -----------------------
         * 0|#####|Q1 |Q2 |Q3 |Q4 |     Quarter
         *  -----------------------
         * 1|Belts|100|110|120|130|
         *  -----------------------
         * 2|Hats |180|190|200|210|
         *  -----------------------
         *  Department
        */
        /*
         *   0     1   2   3   4   5   6   7   8   
         *  --------------------------------------- 
         * 0|#####|2001           |2002           |     Year
         *  ---------------------------------------  
         * 1|#####|Q1 |Q2 |Q3 |Q4 |Q1 |Q2 |Q3 |Q4 |     Quarter
         *  ---------------------------------------  
         * 2|Belts|100|110|120|130|140|150|160|170| 
         *  ---------------------------------------  
         * 3|Hats |180|190|200|210|220|230|240|250| 
         *  --------------------------------------- 
         *  Department
        */
        /*
         *   0     1     2    3    4    5    6    7    8    9
         *  -----------------------------------------------------  
         * 0|#####|#####|2001               |2002               |     Year
         *  -----------------------------------------------------  
         * 1|#####|#####|Q1  |Q2  |Q3  |Q4  |Q1  |Q2  |Q3  |Q4  |     Quarter
         *  -----------------------------------------------------  
         * 2|Belts|Units|100 |110 |120 |130 |140 |150 |160 |170 | 
         *  -     -----------------------------------------------  
         * 3|     |Price|1.00|1.10|1.20|1.30|1.40|1.50|1.60|1.70| 
         *  -----------------------------------------------------  
         * 4|Hats |Units|180 |190 |200 |210 |220 |230 |240 |250 | 
         *  -     -----------------------------------------------  
         * 5|     |Price|1.80|1.90|2.00|2.10|2.20|2.30|2.40|2.50| 
         *  -----------------------------------------------------  
         *  Department, Details
        */
        #endregion layout examples
    }
}
