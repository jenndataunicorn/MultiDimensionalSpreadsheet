using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.Serialization;
using System.Xml.Serialization;
using Ssepan.Collections;
using Ssepan.Utility;
using System.Diagnostics;
using System.Reflection;

namespace MultiDimensionalSpreadsheetLibrary
{
    
    //[Serializable()]
    //[DataContract(IsReference=true)]
    /// <summary>
    /// 
    /// </summary>
    [DataContract(IsReference = true)]
    [KnownType(typeof(Category))]
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class Category :
        IDisposable,
        IEquatable<Category>,
        INotifyPropertyChanged
    {
        #region Declarations
        private Boolean disposed = false;

        public enum SheetCategoryType
        {
            None,
            Filter,
            X,
            Y
        }
        #endregion Declarations

        #region constructors
        public Category()
        { 
            try
            {
                if (Items != null)
                {
                    Items.ListChanged += new ListChangedEventHandler(Items_ListChanged);
                    Items.AddingNew += new AddingNewEventHandler(Items_AddingNew);
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
        }

        public Category(String name)
        {
            try
            {
                Name = name;

                if (Items != null)
                {
                    Items.ListChanged += new ListChangedEventHandler(Items_ListChanged);
                    Items.AddingNew += new AddingNewEventHandler(Items_AddingNew);
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
        }

        public Category(String name, SheetCategoryType sheetCategoryType)
        {
            try
            {
                Name = name;
                CategoryType = sheetCategoryType;

                if (Items != null)
                {
                    Items.ListChanged += new ListChangedEventHandler(Items_ListChanged);
                    Items.AddingNew += new AddingNewEventHandler(Items_AddingNew);
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
        }

        public Category(String name, CategoryItem categoryItem)
        {
            try
            {
                Name = name;
                Items.Add(categoryItem);

                if (Items != null)
                {
                    Items.ListChanged += new ListChangedEventHandler(Items_ListChanged);
                    Items.AddingNew += new AddingNewEventHandler(Items_AddingNew);
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
        }

        public Category(String name, CategoryItem categoryItem, SheetCategoryType sheetCategoryType)
        {
            try
            {
                Name = name;
                Items.Add(categoryItem);
                CategoryType = sheetCategoryType;

                if (Items != null)
                {
                    Items.ListChanged += new ListChangedEventHandler(Items_ListChanged);
                    Items.AddingNew += new AddingNewEventHandler(Items_AddingNew);
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
        }

        public Category(String name, OrderedEquatableBindingList<CategoryItem> categoryItems)
        {
            try
            {
                Name = name;
                Items = categoryItems;

                if (Items != null)
                {
                    Items.ListChanged += new ListChangedEventHandler(Items_ListChanged);
                    Items.AddingNew += new AddingNewEventHandler(Items_AddingNew);
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
        }

        public Category(String name, OrderedEquatableBindingList<CategoryItem> categoryItems, SheetCategoryType sheetCategoryType)
        {
            try
            {
                Name = name;
                Items = categoryItems;
                CategoryType = sheetCategoryType;

                if (Items != null)
                {
                    Items.ListChanged += new ListChangedEventHandler(Items_ListChanged);
                    Items.AddingNew += new AddingNewEventHandler(Items_AddingNew);
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
        }
        #endregion constructors

        #region IDisposable 
        ~Category()
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
                    if (Items != null)
                    {
                        Items.AddingNew -= new AddingNewEventHandler(Items_AddingNew);
                        Items.ListChanged -= new ListChangedEventHandler(Items_ListChanged);
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

        #region IEquatable<T>
        /// <summary>
        /// Compare property values of this object to another.
        /// </summary>
        /// <param name="anotherSettings"></param>
        /// <returns></returns>
        public Boolean Equals(Category other)
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
                    else if (!this.Items.Equals(other.Items))
                    {
                        returnValue = false;
                    }
                    else if (this.CategoryType != other.CategoryType)
                    {
                        returnValue = false;
                    }
                    else if (this.SelectedItemIndex != other.SelectedItemIndex)
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
        #endregion IEquatable<T>

        #region ListChanged handlers
        void Items_ListChanged(object sender, ListChangedEventArgs e)
        {
            try
            {
                this.OnPropertyChanged(String.Format("Category.Items[{0}].{1}", e.NewIndex, (e.PropertyDescriptor == null ? String.Empty : e.PropertyDescriptor.Name)));
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
        }
        #endregion ListChanged handlers

        #region AddingNew handlers
        void Items_AddingNew(object sender, AddingNewEventArgs e)
        {
            try
            {
                e.NewObject = new CategoryItem(CategoryItem.NewItemName, this);
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

        private OrderedEquatableBindingList<CategoryItem> _Items = new OrderedEquatableBindingList<CategoryItem>();
        [DataMember]
        public OrderedEquatableBindingList<CategoryItem> Items
        {
            get { return _Items; }
            set
            {
                if (Items != null)
                {
                    Items.ListChanged -= new ListChangedEventHandler(Items_ListChanged);
                    Items.AddingNew -= new AddingNewEventHandler(Items_AddingNew);
                }
                _Items = value;
                if (Items != null)
                {
                    Items.AddingNew += new AddingNewEventHandler(Items_AddingNew);
                    Items.ListChanged += new ListChangedEventHandler(Items_ListChanged);
                }
                this.OnPropertyChanged("Items");
            }
        }

        /// <summary>
        /// Indicates whether the category is assigned to a sheet category, and if so what type.
        /// </summary>
        private SheetCategoryType _CategoryType = SheetCategoryType.None;
        [DataMember]
        public SheetCategoryType CategoryType
        {
            get { return _CategoryType; }
            set 
            {
                _CategoryType = value;
                this.OnPropertyChanged("CategoryType");

                if (value == SheetCategoryType.Filter)
                {
                    SelectedItemIndex = 0;
                }
                else
                { 
                    SelectedItemIndex = -1;
                }
            }
        }

        private Int32 _SelectedItemIndex = -1;
        /// <summary>
        /// Used when CategoryType == SheetCategoryType.Filter 
        /// </summary>
        [DataMember]
        public Int32 SelectedItemIndex
        {
            get { return _SelectedItemIndex; }
            set 
            {
                _SelectedItemIndex = value;
                this.OnPropertyChanged("SelectedItemIndex");
            }
        }
        #endregion Persisted Properties
        #endregion Properties
    }
}
