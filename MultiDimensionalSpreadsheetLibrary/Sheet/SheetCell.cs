using System;
using System.Collections.Generic;
using System.ComponentModel;
//using System.Linq;
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
    [DataContract(IsReference = true)]
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class SheetCell :
        SettingsComponentBase,
        IDisposable,
        IEquatable<SheetCell>,
        INotifyPropertyChanged
    {
        #region Declarations
        private Boolean disposed = false;
        #endregion Declarations

        #region Constructors
        static SheetCell()
        {
        }

        public SheetCell()
        {
            try
            {
                if (CategoryItems != null)
                {
                    CategoryItems.ListChanged += new ListChangedEventHandler(CategoryItems_ListChanged);
                }

            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
                throw;
            }
        }

        public SheetCell(String value)
        {
            try
            {
                Value = value;

                if (CategoryItems != null)
                {
                    CategoryItems.ListChanged += new ListChangedEventHandler(CategoryItems_ListChanged);
                }

            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
        }

        public SheetCell(String value, EquatableBindingList<CategoryItem> categoryItems)
        {
            try
            {
                Value = value;
                CategoryItems = categoryItems;

                if (CategoryItems != null)
                {
                    CategoryItems.ListChanged += new ListChangedEventHandler(CategoryItems_ListChanged);
                }

            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
        }
        #endregion Constructors

        #region IDisposable 
        ~SheetCell()
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
                    if (CategoryItems != null)
                    {
                        CategoryItems.ListChanged -= new ListChangedEventHandler(CategoryItems_ListChanged);
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
        public Boolean Equals(SheetCell other)
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
                    if (this.Value != other.Value)
                    {
                        returnValue = false;
                    }
                    else if (!this.CategoryItems.Equals(other.CategoryItems))
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
        void CategoryItems_ListChanged(object sender, ListChangedEventArgs e)
        {
            try
            {
                this.OnPropertyChanged(String.Format("Sheet[].Cells[].CategoryItems[{0}].{1}", e.NewIndex, (e.PropertyDescriptor == null ? String.Empty : e.PropertyDescriptor.Name)));
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
        private Boolean _ValueChangingProgrammatically = default(Boolean); 
        /// <summary>
        /// In specific instances, allow notification to be suppressed when changing value.
        /// </summary>
        [XmlIgnore]
        public Boolean ValueChangingProgrammatically 
        {
	        get { return _ValueChangingProgrammatically; } 
	        set { _ValueChangingProgrammatically = value; } 
        }
        #endregion NonPersisted Properties

        #region Persisted Properties
        private String _Value = default(String);
        [DataMember]
        public String Value
        {
            get { return _Value; }
            set
            {
                _Value = value;
                if (!_ValueChangingProgrammatically)
                {
                    this.OnPropertyChanged("Value");
                }
            }
        }

        private EquatableBindingList<CategoryItem> _CategoryItems = new EquatableBindingList<CategoryItem>();
        [DataMember]
        public EquatableBindingList<CategoryItem> CategoryItems
        {
            get { return _CategoryItems; }
            set
            {
                if (CategoryItems != null)
                {
                    CategoryItems.ListChanged -= new ListChangedEventHandler(CategoryItems_ListChanged);
                    //Note: these category items should already have a parent, since they *should be* coming from  items collection in category
                }
                _CategoryItems = value;
                if (CategoryItems != null)
                {
                    CategoryItems.ListChanged += new ListChangedEventHandler(CategoryItems_ListChanged);
                }
                this.OnPropertyChanged("Categories");
            }
        }

        //private EquatableBindingList<Category> _Categories = new EquatableBindingList<Category>();
        //[DataMember]
        //public EquatableBindingList<Category> Categories
        //{
        //    get { return _Categories; }
        //    set
        //    {
        //        if (Categories != null)
        //        {
        //            Categories.ListChanged -= new ListChangedEventHandler(Categories_ListChanged);
        //        }
        //        _Categories = value;
        //        if (Categories != null)
        //        {
        //            Categories.ListChanged += new ListChangedEventHandler(Categories_ListChanged);
        //        }
        //        this.OnPropertyChanged("Categories");
        //    }
        //}
        #endregion Persisted Properties
        #endregion Properties

        #region Static Methods
        #endregion Static Methods

    }
}
