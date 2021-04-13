using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.Serialization;
using System.Xml.Serialization;
using System.Diagnostics;
using System.Reflection;
using Ssepan.Application;
using Ssepan.Application.MVC;
using Ssepan.Utility;

namespace MultiDimensionalSpreadsheetLibrary
{
    /// <summary>
    /// 
    /// </summary>
    //[Serializable()]
    [DataContract(IsReference = true)]
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class CategoryItem :
        SettingsComponentBase,
        IEquatable<CategoryItem>,
        INotifyPropertyChanged
    {
        #region Declarations
        public const String NewItemName = "(new item)";
        #endregion Declarations

        #region Constructors
        static CategoryItem()
        {
        }

        public CategoryItem()
        {
        }

        /// <summary>
        /// Parent name is the category name.
        /// </summary>
        /// <param name="name"></param>
        public CategoryItem(String name) :
            this()
        {
            try
            {
                Name = name;
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
        }

        /// <summary>
        /// Parent  is the category .
        /// </summary>
        /// <param name="name"></param>
        /// <param name="parent"></param>
        public CategoryItem(String name, Category parent) :
            this(name)
        {
            try
            {
                Parent = parent;
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
        }
        #endregion Constructors

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
        public Boolean Equals(CategoryItem other)
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
                    else if (this.Parent != other.Parent)
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

        #region Properties
        #region NonPersisted Properties
        #endregion NonPersisted Properties

        #region Persisted Properties
        private String _Name = default(String);
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

        private Category _Parent = default(Category);
        [DataMember]
        public Category Parent
        {
            get { return _Parent; }
            set 
            { 
                _Parent = value;
                this.OnPropertyChanged("Parent");
            }
        }

        //private Category _Parent = default(Category);
        //[DataMember]
        //public Category Parent
        //{
        //    get { return _Parent; }
        //    set { _Parent = value; }
        //}
        #endregion Persisted Properties
        #endregion Properties
    }
}
