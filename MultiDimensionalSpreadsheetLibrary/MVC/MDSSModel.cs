using System;
using System.Collections;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Collections.Generic;
using System.Text;
using Ssepan.Application;
using Ssepan.Collections;
using System.Reflection;
using Ssepan.Utility;

namespace MultiDimensionalSpreadsheetLibrary
{
    /// <summary>
    /// run-time model; relies on settings
    /// </summary>
    //[DefaultPropertyAttribute("SomePropertyName")]
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class MDSSModel :
        ModelBase
    {
        #region Declarations
        #endregion Declarations

        #region Constructors
        public MDSSModel() 
        {
            if (SettingsController<Settings>.Settings == null)
            {
                SettingsController<Settings>.New();
            }
            Debug.Assert(SettingsController<Settings>.Settings != null, "SettingsController<Settings>.Settings != null");
        }
        #endregion Constructors

        #region IEquatable<IModel>
        /// <summary>
        /// Compare property values of two specified Model objects.
        /// </summary>
        /// <param name="anotherSettings"></param>
        /// <returns></returns>
        public override Boolean Equals(IModelComponent other)
        {
            Boolean returnValue = default(Boolean);
            MDSSModel otherModel = default(MDSSModel);

            try
            {
                otherModel = other as MDSSModel;

                if (this == otherModel)
                {
                    returnValue = true;
                }
                else
                {
                    if (!base.Equals(other))
                    {
                        returnValue = false;
                    }
                    else if (!this.Sheets.Equals(otherModel.Sheets))
                    {
                        returnValue = false;
                    }
                    else if (this.Version != otherModel.Version)
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
        #endregion IEquatable<IModel>

        #region ListChanged handlers
        //void Sheets_ListChanged(object sender, ListChangedEventArgs e)
        //{
        //    try
        //    {
        //        this.OnPropertyChanged(String.Format("Sheet.Cells[{0}].{1}", e.NewIndex, (e.PropertyDescriptor == null ? String.Empty : e.PropertyDescriptor.Name)));
        //    }
        //    catch (Exception ex)
        //    {
        //        Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

        //        throw;
        //    }
        //}
        #endregion ListChanged handlers

        #region AddingNew handlers
        //void Sheets_AddingNew(object sender, AddingNewEventArgs e)
        //{
        //    try
        //    {
        //        e.NewObject = new Sheet(Sheet.NewItemName);
        //    }
        //    catch (Exception ex)
        //    {
        //        Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

        //        throw;
        //    }
        //}
        #endregion AddingNew handlers

        #region Properties
        private String[] _Args = default(String[]);
        public String[] Args
        {
            get { return _Args; }
            set
            {
                _Args = value;
                OnPropertyChanged("Args");
            }
        }


        public EquatableBindingList<Sheet> Sheets
        {
            get { return SettingsController<Settings>.Settings.Sheets; }
            set
            {
                SettingsController<Settings>.Settings.Sheets = value;
                this.OnPropertyChanged("Sheets");
            }
        }

        public String Version
        {
            get { return SettingsController<Settings>.Settings.Version; }
            set
            {
                SettingsController<Settings>.Settings.Version = value;
                this.OnPropertyChanged("Version");
            }
        }
        #endregion Properties

        #region Methods
        #endregion Methods

    }
}
