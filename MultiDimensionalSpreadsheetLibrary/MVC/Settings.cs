using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.Serialization;
using System.Xml.Serialization;
using Ssepan.Application;
using Ssepan.Collections;
using Ssepan.Utility;

namespace MultiDimensionalSpreadsheetLibrary.MVC
{
	/// <summary>
	/// Summary description for Settings.
	/// </summary>
    //[Serializable()]
    [TypeConverter(typeof(ExpandableObjectConverter))]
    [DataContract(IsReference = true)]
    public class Settings : 
        SettingsBase
    {
        #region declarations
        public const String FILE_TYPE_EXTENSION = "mdss"; 
        public const String FILE_TYPE_NAME = "mdssfile";
        public const String FILE_TYPE_DESCRIPTION = "MultiDimensionalSpreadsheet Settings File";
        #endregion declarations

        #region constructors
        public Settings()
        {
            try
            {
                FileTypeExtension = FILE_TYPE_EXTENSION;
                FileTypeName = FILE_TYPE_NAME;
                FileTypeDescription = FILE_TYPE_DESCRIPTION;
                SerializeAs = SerializationFormat.DataContract;//non-default
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
                throw;
            }
        }
        #endregion

        #region IDisposable support
        ~Settings()
        {
            Dispose(false);
        }

        //inherited; override if additional cleanup needed
        protected override void Dispose(Boolean disposeManagedResources)
        {
            // process only if mananged and unmanaged resources have
            // not been disposed of.
            if (!disposed)
            {
                try
                {
                    //Resources not disposed
                    if (disposeManagedResources)
                    {
                        // dispose managed resources
                    }

                    disposed = true;
                }
                finally
                {
                    // dispose unmanaged resources
                    base.Dispose(disposeManagedResources);
                }
            }
            else
            {
                //Resources already disposed
            }
        }
        #endregion

        #region IEquatable<ISettings> Members
        /// <summary>
        /// Compare property values of two specified Settings objects.
        /// </summary>
        /// <param name="anotherSettings"></param>
        /// <returns></returns>
        public override Boolean Equals(ISettingsComponent other)
        {
            Boolean returnValue = default(Boolean);
            Settings otherSettings = default(Settings);

            try
            {
                otherSettings = other as Settings;

                if (this == otherSettings)
                {
                    returnValue = true;
                }
                else
                {
                    if (!base.Equals(other))
                    {
                        returnValue = false;
                    }
                    else if (this.Version != otherSettings.Version)
                    {
                        returnValue = false;
                    }
                    else if (!this.Sheets.Equals(otherSettings.Sheets))
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
        #endregion IEquatable<ISettings> Members

        #region ListChanged handlers
        void Sheets_ListChanged(object sender, ListChangedEventArgs e)
        {
            try
            {
                this.OnPropertyChanged(String.Format("Settings.Sheets[{0}].{1}", e.NewIndex, (e.PropertyDescriptor == null ? String.Empty : e.PropertyDescriptor.Name)));
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
        }
        #endregion ListChanged handlers

        #region Properties
        [XmlIgnore]
        public override Boolean Dirty
        {
            get
            {
                Boolean returnValue = default(Boolean);

                try
                {
                    if (base.Dirty)
                    {
                        returnValue = true;
                    }
                    else if (_Sheets.Count != __Sheets.Count) //(!_Sheets.Equals(__Sheets))//clooning via DataContract ser-deser resulting in unique items?
                    {
                        returnValue = true;
                    }
                    else if (_Version != __Version)
                    {
                        returnValue = true;
                    }
                    else
                    {
                        returnValue = false;
                    }
                }
                catch (Exception ex)
                {
                    Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                    throw;
                }

                return returnValue;
            }
        }

        #region Persisted Properties

        private EquatableBindingList<Sheet> __Sheets = new EquatableBindingList<Sheet>();
        private EquatableBindingList<Sheet> _Sheets = new EquatableBindingList<Sheet>();
        [DataMember]
        public EquatableBindingList<Sheet> Sheets
        {
            get { return _Sheets; }
            set
            {
                if (_Sheets != null)
                {
                    _Sheets.ListChanged -= new ListChangedEventHandler(Sheets_ListChanged);
                }
                _Sheets = value;
                if (_Sheets != null)
                {
                    _Sheets.ListChanged += new ListChangedEventHandler(Sheets_ListChanged);
                }
                this.OnPropertyChanged("Sheets");
            }
        }

        private String __Version = "0";
        private String _Version = "0";
        [DescriptionAttribute("Application major version"),
        CategoryAttribute("Misc"),
        DefaultValueAttribute(null)]
        [DataMember]
        public String Version
        {
            get { return _Version; }
            set
            {
                _Version = value;
                this.OnPropertyChanged("Version");
            }
        }
        #endregion Persisted Properties
        #endregion Properties

        #region Methods

        /// <summary>
        /// Copies property values from source working fields to detination working fields, then optionally syncs destination.
        /// <param name="destinationSettings"></param>
        /// <param name="sync"></param>
        /// </summary>
        public override void CopyTo(ISettingsComponent destination, Boolean sync)
        {
            Settings destinationSettings = default(Settings);

            try
            {
                destinationSettings = destination as Settings;

                destinationSettings.Sheets = this.Sheets;
                destinationSettings.Version = this.Version;

                base.CopyTo(destination, sync);//also checks and optionally performs sync
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
        }

        /// <summary>
        /// Syncs property values from working fields to reference fields.
        /// </summary>
        public override void Sync()
        {
            try
            {
                __Version = _Version;
                __Sheets = ObjectHelper.Clone<EquatableBindingList<Sheet>>(_Sheets, Ssepan.Utility.SerializationFormat.DataContract);

                //DEBUG:(breaks Dirty flag)
                //force Dirty to False in sheets Equal comparison 
                base.Sync();

                //Note:where we have cloned collections; the collection comparison will never find the orignla items in the cloned collection if it is looking at identity vs content--SJS
                //if (Dirty)
                //{
                //    throw new ApplicationException("Sync failed.");
                //}
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
        }
        #endregion Methods
    }
}
