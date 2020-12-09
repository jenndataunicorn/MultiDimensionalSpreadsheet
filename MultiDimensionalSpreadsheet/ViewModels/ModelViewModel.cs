using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using Ssepan.Utility;
using Ssepan.Application;
using Ssepan.Io;
using MultiDimensionalSpreadsheetLibrary;
//using MultiDimensionalSpreadsheetLibrary.Properties;

namespace MultiDimensionalSpreadsheet
{
    /// <summary>
    /// Note: this class can subclass the base without type parameters.
    /// </summary>
    public class ModelViewModel :
        FormsViewModel<Bitmap, Settings, MDSSModel, ModelView>
    {
        #region Declarations

        #region Commands
        //public ICommand FileNewCommand { get; private set; }
        //public ICommand FileOpenCommand { get; private set; }
        //public ICommand FileSaveCommand { get; private set; }
        //public ICommand FileSaveAsCommand { get; private set; }
        //public ICommand FilePrintCommand { get; private set; }
        //public ICommand FileExitCommand { get; private set; }
        //public ICommand EditCopyToClipboardCommand { get; private set; }
        //public ICommand EditPropertiesCommand { get; private set; }
        //public ICommand ViewPreviousMonthCommand { get; private set; }
        //public ICommand ViewPreviousWeekCommand { get; private set; }
        //public ICommand ViewNextWeekCommand { get; private set; }
        //public ICommand ViewNextMonthCommand { get; private set; }
        //public ICommand HelpAboutCommand { get; private set; }
        #endregion Commands
        #endregion Declarations

        #region Constructors
        public ModelViewModel() { }//Note: not called, but need to be present to compile--SJS

        /// <summary>
        /// 
        /// </summary>
        /// <param name="propertyChangedEventHandlerDelegate"></param>
        /// <param name="actionIconImages"></param>
        /// <param name="settingsFileDialogInfo"></param>
        public ModelViewModel
        (
            PropertyChangedEventHandler propertyChangedEventHandlerDelegate,
            Dictionary<String, Bitmap> actionIconImages,
            FileDialogInfo settingsFileDialogInfo
        ) :
            base(propertyChangedEventHandlerDelegate, actionIconImages, settingsFileDialogInfo)
        {
            try
            {
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="propertyChangedEventHandlerDelegate"></param>
        /// <param name="actionIconImages"></param>
        /// <param name="settingsFileDialogInfo"></param>
        /// <param name="view"></param>
        public ModelViewModel
        (
            PropertyChangedEventHandler propertyChangedEventHandlerDelegate,
            Dictionary<String, Bitmap> actionIconImages,
            FileDialogInfo settingsFileDialogInfo,
            ModelView view
        ) :
            base(propertyChangedEventHandlerDelegate, actionIconImages, settingsFileDialogInfo, view)
        {
            try
            {
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
            }
        }
        #endregion Constructors

        #region Properties
        #endregion Properties

        #region Methods
        /// <summary>
        /// model specific, not generioc
        /// </summary>
        public void DoSomething()
        {
            StatusMessage = String.Empty;
            ErrorMessage = String.Empty;

            try
            {
                StartProgressBar
                (
                    "Doing something...",
                    null,
                    null, //_actionIconImages["Xxx"],
                    true,
                    33
                );

                //ModelController<MDSSModel>.Model.Xxx = !ModelController<MDSSModel>.Model.Xxx;

                ModelController<MDSSModel>.Model.Refresh();
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                StopProgressBar(null, String.Format("{0}", ex.Message));
            }
            finally
            {
                StopProgressBar("Did something.");
            }
        }
        #endregion Methods

    }
}
