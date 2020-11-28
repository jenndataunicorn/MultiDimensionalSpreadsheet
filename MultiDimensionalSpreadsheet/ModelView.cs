using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Reflection;
using System.Windows.Forms;
using Ssepan.Application;
using Ssepan.Application.WinForms;
using Ssepan.Application.WinConsole;
using Ssepan.Collections;
using Ssepan.Io;
using Ssepan.Utility;
using MultiDimensionalSpreadsheetLibrary;
using MultiDimensionalSpreadsheetLibrary.UI;
using MultiDimensionalSpreadsheetLibrary.Properties;

namespace MultiDimensionalSpreadsheet
{
    /// <summary>
    /// This is a MVC View
    /// </summary>
    public partial class ModelView :
        Form,
        INotifyPropertyChanged
    {
        #region Declarations
        protected Boolean disposed;

        private Boolean _ValueChangedProgrammatically;

        //cancellation hook
        internal Action cancelDelegate = null;
        protected ModelViewModel ViewModel = default(ModelViewModel);
        #endregion Declarations

        #region Constructors
        public ModelView(/*String[] args*/)
        {
            try
            {
                InitializeComponent();

                ////(re)define default output delegate
                //ConsoleApplication.defaultOutputDelegate = ConsoleApplication.messageBoxWrapperOutputDelegate;

                //subscribe to notifications
                this.PropertyChanged += PropertyChangedEventHandlerDelegate;

                InitViewModel();

                BindSizeAndLocation();
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
            }
        }
        #endregion Constructors

        #region INotifyPropertyChanged
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(String propertyName)
        {
            try
            {
                if (this.PropertyChanged != null)
                {
                    this.PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
                }
            }
            catch (Exception ex)
            {
                ViewModel.ErrorMessage = ex.Message;
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
        }
        #endregion INotifyPropertyChanged

        #region PropertyChangedEventHandlerDelegate
        /// <summary>
        /// Note: property changes update UI manually.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void PropertyChangedEventHandlerDelegate
        (
            Object sender,
            PropertyChangedEventArgs e
        )
        {
            try
            {
                //TODO:"Settings.Sheets[#].Name"?
                if (e.PropertyName == "IsChanged")
                {
                    //ConsoleApplication.defaultOutputDelegate(String.Format("{0}", e.PropertyName));
                    ApplySettings();
                }
                //Status Bar
                else if (e.PropertyName == "ActionIconIsVisible")
                {
                    StatusBarActionIcon.Visible = (ViewModel.ActionIconIsVisible);
                }
                else if (e.PropertyName == "ActionIconImage")
                {
                    StatusBarActionIcon.Image = (ViewModel != null ? ViewModel.ActionIconImage : (Image)null);
                }
                else if (e.PropertyName == "StatusMessage")
                {
                    //replace default action by setting control property
                    StatusBarStatusMessage.Text = ViewModel.StatusMessage;
                    //e = new PropertyChangedEventArgs(e.PropertyName + ".handled");

                    //ConsoleApplication.defaultOutputDelegate(String.Format("{0}", StatusMessage));
                }
                else if (e.PropertyName == "ErrorMessage")
                {
                    //replace default action by setting control property
                    StatusBarErrorMessage.Text = ViewModel.ErrorMessage;
                    //e = new PropertyChangedEventArgs(e.PropertyName + ".handled");

                    //ConsoleApplication.defaultOutputDelegate(String.Format("{0}", ErrorMessage));
                }
                else if (e.PropertyName == "ErrorMessageToolTipText")
                {
                    StatusBarErrorMessage.ToolTipText = ViewModel.ErrorMessageToolTipText;
                }
                else if (e.PropertyName == "ProgressBarValue")
                {
                    StatusBarProgressBar.Value = ViewModel.ProgressBarValue;
                }
                else if (e.PropertyName == "ProgressBarMaximum")
                {
                    StatusBarProgressBar.Maximum = ViewModel.ProgressBarMaximum;
                }
                else if (e.PropertyName == "ProgressBarMinimum")
                {
                    StatusBarProgressBar.Minimum = ViewModel.ProgressBarMinimum;
                }
                else if (e.PropertyName == "ProgressBarStep")
                {
                    StatusBarProgressBar.Step = ViewModel.ProgressBarStep;
                }
                else if (e.PropertyName == "ProgressBarIsMarquee")
                {
                    StatusBarProgressBar.Style = (ViewModel.ProgressBarIsMarquee ? ProgressBarStyle.Marquee : ProgressBarStyle.Blocks);
                }
                else if (e.PropertyName == "ProgressBarIsVisible")
                {
                    StatusBarProgressBar.Visible = (ViewModel.ProgressBarIsVisible);
                }
                else if (e.PropertyName == "DirtyIconIsVisible")
                {
                    StatusBarDirtyMessage.Visible = (ViewModel.DirtyIconIsVisible);
                }
                else if (e.PropertyName == "DirtyIconImage")
                {
                    StatusBarDirtyMessage.Image = ViewModel.DirtyIconImage;
                }
                //use if properties cannot be databound
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
            }
        }
        #endregion PropertyChangedEventHandlerDelegate

        #region Properties
        private String _ViewName = Program.APP_NAME;
        public String ViewName
        {
            get { return _ViewName; }
            set { _ViewName = value; }
        }
        #endregion Properties

        #region Events
        #region Form Events
        /// <summary>
        /// Process Form key presses.
        /// </summary>
        /// <param name="msg"></param>
        /// <param name="keyData"></param>
        /// <returns></returns>
        protected override Boolean ProcessCmdKey(ref Message msg, Keys keyData)
        {
            Boolean returnValue = default(Boolean);
            try
            {
                // Implement the Escape / Cancel keystroke
                if (keyData == Keys.Cancel || keyData == Keys.Escape)
                {
                    //if a long-running cancellable-action has registered 
                    //an escapable-event, trigger it
                    InvokeActionCancel();

                    // This keystroke was handled, 
                    //don't pass to the control with the focus
                    returnValue = true;
                }
                else
                {
                    returnValue = base.ProcessCmdKey(ref msg, keyData);
                }

            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
            }
            return returnValue;
        }

        private void ModelView_Load(object sender, EventArgs e)
        {
            try
            {
                ViewModel.StatusMessage = String.Format("{0} starting...", ViewName);

                ViewModel.StatusMessage = String.Format("{0} started.", ViewName);

                _Run();
            }
            catch (Exception ex)
            {
                ViewModel.ErrorMessage = ex.Message;
                ViewModel.StatusMessage = String.Empty;

                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
            }
        }

        private void ModelView_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                ViewModel.StatusMessage = String.Format("{0} completing...", ViewName);

                DisposeSettings();

                ViewModel.StatusMessage = String.Format("{0} completed.", ViewName);
            }
            catch (Exception ex)
            {
                ViewModel.ErrorMessage = ex.Message.ToString();
                ViewModel.StatusMessage = "";

                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
            }
            finally
            {
                ViewModel = null;
            }
        }
        #endregion Form Events

        #region Menu Events
        private void menuFileNew_Click(Object sender, EventArgs e)
        {
            ViewModel.FileNew();
        }

        private void menuFileOpen_Click(Object sender, EventArgs e)
        {
            ViewModel.FileOpen();
        }

        private void menuFileSave_Click(Object sender, EventArgs e)
        {
            ViewModel.FileSave();
        }

        private void menuFileSaveAs_Click(Object sender, EventArgs e)
        {
            ViewModel.FileSaveAs();
        }

        private void menuFilePrint_Click(object sender, EventArgs e)
        {
            ViewModel.FilePrint();
        }

        private void menuFileExit_Click(Object sender, EventArgs e)
        {
            ViewModel.FileExit();
        }

        private void menuEditProperties_Click(Object sender, EventArgs e)
        {
            ViewModel.EditProperties();
        }

        private void menuEditCopyToClipboard_Click(Object sender, EventArgs e)
        {
            ViewModel.EditCopy();
        }

        private void menuHelpAbout_Click(Object sender, EventArgs e)
        {
            ViewModel.HelpAbout<AssemblyInfo>();
        }
        #endregion Menu Events
        
        //TODO:move into viewmodel as appropriate--SJS
        #region Control Events
        #region Sheets
        private void dgvSheets_CurrentCellChanged(object sender, EventArgs e)
        {
            try
            {
                MDSSController<MDSSModel>.Model.Refresh();
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
        }

        private void dgvSheets_SelectionChanged(object sender, EventArgs e)
        {
            try
            {
                MDSSController<MDSSModel>.Model.Refresh();
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
        }
        #endregion Sheets

        #region Categories
        private void contextMenuStripCategory_Opening(object sender, CancelEventArgs e)
        {
            Int32 count = default(Int32);
            ContextMenuStrip contextMenu = sender as ContextMenuStrip;
            Int32 selectedSheetRowIndex = -1;
            DataGridView source = default(DataGridView);
            
            try
            {
                source = contextMenu.SourceControl as DataGridView;
                count = source.SelectedRows.Count;

                if (count == 1)
                {
                    //one row/cell selected
                    if (source.SelectedRows[0].Cells[0].Value.ToString() != String.Empty)
                    {
                        //cell has a value -- i.e. is not New.
                        selectedSheetRowIndex = DataGridViewHelper.GetIndex(dgvSheets);
                        if (selectedSheetRowIndex != -1)
                        {
                            Category category = SettingsController<Settings>.Settings.Sheets[selectedSheetRowIndex].Categories.Find(c => c.Name == source.SelectedRows[0].Cells[0].Value.ToString());
                            if (category.Items.Count > 0)
                            {
                                //category has items.
                                tsmiCategoryMoveToFilterCategory.Enabled = dgvSheetCategoriesFilter.Enabled;
                                tsmiCategoryMoveToXCategory.Enabled = dgvSheetCategoriesX.Enabled;
                                tsmiCategoryMoveToYCategory.Enabled = dgvSheetCategoriesY.Enabled;
                            }
                            else
                            { 
                                tsmiCategoryMoveToFilterCategory.Enabled = false;
                                tsmiCategoryMoveToXCategory.Enabled = false;
                                tsmiCategoryMoveToYCategory.Enabled = false;
                            }
                        }
                        else
                        { 
                            tsmiCategoryMoveToFilterCategory.Enabled = false;
                            tsmiCategoryMoveToXCategory.Enabled = false;
                            tsmiCategoryMoveToYCategory.Enabled = false;
                        }
                    }
                    else
                    { 
                        tsmiCategoryMoveToFilterCategory.Enabled = false;
                        tsmiCategoryMoveToXCategory.Enabled = false;
                        tsmiCategoryMoveToYCategory.Enabled = false;
                    }
                }
                else
                {
                    tsmiCategoryMoveToFilterCategory.Enabled = false;
                    tsmiCategoryMoveToXCategory.Enabled = false;
                    tsmiCategoryMoveToYCategory.Enabled = false;
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
        }

        private void tsmiCategoryMoveToFilterCategory_Click(object sender, EventArgs e)
        {
            String categoryName = default(String);
            Int32 selectedSheetRowIndex = -1;

            try
            {
                selectedSheetRowIndex = DataGridViewHelper.GetIndex(dgvSheets);
                if (selectedSheetRowIndex != -1)
                {
                    if (this.dgvCategories.SelectedRows.Count == 1)
                    {
                        categoryName = this.dgvCategories.SelectedRows[0].Cells[0].Value.ToString();

                        if (!MDSSController<MDSSModel>.AddSheetCategory(SettingsController<Settings>.Settings.Sheets[selectedSheetRowIndex], categoryName, Category.SheetCategoryType.Filter))
                        { 
                            throw new ApplicationException(String.Format("Category '{0}' not added to '{1}'.", categoryName, Category.SheetCategoryType.Filter.ToString()));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                ViewModel.StopProgressBar(null, String.Format("{0}", ex.Message));
            }
        }

        private void tsmiCategoryMoveToXCategory_Click(object sender, EventArgs e)
        {
            String categoryName = default(String);
            Int32 selectedSheetRowIndex = -1;

            try
            {
                selectedSheetRowIndex = DataGridViewHelper.GetIndex(dgvSheets);
                if (selectedSheetRowIndex != -1)
                {
                    if (this.dgvCategories.SelectedRows.Count == 1)
                    {
                        categoryName = this.dgvCategories.SelectedRows[0].Cells[0].Value.ToString();

                        if (!MDSSController<MDSSModel>.AddSheetCategory(SettingsController<Settings>.Settings.Sheets[selectedSheetRowIndex], categoryName, Category.SheetCategoryType.X))
                        {
                            throw new ApplicationException(String.Format("Category '{0}' not added to '{1}'.", categoryName, Category.SheetCategoryType.X.ToString()));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                ViewModel.StopProgressBar(null, String.Format("{0}", ex.Message));
            }
        }

        private void tsmiCategoryMoveToYCategory_Click(object sender, EventArgs e)
        {
            String categoryName = default(String);
            Int32 selectedSheetRowIndex = -1;

            try
            {
                selectedSheetRowIndex = DataGridViewHelper.GetIndex(dgvSheets);
                if (selectedSheetRowIndex != -1)
                {
                    if (this.dgvCategories.SelectedRows.Count == 1)
                    {
                        categoryName = this.dgvCategories.SelectedRows[0].Cells[0].Value.ToString();

                        if (!MDSSController<MDSSModel>.AddSheetCategory(SettingsController<Settings>.Settings.Sheets[selectedSheetRowIndex], categoryName, Category.SheetCategoryType.Y))
                        {
                            throw new ApplicationException(String.Format("Category '{0}' not added to '{1}'.", categoryName, Category.SheetCategoryType.Y.ToString()));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                ViewModel.StopProgressBar(null, String.Format("{0}", ex.Message));
            }
        }

        private void dgvCategories_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                DataGridView categoryList = (sender as DataGridView);

                if (e.Button == MouseButtons.Right)
                {
                    if (e.RowIndex != -1)
                    {
                        categoryList[0, e.RowIndex].Selected = true;
                        categoryList.Rows[e.RowIndex].Selected = true;
                        categoryList.ContextMenuStrip.Show(categoryList, new Point(e.X, e.Y));
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                ViewModel.StopProgressBar(null, String.Format("{0}", ex.Message));
            }
        }

        private void dgvCategories_UserDeletingRow(object sender, DataGridViewRowCancelEventArgs e)
        {
            String categoryName = default(String);
            Category category = default(Category);
            Int32 selectedSheetRowIndex = -1;

            try
            {
                //identify active sheet
                selectedSheetRowIndex = DataGridViewHelper.GetIndex(dgvSheets);
                if (selectedSheetRowIndex != -1)
                {
                    //find category by name
                    categoryName = e.Row.Cells[0].Value.ToString();
                    category = SettingsController<Settings>.Settings.Sheets[selectedSheetRowIndex].Categories.Find(c => c.Name == categoryName);
                    if (category == null)
                    {
                        throw new ApplicationException(String.Format("Unable to find category: '{0}'", categoryName));
                    }

                    //if current Category is assigned, stop the Delete
                    if (category.CategoryType != Category.SheetCategoryType.None)
                    {
                        e.Cancel = true;
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                ViewModel.StopProgressBar(null, String.Format("{0}", ex.Message));
            }
        }
        #endregion Categories

        #region Category Items
        private void dgvCategoryItems_UserDeletingRow(object sender, DataGridViewRowCancelEventArgs e)
        {
            String categoryItemName = default(String);
            Category category = default(Category);
            Int32 selectedSheetRowIndex = -1;

            try
            {
                //identify active sheet
                selectedSheetRowIndex = DataGridViewHelper.GetIndex(dgvSheets);
                if (selectedSheetRowIndex != -1)
                {
                    //find category by name
                    categoryItemName = e.Row.Cells[0].Value.ToString();
                    category = SettingsController<Settings>.Settings.Sheets[selectedSheetRowIndex].Categories.Find(c => c.Items.Any(ci => ci.Name == categoryItemName));
                    if (category == null)
                    {
                        throw new ApplicationException(String.Format("Unable to find category: '{0}'", categoryItemName));
                    }
                    
                    //TODO:check for more than one result

                    //if current Category is assigned, stop the Delete
                    if (category.CategoryType != Category.SheetCategoryType.None)
                    {
                        e.Cancel = true;
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                ViewModel.StopProgressBar(null, String.Format("{0}", ex.Message));
            }
        }

        private void contextMenuStripCategoryItem_Opening(object sender, CancelEventArgs e)
        {
            Int32 count = default(Int32);
            ContextMenuStrip contextMenu = sender as ContextMenuStrip;
            Int32 selectedSheetRowIndex = -1;
            DataGridView source = default(DataGridView);

            try
            {
                source = contextMenu.SourceControl as DataGridView;
                count = source.SelectedRows.Count;

                if (count == 1)
                {
                    //one row/cell selected
                    if ((source.SelectedRows[0].Cells[0].Value == null) || (source.SelectedRows[0].Cells[0].Value.ToString() == String.Empty) || (source.SelectedRows[0].Cells[0].Value.ToString() == CategoryItem.NewItemName))
                    {
                        tsmiCategoryItemPromoteCategoryItem.Enabled = false;
                        tsmiCategoryItemDemoteCategoryItem.Enabled = false;
                    }
                    else
                    {
                        //cell has a value -- i.e. is not New.
                        selectedSheetRowIndex = DataGridViewHelper.GetIndex(dgvSheets);
                        if (selectedSheetRowIndex != -1)
                        {
                            //a sheet is selected
                            tsmiCategoryItemPromoteCategoryItem.Enabled = dgvCategoryItems.Enabled;
                            tsmiCategoryItemDemoteCategoryItem.Enabled = dgvCategoryItems.Enabled;
                        }
                        else
                        {
                            tsmiCategoryItemPromoteCategoryItem.Enabled = false;
                            tsmiCategoryItemDemoteCategoryItem.Enabled = false;
                        }
                    }
                }
                else
                {
                    tsmiCategoryItemPromoteCategoryItem.Enabled = false;
                    tsmiCategoryItemDemoteCategoryItem.Enabled = false;
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
        }

        private void tsmiCategoryItemPromoteCategoryItem_Click(object sender, EventArgs e)
        {
            Category category = default(Category);
            CategoryItem categoryItem = default(CategoryItem);
            String categoryItemName = default(String);
            Int32 selectedSheetRowIndex = -1;

            try
            {
                selectedSheetRowIndex = DataGridViewHelper.GetIndex(dgvSheets);
                if (selectedSheetRowIndex != -1)
                {
                    if (this.dgvCategoryItems.SelectedRows.Count == 1)
                    {
                        categoryItem = (CategoryItem)this.dgvCategoryItems.SelectedRows[0].DataBoundItem;
                        categoryItemName = categoryItem.Name;
                        category = categoryItem.Parent;

                        if (!MDSSController<MDSSModel>.ShiftSheetCategoryItem(SettingsController<Settings>.Settings.Sheets[selectedSheetRowIndex], category.Name, categoryItemName, ListOfTExtension.ShiftTypes.Promote))
                        {
                            throw new ApplicationException(String.Format("Category Item '{0}' not promoted.", categoryItemName));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                ViewModel.StopProgressBar(null, String.Format("{0}", ex.Message));
            }
        }

        private void tsmiCategoryItemDemoteCategoryItem_Click(object sender, EventArgs e)
        {
            Category category = default(Category);
            CategoryItem categoryItem = default(CategoryItem);
            String categoryItemName = default(String);
            Int32 selectedSheetRowIndex = -1;

            try
            {
                selectedSheetRowIndex = DataGridViewHelper.GetIndex(dgvSheets);
                if (selectedSheetRowIndex != -1)
                {
                    if (this.dgvCategoryItems.SelectedRows.Count == 1)
                    {
                        categoryItem = (CategoryItem)this.dgvCategoryItems.SelectedRows[0].DataBoundItem;
                        categoryItemName = categoryItem.Name;
                        category = categoryItem.Parent;

                        if (!MDSSController<MDSSModel>.ShiftSheetCategoryItem(SettingsController<Settings>.Settings.Sheets[selectedSheetRowIndex], category.Name, categoryItemName, ListOfTExtension.ShiftTypes.Demote))
                        {
                            throw new ApplicationException(String.Format("Category Item '{0}' not demoted.", categoryItemName));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                ViewModel.StopProgressBar(null, String.Format("{0}", ex.Message));
            }
        }

        private void dgvCategoryItems_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {

            try
            {
                DataGridView formulaList = (sender as DataGridView);

                if (e.Button == MouseButtons.Right)
                {
                    if (e.RowIndex != -1)
                    {
                        formulaList[0, e.RowIndex].Selected = true;
                        formulaList.Rows[e.RowIndex].Selected = true;
                        formulaList.ContextMenuStrip.Show(formulaList, new Point(e.X, e.Y));
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                ViewModel.StopProgressBar(null, String.Format("{0}", ex.Message));
            }
        }
        #endregion Category Items

        #region Formulae
        private void contextMenuStripFormula_Opening(object sender, CancelEventArgs e)
        {
            Int32 count = default(Int32);
            ContextMenuStrip contextMenu = sender as ContextMenuStrip;
            Int32 selectedSheetRowIndex = -1;
            DataGridView source = default(DataGridView);

            try
            {
                source = contextMenu.SourceControl as DataGridView;
                count = source.SelectedRows.Count;

                if (count == 1)
                {
                    //one row/cell selected
                    if ((source.SelectedRows[0].Cells[2].Value == null) || (source.SelectedRows[0].Cells[2].Value.ToString() == String.Empty) || (source.SelectedRows[0].Cells[2].Value.ToString() == Formula.NewItemName))
                    {
                        tsmiFormulaPromoteFormula.Enabled = false;
                        tsmiFormulaDemoteFormula.Enabled = false;
                    }
                    else
                    { 
                        //cell has a value -- i.e. is not New.
                        selectedSheetRowIndex = DataGridViewHelper.GetIndex(dgvSheets);
                        if (selectedSheetRowIndex != -1)
                        {
                            //a sheet is selected
                            tsmiFormulaPromoteFormula.Enabled = dgvFormulae.Enabled;
                            tsmiFormulaDemoteFormula.Enabled = dgvFormulae.Enabled;
                        }
                        else
                        { 
                            tsmiFormulaPromoteFormula.Enabled = false;
                            tsmiFormulaDemoteFormula.Enabled = false;
                        }
                    }
                }
                else
                {
                    tsmiFormulaPromoteFormula.Enabled = false;
                    tsmiFormulaDemoteFormula.Enabled = false;
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
        }

        private void tsmiFormulaPromoteFormula_Click(object sender, EventArgs e)
        {
            String formulaName = default(String);
            Int32 selectedSheetRowIndex = -1;

            try
            {
                selectedSheetRowIndex = DataGridViewHelper.GetIndex(dgvSheets);
                if (selectedSheetRowIndex != -1)
                {
                    if (this.dgvFormulae.SelectedRows.Count == 1)
                    {
                        formulaName = this.dgvFormulae.SelectedRows[0].Cells[2].Value.ToString();

                        if (!MDSSController<MDSSModel>.ShiftSheetFormula(SettingsController<Settings>.Settings.Sheets[selectedSheetRowIndex], formulaName, ListOfTExtension.ShiftTypes.Promote))
                        {
                            throw new ApplicationException(String.Format("Formula '{0}' not promoted.", formulaName));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                ViewModel.StopProgressBar(null, String.Format("{0}", ex.Message));
            }
        }

        private void tsmiFormulaDemoteFormula_Click(object sender, EventArgs e)
        {
            String formulaName = default(String);
            Int32 selectedSheetRowIndex = -1;

            try
            {
                selectedSheetRowIndex = DataGridViewHelper.GetIndex(dgvSheets);
                if (selectedSheetRowIndex != -1)
                {
                    if (this.dgvFormulae.SelectedRows.Count == 1)
                    {
                        formulaName = this.dgvFormulae.SelectedRows[0].Cells[2].Value.ToString();

                        if (!MDSSController<MDSSModel>.ShiftSheetFormula(SettingsController<Settings>.Settings.Sheets[selectedSheetRowIndex], formulaName, ListOfTExtension.ShiftTypes.Demote))
                        {
                            throw new ApplicationException(String.Format("Formula '{0}' not demoted.", formulaName));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                ViewModel.StopProgressBar(null, String.Format("{0}", ex.Message));
            }
        }

        private void dgvFormulae_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {

            try
            {
                DataGridView formulaList = (sender as DataGridView);

                if (e.Button == MouseButtons.Right)
                {
                    if (e.RowIndex != -1)
                    {
                        formulaList[0, e.RowIndex].Selected = true;
                        formulaList.Rows[e.RowIndex].Selected = true;
                        formulaList.ContextMenuStrip.Show(formulaList, new Point(e.X, e.Y));
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                ViewModel.StopProgressBar(null, String.Format("{0}", ex.Message));
            }
        }
        #endregion Formulae

        #region Sheet
        private void dgvSheet_CellValueNeeded(object sender, DataGridViewCellValueEventArgs e)
        {
            Int32 selectedSheetRowIndex = -1;

            try
            {
                selectedSheetRowIndex = DataGridViewHelper.GetIndex(dgvSheets);
                if (selectedSheetRowIndex != -1)
                {
                    DataGridViewHelper.SheetCellValueNeeded(sender, e, SettingsController<Settings>.Settings.Sheets[selectedSheetRowIndex]);
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                //throw;
            }
        }

        private void dgvSheet_CellValuePushed(object sender, DataGridViewCellValueEventArgs e)
        {
            Int32 selectedSheetRowIndex = -1;

            try
            {
                selectedSheetRowIndex = DataGridViewHelper.GetIndex(dgvSheets);
                if (selectedSheetRowIndex != -1)
                {
                    DataGridViewHelper.SheetCellValuePushed(sender, e, SettingsController<Settings>.Settings.Sheets[selectedSheetRowIndex]);
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
        }
        #endregion Sheet

        #region Sheet Categories
        #region Filters
        private void contextMenuStripCategoryFilter_Opening(object sender, CancelEventArgs e)
        {
            Int32 count = -1;
            ContextMenuStrip contextMenu = sender as ContextMenuStrip;
            DataGridView source = default(DataGridView);

            try
            {
                source = contextMenu.SourceControl as DataGridView;
                count = source.SelectedCells.Count;

                if (count == 1)
                {
                    tsmiFilterCategorySelectItem.Enabled = dgvSheetCategoriesFilter.Enabled;
                    tsmiFilterCategoryMoveToXCategory.Enabled = dgvSheetCategoriesX.Enabled;
                    tsmiFilterCategoryMoveToYCategory.Enabled = dgvSheetCategoriesY.Enabled;
                    tsmiFilterCategoryRemoveCategory.Enabled = dgvSheetCategoriesFilter.Enabled;
                    tsmiFilterCategoryPromoteCategory.Enabled = dgvSheetCategoriesFilter.Enabled;
                    tsmiFilterCategoryDemoteCategory.Enabled = dgvSheetCategoriesFilter.Enabled;
                }
                else
                {
                    tsmiFilterCategorySelectItem.Enabled = false;
                    tsmiFilterCategoryMoveToXCategory.Enabled = false;
                    tsmiFilterCategoryMoveToYCategory.Enabled = false;
                    tsmiFilterCategoryRemoveCategory.Enabled = false;
                    tsmiFilterCategoryPromoteCategory.Enabled = false;
                    tsmiFilterCategoryDemoteCategory.Enabled = false;
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
        }

        private void dgvSheetCategoriesFilter_CellValueNeeded(object sender, DataGridViewCellValueEventArgs e)
        {
            Int32 selectedSheetRowIndex = -1;

            try
            {
                selectedSheetRowIndex = DataGridViewHelper.GetIndex(dgvSheets);
                if (selectedSheetRowIndex != -1)
                {
                    DataGridViewHelper.SheetCategoryCellValueNeeded(sender, e, SettingsController<Settings>.Settings.Sheets[selectedSheetRowIndex], Category.SheetCategoryType.Filter);
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
        }

        private void tsmiFilterCategorySelectItem_Click(object sender, EventArgs e)
        {
            String categoryName = default(String);
            Int32 selectedSheetRowIndex = -1;

            try
            {
                selectedSheetRowIndex = DataGridViewHelper.GetIndex(dgvSheets);
                if (selectedSheetRowIndex != -1)
                {
                    if (this.dgvSheetCategoriesFilter.SelectedCells.Count > 0)
                    {
                        //only works if header is displayed
                        categoryName = this.dgvSheetCategoriesFilter.SelectedCells[0].OwningColumn.HeaderText;
                        ////use if header is not displayed
                        //category = this.dgvSheetCategoriesFilter.SelectedCells[0].Value.ToString();

                        if (!MDSSController<MDSSModel>.SelectFilterCategoryItem(SettingsController<Settings>.Settings.Sheets[selectedSheetRowIndex], categoryName))
                        {
                            throw new ApplicationException(String.Format("Category '{0}' item not selected.", categoryName));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                ViewModel.StopProgressBar(null, String.Format("{0}", ex.Message));
            }
        }

        private void tsmiFilterCategoryMoveToXCategory_Click(object sender, EventArgs e)
        {
            String categoryName = default(String);
            Int32 selectedSheetRowIndex = -1;

            try
            {
                selectedSheetRowIndex = DataGridViewHelper.GetIndex(dgvSheets);
                if (selectedSheetRowIndex != -1)
                {
                    if (this.dgvSheetCategoriesFilter.SelectedCells.Count > 0)
                    {
                        categoryName = this.dgvSheetCategoriesFilter.SelectedCells[0].OwningColumn.HeaderText;

                        if (!MDSSController<MDSSModel>.MoveSheetCategory(SettingsController<Settings>.Settings.Sheets[selectedSheetRowIndex], categoryName, Category.SheetCategoryType.X))
                        { 
                            throw new ApplicationException(String.Format("Category '{0}' not moved to '{1}'.", categoryName, Category.SheetCategoryType.X.ToString()));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                ViewModel.StopProgressBar(null, String.Format("{0}", ex.Message));
            }
        }

        private void tsmiFilterCategoryMoveToYCategory_Click(object sender, EventArgs e)
        {
            String categoryName = default(String);
            Int32 selectedSheetRowIndex = -1;

            try
            {
                selectedSheetRowIndex = DataGridViewHelper.GetIndex(dgvSheets);
                if (selectedSheetRowIndex != -1)
                {
                    if (this.dgvSheetCategoriesFilter.SelectedCells.Count == 1)
                    {
                        categoryName = this.dgvSheetCategoriesFilter.SelectedCells[0].OwningColumn.HeaderText;

                        if (!MDSSController<MDSSModel>.MoveSheetCategory(SettingsController<Settings>.Settings.Sheets[selectedSheetRowIndex], categoryName, Category.SheetCategoryType.Y))
                        { 
                            throw new ApplicationException(String.Format("Category '{0}' not moved to '{1}'.", categoryName, Category.SheetCategoryType.Y.ToString()));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                ViewModel.StopProgressBar(null, String.Format("{0}", ex.Message));
            }
        }

        private void tsmiFilterCategoryRemoveCategory_Click(object sender, EventArgs e)
        {
            String categoryName = default(String);
            Int32 selectedSheetRowIndex = -1;

            try
            {
                selectedSheetRowIndex = DataGridViewHelper.GetIndex(dgvSheets);
                if (selectedSheetRowIndex != -1)
                {
                    if (this.dgvSheetCategoriesFilter.SelectedCells.Count == 1)
                    {
                        categoryName = this.dgvSheetCategoriesFilter.SelectedCells[0].OwningColumn.HeaderText;

                        if (!MDSSController<MDSSModel>.RemoveSheetCategory(SettingsController<Settings>.Settings.Sheets[selectedSheetRowIndex], categoryName))
                        { 
                            throw new ApplicationException(String.Format("Category '{0}' not removed from '{1}'.", categoryName, Category.SheetCategoryType.Filter.ToString()));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                ViewModel.StopProgressBar(null, String.Format("{0}", ex.Message));
            }
        }

        private void tsmiFilterCategoryPromoteCategory_Click(object sender, EventArgs e)
        {
            String categoryName = default(String);
            Int32 selectedSheetRowIndex = -1;

            try
            {
                selectedSheetRowIndex = DataGridViewHelper.GetIndex(dgvSheets);
                if (selectedSheetRowIndex != -1)
                {
                    if (this.dgvSheetCategoriesFilter.SelectedCells.Count > 0)
                    {
                        categoryName = this.dgvSheetCategoriesFilter.SelectedCells[0].OwningColumn.HeaderText;

                        if (!MDSSController<MDSSModel>.ShiftSheetCategory(SettingsController<Settings>.Settings.Sheets[selectedSheetRowIndex], categoryName, ListOfTExtension.ShiftTypes.Promote))
                        { 
                            throw new ApplicationException(String.Format("Category '{0}' not promoted in '{1}'.", categoryName, Category.SheetCategoryType.Filter.ToString()));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                ViewModel.StopProgressBar(null, String.Format("{0}", ex.Message));
            }
        }

        private void tsmiFilterCategoryDemoteCategory_Click(object sender, EventArgs e)
        {
            String categoryName = default(String);
            Int32 selectedSheetRowIndex = -1;

            try
            {
                selectedSheetRowIndex = DataGridViewHelper.GetIndex(dgvSheets);
                if (selectedSheetRowIndex != -1)
                {
                    if (this.dgvSheetCategoriesFilter.SelectedCells.Count > 0)
                    {
                        categoryName = this.dgvSheetCategoriesFilter.SelectedCells[0].OwningColumn.HeaderText;

                        if (!MDSSController<MDSSModel>.ShiftSheetCategory(SettingsController<Settings>.Settings.Sheets[selectedSheetRowIndex], categoryName, ListOfTExtension.ShiftTypes.Demote))
                        { 
                            throw new ApplicationException(String.Format("Category '{0}' not demoted in '{1}'.", categoryName, Category.SheetCategoryType.Filter.ToString()));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                ViewModel.StopProgressBar(null, String.Format("{0}", ex.Message));
            }
        }

        private void dgvSheetCategoriesFilter_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {

            try
            {
                DataGridView categoryFilterList = (sender as DataGridView);

                if (e.Button == MouseButtons.Right)
                {
                    if (e.RowIndex != -1)
                    {
                        categoryFilterList[e.ColumnIndex, 0].Selected = true;
                        //categoryFilterList.Rows[e.RowIndex].Selected = true;//causes entire row to be selected.
                        categoryFilterList.ContextMenuStrip.Show(categoryFilterList, new Point(e.X, e.Y));
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                ViewModel.StopProgressBar(null, String.Format("{0}", ex.Message));
            }
        }
        #endregion Filters

        #region XCategory
        private void contextMenuStripCategoryX_Opening(object sender, CancelEventArgs e)
        {
            Int32 count = default(Int32);
            ContextMenuStrip contextMenu = sender as ContextMenuStrip;

            try
            {
                DataGridView source = contextMenu.SourceControl as DataGridView;
                count = source.SelectedRows.Count;

                if (count == 1)
                {
                    tsmiXCategoryMoveToFilterCategory.Enabled = dgvSheetCategoriesFilter.Enabled;
                    tsmiXCategoryMoveToYCategory.Enabled = dgvSheetCategoriesY.Enabled;
                    tsmiXCategoryRemoveCategory.Enabled = dgvSheetCategoriesX.Enabled;
                    tsmiXCategoryPromoteCategory.Enabled = dgvSheetCategoriesX.Enabled;
                    tsmiXCategoryDemoteCategory.Enabled = dgvSheetCategoriesX.Enabled;
                }
                else
                {
                    tsmiXCategoryMoveToFilterCategory.Enabled = false;
                    tsmiXCategoryMoveToYCategory.Enabled = false;
                    tsmiXCategoryRemoveCategory.Enabled = false;
                    tsmiXCategoryPromoteCategory.Enabled = false;
                    tsmiXCategoryDemoteCategory.Enabled = false;
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
        }
        
        private void dgvSheetCategoriesX_CellValueNeeded(object sender, DataGridViewCellValueEventArgs e)
        {
            Int32 selectedSheetRowIndex = -1;
            
            try
            {
                selectedSheetRowIndex = DataGridViewHelper.GetIndex(dgvSheets);
                if (selectedSheetRowIndex != -1)
                {
                    DataGridViewHelper.SheetCategoryCellValueNeeded(sender, e, SettingsController<Settings>.Settings.Sheets[selectedSheetRowIndex], Category.SheetCategoryType.X);
                }

            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                //throw;
            }
        }

        private void tsmiXCategoryMoveToFilterCategory_Click(object sender, EventArgs e)
        {
            String categoryName = default(String);
            Int32 selectedSheetRowIndex = -1;

            try
            {
                selectedSheetRowIndex = DataGridViewHelper.GetIndex(dgvSheets);
                if (selectedSheetRowIndex != -1)
                {
                    if (this.dgvSheetCategoriesX.SelectedCells.Count == 1)
                    {
                        categoryName = this.dgvSheetCategoriesX.SelectedCells[0].Value.ToString();

                        if (!MDSSController<MDSSModel>.MoveSheetCategory(SettingsController<Settings>.Settings.Sheets[selectedSheetRowIndex], categoryName, Category.SheetCategoryType.Filter))
                        { 
                            throw new ApplicationException(String.Format("Category '{0}' not moved to '{1}'.", categoryName, Category.SheetCategoryType.Filter.ToString()));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                ViewModel.StopProgressBar(null, String.Format("{0}", ex.Message));
            }
        }

        private void tsmiXCategoryMoveToYCategory_Click(object sender, EventArgs e)
        {
            String categoryName = default(String);
            Int32 selectedSheetRowIndex = -1;

            try
            {
                selectedSheetRowIndex = DataGridViewHelper.GetIndex(dgvSheets);
                if (selectedSheetRowIndex != -1)
                {
                    if (this.dgvSheetCategoriesX.SelectedCells.Count == 1)
                    {
                        categoryName = this.dgvSheetCategoriesX.SelectedCells[0].Value.ToString();

                        if (!MDSSController<MDSSModel>.MoveSheetCategory(SettingsController<Settings>.Settings.Sheets[selectedSheetRowIndex], categoryName, Category.SheetCategoryType.Y))
                        { 
                            throw new ApplicationException(String.Format("Category '{0}' not moved to '{1}'.", categoryName, Category.SheetCategoryType.Y.ToString()));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                ViewModel.StopProgressBar(null, String.Format("{0}", ex.Message));
            }
        }

        private void tsmiXCategoryRemoveCategory_Click(object sender, EventArgs e)
        {
            String categoryName = default(String);
            Int32 selectedSheetRowIndex = -1;

            try
            {
                selectedSheetRowIndex = DataGridViewHelper.GetIndex(dgvSheets);
                if (selectedSheetRowIndex != -1)
                {
                    if (this.dgvSheetCategoriesX.SelectedCells.Count == 1)
                    {
                        categoryName = this.dgvSheetCategoriesX.SelectedCells[0].Value.ToString();

                        if (!MDSSController<MDSSModel>.RemoveSheetCategory(SettingsController<Settings>.Settings.Sheets[selectedSheetRowIndex], categoryName))
                        { 
                            throw new ApplicationException(String.Format("Category '{0}' not removed from '{1}'.", categoryName, Category.SheetCategoryType.X.ToString()));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                ViewModel.StopProgressBar(null, String.Format("{0}", ex.Message));
            }
        }

        private void tsmiXCategoryPromoteCategory_Click(object sender, EventArgs e)
        {
            String categoryName = default(String);
            Int32 selectedSheetRowIndex = -1;
            
            try
            {
                selectedSheetRowIndex = DataGridViewHelper.GetIndex(dgvSheets);
                if (selectedSheetRowIndex != -1)
                {
                    if (this.dgvSheetCategoriesX.SelectedCells.Count == 1)
                    {
                        categoryName = this.dgvSheetCategoriesX.SelectedCells[0].Value.ToString();

                        if (!MDSSController<MDSSModel>.ShiftSheetCategory(SettingsController<Settings>.Settings.Sheets[selectedSheetRowIndex], categoryName, ListOfTExtension.ShiftTypes.Promote))
                        { 
                            throw new ApplicationException(String.Format("Category '{0}' not promoted in '{1}'.", categoryName, Category.SheetCategoryType.X.ToString()));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                ViewModel.StopProgressBar(null, String.Format("{0}", ex.Message));
            }
        }

        private void tsmiXCategoryDemoteCategory_Click(object sender, EventArgs e)
        {
            String categoryName = default(String);
            Int32 selectedSheetRowIndex = -1;

            try
            {
                selectedSheetRowIndex = DataGridViewHelper.GetIndex(dgvSheets);
                if (selectedSheetRowIndex != -1)
                {
                    if (this.dgvSheetCategoriesX.SelectedCells.Count == 1)
                    {
                        categoryName = this.dgvSheetCategoriesX.SelectedCells[0].Value.ToString();

                        if (!MDSSController<MDSSModel>.ShiftSheetCategory(SettingsController<Settings>.Settings.Sheets[selectedSheetRowIndex], categoryName, ListOfTExtension.ShiftTypes.Demote))
                        { 
                            throw new ApplicationException(String.Format("Category '{0}' not demoted in '{1}'.", categoryName, Category.SheetCategoryType.X.ToString()));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                ViewModel.StopProgressBar(null, String.Format("{0}", ex.Message));
            }
        }

        private void dgvSheetCategoriesX_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                DataGridView categoryXList = (sender as DataGridView);

                if (e.Button == MouseButtons.Right)
                {
                    if (e.RowIndex != -1)
                    {
                        categoryXList[0, e.RowIndex].Selected = true;
                        categoryXList.Rows[e.RowIndex].Selected = true;
                        categoryXList.ContextMenuStrip.Show(categoryXList, new Point(e.X, e.Y));
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                ViewModel.StopProgressBar(null, String.Format("{0}", ex.Message));
            }
        }
        #endregion XCategory

        #region YCategory
        private void contextMenuStripCategoryY_Opening(object sender, CancelEventArgs e)
        {
            Int32 count = -1;
            ContextMenuStrip contextMenu = sender as ContextMenuStrip;
            DataGridView source = default(DataGridView);

            try
            {
                source = contextMenu.SourceControl as DataGridView;
                count = source.SelectedCells.Count;

                if (count == 1)
                {
                    tsmiYCategoryMoveToFilterCategory.Enabled = dgvSheetCategoriesFilter.Enabled;
                    tsmiYCategoryMoveToXCategory.Enabled = dgvSheetCategoriesY.Enabled;
                    tsmiYCategoryRemoveCategory.Enabled = dgvSheetCategoriesY.Enabled;
                    tsmiYCategoryPromoteCategory.Enabled = dgvSheetCategoriesY.Enabled;
                    tsmiYCategoryDemoteCategory.Enabled = dgvSheetCategoriesY.Enabled;                
                }
                else
                {
                    tsmiYCategoryMoveToFilterCategory.Enabled = false;
                    tsmiYCategoryMoveToXCategory.Enabled = false;
                    tsmiYCategoryRemoveCategory.Enabled = false;
                    tsmiYCategoryPromoteCategory.Enabled = false;
                    tsmiYCategoryDemoteCategory.Enabled = false;
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
        }

        private void dgvSheetCategoriesY_CellValueNeeded(object sender, DataGridViewCellValueEventArgs e)
        {
            Int32 selectedSheetRowIndex = -1;

            try
            {
                selectedSheetRowIndex = DataGridViewHelper.GetIndex(dgvSheets);
                if (selectedSheetRowIndex != -1)
                {
                    DataGridViewHelper.SheetCategoryCellValueNeeded(sender, e, SettingsController<Settings>.Settings.Sheets[selectedSheetRowIndex], Category.SheetCategoryType.Y);
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                //throw;
            }
        }

        private void tsmiYCategoryMoveToFilterCategory_Click(object sender, EventArgs e)
        {
            String categoryName = default(String);
            Int32 selectedSheetRowIndex = -1;

            try
            {
                selectedSheetRowIndex = DataGridViewHelper.GetIndex(dgvSheets);
                if (selectedSheetRowIndex != -1)
                {
                    if (this.dgvSheetCategoriesY.SelectedCells.Count == 1)
                    {
                        categoryName = this.dgvSheetCategoriesY.SelectedCells[0].Value.ToString();

                        if (!MDSSController<MDSSModel>.MoveSheetCategory(SettingsController<Settings>.Settings.Sheets[selectedSheetRowIndex], categoryName, Category.SheetCategoryType.Filter))
                        { 
                            throw new ApplicationException(String.Format("Category '{0}' not moved to '{1}'.", categoryName, Category.SheetCategoryType.Filter.ToString()));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                ViewModel.StopProgressBar(null, String.Format("{0}", ex.Message));
            }
        }

        private void tsmiYCategoryMoveToXCategory_Click(object sender, EventArgs e)
        {
            String categoryName = default(String);
            Int32 selectedSheetRowIndex = -1;

            try
            {
                selectedSheetRowIndex = DataGridViewHelper.GetIndex(dgvSheets);
                if (selectedSheetRowIndex != -1)
                {
                    if (this.dgvSheetCategoriesY.SelectedCells.Count == 1)
                    {
                        categoryName = this.dgvSheetCategoriesY.SelectedCells[0].Value.ToString();

                        if (!MDSSController<MDSSModel>.MoveSheetCategory(SettingsController<Settings>.Settings.Sheets[selectedSheetRowIndex], categoryName, Category.SheetCategoryType.X))
                        { 
                            throw new ApplicationException(String.Format("Category '{0}' not moved to '{1}'.", categoryName, Category.SheetCategoryType.X.ToString()));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                ViewModel.StopProgressBar(null, String.Format("{0}", ex.Message));
            }
        }

        private void tsmiYCategoryRemoveCategory_Click(object sender, EventArgs e)
        {
            String categoryName = default(String);
            Int32 selectedSheetRowIndex = -1;

            try
            {
                selectedSheetRowIndex = DataGridViewHelper.GetIndex(dgvSheets);
                if (selectedSheetRowIndex != -1)
                {
                    if (this.dgvSheetCategoriesY.SelectedCells.Count == 1)
                    {
                        categoryName = this.dgvSheetCategoriesY.SelectedCells[0].Value.ToString();

                        if (!MDSSController<MDSSModel>.RemoveSheetCategory(SettingsController<Settings>.Settings.Sheets[selectedSheetRowIndex], categoryName))
                        { 
                            throw new ApplicationException(String.Format("Category '{0}' not removed from '{1}'.", categoryName, Category.SheetCategoryType.Y.ToString()));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                ViewModel.StopProgressBar(null, String.Format("{0}", ex.Message));
            }
        }

        private void tsmiYCategoryPromoteCategory_Click(object sender, EventArgs e)
        {
            String categoryName = default(String);
            Int32 selectedSheetRowIndex = -1;

            try
            {
                selectedSheetRowIndex = DataGridViewHelper.GetIndex(dgvSheets);
                if (selectedSheetRowIndex != -1)
                {
                    if (this.dgvSheetCategoriesY.SelectedCells.Count == 1)
                    {
                        categoryName = this.dgvSheetCategoriesY.SelectedCells[0].Value.ToString();

                        if (!MDSSController<MDSSModel>.ShiftSheetCategory(SettingsController<Settings>.Settings.Sheets[selectedSheetRowIndex], categoryName, ListOfTExtension.ShiftTypes.Promote))
                        { 
                            throw new ApplicationException(String.Format("Category '{0}' not promoted in '{1}'.", categoryName, Category.SheetCategoryType.Y.ToString()));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                ViewModel.StopProgressBar(null, String.Format("{0}", ex.Message));
            }
        }

        private void tsmiYCategoryDemoteCategory_Click(object sender, EventArgs e)
        {
            String categoryName = default(String);
            Int32 selectedSheetRowIndex = -1;

            try
            {
                selectedSheetRowIndex = DataGridViewHelper.GetIndex(dgvSheets);
                if (selectedSheetRowIndex != -1)
                {
                    if (this.dgvSheetCategoriesY.SelectedCells.Count == 1)
                    {
                        categoryName = this.dgvSheetCategoriesY.SelectedCells[0].Value.ToString();

                        if (!MDSSController<MDSSModel>.ShiftSheetCategory(SettingsController<Settings>.Settings.Sheets[selectedSheetRowIndex], categoryName, ListOfTExtension.ShiftTypes.Demote))
                        { 
                            throw new ApplicationException(String.Format("Category '{0}' not demoted in '{1}'.", categoryName, Category.SheetCategoryType.Y.ToString()));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                ViewModel.StopProgressBar(null, String.Format("{0}", ex.Message));
            }
        }

        private void dgvSheetCategoriesY_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {

            try
            {
                DataGridView categoryYList = (sender as DataGridView);

                if (e.Button == MouseButtons.Right)
                {
                    if (e.RowIndex != -1)
                    {
                        categoryYList[e.ColumnIndex, 0].Selected = true;
                        //categoryYList.Rows[e.RowIndex].Selected = true;//causes entire row to be selected.
                        categoryYList.ContextMenuStrip.Show(categoryYList, new Point(e.X, e.Y));
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                ViewModel.StopProgressBar(null, String.Format("{0}", ex.Message));
            }
        }
        #endregion YCategory
        #endregion Sheet Categories
        #endregion Control Events
        #endregion Events

        #region Methods

        #region FormAppBase
        protected void InitViewModel()
        {
            try
            {
                //subscribe view to model notifications
                ModelController<MDSSModel>.Model.PropertyChanged += PropertyChangedEventHandlerDelegate;

                FileDialogInfo settingsFileDialogInfo =
                    new FileDialogInfo
                    (
                        SettingsController<Settings>.FILE_NEW,
                        null,
                        null,
                        /*SettingsController<Settings>.*/Settings.FileTypeExtension,
                        /*SettingsController<Settings>.*/Settings.FileTypeDescription,
                        /*SettingsController<Settings>.*/Settings.FileTypeName,
                        new String[] 
                        { 
                            "XML files (*.xml)|*.xml", 
                            "All files (*.*)|*.*" 
                        },
                        false,
                        default(Environment.SpecialFolder),
                        Environment.GetFolderPath(Environment.SpecialFolder.Personal).WithTrailingSeparator()
                    );

                //set dialog caption
                settingsFileDialogInfo.Title = this.Text;

                //class to handle standard behaviors
                ViewModelController<Bitmap, ModelViewModel>.New
                (
                    ViewName,
                    new ModelViewModel
                    (
                        this.PropertyChangedEventHandlerDelegate,
                        new Dictionary<String, Bitmap>() 
                        { 
                            { "Above", Resources.Above }, 
                            { "Below", Resources.Below }, 
                            { "Bottom", Resources.Bottom }, 
                            { "Copy", Resources.Copy },
                            { "Delete", Resources.Delete }, 
                            { "FastForward", Resources.FastForward }, 
                            { "FastRewind", Resources.FastRewind }, 
                            { "Forward", Resources.Forward }, 
                            { "New", Resources.New }, 
                            { "Open", Resources.Open },
                            { "Print", Resources.Print },
                            { "Properties", Resources.Properties },
                            { "Rewind", Resources.Rewind }, 
                            { "Save", Resources.Save },
                            { "Top", Resources.Top }//, 
                            //{ "BoxEmpty", Resources.BoxEmpty }, 
                            //{ "BoxFull", Resources.BoxFull }, 
                            //{ "Download", Resources.Download }, 
                            //{ "Upload", Resources.Upload }, 
                            //{ "Package", Resources.Package }, 
                            //{ "List", Resources.List }, 
                            //{ "ListSplitAbove", Resources.ListSplitAbove }, 
                            //{ "ListSplitBelow", Resources.ListSplitBelow }, 
                            //{ "Network", Resources.Network }, 
                            //{ "Scan", Resources.Scan }, 
                            //{ "Search", Resources.Search }, 
                            //{ "RotateCCW", Resources.RotateCCW }, 
                            //{ "RotateCW", Resources.RotateCW } 
                        },
                        settingsFileDialogInfo,
                        this
                    )
                );
                ViewModel = ViewModelController<Bitmap, ModelViewModel>.ViewModel[ViewName];

                BindFormUi();

                //Init config parameters
                if (!LoadParameters())
                {
                    throw new Exception(String.Format("Unable to load config file parameter(s)."));
                }

                //DEBUG:filename coming in is being converted/passed as DOS 8.3 format equivalent
                //Load
                if ((SettingsController<Settings>.FilePath == null) || (SettingsController<Settings>.Filename == SettingsController<Settings>.FILE_NEW))
                {
                    //NEW
                    ViewModel.FileNew();
                }
                else
                {
                    //OPEN
                    ViewModel.FileOpen(false);
                }

    #if debug
                //debug view
                menuEditProperties_Click(sender, e);
    #endif

                //Display dirty state
                ModelController<MDSSModel>.Model.Refresh();
            }
            catch (Exception ex)
            {
                if (ViewModel != null)
                {
                    ViewModel.ErrorMessage = ex.Message;
                }
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
            }
        }

        protected void DisposeSettings()
        {
            //save user and application settings 
            Properties.Settings.Default.Save();

            if (SettingsController<Settings>.Settings.Dirty)
            {
                //prompt before saving
                DialogResult dialogResult = MessageBox.Show("Save changes?", this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                switch (dialogResult)
                {
                    case DialogResult.Yes:
                        {
                            //SAVE
                            ViewModel.FileSave();

                            break;
                        }
                    case DialogResult.No:
                        {
                            break;
                        }
                    default:
                        {
                            throw new InvalidEnumArgumentException();
                        }
                }
            }

            //unsubscribe from model notifications
            ModelController<MDSSModel>.Model.PropertyChanged -= PropertyChangedEventHandlerDelegate;
        }

        protected void _Run()
        {
        }
        #endregion FormAppBase

        #region Utility
        /// <summary>
        /// Bind Settings controls to SettingsController
        /// </summary>
        private void BindFormUi()
        {
            try
            {
                //Form
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
        }

        /// <summary>
        /// Bind Settings controls to SettingsController
        /// </summary>
        private void BindModelUi()
        {
            try
            {
                //Settings

                //Sheet
                sheetBindingSource.DataSource = SettingsController<Settings>.Settings.Sheets;

                //Sheet.Categories
                categoryBindingSource.DataSource = sheetBindingSource;
                categoryBindingSource.DataMember = "Categories";

                //Sheet.Categories.Items
                itemsBindingSource.DataSource = categoryBindingSource;
                itemsBindingSource.DataMember = "Items";

                //Sheet.Formulae
                formulaBindingSource.DataSource = sheetBindingSource;
                formulaBindingSource.DataMember = "Formulae";

                //Sheet is virtual-mode datagridview bound with CellNeeded / CellPushed events.
                RefreshSheet(this.dgvSheets, this.dgvSheet);

                //Sheet-Categories are virtual datagridview generated from Categories
                RefereshSheetCategories(this.dgvSheets, this.dgvSheetCategoriesFilter, this.dgvSheetCategoriesX, this.dgvSheetCategoriesY);
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
        }

        /// <summary>
        /// Check for current sheet and performs a grid render.
        /// </summary>
        /// <param name="sheetListGrid"></param>
        /// <param name="sheetGrid"></param>
        private void RefreshSheet(DataGridView sheetListGrid, DataGridView sheetGrid)
        {
            Int32 selectedSheetRowIndex = -1;
            Sheet sheet = default(Sheet);
            
            try
            {   
                selectedSheetRowIndex = DataGridViewHelper.GetIndex(sheetListGrid);
                if (selectedSheetRowIndex != -1)
                {
                    //Get bound sheet directly; if null, it will behave like 'else' case.
                    sheet = sheetListGrid.Rows[selectedSheetRowIndex].DataBoundItem as Sheet; 

                    // Add columns and rows to the DataGridView.
                    DataGridViewHelper.RenderSheet(sheetGrid, sheet);
                }
                else
                {
                    // Clear columns and rows in the DataGridView.
                    DataGridViewHelper.RenderSheet(sheetGrid, (Sheet)null);
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
        }

        /// <summary>
        /// Check for current sheet and performs category render.
        /// </summary>
        /// <param name="sheetListGrid"></param>
        /// <param name="filterCategory"></param>
        /// <param name="xCategory"></param>
        /// <param name="yCategory"></param>
        private void RefereshSheetCategories(DataGridView sheetListGrid, DataGridView filterCategory, DataGridView xCategory, DataGridView yCategory)
        {
            Int32 selectedRowIndex = -1;
            Sheet sheet = default(Sheet);
            
            try
            {
                selectedRowIndex = DataGridViewHelper.GetIndex(sheetListGrid);
                if (selectedRowIndex != -1)
                {
                    //Get bound sheet directly; if null, it will behave like 'else' case.
                    sheet = sheetListGrid.Rows[selectedRowIndex].DataBoundItem as Sheet;

                    DataGridViewHelper.RenderSheetCategory(this.dgvSheetCategoriesFilter, sheet, Category.SheetCategoryType.Filter);
                    DataGridViewHelper.RenderSheetCategory(this.dgvSheetCategoriesX, sheet, Category.SheetCategoryType.X);
                    DataGridViewHelper.RenderSheetCategory(this.dgvSheetCategoriesY, sheet, Category.SheetCategoryType.Y);
                }
                else
                {
                    DataGridViewHelper.RenderSheetCategory(this.dgvSheetCategoriesFilter, (Sheet)null, Category.SheetCategoryType.Filter);
                    DataGridViewHelper.RenderSheetCategory(this.dgvSheetCategoriesX, (Sheet)null, Category.SheetCategoryType.X);
                    DataGridViewHelper.RenderSheetCategory(this.dgvSheetCategoriesY, (Sheet)null, Category.SheetCategoryType.Y);
                }

            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
        }

        /// <summary>
        /// Apply Settings to viewer.
        /// </summary>
        private void ApplySettings()
        {
            try
            {
                _ValueChangedProgrammatically = true;

                //apply settings that have databindings
                BindModelUi();

                //apply settings that shouldn't use databindings

                //apply settings that can't use databindings
                Text = Path.GetFileName(SettingsController<Settings>.Filename) + " - " + ViewName;

                //apply settings that don't have databindings
                ViewModel.DirtyIconIsVisible = (SettingsController<Settings>.Settings.Dirty);

                _ValueChangedProgrammatically = false;
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
        }

        /// <summary>
        /// Set function button and menu to enable value, and cancel button to opposite.
        /// For now, do only disabling here and leave enabling based on biz logic 
        ///  to be triggered by refresh?
        /// </summary>
        /// <param name="functionButton"></param>
        /// <param name="functionMenu"></param>
        /// <param name="cancelButton"></param>
        /// <param name="enable"></param>
        internal void SetFunctionControlsEnable
        (
            Button functionButton,
            ToolStripButton functionToolbarButton,
            ToolStripMenuItem functionMenu,
            Button cancelButton,
            Boolean enable
        )
        {
            try
            {
                //stand-alone button
                if (functionButton != null)
                {
                    functionButton.Enabled = enable;
                }

                //toolbar button
                if (functionToolbarButton != null)
                {
                    functionToolbarButton.Enabled = enable;
                }

                //menu item
                if (functionMenu != null)
                {
                    functionMenu.Enabled = enable;
                }

                //stand-alone cancel button
                if (cancelButton != null)
                {
                    cancelButton.Enabled = !enable;
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
            }
        }

        /// <summary>
        /// Invoke any delegate that has been registered 
        ///  to cancel a long-running background process.
        /// </summary>
        private void InvokeActionCancel()
        {
            try
            {
                //execute cancellation hook
                if (cancelDelegate != null)
                {
                    cancelDelegate();
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
            }
        }

        /// <summary>
        /// Load from app config; override with command line if present
        /// </summary>
        /// <returns></returns>
        private Boolean LoadParameters()
        {
            Boolean returnValue = default(Boolean);
#if USE_CONFIG_FILEPATH
            String _settingsFilename = default(String);
#endif

            try
            {
                if ((Program.Filename != null) && (Program.Filename != SettingsController<Settings>.FILE_NEW))
                {
                    //got filename from command line
                    //DEBUG:filename coming in is being converted/passed as DOS 8.3 format equivalent
                    if (!RegistryAccess.ValidateFileAssociation(Application.ExecutablePath, "." + /*SettingsController<Settings>.*/Settings.FileTypeExtension))
                    {
                        throw new ApplicationException(String.Format("Settings filename not associated: '{0}'.\nCheck SettingsFilename in app config file.", Program.Filename));
                    }
                    //it passed; use value from command line
                }
                else
                {
#if USE_CONFIG_FILEPATH
                    //get filename from app.config
                    if (!Configuration.ReadString("SettingsFilename", out _settingsFilename))
                    {
                        throw new ApplicationException(String.Format("Unable to load SettingsFilename: {0}", "SettingsFilename"));
                    }
                    if ((_settingsFilename == null) || (_settingsFilename == SettingsController<Settings>.FILE_NEW))
                    {
                        throw new ApplicationException(String.Format("Settings filename not set: '{0}'.\nCheck SettingsFilename in app config file.", _settingsFilename));
                    }
                    //use with the supplied path
                    SettingsController<Settings>.Filename = _settingsFilename;

                    if (Path.GetDirectoryName(_settingsFilename) == String.Empty)
                    {
                        //supply default path if missing
                        SettingsController<Settings>.Pathname = Environment.GetFolderPath(Environment.SpecialFolder.Personal).WithTrailingSeparator();
                    }
#endif
                }

                returnValue = true;
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
                //throw;
            }
            return returnValue;
        }

        private void BindSizeAndLocation()
        {
            //Note:Size must be done after InitializeComponent(); do Location this way as well.--SJS
            this.DataBindings.Add(new System.Windows.Forms.Binding("Location", global::MultiDimensionalSpreadsheet.Properties.Settings.Default, "Location", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.DataBindings.Add(new System.Windows.Forms.Binding("ClientSize", global::MultiDimensionalSpreadsheet.Properties.Settings.Default, "Size", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.ClientSize = global::MultiDimensionalSpreadsheet.Properties.Settings.Default.Size;
            this.Location = global::MultiDimensionalSpreadsheet.Properties.Settings.Default.Location;
        }
        #endregion Utility
        #endregion Methods
    }
}