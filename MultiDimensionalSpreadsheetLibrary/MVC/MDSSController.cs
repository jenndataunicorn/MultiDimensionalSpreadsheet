using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Ssepan.Application;
using Ssepan.Collections;
using Ssepan.Utility;
using System.Diagnostics;
using System.Reflection;

namespace MultiDimensionalSpreadsheetLibrary
{
    /// <summary>
    /// This is the MVC Controller
    /// </summary>
    public class MDSSController<TModel> : 
        ModelController<TModel>
        where TModel :
            class,
            IModel,
            new()
    {
        #region Declarations
        #endregion Declarations

        #region Constructors
        public MDSSController()
        {
        }
        #endregion Constructors

        #region Properties
        //Note:TModel Model exists in base
        #endregion Properties

        #region Methods

        /// <summary>
        /// Select a category item for a filter category.
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="category"></param>
        public static Boolean SelectFilterCategoryItem(Sheet sheet, String categoryName)
        {
            Boolean returnValue = default(Boolean);
            CategoryItem categoryItem = default(CategoryItem);

            try
            {
                //User must select a Category Item.
                categoryItem = SelectCategoryItem(sheet, categoryName, "Select a Category Item to Filter Sheet Cells.", -1, true);

                if (categoryItem != null)
                {
                    _ValueChanging = true;
                    
                    //move it
                    sheet.SetFilterCategoryItem(categoryItem);

                    _ValueChanging = false;

                    //refresh
                    //Refresh();

                    returnValue = true;
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                _ValueChanging = false;
                throw;

            }
            return returnValue;
        }

        /// <summary>
        /// Search until end or until another category with same sheet-category-type found; swap in latter case.
        /// </summary>
        /// <param name="categoryName"></param>
        /// <param name="categoryShiftType"></param>
        public static Boolean ShiftSheetCategory(Sheet sheet, String categoryName, ListOfTExtension.ShiftTypes shiftType)
        {
            Boolean returnValue = default(Boolean);

            try
            {
                _ValueChanging = true;

                sheet.Categories.ShiftListItem<OrderedEquatableBindingList<Category>, Category>
                    (
                        shiftType,
                        (item => item.Name == categoryName), //match on Category Name property
                        (item, itemSwap) => item.CategoryType == itemSwap.CategoryType //match on item Category Type)
                    );
                
                _ValueChanging = false;

                //refresh
                //Refresh();

                returnValue = true;
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                _ValueChanging = false;
                throw;
            }
            return returnValue;
        }

        /// <summary>
        /// Add an unassigned category to one of the sheet categories.
        /// </summary>
        /// <param name="category"></param>
        /// <param name="source"></param>
        /// <param name="destination"></param>
        public static Boolean AddSheetCategory(Sheet sheet, String categoryName, Category.SheetCategoryType categoryType)
        {
            Boolean returnValue = default(Boolean);
            CategoryItem categoryItem = default(CategoryItem);

            try
            {
                //User must select a Category Item.
                categoryItem = SelectCategoryItem(sheet, categoryName, "Select a Category Item to Add to existing Sheet Cells.", 0, false);

                if (categoryItem != null)
                {
                    _ValueChanging = true;

                    //move it
                    sheet.AssignCategory(categoryName, categoryType, categoryItem);

                    _ValueChanging = false;

                    //refresh
                    //Refresh();

                    returnValue = true;
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                _ValueChanging = false;
                throw;

            }
            return returnValue;
        }

        /// <summary>
        /// Changed an assigned category to another sheet category.
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="category"></param>
        /// <param name="categoryType"></param>
        public static Boolean MoveSheetCategory(Sheet sheet, String categoryName, Category.SheetCategoryType categoryType)
        {
            Boolean returnValue = default(Boolean);

            try
            {
                _ValueChanging = true;
                
                //move it
                sheet.AssignCategory(categoryName, categoryType, (CategoryItem)null);

                _ValueChanging = false;

                //refresh
                //Refresh();

                returnValue = true;
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                _ValueChanging = false;
                throw;

            }
            return returnValue;
        }

        /// <summary>
        /// Remove an assigned category from one of the sheet categories.
        /// </summary>
        /// <param name="categoryName"></param>
        public static Boolean RemoveSheetCategory(Sheet sheet, String categoryName)
        {
            Boolean returnValue = default(Boolean);
            CategoryItem categoryItem = default(CategoryItem);

            try
            {
                //User must select a Category Item.
                categoryItem = SelectCategoryItem(sheet, categoryName, "Select a Category Item to Remove from existing Sheet Cells. Note: Cells without selected Item will be permanently deleted.", 0, false);

                if (categoryItem != null)
                {
                    _ValueChanging = true;
                    
                    //move it
                    sheet.AssignCategory(categoryName, Category.SheetCategoryType.None, categoryItem);

                    _ValueChanging = false;

                    //refresh
                    //Refresh();

                    returnValue = true;
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                _ValueChanging = false;
                throw;

            }
            return returnValue;
        }

        /// <summary>
        /// Search until end or until another category item found; swap in latter case.
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="categoryName"></param>
        /// <param name="categoryItemName"></param>
        /// <param name="categoryShiftType"></param>
        public static Boolean ShiftSheetCategoryItem(Sheet sheet, String categoryName, String categoryItemName, ListOfTExtension.ShiftTypes shiftType)
        {
            Boolean returnValue = default(Boolean);
            Category category = default(Category);

            try
            {
                _ValueChanging = true;

                //find item's parent
                category = sheet.Categories.Find(c => c.Name == categoryName);
                if (category == null)
                {
                    throw new ArgumentException(String.Format("Unable to find Category '{0}'.", categoryName));
                }

                category.Items.ShiftListItem<OrderedEquatableBindingList<CategoryItem>, CategoryItem>
                    (
                        shiftType,
                        (item) => item.Name == categoryItemName, //match on Category Item Name property
                        (item, itemSwap) => true //match on any Category Item)
                    );

                _ValueChanging = false;

                //refresh
                //Refresh();

                returnValue = true;
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                _ValueChanging = false;
                throw;
            }
            return returnValue;
        }

        /// <summary>
        /// Search until end or until another formula found; swap in latter case.
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="formulaName"></param>
        /// <param name="categoryShiftType"></param>
        public static Boolean ShiftSheetFormula(Sheet sheet, String formulaName, ListOfTExtension.ShiftTypes shiftType)
        {
            Boolean returnValue = default(Boolean);

            try
            {
                _ValueChanging = true;

                sheet.Formulae.ShiftListItem<OrderedEquatableBindingList<Formula>, Formula>
                    (
                        shiftType,
                        (item => item.Value == formulaName), //match on Formula Name property
                        (item, itemSwap) => true //match on any Formula)
                    );

                _ValueChanging = false;

                //refresh
                //Refresh();

                returnValue = true;
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                _ValueChanging = false;
                throw;
            }
            return returnValue;
        }

        /// <summary>
        /// Display dialog to allow user to select a category item from specified category.
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="categoryName"></param>
        /// <param name="instructions"></param>
        /// <param name="defaultSelectedIndex"></param>
        /// <param name="useCategorySelectedItemIndex"></param>
        /// <returns></returns>
        public static CategoryItem SelectCategoryItem(Sheet sheet, String categoryName, String instructions, Int32 defaultSelectedIndex, Boolean useCategorySelectedItemIndex)
        {
            CategoryItem returnValue = default(CategoryItem);
            Category category = default(Category);
            String categoryItemName = default(String);

            try
            {
                //User must select a Category Item.
                
                // Display the form as a modal dialog box.
                SelectDialog dialog = new SelectDialog();

                //set dialog icon
                dialog.Icon = MultiDimensionalSpreadsheetLibrary.Properties.Resources.MultiDimensionalSpreadsheet;

                //set dialog title
                dialog.Text = "Select a Category Item";
                
                //set instructions
                dialog.lblInstructions.Text = instructions;
                
                //set category name
                category = sheet.Categories.Find(c => c.Name == categoryName);
                dialog.lblCategoryName.Text = category.Name;

                // Add  category items to the listbox
                dialog.ddlCategoryItems.DataSource = category.Items;
                dialog.ddlCategoryItems.DisplayMember = "Name";
                dialog.ddlCategoryItems.ValueMember = "Name";

                //pre-set selection
                if (useCategorySelectedItemIndex)
                {
                    dialog.ddlCategoryItems.SelectedIndex = category.SelectedItemIndex;
                }
                else if (defaultSelectedIndex != -1)
                {
                    dialog.ddlCategoryItems.SelectedIndex = defaultSelectedIndex;
                }

                //show dialog;
                dialog.ShowDialog();

                // Determine if the OK button was clicked on the dialog box.
                if (dialog.DialogResult == DialogResult.OK)
                {
                    categoryItemName = ((CategoryItem)dialog.ddlCategoryItems.SelectedItem).Name;
                    dialog.Dispose();
                    returnValue = category.Items.Find(ci => ci.Name == categoryItemName);
                }
                else
                {
                    dialog.Dispose();
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
            return returnValue;
        }

        //public new static void Refresh()
        //{
        //    MDSSController<MDSSModel>.Model.IsChanged = true;//Value doesn't matter; fire a changed event;
        //}
        #endregion Methods
    }
}