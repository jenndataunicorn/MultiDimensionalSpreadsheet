using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MultiDimensionalSpreadsheetLibrary;
using MultiDimensionalSpreadsheetLibrary.MVC;
using Ssepan.Application;
using Ssepan.Application.MVC;
using Ssepan.Collections;
using Ssepan.Utility;
using System.Diagnostics;
using System.Reflection;

namespace MultiDimensionalSpreadsheetLibrary.UI
{
    /// <summary>
    /// 
    /// </summary>
    public static class DataGridViewHelper
    {
        #region Declarations
        private static Boolean _IsReshapingGrid = default(Boolean);
        #endregion Declarations

        #region Public Methods
        /// <summary>
        /// Clears grid and builds a new one. Adds rows and columns, and formats them.
        /// Pass sheet explicitly.
        /// </summary>
        /// <param name="grid"></param>
        public static void RenderSheet(DataGridView grid, Sheet sheet)
        {
            try
            {
                BuildSheet(grid, sheet);
                FormatSheet(grid, sheet);
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
                throw;
            }
        }
        
        /// <summary>
        /// Clears grid and builds a new one. Adds rows and columns, and formats them.
        /// Pass sheet name and sheet will be looked up.
        /// </summary>
        /// <param name="grid"></param>
        public static void RenderSheet(DataGridView grid, String sheetName)
        {
            try
            {
                RenderSheet(grid, SettingsController<Settings>.Settings.Sheets.Find(s => s.Name == sheetName));
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
                throw;
            }
        }
        
        /// <summary>
        /// Clears grid and builds a new one. Adds rows and columns, and formats them.
        /// Pass sheet index and sheet will be looked up.
        /// </summary>
        /// <param name="grid"></param>
        public static void RenderSheet(DataGridView grid, Int32 sheetIndex)
        {
            try
            {
                RenderSheet(grid, SettingsController<Settings>.Settings.Sheets[sheetIndex]);
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
                throw;
            }
        }

        /// <summary>
        /// Method to call from Event handler for Sheet DataGridView CellValueNeeded event.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public static void SheetCellValueNeeded(object sender, DataGridViewCellValueEventArgs e, Sheet sheet)
        {
            Int32 cellIndex = -1;
            String cellValue = String.Empty;

            try
            {
                //identify cell type from range given by row/column
                switch (sheet.GetCellTypeAtCoordinates(e.RowIndex, e.ColumnIndex))
                {
                    case Sheet.CellTypes.Empty:
                        {
                            //if empty range, ignore
                            #if debugcell
                            e.Value = String.Format("({0},{1});", e.RowIndex, e.ColumnIndex);
                            #endif

                            break;
                        }
                    case Sheet.CellTypes.XCategory:
                        {
                            //if x category range, identify category-item @ x, y 
                            e.Value = sheet.GetCategoryItemNameAtCoordinates(e.RowIndex, e.ColumnIndex, true);

                            break;
                        }
                    case Sheet.CellTypes.YCategory:
                        {
                            //if y category range, identify category-item @ x, y 
                            e.Value = sheet.GetCategoryItemNameAtCoordinates(e.RowIndex, e.ColumnIndex, true);

                            break;
                        }
                    case Sheet.CellTypes.Value:
                        {

                            cellIndex = sheet.GetValueIndexAtCoordinates(e.RowIndex, e.ColumnIndex);
                            if (cellIndex != -1)
                            {
                                //default to whatever was in sheet cell; may be overridden with formula value(s)
                                cellValue = sheet.Cells[cellIndex].Value;
                                
                                //check for formulae here; apply value to sheet cell if applicable (last one 'wins' )
                                foreach (Formula formula in sheet.Formulae)
                                {
                                    if (formula.Complete())
                                    {
                                        if (formula.IsApplicableFormula(sheet.Cells[cellIndex]))
                                        {
                                            //call assigner operation
                                            if (formula.Run(sheet.Cells[cellIndex]))
                                            {
                                                cellValue = sheet.Cells[cellIndex].Value;
                                            }
                                            #if debugcellformulaaddress
                                            var addressName = Sheet.GetFormattedCellAddress(sheet.Cells[cellIndex].CategoryItems.ToList<CategoryItem>());
                                            Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
                                            #else
                                            #endif
                                        }
                                        else
                                        {
                                            #if debugcellformulaaddress
                                            var addressName = Sheet.GetFormattedCellAddress(sheet.Cells[cellIndex].CategoryItems.ToList<CategoryItem>());
                                            Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
                                            #else
                                            #endif
                                        }
                                    }
                                }
                            }
                            else
                            {
                                throw new ArgumentOutOfRangeException(String.Format("Cell index ({0}) was out of range: ({1}...{2})", cellIndex, 0, sheet.Cells.Count - 1));
                            }

                            //update display with value
                            #if debugcellvalue
                            e.Value = String.Format("({0},{1}):{2}", e.RowIndex, e.ColumnIndex, cellValue); 
                            #else
                            e.Value = cellValue;
                            #endif

                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
                //throw;
            }
        }

        /// <summary>
        /// Method to call from Event handler for Sheet DataGridView CellValuePushed event.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// <param name="sheet"></param>
        public static void SheetCellValuePushed(object sender, DataGridViewCellValueEventArgs e, Sheet sheet)
        {
            try
            {
                //identify cell type from range given by row/column
                switch (sheet.GetCellTypeAtCoordinates(e.RowIndex, e.ColumnIndex))
                {
                    case Sheet.CellTypes.Empty:
                        {
                            //if empty range, ignore
                            #if debugcell
                            e.Value = String.Format("({0},{1});", e.RowIndex, e.ColumnIndex);
                            #endif

                            break;
                        }
                    case Sheet.CellTypes.XCategory:
                        {
                            //if x category range, identify category-item @ x, y 
                            e.Value = sheet.GetCategoryItemNameAtCoordinates(e.RowIndex, e.ColumnIndex, true);

                            break;
                        }
                    case Sheet.CellTypes.YCategory:
                        {
                            //if y category range, identify category-item @ x, y 
                            e.Value = sheet.GetCategoryItemNameAtCoordinates(e.RowIndex, e.ColumnIndex, true);

                            break;
                        }
                    case Sheet.CellTypes.Value:
                        {
                            Int32 cellIndex = -1;
                            SheetCell cell = default(SheetCell);

                            cellIndex = sheet.GetValueIndexAtCoordinates(e.RowIndex, e.ColumnIndex);
                            if (cellIndex != -1)
                            {
                                cell = sheet.Cells[cellIndex];
                            }
                            else
                            {
                                throw new ArgumentOutOfRangeException(String.Format("Cell index ({0}) was out of range: ({1}...{2})", cellIndex, 0, sheet.Cells.Count - 1));
                            }

                            //update cell with display value.
                            cell.Value = e.Value.ToString();

                            break;
                        }
                }

            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
                throw;
            }
        }
        
        /// <summary>
        /// Clears grid and builds a new one. Adds rows and columns, and formats them.
        /// Pass sheet explicitly.
        /// </summary>
        /// <param name="grid"></param>
        public static void RenderSheetCategory(DataGridView grid, Sheet sheet, Category.SheetCategoryType sheetCategoryType)
        {
            try
            {
                BuildSheetCategory(grid, sheet, sheetCategoryType);
                FormatSheetCategory(grid, sheet, sheetCategoryType);
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
                throw;
            }
        }

        /// <summary>
        /// Clears grid and builds a new one. Adds rows and columns, and formats them.
        /// Pass sheet name and sheet will be looked up.
        /// </summary>
        /// <param name="grid"></param>
        public static void RenderSheetCategory(DataGridView grid, String sheetName, Category.SheetCategoryType sheetCategoryType)
        {
            try
            {
                RenderSheetCategory(grid, SettingsController<Settings>.Settings.Sheets.Find(s => s.Name == sheetName), sheetCategoryType);
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
                throw;
            }
        }

        /// <summary>
        /// Clears grid and builds a new one. Adds rows and columns, and formats them.
        /// Pass sheet index and sheet will be looked up.
        /// </summary>
        /// <param name="grid"></param>
        public static void RenderSheetCategory(DataGridView grid, Int32 sheetIndex, Category.SheetCategoryType sheetCategoryType)
        {
            try
            {
                RenderSheetCategory(grid, SettingsController<Settings>.Settings.Sheets[sheetIndex], sheetCategoryType);
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
        }

        /// <summary>
        /// Method to call from Event handler for Sheet-Category DataGridView CellValueNeeded event.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public static void SheetCategoryCellValueNeeded(object sender, DataGridViewCellValueEventArgs e, Sheet sheet, Category.SheetCategoryType sheetCategoryType)
        {
            DataGridView sheetCategory = (sender as DataGridView);
            List<Category> filterCategory = default(List<Category>);

            try
            {
                if (sheetCategoryType == Category.SheetCategoryType.Filter)
                {
                    if (!_IsReshapingGrid)
                    {
                        //only works if header is displayed
                        //category item name; category name is put in header when grid is re-configured. SelectedItemIndex must be valid.
                        filterCategory = sheet.CategoryFilters();
                        e.Value = filterCategory[e.ColumnIndex].Items[filterCategory[e.ColumnIndex].SelectedItemIndex].Name;
                        ////use if header is not displayed
                        ////category name
                        //e.Value = sheet.CategoryFilters()[e.ColumnIndex].Name;
                    }
                    else
                    {
                        e.Value = null;
                    }
                }
                else if (sheetCategoryType == Category.SheetCategoryType.X)
                {
                    if (!_IsReshapingGrid)
                    {
                        //category name
                        e.Value = sheet.CategoryX()[e.RowIndex].Name;
                    }
                    else
                    {
                        e.Value = null;
                    }
                }
                else if (sheetCategoryType == Category.SheetCategoryType.Y)
                {
                    if (!_IsReshapingGrid)
                    {
                        //category name
                        e.Value = sheet.CategoryY()[e.ColumnIndex].Name;
                    }
                    else
                    { 
                        e.Value = null;
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
                throw;
            }
        }

        /// <summary>
        /// Look up a selected row number in the DataGridView, first by checking for selected rows, then by checking row of selected cells
        /// </summary>
        /// <param name="sheetListGrid"></param>
        /// <returns></returns>
        public static Int32 GetIndex(DataGridView grid)
        {
            Int32 returnValue = -1;
            try
            {//TODO:handle apparent 'new' row selection when there are no actual rows
                //look for selected row first
                if (grid.SelectedRows.Count == 1)
                {
                    returnValue = grid.SelectedRows[0].Index;
                }
                if (returnValue == -1)
                {
                    //look for selected cell next
                    if (grid.SelectedCells.Count == 1)
                    {
                        returnValue = grid.SelectedCells[0].OwningRow.Index;
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

        #endregion Public Methods

        #region Private Methods
        /// <summary>
        /// Clears grid and builds a new one. Adds rows and columns, and formats them.
        /// </summary>
        /// <param name="grid"></param>
        private static void BuildSheet(DataGridView grid, Sheet sheet)
        {
            try
            {
                ConfigureVirtualDataGridView(grid);

                // Add columns and rows to the DataGridView.
                grid.Rows.Clear();
                grid.Columns.Clear();

                if (sheet != null)
                {
                    ConfigureVirtualDataGridViewTextCells(grid, sheet.ColumnCount, sheet.RowCount, (i => ""));
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
                throw;
            }
        }

        /// <summary>
        /// Formats grid.
        /// </summary>
        /// <param name="grid"></param>
        /// <param name="sheet"></param>
        private static void FormatSheet(DataGridView grid, Sheet sheet)
        {
            try
            {
                if (sheet != null)
                {
                    //format cells
                    for (int x = 0; x < sheet.ColumnCount; x++)
                    {
                        for (int y = 0; y < sheet.RowCount; y++)
                        {
                            switch (sheet.GetCellTypeAtCoordinates(y, x))
                            {
                                case Sheet.CellTypes.Empty:
                                    {
                                        grid.Rows[y].Cells[x].Style.BackColor = sheet.ColorEmptyBackground;
                                        grid.Rows[y].Cells[x].Style.ForeColor = sheet.ColorEmptyForeground;

                                        break;
                                    }
                                case Sheet.CellTypes.XCategory:
                                    {
                                        grid.Rows[y].Cells[x].Style.BackColor = sheet.ColorCategoryItemBackground;
                                        grid.Rows[y].Cells[x].Style.ForeColor = sheet.ColorCategoryItemForeground;

                                        break;
                                    }
                                case Sheet.CellTypes.YCategory:
                                    {
                                        grid.Rows[y].Cells[x].Style.BackColor = sheet.ColorCategoryItemBackground;
                                        grid.Rows[y].Cells[x].Style.ForeColor = sheet.ColorCategoryItemForeground;

                                        break;
                                    }
                                case Sheet.CellTypes.Value:
                                    {
                                        grid.Rows[y].Cells[x].Style.BackColor = sheet.ColorValueBackground;
                                        grid.Rows[y].Cells[x].Style.ForeColor = sheet.ColorValueForeground;

                                        break;
                                    }
                                default:
                                    {
                                        grid.Rows[y].Cells[x].Style.BackColor = sheet.ColorEmptyBackground;
                                        grid.Rows[y].Cells[x].Style.ForeColor = sheet.ColorEmptyForeground;

                                        break;
                                    }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
        }
        
        /// <summary>
        /// Clears grid and builds a new one. Adds rows and columns, and formats them.
        /// </summary>
        /// <param name="grid"></param>
        /// <param name="sheet"></param>
        /// <param name="sheetCategoryType"></param>
        private static void BuildSheetCategory(DataGridView grid, Sheet sheet, Category.SheetCategoryType sheetCategoryType)
        {
            Int32 columnCount = -1;
            Int32 rowCount = -1;

            try
            {
                ConfigureVirtualDataGridView(grid);

                // Add columns and rows to the DataGridView.
                grid.Rows.Clear();
                grid.Columns.Clear();

                if (sheet != null)
                {
                    if (sheetCategoryType == Category.SheetCategoryType.Filter)
                    {
                        grid.ColumnHeadersVisible = true;

                        columnCount = sheet.CategoryFilters().Count;
                        rowCount = 1;

                        ConfigureVirtualDataGridViewTextCells(grid, columnCount, rowCount, (i => sheet.CategoryFilters()[i].Name));
                    }
                    else if (sheetCategoryType == Category.SheetCategoryType.X)
                    {
                        columnCount = 1;
                        rowCount = sheet.CategoryX().Count;

                        ConfigureVirtualDataGridViewTextCells(grid, columnCount, rowCount, (i => ""));
                    }
                    else if (sheetCategoryType == Category.SheetCategoryType.Y)
                    {
                        columnCount = sheet.CategoryY().Count;
                        rowCount = 1;

                        ConfigureVirtualDataGridViewTextCells(grid, columnCount, rowCount, (i => ""));
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
        }

        /// <summary>
        /// Formats grid.
        /// </summary>
        /// <param name="grid"></param>
        /// <param name="sheet"></param>
        /// <param name="sheetCategoryType"></param>
        private static void FormatSheetCategory(DataGridView grid, Sheet sheet, Category.SheetCategoryType sheetCategoryType)
        {
            try
            {
                if (sheet != null)
                {
                    //format cells
                    if (sheetCategoryType == Category.SheetCategoryType.Filter)
                    {
                        for (int x = 0; x < grid.Columns.Count; x++)
                        {
                            for (int y = 0; y < 1; y++)
                            {
                                grid.Rows[y].Cells[x].Style.BackColor = sheet.ColorCategoryItemBackground;
                                grid.Rows[y].Cells[x].Style.ForeColor = sheet.ColorCategoryItemForeground;
                            }
                        }
                    }
                    else if (sheetCategoryType == Category.SheetCategoryType.X)
                    {
                        for (int x = 0; x < 1; x++)
                        {
                            for (int y = 0; y < grid.Rows.Count; y++)
                            {
                                grid.Rows[y].Cells[x].Style.BackColor = sheet.ColorCategoryItemBackground;
                                grid.Rows[y].Cells[x].Style.ForeColor = sheet.ColorCategoryItemForeground;
                            }
                        }
                    }
                    else if (sheetCategoryType == Category.SheetCategoryType.Y)
                    {
                        for (int x = 0; x < grid.Columns.Count; x++)
                        {
                            for (int y = 0; y < 1; y++)
                            {
                                grid.Rows[y].Cells[x].Style.BackColor = sheet.ColorCategoryItemBackground;
                                grid.Rows[y].Cells[x].Style.ForeColor = sheet.ColorCategoryItemForeground;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
        }

        /// <summary>
        /// Configure a DataGridView for virtual Mode use.
        /// </summary>
        /// <param name="grid"></param>
        private static void ConfigureVirtualDataGridView(DataGridView grid)
        {
            try
            {
                _IsReshapingGrid = true;

                // Enable virtual mode.
                grid.VirtualMode = true;

                grid.ColumnHeadersVisible = false;
                grid.RowHeadersVisible = false;

                grid.AllowUserToAddRows = false;
                grid.AllowUserToDeleteRows = false;
                grid.AllowUserToOrderColumns = false;
                grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                grid.MultiSelect = false;

                // Reset columns and rows to the DataGridView.
                grid.Rows.Clear();
                grid.Columns.Clear();
                
                _IsReshapingGrid = false;
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                _IsReshapingGrid = false;
                throw;

            }
        }

        /// <summary>
        /// Configure a DataGridView with desired cells for virtual Mode use. 
        /// Use simple text cells.
        /// </summary>
        /// <param name="grid"></param>
        /// <param name="columnCount"></param>
        /// <param name="rowCount"></param>
        /// <param name="headerTextFunctor"></param>
        private static void ConfigureVirtualDataGridViewTextCells(DataGridView grid, Int32 columnCount, Int32 rowCount, Func<Int32, String> headerTextFunctor)
        {
            DataGridViewTextBoxColumn textBoxColumn = default(DataGridViewTextBoxColumn);

            try
            {
                // Add columns and rows to the DataGridView.
                //don't define columns if there are no rows
                if (rowCount > 0)
                {
                    _IsReshapingGrid = true;
                    for (int i = 0; i < columnCount; i++)
                    {
                        textBoxColumn = new DataGridViewTextBoxColumn();
                        //don't set this every time; only for filters!
                        textBoxColumn.HeaderText = headerTextFunctor(i); 
                        grid.Columns.Add(textBoxColumn);
                    }
                    if (columnCount > 0)
                    {
                        //only add rows if columns are defined
                        grid.Rows.Add(rowCount);
                        grid.RowCount = rowCount;
                    }
                    _IsReshapingGrid = false;
                    
                    //when done changing shape, resize columns to fit text (CellValueNeeded will be called).
                    grid.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                _IsReshapingGrid = false;
                throw;
            }
        }
        #endregion Private Methods
    }
}
