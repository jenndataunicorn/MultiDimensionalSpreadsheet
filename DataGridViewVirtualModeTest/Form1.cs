using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MultiDimensionalSpreadsheetLibrary;
using MultiDimensionalSpreadsheetLibrary.UI;
using Ssepan.Application;
using Ssepan.Application.WinConsole;
using Ssepan.Collections;
using Ssepan.Utility;
using System.Diagnostics;
using System.Reflection;

namespace DataGridViewVirtualModeTest
{
    public partial class Form1 :
        Form
    {
        private Boolean _ValueChangedProgrammatically;
        private String[] _CommandLineArgs = null;

        // Declare an entity to serve as the data store. 
        private Sheet sheet = new Sheet();

        public Form1(String[] args)
        {
            try
            {
                InitializeComponent();
                _CommandLineArgs = args;
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
            }
        }

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
                //ErrorMessage = ex.Message;
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
        }
        #endregion INotifyPropertyChanged

        #region PropertyChangedEventHandler
        /// <summary>
        /// Note: property changes update UI manually.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        internal void PropertyChangedEventHandler
        (
            Object sender,
            PropertyChangedEventArgs e
        )
        {
            try
            {
                if (e.PropertyName == "IsChanged")
                {
                    ApplySettings();
                }
            }
            catch (Exception ex)
            {
                //ErrorMessage = ex.Message;
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
            }
        }
        #endregion PropertyChangedEventHandler

        private void Form1_Load(object sender, EventArgs e)
        {

            //MDSSController<MDSSModel>.New(new Settings());
            //create Settings before first use by Model
            SettingsController<Settings>.New();
            MDSSController<MDSSModel>.New();
            Category createdCategory = default(Category);

            //subscribe to notifications
            this.PropertyChanged += PropertyChangedEventHandler;
            MDSSController<MDSSModel>.Model.PropertyChanged += PropertyChangedEventHandler;

            #region Add Sheets
            // Add sheet to settings
            SettingsController<Settings>.Settings.Sheets.Add(new Sheet("Sheet1"));//TODO:do to model, not settings
            #endregion Add Sheets

            #region Add Categories with CategoryItems
            //Create categories in settings
            SettingsController<Settings>.Settings.Sheets[0].Categories.Add
            (
                new Category
                (
                    "Year", 
                    Category.SheetCategoryType.X
                )
            );
            createdCategory = SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Year");
            createdCategory.Items =
                new OrderedEquatableBindingList<CategoryItem> 
		        {
			        new CategoryItem("2001", createdCategory),
			        new CategoryItem("2002", createdCategory)
		        };

            SettingsController<Settings>.Settings.Sheets[0].Categories.Add
            (
		        new Category
                (
                    "Quarter",
                    Category.SheetCategoryType.X
                )
            );
            createdCategory = SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Quarter");
            createdCategory.Items =
                    new OrderedEquatableBindingList<CategoryItem> 
		            {
			            new CategoryItem("Q1", createdCategory),
			            new CategoryItem("Q2", createdCategory),
			            new CategoryItem("Q3", createdCategory),
			            new CategoryItem("Q4", createdCategory)
		            };

            SettingsController<Settings>.Settings.Sheets[0].Categories.Add
            (
                new Category
                (
                    "Details", 
                    Category.SheetCategoryType.Y
                )
            );
            createdCategory = SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Details");
            createdCategory.Items =
                    new OrderedEquatableBindingList<CategoryItem> 
		            {
			            new CategoryItem("Units", createdCategory),
			            new CategoryItem("Price", createdCategory)
		            };

            SettingsController<Settings>.Settings.Sheets[0].Categories.Add
            (
                new Category
                (
                    "Department", 
                    Category.SheetCategoryType.Filter
                )
            );
            createdCategory = SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Department");
            createdCategory.Items =
                    new OrderedEquatableBindingList<CategoryItem> 
                    {
                        new CategoryItem("Belts", createdCategory),
                        new CategoryItem("Hats",  createdCategory)
                    };

            SettingsController<Settings>.Settings.Sheets[0].Categories.Add
            (
                new Category
                (
                    "Type",
                    Category.SheetCategoryType.None
                )
            );
            createdCategory = SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Type");
            createdCategory.Items =
                new OrderedEquatableBindingList<CategoryItem> 
                {
                    new CategoryItem("TypeA", createdCategory),
                    new CategoryItem("TypeB", createdCategory)
                };
            #endregion Add Categories with CategoryItems

            #region Add Categories to SheetCategories
            //Add sheet-categories that are reference to categories 
            //SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Year").CategoryType = Category.SheetCategoryType.X;
            //SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Quarter").CategoryType = Category.SheetCategoryType.X;
            //SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Department").CategoryType = Category.SheetCategoryType.X;
            //SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Details").CategoryType = Category.SheetCategoryType.Y;
            //SettingsController<Settings>.Settings.Sheets[0].CategoryX.Add
            #endregion Add Categories to SheetCategories

            #region Add SheetCells with Category Items
            //Add sheet cells with references to sheet-categories
            SettingsController<Settings>.Settings.Sheets[0].Cells.Add
            (
                new SheetCell
                (
                    "100",
                    new EquatableBindingList<CategoryItem> 
                    {
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Year").Items.Find(i => i.Name == "2001"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Quarter").Items.Find(i => i.Name == "Q1"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Department").Items.Find(i => i.Name == "Belts"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Details").Items.Find(i => i.Name == "Units")
                    }
                )
            );
            SettingsController<Settings>.Settings.Sheets[0].Cells.Add
            (
                new SheetCell
                (
                    "110",
                    new EquatableBindingList<CategoryItem> 
                    {
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Year").Items.Find(i => i.Name == "2001"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Quarter").Items.Find(i => i.Name == "Q2"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Department").Items.Find(i => i.Name == "Belts"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Details").Items.Find(i => i.Name == "Units")
                    }
                )
            );
            SettingsController<Settings>.Settings.Sheets[0].Cells.Add
            (
                new SheetCell
                (
                    "120",
                    new EquatableBindingList<CategoryItem> 
                    {
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Year").Items.Find(i => i.Name == "2001"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Quarter").Items.Find(i => i.Name == "Q3"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Department").Items.Find(i => i.Name == "Belts"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Details").Items.Find(i => i.Name == "Units")
                    }
                )
            );
            SettingsController<Settings>.Settings.Sheets[0].Cells.Add
            (
                new SheetCell
                (
                    "130",
                    new EquatableBindingList<CategoryItem> 
                    {
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Year").Items.Find(i => i.Name == "2001"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Quarter").Items.Find(i => i.Name == "Q4"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Department").Items.Find(i => i.Name == "Belts"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Details").Items.Find(i => i.Name == "Units")
                    }
                )
            );
            SettingsController<Settings>.Settings.Sheets[0].Cells.Add
            (
                new SheetCell
                (
                    "140",
                    new EquatableBindingList<CategoryItem> 
                    {
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Year").Items.Find(i => i.Name == "2002"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Quarter").Items.Find(i => i.Name == "Q1"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Department").Items.Find(i => i.Name == "Belts"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Details").Items.Find(i => i.Name == "Units")
                    }
                )
            );
            SettingsController<Settings>.Settings.Sheets[0].Cells.Add
            (
                new SheetCell
                (
                    "150",
                    new EquatableBindingList<CategoryItem> 
                    {
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Year").Items.Find(i => i.Name == "2002"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Quarter").Items.Find(i => i.Name == "Q2"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Department").Items.Find(i => i.Name == "Belts"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Details").Items.Find(i => i.Name == "Units")
                    }
                )
            );
            SettingsController<Settings>.Settings.Sheets[0].Cells.Add
            (
                new SheetCell
                (
                    "160",
                    new EquatableBindingList<CategoryItem> 
                    {
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Year").Items.Find(i => i.Name == "2002"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Quarter").Items.Find(i => i.Name == "Q3"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Department").Items.Find(i => i.Name == "Belts"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Details").Items.Find(i => i.Name == "Units")
                    }
                )
            );
            SettingsController<Settings>.Settings.Sheets[0].Cells.Add
            (
                new SheetCell
                (
                    "170",
                    new EquatableBindingList<CategoryItem> 
                    {
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Year").Items.Find(i => i.Name == "2002"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Quarter").Items.Find(i => i.Name == "Q4"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Department").Items.Find(i => i.Name == "Belts"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Details").Items.Find(i => i.Name == "Units")
                    }
                )
            );
            SettingsController<Settings>.Settings.Sheets[0].Cells.Add
            (
                new SheetCell
                (
                    "180",
                    new EquatableBindingList<CategoryItem> 
                    {
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Year").Items.Find(i => i.Name == "2001"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Quarter").Items.Find(i => i.Name == "Q1"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Department").Items.Find(i => i.Name == "Hats"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Details").Items.Find(i => i.Name == "Units")
                    }
                )
            );
            SettingsController<Settings>.Settings.Sheets[0].Cells.Add
            (
                new SheetCell
                (
                    "190",
                    new EquatableBindingList<CategoryItem> 
                    {
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Year").Items.Find(i => i.Name == "2001"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Quarter").Items.Find(i => i.Name == "Q2"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Department").Items.Find(i => i.Name == "Hats"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Details").Items.Find(i => i.Name == "Units")
                    }
                )
            );
            SettingsController<Settings>.Settings.Sheets[0].Cells.Add
            (
                new SheetCell
                (
                    "200",
                    new EquatableBindingList<CategoryItem> 
                    {
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Year").Items.Find(i => i.Name == "2001"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Quarter").Items.Find(i => i.Name == "Q3"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Department").Items.Find(i => i.Name == "Hats"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Details").Items.Find(i => i.Name == "Units")
                    }
                )
            );
            SettingsController<Settings>.Settings.Sheets[0].Cells.Add
            (
                new SheetCell
                (
                    "210",
                    new EquatableBindingList<CategoryItem> 
                    {
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Year").Items.Find(i => i.Name == "2001"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Quarter").Items.Find(i => i.Name == "Q4"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Department").Items.Find(i => i.Name == "Hats"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Details").Items.Find(i => i.Name == "Units")
                    }
                )
            );
            SettingsController<Settings>.Settings.Sheets[0].Cells.Add
            (
                new SheetCell
                (
                    "220",
                    new EquatableBindingList<CategoryItem> 
                    {
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Year").Items.Find(i => i.Name == "2002"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Quarter").Items.Find(i => i.Name == "Q1"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Department").Items.Find(i => i.Name == "Hats"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Details").Items.Find(i => i.Name == "Units")
                    }
                )
            );
            SettingsController<Settings>.Settings.Sheets[0].Cells.Add
            (
                new SheetCell
                (
                    "230",
                    new EquatableBindingList<CategoryItem> 
                    {
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Year").Items.Find(i => i.Name == "2002"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Quarter").Items.Find(i => i.Name == "Q2"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Department").Items.Find(i => i.Name == "Hats"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Details").Items.Find(i => i.Name == "Units")
                    }
                )
            );
            SettingsController<Settings>.Settings.Sheets[0].Cells.Add
            (
                new SheetCell
                (
                    "240",
                    new EquatableBindingList<CategoryItem> 
                    {
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Year").Items.Find(i => i.Name == "2002"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Quarter").Items.Find(i => i.Name == "Q3"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Department").Items.Find(i => i.Name == "Hats"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Details").Items.Find(i => i.Name == "Units")
                    }
                )
            );
            SettingsController<Settings>.Settings.Sheets[0].Cells.Add
            (
                new SheetCell
                (
                    "250",
                    new EquatableBindingList<CategoryItem> 
                    {
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Year").Items.Find(i => i.Name == "2002"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Quarter").Items.Find(i => i.Name == "Q4"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Department").Items.Find(i => i.Name == "Hats"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Details").Items.Find(i => i.Name == "Units")
                    }
                )
            );
            SettingsController<Settings>.Settings.Sheets[0].Cells.Add
            (
                new SheetCell
                (
                    "1.00",
                    new EquatableBindingList<CategoryItem> 
                    {
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Year").Items.Find(i => i.Name == "2001"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Quarter").Items.Find(i => i.Name == "Q1"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Department").Items.Find(i => i.Name == "Belts"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Details").Items.Find(i => i.Name == "Price")
                    }
                )
            );
            SettingsController<Settings>.Settings.Sheets[0].Cells.Add
            (
                new SheetCell
                (
                    "1.10",
                    new EquatableBindingList<CategoryItem> 
                    {
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Year").Items.Find(i => i.Name == "2001"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Quarter").Items.Find(i => i.Name == "Q2"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Department").Items.Find(i => i.Name == "Belts"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Details").Items.Find(i => i.Name == "Price")
                    }
                )
            );
            SettingsController<Settings>.Settings.Sheets[0].Cells.Add
            (
                new SheetCell
                (
                    "1.20",
                    new EquatableBindingList<CategoryItem> 
                    {
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Year").Items.Find(i => i.Name == "2001"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Quarter").Items.Find(i => i.Name == "Q3"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Department").Items.Find(i => i.Name == "Belts"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Details").Items.Find(i => i.Name == "Price")
                    }
                )
            );
            SettingsController<Settings>.Settings.Sheets[0].Cells.Add
            (
                new SheetCell
                (
                    "1.30",
                    new EquatableBindingList<CategoryItem> 
                    {
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Year").Items.Find(i => i.Name == "2001"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Quarter").Items.Find(i => i.Name == "Q4"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Department").Items.Find(i => i.Name == "Belts"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Details").Items.Find(i => i.Name == "Price")
                    }
                )
            );
            SettingsController<Settings>.Settings.Sheets[0].Cells.Add
            (
                new SheetCell
                (
                    "1.40",
                    new EquatableBindingList<CategoryItem> 
                    {
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Year").Items.Find(i => i.Name == "2002"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Quarter").Items.Find(i => i.Name == "Q1"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Department").Items.Find(i => i.Name == "Belts"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Details").Items.Find(i => i.Name == "Price")
                    }
                )
            );
            SettingsController<Settings>.Settings.Sheets[0].Cells.Add
            (
                new SheetCell
                (
                    "1.50",
                    new EquatableBindingList<CategoryItem> 
                    {
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Year").Items.Find(i => i.Name == "2002"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Quarter").Items.Find(i => i.Name == "Q2"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Department").Items.Find(i => i.Name == "Belts"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Details").Items.Find(i => i.Name == "Price")
                    }
                )
            );
            SettingsController<Settings>.Settings.Sheets[0].Cells.Add
            (
                new SheetCell
                (
                    "1.60",
                    new EquatableBindingList<CategoryItem> 
                    {
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Year").Items.Find(i => i.Name == "2002"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Quarter").Items.Find(i => i.Name == "Q3"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Department").Items.Find(i => i.Name == "Belts"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Details").Items.Find(i => i.Name == "Price")
                    }
                )
            );
            SettingsController<Settings>.Settings.Sheets[0].Cells.Add
            (
                new SheetCell
                (
                    "1.70",
                    new EquatableBindingList<CategoryItem> 
                    {
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Year").Items.Find(i => i.Name == "2002"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Quarter").Items.Find(i => i.Name == "Q4"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Department").Items.Find(i => i.Name == "Belts"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Details").Items.Find(i => i.Name == "Price")
                    }
                )
            );
            SettingsController<Settings>.Settings.Sheets[0].Cells.Add
            (
                new SheetCell
                (
                    "1.80",
                    new EquatableBindingList<CategoryItem> 
                    {
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Year").Items.Find(i => i.Name == "2001"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Quarter").Items.Find(i => i.Name == "Q1"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Department").Items.Find(i => i.Name == "Hats"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Details").Items.Find(i => i.Name == "Price")
                    }
                )
            );
            SettingsController<Settings>.Settings.Sheets[0].Cells.Add
            (
                new SheetCell
                (
                    "1.90",
                    new EquatableBindingList<CategoryItem> 
                    {
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Year").Items.Find(i => i.Name == "2001"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Quarter").Items.Find(i => i.Name == "Q2"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Department").Items.Find(i => i.Name == "Hats"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Details").Items.Find(i => i.Name == "Price")
                    }
                )
            );
            SettingsController<Settings>.Settings.Sheets[0].Cells.Add
            (
                new SheetCell
                (
                    "2.00",
                    new EquatableBindingList<CategoryItem> 
                    {
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Year").Items.Find(i => i.Name == "2001"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Quarter").Items.Find(i => i.Name == "Q3"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Department").Items.Find(i => i.Name == "Hats"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Details").Items.Find(i => i.Name == "Price")
                    }
                )
            );
            SettingsController<Settings>.Settings.Sheets[0].Cells.Add
            (
                new SheetCell
                (
                    "2.10",
                    new EquatableBindingList<CategoryItem> 
                    {
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Year").Items.Find(i => i.Name == "2001"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Quarter").Items.Find(i => i.Name == "Q4"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Department").Items.Find(i => i.Name == "Hats"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Details").Items.Find(i => i.Name == "Price")
                    }
                )
            );
            SettingsController<Settings>.Settings.Sheets[0].Cells.Add
            (
                new SheetCell
                (
                    "2.20",
                    new EquatableBindingList<CategoryItem> 
                    {
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Year").Items.Find(i => i.Name == "2002"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Quarter").Items.Find(i => i.Name == "Q1"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Department").Items.Find(i => i.Name == "Hats"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Details").Items.Find(i => i.Name == "Price")
                    }
                )
            );
            SettingsController<Settings>.Settings.Sheets[0].Cells.Add
            (
                new SheetCell
                (
                    "2.30",
                    new EquatableBindingList<CategoryItem> 
                    {
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Year").Items.Find(i => i.Name == "2002"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Quarter").Items.Find(i => i.Name == "Q2"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Department").Items.Find(i => i.Name == "Hats"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Details").Items.Find(i => i.Name == "Price")
                    }
                )
            );
            SettingsController<Settings>.Settings.Sheets[0].Cells.Add
            (
                new SheetCell
                (
                    "2.40",
                    new EquatableBindingList<CategoryItem> 
                    {
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Year").Items.Find(i => i.Name == "2002"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Quarter").Items.Find(i => i.Name == "Q3"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Department").Items.Find(i => i.Name == "Hats"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Details").Items.Find(i => i.Name == "Price")
                    }
                )
            );
            SettingsController<Settings>.Settings.Sheets[0].Cells.Add
            (
                new SheetCell
                (
                    "2.50",
                    new EquatableBindingList<CategoryItem> 
                    {
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Year").Items.Find(i => i.Name == "2002"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Quarter").Items.Find(i => i.Name == "Q4"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Department").Items.Find(i => i.Name == "Hats"),
                        SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Details").Items.Find(i => i.Name == "Price")
                    }
                )
            );
            #endregion Add SheetCells with Category Items
            
            SettingsController<Settings>.Settings.Sync();

            // Connect the virtual-mode events to event handlers. 
            this.dataGridView1.CellValueNeeded += new DataGridViewCellValueEventHandler(dataGridView1_CellValueNeeded);
            this.dataGridView1.CellValuePushed += new DataGridViewCellValueEventHandler(dataGridView1_CellValuePushed);
            this.dgvSheetCategoriesX.CellValueNeeded += new DataGridViewCellValueEventHandler(dgvSheetCategoriesX_CellValueNeeded);
            this.dgvSheetCategoriesY.CellValueNeeded += new DataGridViewCellValueEventHandler(dgvSheetCategoriesY_CellValueNeeded);


            // Add columns and rows to the DataGridView.
            //format cells
            DataGridViewHelper.RenderSheet(this.dataGridView1, SettingsController<Settings>.Settings.Sheets[0]);
            //DataGridViewHelper.RenderSheet(this.dataGridView1, (Sheet)null);

            //SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Details").CategoryType = Category.SheetCategoryType.Y;
            this.dataGridView2.DataSource = SettingsController<Settings>.Settings.Sheets[0].Categories;

            DataGridViewHelper.RenderSheetCategory(this.dgvSheetCategoriesFilter, SettingsController<Settings>.Settings.Sheets[0], Category.SheetCategoryType.Filter);
            DataGridViewHelper.RenderSheetCategory(this.dgvSheetCategoriesX, SettingsController<Settings>.Settings.Sheets[0], Category.SheetCategoryType.X);
            DataGridViewHelper.RenderSheetCategory(this.dgvSheetCategoriesY, SettingsController<Settings>.Settings.Sheets[0], Category.SheetCategoryType.Y);

            dgCells.DataSource = SettingsController<Settings>.Settings.Sheets[0].Cells;



#if debug
                //debug view
             menuEditProperties_Click(sender, e);
#endif
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.PropertyChanged -= PropertyChangedEventHandler;
            MDSSController<MDSSModel>.Model.PropertyChanged -= PropertyChangedEventHandler;
        }

        private void menuEditProperties_Click(object sender, EventArgs e)
        {
            String statusMessage = "";
            String errorMessage = "";

            try
            {
                //MDSSController<MDSSModel>.ShowProperties(() => { });
                throw new NotImplementedException("menuEditProperties_Click -- implement viewmodel");
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message.ToString();
                statusMessage = "";

                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

            }
            finally
            {
                //StopProgressBar(statusMessage, errorMessage);
            }
        }

        private void dataGridView1_CellValueNeeded(object sender, DataGridViewCellValueEventArgs e)
        {
            DataGridViewHelper.SheetCellValueNeeded(sender, e, SettingsController<Settings>.Settings.Sheets[0]);
        }

        private void dgvSheetCategoriesFilter_CellValueNeeded(object sender, DataGridViewCellValueEventArgs e)
        {
            DataGridViewHelper.SheetCategoryCellValueNeeded(sender, e, SettingsController<Settings>.Settings.Sheets[0], Category.SheetCategoryType.Filter);
        }

        private void dgvSheetCategoriesX_CellValueNeeded(object sender, DataGridViewCellValueEventArgs e)
        {
            DataGridViewHelper.SheetCategoryCellValueNeeded(sender, e, SettingsController<Settings>.Settings.Sheets[0], Category.SheetCategoryType.X);
        }

        private void dgvSheetCategoriesY_CellValueNeeded(object sender, DataGridViewCellValueEventArgs e)
        {
            DataGridViewHelper.SheetCategoryCellValueNeeded(sender, e, SettingsController<Settings>.Settings.Sheets[0], Category.SheetCategoryType.Y);
        }

        private void cmdRender_Click(object sender, EventArgs e)
        {
            DataGridViewHelper.RenderSheet(this.dataGridView1, SettingsController<Settings>.Settings.Sheets[0]);

            DataGridViewHelper.RenderSheetCategory(this.dgvSheetCategoriesFilter, SettingsController<Settings>.Settings.Sheets[0], Category.SheetCategoryType.Filter);
            DataGridViewHelper.RenderSheetCategory(this.dgvSheetCategoriesX, SettingsController<Settings>.Settings.Sheets[0], Category.SheetCategoryType.X);
            DataGridViewHelper.RenderSheetCategory(this.dgvSheetCategoriesY, SettingsController<Settings>.Settings.Sheets[0], Category.SheetCategoryType.Y);

            dgCells.DataSource = SettingsController<Settings>.Settings.Sheets[0].Cells;
        
        }

        /// <summary>
        /// add createdCategory
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cmdTest1_Click(object sender, EventArgs e)
        {
            CategoryItem categoryItem = default(CategoryItem);

            //User must select a Category Item.
            categoryItem = MDSSController<MDSSModel>.SelectCategoryItem(SettingsController<Settings>.Settings.Sheets[0], "Type", "Select a Category Item to Add to existing Sheet Cells.", 0, false);
            if (categoryItem != null)
            {
                //_ValueChanging = true;

                //move it
                SettingsController<Settings>.Settings.Sheets[0].AssignCategory("Type", Category.SheetCategoryType.Y, categoryItem);

                //_ValueChanging = false;

                //refresh
                //MDSSController<MDSSModel>.Refresh();
                dgCells.DataSource = SettingsController<Settings>.Settings.Sheets[0].Cells;

                //returnValue = true;
            }

        }

        /// <summary>
        /// remove createdCategory
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cmdTest2_Click(object sender, EventArgs e)
        {
            CategoryItem categoryItem = default(CategoryItem);

 
            //User must select a Category Item.
            categoryItem = MDSSController<MDSSModel>.SelectCategoryItem(SettingsController<Settings>.Settings.Sheets[0], "Type", "Select a Category Item to Add to existing Sheet Cells.", 0, false);
            if (categoryItem != null)
            {
                //_ValueChanging = true;

                //move it
                SettingsController<Settings>.Settings.Sheets[0].AssignCategory("Type", Category.SheetCategoryType.None, categoryItem);

                //_ValueChanging = false;

                //refresh
                //MDSSController<MDSSModel>.Refresh();
                dgCells.DataSource = SettingsController<Settings>.Settings.Sheets[0].Cells;

                //returnValue = true;
            }

            //List<CategoryItem> sheetCategoryItems = new List<CategoryItem> { SettingsController<Settings>.Settings.Sheets[0].Categories.Find(c => c.Name == "Department").Items.Find(ci => ci.Name == "Belts") };
            //List<SheetCell> cellsWithSelectedCategoryItem = (from cell in SettingsController<Settings>.Settings.Sheets[0].Cells
            //                                                 where sheetCategoryItems.All(ci1 => cell.CategoryItems.Any(ci2 => ci2 == ci1))
            //                                                 select cell).ToList<SheetCell>();
            //List<SheetCell> cellsWithUnSelectedCategoryItem = (from cell in SettingsController<Settings>.Settings.Sheets[0].Cells
            //                                                   where sheetCategoryItems.All(ci1 => cell.CategoryItems.Any(ci2 => ((ci2.Parent == ci1.Parent) && (ci2.Name != ci1.Name))))
            //                                                   select cell).ToList<SheetCell>();
        }

        private void dataGridView1_CellValuePushed(object sender, DataGridViewCellValueEventArgs e)
        {
            DataGridViewHelper.SheetCellValuePushed(sender, e, SettingsController<Settings>.Settings.Sheets[0]);
        }

        private void cmdPromote_Click(object sender, EventArgs e)
        {
            String categoryName = default(String);

            if (this.dataGridView2.SelectedRows.Count == 1)
            {
                categoryName = this.dataGridView2.SelectedRows[0].Cells[0].Value.ToString();

                MDSSController<MDSSModel>.ShiftSheetCategory(SettingsController<Settings>.Settings.Sheets[0], categoryName, ListOfTExtension.ShiftTypes.Promote);
            }
        }

        private void cmdDemote_Click(object sender, EventArgs e)
        {
            String categoryName = default(String);

            if (this.dataGridView2.SelectedRows.Count == 1)
            {
                categoryName = this.dataGridView2.SelectedRows[0].Cells[0].Value.ToString();

                MDSSController<MDSSModel>.ShiftSheetCategory(SettingsController<Settings>.Settings.Sheets[0], categoryName, ListOfTExtension.ShiftTypes.Demote);
            }
        }

        /// <summary>
        /// Bind Settings controls to SettingsController
        /// </summary>
        private void BindSettingsUi()
        {
            try
            {
                //Settings
                //bindingSource1.DataSource = SettingsController<Settings>.Settings.Sheets;
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
                BindSettingsUi();

                ////apply settings that shouldn't use databindings
                //this.Size = SettingsController<Settings>.Settings.Size;
                //this.Location = SettingsController<Settings>.Settings.Location;

                ////apply settings that can't use databindings
                //this.Text = Path.GetFileName(SettingsController.Filename) + " - " + _OriginalFormTitle;

                ////apply settings that don't have databindings
                //this.StatusBarDirtyMessage.Visible = (SettingsController.Dirty);

                _ValueChangedProgrammatically = false;
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

            }
        }

    }
}
