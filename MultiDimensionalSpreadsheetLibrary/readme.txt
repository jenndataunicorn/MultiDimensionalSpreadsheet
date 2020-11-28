MultiDimensionalSpreadsheet v0.14 (Alpha)


Purpose:
To let the user create spreadsheets of more than two dimensions.
To allow formulas to use named cell values, using category-item or category names.
To allow the user to reconfigure the views of the data on the fly, and save the configuration.
To produce an open storage format in XML for sheets.
To produce a long-overdue open-source alternative to the proprietary applications that pioneered this concept in the past or embody the notion today.


Usage notes:
~Define one or more sheets in the Sheets pane.
~For each sheet, define two or more categories in the Categories pane.
~For each category, define one or more category items in the Category Items pane
~For each sheet, define zero or more formulas in the Formulas pane.
~Right click a category in the Categories pane to add it to a sheet category in the Sheet pane.
~Select a sheet in the Sheets pane to view it in the Sheet pane.
~Add / enter values in the cells in the Sheet pane.
~Define formulas in the Formulas pane that define calculations for ranges of cells by Category-Item name.  Formulas support assignment (=) to a cell address and simple arithmetic operations (+-*/) on integer and floating-point literals, and cell addresses. Cell addresses are designated with un-quoted strings containing one or more Category-Item names, separated by ':' like "Item1:Item2". Order is NOT significant; "Item1:Item2" is equivalent to "Item2:Item1". Category-Items may optionally explicitly indicate their Category(*) using a '.' separator like "CategoryName.ItemName". Order IS significant; "CategoryName.ItemName" is NOT equivalent to "ItemName.CategoryName". Use parenthesis to override the precedence of arithmetic operators. 
~*: Explicit Category names are allowed so that the use of the same Category-Item name in two different Categories is not a problem.
~Right click a sheet category in the Sheet pane to move to another sheet category in the Sheet pane or remove it from a sheet category in the Sheet pane.
~Functions can be specified in formulas, in the format @FUNCTION{parameter1, parameter2, ...}. See 'Functions Supported' for implemented formulas.

Design Notes:
Implementation Choices
~Uses MVC.
~Uses DataGridView (in call-back mode) for sheet and sheet categories (x- and y-categories, and filter);.
~Uses context menus to reconfigure categories.
~Uses xml serialization (specifically, DataContractSerialization -- for the preservation of object references) for settings / data files.

Understanding The Application User Interface
	The interface is laid out in a set of panes. The Sheets pane controls all of the others, which display values for the currently selected sheet. All of the others but one (the Sheet pane) define the  components of a sheet (except for cells). The Sheet pane is where an individual sheet is displayed and configured, as well as where cell data is edited.
	Multiple sheets can be defined in the Sheets pane. For each sheet, at least two categories must be defined in the Category pane. For each category, at least one item must be defined in the Category Item pane. For each sheet, zero or more formulas may be defined in the Formulas pane. Once these components are defined, categories can be assigned to specific regions in the Sheet pane.
	The Sheet pane is laid out with four regions. The large one in the center is where the cells will be displayed. The other three are where the categories will be placed.
	Categories can be placed in one of three locations on a sheet: X, Y, and Filter categories. X categories are categories placed to the right of the sheet, and they control the vertical placement of category items along the top edge of the sheet. Y categories are categories placed under the bottom of the sheet, and they control the horizontal placement of category items along the left edge of the sheet. Filter categories are categories placed above the top of the sheet, and they control which of those category's items are displayed in the sheet. Once at least one X and one Y category are assigned, the sheet can be rendered.
	Categories can be added or removed from the three regions, moved between regions, or promoted or demoted within a region. Filter categories can also have an item selected.
	Note that a category can only be assigned to one region at a time. Also, if a category is not assigned, it will not have any cells associated with it, because all cells, whether visible or hidden by a filter, are associated with all of the categories in the three regions X, Y, and Filter (More specifically, they are associated with one  of each categories' items.)
	The number of dimensions in a sheet is defined by the categories which are assigned a sheet category type other than 'None'. To associate a category to a sheet category type is to apply a new dimension to a sheet, a dimension with a range equal to the category's items. To dis-associate a category from a sheet category type is to remove that dimension from the sheet.
	If an unassigned category is assigned to a region in a sheet that is already rendered, then that category must also be associated with the existing cells. This is done in two steps. First, the user must be prompted to select one of the category's items. Second, that item must be added to the criteria of every cell. If the category contains other items besides the one that the user selects, those will generate new cells and they will be empty initially. Whether they are also rendered depends upon the filter.
	If a category is removed altogether (not moved) from a region in a sheet that is already rendered, then that category must be dis-associated from the existing cells. This is done in three steps.  First, the user must be prompted to select one of the category's items. Second, that item must be removed from the criteria of every cell that contains it. Third, if the category contains other items besides the one that the user selects, the cells that contain those items in their criteria will be deleted.
	Note that deleting categories or items while the category is assigned is not recommended, as the application does not currently safeguard against the errors this will cause. For now, always dis-associate categories from sheet category types first, before deleting them or their items. Future releases will address this shortcoming.

Identifying A Cell By Row And Column
	In MultiDimensionalSpreadsheet, the flexible nature of a sheet means that traditional assumptions about referencing cells need to be abandoned. Cells have no permanent association with, and no knowledge of, rows and columns. 
	Instead, a different construct is needed to identify and arrange cells. This construct is the Category and it's Items. A Category is a group of headings (Items) that can be managed and organized together. The Category is the unit of information that is used to reorganize a sheet's cells. Items are the unit of information used to identify a sheet's cells.
	Individual cells are associated with one or more Category Items (although they need two or more -- one in the X category and one in the Y -- in order to be displayed.) Cells are positioned according to the placement of their associated category items; category items are placed according to their parent categories. So to describe cell placement it is necessary to start with an understanding of category placement.
	Regions were described earlier, in 'Understanding The Application User Interface'.
	The order of Filter categories is not significant. The order of X and Y categories controls how the category items will be arranged. Categories specified first (top-most for X and left-most for Y ) will have their items rendered once. The next category in the list will have it's items rendered once for each item in the category 'above' it. This means the X dimension must account for as many columns as there are in the product of all the items in the X categories plus the sum of all the categories in the Y category. Likewise,  the Y dimension must account for as many rows as there are in the product of all the items in the Y categories plus the sum of all the categories in the X category.
	This implies four things about the layout of the grid. First, the upper-left corner will contain empty cells. This region will have as many columns as there are Y categories and as many rows as there are X categories Second, the X category items will occupy the upper-right cells in the grid. This region will have as many columns as there are in the product of X category items and as many rows as there are X categories. Third, the Y category items will occupy the lower-left cells in the grid. This region will have as many columns as there are Y categories and as many rows as there are in the product of Y category items.  Fourth, the lower-right corner will contain value cells.  This region will have as many columns as the product of X category items and as many rows as the product of Y category items.
	As mentioned earlier, a given category's items are repeated for each superior category's item. This is accomplished by calculating the product of all superior category's items and determining how many times to repeat the items.
	Items are not repeated for each inferior category's items. This is accomplished by calculating the product of all inferior category's items and using a modulus to determine if the cell is the first instance in that group.
	For a given X or Y category, items are placed in the same row or column as their category. The item's other coordinate is determined by it's order in the category and any repetition of the given category within a superior category's items.
	When the DataGridView control needs to retrieve or set a cell value, it fires events. In the event handlers we must translate the row and column into the criteria that describe a cell. The cell at the intersection of a row and column is defined by three things. First, it must be associated with all of the category items in that column. Second, it must be associated with all of the category items in that row.  Third, it must be associated with the category items selected in each category in the filter.
	Finding a cell by row and column is performed with the following two steps. First, for each of the three category types (X, Y, and Filter) determine which category items in the sheet apply to the row and column. In the case of the Filter category it is every item selected in each category. In the case of the X and Y categories it is every item in that column or row, respectively. All of these category items together are the Criteria, and they are gathered into a single list. Second, the criteria list is passed to a search function where all items in the criteria must match any item in a cell's list of category items. If a single cell is found then it is an exact match.

Identifying A Cell Or Range Of Cells By Address
	Like cells, formulas also need to be more flexible than their traditional counterparts. Formulas in MultiDimensionalSpreadsheet differ from those in traditional spreadsheets in two ways. First, formulas are not attached to cells, but are defined independently. Second, as a consequence of not implicitly knowing the identity of the assignee, the formula must explicitly define one on the left side of the equals sign. 
	The address scheme allows for a varying degree of specificity; it can be made to identify a range of cells or a single cell. An address consists of one or more category item names, separated by a colon. Each item may optionally be qualified by it's category name, separated by a period. When the application needs to evaluate a cell address, it does so in a manner similar to the row and column lookup. It compiles the category items into a criteria list, and then uses the list to search for cells. In this case, however; it is acceptable if more than one cell is matched. The manner in which multiple matches are narrowed down to a single cell is described next.
	When a cell value lookup occurs, before the value is retrieved from the cell data structure, the formulas are searched for assignee address references to the cell. If the cell's criteria produce a match when run against the results of an assignee address criteria search, then the formula is one that applies to the cell. In that case, the  formula is evaluated, the cell data structure is updated with the value, and it is that value which is returned. That is how assignee references are narrowed to a single cell; assigner references on the right ride of the equals sign are discussed next.
	Unlike the assignee portion of the formula on the left of the equals sign, the addresses that appear in the right side are not simply a subset of the criteria of a cell that is known when the formula is evaluated. That means that there is no readily available set of criteria with which to narrow down assigner address criteria results, as was done for the assignee.  However, a set of criteria can still be constructed by starting with the criteria list from the assignee cell. This is because the intended assignee will likely share many categories in common with the assigner, and any missing from the latter should be inferred. So, two changes are made. First, the application removes any items in the list where there are items in the assigner address with  the same category  but different items. Second, it adds items to the list from the  assigner address.  If this list  produces a single match when run against all of the sheet's cells, then that address operand will return the cell's value.
	At this point it must be noted that more than one formula may apply to an assignee cell. Formulas will be processed in the order that they are defined. The last formula to operate on a cell will be the one whose value is ultimately displayed. Formulas can be reordered using the context menu. When writing formulas that differ only in how specific they are (that is, how many items are specified in any addresses), it is more effective to place more specific versions of a formula after more general versions.
     



Enhancements:
0.Next: (PLANNED)
~TODO:researching change of use of OperationBase objects to System.Linq.Expressions...
~TODO:rename OperandReference to OperandCellCriteria, and create an OperandReference for use with scalar variables.
~TODO:make operands of T, where T is numeric; continue to use operand of Single for now

0.14: (DEVELOPMENT)
~Fixed bug when renaming sheet or category. Changed ParentName to Parent object reference in CategoryItem and Formula, and changed Ssepan.Utility ObjectHelper to use DataContract serializer as an option. (Background: Renaming a sheet or category hangs program in call to GetCellsWhereAllSearchCategoryItemsMatchAnyInGivenCell() in IsApplicableFormula() in Formula.cs. The name in 'this.ParentName' still has the previous value. Child controls had name instead of parent reference.) 
~Also had to modify Ssepan.Utility.ObjectHelper.Clone() to optionally use DataContract, including circuular-reference awareness.
~Also had to fix circular reference when serializing DataContract. Added  'IsReference=true' to DataContract attribute on objects being serialized under Settings. 
~Also exposed a problem deserializing DataContract, in which properties not included in serialization (including but not limited to SerialiseAs) were not reconstituted because Settigns ctor did not run. Converted these properties to static. (Note: This will have the side efect of limiting the SettingsController to one configuration at a time if we ever allow muktiple settgins files per controller.)
~Handler in Program.cs was not wired up in forms or console; added 'PropertyChanged += PropertyChangedEventHandlerDelegate;' to Main.
~No handler wired up to Settings in forms or console. Renamed 'PropertyChangedEventHandlerDelegate' to 'ModelPropertyChangedEventHandlerDelegate'. Wired up latter in 'InitViewModel' after model handler in view. Currently handling 'Dirty' property.
~Fixed bug in GetPathForSave in Ssepan.Io where SaveAs did not display dialog for '(new)'.
~Modified Settings in Ssepan.Application to implement part of it interfaces as a new interface ISettingsComponent, and implemented an example as a property SomeComponent which copies the PropertyChanged handlers from settings and implements its own Equals, Dirty, CopyTo, Sync, etc, so Settings does not need to know the details.


0.13: (RELEASED)
~Refactored to use new MVVM / MVC hybrid.
~Updated Ssepan.* to 2.6
~Fixed DataGridViewVirtualModeTest to work with changes to Settings / SettingsController from 0.12.
~Modified Ssepan.Collections to provide an extension method ShiftListItem.
~Changed MultimensionalSpreadsheetLibrary to use ShiftListItem method.
~Fixed progress bar and status message logic.
~Using version 2.5 of Ssepan.* libraries.
~Minor changes to documentation paragraphs above.
~Renaming local variables and parameters, and correcting XML comments.
~Encapsulate and reviewed sheet-redimensioning logic.
~Further documented sheet-redimensioning logic.
~Fixed problem with Model not being instantiated before View form move tries to access it and trigger a Refresh.

0.12: 
~Refactored Settings / SettingsController, and their bases in Ssepan.Application, to put the static Settings property into the Settings class instead of the SettingsController class. This will make Settings more like SettingsController and the model / controller classes, and hopefully make Settings easier to understand and maintain.
~Projects are using .Net Framework 4.0.
~Using version 2.0 of Ssepan.* libraries, all of which are using .Net Framework 4.0.

0.11:
~Refactor image resources.

0.10: 
~Refactor UI File IO logic in menu events to eliminate duplication.
~Added images to promote / demote context menu items.
~Modified Sepan.Application to include _ValueChanging flag from sub-class, and to check and set it from the Controller base class Refresh method.
~Fixed errors thrown in RefreshSheet and RefreshSheetCategories that occurred when an empty grid Sheets had a row in add-mode but the underlying Sheets collection was empty.

0.9: 
~Moved common portions of settings I/O into base classes in Ssepan.Application library.

0.8: 
~Created EquatableOrderedBindingList(Of T) which considers the order of contents as also significant, unlike EquatableBindingList(Of T) which only compares the contents. 
~Also created EquatableOrderedList(Of T). 
~Also created addditional extensions for conversion.
~Sheets and Cells continue to use the unordered variety of equatable list; Categories, Category Items, and Formulas use the ordered variety of equatable list. 

0.7: 
~Documented portions of the design pertaining to the identification and arrangement of cells.
~Added labels to the sheet categories

0.6:
~Implemented functions in formulas.
~Moved Formula currentFormula, List<CategoryItem> assigneeCellCriteria out of operand / operator / operation and into new object FormulaExecutionContext.
~Replaced Left, Right parameters in operator / operation with Dictionary<String, OperandBase>.
~Reorganized library's source files into folders.
~Implemented a few functions from System.Math for use in formulas.

0.5:
~Implemented parentheses parsing in formulas.
~Fixed bug in operand parsing. (See 'Fixes'.)

0.4:
~Implemented formula re-ordering.
~Implemented category-item re-ordering.
~Fixed formula tooltip. (See 'Fixes'.)

0.3:
~Implemented simple formulas. (See 'Usage notes'.)
~There is a sample file (TEST.MDSS) in the settings folder.
~The project can be compiled with the conditional 'debug', which will cause the Properties dialog to load with the main form. This can be handy for inspecting the state of the settings, including Sheets, Cells, and Categories/Items.

0.2:
~Implemented filter category.
~Changed category DataGridViews to read-only cells.
~Included sample settings file in source.

0.1:
~Attempting to produce a Proof-of-Concept of the UI, the data model, and the storage format.
~Unlike previous projects, the serialization used in MDSS is DataContractSerialization instead of XmlSerialization, because the former, while still using XML, also preserves object references.
~The Filter Category feature is not enabled.
~The Formulas list can be maintained, but formulas are not implemented in this version.


Functions Supported:
~@ABS: Absolute value. Takes a single numeric parameter. (See 'Known Issues' regarding expressions as parameters, function names.)
~@MAX: Maximum of two numbers. Takes two numeric parameters.
~@MIN: Minumum of two numbers. Takes two numeric parameters.

Fixes:
~0.10: Fixed errors thrown in RefreshSheet and RefreshSheetCategories that occurred when an empty grid Sheets had a row in add-mode but the underlying Sheets collection was empty.
~0.8: Fixed bug in list equality testing that would fail on the same contents in a different order in two lists. EquatableList<T> and EquatableBindingList<T> Equal() implementation returns true if two lists contain the same items (regardless of order). So the promote/demote of some lists (i.e. Formulas) do not trigger Dirty flag. Need to create EquatableOrderedBindingList(Of T).
~0.5: Fixed bug in operand parsing that would fail on floating point literal in formula.
~0.4: Fixed formula tooltip that still indicated that formulas were not implemented.


Known Issues:
~Filename passed in command line argument is converted/passed in DOS 8.3 equivalent format. Cannot compare file extension directly. Status: research. 
~Running this app under Vista or Windows 7 requires that the library that writes to the event log (Ssepan.Utility.dll) have its name added to the list of allowed 'sources'. Rather than do it manually, the one way to be sure to get it right is to simply run the application the first time As Administrator, and the settings will be added. After that you may run it normally. To register additional DLLs for the event log, you can use this trick any time you get an error indicating that you cannot write to it. Or you can manually register DLLs by adding a key called '<filename>.dll' under HKLM\System\CurrentControlSet\services\eventlog\Application\, and adding the string value 'EventMessageFile' with the value like <C>:\<Windows>\Microsoft.NET\Framework\v2.0.50727\EventLogMessages.dll (where the drive letter and Windows folder match your system). Status: work-around. 
~The focus goes to the Sheets list by default, and when the DataGridView receives the focus in an empty grid, it goes into Insert mode, adding a temporary row. This will result in the model having a temmporary 'Dirty' state, which you will notice if you immediately open a settings file. Work-around: You can ignore it and choose No, or you can tab out to another control. I will look into putting the focus elsewhere.  Status: research.
~Need to trap category delete and handle sheet cells like category remove and prompt user for slice to remain. Status: research determined that BindingNavigator delete code must be replaced with custom delete code; this is postponed for a future release. Work-around: remove category from X, Y, or Filter before deleting.
~Need to trap category-item delete and possibly delete cells containing it (prompt user first).  Status: research determined that BindingNavigator delete code must be replaced with custom delete code; this is postponed for a future release. Work-around: remove category from X, Y, or Filter before deleting.
~Cells that have no value stored, but have a formula that apply to them, are not being drawn automatically until the file is saved. This is because the CellValueNeeded event is not firing for those cells. I cannot fire Refresh method here because it will cause recursive calls. I had to wrap a flag around the cell assignment to prevent such an occurance when setting the cell value in the formula code. Note: it is the formula-dependent cells that are not being requested until after the first time a value is applied.
~Function parameters in formulas do not support expressions unless wrapped in parenthesis.
~Formulas do not support literal values other than integer and floating-point.
~Formulas do not support unary operators, such as negative ('-').
~Formulas do not detect circular references.
~Settings files from previous versions(**) may not load, because DataContractSerialization is less forgiving about de-serializing settings-files with structural differences than XMLSerialization.
~**:  0.1 -> 0.2,  0.2 -> 0.3
~Some errors are being trapped and logged silently, and not bubbling up to the display.
~Moving category from X or Y to Filter and back leave condition that shows dirty flag.

Possible Enhancements:
~Implement a larger set of System.Math functions for use in formulas.
~Now that I've seen how to accomplish some of the features and what sort of checks I need to do when adding / removing categories and category-items in an existing sheet, I should re-design the model to account for those changes instead of handling the validation in the UI.
~(When DataContractSerialier fixes inability to serialize circular-references) change String ParentName properties of Formula and CategoryItem to Sheet Parent and Category Parent respectively.

Steve Sepan
ssepanus@yahoo.com
5/23/2014