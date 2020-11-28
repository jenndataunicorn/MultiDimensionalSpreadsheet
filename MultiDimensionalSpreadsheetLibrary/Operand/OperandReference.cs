using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ssepan.Application;
using Ssepan.Collections;
using Ssepan.Utility;
using System.Diagnostics;
using System.Reflection;

namespace MultiDimensionalSpreadsheetLibrary
{
    public class OperandReference:
        OperandBase 
    {
        #region Declarations
        #endregion Declarations

        #region Constructors
        public OperandReference()
        { 
            try
            {
                Type = OperandType.CellReference;

            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
                throw;
            }
        }

        public OperandReference(List<CategoryItem> value)
        {
            try
            {
                Type = OperandType.CellReference;
                _Value = value;

            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
                throw;
            }
        }
        #endregion Constructors

        #region Properties
        private List<CategoryItem> _Value = default(List<CategoryItem>);
        #endregion Properties

        #region IOperand Members
        /// <summary>
        /// Evaluate the operand. Must process the address reference and find a single sheet cell.
        /// </summary>
        /// <param name="context"></param>
        /// <returns></returns>
        public override Single Evaluate(FormulaExecutionContext context)
        {
            Single returnValue = default(Single);
            List<CategoryItem> assignerCellCriteria = default(List<CategoryItem>);
            List<SheetCell> formulaSearchResults = default(List<SheetCell>);
            Sheet sheet = default(Sheet);

            try
            {
                //locate parent sheet, containing cells
                sheet = context.ExecutingFormula.Parent;
                if (sheet == null)
                {
                    throw new ApplicationException(String.Format("Unable to find parent sheet for formula: '{0}'", context.ExecutingFormula.Value));
                }

                //select all criteria in assignee cell EXCEPT those categories defined for assigner in the formula criteria.
                assignerCellCriteria = context.AssigneeCell.CategoryItems.Except(_Value, new EqualityComparerOfT<CategoryItem>((x, y) => x.Parent == y.Parent)).ToList<CategoryItem>();
                //add formula criteria to produce a complete set of criteria for an assigner cell.
                assignerCellCriteria.AddRange(_Value);


                //run assigner criteria
                formulaSearchResults = Sheet.GetCellsWhereAllSearchCategoryItemsMatchAnyInGivenCell(sheet.Cells.ToList<SheetCell>(), assignerCellCriteria);
                if ((formulaSearchResults == null) || (formulaSearchResults.Count == 0))
                {
                    throw new ApplicationException(String.Format("Unable to find assigner cell for parent sheet for formula '{0}'", context.ExecutingFormula.Value));
                }
                if (formulaSearchResults.Count > 1)
                {
                    throw new ApplicationException(String.Format("Too many results found for assigner cell for parent sheet for formula '{0}'", context.ExecutingFormula.Value));
                }

                //retrieve value from single cell in results
                returnValue = Single.Parse(formulaSearchResults[0].Value);

            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
            return returnValue;
        }
        #endregion IOperand Members
    }
}
