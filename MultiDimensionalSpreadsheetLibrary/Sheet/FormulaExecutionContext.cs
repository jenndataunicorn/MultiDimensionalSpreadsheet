using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ssepan.Utility;
using System.Diagnostics;
using System.Reflection;

namespace MultiDimensionalSpreadsheetLibrary
{
    /// <summary>
    /// Information about a given formula execution; needed by child objects (operations, operators, operands)
    /// </summary>
    public class FormulaExecutionContext
    {
        public FormulaExecutionContext()
        {

        }
        public FormulaExecutionContext(Formula formula, SheetCell cell)
        {
            try
            {
                ExecutingFormula = formula;
                AssigneeCell = cell;
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
                throw;
            }
        }

        private Formula _ExecutingFormula = default(Formula);
        public Formula ExecutingFormula
        {
            get { return _ExecutingFormula; }
            set { _ExecutingFormula = value; }
        }

        private SheetCell _AssigneeCell = default(SheetCell);
        public SheetCell AssigneeCell
        {
            get { return _AssigneeCell; }
            set { _AssigneeCell = value; }
        }
    }
}
