using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MultiDimensionalSpreadsheetLibrary
{
    /// <summary>
    /// Defines values to be operated upon.
    /// </summary>
    public interface IOperand
    {
        Single Evaluate(FormulaExecutionContext context);
    }
}
