using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Text;

namespace MultiDimensionalSpreadsheetLibrary
{
    /// <summary>
    /// Defines function to be performed.
    /// </summary>
    public interface IOperator
    {
        OperandBase Run(Dictionary<String, OperandBase> dictionary, FormulaExecutionContext context);
    }
}
