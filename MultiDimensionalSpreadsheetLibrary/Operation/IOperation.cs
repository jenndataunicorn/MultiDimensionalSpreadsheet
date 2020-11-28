using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Reflection;

namespace MultiDimensionalSpreadsheetLibrary
{
    /// <summary>
    /// Defines context of operation (such as parameters).
    /// </summary>
    public interface IOperation
    {
        Single Run(FormulaExecutionContext context);
    }
}
