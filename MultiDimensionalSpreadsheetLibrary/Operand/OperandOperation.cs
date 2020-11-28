using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ssepan.Utility;
using System.Diagnostics;
using System.Reflection;

namespace MultiDimensionalSpreadsheetLibrary
{
    public class OperandOperation:
        OperandBase 
    {
        #region Declarations
        #endregion Declarations

        #region Constructors
        public OperandOperation()
        { 
            try
            {
                Type = OperandType.Operation;

            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
                throw;
            }
        }

        public OperandOperation(OperationBase value)
        {
            try
            {
                Type = OperandType.Operation;
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
        private OperationBase _Value = default(OperationBase);
        #endregion Properties

        #region IOperand Members
        /// <summary>
        /// Evaluate the operand. Must run the operation, which must evaluate the child operands.
        /// Must pass down context for any children that must do cell lookups.
        /// </summary>
        /// <param name="context"></param>
        /// <returns></returns>
        public override Single Evaluate(FormulaExecutionContext context)
        {
            Single returnValue = default(Single);
            try
            {
                returnValue = _Value.Run(context);
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
