using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ssepan.Utility;
using System.Diagnostics;
using System.Reflection;

namespace MultiDimensionalSpreadsheetLibrary
{
    public class OperandLiteral:
        OperandBase 
    {
        #region Declarations
        #endregion Declarations

        #region Constructors
        public OperandLiteral()
        { 
            try
            {
                Type = OperandType.Literal;

            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
                throw;
            }
        }

        public OperandLiteral(Single value)
        {
            try
            {
                Type = OperandType.Literal;
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
        private Single _Value = default(Single);
        #endregion Properties

        #region IOperand Members
        /// <summary>
        /// Evaluate the operand. Must return the literal value.
        /// </summary>
        /// <param name="context"></param>
        /// <returns></returns>
        public override Single Evaluate(FormulaExecutionContext context)
        {
            Single returnValue = default(Single);
            try
            {
                returnValue = _Value;
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
