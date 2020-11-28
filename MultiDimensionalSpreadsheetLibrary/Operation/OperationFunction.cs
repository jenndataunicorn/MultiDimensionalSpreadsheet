using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ssepan.Utility;
using System.Diagnostics;
using System.Reflection;

namespace MultiDimensionalSpreadsheetLibrary
{
    public class OperationFunction :
        OperationBase
    {
        #region Declarations
        #endregion Declarations

        #region Constructors
        public OperationFunction()
        {
            try
            {
                this.Operands = new Dictionary<String, OperandBase>();

            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
        }
        
        public OperationFunction(OperatorBase operatorBase)
        {
            try
            {
                this.Operator = operatorBase;
                this.Operands = new Dictionary<String, OperandBase>();
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
        }
        
        public OperationFunction(OperatorBase operatorBase, Dictionary<String, OperandBase> operands)
        {
            try
            {
                this.Operator = operatorBase;
                this.Operands = operands;
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
        }
        #endregion Constructors

        #region Properties
        #endregion Properties

        #region IOperation Members
        /// <summary>
        /// Run operation on child operands. Must pass through sheet cell and formula to child operands that may need to do cell lookups.
        /// </summary>
        /// <param name="context"></param>
        /// <returns></returns>
        public override Single Run(FormulaExecutionContext context)
        {
            Single returnValue = default(Single);
            OperandBase operationResult = default(OperandBase);

            try
            {
                //The operand (of type OperandOperation) that contains this operation will have its Evaluate() called; Evaluate() will call this operation's Run().
                operationResult = Operator.Run(this.Operands, context);

                returnValue = operationResult.Evaluate(context);
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
            return returnValue;
        }
        #endregion IOperation Members
    }
}
