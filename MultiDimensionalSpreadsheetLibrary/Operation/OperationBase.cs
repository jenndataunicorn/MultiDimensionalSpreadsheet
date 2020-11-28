using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ssepan.Utility;
using System.Diagnostics;
using System.Reflection;

namespace MultiDimensionalSpreadsheetLibrary
{
    public abstract class OperationBase :
        IOperation
    {
        #region Declarations
        #endregion Declarations

        #region Properties
        private OperatorBase _Operator = default(OperatorBase);
        public OperatorBase Operator
        {
            get { return _Operator; }
            set { _Operator = value; }
        }

        private Dictionary<String, OperandBase> _Operands = default(Dictionary<String, OperandBase>);
        public Dictionary<String, OperandBase> Operands
        {
            get { return _Operands; }
            set { _Operands = value; }
        }
        #endregion Properties

        #region IOperation Members
        /// <summary>
        /// Run operation on child operands. Must pass through sheet cell and formula to child operands that may need to do cell lookups.
        /// </summary>
        /// <param name="context"></param>
        /// <returns></returns>
        public abstract Single Run(FormulaExecutionContext context);
        #endregion IOperation Members
    }
}
