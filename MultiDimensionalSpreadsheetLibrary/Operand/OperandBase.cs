using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Reflection;

namespace MultiDimensionalSpreadsheetLibrary
{
    public abstract class OperandBase :
        IOperand
    {
        #region Declarations
        public enum OperandType
        { 
            Literal,
            CellReference,
            Operation
        }
        #endregion Declarations

        #region Constructors
        #endregion Constructors

        #region Properties
        private OperandType _Type = default(OperandType);
        public OperandType Type
        {
            get { return _Type; }
            set { _Type = value; }
        }
        #endregion Properties

        #region IOperand Members
        public abstract Single Evaluate(FormulaExecutionContext context);
        #endregion IOperand Members
    }
}
