using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ssepan.Utility;
using System.Diagnostics;
using System.Reflection;

namespace MultiDimensionalSpreadsheetLibrary
{
    public abstract class OperatorBase :
        IOperator
    {
        #region Declarations
        #endregion Declarations

        #region Properties
        private String _Name = default(String);
        /// <summary>
        /// The operator name (and character(s)).
        /// </summary>
        public String Name
        {
            get { return _Name; }
            set { _Name = value; }
        }

        private Int32 _Precedence = default(Int32);
        /// <summary>
        /// The operator's precedence relative to all other operators.
        /// </summary>
        public Int32 Precedence
        {
            get { return _Precedence; }
            set { _Precedence = value; }
        }

        /// <summary>
        /// The operation's OperandBase implementation is stored here.
        /// </summary>
        public Func<Dictionary<String, OperandBase>, FormulaExecutionContext, OperandBase> functorOperandBase;
        #endregion Properties

        #region IOperator Members
        /// <summary>
        /// The operation's OperandBase implementation is invoked here.
        /// </summary>
        /// <param name="dictionary"></param>
        /// <param name="context"></param>
        /// <returns></returns>
        public OperandBase Run(Dictionary<String, OperandBase> dictionary, FormulaExecutionContext context)
        {
            OperandBase returnValue = default(OperandBase);
            
            try
            {
                returnValue = functorOperandBase(dictionary, context);
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);

                throw;
            }
            return returnValue;
        }
        #endregion IOperator Members
    }
}
