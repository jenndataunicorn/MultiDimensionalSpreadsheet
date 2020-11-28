using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ssepan.Utility;

namespace MultiDimensionalSpreadsheetLibrary
{
    public class OperatorSubtract :
        OperatorBase
    {
        public const Char Operator = '-';

        public OperatorSubtract()
        {
            Name = OperatorSubtract.Operator.ToString();
            Precedence = 1;

            //define Subtract for OperandBase
            functorOperandBase = (dictionary, context) => new OperandLiteral(dictionary["Left"].Evaluate(context) - dictionary["Right"].Evaluate(context));
        }
    }
}
