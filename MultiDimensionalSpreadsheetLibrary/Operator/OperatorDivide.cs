using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ssepan.Utility;

namespace MultiDimensionalSpreadsheetLibrary
{
    public class OperatorDivide :
        OperatorBase
    {
        public const Char Operator = '/';

        public OperatorDivide()
        {
            Name = OperatorDivide.Operator.ToString();
            Precedence = 0;

            //define Divide for OperandBase
            functorOperandBase = (dictionary, context) => new OperandLiteral(dictionary["Left"].Evaluate(context) / dictionary["Right"].Evaluate(context));
        }
    }
}
