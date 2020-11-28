using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ssepan.Utility;

namespace MultiDimensionalSpreadsheetLibrary
{
    public class OperatorMultiply :
        OperatorBase
    {
        public const Char Operator = '*';

        public OperatorMultiply()
        {
            Name = OperatorMultiply.Operator.ToString();
            Precedence = 0;

            //define Multiply for OperandBase
            functorOperandBase = (dictionary, context) => new OperandLiteral(dictionary["Left"].Evaluate(context) * dictionary["Right"].Evaluate(context));
        }
    }
}
