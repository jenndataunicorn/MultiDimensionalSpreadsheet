using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ssepan.Utility;

namespace MultiDimensionalSpreadsheetLibrary
{
    public class OperatorAbs :
        OperatorBase
    {
        public const String Operator = "ABS";

        public OperatorAbs()
        {
            Name = OperatorAbs.Operator;
            Precedence = -1;

            //define Abs for OperandBase
            functorOperandBase = (dictionary, context) => new OperandLiteral(Math.Abs(dictionary["Value0"].Evaluate(context)));
        }
    }
}
