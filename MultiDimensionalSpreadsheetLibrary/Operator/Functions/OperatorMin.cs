using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ssepan.Utility;

namespace MultiDimensionalSpreadsheetLibrary
{
    public class OperatorMin :
        OperatorBase
    {
        public const String Operator = "MIN";

        public OperatorMin()
        {
            Name = OperatorMin.Operator;
            Precedence = -1;

            //define Min for OperandBase
            functorOperandBase = (dictionary, context) => new OperandLiteral(Math.Min(dictionary["Value0"].Evaluate(context), dictionary["Value1"].Evaluate(context)));
        }
    }
}
