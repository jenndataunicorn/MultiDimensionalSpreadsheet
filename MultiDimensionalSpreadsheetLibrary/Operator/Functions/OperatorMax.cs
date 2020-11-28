using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ssepan.Utility;

namespace MultiDimensionalSpreadsheetLibrary
{
    public class OperatorMax :
        OperatorBase
    {
        public const String Operator = "MAX";

        public OperatorMax()
        {
            Name = OperatorMax.Operator; 
            Precedence = -1;

            //define Max for OperandBase
            functorOperandBase = (dictionary, context) => new OperandLiteral(Math.Max(dictionary["Value0"].Evaluate(context), dictionary["Value1"].Evaluate(context)));
        }
    }
}
