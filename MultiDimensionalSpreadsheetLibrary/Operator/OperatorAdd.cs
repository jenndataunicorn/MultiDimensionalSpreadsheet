using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ssepan.Utility;

namespace MultiDimensionalSpreadsheetLibrary
{
    public class OperatorAdd :
        OperatorBase
    {
        public const Char Operator = '+';

        public OperatorAdd()
        {
            Name = OperatorAdd.Operator.ToString();
            Precedence = 1;

            //define Add for OperandBase
            functorOperandBase = (dictionary, context) => new OperandLiteral(dictionary["Left"].Evaluate(context) + dictionary["Right"].Evaluate(context));
        }
    }
}
