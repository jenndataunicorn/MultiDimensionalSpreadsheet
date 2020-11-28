using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ssepan.Utility;

namespace MultiDimensionalSpreadsheetLibrary
{
    public class OperatorAssignment :
        OperatorBase
    {
        public const Char Operator = '=';

        public OperatorAssignment()
        {
            Name = OperatorAssignment.Operator.ToString();
            Precedence = -1;

            //define Assignment for OperandBase
            functorOperandBase = null;
        }
    }
}
