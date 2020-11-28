//#define expression_vs_operation
#define operation_statement
#define expression_statement
//#define example

using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using MultiDimensionalSpreadsheetLibrary;
using System.Diagnostics;
using System.Reflection;

namespace ExpressionTreeTestConsole
{
    class Program
    {
        //private static int _p = default(int);
        //public static int p
        //{
        //    get { return _p; }
        //    set { _p = value; }
        //}
        static void Main(string[] args)
        {
            SheetCell cellP = new SheetCell("2");
            SheetCell cellQ = new SheetCell("3");
            SheetCell cellR = new SheetCell("4");
            SheetCell cellS = new SheetCell("5");
            SheetCell cellT = new SheetCell("6");
            SheetCell cellResult = new SheetCell();

            //float p = 2.0F;
            //float q = 3.0F;
            //float r = 4.0F;
            //float s = 5.0F;
            //float t = 6.0F;
            //float result = default(float);

            //y = p + q * r - s / t
            Console.WriteLine("======");
            Console.WriteLine(" calculate it straight up ");
            Console.WriteLine(String.Format("p + q * r - s / t  =  {0} + {1} * {2} - {3} / {4}  =  {5}", Single.Parse(cellP.Value), Single.Parse(cellQ.Value), Single.Parse(cellR.Value), Single.Parse(cellS.Value), Single.Parse(cellT.Value), Single.Parse(cellP.Value) + (Single.Parse(cellQ.Value) * Single.Parse(cellR.Value)) - (Single.Parse(cellS.Value) / Single.Parse(cellT.Value))));

#if operation_statement
            //this is how MDSS formulas work now
            //these are parsed rescursively, and run as each node is identified
            OperationBase rootOperation = null;
            OperationBase parentOperation1 = null;
            OperationBase childOperation1 = null;
            OperationBase childOperation2 = null;
            rootOperation = new OperationBinary();
            parentOperation1 = new OperationBinary();
            childOperation1 = new OperationBinary();
            childOperation2 = new OperationBinary();
            //define q * r
            childOperation1.Operands["Left"] = new OperandLiteral(/*q*/Single.Parse(cellQ.Value));
            childOperation1.Operands["Right"] = new OperandLiteral(/*r*/Single.Parse(cellR.Value));
            childOperation1.Operator = new OperatorMultiply();
            //define p + (q*r)
            parentOperation1.Operands["Left"] = new OperandLiteral(/*p*/Single.Parse(cellP.Value));
            parentOperation1.Operands["Right"] = new OperandOperation(childOperation1);
            parentOperation1.Operator = new OperatorAdd();
            //define s / t
            childOperation2.Operands["Left"] = new OperandLiteral(/*s*/Single.Parse(cellS.Value));
            childOperation2.Operands["Right"] = new OperandLiteral(/*t*/Single.Parse(cellT.Value));
            childOperation2.Operator = new OperatorDivide();
            //define p+(q*r) - (s/t)
            rootOperation.Operands["Left"] = new OperandOperation(parentOperation1);
            rootOperation.Operands["Right"] = new OperandOperation(childOperation2);
            rootOperation.Operator = new OperatorSubtract();
            Console.WriteLine("======");
            Console.WriteLine(" calculate it with Operations ");
            cellResult.Value = rootOperation.Run(new FormulaExecutionContext()).ToString();
            Console.WriteLine("rootOperation:" + cellResult.Value);

#endif
#if expression_statement
            //this is how I would like MDSS formulas to work  in the future
            //these will need to be parsed recursively, 
            // but instead of being Run, each node will need to contribute an expression component 
            // to the overall composition of a statement expression.
            //the formula must be parsed when first entered or subsequently changed, 
            // and stored for later invocation.
            //the means by which parameters (cells) are attached is TBD.
            Console.WriteLine("======");
            Console.WriteLine(" do it without invoke calls on every sub-expression; embed one within another ");
            ParameterExpression parameterP = Expression.Parameter(typeof(float), "p");
            ParameterExpression parameterQ = Expression.Parameter(typeof(float), "q");
            ParameterExpression parameterR = Expression.Parameter(typeof(float), "r");
            ParameterExpression parameterS = Expression.Parameter(typeof(float), "s");
            ParameterExpression parameterT = Expression.Parameter(typeof(float), "t");

            BinaryExpression multiplyExpressionQ_R = Expression.Multiply(parameterQ, parameterR);
            BinaryExpression addExpressionP_QR = Expression.Add(parameterP, multiplyExpressionQ_R);
            BinaryExpression divideExpressionS_T = Expression.Divide(parameterS, parameterT);
            BinaryExpression subtractExpressionPQR_ST = Expression.Subtract(addExpressionP_QR, divideExpressionS_T);
            ParameterExpression parameterResultExpression = Expression.Variable(typeof(float), "result");
            ParameterExpression[] parameterResultExpressionList = new ParameterExpression[] { parameterResultExpression };
            ParameterExpression[] inputParametersExpression = new ParameterExpression[] { parameterP, parameterQ, parameterR, parameterS, parameterT };
            //adding cell *values* to array will not automatically update
            //adding lambdas that *returns* new array of current cell values may automatically update???
            Func<Object[]> inputParametersExpressionArgumentsDelegate = 
                () =>
                {
                    return new Object[] { Single.Parse(cellP.Value), Single.Parse(cellQ.Value), Single.Parse(cellR.Value), Single.Parse(cellS.Value), Single.Parse(cellT.Value) };
                };

            Delegate someFormulaExpressionDelegate =
                Expression.Lambda
                (
                    Expression.Block
                    (
                        parameterResultExpressionList,
                        Expression.Assign
                        (
                            parameterResultExpression,
                            subtractExpressionPQR_ST
                        )
                    ),
                    inputParametersExpression 
                ).Compile();
            cellResult.Value = someFormulaExpressionDelegate.DynamicInvoke(inputParametersExpressionArgumentsDelegate()).ToString();
            Console.WriteLine(String.Format("someFormulaExpressionDelegate DynamicInvoke (with p={0}):{1}", Single.Parse(cellP.Value), cellResult.Value));
#endif
            //change variable p in y = p + q * r - s / t
            Console.WriteLine();
            Console.WriteLine("======");
            cellP.Value = 10.ToString();
            Console.WriteLine(String.Format("changed variable p to {0} in y = p + q * r - s / t", Single.Parse(cellP.Value)));
            Console.WriteLine("calculate it straight up");
            Console.WriteLine(String.Format("p + q * r - s / t  =  {0} + {1} * {2} - {3} / {4}  =  {5}", Single.Parse(cellP.Value), Single.Parse(cellQ.Value), Single.Parse(cellR.Value), Single.Parse(cellS.Value), Single.Parse(cellT.Value), Single.Parse(cellP.Value) + (Single.Parse(cellQ.Value) * Single.Parse(cellR.Value)) - (Single.Parse(cellS.Value) / Single.Parse(cellT.Value))));


            Console.WriteLine("======");
            Console.WriteLine("recalc operation with new value");
            cellResult.Value = rootOperation.Run(new FormulaExecutionContext()).ToString();
            Console.WriteLine(String.Format("rootOperation uses constants, not variables, (with p={0}){1}:", Single.Parse(cellP.Value), cellResult.Value));

            Console.WriteLine("======");
            Console.WriteLine("recalc expression with new value");
            cellResult.Value = someFormulaExpressionDelegate.DynamicInvoke(inputParametersExpressionArgumentsDelegate()).ToString();
            Console.WriteLine(String.Format("someFormulaExpressionDelegate DynamicInvoke (with p={0}):{1}", Single.Parse(cellP.Value), cellResult.Value));

            //to replace formula made of operations with formula made of expressions, need to
            //a) be able to build formula dynamically at runtime from text formula string, and
            //b) invoke with cell references for assigner and assignee cells.
            //
            //Unfortunately, the Expressions above are not tied to variable references (to cells) but inputParametersExpression,
            //so formula must be handed variables (cells) at the time that the formula is executed.
            //Need to embed variable reference for this to work right with expressions; the way to do this
            // may be with closures, and I'll need to create a Lambda expression, not just a simple expression.
            //If I can hand off a lambda that will evaluate the current value of a cell, then this should work.
            //
            //Operations are using cell reference that are evaluated at the time that the formula is executed, 
            //which works but is slow. I would like to use a mechanism that can find the cell reference once 
            //when the formula is evaluated and constructed,
            //and then use that every time the formula is used, until the formula changes. However, I need to 
            //also be sure that changing the layout of the 
            //sheet will not cause problems. (I think it should not, as the criteria used to find a cell 
            //should find the same one regardless of where it is displayed.)
            //
            //Expression lambda should be an Action of T, but with no input param or return values.
            //Instead, all cell references, assigner and assignee, should be embedded as variables in a code block.

            //see also
            //http://community.bartdesmet.net/blogs/bart/archive/2009/08/10/expression-trees-take-two-introducing-system-linq-expressions-v4-0.aspx

            #region Statement
#if example
            var to = Expression.Parameter(typeof(int), "to");
            var res = Expression.Variable(typeof(List<int>), "res");
            var n = Expression.Variable(typeof(int), "n");
            var found = Expression.Variable(typeof(bool), "found");
            var d = Expression.Variable(typeof(int), "d");
            var breakOuter = Expression.Label();
            var breakInner = Expression.Label();
            var getPrimes = 
                // Func<int, List<int>> getPrimes =
                Expression.Lambda<Func<int, List<int>>>(
                    // {
                    Expression.Block(
                        // List<int> res;
                        new [] { res },
                        // res = new List<int>();
                        Expression.Assign(
                            res,
                            Expression.New(typeof(List<int>))
                        ),
                        // {
                        Expression.Block(
                            // int n;
                            new [] { n },
                            // n = 2;
                            Expression.Assign(
                                n,
                                Expression.Constant(2)
                            ),
                            // while (true)
                            Expression.Loop(
                                // {
                                Expression.Block(
                                    // if
                                    Expression.IfThen(
                                        // (!
                                        Expression.Not(
                                            // (n <= to)
                                            Expression.LessThanOrEqual(
                                                n,
                                                to
                                            )
                                        // )
                                        ),
                                        // break;
                                        Expression.Break(breakOuter)
                                    ),
                                    // {
                                    Expression.Block(
                                        // bool found;
                                        new[] { found },
                                        // found = false;
                                        Expression.Assign(
                                            found,
                                            Expression.Constant(false)
                                        ),
                                        // {
                                        Expression.Block(
                                            // int d;
                                            new [] { d },
                                            // d = 2;
                                            Expression.Assign(
                                                d,
                                                Expression.Constant(2)
                                            ),
                                            // while (true)
                                            Expression.Loop(
                                                // {
                                                Expression.Block(
                                                    // if
                                                    Expression.IfThen(
                                                        // (!
                                                        Expression.Not(
                                                            // d <= Math.Sqrt(n)
                                                            Expression.LessThanOrEqual(
                                                                d,
                                                                Expression.Convert(
                                                                    Expression.Call(
                                                                        null,
                                                                        typeof(Math).GetMethod("Sqrt"),
                                                                        Expression.Convert(
                                                                            n,
                                                                            typeof(double)
                                                                        )
                                                                    ),
                                                                    typeof(int)
                                                                )
                                                            )
                                                        // )
                                                        ),
                                                        // break;
                                                        Expression.Break(breakInner)
                                                    ),
                                                    // {
                                                    Expression.Block(
                                                        // if (n % d == 0)
                                                        Expression.IfThen(
                                                            Expression.Equal(
                                                                Expression.Modulo(
                                                                    n,
                                                                    d
                                                                ),
                                                                Expression.Constant(0)
                                                            ),
                                                            // {
                                                            Expression.Block(
                                                                // found = true;
                                                                Expression.Assign(
                                                                    found,
                                                                    Expression.Constant(true)
                                                                ),
                                                                // break;
                                                                Expression.Break(breakInner)
                                                            // }
                                                            )
                                                        )
                                                    // }
                                                    ),
                                                    // d++;
                                                    Expression.PostIncrementAssign(d)
                                                // }
                                                ),
                                                breakInner
                                            )
                                        ),
                                        // if
                                        Expression.IfThen(
                                            // (!found)
                                            Expression.Not(found),
                                            //    res.Add(n);
                                            Expression.Call(
                                                res,
                                                typeof(List<int>).GetMethod("Add"),
                                                n
                                            )
                                        )
                                    ),
                                    // n++;
                                    Expression.PostIncrementAssign(n)
                                // }
                                ),
                                breakOuter
                            )
                        ),
                        res
                    ),
                    to
                // }
                ).Compile();

            foreach (var num in getPrimes(100))
                Console.WriteLine(num);
#endif
            #endregion Statement

            #region article
            //see also
            //http://stackoverflow.com/questions/1644146/user-defined-formulas-in-c-sharp

            //I have written an open source project, Dynamic Expresso, that can convert text expression written using a C# syntax into delegates (or expression tree). Expressions are parsed and transformed into Expression Trees without using compilation or reflection.

            //You can write something like:

            //var interpreter = new Interpreter();
            //var result = interpreter.Eval("8 / 2 + 2");

            //or

            //var interpreter = new Interpreter()
            //                .SetVariable("service", new ServiceExample());

            //string expression = "x > 4 ? service.SomeMethod() : service.AnotherMethod()";

            //Lambda parsedExpression = interpreter.Parse(expression, 
            //                        new Parameter("x", typeof(int)));

            //parsedExpression.Invoke(5);

            //My work is based on Scott Gu article http://weblogs.asp.net/scottgu/archive/2008/01/07/dynamic-linq-part-1-using-the-linq-dynamic-query-library.aspx .

            //see also
            //https://github.com/davideicardi/DynamicExpresso/
            #endregion Article

            #region Article
            //see also
           

            #endregion Article

            Console.ReadLine();
        }
    }
}