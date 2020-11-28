using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExpressionEvaluatorTestConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine(Parse.ProcessCommand("1+1").ToString()); //Displays 2
            Console.WriteLine(Parse.ProcessCommand("Math.PI").ToString()); //Displays 3.14159265358979
            Console.WriteLine(Parse.ProcessCommand("Math.Abs(-22)").ToString()); //Displays 22
            Console.WriteLine(Parse.ProcessCommand("3-4+6+7+22/3+66*(55)").ToString()); //Displays 3649

            Console.ReadLine();
        }
    }
}
