using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.Serialization;
using System.Linq;
using System.Linq.Expressions;
using System.Text.RegularExpressions;
using System.Xml.Serialization;
using System.Diagnostics;
using System.Reflection;
using Ssepan.Application;
using Ssepan.Application.MVC;
using Ssepan.Collections;
using Ssepan.Utility;
using MultiDimensionalSpreadsheetLibrary.MVC;

namespace MultiDimensionalSpreadsheetLibrary
{
    /// <summary>
    /// </summary>
    //[Serializable()]
    [DataContract(IsReference = true)]
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public class Formula :
        SettingsComponentBase,
        IEquatable<Formula>,
        INotifyPropertyChanged
    {
        #region Declarations

        private static List<OperatorBase> arithmeticOperators = default(List<OperatorBase>);
        private static List<OperatorBase> functionOperators = default(List<OperatorBase>);

        //private const String RegExFindOutermostParenthesisContents = @"(?<!\\)\((\\\(|\\\)|[^\(\)]|(?<!\\)\(.*(?<!\\)\))*(?<!\\)\)";
        private const String RegExFindInnermostParenthesisAndContents = @"\([^\(\)]*\)";
        private const String RegExFindInnermostBracesAndContents = @"\{[^\{\}]*\}";
        private const String RegExFindFunctionNameDelimitersAndContents = @"\@[^\@\{]*\{";
        private const Char LeftParenthesis = '(';
        private const Char RightParenthesis = ')';
        private const Char LeftBrace = '{';
        private const Char RightBrace = '}';
        private const String ParsingVariableIndicator = @"#";
        private const Char FunctionParameterDelimiter = ',';
        
        public const String NewItemName = "(new formula)";
        public const String FunctionIndicator = @"@"; //Reserved
        public const String VariableIndicator = @"$"; //Reserved
        /*
        y=(p+q)*((r-s)/t)
        */
        #endregion Declarations

        #region Constructors
        static Formula()
        {
            try
            {
                //create a list of arithmetic operators found on the right side of an assignment; sort them by precedence, descending
                arithmeticOperators = new List<OperatorBase>();
                arithmeticOperators.Add(new OperatorAdd());
                arithmeticOperators.Add(new OperatorSubtract());
                arithmeticOperators.Add(new OperatorMultiply());
                arithmeticOperators.Add(new OperatorDivide());
                arithmeticOperators.Sort((operator1, operator2) => operator1.Precedence.CompareTo(operator2.Precedence));
                arithmeticOperators.Reverse();

                //create a list of function operators found on the right side of an assignment
                functionOperators = new List<OperatorBase>();
                functionOperators.Add(new OperatorAbs());
                functionOperators.Add(new OperatorMax());
                functionOperators.Add(new OperatorMin());
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
                throw;
            }
        }

        public Formula()
        {
            try
            {
                //ParsingVariables = new Dictionary<String, String>();
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
                throw;
            }
        }

        /// <summary>
        /// </summary>
        /// <param name="value"></param>
        public Formula(String value) :
            this()
        {
            try
            {
                Value = value;
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
                throw;
            }
        }

        /// <summary>
        /// Parent  is the sheet 
        /// </summary>
        /// <param name="value"></param>
        /// <param name="parent"></param>
        public Formula(String value, Sheet parent) :
            this(value)
        {
            try
            {
                Parent = parent;
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
                throw;
            }
        }
        #endregion Constructors

        #region INotifyPropertyChanged support
        public event PropertyChangedEventHandler PropertyChanged;

        void OnPropertyChanged(String propertyName)
        {
            try
            {
                if (this.PropertyChanged != null)
                {
                    this.PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
#if debug
                    Log.Write(propertyName, EventLogEntryType.Error);
#endif
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
                //throw;
            }
        }
        #endregion INotifyPropertyChanged support

        #region IEquatable<Formula>
        /// <summary>
        /// Compare property values of this object to another.
        /// </summary>
        /// <param name="anotherSettings"></param>
        /// <returns></returns>
        public Boolean Equals(Formula other)
        {
            Boolean returnValue = default(Boolean);

            try
            {
                if (this == other)
                {
                    returnValue = true;
                }
                else
                {
                    if (this.Value != other.Value)
                    {
                        returnValue = false;
                    }
                    else
                    {
                        returnValue = true;
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
                throw;
            }

            return returnValue;
        }

        #endregion IEquatable<Formula>

        #region Properties
        #region NonPersisted Properties
        private String _ErrorMessage = default(String);
        [XmlIgnore]
        public String ErrorMessage
        {
            get { return _ErrorMessage; }
            set { _ErrorMessage = value; }
        }

        [XmlIgnore]
        public Boolean IsValid
        {
            get
            {
                Boolean returnValue = default(Boolean);
                try
                {
                    if (Valid())
                    {
                        returnValue = true;
                    }
                }
                catch (Exception ex)
                {
                    Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
                }
                return returnValue;
            }
            //set { _Valid = value; }
        }

        [XmlIgnore]
        public Boolean IsCompiled
        {
            get
            {
                Boolean returnValue = default(Boolean);
                try
                {
                    if (Compiled())
                    {
                        returnValue = true;
                    }
                }
                catch (Exception ex)
                {
                    Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
                }
                return returnValue;
            }
        }

        private List<CategoryItem> _AssigneeCriteria = default(List<CategoryItem>);
        /// <summary>
        /// Category item criteria identifying cells to which the formula will apply.
        /// It is built when formula is validated. 
        /// When a cell needs a value, AssigneeCriteria will be run  
        ///  and the specific cell sought in that list. 
        /// If found, the AssignerOperand is evaluated and the result assigned to the cell.
        /// </summary>
        [XmlIgnore]
        public List<CategoryItem> AssigneeCriteria
        {
            get { return _AssigneeCriteria; }
            set { _AssigneeCriteria = value; }
        }

        private OperandBase _AssignerOperand = default(OperandBase);
        /// <summary>
        /// The operand that will be evaluated for a given matching cell.
        /// It is composed when the formula is validated.
        /// When a cell needs a value, and it matches the assignee criteria for this formula, 
        ///  the root assigner operand for this formula is evaluated.
        /// 1) Literal operands return  the value that they contain.
        /// 2) Reference operands use their own category item criteria 
        ///  to narrow down the list of cells to which the formula applies. Then the specific cell's criteria 
        ///  is used to identify categories not in the criteria, and use those to narrow down to a single source cell.
        /// 3) Operation operands perform their operation and trigger the evaulation of their child operands.
        /// </summary>
        [XmlIgnore]
        public OperandBase AssignerOperand
        {
            get { return _AssignerOperand; }
            set { _AssignerOperand = value; }
        }

        private Dictionary<String, String> _ParsingVariables = new Dictionary<String, String>();
        /// <summary>
        /// Variables used for substitution during parsing.
        /// For now, these are only used for pre-parsing expressions in parenthesis.
        /// </summary>
        [XmlIgnore]
        public Dictionary<String, String> ParsingVariables
        {
            get { return _ParsingVariables; }
            set { _ParsingVariables = value; }
        }
        
        //TODO:Need temp value for a cell that needs a value; must be reentrant! 
        // Keep a list, add when starting and remove when finished; can also use to check for circular references.
        #endregion NonPersisted Properties

        #region Persisted Properties
        private String _Value = default(String);
        [DataMember]
        public String Value
        {
            get { return _Value; }
            set
            {
                _Value = value;
                this.OnPropertyChanged("Value");

                //any change to the formula source invalidates the compiled formula
                Delegate = default(Delegate);
            }
        }

        private Sheet _Parent = default(Sheet);
        [DataMember]
        public Sheet Parent
        {
            get { return _Parent; }
            set
            {
                _Parent = value;
                this.OnPropertyChanged("Parent");
            }
        }

        private Delegate _Delegate = default(Delegate);
        [DataMember]
        public Delegate Delegate
        {
            get { return _Delegate;  }
            set 
            {
                _Delegate = value;
                this.OnPropertyChanged("Delegate");
            }
        }
        #endregion Persisted Properties
        #endregion Properties

        #region Methods
        #region Public Methods
        /// <summary>   
        /// validate settings entered by user
        /// because this is being called by grid databinding, it is run on every refresh
        /// </summary>
        public Boolean Valid()
        {
            Boolean returnValue = default(Boolean);
            String errorMessage = default(String);

            try
            {
                ErrorMessage = String.Empty;

                if ((Value == null) || (Value == String.Empty))
                {
                    ErrorMessage = String.Format("Formula Value must contain a value.");
                }
                //don't force re-parsing of the source if the delegate is compiled
                else if (!Compiled())
                {
                    if (!Parse(Value, ref errorMessage))
                    {
                        ErrorMessage = String.Format("The formula is not valid. \n'{0}': '{1}'", errorMessage, Value);
                    }
                    else
                    { 
                        returnValue = true;
                    }
                }
                else
                {
                    returnValue = true;
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
                throw;
            }

            return returnValue;
        }

        /// <summary>   
        /// check for presence of compiled expression
        /// </summary>
        public Boolean Compiled()
        {
            Boolean returnValue = default(Boolean);

            try
            {
                ErrorMessage = String.Empty;

                if (Delegate == default(Delegate))
                {
                    ErrorMessage = String.Format("The formula is not compiled.\n{0}", Value);
                }
                else
                {
                    returnValue = true;
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
                throw;
            }

            return returnValue;
        }

        /// <summary>   
        /// check for presence of assignee operand
        /// </summary>
        public Boolean HasAssignee()
        {
            Boolean returnValue = default(Boolean);

            try
            {
                ErrorMessage = String.Empty;

                if ((AssigneeCriteria == null) || (AssigneeCriteria.Count == 0))
                {
                    returnValue = false;
                    ErrorMessage = String.Format("Assignee criteria not defined.");
                }
                else
                {
                    returnValue = true;
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
                throw;
            }

            return returnValue;
        }

        /// <summary>   
        /// check for presence of assigner operand
        /// </summary>
        public Boolean HasAssigner()
        {
            Boolean returnValue = default(Boolean);

            try
            {
                ErrorMessage = String.Empty;

                if (AssignerOperand == null)
                {
                    returnValue = false;
                    ErrorMessage = String.Format("Assigner operand not defined.");
                }
                else
                {
                    returnValue = true;
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
                throw;
            }

            return returnValue;
        }

        /// <summary>
        /// validate settings determined at the time of execution
        /// </summary>
        public Boolean Complete()
        {
            Boolean returnValue = default(Boolean);

            try
            {
                ErrorMessage = String.Empty;

                if (!Compiled())
                {
                    returnValue = false;
                }
                else if (!Valid())
                {
                    returnValue = false;
                }
                else if (!HasAssignee())
                {
                    returnValue = false;
                }
                else if (!HasAssigner())
                {
                    returnValue = false;
                }
                else
                {
                    returnValue = true;
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
                throw;
            }

            return returnValue;
        }

        /// <summary>
        /// Determine if passed cell is among cells to which the formula criteria are applicable.
        /// </summary>
        /// <param name="cell">Specific sheet cell seeking value.</param>
        /// <returns></returns>
        public Boolean IsApplicableFormula(SheetCell cell)
        {
            Boolean returnValue = default(Boolean);
            List<SheetCell> formulaSearchResults = default(List<SheetCell>);
            List<SheetCell> cellSearchResults = default(List<SheetCell>);
            Sheet sheet = default(Sheet);

            try
            {
                if (this.Complete())
                {
                    //locate parent sheet, containing cells
                    sheet = this.Parent;
                    if (sheet == null)
                    {
                        throw new ApplicationException(String.Format("Unable to find parent sheet for formula: '{0}'", this.Value));
                    }

                    //run formula assigner criteria
                    formulaSearchResults = Sheet.GetCellsWhereAllSearchCategoryItemsMatchAnyInGivenCell(sheet.Cells.ToList<SheetCell>(), AssigneeCriteria);
                    if ((formulaSearchResults == null) || (formulaSearchResults.Count == 0))
                    {
                        throw new ApplicationException(String.Format("Unable to find cells for parent sheet of formula '{0}'", this.Value));
                    }

                    //find single cell in results
                    cellSearchResults = Sheet.GetCellsWhereAllSearchCategoryItemsMatchAnyInGivenCell(formulaSearchResults, cell.CategoryItems.ToList<CategoryItem>());
                    if ((cellSearchResults == null) || (cellSearchResults.Count == 0))
                    {
                        //cell not found; not an error
                    }
                    else
                    {
                        returnValue = true;
                    }
                }
                else
                { 
                    //skip over any invalid formulae
                    Log.Write(this.ErrorMessage, EventLogEntryType.Error);
                }
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
                //throw;
            }
            return returnValue;
        }

        /// <summary>
        /// Run formula and apply to passed cell.
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public Boolean Run(SheetCell cell)
        {
            Boolean returnValue = default(Boolean);
            FormulaExecutionContext context = default(FormulaExecutionContext);

            try
            {
                //validate settings determined at runtime
                if (!this.Complete())
                {
                    throw new ApplicationException(String.Format("{0}", this.ErrorMessage));
                }

                //run operation
                context = new FormulaExecutionContext(this, cell);
                cell.ValueChangingProgrammatically = true;
                //TODO:invoke with parameter list
                cell.Value = this.Delegate.DynamicInvoke(new ParameterExpression[] { /*Expression.Parameter(typeof(float), "p")*/ }).ToString(); //AssignerOperand.Evaluate(context).ToString();
                //TODO:mvoe this or something like it to wwhere formula is changed.
                //cell.Value = AssignerOperand.Evaluate(context).ToString();
                cell.ValueChangingProgrammatically = false;

                returnValue = true;
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
                cell.ValueChangingProgrammatically = false;
                throw;
            }
            return returnValue;
        }
        #endregion Public Methods

        #region Private Methods
        /// <summary>
        /// Parse the formula to get assignee and assigner.
        /// </summary>
        /// <param name="value"></param>
        /// <param name="errorMessage"></param>
        /// <returns></returns>
        private Boolean Parse(String value, ref String errorMessage)
        {
            Boolean returnValue = default(Boolean);
            String[] tokens = default(String[]);
            String assigneeToken = default(String);
            String assignerToken = default(String);

            try
            {
                tokens = value.Split(OperatorAssignment.Operator);
                if (tokens.Length != 2)
                {
                    throw new ArgumentException(String.Format("Incorrectly formatted formula; there does not appear to be an assignment: '{0}'", value));
                }

                assigneeToken = tokens[0].Trim();
                if (!ParseAssignee(assigneeToken, ref errorMessage))
                {
                    throw new ApplicationException(String.Format("Unable to parse assignee: '{0}'", assigneeToken));
                }

                assignerToken = tokens[1].Trim();
                if (!ParseAssigner(assignerToken, ref errorMessage))
                {
                    throw new ApplicationException(String.Format("Unable to parse assigner: '{0}'", assignerToken));
                }

                returnValue = true;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
                //throw;
            }
            return returnValue;
        }

        /// <summary>
        /// Parse the assignee to get the category item criteria list
        /// </summary>
        /// <param name="assigneeToken"></param>
        /// <param name="errorMessage"></param>
        /// <returns></returns>
        private Boolean ParseAssignee(String assigneeToken, ref String errorMessage)
        {
            Boolean returnValue = default(Boolean);
            Sheet parentSheet = default(Sheet);
            List<CategoryItem> criteria = default(List<CategoryItem>);

            try
            {
                if (assigneeToken == String.Empty)
                {
                    throw new ArgumentException(String.Format("Incorrectly formatted assignee; the assignee address is not defined: '{0}'", assigneeToken));
                }
                assigneeToken.Trim();

                parentSheet = SettingsController<Settings>.Settings.Sheets.Find(s => s.Equals(this.Parent));
                //parentSheet = SettingsController<Settings>.Settings.Sheets.Find(s => s.Name == this.Parent);
                if (parentSheet == null)
                {
                    throw new ApplicationException(String.Format("Unable to find sheet '{0}' for formula '{1}'.", this.Parent, this.Value));
                }

                criteria = parentSheet.GetCategoryItemsForAddressName(assigneeToken, true);
                if ((criteria == null) || (criteria.Count < 1))
                {
                    throw new ApplicationException(String.Format("Unable to find category-item criteria in sheet '{0}' for formula '{1}'.", this.Parent, this.Value));
                }

                //store criteria
                AssigneeCriteria = criteria;

                returnValue = true;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
                //throw;
            }
            return returnValue;
        }

        /// <summary>
        /// Parse the assigner to get the tree of operations and their operands (of type literal, reference, or operation)
        /// </summary>
        /// <param name="assignerToken"></param>
        /// <param name="errorMessage"></param>
        /// <returns></returns>
        private Boolean ParseAssigner(String assignerToken, ref String errorMessage)
        {
            Boolean returnValue = default(Boolean);
            OperandBase operand = default(OperandBase);

            try
            {
                if (assignerToken == String.Empty)
                {
                    throw new ArgumentException(String.Format("Incorrectly formatted formula; the assigner is not defined: '{0}'", assignerToken));
                }

                //init parsing variables dictionary
                if (ParsingVariables == null)
                {
                    ParsingVariables = new Dictionary<String, String>();
                }
                else
                {
                    ParsingVariables.Clear();
                }

                //pre-parse token for parenthesis, and replace expressions in parenthesis with variables
                if (!ParseParenthesis(ref assignerToken, ref errorMessage))
                {
                    throw new ApplicationException(errorMessage);
                }

                //recursively parse assignerToken and build operation tree
                if (!ParseToken(assignerToken, ref operand, ref errorMessage))
                {
                    throw new ApplicationException(errorMessage);
                }

                //clear parsing variables dictionary
                ParsingVariables.Clear();

                AssignerOperand = operand;

                returnValue = true;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
                //throw;
            }
            return returnValue;
        }

        /// <summary>
        /// Pre-parse the token for parenthesis, and replace (from the inside out) 
        /// each expression in parenthesis with a variable that acts as a placeholder.
        /// The expression is stored in a dictionary identified by @variable_name.
        /// The variable is used later during regular parsing to look up the expression.
        /// </summary>
        /// <param name="assignerToken"></param>
        /// <param name="errorMessage"></param>
        /// <returns></returns>
        private Boolean ParseParenthesis(ref String assignerToken, ref String errorMessage)
        {
            Boolean returnValue = default(Boolean);
            Regex regEx = default(Regex);
            Int32 countOfLeftParenthesis = default(Int32);
            Int32 countOfRightParenthesis = default(Int32);
            Match match = default(Match);
            String variableName = default(String);
            String variableValue = default(String);

            try
            {
                countOfLeftParenthesis = (from Char ch in assignerToken
                                          where ch == Formula.LeftParenthesis
                                          select ch).Count();
                countOfRightParenthesis = (from Char ch in assignerToken
                                           where ch == Formula.RightParenthesis
                                           select ch).Count();
                if ((countOfLeftParenthesis > 0) || (countOfRightParenthesis > 0))
                {
                    //parenthesis found

                    if (countOfLeftParenthesis != countOfRightParenthesis)
                    {
                        //mismatched pair
                        throw new ApplicationException(String.Format("Mis-matched parenthesis in token: '{0}'", assignerToken));
                    }

                    regEx = new Regex(Formula.RegExFindInnermostParenthesisAndContents);
                    while ((assignerToken.Contains(Formula.LeftParenthesis)) || (assignerToken.Contains(Formula.RightParenthesis)))
                    {
                        //process until there are no more pairs left

                        match = regEx.Match(assignerToken);
                        if ((match.Value == null) || (match.Value == String.Empty))
                        {
                            //pair did not match pattern
                            throw new ApplicationException(String.Format("Parenthesis in token not formatted correctly: '{0}'", assignerToken));
                        }

                        //generate variable
                        variableName = String.Format("{0}{1}", Formula.ParsingVariableIndicator, Guid.NewGuid().ToString().Replace('-'.ToString(), String.Empty));
                        variableValue = match.Value.Replace(Formula.LeftParenthesis.ToString(), String.Empty).Replace(Formula.RightParenthesis.ToString(), String.Empty);
                        if ((variableValue == null) || (variableValue == String.Empty))
                        {
                            //pair did not contain an expression
                            throw new ApplicationException(String.Format("Value in parenthesis in token not formatted correctly: '{0}'", assignerToken));
                        }

                        //store variable in dictionary
                        ParsingVariables.Add(variableName, variableValue);

                        //replace expression (including parenthesis) in token w/ variable name
                        assignerToken = assignerToken.Replace(match.Value, variableName);
                    }
                }

                returnValue = true;
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
                //throw;
            }
            return returnValue;
        }
        
        /// <summary>
        /// Parse the string token to get an operand of type Literal, Reference, or Operation.
        /// 1) Literal operands return  the value that they contain.
        /// 2) Reference operands use their own category item criteria 
        ///  to narrow down the list of cells to which the formula applies. Then the specific cell's criteria 
        ///  is used to identify categories not in the criteria, and use those to narrow down to a single source cell.
        /// 3) Operation operands perform their operation and trigger the evaulation of their child operands.
        /// </summary>
        /// <param name="token"></param>
        /// <param name="operand"></param>
        /// <param name="errorMessage"></param>
        /// <returns></returns>
        private Boolean ParseToken(String token, ref OperandBase operand, ref String errorMessage)
        {
            Boolean returnValue = default(Boolean);
            String[] tokens = default(String[]);
            String leftToken = default(String);
            String rightToken = default(String);
            Boolean isFoundOperator = default(Boolean);
            OperatorBase foundOperator = default(OperatorBase);
            OperandBase leftOperand = default(OperandBase);
            OperandBase rightOperand = default(OperandBase);
            OperationBase operation = default(OperationBase);

            try
            {
                //look for operators, if any, in reverse precedence order
                foreach (OperatorBase op in arithmeticOperators)
                {
                    if (token.Contains(op.Name))
                    {
                        isFoundOperator = true;
                        foundOperator = op;
                        
                        break;
                    }
                }

                //create an array of one or two sub-tokens
                if (isFoundOperator)
                {
                    //split into 2 parts only
                    tokens = token.Split(new Char[] { foundOperator.Name[0] }, 2);
                }
                else
                { 
                    //create list with single item
                    String[] newList =  { token };
                    tokens = newList;
                }
                
                //process sub-tokens
                if (tokens.Length == 0)
                {
                    throw new ArgumentException(String.Format("Incorrectly formatted formula; there does not appear to be an operand: '{0}'", token));
                }
                else if (tokens.Length == 1)
                {
                    //no operator; process single operand 
                    leftToken = tokens[0].Trim();
                    if (!ParseOperand(leftToken, ref leftOperand, ref errorMessage))
                    {
                        throw new ApplicationException(errorMessage);
                    }
                    operand = leftOperand;

                    returnValue = true;
                }
                else if (tokens.Length == 2)
                {
                    //operator found; process dual operands with recursive call
                    leftToken = tokens[0].Trim();
                    if (!ParseToken(leftToken, ref leftOperand, ref errorMessage))
                    {
                        throw new ApplicationException(errorMessage);
                    }

                    rightToken = tokens[1].Trim();
                    if (!ParseToken(rightToken, ref rightOperand, ref errorMessage))
                    {
                        throw new ApplicationException(errorMessage);
                    }

                    //create operation from left/right operands and operator, and create OperandOperation
                    //create OperationBinary which configures 2 OperandBase dictionary entries called Left, Right.
                    operation = new OperationBinary();
                    operation.Operands["Left"] = leftOperand;
                    operation.Operands["Right"] = rightOperand;
                    operation.Operator = foundOperator;

                    operand = new OperandOperation(operation);

                    returnValue = true;
                }
                else if (tokens.Length > 2)
                {
                    throw new ArgumentException(String.Format("Unexpected formula format: '{0}'", token));
                }

            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
                //throw;
            }
            return returnValue;
        }

        /// <summary>
        /// Parse token to identify whether it is a pre-parsing variable, a literal, or a cell-reference.
        /// </summary>
        /// <param name="token"></param>
        /// <param name="operand"></param>
        /// <param name="errorMessage"></param>
        /// <returns></returns>
        private Boolean ParseOperand(String token, ref OperandBase operand, ref String errorMessage)
        {
            Boolean returnValue = default(Boolean);
            Single numericResult = default(Single);
            Sheet parentSheet = default(Sheet);
            List<CategoryItem> criteria = default(List<CategoryItem>);
            String variableValue = default(String);

            try
            {
                //trap string leterals
                if ((token.Contains("'")) || (token.Contains("\"")))
                {
                    throw new ApplicationException(String.Format("String literals not supported: '{0}'", token));
                }
                //trap parentheses
                if ((token.Contains("(")) || (token.Contains(")")))
                {
                    throw new ApplicationException(String.Format("Parentheses found after pre-processing: '{0}'", token));
                }

                token = token.Trim();

                if (token.StartsWith(Formula.ParsingVariableIndicator))
                {
                    //token is a parsing variable; look up value and process token as an expression

                    //look up parsing variable
                    if (!ParsingVariables.TryGetValue(token, out variableValue))
                    {
                        throw new ApplicationException(String.Format("Unable to look up variable value for variable name: '{0}'", token));
                    }
                    
                    //parse expression and return operand
                    if (!ParseToken(variableValue, ref operand, ref errorMessage))
                    {
                        throw new ApplicationException(errorMessage);
                    }

                    returnValue = true;
                }
                else if (token.StartsWith(Formula.FunctionIndicator))
                {
                    //token is a function

                    //parse function and return operand
                    if (!ParseFunction(token, ref operand, ref errorMessage))
                    {
                        throw new ApplicationException(errorMessage);
                    }

                    returnValue = true;
                }
                else if (Single.TryParse(token, out numericResult))
                {
                    //token is a numeric literal
                    //return literal operand with numeric value
                    operand = new OperandLiteral(numericResult);

                    returnValue = true;
                }
                else
                {
                    //token should be reference-operand
                    parentSheet = SettingsController<Settings>.Settings.Sheets.Find(s => s.Equals(this.Parent));
                    //parentSheet = SettingsController<Settings>.Settings.Sheets.Find(s => s.Name == this.Parent);
                    if (parentSheet == null)
                    {
                        throw new ApplicationException(String.Format("Unable to find sheet '{0}' for formula '{1}'.", this.Parent, this.Value));
                    }

                    criteria = parentSheet.GetCategoryItemsForAddressName(token, true);
                    if ((criteria == null) || (criteria.Count < 1))
                    {
                        throw new ApplicationException(String.Format("Unable to find category-item criteria in sheet '{0}' for formula '{1}'.", this.Parent, this.Value));
                    }

                    //store criteria
                    operand = new OperandReference(criteria);

                    returnValue = true;
                }
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
                //throw;
            }
            return returnValue;
        }

        /// <summary>
        /// Parse the string token to get an operand of type Operation, 
        /// whose operation is of type Function(?), 
        /// and where the operator is the defined in the token.
        /// </summary>
        /// <param name="token"></param>
        /// <param name="operand"></param>
        /// <param name="errorMessage"></param>
        /// <returns></returns>
        private Boolean ParseFunction(String token, ref OperandBase operand, ref String errorMessage)
        {
            Boolean returnValue = default(Boolean);
            Boolean isFoundOperator = default(Boolean);
            OperatorBase foundOperator = default(OperatorBase);
            Int32 countOfLeftBraces = default(Int32);
            Int32 countOfRightBraces = default(Int32);
            Regex regEx = default(Regex);
            Match match = default(Match);
            String parametersToken = default(String);
            String[] parameterTokens = default(String[]);
            String functionNameToken = default(String);
            Dictionary<String, OperandBase> parameterDictionary = default(Dictionary<String, OperandBase>);
            OperandBase parameterOperand = default(OperandBase);
            Int32 parameterIndex = default(Int32);

            try
            {
                countOfLeftBraces = (from Char ch in token
                                          where ch == Formula.LeftBrace
                                          select ch).Count();
                countOfRightBraces = (from Char ch in token
                                           where ch == Formula.RightBrace
                                           select ch).Count();
                if ((countOfLeftBraces > 0) || (countOfRightBraces > 0))
                {
                    //braces found

                    if (countOfLeftBraces != countOfRightBraces)
                    {
                        //mismatched pair
                        throw new ApplicationException(String.Format("Mis-matched braces in token: '{0}'", token));
                    }

                    //get parameters
                    
                    regEx = new Regex(Formula.RegExFindInnermostBracesAndContents);
                    match = regEx.Match(token);
                    if ((match.Value == null) || (match.Value == String.Empty))
                    {
                        //pair did not match pattern
                        throw new ApplicationException(String.Format("Braces in token not formatted correctly: '{0}'", token));
                    }

                    parametersToken = match.Value.Replace(Formula.LeftBrace.ToString(), String.Empty).Replace(Formula.RightBrace.ToString(), String.Empty);
                    if ((parametersToken == null) || (parametersToken == String.Empty))
                    {
                        //pair did not contain an expression
                        throw new ApplicationException(String.Format("Value in parenthesis in token not formatted correctly: '{0}'", token));
                    }

                    parameterTokens = parametersToken.Split(Formula.FunctionParameterDelimiter);
                    if (parameterTokens.Length == 0)
                    {
                        //pair did not contain any values
                        throw new ApplicationException(String.Format("Value in parenthesis in token not formatted correctly: '{0}'", token));
                    }

                    //get function name

                    regEx = new Regex(Formula.RegExFindFunctionNameDelimitersAndContents);
                    match = regEx.Match(token);
                    if ((match.Value == null) || (match.Value == String.Empty))
                    {
                        //pair did not match pattern
                        throw new ApplicationException(String.Format("Function name delimiters in token not formatted correctly: '{0}'", token));
                    }

                    functionNameToken = match.Value.Replace(Formula.FunctionIndicator.ToString(), String.Empty).Replace(Formula.LeftBrace.ToString(), String.Empty);
                    if ((functionNameToken == null) || (functionNameToken == String.Empty))
                    {
                        //pair did not contain an expression
                        throw new ApplicationException(String.Format("Function name in token not formatted correctly: '{0}'", token));
                    }

                    //look for operators, if any
                    foreach (OperatorBase op in functionOperators)
                    {
                        if (functionNameToken.ToUpper().Contains(op.Name))
                        {
                            isFoundOperator = true;
                            foundOperator = op;

                            break;
                        }
                    }

                    //get function

                    if (isFoundOperator)
                    {
                        //build operation operand with function operation consisting of operator and each parameter token parsed into an operand.
                        parameterDictionary = new Dictionary<String, OperandBase>();
                        foreach (String parameterToken in parameterTokens)
                        {
                            if (!ParseToken(parameterToken, ref parameterOperand, ref errorMessage))
                            {
                                //error parsing operand token
                                throw new ApplicationException(String.Format("Unable to parse parameter '{0}' in function token '{1}'.", parameterToken, token));
                            }
                            parameterDictionary.Add(String.Format("Value{0}", parameterIndex++), parameterOperand);
                        }
                        
                        //set operand to new OperandOperation of type OperationFunction using operator of type OperatorXXX (foundOperator)
                        operand = new OperandOperation(new OperationFunction(foundOperator, parameterDictionary));
                    }
                    else
                    { 
                        throw new ApplicationException(String.Format("Unable to find function: '{0}'", token));
                    }
                }
                else
                { 
                    throw new ApplicationException(String.Format("Unable to find braces for parameters: '{0}'", token));
                }

                returnValue = true;
            }
            catch (Exception ex)
            {
                Log.Write(ex, MethodBase.GetCurrentMethod(), EventLogEntryType.Error);
                //throw;
            }
            return returnValue;
        }
        #endregion Private Methods
        #endregion Methods
    }
}
