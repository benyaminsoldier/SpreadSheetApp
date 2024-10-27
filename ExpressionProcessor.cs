using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace spreadsheetApp
{
    abstract class Operator
    {
        public char Symbol { get; set; }
        public int Precedence { get; set; }
        public bool IsLeftAssociative { get; set; }
        //Since we are using just left associative operators this property is useless.
        public abstract float Execute(float op1 = 0, float op2 = 0);
        //method overloading cause sometimes I have chars and sometime I have strings.
        public static bool IsOperator(char op)
        {
            if (op == '+' || op == '-' || op == '*' || op == '/')
                return true;

            return false;
        }
        public static bool IsOperator(string op)
        {
            if (op == "+" || op == "-" || op == "*" || op == "/")
                return true;

            return false;
        }

    }
    class Addition : Operator
    {
        public Addition()
        {
            Symbol = '+';
            Precedence = 1;
            IsLeftAssociative = true;
        }

        public override float Execute(float op1, float op2)
        {
            return op1 + op2;
        }
    }
    class Subtraction : Operator
    {
        public Subtraction()
        {
            Symbol = '-';
            Precedence = 1;
            IsLeftAssociative = true;
        }

        public override float Execute(float op1, float op2)
        {
            return op1 - op2;
        }
    }
    class Multiplication : Operator
    {
        public Multiplication()
        {
            Symbol = '*';
            Precedence = 2;
            IsLeftAssociative = true;
        }

        public override float Execute(float op1, float op2)
        {
            return op1 * op2;
        }
    }
    class Division : Operator
    {
        public Division()
        {
            Symbol = '/';
            Precedence = 2;
            IsLeftAssociative = true;
        }

        public override float Execute(float op1, float op2)
        {
            if (op2 == 0) throw new DivideByZeroException(); //Where to catch it???
            return op1 / op2;
        }
    }
    //
    //Left_Parentheses was included as a class since it needs to go into the Stack<Operator> and it cannot go in as a mortal string XD
    //
    class Left_Parentheses : Operator
    {
        public Left_Parentheses()
        {
            Symbol = '(';
            Precedence = 3;
            IsLeftAssociative = true;
        }

        public override float Execute(float op1, float op2)
        {
            throw new NotImplementedException("Parentheses really do not perform any operation");
        }
    }
    static class ExpressionProcessor
    {

        // PROPS 
        //
        //
        private static List<string> OutputList { get; set; } = new List<string>();
        //This list receives operands and operators that compose the postfix exp returned by the ToPostfix() execution.
        private static Stack<Operator> OperatorStack { get; set; } = new Stack<Operator>();
        //Operators found in the infix exp are stacked and pop out to the output list depending on their precedence and association.
        //Stacks are like dumb lists with constraints.... FILO (First In, Last Out)
        //
        //
        //infix expression coming from the user input in the program class/ProcessCommand method.
        //This method separates each number, operator and parentheses into individual strings.
        //must be strings for double digits or float numbers.
        //Negative reckon logic goes here.
        //Validation could go here...
        //Returns the infix exp but tokenized.

        // PRIVATE METHODS
        //
        //
        private static bool _validateParentheses(char[] expArr)
        {
            bool order = true;

            for (int j = 0; j < expArr.Length; j++)
            {
                if (expArr[0] == '(') continue;
                if (expArr[j] == '(' && !Operator.IsOperator(expArr[j - 1])) order = false;

            }

            for (int j = 0; j < expArr.Length; j++)
            {
                if (expArr[expArr.Length - 1] == ')') continue;
                if (expArr[j] == ')' && !Operator.IsOperator(expArr[j + 1])) order = false;

            }


            int lp = 0, rp = 0;
            foreach (var num in expArr)
            {
                if (num == '(') lp++;
                else if (num == ')') rp++;
            }

            if (lp - rp == 0 && order) return true;

            return false;
        }
        
        private static bool _validateOperators(char[] expArr)
        {
            foreach (var num in expArr)
            {
                if (Operator.IsOperator(num) || char.IsDigit(num) || num == '(' || num == ')')
                {
                    continue;
                }
                else  return false ;
                
                
            }
            return true;

        }
        private static char _getInvalidOperators(char[] expArr)
        {
            char invalidChar='a';
            foreach (var num in expArr)
            {
                if (Operator.IsOperator(num) || char.IsDigit(num) || num == '(' || num == ')')
                {
                    continue;
                }
                else invalidChar = num;


            }
            return invalidChar;

        }
        private static bool _validateExpOrder(List<string> tokenizedExp)
        {
            for (int j = 0; j < tokenizedExp.Count; j++)
                if (Operator.IsOperator(tokenizedExp[j]))
                    if (Operator.IsOperator(tokenizedExp[j + 1]))
                        return false;

            return true;
        }
        private static bool _validateDecimals(List<string> tokenizedExp)
        {
            bool state = true;
            tokenizedExp.ForEach((num) =>
            {
                if (!float.TryParse(num, out float num1))
                {
                    if (!(Operator.IsOperator(num) || num == "(" || num == ")"))
                    {
                        state = false;
                    }
                }

            });
            return state;
        }
        private static Stack<string> _convertListToStack(List<string> list)
        {
            Stack<string> stack = new Stack<string>();
            for (int j = 0; j < list.Count; j++) stack.Push(list[j]);
            return stack;
        }
        

        // PUBLIC METHODS
        //
        //
        private static List<string> Tokenize(string input)
        {
            //
            //Spaces removal         
            while (input.Contains(" "))
            {
                input = input.Replace(" ", ""); // finds first occurrence and replace it until got no more ocurrences.
            }
            //Now input is bulletproof...
            List<string> infixExp = new List<string>();   /// This list will be returned with the tokenized expression.               
            char[] expArr = input.ToCharArray();
            string substrng = "";
            int numero = 0;
            if (!_validateParentheses(expArr)) throw new Exception("Please close parentheses properly");
            if (!_validateOperators(expArr)) throw new InvalidFormulaException("This calculator only add,substracts, multiplies and divides",_getInvalidOperators(expArr));
           

            for (int i = 0; i < expArr.Length;)
            {
                //Validation implemented for string inputs.
                //Should I include a try/catch block in here?? or wait for Pedro's catcher?
                //
                if (char.IsLetter(expArr[i])) throw new Exception("This calculator does not accept variables/characters");
                //
                //
                //Reckons implicit decimals at the start
                if (i == 0 && expArr[i] == '.')
                {
                    int skippedchars = 1; //The decimal counts... so starts at 1.
                    for (int j = 0; j < input.Substring(i + 1).Length; j++)
                    {
                        if (char.IsDigit(input.Substring(i + 1)[j]))
                        {
                            substrng += input.Substring(i + 1)[j];
                            skippedchars++;
                            continue;
                        }
                        else break; //It ends when something that is not a number is detected.

                    }
                    infixExp.Add($"0{expArr[i]}{substrng}"); //substrng collects the whole number and adds it to the negative sign.
                    i += skippedchars; //Next look up will begin in the next operator/parentheses.
                    substrng = "";
                    continue;
                }
                //
                //Reckons implicit decimals:
                if (i > 0 && Operator.IsOperator(expArr[i - 1]) && expArr[i] == '.')
                {
                    int skippedchars = 1; //The decimal counts... so starts at 1.
                    for (int j = 0; j < input.Substring(i + 1).Length; j++)
                    {
                        if (char.IsDigit(input.Substring(i + 1)[j]))
                        {

                            substrng += input.Substring(i + 1)[j];
                            skippedchars++;
                            continue;
                        }
                        else break; //It ends when something that is not a number is detected.

                    }

                    infixExp.Add($"0{expArr[i]}{substrng}"); //substrng collects the whole number and adds it to the negative sign.
                    i += skippedchars; //Next look up will begin in the next operator/parentheses.
                    substrng = "";
                    continue;
                }
                //
                //Reckons negatives at the start:
                //
                if (i == 0 && expArr[i] == '-')
                {
                    //Looking for the complete negative number.
                    int skippedchars = 1; //The minus counts... so starts at 1.
                    for (int j = 0; j < input.Substring(i + 1).Length; j++)
                    {
                        if (char.IsDigit(input.Substring(i + 1)[j]) || input.Substring(i + 1)[j] == '.')
                        {
                            //input.substring(i+1) represents a substring taken from 1 char offset to the right  from where the negative showed up.
                            if (input.Substring(i + 1)[j] == '.')
                            {
                                substrng += $"0{input.Substring(i + 1)[j]}";
                                skippedchars++;
                                continue;
                            }
                            substrng += input.Substring(i + 1)[j];
                            skippedchars++;
                            continue;
                        }
                        else break; //It ends when something that is not a number is detected.

                    }

                    infixExp.Add($"{expArr[i]}{substrng}"); //substrng collects the whole number and adds it to the negative sign.
                    i += skippedchars; //Next look up will begin in the next operator/parentheses.
                    substrng = "";
                    continue;
                }
                //
                //Reckons negatives after start: bool expression crashes if negative is the first char.
                //
                if (i > 0 && expArr[i] == '-' && !char.IsDigit(expArr[i - 1]))
                {
                    if (expArr[i - 1] == ')')
                    {
                        infixExp.Add(expArr[i] + "");
                        i++;
                        continue;
                    }
                    //
                    //Looking for the complete negative number.
                    //
                    int skippedchars = 1; //The minus counts... so starts at 1.
                    for (int j = 0; j < input.Substring(i + 1).Length; j++)
                    {
                        if (char.IsDigit(input.Substring(i + 1)[j]) || input.Substring(i + 1)[j] == '.')
                        {
                            //input.substring(i+1) represents a substring taken from 1 char offset to the right  from where the negative showed up.
                            if (input.Substring(i + 1)[j] == '.' && input.Substring(i)[j] == '-')
                            {
                                substrng += $"0{input.Substring(i + 1)[j]}";
                                skippedchars++;
                                continue;
                            }
                            substrng += input.Substring(i + 1)[j];
                            skippedchars++;
                            continue;
                        }
                        else break; //It ends when something that is not a number is detected.

                    }

                    infixExp.Add($"{expArr[i]}{substrng}"); //substrng collects the whole number and adds it to the negative sign.
                    i += skippedchars; //Next look up will begin in the next operator/parentheses.
                    substrng = "";
                    continue;
                }
                //
                //Reckons operators n parentheses:
                //
                if (Operator.IsOperator(expArr[i]) || (expArr[i] == '-' && char.IsDigit(expArr[i - 1])) || expArr[i] == '(' || expArr[i] == ')' && i > 0)
                {
                    infixExp.Add(expArr[i] + "");
                    i++;
                    continue;
                }
                //
                //Checking for full numbers not just digits.
                //
                if (char.IsDigit(expArr[i]))
                {

                    int skippedchars = 0; //No minus here... XD
                    for (int j = 0; j < input.Substring(i).Length; j++)
                    {
                        if (char.IsDigit(input.Substring(i)[j]) || input.Substring(i)[j] == '.')
                        {
                            substrng += input.Substring(i)[j];
                            skippedchars++;
                        }
                        else break;
                    }
                    infixExp.Add(substrng);
                    i += skippedchars;
                    substrng = "";
                    continue;
                }



            } //Main For Loop ends here.

            //Final validations that were accomplished much easier with defined numbers and operators. 
            //Some strings validations might be missing regards to algebra operations.
            if (!_validateExpOrder(infixExp)) throw new Exception("Please do not type one operator after another one.");
            if (!_validateDecimals(infixExp)) throw new Exception("This calculator is not currently accepting IPs XD");


            return infixExp;
        }
        //
        //ToPostfix method converts the tockenized infix expression to a postfix expression via the Shunting-yard algorithm
        //Parentheses logic is placed here overriding '*'and '/' precedence.
        //Operands will go into the outputlist.
        //Operators with higher precedence will go on the top stack.
        //Operators with lower precedence will cause stacked operators pop out into the output list in the stack order (higher order). 
        //Current operator will be stacker until a lower is found or the expression comes to its end.
        //Returns postfixExp ready to be executed by ExecutePostfix() method.



        private static Stack<string> ToPostfix(List<string> infixExp)
        {

            Operator addition = new Addition();
            Operator subtraction = new Subtraction();
            Operator multiplication = new Multiplication();
            Operator division = new Division();
            Operator left_parentheses = new Left_Parentheses();

            void StackByPrecedence(string op)
            {
                Operator operador = null;

                if (op == "+") operador = addition;
                else if (op == "-") operador = subtraction;
                else if (op == "*") operador = multiplication;
                else if (op == "/") operador = division;

                if (OperatorStack.Count == 0 || OperatorStack.Peek() == left_parentheses)
                {
                    OperatorStack.Push(operador);
                }
                else
                {
                    if (operador.Precedence > OperatorStack.Peek().Precedence)
                        OperatorStack.Push(operador);

                    else if (operador.Precedence <= OperatorStack.Peek().Precedence)
                    {
                        while (OperatorStack.Count > 0 && OperatorStack.Peek() != left_parentheses && OperatorStack.Peek().Precedence >= operador.Precedence)
                            OutputList.Add(OperatorStack.Pop().Symbol + "");

                        OperatorStack.Push(operador);
                    }
                }
            }

            infixExp.ForEach((num) =>
            {

                if (float.TryParse(num, out float numero))
                    OutputList.Add(numero.ToString());

                else if (num == "(")
                    OperatorStack.Push(left_parentheses);

                else if (Operator.IsOperator(num))
                    StackByPrecedence(num);

                else if (num == ")")
                {
                    string op;
                    while ((op = OperatorStack.Pop().Symbol + "") != "(")
                        OutputList.Add(op);
                }

            });

            while (OperatorStack.Count > 0)
                OutputList.Add(OperatorStack.Pop().Symbol + "");


            Stack<string> postfixStack = _convertListToStack(OutputList);
            OutputList.Clear();
            return postfixStack;
        }
        //
        //receives postfix exp and feeds stack while the stack operates and stacks the result.
        //some private functions to convert strings to floats and vis could be used here.

        private static bool _detectedOperation(List<string> list, ref int index) // 1 2 + 3 * -->  * 3 + 2 1  --> * 3
        {
            bool detected = false;
            for (int j = 0; j < list.Count - 2; j++)
            {
                if (Operator.IsOperator(list[j]) && float.TryParse(list[j + 1], out float num1) && float.TryParse(list[j + 2], out float num2))
                {
                    index = j;
                    detected = true;
                    break;
                }
                continue;
            }


            return detected;

        }
        private static float _executeOperator(List<string> list, int index)
        {

            Operator op = null;
            float partialResult = 0;
            /*
            void _displayList(List<string> list)
            {
                list.ForEach((item) =>
                {
                    Console.WriteLine(item);
                });
            }*/

            if (list[index] == "+")
            {
                op = new Addition();
                partialResult = op.Execute(float.Parse(list[index + 2]), float.Parse(list[index + 1]));
                list.RemoveRange(index, 3);
                list.Add(partialResult.ToString());
                //Console.WriteLine()
            }
            else if (list[index] == "-")
            {
                op = new Subtraction();
                partialResult = op.Execute(float.Parse(list[index + 2]), float.Parse(list[index + 1]));
                list.RemoveRange(index, 3);
                list.Add(partialResult.ToString());
            }
            else if (list[index] == "*")
            {
                op = new Multiplication();
                partialResult = op.Execute(float.Parse(list[index + 2]), float.Parse(list[index + 1]));
                list.RemoveRange(index, 3);
                list.Add(partialResult.ToString());
            }
            else if (list[index] == "/")
            {
                op = new Division();
                partialResult = op.Execute(float.Parse(list[index + 2]), float.Parse(list[index + 1]));
                list.RemoveRange(index, 3);
                list.Add(partialResult.ToString());
            }



            return partialResult;
        }

        private static string ExecutePostfix(Stack<string> postfixStack)
        {

            List<string> tempList = new List<string>(); // 
            int opIndex = 0;
            bool opDetected = false;
            //1+2 -> 1 2 + ->  + 2 1
            while (postfixStack.Count + tempList.Count > 1)
            {
                if (postfixStack.Count > 0)
                    tempList.Add(postfixStack.Pop());

                opDetected = _detectedOperation(tempList, ref opIndex);
                if (tempList.Count > 2 && opDetected)
                {

                    _executeOperator(tempList, opIndex);
                    //le.WriteLine();
                    for (int j = tempList.Count - 1; j >= 0; j--)
                    {
                        postfixStack.Push(tempList[j]);
                        //Console.WriteLine(postfixStack.Peek());

                    }
                    tempList.Clear();
                }

            }
            return postfixStack.Pop();
        }
        public static string ProcessCommand(string input)
        {
            try
            {
                var list = ExpressionProcessor.Tokenize(input);
                var postfixOutput = ExpressionProcessor.ToPostfix(list);  
                string result = $"{ExpressionProcessor.ExecutePostfix(postfixOutput)}";
                return result;
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
        }
    }
}
