using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis
{
    public class OptimizedSourceCodeTokenizer : ISourceCodeTokenizer
    {
        private static readonly Dictionary<char, Token> _charAddressTokens = new Dictionary<char, Token>
        {
            {'!', new Token("!", TokenType.ExcelAddress)},
            {'$', new Token("$", TokenType.ExcelAddress)},
            {'[', new Token("[", TokenType.OpeningBracket)},
            {']', new Token("]", TokenType.ClosingBracket)},
            {':', new Token(":", TokenType.Colon) },
        };
        private static readonly Dictionary<char, Token> _charTokens = new Dictionary<char, Token>
        {
            {'+', new Token("+",TokenType.Operator)},
            {'-', new Token("-", TokenType.Operator)},
            {'*', new Token("*", TokenType.Operator)},
            {'/', new Token("/", TokenType.Operator)},
            {'^', new Token("^", TokenType.Operator)},
            {'&', new Token("&", TokenType.Operator)},
            {'>', new Token(">", TokenType.Operator)},
            {'<', new Token("<", TokenType.Operator)},
            {'=', new Token("=", TokenType.Operator)},
            {'(', new Token("(", TokenType.OpeningParenthesis)},
            {')', new Token(")", TokenType.ClosingParenthesis)},
            {'{', new Token("{", TokenType.OpeningEnumerable)},
            {'}', new Token("}", TokenType.ClosingEnumerable)},
            {'\"', new Token("\"", TokenType.String)},
            {',', new Token(",", TokenType.Comma)},
            {';', new Token(";", TokenType.SemiColon) },
            {'%', new Token("%", TokenType.Percent) },
            {' ', new Token(" ", TokenType.WhiteSpace) }
        };
        private static readonly Dictionary<string, Token> _stringTokens = new Dictionary<string, Token>
        {
            {">=", new Token(">=", TokenType.Operator)},
            {"<=", new Token("<=", TokenType.Operator)},
            {"<>", new Token("<>", TokenType.Operator)},
        };


        public static ISourceCodeTokenizer Default
        {
            get { return new OptimizedSourceCodeTokenizer(FunctionNameProvider.Empty, NameValueProvider.Empty, false); }
        }
        public static ISourceCodeTokenizer R1C1
        {
            get { return new OptimizedSourceCodeTokenizer(FunctionNameProvider.Empty, NameValueProvider.Empty, true); }
        }

        public OptimizedSourceCodeTokenizer(IFunctionNameProvider functionRepository, INameValueProvider nameValueProvider, bool r1c1 = false)
            : this(new TokenFactory(functionRepository, nameValueProvider, r1c1))
        {
        }
        public OptimizedSourceCodeTokenizer(ITokenFactory tokenFactory)
        {
            _tokenFactory = tokenFactory;
        }

        private readonly ITokenFactory _tokenFactory;

        public IEnumerable<Token> Tokenize(string input)
        {
            return Tokenize(input, null);
        }
        [Flags]
        enum statFlags : int
        {
           isString =   0x1,
           isOperator = 0x2,
           isAddress =  0x4,
           isNonNumeric=0x8,
           isNumeric = 0x10,
           isDecimal = 0x20,
           isPercent = 0x40,
           isNegator = 0x80
        }
        public IEnumerable<Token> Tokenize(string input, string worksheet)
        {
            var l = new List<Token>();
            int ix = 0;

            statFlags flags = 0;
            var isInString = false;
            var current =new StringBuilder();
            var length = input.Length;
            var pc = '\0';
            var separatorTokens = TokenSeparatorProvider.Instance.Tokens;
            while (ix < length)
            {
                var c = input[ix];
                if (ix==0 && (c=='+' || c=='='))
                {
                    ix++;
                }

                if (c == '\"')
                {
                    if (pc == c && isInString)
                    {
                        current.Append(c);
                    }
                    else
                    {
                        flags |= statFlags.isString;
                    }
                    isInString = !isInString;
                }
                else
                { 
                    if(isInString==false && _charTokens.ContainsKey(c))
                    {
                        HandleToken(l, c, current, ref flags);
                        if(c=='-')
                        {
                            flags |= statFlags.isNegator;
                        }
                        else if (ix+1 < length && _stringTokens.ContainsKey(input.Substring(ix,2)))
                        {
                            l.Add(_stringTokens[input.Substring(ix, 2)]);
                        }
                        else
                        {
                            l.Add(_charTokens[c]);
                        }
                    }
                    else
                    {
                        if(!isInString)
                        {
                            if (_charAddressTokens.ContainsKey(c))
                            {
                                flags |= statFlags.isAddress;
                            }
                            if (c >= '0' && c <= '9')
                            {
                                flags |= statFlags.isNumeric;
                            }
                            else if (c == '.')
                            {
                                flags |= statFlags.isDecimal;
                            }
                            else if (c == '%')
                            {
                                flags |= statFlags.isPercent;
                            }
                            else
                            {
                                flags |= statFlags.isNonNumeric;
                            }
                        }
                        current.Append(c);
                    }
                }
                ix++;
                pc = c;
            }
            HandleToken(l, pc, current, ref flags);
            return l;
        }
        private void HandleToken(List<Token> l,char c, StringBuilder current, ref statFlags flags)
        {
            if ((flags & statFlags.isNegator) == statFlags.isNegator)
            {
                var pt = l[l.Count - 1];
                if (pt.TokenTypeIsSet(TokenType.Operator) ||
                        pt.TokenTypeIsSet(TokenType.OpeningParenthesis)
                        ||
                        pt.TokenTypeIsSet(TokenType.Comma)
                        ||
                        pt.TokenTypeIsSet(TokenType.SemiColon)
                        ||
                        pt.TokenTypeIsSet(TokenType.OpeningEnumerable))
                {
                    l.Add(new Token("-", TokenType.Negator));
                }
                else
                {
                    l.Add(_charTokens['-']);
                }
            }
            if (current.Length == 0)
            {
                return;
            }
            var currentString = current.ToString();
            if ((flags & statFlags.isString) == statFlags.isString)
            {
                l.Add(new Token(currentString, TokenType.StringContent));
            }
            else if (c == '(')
            {
                l.Add(new Token(currentString, TokenType.Function));
            }
            else if ((flags & statFlags.isAddress) == statFlags.isAddress)
            {
                if (currentString.Equals("#REF!", StringComparison.OrdinalIgnoreCase))
                {
                    l.Add(new Token(currentString, TokenType.InvalidReference));
                }
                else if (currentString.Equals("#NUM!", StringComparison.OrdinalIgnoreCase))
                {
                    l.Add(new Token(currentString, TokenType.NumericError));
                }
                else if (currentString.Equals("#VALUE!", StringComparison.OrdinalIgnoreCase))
                {
                    l.Add(new Token(currentString, TokenType.ValueDataTypeError));
                }
                else if (currentString.Equals("#NULL!", StringComparison.OrdinalIgnoreCase))
                {
                    l.Add(new Token(currentString, TokenType.Null));
                }
                else
                {
                    l.Add(new Token(currentString, TokenType.ExcelAddress));
                }
            }
            else if ((flags & statFlags.isNonNumeric) == statFlags.isNonNumeric)
            {
                if (currentString.Equals("true", StringComparison.OrdinalIgnoreCase) ||
                   currentString.Equals("false", StringComparison.OrdinalIgnoreCase))
                {
                    l.Add(new Token(currentString, TokenType.Boolean));
                }
                else if (IsValidCellAddress(currentString))
                {
                    l.Add(new Token(currentString, TokenType.ExcelAddress));
                }
                else
                {
                    l.Add(new Token(currentString, TokenType.NameValue));
                }
            }
            else
            {
                //We can set the value as negative instead of an extra token for negator.
                //if(EnumUtil.HasFlag(flags, statFlags.isNegator))
                //{
                //    currentString = "-" + currentString;
                //}
                if ((flags & statFlags.isPercent) == statFlags.isPercent)
                {
                    l.Add(new Token(currentString, TokenType.Percent));
                }
                else if ((flags & statFlags.isDecimal) == statFlags.isDecimal)
                {
                    l.Add(new Token(currentString, TokenType.Decimal));
                }
                else if ((flags & statFlags.isNumeric) == statFlags.isNumeric)
                {
                    l.Add(new Token(currentString, TokenType.Integer));
                }
            }
            flags = 0;
            //Clear sb
#if (NET35)
    current=new StringBuilder();
#else
            current.Clear();
#endif
            
        }
        private bool IsValidCellAddress(string address)
        {
            var numPos = -1;
            for (var i=0; i < address.Length; i++)
            {
                var c = address[i];
                if (c>='0' && c<='9')
                {
                    if (i == 0) return false;
                    numPos = i;
                }
                else if ((c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z'))
                {
                    if (numPos != -1 || i > 3) return false;
                }
                else
                {
                    return false;
                }
            }
            if (numPos < 1 || numPos > 3) return false;
            var col = ExcelAddressBase.GetColumnNumber(address.Substring(numPos));
            if (col <= 0 || col > ExcelPackage.MaxColumns) return false;
            if(int.TryParse(address.Substring(numPos), out int row))
            {
                return row > 0 && row <= ExcelPackage.MaxRows; 

            }
            return false;
        }
    }
}
