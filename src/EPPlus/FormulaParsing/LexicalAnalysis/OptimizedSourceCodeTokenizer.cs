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
           isNegator = 0x80,
           isColon =   0x100
        }
        public IEnumerable<Token> Tokenize(string input, string worksheet)
        {
            var l = new List<Token>();
            int ix;
            var length = input.Length;
            
            if (length > 0 && (input[0] == '+' /*|| input[0] == '='*/))
            {
                ix=1;
            }
            else
            {
                ix = 0;
            }

            statFlags flags = 0;



            short isInString = 0;
            short bracketCount = 0;
            var current =new StringBuilder();
            var pc = '\0';
            var separatorTokens = TokenSeparatorProvider.Instance.Tokens;
            while (ix < length)
            {
                var c = input[ix];
                if(c == '\"' && isInString != 2)
                {
                    if (pc == c && isInString == 0)
                    {
                        current.Append(c);
                    }
                    else
                    {
                        flags |= statFlags.isString;
                    }
                    isInString ^= 1;
                }
                else if (c == '\'' && isInString !=1)
                {
                    current.Append(c);
                    flags |= statFlags.isAddress;
                    isInString ^= 2;
                }
                else if(c=='[' && isInString == 0)
                {
                    current.Append(c);
                    flags |= statFlags.isAddress;
                    bracketCount++;
                }
                else if (c == ']' && isInString == 0)
                {
                    current.Append(c);
                    bracketCount--;
                }
                else
                { 
                    if(isInString==0 && bracketCount == 0 && _charTokens.ContainsKey(c))
                    {
                        if (c == ' ') //white space, we ignore for now. Implement intersect operation handling here.
                        {

                        }
                        else
                        {
                            HandleToken(l, c, current, ref flags);
                            if (c == '-')
                            {
                                flags |= statFlags.isNegator;
                            }
                            else if(c=='+' && l.Count>0) //remove leading + and add + operator.
                            {
                                var pt = l[l.Count - 1];

                                //Remove prefixing +
                                if (!(pt.TokenTypeIsSet(TokenType.Operator)
                                    ||
                                    pt.TokenTypeIsSet(TokenType.Negator)
                                    ||
                                    pt.TokenTypeIsSet(TokenType.OpeningParenthesis)
                                    ||
                                    pt.TokenTypeIsSet(TokenType.Comma)
                                    ||
                                    pt.TokenTypeIsSet(TokenType.SemiColon)
                                    ||
                                    pt.TokenTypeIsSet(TokenType.OpeningEnumerable)))
                                {
                                    l.Add(_charTokens[c]);
                                }
                            }
                            else if (ix + 1 < length && _stringTokens.ContainsKey(input.Substring(ix, 2)))
                            {
                                l.Add(_stringTokens[input.Substring(ix, 2)]);
                                ix++;
                            }
                            else
                            {
                                l.Add(_charTokens[c]);
                            }
                        }
                    }
                    else
                    {
                        if(isInString==0)
                        {
                            if (current.Length == 0 && c == ':' && pc == ')')
                            {
                                l.Add(new Token(":", TokenType.Colon));
                                SetRangeOffsetToken(l);
                                flags |= statFlags.isColon;
                            }
                            else
                            {
                                if (_charAddressTokens.ContainsKey(c)) //handel :
                                {
                                    flags |= statFlags.isAddress;
                                }
                                else if (c >= '0' && c <= '9')
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
                                current.Append(c);
                            }
                        }
                        else
                        {
                            current.Append(c);
                        }
                    }
                }
                ix++;
                pc = c;
            }
            HandleToken(l, pc, current, ref flags);
            return l;
        }

        private void SetRangeOffsetToken(List<Token> l)
        {
            int i = l.Count - 1;
            int p= 0;
            while (i >= 0)
            {
                if (l[i].TokenTypeIsSet(TokenType.OpeningParenthesis))
                {
                    p--;
                }
                else if(l[i].TokenTypeIsSet(TokenType.ClosingParenthesis))
                {
                    p++;
                }
                else if (l[i].TokenTypeIsSet(TokenType.Function) && l[i].Value.Equals("offset", StringComparison.OrdinalIgnoreCase) && p==0)
                {
                    l[i] = new Token(l[i].Value, TokenType.RangeOffset | TokenType.Function);
                }
                i--;
            }
        }

        private void HandleToken(List<Token> l,char c, StringBuilder current, ref statFlags flags)
        {
            if ((flags & statFlags.isNegator) == statFlags.isNegator)
            {
                if (l.Count == 0)
                {
                    l.Add(new Token("-", TokenType.Negator));
                }
                else
                {
                    var pt = l[l.Count - 1];
                    if (pt.TokenTypeIsSet(TokenType.Operator)
                        ||
                        pt.TokenTypeIsSet(TokenType.Negator)
                        ||
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
            }
            if (current.Length == 0)
            {
                if((flags & statFlags.isString) == statFlags.isString)
                {
                    l.Add(new Token("", TokenType.StringContent));
                }
                flags = 0;
                return;
            }
            var currentString = current.ToString();
            if ((flags & statFlags.isString) == statFlags.isString)
            {
                l.Add(new Token(currentString, TokenType.StringContent));
            }
            else if (c == '(')
            {
                if((flags & statFlags.isColon) == statFlags.isColon)
                {
                    l.Add(new Token(currentString, TokenType.Function | TokenType.RangeOffset));
                }
                else
                {
                    l.Add(new Token(currentString, TokenType.Function));
                }
            }            
            else if ((flags & statFlags.isAddress) == statFlags.isAddress)
            {
                if (currentString.EndsWith("#REF!", StringComparison.OrdinalIgnoreCase))
                {
                    l.Add(new Token(currentString, TokenType.InvalidReference));
                }
                else if (currentString.EndsWith("#NUM!", StringComparison.OrdinalIgnoreCase))
                {
                    l.Add(new Token(currentString, TokenType.NumericError));
                }
                else if (currentString.EndsWith("#VALUE!", StringComparison.OrdinalIgnoreCase))
                {
                    l.Add(new Token(currentString, TokenType.ValueDataTypeError));
                }
                else if (currentString.EndsWith("#NULL!", StringComparison.OrdinalIgnoreCase))
                {
                    l.Add(new Token(currentString, TokenType.Null));
                }
                else
                {
                    if(IsName(currentString))
                    {
                        l.Add(new Token(currentString, TokenType.NameValue));
                    }
                    else
                    {
                        l.Add(new Token(currentString, TokenType.ExcelAddress));
                    }
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
                    if ((flags & statFlags.isColon) == statFlags.isColon)
                    {
                        l.Add(new Token(currentString, TokenType.ExcelAddress| TokenType.RangeOffset));
                    }
                    else
                    {
                        l.Add(new Token(currentString, TokenType.ExcelAddress));
                    }
                }
                else
                {
                    l.Add(new Token(currentString, TokenType.NameValue));
                }
            }
            else
            {
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
    private static readonly char[] _addressChars = new char[]{':','$', '[', ']'};
    private static bool IsName(string s)
    {
        var ix = s.LastIndexOf('!');
        if(ix>=0)
        {
            s = s.Substring(ix + 1);
        }        
        if (s.IndexOfAny(_addressChars) >=0) return false;
        return IsValidCellAddress(s)==false;
    }

        private static bool IsValidCellAddress(string address)
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
            var col = ExcelAddressBase.GetColumnNumber(address.Substring(0,numPos));
            if (col <= 0 || col > ExcelPackage.MaxColumns) return false;
            if(int.TryParse(address.Substring(numPos), out int row))
            {
                return row > 0 && row <= ExcelPackage.MaxRows; 

            }
            return false;
        }
    }
}
