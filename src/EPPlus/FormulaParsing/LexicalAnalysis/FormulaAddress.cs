using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;

namespace OfficeOpenXml.FormulaParsing.LexicalAnalysis
{
    internal class Formula
    {
        ExcelWorksheet _ws;
        ExcelRangeBase _range;
        internal static ISourceCodeTokenizer _tokenizer= OptimizedSourceCodeTokenizer.Default;
        public Formula(ExcelRangeBase range, string formula)
        {            
            _range = range;
            _ws = range.Worksheet;
            Tokens = _tokenizer.Tokenize(formula);
            SetTokenInfos();
        
        }
        internal IList<Token> Tokens;
        internal Dictionary<int, TokenInfo> TokenInfos;

        private void SetTokenInfos()
        {
            TokenInfos = new Dictionary<int, TokenInfo>();
            string er = "", ws = "";
            short startToken = -1;
            for (short i = 0; i < Tokens.Count; i++)
            {
                var t = Tokens[i];
                switch(t.TokenType)
                {

                    case TokenType.ExcelAddress:
                        var fa = new FormulaCellAddress(i, t.Value);
                        TokenInfos.Add(i, fa);
                        er = ws = "";
                        break;
                    case TokenType.NameValue:
                        AddNameInfo(startToken==-1 ? i : startToken, i, er, ws);
                        er = ws = "";
                        break;
                    case TokenType.TableName:
                        AddTableAddress(i);
                        er = ws = "";
                        break;
                    case TokenType.WorksheetNameContent:
                        if (startToken == -1)
                        {
                            startToken = i;
                        }
                        ws = t.Value;
                        break;
                    case TokenType.ExternalReference:
                        er = t.Value;
                        break;
                    case TokenType.OpeningBracket:
                        if (startToken == -1)
                        {
                            startToken = i;
                        }
                        break;
                }
            }
        }

        int _rowOffset=0, _colOffset = 0;
        internal void SetOffset(int rowOffset, int colOffset)
        {
            var changeRowOffset = rowOffset - _rowOffset;
            var changeColOffset = colOffset - _colOffset;
            foreach (var t in TokenInfos.Values)
            {
                switch(t.Type)
                {
                    case FormulaType.CellAddress:
                    case FormulaType.FormulaRange:
                        t.SetOffset(changeRowOffset, changeColOffset);
                        break;
                }
            }
            _rowOffset = rowOffset;
            _colOffset = colOffset;
        }

        private void AddTableAddress(short pos)
        {
            short i = pos;
            var t = Tokens[i];
            TokenInfos = new Dictionary<int, TokenInfo>();
            var table = _range.Worksheet.Workbook.GetTable(t.Value);
            if(table != null)
            {
                if(Tokens[++i].TokenType == TokenType.OpeningBracket)
                {
                    int fromRow = 0, toRow = 0, fromCol = 0, toCol = 0;
                    FixedFlag fixedFlag = FixedFlag.All;
                    bool lastColon=false;
                    var bc = 1;
                    i++;
                    while(bc>0 && i < Tokens.Count)
                    {
                        switch(Tokens[i].TokenType)
                        {
                            case TokenType.OpeningBracket:
                                bc++;
                                break;
                            case TokenType.ClosingBracket:
                                bc--;
                                break;
                            case TokenType.TablePart:
                                SetRowFromTablePart(Tokens[i].Value, table, ref fromRow, ref toRow, ref fixedFlag);
                                break;
                            case TokenType.TableColumn:
                                SetColFromTablePart(Tokens[i].Value, table, ref fromCol, ref toCol, lastColon);
                                break;
                            case TokenType.Colon:
                                lastColon=true;
                                break;
                            default:
                                lastColon = false;
                                break;
                        }
                        i++;
                    }
                    if(bc==0)
                    {
                        if(fromRow == 0)
                        {
                            fromRow = table.DataRange._fromRow;
                            toRow = table.DataRange._toRow;
                        }
                        
                        if(fromCol == 0)
                        {
                            fromCol = table.DataRange._fromCol;
                            toCol = table.DataRange._toCol;
                        }

                        i--;
                        TokenInfos.Add(pos, new FormulaRange(pos, i, fromRow, fromCol, toRow, toCol, fixedFlag)); 
                    }
                }
                else
                {
                    TokenInfos.Add(pos, new FormulaRange(pos, i, table.DataRange._fromRow, table.DataRange._fromCol, table.DataRange._toRow, table.DataRange._toCol, 0));
                }
            }
        }
        private void SetColFromTablePart(string value, ExcelTable table, ref int fromCol, ref int toCol, bool lastColon)
        {
            var col = table.Columns[value];
            if (col == null) return;
            if (lastColon)
            {
                toCol = table.Range._fromCol + col.Position;
            }
            else
            {
                fromCol = toCol = table.Range._fromCol + col.Position;
            }
        }
        private void SetRowFromTablePart(string value, ExcelTable table, ref int fromRow, ref int toRow, ref FixedFlag fixedFlag)
        {
            switch(value.ToLower())
            {
                case "#all":
                    fromRow = table.Address._fromRow;
                    toRow = table.Address._toRow;
                    break;
                case "#headers":
                    if (table.ShowHeader)
                    {
                        fromRow = table.Address._fromRow;
                        if (toRow == 0)
                        {
                            toRow = table.Address._fromRow;
                        }
                    }
                    else if (fromRow == 0)
                    {
                        fromRow = toRow = -1;
                    }
                    break;
                case "#data":
                    if (fromRow == 0 || table.DataRange._fromRow < fromRow)
                    {
                        fromRow = table.DataRange._fromRow;
                    }
                    if (table.DataRange._toRow > toRow)
                    {
                        toRow = table.DataRange._toRow;
                    }
                    break;
                case "#totals":
                    if (table.ShowTotal)
                    {
                        if (fromRow == 0)
                            fromRow = table.Range._toRow;
                        toRow = table.Range._toRow;
                    }
                    else if (fromRow == 0)
                    {
                        fromRow = toRow = -1;
                    }
                    break;
                case "#this row":
                    var dr = table.DataRange;
                    if (_ws != table.WorkSheet || _range._fromRow < dr._fromRow || _range._fromRow > dr._toRow)
                    {
                        fromRow = toRow = -1;
                    }
                    else
                    {
                        fromRow = _range._fromRow;
                        toRow = _range._fromRow;
                        fixedFlag = FixedFlag.FromColFixed | FixedFlag.ToColFixed;
                    }
                    break;
            }
        }
        private void AddNameInfo(short startPos, short namePos, string er, string ws)
        {
            var t = Tokens[namePos];
            if(string.IsNullOrEmpty(er))    //TODO: add support for external refrence
            {
                ExcelNamedRange n=null;
                if (string.IsNullOrEmpty(ws))
                {
                    if (_ws.Names.ContainsKey(t.Value))
                    {
                        n = _ws.Names[t.Value];
                    }
                    else if (_ws.Workbook.Names.ContainsKey(t.Value))
                    {
                        n = _ws.Workbook.Names[t.Value];
                    }
                }
                else
                {
                    var wsRef = _ws.Workbook.Worksheets[ws];
                    if(wsRef != null)
                    {
                        n = wsRef.Names[t.Value];
                    }
                }
                if(n==null)
                {
                    //The name is a table.
                    var tbl = _ws.Workbook.GetTable(t.Value);
                    if(tbl!=null)
                    {
                        var fr = new FormulaRange(startPos, namePos, tbl.DataRange);
                        fr.Ranges[0].FixedFlag = FixedFlag.All; //a Tables data range is allways fixed.
                        TokenInfos.Add(startPos, fr);
                    }
                }
                else
                {
                    if (n.NameValue != null)
                    {
                        TokenInfos.Add(startPos, new FormulaFixedValue(startPos, namePos, n.NameValue));
                    }
                    else if(n.Formula != null)
                    {
                        TokenInfos.Add(startPos, new FormulaNamedFormula(startPos, namePos, n.NameFormula));
                    }
                    else
                    {
                        TokenInfos.Add(startPos, new FormulaRange(startPos, namePos, n));
                    }
                }
            }
        }
    }
    internal enum FormulaType
    {
        CellAddress,
        FormulaRange,
        FixedValue,
        Formula
    }
    [Flags]
    internal enum FixedFlag : byte
    {
        None = 0,
        FromRowFixed = 0x1,
        FromColFixed = 0x2,
        ToRowFixed = 0x4,
        ToColFixed = 0x8,
        All = 0xF,
    }

    internal abstract class TokenInfo
    {
        internal FormulaType Type;
        internal short TokenStartPosition;
        internal short TokenEndPosition;
        internal virtual void SetOffset(int rowOffset, int colOffset) { }
        internal virtual bool IsFixed { get { return true; } }
    }
    internal class FormulaCellAddress : TokenInfo
    {
        internal FormulaCellAddress(short pos, string cellAddress)
        {
            Type = FormulaType.CellAddress; 
            TokenStartPosition = TokenEndPosition = pos;
            ExcelCellBase.GetRowColFromAddress(cellAddress, out Row, out Col, out FixedRow, out FixedCol);
        }

        internal int Row, Col;
        internal bool FixedRow, FixedCol;
        internal override void SetOffset(int rowOffset, int colOffset)
        {
            if (!FixedRow) Row += rowOffset;
            if (!FixedCol) Col += colOffset;
        }
        internal override bool IsFixed { get { return FixedRow & FixedCol; } }
    }
    internal class FormulaFixedValue : TokenInfo
    {
        public FormulaFixedValue(short startPos, short endPos, object v)
        {
            Type = FormulaType.FixedValue;
            TokenStartPosition = startPos;
            TokenEndPosition = endPos;
            Value = v;
        }
        internal object Value;
    }
    internal class FormulaNamedFormula : TokenInfo
    {
        public FormulaNamedFormula(short startPos, short endPos, string f)
        {
            Type = FormulaType.Formula;
            TokenStartPosition = startPos;
            TokenEndPosition = endPos;
            Formula = f;
        }
        internal string Formula;
        internal override bool IsFixed { get { return false; } } //TODO: Check here if we can us fixed from the acutal formula in  later stadge.
    }
    internal class FormulaRange : TokenInfo
    {
        internal override void SetOffset(int rowOffset, int colOffset)
        { 
            for(int i=0;i < Ranges.Count;i++)
            {
                var r=Ranges[i];
                if ((r.FixedFlag & FixedFlag.FromRowFixed) == FixedFlag.None) r.FromRow += rowOffset;
                if ((r.FixedFlag & FixedFlag.ToRowFixed) == FixedFlag.None) r.ToRow += rowOffset;
                if ((r.FixedFlag & FixedFlag.FromColFixed) == FixedFlag.None) r.FromCol += colOffset;
                if ((r.FixedFlag & FixedFlag.ToColFixed) == FixedFlag.None) r.ToCol += colOffset;
            }
        }
        internal class FormulaRangeAddress
        {
            internal short er, ws=-1;
            internal int FromRow, FromCol, ToRow, ToCol;
            internal FixedFlag FixedFlag;
        }
        internal override bool IsFixed 
        {
            get
            {
                foreach(var r in Ranges)
                {
                    if(r.FixedFlag != FixedFlag.All)
                    {
                        return false;
                    }
                }
                return true;
            }
        }
        internal List<FormulaRangeAddress> Ranges;
        internal FormulaRange(short startPos, short endPos, int fromRow, int fromCol, int toRow, int toCol, FixedFlag fixedFlag)
        {
            Type = FormulaType.FormulaRange;
            TokenStartPosition = startPos;
            TokenEndPosition = endPos;
            Ranges = new List<FormulaRangeAddress>();
            Ranges.Add(
                new FormulaRangeAddress()
                {
                    FromRow = fromRow,
                    FromCol = fromCol,
                    ToRow = toRow,
                    ToCol = toCol,
                    FixedFlag = fixedFlag
                });
        }
        internal FormulaRange(short startPos, short endPos, ExcelRangeBase range)
        {
            Type = FormulaType.FormulaRange;
            TokenStartPosition = startPos;
            TokenEndPosition = endPos;
            Ranges = new List<FormulaRangeAddress>();
            if (range.Addresses == null)
            {
                Ranges.Add(
                    new FormulaRangeAddress()
                    {
                        er = (short)(string.IsNullOrEmpty(range._wb) ? 0 : range._workbook.ExternalLinks.GetExternalLink(range._wb)),
                        ws = (short)range.Worksheet.PositionId,
                        FromRow = range._fromRow,
                        FromCol = range._fromCol,
                        ToRow = range._toRow,
                        ToCol = range._toCol,

                        FixedFlag = (range._fromRowFixed ? FixedFlag.FromRowFixed : 0) |
                                    (range._fromColFixed ? FixedFlag.FromColFixed : 0) |
                                    (range._toRowFixed ? FixedFlag.ToRowFixed : 0) |
                                    (range._toColFixed ? FixedFlag.ToColFixed : 0)
                    }); 
            }
            else
            {
                foreach (var a in range.Addresses)
                {
                    Ranges.Add(
                        new FormulaRangeAddress()
                        {
                            er = (short)(string.IsNullOrEmpty(a._wb) ? -1 : range._workbook.ExternalLinks.GetExternalLink(a._wb)),
                            ws = (short)(string.IsNullOrEmpty(a.WorkSheetName) ? range.Worksheet.PositionId : (range._workbook.Worksheets[a.WorkSheetName]==null ? -1 : range._workbook.Worksheets[a.WorkSheetName].PositionId)),
                            FromRow = a._fromRow,
                            FromCol = a._fromCol,
                            ToRow = a._toRow,
                            ToCol = a._toCol,
                            FixedFlag = (a._fromRowFixed ? FixedFlag.FromRowFixed : 0) |
                                        (a._fromColFixed ? FixedFlag.FromColFixed : 0) |
                                        (a._toRowFixed ? FixedFlag.ToRowFixed : 0) |
                                        (a._toColFixed ? FixedFlag.ToColFixed : 0) 

                        });
                }
            }
        }
    }
}
