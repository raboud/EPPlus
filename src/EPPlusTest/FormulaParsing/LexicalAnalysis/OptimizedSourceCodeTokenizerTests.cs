using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.LexicalAnalysis
{
    [TestClass]
    public class OptimizedSourceCodeTokenizerTests
    {
        [TestMethod]
        public void TokenizePerformance()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                var tOld = OfficeOpenXml.FormulaParsing.LexicalAnalysis.SourceCodeTokenizer.Default;
                var tNew = OfficeOpenXml.FormulaParsing.LexicalAnalysis.OptimizedSourceCodeTokenizer.Default;
                var formula = "VLOOKUP(CONCAT(ORRange30,$H20,$F$17),Ranking!$A$1:$M$3775,MATCH(\"\"\"Value\"\"\",Ranking!$A$1:$M$1,0),0)";
                //var formula = "(-1+-2*3)*12";
                //RunTokenize(tOld, formula);
                RunTokenize(tNew, formula);

                formula = "SUM(A1:OFFSET(B1;1;3))";
                //RunTokenize(tOld, formula);
                RunTokenize(tNew, formula);
            }
        }

        private static void RunTokenize(OfficeOpenXml.FormulaParsing.LexicalAnalysis.ISourceCodeTokenizer t, string formula)
        {
            var time = DateTime.Now;
            for (int i = 0; i < 1000000; i++)
            {
                var tokens = t.Tokenize(formula, "sheet1");
            }
            var offset = new TimeSpan((DateTime.Now - time).Ticks);
            Debug.WriteLine(offset.TotalMilliseconds);
        }
    }
}
