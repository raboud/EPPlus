/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * Required Notice: Copyright (C) EPPlus Software AB. 
 * https://epplussoftware.com
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "" as is "" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
  Date               Author                       Change
 *******************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *******************************************************************************/
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Globalization;
using System.Threading;

namespace EPPlusTest.Style
{
    [TestClass]
    public class TableStyleTests : TestBase
    {
        static ExcelPackage _pck;

        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("TableStyle.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }
        [TestMethod]
        public void VerifyColumnStyle()
        {
            using (var p = OpenTemplatePackage("PivotTableNamedStyles.xlsx"))
            {
                Assert.AreEqual(2, p.Workbook.Styles.TableStyles.Count);
            }
        }
        [TestMethod]
        public void AddTableStyle()
        {
            var ws = _pck.Workbook.Worksheets.Add("TableStyle");
            var s=_pck.Workbook.Styles.CreateTableStyle("CustomTableStyle1");
            s.WholeTable.Style.Font.Color.SetColor(Color.Red);
            s.FirstRowStripe.Style.Fill.PatternType = ExcelFillStyle.Solid;
            s.FirstRowStripe.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
            s.SecondRowStripe.Style.Fill.PatternType = ExcelFillStyle.Solid;
            s.SecondRowStripe.Style.Fill.BackgroundColor.SetColor(Color.LightYellow);
        }
    }
}

