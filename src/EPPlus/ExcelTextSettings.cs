﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/26/2021         EPPlus Software AB       EPPlus 6.0
 *************************************************************************************************/
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.GenericMeasurements;
using OfficeOpenXml.Interfaces.Text;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml
{
    /// <summary>
    /// This class contains settings for text measurement.
    /// </summary>
    public class ExcelTextSettings
    {
        public ExcelTextSettings()
        {
#if (Core)
            PrimaryTextMeasurer = new SkiaSharp.Text.SkiaSharpTextMeasurer();
            FallbackTextMeasurer = new GenericFontMetricsTextMeasurer();
#else
            PrimaryTextMeasurer = new GenericFontMetricsTextMeasurer();
#endif
            AutofitScaleFactor = 1f;
        }

        /// <summary>
        /// This is the primary text measurer
        /// </summary>
        public ITextMeasurer PrimaryTextMeasurer { get; set; }

        /// <summary>
        /// If the primary text measurer fails to measure the text, this one will be used.
        /// </summary>
        public ITextMeasurer FallbackTextMeasurer { get; set; }

        /// <summary>
        /// All measurements of texts will be multiplied with this value. Default is 1.
        /// </summary>
        public float AutofitScaleFactor { get; set; }
    }
}
