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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.Tables.Cmap
{
    /// <summary>
    /// This table defines the mapping of character codes to the glyph index values used in the font. It may contain more than one subtable, in order to support more than one character encoding scheme.
    /// </summary>
    public class CmapTable
    {
        public CmapTable()
        {
            EncodingRecords = new List<EncodingRecord>();
        }
        /// <summary>
        /// Table version number (0).
        /// </summary>
        public ushort Version { get; set; }

        /// <summary>
        /// Number of encoding tables that follow.
        /// </summary>
        public ushort NumTables { get; set; }

        /// <summary>
        /// The array of encoding records specifies particular encodings and the offset to the subtable for each encoding.
        /// </summary>
        public List<EncodingRecord> EncodingRecords { get; private set; }
    }
}
