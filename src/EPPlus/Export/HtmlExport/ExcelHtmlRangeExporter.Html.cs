﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/16/2020         EPPlus Software AB           ExcelTable Html Export
 *************************************************************************************************/
using OfficeOpenXml.Core;
using OfficeOpenXml.Export.HtmlExport.Accessibility;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;


#if !NET35 && !NET40
using System.Threading.Tasks;
#endif

namespace OfficeOpenXml.Export.HtmlExport
{
    /// <summary>
    /// Exports a <see cref="ExcelTable"/> to Html
    /// </summary>
    public partial class ExcelHtmlRangeExporter : HtmlExporterBase
    {
        private readonly EPPlusReadOnlyList<ExcelRangeBase> _ranges;
        private readonly CellDataWriter _cellDataWriter = new CellDataWriter();
        internal ExcelHtmlRangeExporter
            (ExcelRangeBase range)
        {
            Require.Argument(range).IsNotNull("range");
            _ranges = new EPPlusReadOnlyList<ExcelRangeBase>();

            if(range.Addresses==null)
            {
                AddRange(range);
            }
            else
            {
                foreach(var address in range.Addresses)
                {
                    AddRange(range.Worksheet.Cells[address.Address]);
                }
            }

            LoadRangeImages(_ranges._list);
        }
        internal ExcelHtmlRangeExporter
            (ExcelRangeBase[] ranges)
        {
            Require.Argument(ranges).IsNotNull("ranges");
            _ranges = new EPPlusReadOnlyList<ExcelRangeBase>();

            foreach (var range in ranges)
            {
                AddRange(range);
            }

            LoadRangeImages(_ranges._list);
        }

        private void AddRange(ExcelRangeBase range)
        {
            if (range.IsFullColumn && range.IsFullRow)
            {
                _ranges.Add(new ExcelRangeBase(range.Worksheet, range.Worksheet.Dimension.Address));
            }
            else
            {
                _ranges.Add(range);
            }
        }

        /// <summary>
        /// Setting used for the export.
        /// </summary>
        public HtmlRangeExportSettings Settings { get; } = new HtmlRangeExportSettings();
        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <returns>A html table</returns>
        public string GetHtmlString()
        {
            using (var ms = RecyclableMemory.GetStream())
            {
                RenderHtml(ms, 0);
                ms.Position = 0;
                using (var sr = new StreamReader(ms))
                {
                    return sr.ReadToEnd();
                }
            }
        }
        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <returns>A html table</returns>
        public string GetHtmlString(int rangeIndex)
        {
            ValidateRangeIndex(rangeIndex);
            using (var ms = RecyclableMemory.GetStream())
            {
                RenderHtml(ms, rangeIndex);
                ms.Position = 0;
                using (var sr = new StreamReader(ms))
                {
                    return sr.ReadToEnd();
                }
            }
        }

        private void ValidateRangeIndex(int rangeIndex)
        {
            if (rangeIndex < 0 || rangeIndex >= _ranges.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(rangeIndex));
            }
        }

        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <param name="stream">The stream to write to</param>
        /// <returns>A html table</returns>
        public void RenderHtml(Stream stream)
        {
            RenderHtml(stream, 0);
        }
        /// <summary>
        /// Exports an <see cref="ExcelTable"/> to a html string
        /// </summary>
        /// <param name="stream">The stream to write to</param>
        /// <param name="rangeIndex">The index of the range to output.</param>
        /// <returns>A html table</returns>
        public void RenderHtml(Stream stream, int rangeIndex)
        {
            ValidateRangeIndex(rangeIndex);

            if (!stream.CanWrite)
            {
                throw new IOException("Parameter stream must be a writeable System.IO.Stream");
            }
            var range = _ranges[rangeIndex];
            GetDataTypes(range);
            
            var writer = new EpplusHtmlWriter(stream, Settings.Encoding);
            AddClassesAttributes(writer);
            AddTableAccessibilityAttributes(Settings, writer);
            writer.RenderBeginTag(HtmlElements.Table);

            writer.ApplyFormatIncreaseIndent(Settings.Minify);
            LoadVisibleColumns(range);
            if (Settings.SetColumnWidth || Settings.HorizontalAlignmentWhenGeneral==eHtmlGeneralAlignmentHandling.ColumnDataType)
            {
                SetColumnGroup(writer, range, Settings);
            }

            if (Settings.HeaderRows > 0 || Settings.Headers.Count > 0)
            {
                RenderHeaderRow(range, writer);
            }
            // table rows
            RenderTableRows(range, writer);

            // end tag table
            writer.RenderEndTag();

        }
        /// <summary>
        /// The ranges used in the export.
        /// </summary>
        public EPPlusReadOnlyList<ExcelRangeBase> Ranges 
        { 
            get 
            { 
                return _ranges;
            } 
        }
        private void AddClassesAttributes(EpplusHtmlWriter writer)
        {
            writer.AddAttribute(HtmlAttributes.Class, $"{TableClass}");
            if (!string.IsNullOrEmpty(Settings.TableId))
            {
                writer.AddAttribute(HtmlAttributes.Id, Settings.TableId);
            }
        }

        private void LoadVisibleColumns(ExcelRangeBase range)
        {
            var ws = range.Worksheet;
            _columns = new List<int>();
            for (int col = range._fromCol; col <= range._toCol; col++)
            {
                var c = ws.GetColumn(col);
                if (c == null || (c.Hidden == false && c.Width > 0))
                {
                    _columns.Add(col);
                }
            }
        }

        /// <summary>
        /// Renders both the Html and the Css to a single page. 
        /// </summary>
        /// <param name="htmlDocument">The html string where to insert the html and the css. The Html will be inserted in string parameter {0} and the Css will be inserted in parameter {1}.</param>
        /// <returns>The html document</returns>
        public string GetSinglePage(string htmlDocument = "<html>\r\n<head>\r\n<style type=\"text/css\">\r\n{1}</style></head>\r\n<body>\r\n{0}</body>\r\n</html>")
        {
            if (Settings.Minify) htmlDocument = htmlDocument.Replace("\r\n", "");
            var html = GetHtmlString();
            var css = GetCssString();
            return string.Format(htmlDocument, html, css);
        }
        List<ExcelAddressBase> _mergedCells = new List<ExcelAddressBase>();
        private void RenderTableRows(ExcelRangeBase range, EpplusHtmlWriter writer)
        {
            if (Settings.Accessibility.TableSettings.AddAccessibilityAttributes && !string.IsNullOrEmpty(Settings.Accessibility.TableSettings.TbodyRole))
            {
                writer.AddAttribute("role", Settings.Accessibility.TableSettings.TbodyRole);
            }
            writer.RenderBeginTag(HtmlElements.Tbody);
            writer.ApplyFormatIncreaseIndent(Settings.Minify);
            var row = range._fromRow + Settings.HeaderRows;
            var endRow = range._toRow;
            var ws = range.Worksheet;
            HtmlImage image = null;
            while (row <= endRow)
            {
                if (HandleHiddenRow(writer, range.Worksheet, Settings, ref row))
                {
                    continue; //The row is hidden and should not be included.
                }

                if (Settings.Accessibility.TableSettings.AddAccessibilityAttributes)
                {
                    writer.AddAttribute("role", "row");
                    writer.AddAttribute("scope", "row");
                }

                if (Settings.SetRowHeight) AddRowHeightStyle(writer, range, row, Settings.StyleClassPrefix);
                writer.RenderBeginTag(HtmlElements.TableRow);
                writer.ApplyFormatIncreaseIndent(Settings.Minify);
                foreach (var col in _columns)
                {
                    if (InMergeCellSpan(row, col)) continue;
                    var colIx = col - range._fromCol;
                    var cell = ws.Cells[row, col];
                    var dataType = HtmlRawDataProvider.GetHtmlDataTypeFromValue(cell.Value);

                    SetColRowSpan(range, writer, cell);

                    if (Settings.Pictures.Include == ePictureInclude.Include)
                    {
                        image = GetImage(cell._fromRow, cell._fromCol);
                    }
                    if (cell.Hyperlink == null)
                    {
                        _cellDataWriter.Write(cell, dataType, writer, Settings, false, image);
                    }
                    else
                    {
                        writer.RenderBeginTag(HtmlElements.TableData);
                        AddImage(writer, Settings, image, cell.Value);
                        var imageCellClassName = GetImageCellClassName(image, Settings);
                        writer.SetClassAttributeFromStyle(cell, false, Settings, imageCellClassName);
                        RenderHyperlink(writer, cell);
                        writer.RenderEndTag();
                        writer.ApplyFormat(Settings.Minify);
                    }
                }

                // end tag tr
                writer.Indent--;
                writer.RenderEndTag();
                writer.ApplyFormat(Settings.Minify);
                row++;
            }

            writer.ApplyFormatDecreaseIndent(Settings.Minify);
            // end tag tbody
            writer.RenderEndTag();
            writer.ApplyFormat(Settings.Minify);
        }
        private void RenderHeaderRow(ExcelRangeBase range, EpplusHtmlWriter writer)
        {
            if (Settings.Accessibility.TableSettings.AddAccessibilityAttributes && !string.IsNullOrEmpty(Settings.Accessibility.TableSettings.TheadRole))
            {
                writer.AddAttribute("role", Settings.Accessibility.TableSettings.TheadRole);
            }
            writer.RenderBeginTag(HtmlElements.Thead);
            writer.ApplyFormatIncreaseIndent(Settings.Minify);
            var headerRows = Settings.HeaderRows == 0 ? 1 : Settings.HeaderRows;
            HtmlImage image = null;
            for (int i = 0; i < headerRows; i++)
            {
                if (Settings.Accessibility.TableSettings.AddAccessibilityAttributes)
                {
                    writer.AddAttribute("role", "row");
                }
                var row = range._fromRow + i;
                if (Settings.SetRowHeight) AddRowHeightStyle(writer, range, row, Settings.StyleClassPrefix);
                writer.RenderBeginTag(HtmlElements.TableRow);
                writer.ApplyFormatIncreaseIndent(Settings.Minify);
                foreach (var col in _columns)
                {
                    if (InMergeCellSpan(row, col)) continue;
                    var cell = range.Worksheet.Cells[row, col];
                    if (Settings.RenderDataTypes)
                    {
                        writer.AddAttribute("data-datatype", _datatypes[col - range._fromCol]);
                    }
                    SetColRowSpan(range, writer, cell);
                    if(Settings.IncludeCssClassNames)
                    {
                        var imageCellClassName = GetImageCellClassName(image, Settings);
                        writer.SetClassAttributeFromStyle(cell, true, Settings, imageCellClassName);
                    }
                    if (Settings.Pictures.Include == ePictureInclude.Include)
                    {
                        image = GetImage(cell._fromRow, cell._fromCol);
                    }
                    AddImage(writer, Settings, image, cell.Value);
                    writer.RenderBeginTag(HtmlElements.TableHeader);

                    if (Settings.HeaderRows > 0)
                    {
                        if (cell.Hyperlink == null)
                        {
                            writer.Write(GetCellText(cell));
                        }
                        else
                        {
                            RenderHyperlink(writer, cell);
                        }
                    }
                    else if (Settings.Headers.Count < col)
                    {
                        writer.Write(Settings.Headers[col]);
                    }
                    
                    writer.RenderEndTag();
                    writer.ApplyFormat(Settings.Minify);
                }
                writer.Indent--;
                writer.RenderEndTag();
            }
            writer.ApplyFormatDecreaseIndent(Settings.Minify);
            writer.RenderEndTag();
            writer.ApplyFormat(Settings.Minify);
        }
        private bool InMergeCellSpan(int row, int col)
        {
            for(int i=0; i < _mergedCells.Count;i++)
            {
                var adr = _mergedCells[i];
                if(adr._toRow < row || (adr._toRow==row && adr._toCol<col))
                {
                    _mergedCells.RemoveAt(i);
                    i--;
                }
                else
                {
                    if(row >= adr._fromRow && row <= adr._toRow &&
                       col >= adr._fromCol && col <= adr._toCol)
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        private void SetColRowSpan(ExcelRangeBase range, EpplusHtmlWriter writer, ExcelRange cell)
        {
            if(cell.Merge)
            {
                var address = cell.Worksheet.MergedCells[cell._fromRow, cell._fromCol];
                if(address!=null)
                {
                    var ma = new ExcelAddressBase(address);
                    bool added = false;
                    //ColSpan
                    if(ma._fromCol==cell._fromCol || range._fromCol==cell._fromCol)
                    {
                        var maxCol = Math.Min(ma._toCol, range._toCol);
                        var colSpan = maxCol - ma._fromCol+1;
                        if(colSpan>1)
                        {
                            writer.AddAttribute("colspan", colSpan.ToString(CultureInfo.InvariantCulture));
                        }
                        _mergedCells.Add(ma);
                        added = true;
                    }
                    //RowSpan
                    if (ma._fromRow == cell._fromRow || range._fromRow == cell._fromRow)
                    {
                        var maxRow = Math.Min(ma._toRow, range._toRow);
                        var rowSpan = maxRow - ma._fromRow+1;
                        if (rowSpan > 1)
                        {
                            writer.AddAttribute("rowspan", rowSpan.ToString(CultureInfo.InvariantCulture));
                        }
                        if(added==false) _mergedCells.Add(ma);
                    }
                }
            }
        }

        private void RenderHyperlink(EpplusHtmlWriter writer, ExcelRangeBase cell)
        {
            if (cell.Hyperlink is ExcelHyperLink eurl)
            {
                if (string.IsNullOrEmpty(eurl.ReferenceAddress))
                {
                    if(string.IsNullOrEmpty(eurl.AbsoluteUri))
                    {
                        writer.AddAttribute("href", eurl.OriginalString);
                    }
                    else
                    {
                        writer.AddAttribute("href", eurl.AbsoluteUri);
                    }
                    writer.RenderBeginTag(HtmlElements.A);
                    writer.Write(string.IsNullOrEmpty(eurl.Display) ? cell.Text : eurl.Display);
                    writer.RenderEndTag();
                }
                else
                {
                    //Internal
                    writer.Write(GetCellText(cell));
                }
            }
            else
            {
                writer.AddAttribute("href", cell.Hyperlink.OriginalString);
                writer.RenderBeginTag(HtmlElements.A);
                writer.Write(GetCellText(cell));
                writer.RenderEndTag();
            }
        }

        private void AddTableAccessibilityAttributes(HtmlRangeExportSettings settings, EpplusHtmlWriter writer)
        {
            if (!settings.Accessibility.TableSettings.AddAccessibilityAttributes) return;
            if (!string.IsNullOrEmpty(settings.Accessibility.TableSettings.TableRole))
            {
                writer.AddAttribute("role", settings.Accessibility.TableSettings.TableRole);
            }
            if (!string.IsNullOrEmpty(settings.Accessibility.TableSettings.AriaLabel))
            {
                writer.AddAttribute(AriaAttributes.AriaLabel.AttributeName, settings.Accessibility.TableSettings.AriaLabel);
            }
            if (!string.IsNullOrEmpty(settings.Accessibility.TableSettings.AriaLabelledBy))
            {
                writer.AddAttribute(AriaAttributes.AriaLabelledBy.AttributeName, settings.Accessibility.TableSettings.AriaLabelledBy);
            }
            if (!string.IsNullOrEmpty(settings.Accessibility.TableSettings.AriaDescribedBy))
            {
                writer.AddAttribute(AriaAttributes.AriaDescribedBy.AttributeName, settings.Accessibility.TableSettings.AriaDescribedBy);
            }
        }

        private string GetCellText(ExcelRangeBase cell)
        {
            if (cell.IsRichText)
            {
                return cell.RichText.HtmlText;
            }
            else
            {
                return ValueToTextHandler.GetFormattedText(cell.Value, cell.Worksheet.Workbook, cell.StyleID, false, Settings.Culture);
            }
        }

        private void GetDataTypes(ExcelRangeBase range)
        {
            if (range._fromRow + Settings.HeaderRows > ExcelPackage.MaxRows)
            {
                throw new InvalidOperationException("Range From Row + Header rows is out of bounds");
            }

            _datatypes = new List<string>();
            for (int col = range._fromCol; col <= range._toCol; col++)
            {
                _datatypes.Add(
                    ColumnDataTypeManager.GetColumnDataType(range.Worksheet, range, range._fromRow + Settings.HeaderRows, col));
            }
        }
    }
}

