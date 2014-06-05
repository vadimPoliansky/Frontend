using System;
using System.IO;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using IndInv.Models;
using IndInv.Models.ViewModels;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using ClosedXML.Excel;
using SpreadsheetLight;
using SpreadsheetLight.Drawing;

namespace IndInv.Controllers
{

    public class IndicatorController : Controller
    {
        private InventoryDBContext db = new InventoryDBContext();

        //
        // GET: /Indicator/

        [HttpGet]
        public ActionResult Index()
        {
            /*var viewModel = new indexViewModel
            {
                allIndicators = db.Indicators.ToList(),
                allCoEs = db.CoEs.ToList(),
                allAreas = db.Areas.ToList(),
                allFootnotes= db.Footnotes.ToList()
            };*/

            var viewModel = new indexViewModel
            {
                allIndicators = db.Indicators.ToList(),
                allCoEs = db.CoEs.Where(x => x.CoE_ID == 10 || x.CoE_ID == 27 || x.CoE_ID == 30 || x.CoE_ID == 40 || x.CoE_ID == 50).ToList(),
                allAreas = db.Areas.Where(x => x.Area_ID == 1).ToList(),
                //allFootnotes= db.Footnotes.Where(x => x.Footnote_ID == 9999).ToList()
            };

            return View(viewModel);
        }

        [HttpPost]
        public ActionResult searchAdvanced(IList<searchViewModel> advancedSearch)
        {
                TempData["search"] = advancedSearch.FirstOrDefault();
                return Json(Url.Action("searchResults", "Indicator")); 
        }

        public ActionResult searchResults()
        {
            TempData.Keep();
            searchViewModel advancedSearch = (searchViewModel)TempData["search"];

            if (advancedSearch == null)
            {
                return RedirectToAction("Index");
            }

            //List<Indicators> indicatorList = db.Indicators.ToList();
            List<Indicators> indicatorList = db.Indicators.Where(x => x.Area_ID.Equals(1)).Where(y => y.Indicator_CoE_Map.Any(x => x.CoE_ID.Equals(10) || x.CoE_ID.Equals(27) || x.CoE_ID.Equals(30) || x.CoE_ID.Equals(40) || x.CoE_ID.Equals(50))).ToList();
            List<Indicators> indicatorListString = new List<Indicators>();

            string searchString = advancedSearch.searchString;
            if (searchString != null)
            {
                string[] searchStrings;
                searchStrings = searchString.Split(' ');
                foreach (var sS in searchStrings)
                {
                    indicatorList = indicatorList.Where(s => s.Indicator != null && s.Indicator.ToLower().Contains(sS.ToLower())).ToList();
                }
            }

            List<Indicators> indicatorListCoE = new List<Indicators>();
            List<selectedCoEs> searchCoEs;
            searchCoEs = advancedSearch.selectedCoEs;
            if (searchCoEs != null)
            {
                foreach (var coe in searchCoEs)
                {
                    indicatorListCoE.AddRange(db.Indicators.Where(s => s.Indicator_CoE_Map.Any(x => x.CoE_ID == coe.CoE_ID)).ToList());
                }
                indicatorList = indicatorList.Intersect(indicatorListCoE).ToList();
            }
            
            List<Indicators> indicatorListAreas = new List<Indicators>();
            List<selectedAreas> searchAreas;
            searchAreas = advancedSearch.selectedAreas;
            if (searchAreas != null)
            {
                foreach (var area in searchAreas)
                {
                    indicatorListAreas.AddRange(db.Indicators.Where(s => s.Area_ID == area.Area_ID).ToList());
                }
                indicatorList = indicatorList.Intersect(indicatorListAreas).ToList();
            }

            List<Indicators> indicatorListTypes = new List<Indicators>();
            List<selectedTypes> searchTypes;
            searchTypes = advancedSearch.selectedTypes;
            if (searchTypes != null)
            {
                foreach (var type in searchTypes)
                {
                    indicatorListTypes.AddRange(db.Indicators.Where(s => s.Indicator_Type.Replace("/","").Replace("&","").Replace(" ","") == type.Indicator_Type).ToList());
                }
                indicatorList = indicatorList.Intersect(indicatorListTypes).ToList();
            }

            List<Indicators> indicatorListFootnotes = new List<Indicators>();
            List<selectedFootnotes> searchFootnotes;
            searchFootnotes = advancedSearch.selectedFootnotes;
            if (searchFootnotes != null)
            {
                foreach (var footnote in searchFootnotes)
                {
                    indicatorListFootnotes.AddRange(db.Indicators.Where(s => s.Indicator_Footnote_Map.Any(x => x.Footnote_ID == footnote.Footnote_ID)).ToList());
                }
                indicatorList = indicatorList.Intersect(indicatorListFootnotes).ToList();
            }

            if (ModelState.IsValid)
            {
                var viewModel = new indexViewModel
                {
                    allIndicators = indicatorList.Distinct().ToList(),
                    
//                    allCoEs = db.CoEs.ToList(),
//                    allAreas = db.Areas.ToList(),
 //                   allFootnotes = db.Footnotes.ToList()
                };
                return View(viewModel);
            }

            return View();
        }

        public ActionResult viewPR()
        {
            ModelState.Clear();
            var viewModel = new PRViewModel
            {
                //allCoEs = db.CoEs.ToList(),
                allCoEs = db.CoEs.ToList(),
                allMaps = db.Indicator_CoE_Maps.ToList(),
                allFootnoteMaps = db.Indicator_Footnote_Maps.ToList()
            };

            return View(viewModel);
        }

        public ActionResult viewPRExcel()
        {
            ModelState.Clear();
            var viewModel = new PRViewModel
            {
                //allCoEs = db.CoEs.ToList(),
                allCoEs = db.CoEs.ToList(),
                allMaps = db.Indicator_CoE_Maps.ToList(),
                allFootnoteMaps = db.Indicator_Footnote_Maps.ToList()
            };

            // Create the workbook
            var wb = new XLWorkbook();

            var prBlue = XLColor.FromArgb(0, 51, 102);
            var prGreen = XLColor.FromArgb(0,118,53);
            var prYellow = XLColor.FromArgb(255,192,0);
            var prRed = XLColor.FromArgb(255,0,0);
            var prHeader1Fill = prBlue;
            var prHeader1Font = XLColor.White;
            var prHeader2Fill = XLColor.White;
            var prHeader2Font = XLColor.Black;
            var prBorder = XLColor.FromArgb(0, 0, 0);
            var prAreaFill = XLColor.FromArgb(192, 192, 192);
            var prAreaFont = XLColor.Black;
            var prBorderWidth = XLBorderStyleValues.Thin;
            var prFontSize = 10;
            var prTitleFont = 20;
            var prFootnoteSize = 8;
            var prHeightSeperator = 7.5;

            var prAreaObjectiveFontsize = 8;
            var indentLength = 24;
            var firstIndentLength = 20;
            var innerIndentLength = 5;
            var newLineHeight = 15;

            var defNote = "Portal data from the Canadian Institute for Health Information (CIHI) has been used to generate data within this report with acknowledgement to CIHI, the Ministry of Health and Long-Term Care (MOHLTC) and Stats Canada (as applicable). Views are not those of the acknowledged sources. Facility identifiable data other than Mount Sinai Hospital (MSH) is not to be published without the consent of that organization (except where reported at an aggregate level). As this is not a database supported by MSH, please demonstrate caution with use and interpretation of the information. MSH is not responsible for any changes derived from the source data/canned reports. Data may be subject to change.";

            var prNumberWidth = 4;
            var prIndicatorWidth = 55;
            var prValueWidth = 11;
            var prDefWidth = 100;
            var prRatiWidth = 50;
            var prCompWidth = 50;

            var fitRatio = 3.77;
            var fitHeight = 672;
            var fitWidth = 178.18;
            var fitAdjust = 1.325*0.1;
            List<int> fitAdjustableRows = new List<int>();

            var prFootnoteCharsNewLine = 125;
            var prObjectivesCharsNewLine = 226;

            foreach (var coe in viewModel.allCoEs)
            {
                var wsPRName = coe.CoE_Abbr;
                var wsDefName = "Def_" + coe.CoE_Abbr;
                var wsPR = wb.Worksheets.Add(wsPRName);
                var wsDef = wb.Worksheets.Add(wsDefName);
                List<IXLWorksheet> wsList = new List<IXLWorksheet>();
                wsList.Add(wsPR);
                wsList.Add(wsDef);

                foreach (var ws in wsList)
                {
                    var currentRow = 3;
                    ws.Row(2).Height = 21;
                    int startRow;
                    int indicatorNumber = 1;

                    ws.PageSetup.Margins.Top = 0;
                    ws.PageSetup.Margins.Header = 0;
                    ws.PageSetup.Margins.Left = 0.5;
                    ws.PageSetup.Margins.Right = 0.5;
                    ws.PageSetup.Margins.Bottom = 0.5;
                    ws.PageSetup.PageOrientation = XLPageOrientation.Landscape;
                    ws.PageSetup.PaperSize = XLPaperSize.LegalPaper;
                    ws.PageSetup.FitToPages(1, 1);

                    string[,] columnHeaders = new string[0,0];
                    if (ws.Name == wsPRName)
                    {
                        columnHeaders = new string[,]{
                            {"Number",""},
                            {"Indicator",""},
                            {"FY_10_11",""},
                            {"FY_11_12",""},
                            {"FY_12_13",""},
                            {"FY 13 14 Performance","FY_13_14_Q1"},
                            {"FY 13 14 Performance","FY_13_14_Q2"},
                            {"FY 13 14 Performance","FY_13_14_Q3"},
                            {"FY 13 14 Performance","FY_13_14_Q4"},
                            {"FY 13 14 Performance","FY_13_14_YTD"},
                            {"Target",""},
                            {"Performance_Threshold",""},
                            {"Comparator",""}
                        };
                    }
                    else if (ws.Name == wsDefName)
                    {
                         columnHeaders = new string[,]{
                            {"Number",""},
                            {"Indicator",""},
                            {"Definition_Calculation",""},
                            {"Target_Rationale",""},
                            {"Comparator_Source",""}
                        };
                    }

                    var currentCol = 1;
                    var prHeader2ColStart = 99;
                    var prHeader2ColEnd = 1;
                    int maxCol = columnHeaders.GetUpperBound(0) + 1;

                    var prTitle = ws.Cell(currentRow, 1);
                    prTitle.Value = coe.CoE;
                    prTitle.Style.Font.FontSize = prTitleFont;
                    prTitle.Style.Font.Bold = true;
                    prTitle.Style.Font.FontColor = prHeader1Font;
                    prTitle.Style.Fill.BackgroundColor = prHeader1Fill;
                    ws.Range(ws.Cell(currentRow, 1), ws.Cell(currentRow, maxCol)).Merge();
                    ws.Range(ws.Cell(currentRow + 1, 1), ws.Cell(currentRow + 1, maxCol)).Merge();
                    ws.Row(currentRow + 1).Height = prHeightSeperator;
                    currentRow += 2;
                    startRow = currentRow;

                    for (int i = 0; i <= columnHeaders.GetUpperBound(0); i++)
                    {
                        if (columnHeaders[i, 1] == "")
                        {
                            var columnField = columnHeaders[i, 0];
                            string cellValue;
                            Type t = typeof(Indicators);
                            cellValue = t.GetProperty(columnField) != null ?
                                ModelMetadataProviders.Current.GetMetadataForProperty(null, typeof(Indicators), columnField).DisplayName :
                                ModelMetadataProviders.Current.GetMetadataForProperty(null, typeof(Indicator_CoE_Maps), columnField).DisplayName;
                            ws.Cell(currentRow, currentCol).Value = cellValue;
                            ws.Range(ws.Cell(currentRow, currentCol), ws.Cell(currentRow + 1, currentCol)).Merge();
                            currentCol++;
                        }
                        else
                        {
                            var columnField = columnHeaders[i, 1];
                            var columnFieldTop = columnHeaders[i, 0];
                            ws.Cell(currentRow + 1, currentCol).Value = ModelMetadataProviders.Current.GetMetadataForProperty(null, typeof(Indicators), columnField).DisplayName;
                            ws.Cell(currentRow, currentCol).Value = columnFieldTop;
                            if (currentCol < prHeader2ColStart) { prHeader2ColStart = currentCol; }
                            if (currentCol > prHeader2ColEnd) { prHeader2ColEnd = currentCol; }
                            currentCol++;
                        }
                    }
                    currentCol--;
                    ws.Range(ws.Cell(currentRow, prHeader2ColStart).Address, ws.Cell(currentRow, prHeader2ColEnd).Address).Merge();
                    var prHeader1 = ws.Range(ws.Cell(currentRow, 1).Address, ws.Cell(currentRow + 1, currentCol).Address);
                    var prHeader2 = ws.Range(ws.Cell(currentRow + 1, prHeader2ColStart).Address, ws.Cell(currentRow + 1, prHeader2ColEnd).Address);

                    prHeader1.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    prHeader1.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    prHeader1.Style.Fill.BackgroundColor = prHeader1Fill;
                    prHeader1.Style.Font.FontColor = prHeader1Font;

                    prHeader2.Style.Fill.BackgroundColor = prHeader2Fill;
                    prHeader2.Style.Font.FontColor = prHeader2Font;

                    currentRow += 2;

                    List<Footnotes> footnotes = new List<Footnotes>();
                    foreach (var areaMap in coe.Area_CoE_Map.OrderBy(x=> x.Area.Sort))
                    {
                        var cellLengthObjective = 0;
                        var prArea = ws.Range(ws.Cell(currentRow, 1), ws.Cell(currentRow, maxCol));
                        fitAdjustableRows.Add(currentRow);
                        prArea.Merge();
                        prArea.Style.Fill.BackgroundColor = prAreaFill;
                        prArea.Style.Font.FontColor = prAreaFont;
                        prArea.FirstCell().RichText.AddText(areaMap.Area.Area).Bold = true;
                        cellLengthObjective += areaMap.Area.Area.Length;

                        if (ws == wsPR)
                        {
                            var indent = new string('_', indentLength);
                            var innerIndent = new string('_', innerIndentLength);
                            var firstIndent = indent.Substring(0, firstIndentLength - areaMap.Area.Area.Length);

                            var stringSeperators = new string[] { "•" };
                            if (areaMap.Objective != null)
                            {
                                var objectives = areaMap.Objective.Split(stringSeperators, StringSplitOptions.None);
                                //var maxObjectiveLength = 0;
                                //foreach (var obj in objectives) {
                                //    if (obj.Length > maxObjectiveLength && obj.Length < prObjectivesCharsNewLine/2){
                                //        maxObjectiveLength = obj.Length;
                                //    }
                                //}
                                for (var i = 1; i < objectives.Length; i++)
                                {
                                    if (i == 1)
                                    {
                                        prArea.FirstCell().RichText.AddText(firstIndent).SetFontColor(prAreaFill).SetFontSize(prAreaObjectiveFontsize);
                                        cellLengthObjective += firstIndent.Length;
                                    }
                                    //var innerIndentAdj = new string('_', maxObjectiveLength < objectives[i].Length ? 0 : maxObjectiveLength - objectives[i].Length);
                                    var innerIndentAdj = "";

                                    cellLengthObjective += objectives[i].Length + innerIndent.Length + innerIndentAdj.Length;
                                    if (cellLengthObjective > prObjectivesCharsNewLine)
                                    {
                                        prArea.FirstCell().RichText.AddNewLine();
                                        ws.Row(currentRow).Height += newLineHeight;
                                        //prArea.FirstCell().RichText.AddText(indent).FontColor = prAreaFill;
                                        prArea.FirstCell().RichText.AddText(indent).SetFontColor(prAreaFill).SetFontSize(prAreaObjectiveFontsize);
                                        cellLengthObjective = indent.Length;
                                    }
                                    prArea.FirstCell().RichText.AddText(innerIndent + innerIndentAdj).FontColor = prAreaFill;
                                    prArea.FirstCell().RichText.AddText(" •" + objectives[i]).FontSize = prAreaObjectiveFontsize;
                                    cellLengthObjective += objectives[i].Length;
                                }
                            }
                        }

                        currentRow++;

                        foreach (var map in viewModel.allMaps.Where(e => e.Indicator.Area.Equals(areaMap.Area)).Where(d => d.CoE.CoE.Contains(coe.CoE)).OrderBy(f => f.Number))
                        {
                            fitAdjustableRows.Add(currentRow);
                            currentCol = 1;

                            ws.Cell(currentRow, currentCol).Value = indicatorNumber;
                            indicatorNumber++;
                            ws.Cell(currentRow, currentCol).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            currentCol++;

                            int j = 0;
                            ws.Cell(currentRow, currentCol).Value = map.Indicator.Indicator;
                            foreach (var footnote in map.Indicator.Indicator_Footnote_Map.Where(e => e.Indicator_ID == map.Indicator_ID).OrderBy(e => e.Indicator_ID))
                            {
                                if (!footnotes.Contains(footnote.Footnote)) { footnotes.Add(footnote.Footnote); }
                                if (j != 0)
                                {
                                    ws.Cell(currentRow, currentCol).RichText.AddText(",").VerticalAlignment = XLFontVerticalTextAlignmentValues.Superscript;
                                }
                                ws.Cell(currentRow, currentCol).RichText.AddText(footnote.Footnote_ID).VerticalAlignment = XLFontVerticalTextAlignmentValues.Superscript;
                                j++;
                            }
                            ws.Cell(currentRow, currentCol).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                            currentCol++;

                            if (ws.Name == wsPRName)
                            {
                                string[,] columnIndicators = new string[,]{
                                    {map.Indicator.FY_10_11, map.Indicator.FY_10_11_Sup,""},
                                    {map.Indicator.FY_11_12, map.Indicator.FY_11_12_Sup,""},
                                    {map.Indicator.FY_12_13, map.Indicator.FY_12_13_Sup,""},
                                    {map.Indicator.FY_13_14_Q1, map.Indicator.FY_13_14_Q1_Sup,map.Indicator.Q1_Color,},
                                    {map.Indicator.FY_13_14_Q2, map.Indicator.FY_13_14_Q2_Sup,map.Indicator.Q2_Color,},
                                    {map.Indicator.FY_13_14_Q3, map.Indicator.FY_13_14_Q3_Sup,map.Indicator.Q3_Color,},
                                    {map.Indicator.FY_13_14_Q4, map.Indicator.FY_13_14_Q4_Sup,map.Indicator.Q4_Color,},
                                    {map.Indicator.FY_13_14_YTD, map.Indicator.FY_13_14_YTD_Sup,map.Indicator.YTD_Color,},
                                    {map.Indicator.Target, map.Indicator.Target_Sup,""},
                                    {map.Indicator.Performance_Threshold, map.Indicator.Performance_Threshold_Sup,""},
                                    {map.Indicator.Comparator, map.Indicator.Comparator_Sup,""},
                                };
                                var startCol = currentCol;
                                int k = 1;
                                for (var i = 0; i <= columnIndicators.GetUpperBound(0); i++)
                                {
                                    for (j = 0; j <= columnIndicators.GetUpperBound(1); j++)
                                    {
                                        if (columnIndicators[i, j] != null)
                                        {
                                            columnIndicators[i, j] = columnIndicators[i, j].Replace("<b>", "");
                                            columnIndicators[i, j] = columnIndicators[i, j].Replace("</b>", "");
                                            columnIndicators[i, j] = columnIndicators[i, j].Replace("<u>", "");
                                            columnIndicators[i, j] = columnIndicators[i, j].Replace("</u>", "");
                                            columnIndicators[i, j] = columnIndicators[i, j].Replace("<i>", "");
                                            columnIndicators[i, j] = columnIndicators[i, j].Replace("</i>", "");
                                            columnIndicators[i, j] = columnIndicators[i, j].Replace("<sup>", "");
                                            columnIndicators[i, j] = columnIndicators[i, j].Replace("</sup>", "");
                                            columnIndicators[i, j] = columnIndicators[i, j].Replace("<sub>", "");
                                            columnIndicators[i, j] = columnIndicators[i, j].Replace("</sub>", "");
                                        }
                                    }
                                    if (i != columnIndicators.GetUpperBound(0) && columnIndicators[i, 0] == "=")
                                    {
                                        k = 1;
                                        while (columnIndicators[i + k, 0] == "=") { k++; }
                                        ws.Range(ws.Cell(currentRow, startCol + i - 1), ws.Cell(currentRow, startCol + i + k - 1)).Merge();
                                        i += k - 1;
                                        k = 1;
                                    }
                                    else if (columnIndicators[i, 0] != "=")
                                    {
                                        var cell = ws.Cell(currentRow, currentCol + i);
                                        string cellValue = "";

                                        if (columnIndicators[i, 0] != null)
                                        {
                                            cellValue = columnIndicators[i, 0].ToString();
                                        }

                                        if (cellValue.Contains("$"))
                                        {
                                        }

                                        cell.Value = "'" + cellValue;
                                        if (columnIndicators[i, 1] != null)
                                        {
                                            cell.RichText.AddText(columnIndicators[i, 1]).VerticalAlignment = XLFontVerticalTextAlignmentValues.Superscript;
                                        }
                                        switch (columnIndicators[i, 2])
                                        {
                                            case "cssWhite":
                                                cell.RichText.SetFontColor(XLColor.Black);
                                                cell.Style.Fill.BackgroundColor = XLColor.White;
                                                break;
                                            case "cssGreen":
                                                cell.RichText.SetFontColor(XLColor.White);
                                                cell.Style.Fill.BackgroundColor = prGreen;
                                                break;
                                            case "cssYellow":
                                                cell.RichText.SetFontColor(XLColor.Black);
                                                cell.Style.Fill.BackgroundColor = prYellow;
                                                break;
                                            case "cssRed":
                                                cell.RichText.SetFontColor(XLColor.White);
                                                cell.Style.Fill.BackgroundColor = prRed;
                                                break;
                                        }
                                        cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                                    }
                                }
                                currentRow++;
                            }
                            else if (ws.Name == wsDefName)
                            {
                                ws.Cell(currentRow, currentCol).Value = map.Indicator.Definition_Calculation;
                                ws.Cell(currentRow, currentCol).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                                currentCol++;
                                ws.Cell(currentRow, currentCol).Value = map.Indicator.Target_Rationale;
                                ws.Cell(currentRow, currentCol).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                                currentCol++;
                                ws.Cell(currentRow, currentCol).Value = map.Indicator.Comparator_Source;
                                ws.Cell(currentRow, currentCol).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                                currentCol++;
                                currentRow++;
                            }
                        }
                    }

                    var footnoteRow = ws.Range(ws.Cell(currentRow, 1), ws.Cell(currentRow, maxCol));
                    footnoteRow.Merge();
                    footnoteRow.Style.Font.FontSize = prFootnoteSize;

                    Footnotes defaultFootnote = db.Footnotes.FirstOrDefault(x => x.Footnote_Symbol == "*");
                    if (!footnotes.Contains(defaultFootnote))
                    {
                        footnotes.Add(defaultFootnote);
                    }

                    int cellLengthFootnote = 0;
                    if (ws.Name == wsPRName)
                    {
                        foreach (var footnote in footnotes.OrderBy(x => x.Footnote_Symbol))
                        {
                            ws.Cell(currentRow, 1).RichText.AddText(footnote.Footnote_Symbol).VerticalAlignment = XLFontVerticalTextAlignmentValues.Superscript;
                            ws.Cell(currentRow, 1).RichText.AddText(" " + footnote.Footnote + ";");
                            ws.Cell(currentRow, 1).Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                            cellLengthFootnote += footnote.Footnote_Symbol.ToString().Length + footnote.Footnote.ToString().Length + 2;
                            if (cellLengthFootnote > prFootnoteCharsNewLine)
                            {
                                ws.Cell(currentRow, 1).RichText.AddNewLine();
                                cellLengthFootnote = 0;
                                ws.Row(currentRow).Height += newLineHeight;
                            }
                        }
                    }
                    else
                    {
                        ws.Cell(currentRow, 1).Value = defNote;
                        ws.Row(currentRow).Height = 28;
                    }

                    var pr = ws.Range(ws.Cell(startRow, 1), ws.Cell(currentRow - 1, maxCol));

                    pr.Style.Border.InsideBorder = prBorderWidth;
                    pr.Style.Border.InsideBorderColor = prBorder;
                    pr.Style.Border.OutsideBorder = prBorderWidth;
                    pr.Style.Border.OutsideBorderColor = prBorder;
                    pr.Style.Font.FontSize = prFontSize;

                    pr = ws.Range(ws.Cell(startRow, 1), ws.Cell(currentRow, maxCol));
                    pr.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    pr.Style.Alignment.WrapText = true;

                    ws.Column(1).Width = prNumberWidth;
                    ws.Column(2).Width = prIndicatorWidth;
                    if (ws.Name == wsPR.Name)
                    {
                        for (var i = 3; i <= 15; i++)
                        {
                            ws.Column(i).Width = ws.Name == wsPRName ? prValueWidth : prDefWidth;
                        }
                    }
                    else
                    {
                        ws.Column(3).Width = prDefWidth;
                        ws.Column(4).Width = prRatiWidth;
                        ws.Column(5).Width = prCompWidth;
                    }
                    footnotes.Clear();
                    indicatorNumber = 1;

                    var totalHeight = 0.0;
                    foreach (var row in ws.Rows(1, ws.LastRowUsed().RowNumber()))
                    {
                        totalHeight += row.Height;
                    }
                    var totalWidth = 0.0;
                    foreach (var col in ws.Columns(1, ws.LastColumnUsed().ColumnNumber()))
                    {
                        totalWidth += col.Width;
                    }
                    var totalRatio = totalHeight / totalWidth;
                    if (totalRatio > fitRatio)
                    {
                        var fitAddWidthTotal = (totalHeight - fitHeight);
                        var fitAddWidthPer = fitAddWidthTotal / (ws.LastRowUsed().RowNumber() - 1) / fitAdjust;
                        foreach (var col in ws.Columns().Skip(1))
                        {
                            col.Width += fitAddWidthPer;
                        }
                    }
                    else
                    {
                        var fitAddHeightTotal = -(totalHeight - fitHeight);
                        var fitAddHeightPer = fitAddHeightTotal / ws.LastRowUsed().RowNumber() / fitAdjust;
                        foreach (var row in ws.Rows().Skip(3))
                        {
                            row.Height += fitAddHeightPer;
                        }
                    }
                }
            }

            MemoryStream preImage = new MemoryStream();
            wb.SaveAs(preImage);

            MemoryStream postImage = new MemoryStream();
            SLDocument postImageWb = new SLDocument(preImage);

            string picPath = this.HttpContext.ApplicationInstance.Server.MapPath("~/Images/logo.png");
            SLPicture picLogo = new SLPicture(picPath);
            string picPathOPEO = this.HttpContext.ApplicationInstance.Server.MapPath("~/Images/logoOPEO.png");
            SLPicture picLogoOPEO = new SLPicture(picPathOPEO);
            string picMonthlyPath = this.HttpContext.ApplicationInstance.Server.MapPath("~/Images/Monthly.png");
            SLPicture picMonthly = new SLPicture(picMonthlyPath);
            string picQuaterlyPath = this.HttpContext.ApplicationInstance.Server.MapPath("~/Images/quaterly.png");
            SLPicture picQuaterly = new SLPicture(picQuaterlyPath);
            string picNAPath = this.HttpContext.ApplicationInstance.Server.MapPath("~/Images/na.png");
            SLPicture picNA = new SLPicture(picNAPath);
            string picTargetPath = this.HttpContext.ApplicationInstance.Server.MapPath("~/Images/target.png");
            SLPicture picTarget = new SLPicture(picTargetPath);

            foreach (var ws in wb.Worksheets)
            {
                postImageWb.SelectWorksheet(ws.Name);

                picLogo.SetPosition(0, 0);
                picLogo.ResizeInPercentage(25, 25);
                postImageWb.InsertPicture(picLogo);

                picLogoOPEO.SetRelativePositionInPixels(0, ws.LastColumnUsed().ColumnNumber() + 1, -140, 0);
                picLogoOPEO.ResizeInPercentage(45, 45);
                postImageWb.InsertPicture(picLogoOPEO);

                if (ws.Name.Substring(0, 3) != "Def")
                {
                    picTarget.SetRelativePositionInPixels(ws.LastRowUsed().RowNumber(), ws.LastColumnUsed().ColumnNumber() + 1, -240, 1);
                    picNA.SetRelativePositionInPixels(ws.LastRowUsed().RowNumber(), ws.LastColumnUsed().ColumnNumber() + 1, -400, 1);
                    picMonthly.SetRelativePositionInPixels(ws.LastRowUsed().RowNumber(), ws.LastColumnUsed().ColumnNumber() + 1, -500, 1);
                    picQuaterly.SetRelativePositionInPixels(ws.LastRowUsed().RowNumber(), ws.LastColumnUsed().ColumnNumber() + 1, -490, 1);

                    picMonthly.ResizeInPercentage(70, 70);
                    picQuaterly.ResizeInPercentage(70, 70);
                    picNA.ResizeInPercentage(70, 70);
                    picTarget.ResizeInPercentage(70, 70);

                    postImageWb.InsertPicture(picMonthly);
                    postImageWb.InsertPicture(picQuaterly);
                    postImageWb.InsertPicture(picNA);
                    postImageWb.InsertPicture(picTarget);
                }
            }

            // Prepare the response
            HttpResponse httpResponse = this.HttpContext.ApplicationInstance.Context.Response;
            httpResponse.Clear();
            httpResponse.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            httpResponse.AddHeader("content-disposition", "attachment;filename=\"test.xlsx\"");

            // Flush the workbook to the Response.OutputStream
            using (MemoryStream memoryStream = new MemoryStream())
            {
                postImageWb.SaveAs(memoryStream);
                memoryStream.WriteTo(httpResponse.OutputStream);
                memoryStream.Close();
            }

            httpResponse.End();

            return View(viewModel);
        }

        public ActionResult editCoEMaps()
        {
            var viewModel = new Indicator_CoE_MapsViewModel
            {
                allIndicators = db.Indicators.ToList(),
                allCoEs = db.CoEs.ToList(),
                allMaps = db.Indicator_CoE_Maps.ToList()
            };
            return View(viewModel);
        }

        [HttpPost]
        public ActionResult editCoEMaps(IList<Indicator_CoE_MapsViewModel> mapChange)
        {
            Indicator_CoE_Maps map = mapChange.FirstOrDefault().allMaps.FirstOrDefault();
            Indicator_CoE_Maps existingMap = db.Indicator_CoE_Maps.Where(x => x.Indicator_ID == map.Indicator_ID && x.CoE_ID == map.CoE_ID).FirstOrDefault();
            if (existingMap != null)
            {
                map.Map_ID = existingMap.Map_ID;
            }
            if (map.Map_ID != 0)
            {
                var mapID = map.Map_ID;
                var deleteMap = db.Indicator_CoE_Maps.Find(mapID);
                if (deleteMap != null)
                {
                    db.Indicator_CoE_Maps.Remove(deleteMap);
                    db.SaveChanges();
                }
            }
            else
            {
                db.Indicator_CoE_Maps.Add(map);
                db.SaveChanges();
            }

            return RedirectToAction("editCoEMaps");
        }

        [HttpGet]
        public ActionResult editFootnotes(String Footnote_ID_Filter)
        {
            var viewModelItems = db.Footnotes.ToArray();
            var viewModel = viewModelItems.OrderBy(x => x.Footnote_ID).Select(x => new FootnotesViewModel
            {
                Footnote_ID = x.Footnote_ID,
                Footnote = x.Footnote,
                Footnote_Symbol = x.Footnote_Symbol,
            }).ToList();
            if (Request.IsAjaxRequest())
            {
                return Json(viewModel.Where(x => x.Footnote_ID.ToString().Contains(Footnote_ID_Filter == null ? "" : Footnote_ID_Filter)), JsonRequestBehavior.AllowGet);
            }
            else
            {
                return View(viewModel);
            }

        }

        [HttpPost]
        public ActionResult editFootnotes(IList<Footnotes> footnoteChange)
        {
            var footnoteID = footnoteChange[0].Footnote_ID;
            if (db.Footnotes.Any(x => x.Footnote_ID == footnoteID))
            {
                if (ModelState.IsValid)
                {
                    db.Entry(footnoteChange[0]).State = EntityState.Modified;
                    db.SaveChanges();
                    return View();
                }
                return View();
            }
            else
            {
                if (ModelState.IsValid)
                {
                    db.Footnotes.Add(footnoteChange[0]);
                    db.SaveChanges();
                    return View();
                }
                return View();
            }

        }

        public ActionResult editFootnoteMaps()
        {
            List<Indicator_Footnote_Maps> footnoteMaps = new List<Indicator_Footnote_Maps>();
            foreach (var footnote in db.Indicator_Footnote_Maps.OrderBy(e => e.Map_ID).ToList())
            {
                footnoteMaps.Add(footnote);
            }

            var viewModel = db.Indicators.Select(x => new Indicator_Footnote_MapsViewModel
            {
                Indicator_ID = x.Indicator_ID,
                Indicator= x.Indicator,
            }).ToList();

            foreach (var Indicator in viewModel)
            {
                if (footnoteMaps.Count(e => e.Indicator_ID == Indicator.Indicator_ID) > 0)
                {
                    Indicator.Footnote_ID_1 = footnoteMaps.Where(e => e.Indicator_ID == Indicator.Indicator_ID).OrderBy(e => e.Map_ID).First().Footnote_ID;
                    Indicator.Map_ID_1 = footnoteMaps.Where(e => e.Indicator_ID == Indicator.Indicator_ID).OrderBy(e => e.Map_ID).First().Map_ID;
                    Indicator.Footnote_Symbol_1 = db.Footnotes.FirstOrDefault(e => e.Footnote_ID == Indicator.Footnote_ID_1).Footnote_Symbol;
                }
                if (footnoteMaps.Count(e => e.Indicator_ID == Indicator.Indicator_ID) > 1)
                {
                    Indicator.Footnote_ID_2 = footnoteMaps.Where(e => e.Indicator_ID == Indicator.Indicator_ID).OrderBy(e => e.Map_ID).Skip(1).First().Footnote_ID;
                    Indicator.Map_ID_2 = footnoteMaps.Where(e => e.Indicator_ID == Indicator.Indicator_ID).OrderBy(e => e.Map_ID).Skip(1).First().Map_ID;
                    Indicator.Footnote_Symbol_2 = db.Footnotes.FirstOrDefault(e => e.Footnote_ID == Indicator.Footnote_ID_2).Footnote_Symbol;
                }
                if (footnoteMaps.Count(e => e.Indicator_ID == Indicator.Indicator_ID) > 2)
                {
                    Indicator.Footnote_ID_3 = footnoteMaps.Where(e => e.Indicator_ID == Indicator.Indicator_ID).OrderBy(e => e.Map_ID).Skip(2).First().Footnote_ID;
                    Indicator.Map_ID_3 = footnoteMaps.Where(e => e.Indicator_ID == Indicator.Indicator_ID).OrderBy(e => e.Map_ID).Skip(2).First().Map_ID;
                    Indicator.Footnote_Symbol_3 = db.Footnotes.FirstOrDefault(e => e.Footnote_ID == Indicator.Footnote_ID_3).Footnote_Symbol;
                }
                if (footnoteMaps.Count(e => e.Indicator_ID == Indicator.Indicator_ID) > 3)
                {
                    Indicator.Footnote_ID_4 = footnoteMaps.Where(e => e.Indicator_ID == Indicator.Indicator_ID).OrderBy(e => e.Map_ID).Skip(3).First().Footnote_ID;
                    Indicator.Map_ID_4 = footnoteMaps.Where(e => e.Indicator_ID == Indicator.Indicator_ID).OrderBy(e => e.Map_ID).Skip(3).First().Map_ID;
                    Indicator.Footnote_Symbol_4 = db.Footnotes.FirstOrDefault(e => e.Footnote_ID == Indicator.Footnote_ID_4).Footnote_Symbol;
                }
                if (footnoteMaps.Count(e => e.Indicator_ID == Indicator.Indicator_ID) > 4)
                {
                    Indicator.Footnote_ID_5 = footnoteMaps.Where(e => e.Indicator_ID == Indicator.Indicator_ID).OrderBy(e => e.Map_ID).Skip(4).First().Footnote_ID;
                    Indicator.Map_ID_5 = footnoteMaps.Where(e => e.Indicator_ID == Indicator.Indicator_ID).OrderBy(e => e.Map_ID).Skip(4).First().Map_ID;
                    Indicator.Footnote_Symbol_5 = db.Footnotes.FirstOrDefault(e => e.Footnote_ID == Indicator.Footnote_ID_5).Footnote_Symbol;
                }
                if (footnoteMaps.Count(e => e.Indicator_ID == Indicator.Indicator_ID) > 5)
                {
                    Indicator.Footnote_ID_6 = footnoteMaps.Where(e => e.Indicator_ID == Indicator.Indicator_ID).OrderBy(e => e.Map_ID).Skip(5).First().Footnote_ID;
                    Indicator.Map_ID_6 = footnoteMaps.Where(e => e.Indicator_ID == Indicator.Indicator_ID).OrderBy(e => e.Map_ID).Skip(5).First().Map_ID;
                    Indicator.Footnote_Symbol_6 = db.Footnotes.FirstOrDefault(e => e.Footnote_ID == Indicator.Footnote_ID_6).Footnote_Symbol;
                }
                if (footnoteMaps.Count(e => e.Indicator_ID == Indicator.Indicator_ID) > 6)
                {
                    Indicator.Footnote_ID_7 = footnoteMaps.Where(e => e.Indicator_ID == Indicator.Indicator_ID).OrderBy(e => e.Map_ID).Skip(6).First().Footnote_ID;
                    Indicator.Map_ID_7 = footnoteMaps.Where(e => e.Indicator_ID == Indicator.Indicator_ID).OrderBy(e => e.Map_ID).Skip(6).First().Map_ID;
                    Indicator.Footnote_Symbol_7 = db.Footnotes.FirstOrDefault(e => e.Footnote_ID == Indicator.Footnote_ID_7).Footnote_Symbol;
                }
                if (footnoteMaps.Count(e => e.Indicator_ID == Indicator.Indicator_ID) > 7)
                {
                    Indicator.Footnote_ID_8 = footnoteMaps.Where(e => e.Indicator_ID == Indicator.Indicator_ID).OrderBy(e => e.Map_ID).Skip(7).First().Footnote_ID;
                    Indicator.Map_ID_8 = footnoteMaps.Where(e => e.Indicator_ID == Indicator.Indicator_ID).OrderBy(e => e.Map_ID).Skip(7).First().Map_ID;
                    Indicator.Footnote_Symbol_8 = db.Footnotes.FirstOrDefault(e => e.Footnote_ID == Indicator.Footnote_ID_8).Footnote_Symbol;
                }
                if (footnoteMaps.Count(e => e.Indicator_ID == Indicator.Indicator_ID) > 8)
                {
                    Indicator.Footnote_ID_9 = footnoteMaps.Where(e => e.Indicator_ID == Indicator.Indicator_ID).OrderBy(e => e.Map_ID).Skip(8).First().Footnote_ID;
                    Indicator.Map_ID_9 = footnoteMaps.Where(e => e.Indicator_ID == Indicator.Indicator_ID).OrderBy(e => e.Map_ID).Skip(8).First().Map_ID;
                    Indicator.Footnote_Symbol_9 = db.Footnotes.FirstOrDefault(e => e.Footnote_ID == Indicator.Footnote_ID_9).Footnote_Symbol;
                }
                if (footnoteMaps.Count(e => e.Indicator_ID == Indicator.Indicator_ID) > 9)
                {
                    Indicator.Footnote_ID_10 = footnoteMaps.Where(e => e.Indicator_ID == Indicator.Indicator_ID).OrderBy(e => e.Map_ID).Skip(9).First().Footnote_ID;
                    Indicator.Map_ID_10 = footnoteMaps.Where(e => e.Indicator_ID == Indicator.Indicator_ID).OrderBy(e => e.Map_ID).Skip(9).First().Map_ID;
                    Indicator.Footnote_Symbol_10 = db.Footnotes.FirstOrDefault(e => e.Footnote_ID == Indicator.Footnote_ID_10).Footnote_Symbol;
                }
            }

            return View(viewModel);
        }

        [HttpPost]
        public ActionResult editFootnoteMaps(IList<Indicator_Footnote_MapsViewModel> newMapsViewModel)
        {

            var newMaps = new Indicator_Footnote_Maps();
            newMaps.Indicator_ID = newMapsViewModel.FirstOrDefault().Indicator_ID;
            newMaps.Footnote_ID = newMapsViewModel.FirstOrDefault().Footnote_Symbol_1 == null ? null : db.Footnotes.ToList().FirstOrDefault(x => x.Footnote_Symbol == newMapsViewModel.FirstOrDefault().Footnote_Symbol_1).Footnote_ID;
            newMaps.Map_ID = newMapsViewModel.FirstOrDefault().Map_ID_1;

            var mapID = newMaps.Map_ID;
            var footnoteID = newMaps.Footnote_ID;
            if (footnoteID == null)
            {
                var deleteMap = db.Indicator_Footnote_Maps.Find(newMaps.Map_ID);
                if (deleteMap != null)
                {
                    db.Indicator_Footnote_Maps.Remove(deleteMap);
                    db.SaveChanges();
                }
                return View();
            }
            else if (db.Indicator_Footnote_Maps.Any(x => x.Map_ID == mapID))
            {
                if (ModelState.IsValid && db.Footnotes.Any(x => x.Footnote_ID == footnoteID))
                {
                    db.Entry(newMaps).State = EntityState.Modified;
                    db.SaveChanges();
                    return View();
                }
                else
                {
                    var oldMap = db.Indicator_Footnote_Maps.Find(newMaps.Map_ID);
                    var viewModel = new
                    {
                        Map_ID = oldMap.Map_ID,
                        Footnote_ID = oldMap.Footnote_ID,
                        Indicator_ID = oldMap.Indicator_ID,
                        State = "InvalidChange"
                    };
                    return Json(viewModel, JsonRequestBehavior.AllowGet);
                }
            }
            else
            {
                if (ModelState.IsValid && db.Footnotes.Any(x => x.Footnote_ID == footnoteID))
                {
                    db.Indicator_Footnote_Maps.Add(newMaps);
                    db.SaveChanges();
                    var viewModel = new  
                    {
                        Map_ID = newMaps.Map_ID,
                        Footnote_ID  = newMaps.Footnote_ID ,
                        Indicator_ID = newMaps.Indicator_ID,
                        State = "NewID"
                    };
                    return Json(viewModel, JsonRequestBehavior.AllowGet);
                }
                else if (ModelState.IsValid && !db.Footnotes.Any(x => x.Footnote_ID == footnoteID))
                {
                    var viewModel = new
                    {
                        State = "InvalidAdd"
                    };
                    return Json(viewModel, JsonRequestBehavior.AllowGet);
                }
            }
            return View();
        }

        [HttpGet]
        public ActionResult getIndicatorList()
        {
            var viewModel = db.Indicators.Select(x => new IndicatorListViewModel
            {
                Indicator_ID = x.Indicator_ID,
                Indicator = x.Indicator,
            }).ToList();
            return Json(viewModel, JsonRequestBehavior.AllowGet);
        }

        [HttpGet]
        public ActionResult editInventory(String Indicator_ID_Filter)
        {
            var viewModelItems = db.Indicators.ToArray();
            //var viewModelItems = db.Indicators.Where(x => x.Area_ID.Equals(1)).Where(y => y.Indicator_CoE_Map.Any(x => x.CoE_ID.Equals(10) || x.CoE_ID.Equals(27) || x.CoE_ID.Equals(30) || x.CoE_ID.Equals(40) || x.CoE_ID.Equals(50))).ToArray();
            var viewModel = viewModelItems.OrderBy(x => x.Indicator_ID).Select(x => new InventoryViewModel
            {
                Indicator_ID = x.Indicator_ID,
                Area_ID = x.Area_ID,
                Indicator = x.Indicator,
                FY_10_11 = x.FY_10_11,
                FY_10_11_Sup = x.FY_10_11_Sup,
                FY_11_12 = x.FY_11_12,
                FY_11_12_Sup = x.FY_11_12_Sup,
                FY_12_13 = x.FY_12_13,
                FY_12_13_Sup = x.FY_12_13_Sup,
                FY_13_14_Q1 = x.FY_13_14_Q1,
                FY_13_14_Q1_Sup = x.FY_13_14_Q1_Sup,
                FY_13_14_Q2 = x.FY_13_14_Q2,
                FY_13_14_Q2_Sup = x.FY_13_14_Q2_Sup,
                FY_13_14_Q3 = x.FY_13_14_Q3,
                FY_13_14_Q3_Sup = x.FY_13_14_Q3_Sup,
                FY_13_14_Q4 = x.FY_13_14_Q4,
                FY_13_14_Q4_Sup = x.FY_13_14_Q4_Sup,
                FY_13_14_YTD = x.FY_13_14_YTD,
                FY_13_14_YTD_Sup = x.FY_13_14_YTD_Sup,
                Target = x.Target,
                Target_Sup = x.Target_Sup,
                Comparator = x.Comparator,
                Comparator_Sup = x.Comparator_Sup,
                Performance_Threshold = x.Performance_Threshold,
                Performance_Threshold_Sup = x.Performance_Threshold_Sup,

                Colour_ID = x.Colour_ID,
                Custom_YTD = x.Custom_YTD,
                Custom_Q1 = x.Custom_Q1,
                Custom_Q2 = x.Custom_Q2,
                Custom_Q3 = x.Custom_Q3,
                Custom_Q4 = x.Custom_Q4,

                Definition_Calculation = x.Definition_Calculation,
                Target_Rationale = x.Target_Rationale,
                Comparator_Source = x.Comparator_Source,

                Data_Source_MSH = x.Data_Source_MSH,
                Data_Source_Benchmark = x.Data_Source_Benchmark,
                OPEO_Lead = x.OPEO_Lead,

                Q1_Color = x.Q1_Color,
                Q2_Color = x.Q2_Color,
                Q3_Color = x.Q3_Color,
                Q4_Color = x.Q4_Color,
                YTD_Color = x.YTD_Color
            }).ToList();
            if (Request.IsAjaxRequest())
            {
                return Json(viewModel.Where(x => x.Indicator_ID.ToString().Contains(Indicator_ID_Filter == null ? "" : Indicator_ID_Filter)), JsonRequestBehavior.AllowGet);
            }
            else
            {
                return View(viewModel);
            }
            
        }

        [HttpPost]
        public ActionResult editInventory(IList<Indicators> indicatorChange)
        {
            var indicatorID = indicatorChange[0].Indicator_ID;
            if (db.Indicators.Any(x => x.Indicator_ID == indicatorID ))
            {
                if (ModelState.IsValid)
                {
                    db.Entry(indicatorChange[0]).State = EntityState.Modified;
                    db.SaveChanges();
                    return View();
                }
                return View();
            } 
            else
            {
                if (ModelState.IsValid)
                {
                    db.Indicators.Add(indicatorChange[0]);
                    db.SaveChanges();
                    return View();
                }
                return View();
            }

        }

        //
        // GET: /Indicator/Details/5

        public ActionResult Details(string Indicator_ID)
        {
            Indicators indicators = db.Indicators.Find(Indicator_ID);
            if (indicators == null)
            {
                return HttpNotFound();
            }
            return View(indicators);
        }

        //
        // GET: /Indicator/Create

        public ActionResult Create()
        {
            return View();
        }

        //
        // POST: /Indicator/Create

        [HttpPost]
        public ActionResult Create(Indicators indicators)
        {
            if (ModelState.IsValid)
            {
                db.Indicators.Add(indicators);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(indicators);
        }

        //
        // GET: /Indicator/Edit/5

        public ActionResult Edit(string Indicator_ID)
        {
            Indicators indicators = db.Indicators.Find(Indicator_ID);
            if (indicators == null)
            {
                return HttpNotFound();
            }
            return View(indicators);
        }

        //
        // POST: /Indicator/Edit/5

        [HttpPost, ValidateInput(false)]
        public ActionResult Edit(Indicators indicators)
        {
            if (ModelState.IsValid)
            {
                db.Entry(indicators).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(indicators);
        }

        //
        // GET: /Indicator/Delete/5

        public ActionResult Delete(string Indicator_ID)
        {
            Indicators indicators = db.Indicators.Find(Indicator_ID);
            if (indicators == null)
            {
                return HttpNotFound();
            }
            return View(indicators);
        }

        //
        // POST: /Indicator/Delete/5

        [HttpPost, ActionName("Delete")]
        public ActionResult DeleteConfirmed(string Indicator_ID)
        {
            Indicators indicators = db.Indicators.Find(Indicator_ID);
            db.Indicators.Remove(indicators);
            db.SaveChanges();
            return RedirectToAction("Index");
        }

        protected override void Dispose(bool disposing)
        {
            db.Dispose();
            base.Dispose(disposing);
        }
    }
}