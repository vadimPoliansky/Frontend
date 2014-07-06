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
using IndInv.Helpers;
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
            return RedirectToAction("viewPR", "Indicator", new { fiscalYear = 1 });
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
//                    allFootnotes = db.Footnotes.ToList()
                };
                return View(viewModel);
            }

            return View();
        }

        public ActionResult viewPR(Int16 fiscalYear, Int16? analystID)
        {
            var allMaps = new List<Indicator_CoE_Maps>();

            if (analystID.HasValue)
            {
                allMaps = db.Indicator_CoE_Maps.Where(x=>x.Indicator.Analyst_ID == analystID).ToList();
            }
            else
            {
                allMaps = db.Indicator_CoE_Maps.ToList();
            }

            ModelState.Clear();
            var viewModel = new PRViewModel
            {
                //allCoEs = db.CoEs.ToList(),
                allAnalysts = db.Analysts.ToList(),
                allCoEs = db.CoEs.ToList(),
                allMaps = allMaps,
                allFootnoteMaps = db.Indicator_Footnote_Maps.ToList(),
                Fiscal_Year = fiscalYear,
                Analyst_ID = analystID,
            };

            return View(viewModel);
        }

        public ActionResult viewPRExcel(Int16 fiscalYear, Int16? coeID)
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
            var prGreen = XLColor.FromArgb(0, 118, 53);
            var prYellow = XLColor.FromArgb(255, 192, 0);
            var prRed = XLColor.FromArgb(255, 0, 0);
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
            var fitAdjust = 1.325 * 0.1;
            List<int> fitAdjustableRows = new List<int>();

            var prFootnoteCharsNewLine = 125;
            var prObjectivesCharsNewLine = 226;

            var allCoes = new List<CoEs>();
            if (coeID != 0 && coeID != null)
            {
                allCoes = viewModel.allCoEs.Where(x => x.CoE_ID == coeID).ToList();
            }
            else
            {
                allCoes = viewModel.allCoEs.ToList();
            }

            foreach (var coe in allCoes)
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

                    string[,] columnHeaders = new string[0, 0];
                    if (ws.Name == wsPRName)
                    {
                        var prHeadder2Title = FiscalYear.FYStrFull("FY_", fiscalYear) + "Performance";
                        prHeadder2Title = prHeadder2Title.Replace("_", " ");
                        columnHeaders = new string[,]{
                            {"Number",""},
                            {"Indicator",""},
                            {FiscalYear.FYStrFull("FY_3", fiscalYear), ""},
                            {FiscalYear.FYStrFull("FY_2", fiscalYear),""},
                            {FiscalYear.FYStrFull("FY_1", fiscalYear),""},
                            {prHeadder2Title,"Q1"},
                            {prHeadder2Title,"Q2"},
                            {prHeadder2Title,"Q3"},
                            {prHeadder2Title,"Q4"},
                            {prHeadder2Title,"YTD"},
                            {FiscalYear.FYStrFull("FY_", fiscalYear) + "Target",""},
                            {FiscalYear.FYStrFull("FY_", fiscalYear) + "Performance_Threshold",""},
                            {FiscalYear.FYStrFull("FY_", fiscalYear) + "Comparator",""}
                        };
                    }
                    else if (ws.Name == wsDefName)
                    {
                        columnHeaders = new string[,]{
                            {"Number",""},
                            {"Indicator",""},
                            {FiscalYear.FYStrFull("FY_", fiscalYear) + "Definition_Calculation",""},
                            {FiscalYear.FYStrFull("FY_", fiscalYear) + "Target_Rationale",""},
                            {FiscalYear.FYStrFull("FY_", fiscalYear) + "Comparator_Source",""}
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
                            ws.Cell(currentRow + 1, currentCol).Value = columnField;
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
                    foreach (var areaMap in coe.Area_CoE_Map.Where(x=>x.Fiscal_Year == fiscalYear).OrderBy(x => x.Area.Sort))
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

                        foreach (var map in viewModel.allMaps.Where(x => x.Fiscal_Year == fiscalYear).Where(e => e.Indicator.Area.Equals(areaMap.Area)).Where(d => d.CoE.CoE.Contains(coe.CoE)).OrderBy(f => f.Number))
                        {
                            fitAdjustableRows.Add(currentRow);
                            currentCol = 1;

                            ws.Cell(currentRow, currentCol).Value = indicatorNumber;
                            indicatorNumber++;
                            ws.Cell(currentRow, currentCol).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            currentCol++;

                            int j = 0;
                            ws.Cell(currentRow, currentCol).Value = map.Indicator.Indicator;
                            foreach (var footnote in map.Indicator.Indicator_Footnote_Map.Where(x => x.Fiscal_Year == fiscalYear).Where(e => e.Indicator_ID == map.Indicator_ID).OrderBy(e => e.Indicator_ID))
                            {
                                if (!footnotes.Contains(footnote.Footnote)) { footnotes.Add(footnote.Footnote); }
                                if (j != 0)
                                {
                                    ws.Cell(currentRow, currentCol).RichText.AddText(",").VerticalAlignment = XLFontVerticalTextAlignmentValues.Superscript;
                                }
                                ws.Cell(currentRow, currentCol).RichText.AddText(footnote.Footnote.Footnote_Symbol).VerticalAlignment = XLFontVerticalTextAlignmentValues.Superscript;
                                j++;
                            }
                            ws.Cell(currentRow, currentCol).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                            currentCol++;

                            if (ws.Name == wsPRName)
                            {
                                var obj = map.Indicator;
                                var type = obj.GetType();
                                string[,] columnIndicators = new string[,]{
                                    {(string)type.GetProperty(FiscalYear.FYStrFull("FY_3",fiscalYear)).GetValue(obj,null),
                                     (string)type.GetProperty(FiscalYear.FYStrFull("FY_3",fiscalYear) + "_Sup").GetValue(obj,null),
                                     ""
                                    },
                                    {(string)type.GetProperty(FiscalYear.FYStrFull("FY_2",fiscalYear)).GetValue(obj,null),
                                     (string)type.GetProperty(FiscalYear.FYStrFull("FY_2",fiscalYear) + "_Sup").GetValue(obj,null),
                                     ""
                                    },
                                    {(string)type.GetProperty(FiscalYear.FYStrFull("FY_1",fiscalYear)).GetValue(obj,null),
                                     (string)type.GetProperty(FiscalYear.FYStrFull("FY_1",fiscalYear) + "_Sup").GetValue(obj,null),
                                     ""
                                    },
                                    {(string)type.GetProperty(FiscalYear.FYStrFull("FY_",fiscalYear) + "Q1").GetValue(obj,null),
                                     (string)type.GetProperty(FiscalYear.FYStrFull("FY_",fiscalYear) + "Q1_Sup").GetValue(obj,null),
                                     (string)type.GetProperty(FiscalYear.FYStrFull("FY_",fiscalYear) + "Q1_Color").GetValue(obj,null),
                                    },
                                    {(string)type.GetProperty(FiscalYear.FYStrFull("FY_",fiscalYear) + "Q2").GetValue(obj,null),
                                     (string)type.GetProperty(FiscalYear.FYStrFull("FY_",fiscalYear) + "Q2_Sup").GetValue(obj,null),
                                     (string)type.GetProperty(FiscalYear.FYStrFull("FY_",fiscalYear) + "Q2_Color").GetValue(obj,null),
                                    },
                                    {(string)type.GetProperty(FiscalYear.FYStrFull("FY_",fiscalYear) + "Q3").GetValue(obj,null),
                                     (string)type.GetProperty(FiscalYear.FYStrFull("FY_",fiscalYear) + "Q3_Sup").GetValue(obj,null),
                                     (string)type.GetProperty(FiscalYear.FYStrFull("FY_",fiscalYear) + "Q3_Color").GetValue(obj,null),
                                    },
                                    {(string)type.GetProperty(FiscalYear.FYStrFull("FY_",fiscalYear) + "Q4").GetValue(obj,null),
                                     (string)type.GetProperty(FiscalYear.FYStrFull("FY_",fiscalYear) + "Q4_Sup").GetValue(obj,null),
                                     (string)type.GetProperty(FiscalYear.FYStrFull("FY_",fiscalYear) + "Q4_Color").GetValue(obj,null),
                                    },
                                    {(string)type.GetProperty(FiscalYear.FYStrFull("FY_",fiscalYear) + "YTD").GetValue(obj,null),
                                     (string)type.GetProperty(FiscalYear.FYStrFull("FY_",fiscalYear) + "YTD_Sup").GetValue(obj,null),
                                     (string)type.GetProperty(FiscalYear.FYStrFull("FY_",fiscalYear) + "YTD_Color").GetValue(obj,null),
                                    },
                                    {(string)type.GetProperty(FiscalYear.FYStrFull("FY_",fiscalYear) + "Target").GetValue(obj,null),
                                     (string)type.GetProperty(FiscalYear.FYStrFull("FY_",fiscalYear) + "Target_Sup").GetValue(obj,null),
                                     ""
                                    },
                                    {(string)type.GetProperty(FiscalYear.FYStrFull("FY_",fiscalYear) + "Performance_Threshold").GetValue(obj,null),
                                     (string)type.GetProperty(FiscalYear.FYStrFull("FY_",fiscalYear) + "Performance_Threshold_Sup").GetValue(obj,null),
                                     ""
                                    },
                                    {(string)type.GetProperty(FiscalYear.FYStrFull("FY_",fiscalYear) + "Comparator").GetValue(obj,null),
                                     (string)type.GetProperty(FiscalYear.FYStrFull("FY_",fiscalYear) + "Comparator_Sup").GetValue(obj,null),
                                     ""
                                    },
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
                                var obj = map.Indicator;
                                var type = obj.GetType();
                                ws.Cell(currentRow, currentCol).Value = (string)type.GetProperty(FiscalYear.FYStrFull("FY_", fiscalYear) + "Definition_Calculation").GetValue(obj, null);
                                ws.Cell(currentRow, currentCol).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                                currentCol++;
                                ws.Cell(currentRow, currentCol).Value = (string)type.GetProperty(FiscalYear.FYStrFull("FY_", fiscalYear) + "Target_Rationale").GetValue(obj, null);
                                ws.Cell(currentRow, currentCol).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                                currentCol++;
                                ws.Cell(currentRow, currentCol).Value = (string)type.GetProperty(FiscalYear.FYStrFull("FY_", fiscalYear) + "Comparator_Source").GetValue(obj, null);
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

                    //var totalHeight = 0.0;
                    //foreach (var row in ws.Rows(1, ws.LastRowUsed().RowNumber()))
                    //{
                    //    totalHeight += row.Height;
                    //}
                    //var totalWidth = 0.0;
                    //foreach (var col in ws.Columns(1, ws.LastColumnUsed().ColumnNumber()))
                    //{
                    //    totalWidth += col.Width;
                    //}
                    //var totalRatio = totalHeight / totalWidth;
                    //if (totalRatio > fitRatio)
                    //{
                    //    var fitAddWidthTotal = (totalHeight - fitHeight);
                    //    var fitAddWidthPer = fitAddWidthTotal / (ws.LastRowUsed().RowNumber() - 1) / fitAdjust;
                    //    foreach (var col in ws.Columns().Skip(1))
                    //    {
                    //        col.Width += fitAddWidthPer;
                    //    }
                    //}
                    //else
                    //{
                    //    var fitAddHeightTotal = -(totalHeight - fitHeight);
                    //    var fitAddHeightPer = fitAddHeightTotal / ws.LastRowUsed().RowNumber() / fitAdjust;
                    //    foreach (var row in ws.Rows().Skip(3))
                    //    {
                    //        row.Height += fitAddHeightPer;
                    //    }
                    //}
                }
            }

            MemoryStream preImage = new MemoryStream();
            wb.SaveAs(preImage);

            MemoryStream postImage = new MemoryStream();
            SLDocument postImageWb = new SLDocument(preImage);

            string picPath = this.HttpContext.ApplicationInstance.Server.MapPath("~/App_Data/logo.png");
            SLPicture picLogo = new SLPicture(picPath);
            string picPathOPEO = this.HttpContext.ApplicationInstance.Server.MapPath("~/App_Data/logoOPEO.png");
            SLPicture picLogoOPEO = new SLPicture(picPathOPEO);
            string picMonthlyPath = this.HttpContext.ApplicationInstance.Server.MapPath("~/App_Data/Monthly.png");
            SLPicture picMonthly = new SLPicture(picMonthlyPath);
            string picQuaterlyPath = this.HttpContext.ApplicationInstance.Server.MapPath("~/App_Data/quaterly.png");
            SLPicture picQuaterly = new SLPicture(picQuaterlyPath);
            string picNAPath = this.HttpContext.ApplicationInstance.Server.MapPath("~/App_Data/na.png");
            SLPicture picNA = new SLPicture(picNAPath);
            string picTargetPath = this.HttpContext.ApplicationInstance.Server.MapPath("~/App_Data/target.png");
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

        public ActionResult editCoEMaps(Int16 fiscalYear)
        {
            var viewModel = new Indicator_CoE_MapsViewModel
            {
                allIndicators = db.Indicators.ToList(),
                allCoEs = db.CoEs.ToList(),
                allMaps = db.Indicator_CoE_Maps.Where(x=>x.Fiscal_Year == fiscalYear).ToList(),
                fiscalYear = fiscalYear,
            };
            return View(viewModel);
        }

        [HttpPost]
        public void editCoEMaps(IList<Indicator_CoE_MapsViewModel> mapChange)
        {
            Indicator_CoE_Maps map = mapChange.FirstOrDefault().allMaps.FirstOrDefault();
            if (map.Fiscal_Year == 0)
            {
                map.Fiscal_Year = db.Indicator_CoE_Maps.Max(x => x.Fiscal_Year);
            }

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
                if (Footnote_ID_Filter == "")
                {
                    var newFootnote = db.Footnotes.Create();
                    db.Footnotes.Add(newFootnote);
                    db.SaveChanges();

                    viewModel = new List<FootnotesViewModel>();
                    var newViewModelItem = new FootnotesViewModel {
                        Footnote_ID = newFootnote.Footnote_ID,
                        Footnote = newFootnote.Footnote,
                        Footnote_Symbol = newFootnote.Footnote_Symbol,
                    };
                    viewModel.Add(newViewModelItem);

                    return Json(viewModel, JsonRequestBehavior.AllowGet);
                }
                else
                {
                    return Json(viewModel.Where(x => x.Footnote_ID.ToString().Contains(Footnote_ID_Filter == null ? "" : Footnote_ID_Filter)), JsonRequestBehavior.AllowGet);
                }
            }
            else
            {
                return View(viewModel);
            }

        }

        [HttpPost]
        public void deleteFootnotes(Int16 footnoteID)
        {
            var deleteFootnote = db.Footnotes.FirstOrDefault(x => x.Footnote_ID == footnoteID);
            db.Footnotes.Remove(deleteFootnote);
            db.SaveChanges();
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

        [HttpGet]
        public ActionResult editAnalysts(String Analyst_ID_Filter)
        {
            var viewModelItems = db.Analysts.ToArray();
            var viewModel = viewModelItems.OrderBy(x => x.Analyst_ID).Select(x => new AnalystViewModel
            {
                Analyst_ID = x.Analyst_ID,
                First_Name = x.First_Name,
                Last_Name = x.Last_Name,
                Position = x.Position,
                Order = x.Order,
            }).ToList();
            if (Request.IsAjaxRequest())
            {
                if (Analyst_ID_Filter == "")
                {
                    var newAnalyst = db.Analysts.Create();
                    db.Analysts.Add(newAnalyst);
                    db.SaveChanges();

                    viewModel = new List<AnalystViewModel>();
                    var newViewModelItem = new AnalystViewModel
                    {
                        Analyst_ID = newAnalyst.Analyst_ID,
                        First_Name = newAnalyst.First_Name,
                        Last_Name = newAnalyst.Last_Name,
                        Position = newAnalyst.Position,
                        Order = newAnalyst.Order,
                    };
                    viewModel.Add(newViewModelItem);

                    return Json(viewModel, JsonRequestBehavior.AllowGet);
                }
                else
                {
                    return Json(viewModel.Where(x => x.Analyst_ID.ToString().Contains(Analyst_ID_Filter == null ? "" : Analyst_ID_Filter)), JsonRequestBehavior.AllowGet);
                }
            }
            else
            {
                return View(viewModel);
            }

        }

        [HttpPost]
        public void deleteAnalyst(Int16 analystID)
        {
            var deleteAnalyst = db.Analysts.FirstOrDefault(x => x.Analyst_ID == analystID);
            db.Analysts.Remove(deleteAnalyst);
            db.SaveChanges();
        }

        [HttpPost]
        public ActionResult editAnalysts(IList<Analysts> analystChange)
        {
            var analystID = analystChange[0].Analyst_ID;
            if (db.Analysts.Any(x => x.Analyst_ID == analystID))
            {
                if (ModelState.IsValid)
                {
                    db.Entry(analystChange[0]).State = EntityState.Modified;
                    db.SaveChanges();
                    return View();
                }
                return View();
            }
            else
            {
                if (ModelState.IsValid)
                {
                    db.Analysts.Add(analystChange[0]);
                    db.SaveChanges();
                    return View();
                }
                return View();
            }

        }

        [HttpGet]
        public JsonResult getAreaMap (Int16 mapID, Int16 fiscalYear)
        {
            var objectives = db.Area_CoE_Maps.Where(x=>x.Fiscal_Year == fiscalYear).FirstOrDefault(x => x.Map_ID == mapID).Objective;
            return Json(objectives, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public void setAreaMap (Int16 mapID, string objective, Int16 fiscalYear)
        {
            var map = db.Area_CoE_Maps.FirstOrDefault(x => x.Map_ID == mapID);
            map.Objective = objective;

            if (ModelState.IsValid)
            {
                db.Entry(map).State = EntityState.Modified;
                db.SaveChanges();
            }
        }

        public ActionResult editFootnoteMaps(Int16 fiscalYear,  string indicatorID)
        {
            List<Indicator_Footnote_Maps> footnoteMaps = new List<Indicator_Footnote_Maps>();
            foreach (var footnote in db.Indicator_Footnote_Maps.Where(x=>x.Fiscal_Year == fiscalYear).OrderBy(e => e.Map_ID).ToList())
            {
                footnoteMaps.Add(footnote);
            }

            var allIndicator = new List<Indicators>();
            if (indicatorID != null )
            {
                allIndicator = db.Indicators.Where(x => x.Indicator_ID == indicatorID).OrderBy(x=>x.Indicator_ID).ToList();
            }
            else
            {
                allIndicator = db.Indicators.OrderBy(x => x.Indicator_ID).ToList();
            }

            var viewModel = allIndicator.Select(x => new Indicator_Footnote_MapsViewModel
            {
                Indicator_ID = x.Indicator_ID,
                Indicator= x.Indicator,
                Fiscal_Year = fiscalYear,
            }).ToList();

            viewModel.FirstOrDefault().allFootnotes = new List<string>();

            viewModel.FirstOrDefault().allFootnotes.AddRange(db.Footnotes.Select(x => x.Footnote_Symbol + ", " + x.Footnote).ToList());

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
            newMaps.Footnote_ID = newMapsViewModel.FirstOrDefault().Footnote_Symbol_1 == null ? (Int16)0 : db.Footnotes.ToList().FirstOrDefault(x => x.Footnote_Symbol == newMapsViewModel.FirstOrDefault().Footnote_Symbol_1).Footnote_ID;
            newMaps.Map_ID = newMapsViewModel.FirstOrDefault().Map_ID_1;
            newMaps.Fiscal_Year = newMapsViewModel.FirstOrDefault().Fiscal_Year;

            var mapID = newMaps.Map_ID;
            var footnoteID = newMaps.Footnote_ID;
            if (footnoteID == 0)
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
        public ActionResult getValue(string indicatorID, string field, Int16 fiscalYear)
        {
            var indicator = db.Indicators.FirstOrDefault(x => x.Indicator_ID == indicatorID);

            var property = indicator.GetType().GetProperty(field);
            var propertySup = indicator.GetType().GetProperty(field + "_Sup");
            var value = (string)property.GetValue(indicator, null);
            var valueSup = "";
            if (propertySup != null)
            {
                valueSup = (string)propertySup.GetValue(indicator, null);
            }
            else
            {
                string footnoteStr = "";
                var allFootnotes = db.Footnotes.ToList();
                int j = 0;
                foreach (var footnote in db.Indicator_Footnote_Maps.Where(x => x.Fiscal_Year == fiscalYear).Where(e => e.Indicator_ID == indicatorID).OrderBy(e => e.Indicator_ID))
                {
                    if (j != 0) { footnoteStr += ","; }
                    footnoteStr += allFootnotes.FirstOrDefault(x => x.Footnote_ID == footnote.Footnote_ID).Footnote_Symbol;
                    j++;
                }
                valueSup = footnoteStr;
            }

            var viewModel = new valueViewModel()
            {
                Value = value,
                Value_Sup = valueSup,
            };

            return Json(viewModel, JsonRequestBehavior.AllowGet) ;
        }

        [HttpPost]
        public JsonResult setValue(string indicatorID, string updateProperty, string updateValue, string updateValueSup, int fiscalYear)
        {
            var indicator = db.Indicators.FirstOrDefault(x => x.Indicator_ID == indicatorID);

            var type = indicator.GetType();
            var property = type.GetProperty(updateProperty);
            property.SetValue(indicator, Convert.ChangeType(updateValue, property.PropertyType), null);

            if (updateValueSup != "%NULL%")
            {
                var propertySup = indicator.GetType().GetProperty(updateProperty + "_Sup");
                if (propertySup != null)
                {
                    propertySup.SetValue(indicator, Convert.ChangeType(updateValueSup, property.PropertyType), null);
                }
            }

            if (ModelState.IsValid)
            {
                db.Entry(indicator).State = EntityState.Modified;
                db.SaveChanges();
            }

            var propertyColor = type.GetProperty(updateProperty + "_Color");
            if (propertyColor != null)
            {
                var color = propertyColor.GetValue(indicator,null);
                return Json(color, JsonRequestBehavior.AllowGet);
            }
            else
            {
                return Json("", JsonRequestBehavior.AllowGet);
            }
            //var indicatorID = indicatorChange[0].Indicator_ID;
            //if (db.Indicators.Any(x => x.Indicator_ID == indicatorID ))
            //{
            //    if (ModelState.IsValid)
            //    {
            //        db.Entry(indicatorChange[0]).State = EntityState.Modified;
            //        db.SaveChanges();
            //    }
            //} 
            //else
            //{
            //    if (ModelState.IsValid)
            //    {
            //        db.Indicators.Add(indicatorChange[0]);
            //        db.SaveChanges();
            //    }
            //}

        }

        [HttpGet]
        public ActionResult editInventory(String indicatorID, Int16? analystID, int fiscalYear)
        {
            //}
            //var viewModelItems = db.Indicators.Where(x => x.Area_ID.Equals(1)).Where(y => y.Indicator_CoE_Map.Any(x => x.CoE_ID.Equals(10) || x.CoE_ID.Equals(27) || x.CoE_ID.Equals(30) || x.CoE_ID.Equals(40) || x.CoE_ID.Equals(50))).ToArray();

            var viewModelItems = new List<Indicators>();
            if (analystID.HasValue)
            {
                viewModelItems = db.Indicators.Where(x => x.Analyst_ID == analystID).ToList();
            }
            else
            {
                viewModelItems = db.Indicators.ToList();
            }

            var viewModel = viewModelItems.OrderBy(x => x.Indicator_ID).Select(x => new InventoryViewModel
            {
                Indicator_ID = x.Indicator_ID,
                Area_ID = x.Area_ID,
                CoE = x.Indicator_CoE_Map.Count != 0 ? x.Indicator_CoE_Map.Where(y => y.Fiscal_Year == fiscalYear).FirstOrDefault().CoE.CoE : "",
                Indicator = x.Indicator,
                FY_3 = (string)x.GetType().GetProperty(FiscalYear.FYStr(fiscalYear, 3) + "_YTD").GetValue(x, null),
                FY_3_Sup = (string)x.GetType().GetProperty(FiscalYear.FYStr(fiscalYear, 3) + "_YTD_Sup").GetValue(x, null),
                FY_2 = (string)x.GetType().GetProperty(FiscalYear.FYStr(fiscalYear, 2) + "_YTD").GetValue(x, null),
                FY_2_Sup = (string)x.GetType().GetProperty(FiscalYear.FYStr(fiscalYear, 2) + "_YTD_Sup").GetValue(x, null),
                FY_1 = (string)x.GetType().GetProperty(FiscalYear.FYStr(fiscalYear, 1) + "_YTD").GetValue(x, null),
                FY_1_Sup = (string)x.GetType().GetProperty(FiscalYear.FYStr(fiscalYear, 1) + "_YTD_Sup").GetValue(x, null),
                FY_Q1 = (string)x.GetType().GetProperty(FiscalYear.FYStr(fiscalYear, 0) + "_Q1").GetValue(x, null),
                FY_Q1_Sup = (string)x.GetType().GetProperty(FiscalYear.FYStr(fiscalYear, 0) + "_Q1_Sup").GetValue(x, null),
                FY_Q2 = (string)x.GetType().GetProperty(FiscalYear.FYStr(fiscalYear, 0) + "_Q2").GetValue(x, null),
                FY_Q2_Sup = (string)x.GetType().GetProperty(FiscalYear.FYStr(fiscalYear, 0) + "_Q2_Sup").GetValue(x, null),
                FY_Q3 = (string)x.GetType().GetProperty(FiscalYear.FYStr(fiscalYear, 0) + "_Q3").GetValue(x, null),
                FY_Q3_Sup = (string)x.GetType().GetProperty(FiscalYear.FYStr(fiscalYear, 0) + "_Q3_Sup").GetValue(x, null),
                FY_Q4 = (string)x.GetType().GetProperty(FiscalYear.FYStr(fiscalYear, 0) + "_Q4").GetValue(x, null),
                FY_Q4_Sup = (string)x.GetType().GetProperty(FiscalYear.FYStr(fiscalYear, 0) + "_Q4_Sup").GetValue(x, null),
                FY_YTD = (string)x.GetType().GetProperty(FiscalYear.FYStr(fiscalYear, 0) + "_YTD").GetValue(x, null),
                FY_YTD_Sup = (string)x.GetType().GetProperty(FiscalYear.FYStr(fiscalYear, 0) + "_YTD_Sup").GetValue(x, null),
                Target = (string)x.GetType().GetProperty(FiscalYear.FYStr(fiscalYear, 0) + "_Target").GetValue(x, null),
                Target_Sup = (string)x.GetType().GetProperty(FiscalYear.FYStr(fiscalYear, 0) + "_Target_Sup").GetValue(x, null),
                Comparator = (string)x.GetType().GetProperty(FiscalYear.FYStr(fiscalYear, 0) + "_Comparator").GetValue(x, null),
                Comparator_Sup = (string)x.GetType().GetProperty(FiscalYear.FYStr(fiscalYear, 0) + "_Comparator_Sup").GetValue(x, null),
                Performance_Threshold = (string)x.GetType().GetProperty(FiscalYear.FYStr(fiscalYear, 0) + "_Performance_Threshold").GetValue(x, null),
                Performance_Threshold_Sup = (string)x.GetType().GetProperty(FiscalYear.FYStr(fiscalYear, 0) + "_Performance_Threshold_Sup").GetValue(x, null),

                Colour_ID = (Int16)x.GetType().GetProperty(FiscalYear.FYStr(fiscalYear, 0) + "_Colour_ID").GetValue(x, null),
                Custom_YTD = (string)x.GetType().GetProperty(FiscalYear.FYStr(fiscalYear, 0) + "_Custom_YTD").GetValue(x, null),
                Custom_Q1 = (string)x.GetType().GetProperty(FiscalYear.FYStr(fiscalYear, 0) + "_Custom_Q1").GetValue(x, null),
                Custom_Q2 = (string)x.GetType().GetProperty(FiscalYear.FYStr(fiscalYear, 0) + "_Custom_Q2").GetValue(x, null),
                Custom_Q3 = (string)x.GetType().GetProperty(FiscalYear.FYStr(fiscalYear, 0) + "_Custom_Q3").GetValue(x, null),
                Custom_Q4 = (string)x.GetType().GetProperty(FiscalYear.FYStr(fiscalYear, 0) + "_Custom_Q4").GetValue(x, null),

                Definition_Calculation = (string)x.GetType().GetProperty(FiscalYear.FYStr(fiscalYear, 0) + "_Definition_Calculation").GetValue(x, null),
                Target_Rationale = (string)x.GetType().GetProperty(FiscalYear.FYStr(fiscalYear, 0) + "_Target_Rationale").GetValue(x, null),
                Comparator_Source = (string)x.GetType().GetProperty(FiscalYear.FYStr(fiscalYear, 0) + "_Comparator_Source").GetValue(x, null),

                Data_Source_MSH = (string)x.GetType().GetProperty(FiscalYear.FYStr(fiscalYear, 0) + "_Data_Source_MSH").GetValue(x, null),
                Data_Source_Benchmark = (string)x.GetType().GetProperty(FiscalYear.FYStr(fiscalYear, 0) + "_Data_Source_Benchmark").GetValue(x, null),
                OPEO_Lead = (string)x.GetType().GetProperty(FiscalYear.FYStr(fiscalYear, 0) + "_OPEO_Lead").GetValue(x, null),

                Q1_Color = (string)x.GetType().GetProperty(FiscalYear.FYStr(fiscalYear, 0) + "_Q1_Color").GetValue(x, null),
                Q2_Color = (string)x.GetType().GetProperty(FiscalYear.FYStr(fiscalYear, 0) + "_Q2_Color").GetValue(x, null),
                Q3_Color = (string)x.GetType().GetProperty(FiscalYear.FYStr(fiscalYear, 0) + "_Q3_Color").GetValue(x, null),
                Q4_Color = (string)x.GetType().GetProperty(FiscalYear.FYStr(fiscalYear, 0) + "_Q4_Color").GetValue(x, null),
                YTD_Color = (string)x.GetType().GetProperty(FiscalYear.FYStr(fiscalYear, 0) + "_YTD_Color").GetValue(x, null),

                Fiscal_Year = fiscalYear,

            }).ToList();
            if (viewModel.Count == 0)
            {
                viewModel.Add(new InventoryViewModel());
            }
            viewModel.FirstOrDefault().allAnalysts = db.Analysts.ToList();
            if (Request.IsAjaxRequest())
            {
                return Json(viewModel.Where(x => x.Indicator_ID.ToString().Contains(indicatorID == null ? "" : indicatorID)), JsonRequestBehavior.AllowGet);
            }
            else
            {
                return View(viewModel);
            }
            
        }

        [HttpPost]
        public void editInventory(string indicatorID, string updateProperty, string updateValue, int fiscalYear)
        {
            var updatePropertyFull = updateProperty;
            if (fiscalYear != 0)
            {
                updatePropertyFull = FiscalYear.FYStrFull(updateProperty, fiscalYear);
            }

            var indicator = db.Indicators.FirstOrDefault(x => x.Indicator_ID == indicatorID);

            if (indicator == null)
            {
                indicator = db.Indicators.Create();
                indicator.Indicator_ID = updateValue;
                db.Indicators.Add(indicator);
                db.SaveChanges();
            } else{
                var property = indicator.GetType().GetProperty(updatePropertyFull);
                property.SetValue(indicator, Convert.ChangeType(updateValue, property.PropertyType), null);

                if (ModelState.IsValid)
                {
                    db.Entry(indicator).State = EntityState.Modified;
                    db.SaveChanges();
                }
            }
            //var indicatorID = indicatorChange[0].Indicator_ID;
            //if (db.Indicators.Any(x => x.Indicator_ID == indicatorID ))
            //{
            //    if (ModelState.IsValid)
            //    {
            //        db.Entry(indicatorChange[0]).State = EntityState.Modified;
            //        db.SaveChanges();
            //    }
            //} 
            //else
            //{
            //    if (ModelState.IsValid)
            //    {
            //        db.Indicators.Add(indicatorChange[0]);
            //        db.SaveChanges();
            //    }
            //}

        }

        //
        // GET: /Indicator/Details/5

        [HttpPost]
        public JsonResult newIndicatorAtPR(Int16 fiscalYear, Int16 areaID, Int16 coeID, Int16 number)
        {
            var newIndicator = new Indicators();
            newIndicator.Area_ID = areaID;
            var coeNumLow = (coeID * 100).ToString();
            var coeNumHigh = (coeID * 100 + 100).ToString();
            var id = db.Indicators.Where(x => x.Indicator_ID >= coeNum && Int32.Parse(x.Indicator_ID) <= coeNum + 100).Max(x => x.Indicator_ID);
            newIndicator.Indicator_ID = id;
            db.Indicators.Add(newIndicator);
            db.SaveChanges();
            var newMap = new Indicator_CoE_Maps();
            newMap.Indicator_ID = newIndicator.Indicator_ID;
            newMap.CoE_ID = coeID;
            newMap.Fiscal_Year = fiscalYear;
            db.Indicator_CoE_Maps.Add(newMap);
            db.SaveChanges();

            return Json(newMap.Map_ID, JsonRequestBehavior.AllowGet);
        }

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

        public ActionResult edit(string indicatorID)
        {
            Indicators indicator = db.Indicators.Find(indicatorID);

            if (indicator == null)
            {
                indicator = db.Indicators.Create();
            }

            var viewModel = new editViewModel
            {
                Indicator = indicator,
                allCoEs = db.CoEs.ToList(),
            };

            return View(viewModel);
        }

        //
        // POST: /Indicator/Edit/5

        [HttpPost, ValidateInput(false)]
        public ActionResult edit(Indicators indicators)
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

        public ActionResult Delete(string indicatorID)
        {
            Indicators indicators = db.Indicators.Find(indicatorID);
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