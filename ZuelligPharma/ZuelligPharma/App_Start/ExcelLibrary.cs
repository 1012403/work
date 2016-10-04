using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Runtime.InteropServices;
using ZuelligPharma.Models;

namespace ZuelligPharma.App_Start
{
    public class ExcelLibrary
    {
        /// <summary>
        /// appliation excel mo file excel
        /// </summary>
        public Excel.Application xlApp;

        /// <summary>
        /// Workbook hien tai dang su dung
        /// </summary>
        public Excel.Workbook xlWorkBook;

        /// <summary>
        /// Sheet hien tai dang sua dung
        /// </summary>
        public Excel._Worksheet xlSheet;

        /// <summary>
        /// Constructor mo file excel
        /// </summary>
        /// <param name="path"></param>
        public ExcelLibrary(string path)
        {
            try
            {
                this.xlApp = new Excel.Application();
                this.xlWorkBook = xlApp.Workbooks.Open(path);
            }
            catch (Exception ex)
            {
                if (this.xlApp != null)
                {
                    this.xlApp.Quit();
                    this.xlApp = null;
                }
                throw ex;
            }
            finally
            {

            }
        }
        public void ReadData()
        {
            foreach (var obj in this.xlWorkBook.Worksheets)
            {
                Excel.Range currentFind = null;
                Excel.Range nextFind = null;
                Excel._Worksheet sheet = (Excel._Worksheet)obj;
                if (sheet.Name == @"MAT.")
                {
                    List<ZuelligPharma_MAT> ZuelligPharma_MATs = new List<ZuelligPharma_MAT>();
                    this.xlSheet = sheet;
                    Excel.Range xlRange = this.xlSheet.UsedRange;
                    Excel.Range curCell = null;
                    currentFind = xlRange.Find("Gros MAT");
                    int seqno = 0;
                    object _temp = null;
                    string date = String.Empty;
                    double gros = 0;
                    double net = 0;
                    double sale = 0;
                    for (int i = currentFind.Row + 1; i <= xlRange.Rows.Count; i++)
                    {
                        seqno++;// get seqno for rowinsert
                        var adddt = (string)DateTime.Now.ToString("yyyyMMdd");
                        curCell = (Excel.Range)this.xlSheet.Cells[i, currentFind.Column];
                        _temp = curCell.Value;
                        if (_temp == null || String.IsNullOrEmpty(_temp.ToString()) == true)
                        {
                            break;
                        }
                        else
                        {
                            DateTime _temp2 = Convert.ToDateTime(_temp.ToString());
                            // get date
                            date = (string)_temp2.ToString("MMM-yy");

                            //get gros
                            curCell = (Excel.Range)this.xlSheet.Cells[i, currentFind.Column+1];
                            _temp = curCell.Value;
                            gros = Math.Round(Convert.ToDouble(_temp.ToString()), 0, MidpointRounding.AwayFromZero);
                            
                            // get net
                            curCell = (Excel.Range)this.xlSheet.Cells[i, currentFind.Column + 9];
                            _temp = curCell.Value;
                            net = Math.Round(Convert.ToDouble(_temp.ToString()), 0, MidpointRounding.AwayFromZero);

                            // get sale
                            curCell = (Excel.Range)this.xlSheet.Cells[i, currentFind.Column - 2];
                            _temp = curCell.Value;
                            sale = Math.Round(Convert.ToDouble(_temp.ToString()), 0, MidpointRounding.AwayFromZero);

                            ZuelligPharma_MATs.Add(new ZuelligPharma_MAT()
                            {
                                adddt = (string)DateTime.Now.ToString("yyyyMMdd"),
                                seqno = (String)seqno.ToString().PadLeft(6, '0'),
                                date = date,
                                gros = gros,
                                net = net,
                                sale = sale,
                                timestamp = DateTime.Now.ToString("yyyyMMddHHmmss")
                            });
                        }
                    }
                }
                //else if (sheet.Name == @"Calculated")
                //{
                //    List<ZuelligPharma_Calculated> ZuelligPharma_Calculateds = new List<ZuelligPharma_Calculated>();
                //    this.xlSheet = sheet;
                //    Excel.Range xlRange = this.xlSheet.UsedRange;
                //    Excel.Range curCell = null;
                //    currentFind = xlRange.Find("Gros MAT");
                //    int seqno = 0;
                //    object _temp = null;
                //    double gros = 0;
                //    double net = 0;
                //    double sale = 0;
                //    for (int i = currentFind.Row + 1; i <= xlRange.Rows.Count; i++)
                //    {
                //        seqno++;// get seqno for rowinsert
                //        var adddt = (string)DateTime.Now.ToString("yyyyMMdd");
                //        curCell = (Excel.Range)this.xlSheet.Cells[i, currentFind.Column];
                //        _temp = curCell.Value;
                //        if (_temp == null || String.IsNullOrEmpty(_temp.ToString()) == true)
                //        {
                //            break;
                //        }
                //        else
                //        {
                //            DateTime _temp2 = Convert.ToDateTime(_temp.ToString());
                //            // get date
                //            string date = (string)_temp2.ToString("MMM-yy");

                //            //get gros
                //            curCell = (Excel.Range)this.xlSheet.Cells[i, currentFind.Column + 1];
                //            _temp = curCell.Value;
                //            gros = Math.Round(Convert.ToDouble(_temp.ToString()), 0, MidpointRounding.AwayFromZero);

                //            // get net
                //            curCell = (Excel.Range)this.xlSheet.Cells[i, currentFind.Column + 9];
                //            _temp = curCell.Value;
                //            net = Math.Round(Convert.ToDouble(_temp.ToString()), 0, MidpointRounding.AwayFromZero);

                //            // get sale
                //            curCell = (Excel.Range)this.xlSheet.Cells[i, currentFind.Column - 2];
                //            _temp = curCell.Value;
                //            sale = Math.Round(Convert.ToDouble(_temp.ToString()), 0, MidpointRounding.AwayFromZero);

                //            ZuelligPharma_MATs.Add(new ZuelligPharma_MAT()
                //            {
                //                adddt = (string)DateTime.Now.ToString("yyyyMMdd"),
                //                seqno = (String)seqno.ToString().PadLeft(6, '0'),
                //                date = date,
                //                gros = gros,
                //                net = net,
                //                sale = sale,
                //                timestamp = DateTime.Now.ToString("yyyyMMddHHmmss")
                //            });
                //        }
                //    }
                //}
                else if(sheet.Name == @"Frequency")
                {
                    List<ZuelligPharma_Frequency> ZuelligPharma_Frequencies = new List<ZuelligPharma_Frequency>();
                    this.xlSheet = sheet;
                    Excel.Range xlRange = this.xlSheet.UsedRange;
                    Excel.Range curCell = null;
                    currentFind = xlRange.Find("Frequency");
                    int seqno = 0;
                    object _temp = null;
                    string freqno = String.Empty;
                    int numofcust = 0;
                    double percentofcust = 0;

                    for (int i = currentFind.Column + 1; i <= 12; i++)
                    {
                        seqno++;// get seqno for rowinsert
                        var adddt = (string)DateTime.Now.ToString("yyyyMMdd");
                        curCell = (Excel.Range)this.xlSheet.Cells[currentFind.Row, i];
                        _temp = curCell.Value;
                        if (_temp == null || String.IsNullOrEmpty(_temp.ToString()) == true)
                        {
                            break;
                        }
                        else
                        {
                            // get freqno
                            freqno = _temp.ToString();

                            //get numofcust
                            curCell = (Excel.Range)this.xlSheet.Cells[currentFind.Row + 1, i];
                            _temp = curCell.Value;
                            numofcust = Convert.ToInt16(_temp.ToString());

                            // get percentofcust
                            curCell = (Excel.Range)this.xlSheet.Cells[currentFind.Row + 2, i];
                            _temp = curCell.Value;
                            percentofcust = Math.Round(Convert.ToDouble(_temp.ToString()), 2, MidpointRounding.AwayFromZero);

                            ZuelligPharma_Frequencies.Add(new ZuelligPharma_Frequency()
                            {
                                adddt = (string)DateTime.Now.ToString("yyyyMMdd"),
                                seqno = (String)seqno.ToString().PadLeft(6, '0'),
                                freqno = freqno,
                                numofcust = numofcust,
                                percentofcust = percentofcust,
                                timestamp = DateTime.Now.ToString("yyyyMMddHHmmss")
                            });
                        }
                    }
                }
                else if (sheet.Name == @"Frequency per week")
                {
                    List<ZuelligPharma_FrequencyPerWeek> ZuelligPharma_FrequencyPerWeeks = new List<ZuelligPharma_FrequencyPerWeek>();
                    this.xlSheet = sheet;
                    Excel.Range xlRange = this.xlSheet.UsedRange;
                    Excel.Range curCell = null;
                    currentFind = xlRange.Find("Order no. via eZRx");
                    int seqno = 0;
                    object _temp = null;
                    string week = String.Empty;
                    int twice = 0;
                    int three = 0;
                    int more = 0;

                    for (int i = currentFind.Column + 1; i <= xlRange.Columns.Count; i++)
                    {
                        seqno++;// get seqno for rowinsert
                        var adddt = (string)DateTime.Now.ToString("yyyyMMdd");
                        curCell = (Excel.Range)this.xlSheet.Cells[currentFind.Row, i];
                        _temp = curCell.Value;
                        if (_temp == null || String.IsNullOrEmpty(_temp.ToString()) == true)
                        {
                            break;
                        }
                        else
                        {
                            // get week
                            week = _temp.ToString();

                            //get twice
                            curCell = (Excel.Range)this.xlSheet.Cells[currentFind.Row + 1, i];
                            _temp = curCell.Value;
                            twice = Convert.ToInt16(_temp.ToString());

                            //get three
                            curCell = (Excel.Range)this.xlSheet.Cells[currentFind.Row + 2, i];
                            _temp = curCell.Value;
                            three = Convert.ToInt16(_temp.ToString());

                            //get more
                            curCell = (Excel.Range)this.xlSheet.Cells[currentFind.Row + 3, i];
                            _temp = curCell.Value;
                            more = Convert.ToInt16(_temp.ToString());

                            ZuelligPharma_FrequencyPerWeeks.Add(new ZuelligPharma_FrequencyPerWeek()
                            {
                                adddt = (string)DateTime.Now.ToString("yyyyMMdd"),
                                seqno = (String)seqno.ToString().PadLeft(6, '0'),
                                week = week,
                                twice = twice,
                                three = three,
                                more = more,
                                timestamp = DateTime.Now.ToString("yyyyMMddHHmmss")
                            });
                        }
                    }
                }
                else if (sheet.Name == @"top PRN")
                {
                    List<ZuelligPharma_TopPRN> ZuelligPharma_TopPRNs = new List<ZuelligPharma_TopPRN>();
                    this.xlSheet = sheet;
                    Excel.Range xlRange = this.xlSheet.UsedRange;
                    Excel.Range curCell = null;
                    currentFind = xlRange.Find("Master PRN1");
                    nextFind = xlRange.Find(currentFind);
                    int seqno = 0;
                    object _temp = null;
                    string prnkey = String.Empty;
                    string monthfr = String.Empty;
                    string monthto = String.Empty;
                    double sale_monthfr = 0;
                    double sale_monthto = 0;
                    double month_growth = 0;
                    double month_share = 0;
                    string yearfr = String.Empty;
                    string yearto = String.Empty;
                    double sale_yearfr = 0;
                    double sale_yearto = 0;
                    double year_growth = 0;
                    double year_share = 0;

                    // get monthfr
                    curCell = (Excel.Range)this.xlSheet.Cells[currentFind.Row, currentFind.Column+1];
                    _temp = curCell.Value;
                    monthfr = _temp.ToString();

                    // get monthto
                    curCell = (Excel.Range)this.xlSheet.Cells[currentFind.Row, currentFind.Column + 2];
                    _temp = curCell.Value;
                    monthto = _temp.ToString();

                    // get yearfr
                    curCell = (Excel.Range)this.xlSheet.Cells[nextFind.Row, nextFind.Column + 1];
                    _temp = curCell.Value;
                    yearfr = _temp.ToString();

                    // get yearto
                    curCell = (Excel.Range)this.xlSheet.Cells[nextFind.Row, nextFind.Column + 2];
                    _temp = curCell.Value;
                    yearto = _temp.ToString();

                    for (int i = currentFind.Row + 1; i <= 10; i++)
                    {
                        seqno++;// get seqno for rowinsert
                        var adddt = (string)DateTime.Now.ToString("yyyyMMdd");

                        curCell = (Excel.Range)this.xlSheet.Cells[i, currentFind.Column + 1];
                        _temp = curCell.Value;
                        if (_temp == null || String.IsNullOrEmpty(_temp.ToString()) == true)
                        {
                            break;
                        }
                        else
                        {
                            // get sale_monthfr
                            sale_monthfr = Math.Round(Convert.ToDouble(_temp.ToString()), 0, MidpointRounding.AwayFromZero);

                            //get sale_monthto
                            curCell = (Excel.Range)this.xlSheet.Cells[i, currentFind.Column + 2];
                            _temp = curCell.Value;
                            sale_monthto = Math.Round(Convert.ToDouble(_temp.ToString()), 0, MidpointRounding.AwayFromZero);

                            // get month_growth
                            curCell = (Excel.Range)this.xlSheet.Cells[i, currentFind.Column + 3];
                            _temp = curCell.Value;
                            month_growth = Math.Round(Convert.ToDouble(_temp.ToString()), 0, MidpointRounding.AwayFromZero);
                            // get month_share
                            curCell = (Excel.Range)this.xlSheet.Cells[i, currentFind.Column + 4];
                            _temp = curCell.Value;
                            month_share = Math.Round(Convert.ToDouble(_temp.ToString()), 0, MidpointRounding.AwayFromZero);

                            //get sale_yearfr
                            curCell = (Excel.Range)this.xlSheet.Cells[i, nextFind.Column + 1];
                            _temp = curCell.Value;
                            sale_yearfr = Math.Round(Convert.ToDouble(_temp.ToString()), 0, MidpointRounding.AwayFromZero);

                            //get sale_yearto
                            curCell = (Excel.Range)this.xlSheet.Cells[i, nextFind.Column + 2];
                            _temp = curCell.Value;
                            sale_yearto = Math.Round(Convert.ToDouble(_temp.ToString()), 0, MidpointRounding.AwayFromZero);

                            // get year_growth
                            curCell = (Excel.Range)this.xlSheet.Cells[i, nextFind.Column + 3];
                            _temp = curCell.Value;
                            year_growth = Math.Round(Convert.ToDouble(_temp.ToString()), 0, MidpointRounding.AwayFromZero);

                            // get year_share
                            curCell = (Excel.Range)this.xlSheet.Cells[i, nextFind.Column + 4];
                            _temp = curCell.Value;
                            year_share = Math.Round(Convert.ToDouble(_temp.ToString()), 0, MidpointRounding.AwayFromZero);

                            // add into List
                            ZuelligPharma_TopPRNs.Add(new ZuelligPharma_TopPRN()
                            {
                                adddt = (string)DateTime.Now.ToString("yyyyMMdd"),
                                seqno = (String)seqno.ToString().PadLeft(6, '0'),
                                prnkey = "PRN" + seqno.ToString(),
                                monthfr = monthfr,
                                monthto = monthto,
                                sale_monthfr = sale_monthfr,
                                sale_monthto = sale_monthto,
                                yearfr = yearfr,
                                yearto = yearto,
                                sale_yearfr = sale_yearfr,
                                sale_yearto = sale_yearto,
                                timestamp = DateTime.Now.ToString("yyyyMMddHHmmss")
                            });
                        }

                        // get total
                        //// sale_monthfr
                        curCell = (Excel.Range)this.xlSheet.Cells[currentFind.Row+ 11, currentFind.Column + 1];
                        _temp = curCell.Value;
                        sale_monthfr = Math.Round(Convert.ToDouble(_temp.ToString()), 0, MidpointRounding.AwayFromZero);

                        //// sale_monthto
                        curCell = (Excel.Range)this.xlSheet.Cells[currentFind.Row + 11, currentFind.Column + 2];
                        _temp = curCell.Value;
                        sale_monthto = Math.Round(Convert.ToDouble(_temp.ToString()), 0, MidpointRounding.AwayFromZero);

                        //// sale_yearfr
                        curCell = (Excel.Range)this.xlSheet.Cells[nextFind.Row + 11, nextFind.Column + 1];
                        _temp = curCell.Value;
                        sale_yearfr = Math.Round(Convert.ToDouble(_temp.ToString()), 0, MidpointRounding.AwayFromZero);

                        //// sale_yearto
                        curCell = (Excel.Range)this.xlSheet.Cells[nextFind.Row + 11, nextFind.Column + 2];
                        _temp = curCell.Value;
                        sale_yearto = Math.Round(Convert.ToDouble(_temp.ToString()), 0, MidpointRounding.AwayFromZero);

                        ZuelligPharma_TopPRNs.Add(new ZuelligPharma_TopPRN()
                        {
                            adddt = (string)DateTime.Now.ToString("yyyyMMdd"),
                            seqno = (String)(seqno+1).ToString().PadLeft(6, '0'),
                            prnkey = "Total",
                            monthfr = monthfr,
                            monthto = monthto,
                            sale_monthfr = sale_monthfr,
                            sale_monthto = sale_monthto,
                            yearfr = yearfr,
                            yearto = yearto,
                            sale_yearfr = sale_yearfr,
                            sale_yearto = sale_yearto,
                            timestamp = DateTime.Now.ToString("yyyyMMddHHmmss")
                        });
                    }
                }
            }
        }
        public void Quit()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();

            if (xlSheet != null)
            {
                Marshal.ReleaseComObject(xlSheet);
            }

            if (xlWorkBook != null)
            {
                xlWorkBook.Close(0);
                Marshal.ReleaseComObject(xlWorkBook);
            }

            if (xlApp != null)
            {
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
                xlApp = null;
            }
        }
    }
}