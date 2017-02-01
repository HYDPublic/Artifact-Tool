using System;
using System.Windows.Forms;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Linq.Expressions;

namespace Artifact_Tool
{
    public partial class Form1 : Form
    {
        private DateTime pStart;
        private DateTime pFinish;
        private DateTime aFinish;
        private DateTime aStart;
        private DateTime dateOnly;

        public Form1()
        {
            InitializeComponent();
        }



        public void Read_File(string filename)
        {
            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook xlWorkbook = xlApp.Workbooks.Open(filename, 0, false, 5, "", "", true,
                XlPlatform.xlWindows, ",", false, false, 0, true, 1, 0);
            _Worksheet xlWorksheet = (_Worksheet)xlWorkbook.Sheets[1];
            Range xlRange = xlWorksheet.UsedRange;

            xlWorksheet.ListObjects.AddEx(XlListObjectSourceType.xlSrcRange, xlRange, Type.Missing,
                XlYesNoGuess.xlYes).Name = "WFTableStyle";

            xlWorksheet.ListObjects.get_Item("WFTableStyle").TableStyle = "TableStyleMedium16";
            

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            int lastRow = xlWorksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row;

            for (int index = 2; index <= lastRow; index++)
            {
                Range rows = xlRange.Rows[index];
                try
                {
                    System.Array myValues = (System.Array)xlWorksheet.get_Range("A" + index.ToString(),
                        "V" + index.ToString()).Cells.Value;

                    Check_Line(myValues, rows, xlWorksheet, index);

                }
                catch (NullReferenceException e)
                {
                    MessageBox.Show(e.Message);
                    CloseWorkbook(xlApp, xlRange, xlWorkbook, xlWorksheet);
                }
            }

            xlWorkbook.Save();
            CloseWorkbook(xlApp, xlRange, xlWorkbook, xlWorksheet);
        }

        public void Check_Line(Array item, Range row, _Worksheet page, int currentRow )
        {

            dateOnly = Convert.ToDateTime(DateTime.Now.ToString("MM/d/yyyy"));

            if (item.GetValue(1, 15) != null)
                if (item.GetValue(1, 15).ToString() != "")
            {
                pStart = Convert.ToDateTime(item.GetValue(1, 15).ToString());
            }

            if (item.GetValue(1, 14) != null)
                if (item.GetValue(1, 14).ToString() != "")
            {
                pFinish = Convert.ToDateTime(item.GetValue(1, 14).ToString());
            }

            if (item.GetValue(1, 4) != null)
                if (item.GetValue(1, 4).ToString() != "")
            {
                aStart = Convert.ToDateTime(item.GetValue(1, 4).ToString());
            }

            if (item.GetValue(1, 3) != null)
                if(item.GetValue(1, 3).ToString() != "")
            {
                aFinish = Convert.ToDateTime(item.GetValue(1, 3).ToString());
            }

            // If planned start is not empty then...
            if (item.GetValue(1, 15) != null)
                if (item.GetValue(1, 15).ToString() != "")
            {
                // If todays date is earlier then planned start then...
                if (dateOnly < pStart)
                {
                    row.Interior.Color = System.Drawing.Color.LightGreen;
                }


                // If todays date is the same or later then the planned start and earlier then the planned finish date
                // then...
                if (dateOnly >= pStart && dateOnly < pFinish)
                {
                    row.Interior.Color = System.Drawing.Color.LightGreen;
                }
                // If todays date is the same or later then the planned start and the same or
                // later the planned finish then..
                else if(dateOnly >= pStart && dateOnly >= pFinish)
                {
                    // Color the Planned Finish cell red
                    page.Cells[currentRow, 14].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.HotPink);

                    // If actual start is not greater or equal to planned start or empty then...
                    if (item.GetValue(1, 4) != null)
                        if(item.GetValue(1, 4).ToString() == "" || item.GetValue(1, 8).ToString() == null)
                    {
                        // Color the actual start cell red
                        page.Cells[currentRow, 4].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
                    }
                    else
                        {
                            // Color the actual start cell red
                            page.Cells[currentRow, 4].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.HotPink);
                        }

                    // If actual finish is empty or is less or equal to todays date then ...
                    if (item.GetValue(1, 3) != null)
                        if (item.GetValue(1, 3).ToString() != "" &&  dateOnly <= aFinish)
                    {
                        // Color the actual finish cell red
                       page.Cells[currentRow, 3].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
                    }
                    else
                    {
                        // Color the actual finish cell red
                        page.Cells[currentRow, 3].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.HotPink);
                    }

                    // If Estimated Coding is empty then ... 
                    if (item.GetValue(1, 11) != null)
                        if (item.GetValue(1, 11).ToString() != "" )
                    {
                        // Color the estimated coding cell green
                        page.Cells[currentRow, 11].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
                    }
                    else
                        {
                            // Color the estimated coding cell red
                            page.Cells[currentRow, 11].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.HotPink);
                        }

                    // If Estimated Design is empty then ...
                    if (item.GetValue(1, 12) != null)
                        if (item.GetValue(1, 12).ToString() != "" )
                    {
                        // Color the estimated design cell green
                        page.Cells[currentRow, 12].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
                    }
                    else
                        {
                            // Color the estimated design cell red
                            page.Cells[currentRow, 12].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.HotPink);
                        }

                    // if Actual Coding is empty then ...
                    if (item.GetValue(1, 1) != null)
                        if (item.GetValue(1, 1).ToString() != "")
                    {
                        // Color the actual coding cell red
                        page.Cells[currentRow, 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
                    }
                    else
                    {
                        page.Cells[currentRow, 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.HotPink);
                    }

                    // if Actual Design is empty then ...
                    if(item.GetValue(1, 2) != null)
                        if(item.GetValue(1, 2).ToString() != "")
                    {
                        // Color the actual design cell red
                        page.Cells[currentRow, 2].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
                    }
                    else
                    {
                        // Color the actual design cell red
                        page.Cells[currentRow, 2].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.HotPink);
                    }

                    // if Code Review Passed is yes then ...
                    if (item.GetValue(1, 9) != null)
                    if (item.GetValue(1, 9).ToString() != "")
                    if (item.GetValue(1, 9).ToString() != "No")
                        if (item.GetValue(1, 9).ToString() == "Yes" )
                    {
                        // Color the Code Review Comments cell green
                        page.Cells[currentRow, 9].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
                    }
                    else
                    {
                        // Color the Code Review Comments cell red
                        page.Cells[currentRow, 9].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.HotPink);
                    }

                    // if Code Review Comments is empty then ...
                    if (item.GetValue(1, 8) != null)
                    if (item.GetValue(1, 8).ToString() != "")
                    {
                        // Color the Code Review Comments cell green
                        page.Cells[currentRow, 8].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
                    }
                    else
                    {
                        // Color the Code Review Comments cell red
                        page.Cells[currentRow, 8].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.HotPink);
                    }

                    // if Unit Test Passed is yes then ...
                    if (item.GetValue(1, 21) != null)
                    if (item.GetValue(1, 21).ToString() != "")
                    if (item.GetValue(1, 21).ToString() != "No")
                        if (item.GetValue(1, 21).ToString() == "Yes" && item.GetValue(1, 22) != null)
                    {
                        // Color the unit test cell green
                        page.Cells[currentRow, 21].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
                    }
                    else
                    {
                        // Color the unit test cell red
                        page.Cells[currentRow, 21].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.HotPink);
                    }


                    // if Unit Test Passed is no and Unit Test is empty then ...
                    if (item.GetValue(1, 22) != null)
                    if (item.GetValue(1, 22).ToString() != "")
                    {
                        // Color the Unit Test cell green
                        page.Cells[currentRow, 22].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
                    }
                    else
                    {
                        // Color the Unit Test cell red
                        page.Cells[currentRow, 22].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.HotPink);
                    }
                }
            }
            else
            {
                row.Interior.Color = System.Drawing.Color.HotPink;
            }


        }

        public void CloseWorkbook(Microsoft.Office.Interop.Excel.Application aName, Range rName, Workbook wbName, _Worksheet wsName)
        {


            // Clean Up
            GC.Collect();
            GC.WaitForPendingFinalizers();

            // Rule of thumb for releasing COM objects:
            // Never use two dots, all COM objects must be referenced and released
            // ex: [something.[something].[something] is bad

            // Release COM objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(rName);
            Marshal.ReleaseComObject(wsName);

            // Close and release
            wbName.Close();
            Marshal.ReleaseComObject(wbName);

            // Quit and release
            aName.Quit();
            Marshal.ReleaseComObject(aName);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            string chosenFile = openFileDialog1.FileName;
            MessageBox.Show("Your file is " + chosenFile);
            Read_File(chosenFile);
        }
    }
}
