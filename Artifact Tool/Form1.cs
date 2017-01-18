using System;
using System.Windows.Forms;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace Artifact_Tool
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public void Read_File(string filename)
        {




            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook xlWorkbook = xlApp.Workbooks.Open(filename, 0, false, 5, "", "", true,
                XlPlatform.xlWindows, ",", false, false, 0, true, 1, 0);
            _Worksheet xlWorksheet = (Microsoft.Office.Interop.Excel._Worksheet)xlWorkbook.Sheets[1];
            Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            int lastRow = xlWorksheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell).Row;



            List<Artifact> idList = new List<Artifact>(50);

            Console.WriteLine("Capacity: {0}", idList.Capacity);

            for (int index = 2; index <= lastRow; index++)
            {
                System.Array myValues = (System.Array)xlWorksheet.get_Range("A" + index.ToString(), 
                    "V" + index.ToString()).Cells.Value;

                if (!string.IsNullOrEmpty(myValues.GetValue(1, 1).ToString()))
                {
                    idList.Add(new Artifact() { Artifact_ID = myValues.GetValue(1, 1).ToString() });
                }
                else
                {
                    idList.Add(new Artifact() { Artifact_ID = null });
                }

                if (!string.IsNullOrEmpty(myValues.GetValue(1, 2).ToString()))
                {
                    idList.Add(new Artifact() { Title = myValues.GetValue(1, 2).ToString() });
                }
                else
                {
                    idList.Add(new Artifact() { Title = null });
                }
                if (!string.IsNullOrEmpty(myValues.GetValue(1, 3).ToString()))
                {
                    idList.Add(new Artifact() { Assigned_To = myValues.GetValue(1, 3).ToString() });
                }
                else
                {
                    idList.Add(new Artifact() { Assigned_To = null });
                }
                if (!string.IsNullOrEmpty(myValues.GetValue(1, 4).ToString()))
                {
                    idList.Add(new Artifact() { Status = myValues.GetValue(1, 4).ToString() });
                }
                else
                {
                    idList.Add(new Artifact() { Status = null });
                }
                if (!string.IsNullOrEmpty(myValues.GetValue(1, 5).ToString()))
                {
                    idList.Add(new Artifact() { Actual_Coding = myValues.GetValue(1, 5).ToString() });
                }
                else
                {
                    idList.Add(new Artifact() { Actual_Coding = null });
                }
                if (!string.IsNullOrEmpty(myValues.GetValue(1, 6).ToString()))
                {
                    idList.Add(new Artifact() { Actual_Design = myValues.GetValue(1, 6).ToString() });
                }
                else
                {
                    idList.Add(new Artifact() { Actual_Design = null });
                }
                if (!string.IsNullOrEmpty(myValues.GetValue(1, 7).ToString()))
                {
                    idList.Add(new Artifact() { Actual_Finish = myValues.GetValue(1, 7).ToString() });
                }
                else
                {
                    idList.Add(new Artifact() { Actual_Finish = null });
                }
                idList.Add(new Artifact() { Actual_Start = myValues.GetValue(1, 8).ToString() });
                idList.Add(new Artifact() { Actual_Testing = myValues.GetValue(1, 9).ToString() });
                idList.Add(new Artifact() { Code_Review_Comments = myValues.GetValue(1, 10).ToString() });
                idList.Add(new Artifact() { Code_Review_Passed = myValues.GetValue(1, 11).ToString() });
                idList.Add(new Artifact() { Code_Reviewer = myValues.GetValue(1, 12).ToString() });
                idList.Add(new Artifact() { Estimated_Coding = myValues.GetValue(1, 13).ToString() });
                idList.Add(new Artifact() { Estimated_Design = myValues.GetValue(1, 14).ToString() });
                idList.Add(new Artifact() { Estimated_Testing = myValues.GetValue(1, 15).ToString() });
                idList.Add(new Artifact() { Planned_Finish = myValues.GetValue(1, 16).ToString() });
                idList.Add(new Artifact() { Planned_Start = myValues.GetValue(1, 17).ToString() });
                idList.Add(new Artifact() { System_Test_Required = myValues.GetValue(1, 18).ToString() });
                idList.Add(new Artifact() { Test_Tools_Description = myValues.GetValue(1, 19).ToString() });
                idList.Add(new Artifact() { Tested_By = myValues.GetValue(1, 20).ToString() });
                idList.Add(new Artifact() { Unit_Test_Passed = myValues.GetValue(1, 21).ToString() });
                idList.Add(new Artifact() { Unit_Test = myValues.GetValue(1, 22).ToString() }); 
            }

            foreach (Artifact artf in idList)
            {
                Console.WriteLine(artf);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            string chosenFile = openFileDialog1.FileName;
            MessageBox.Show("Your file is " + chosenFile);
            Read_File(chosenFile);


        }
    }
    public class Artifact
    {
        public string Artifact_ID, Title, Assigned_To, Status, Actual_Coding, Actual_Design, Actual_Finish, Actual_Start,
            Actual_Testing, Code_Review_Comments, Code_Review_Passed, Code_Reviewer, Estimated_Coding, Estimated_Design,
            Estimated_Testing, Planned_Finish, Planned_Start, System_Test_Required, Test_Tools_Description, Tested_By,
            Unit_Test_Passed, Unit_Test;
    }
}
