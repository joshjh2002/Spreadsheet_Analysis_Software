using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace Spreadsheet_Analysis_Software
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        int worksheetrows, worksheetcolumns;
        int rows, columns;

        string[,] resizedWorksheet;

        private void import_spreadsheet_button_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog() { Filter = "Spreadsheet (*.xls)|*.xls| Spreadsheet (*.xlsx)|*.xlsx" };
            
            //Import Spreadsheet
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                spreadsheet_dir_textbox.Text = ofd.FileName;
                GetWorksheet(ofd.FileName);
                MessageBox.Show("Spreadsheet Sucessfully Imported");
                export_spreadsheet_button.Enabled = true;
            }
        }

        private void export_spreadsheet_button_Click(object sender, EventArgs e)
        {
            if (Properties.Settings.Default.default_directory != "no_dir")
            {
                string filename = Properties.Settings.Default.default_directory + "/" + DateTime.Now.Month + " " + DateTime.Now.Day + " " + DateTime.Now.Year + "_Report.xlsx";

                //MessageBox.Show(filename);

                Excel.Application xlApp = new Excel.Application();

                Excel.Workbook xlWorkbook = xlApp.Workbooks.Add();

                Excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1];

                //Write array to spreadsheet
                int titleColumn = 0;
                int sessionColumn = 0;
                int checkInColumn = 0;
                for (int i = 0; i <= rows; i++)
                {
                    for (int j = 0; j < 7; j++)
                    {
                        //MessageBox.Show("i, j: " + i + "," + j);
                        xlWorksheet.Cells[i + 1, j + 1] = resizedWorksheet[i, j];

                        if (resizedWorksheet[i, j] == "Title")
                        {
                            titleColumn = j;
                        }
                        else if (resizedWorksheet[i, j] == "Session Title")
                        {
                            sessionColumn = j;
                        }
                        else if (resizedWorksheet[i, j] == "Check In Date")
                        {
                            checkInColumn = j;
                        }

                    }
                }

                

                //Analyse Data
                int meals_served = rows;
                int[] breakfast = new int[3];
                int[] lunch = new int[3];
                int[] snack = new int[3];

                //counts the number of paid, reduced and free snacks, lunches and breakfasts
                for (int i = 0; i <= rows; i++)
                {
                    if (resizedWorksheet[i, titleColumn].ToLower() == "paid" && resizedWorksheet[i, sessionColumn].ToLower() == "breakfast session")
                    {
                        breakfast[0] += 1;
                    }
                    else if (resizedWorksheet[i, titleColumn].ToLower() == "paid" && resizedWorksheet[i, sessionColumn].ToLower() == "lunch session")
                    {
                        lunch[0] += 1;
                    }
                    else if (resizedWorksheet[i, titleColumn].ToLower() == "paid" && resizedWorksheet[i, sessionColumn].ToLower() == "snack session")
                    {
                        snack[0] += 1;
                    }
                    else if (resizedWorksheet[i, titleColumn].ToLower() == "reduced" && resizedWorksheet[i, sessionColumn].ToLower() == "breakfast session")
                    {
                        breakfast[1] += 1;
                    }
                    else if (resizedWorksheet[i, titleColumn].ToLower() == "reduced" && resizedWorksheet[i, sessionColumn].ToLower() == "lunch session")
                    {
                        lunch[1] += 1;
                    }
                    else if (resizedWorksheet[i, titleColumn].ToLower() == "reduced" && resizedWorksheet[i, sessionColumn].ToLower() == "snack session")
                    {
                        snack[1] += 1;
                    }
                    else if (resizedWorksheet[i, titleColumn].ToLower() == "free" && resizedWorksheet[i, sessionColumn].ToLower() == "breakfast session")
                    {
                        breakfast[2] += 1;
                    }
                    else if (resizedWorksheet[i, titleColumn].ToLower() == "free" && resizedWorksheet[i, sessionColumn].ToLower() == "lunch session")
                    {
                        lunch[2] += 1;
                    }
                    else if (resizedWorksheet[i, titleColumn].ToLower() == "free" && resizedWorksheet[i, sessionColumn].ToLower() == "snack session")
                    {
                        snack[2] += 1;
                    }
                }

                //count the number of individual days
                int days = -1;
                for (int i = 0; i <= rows; i++)
                {
                    bool found = false;
                    for (int j = 0; j < i; j++)
                    {
                        if (resizedWorksheet[i, checkInColumn] == resizedWorksheet[j, checkInColumn])
                        {
                            found = true;
                        }
                    }
                    if (!found)
                    {
                        days++;
                    }
                }

                //Adding in the final tables
                xlWorksheet.Cells[3, 9] = "Breakfast"; 
                xlWorksheet.Cells[4, 9] = "Lunch"; 
                xlWorksheet.Cells[5, 9] = "Snacks"; 
                xlWorksheet.Cells[6, 9] = "Totals";

                xlWorksheet.Cells[2, 10] = "Paid";
                xlWorksheet.Cells[2, 11] = "Reduced";
                xlWorksheet.Cells[2, 12] = "Free";
                xlWorksheet.Cells[2, 13] = "Totals";

                xlWorksheet.Cells[3, 10] = breakfast[0].ToString();
                xlWorksheet.Cells[3, 11] = breakfast[1].ToString();
                xlWorksheet.Cells[3, 12] = breakfast[2].ToString();
                xlWorksheet.Cells[3, 13] = (breakfast[0] + breakfast[1] + breakfast[2]).ToString();

                xlWorksheet.Cells[4, 10] = lunch[0].ToString();
                xlWorksheet.Cells[4, 11] = lunch[1].ToString();
                xlWorksheet.Cells[4, 12] = lunch[2].ToString();
                xlWorksheet.Cells[4, 13] = (lunch[0] + lunch[1] + lunch[2]).ToString();

                xlWorksheet.Cells[5, 10] = snack[0].ToString();
                xlWorksheet.Cells[5, 11] = snack[1].ToString();
                xlWorksheet.Cells[5, 12] = snack[2].ToString();
                xlWorksheet.Cells[5, 13] = (snack[0] + snack[1] + snack[2]).ToString();

                xlWorksheet.Cells[6, 10] = (breakfast[0] + lunch[0] + snack[0]).ToString();
                xlWorksheet.Cells[6, 11] = (breakfast[1] + lunch[1] + snack[1]).ToString();
                xlWorksheet.Cells[6, 12] = (breakfast[2] + lunch[2] + snack[2]).ToString();

                xlWorksheet.Cells[7, 14] = "Total Days:";
                xlWorksheet.Cells[8, 14] = "Total Meals Served:";

                xlWorksheet.Cells[7, 15] = days;
                xlWorksheet.Cells[8, 15] = meals_served;

                try
                {
                    xlWorkbook.SaveAs(filename);
                }
                catch
                {
                    MessageBox.Show("The application cannot access this file. Make sure it is not open.");
                }

                xlApp.Quit();

            }
            else
            {
                MessageBox.Show("You must first set a default directory.");
            }
        }

        private void GetWorksheet(string file_dir)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(file_dir);

            Excel.Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            
            rows = xlRange.Rows.Count - 1;
            columns = xlRange.Columns.Count;

            string[,] worksheet = new string[rows + 2,columns + 2];
            worksheetrows = rows + 2;
            worksheetcolumns = worksheetcolumns + 2;
            //MessageBox.Show("columns: " + columns);

            for (int curCol = 1; curCol <= columns;  curCol++)
            {
                for (int curRow = 1; curRow <= rows + 1; curRow++)
                {
                    worksheet[curRow - 1, curCol - 1] = xlRange.Cells[curRow, curCol].Value2.ToString() + "";
                    //MessageBox.Show((curCol - 1 ) + " , " + (curRow - 1) + "   " +  worksheet[curRow - 1, curCol - 1]);
                }
            }

            xlWorkbook.Close();
            xlApp.Quit();

            int numberOfColumnsFound = 0;
            resizedWorksheet = new string[rows + 1, 7];

            //Generates a small array to reduce space
            for (int i = 0; i < columns; i++)
            {
                if (worksheet[0, i] == "Last Name" || worksheet[0, i] == "First Name" || worksheet[0, i] == "Title" || worksheet[0, i] == "Badge Number" || worksheet[0, i] == "Session ID"
                    || worksheet[0, i] == "Session Title" || worksheet[0, i] == "Check In Date")
                {
                    for (int j = 0; j <= rows; j++)
                    {
                        resizedWorksheet[j, numberOfColumnsFound] = worksheet[j, i];
                        //MessageBox.Show(i + ", " + j + "    " + resizedWorksheet[j, numberOfColumnsFound]);
                    }
                    numberOfColumnsFound++;
                }
            }
        }

        private void set_export_button_Click(object sender, EventArgs e)
        {
            //Sets export folder
            FolderBrowserDialog fbd = new FolderBrowserDialog();

            if (fbd.ShowDialog() == DialogResult.OK)
            {
                Properties.Settings.Default.default_directory = fbd.SelectedPath;
                Properties.Settings.Default.Save();
            }
        }
    }
}
