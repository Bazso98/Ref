using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace ExcelP
{
        
        public partial class Form1 : Form
        {
            RealEstateEntities context = new RealEstateEntities();

            List<Flat> Flats;
            Excel.Application xlApp;
            Excel.Workbook xlWB;
            Excel.Worksheet xlSheet;

            string[] headers;
            public Form1()
            {

                InitializeComponent();

                LoadData();
                CreatExcel();
                FormatTable();

            }

            public void LoadData()
            {
                Flats = context.Flats.ToList();
            }

            public void CreatExcel()
            {
                try
                {

                    xlApp = new Excel.Application();


                    xlWB = xlApp.Workbooks.Add(Missing.Value);


                    xlSheet = xlWB.ActiveSheet;


                    CreateTable();


                    xlApp.Visible = true;
                    xlApp.UserControl = true;
                }
                catch (Exception ex)
                {
                    string errMsg = string.Format("Error: {0}\nLine: {1}", ex.Message, ex.Source);
                    MessageBox.Show(errMsg, "Error");


                    xlWB.Close(false, Type.Missing, Type.Missing);
                    xlApp.Quit();
                    xlWB = null;
                    xlApp = null;
                }
            }

            public void CreateTable()
            {
                headers = new string[]
                {
                "Kód",
                "Eladó",
                "Oldal",
                "Kerület",
                "Lift",
                "Szobák száma",
                "Alapterület (m2)",
                "Ár (mFt)",
                "Négyzetméter ár (Ft/m2)"

                };

                for (int i = 0; i < headers.Length; i++)
                {
                    xlSheet.Cells[1, i + 1] = headers[i];
                }

                object[,] values = new object[Flats.Count(), headers.Length];

                int counter = 0;

                foreach (Flat f in Flats)
                {
                    values[counter, 0] = f.Code;
                    values[counter, 1] = f.Vendor;
                    values[counter, 2] = f.Side;
                    values[counter, 3] = f.District;
                    values[counter, 4] = f.Elevator ? "Van" : "Nincs";
                    values[counter, 5] = f.NumberOfRooms;
                    values[counter, 6] = f.FloorArea;
                    values[counter, 7] = f.Price;
                    values[counter, 8] = "=" + GetCell(counter + 2, 8) + "/" + GetCell(counter + 2, 7) + "*1000000";
                    counter++;

                }

                xlSheet.get_Range(
                GetCell(2, 1),
                GetCell(1 + values.GetLength(0), values.GetLength(1))).Value2 = values;




            }

            public void FormatTable()
            {
                Excel.Range headerRange = xlSheet.get_Range(GetCell(1, 1), GetCell(1, headers.Length));
                headerRange.Font.Bold = true;
                headerRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                headerRange.EntireColumn.AutoFit();
                headerRange.RowHeight = 40;
                headerRange.Interior.Color = Color.LightBlue;
                headerRange.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);


                int lastRowID = xlSheet.UsedRange.Rows.Count;
                int lastColumnID = xlSheet.UsedRange.Columns.Count;
                Excel.Range tableRange = xlSheet.get_Range(GetCell(2, 1), GetCell(lastRowID, lastColumnID));
                tableRange.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);


                Excel.Range tableRange2 = xlSheet.get_Range(GetCell(2, 1), GetCell(lastRowID, 1));
                tableRange2.Interior.Color = Color.LightYellow;
                tableRange2.Font.Bold = true;

                Excel.Range tableRange3 = xlSheet.get_Range(GetCell(2, lastColumnID), GetCell(lastRowID, lastColumnID));
                tableRange3.Interior.Color = Color.LightGreen;
                tableRange3.Cells.NumberFormat = "0.00";

            }

            private string GetCell(int x, int y)
            {
                string ExcelCoordinate = "";
                int dividend = y;
                int modulo;

                while (dividend > 0)
                {
                    modulo = (dividend - 1) % 26;
                    ExcelCoordinate = Convert.ToChar(65 + modulo).ToString() + ExcelCoordinate;
                    dividend = (int)((dividend - modulo) / 26);
                }
                ExcelCoordinate += x.ToString();

                return ExcelCoordinate;
            }


        }

    }

