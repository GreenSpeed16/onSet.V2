using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace onSet
{
    class Route
    {
        //Fields
        protected string color, wall, setter;

        public string grade { get; set; }
        public int Row { get; set; }
        public string Column { get; set; }

        //Constructor
        public Route(string grade, string wall, string setter, string color)
        {
            this.grade = grade;
            this.color = color;
            this.wall = wall;
            this.setter = setter;
        }

        //Methods
        public void submitRoute(Excel.Worksheet routeSheet, Excel.Workbook routeBook, Excel.Application reader)
        {

            int lastUsedRow = routeSheet.Cells.Find("*", System.Reflection.Missing.Value,
                               System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                               Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                               false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

            //Update last row
            routeSheet.Cells[(lastUsedRow + 1), "A"].Value = "'" + grade;
            routeSheet.Cells[(lastUsedRow + 1), "B"].Value = "'" + wall;
            routeSheet.Cells[(lastUsedRow + 1), "C"].Value = "'" + setter;
            routeSheet.Cells[(lastUsedRow + 1), "D"].Value = "'" + color;

            routeBook.Save();
        }

        public void deleteRoute(Excel.Worksheet routeSheet, Excel.Workbook routeBook, bool isPreset, Excel.Application reader)
        {
            /*            //Instantiate variables
                        Range rowToDelete;
                        Range currentSpreadSheet;*/

            int lastUsedRow = routeSheet.Cells.Find("*", System.Reflection.Missing.Value,
                               System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                               Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                               false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;

            //Delete route from spreadsheet
            if (isPreset)
            {
                routeSheet.Range[string.Format("{0}{1}:{2}{3}", Column, 2, Column, lastUsedRow)].Clear();
                /*currentSpreadSheet = routeSheet[string.Format("{0}{1}:{2}{3}", Column, 2, Column, routeSheet.RowCount)];*/
            }
            else
            {
                routeSheet.Range[string.Format("A{0}:D{1}", Row, Row)].Delete();
                /*currentSpreadSheet = routeSheet[string.Format("A{0}:D{1}", Row, Row)];*/
            }

            
            /*rowToDelete.ClearContents();
            currentSpreadSheet.Trim();*/

            routeBook.Save();
        }

        public override string ToString()
        {
            if (color == "Preset") return grade;
            else return string.Format("{0}, {1} {2}, {3}", setter, color, grade, wall);
        }
    }
}
