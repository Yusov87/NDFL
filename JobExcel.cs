using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace namespaceJobExcel
{
    public class JobExcel
    {
        public String FileName = "";
        public Decimal Sell, buy, Pribil, Ndfl;
        public void ReadExcel()
        {
            int[] Collumns = new int[2] { 7, 8 };

            Excel.Application ObjWorkExcel = new Excel.Application(); //открыть эксель
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //открыть файл
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1]; //получить 1 лист
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);//1 ячейку
            string[,] list = new string[lastCell.Column, lastCell.Row]; // массив значений с листа равен по размеру листу
                                                                       
            for (int j = 1; j < (int)lastCell.Row; j++) // по всем строкам

            {
                foreach (int k in Collumns)
                    FillN(ObjWorkSheet.Cells[j, k].Text.ToString(), k);
            }
            Pribil = buy - Sell;
            Ndfl = Decimal.Round(Pribil * 13 / 100, 2);

            ObjWorkBook.Close(false); //закрыть не сохраняя
            ObjWorkExcel.Quit(); // выйти из экселя
            GC.Collect(); // убрать за собой

        }

        private void FillN(String T, int k)
        {
            bool result;
            Decimal delta;

            result = Decimal.TryParse(T, out delta);

            if (result) 
                if (k == 7) buy = buy + delta;
                else Sell = Sell + delta;
        }
    }
}
