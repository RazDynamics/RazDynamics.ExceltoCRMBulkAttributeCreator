using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace CRMConsultants.CreateAttributes
{
    public class CreateExcelDoc
    {
        public Microsoft.Office.Interop.Excel.Application app = null;
        public Microsoft.Office.Interop.Excel.Workbook workbook = null;
        public Microsoft.Office.Interop.Excel.Worksheet worksheet = null;
        public Microsoft.Office.Interop.Excel.Range workSheet_range = null;
        public CreateExcelDoc(string filePath)
        {
            createDoc(filePath);
        }
        internal void createDoc(string filePath)
        {
            try
            {
                app = new Microsoft.Office.Interop.Excel.Application();
                //app.Visible = true;
                workbook = app.Workbooks.Open(filePath);
                worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];
                workSheet_range = worksheet.UsedRange;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        internal void createHeaders(int row, int col, string htext, string cell1,
        string cell2, int mergeColumns, string b, bool font, int size, string
        fcolor)
        {
            worksheet.Cells[row, col] = htext;
            workSheet_range = worksheet.get_Range(cell1, cell2);
            workSheet_range.Merge(mergeColumns);
            switch (b)
            {
                case "YELLOW":
                    workSheet_range.Interior.Color = System.Drawing.Color.Yellow.ToArgb();
                    break;
                case "GRAY":
                    workSheet_range.Interior.Color = System.Drawing.Color.Gray.ToArgb();
                    break;
                case "GAINSBORO":
                    workSheet_range.Interior.Color =
            System.Drawing.Color.Gainsboro.ToArgb();
                    break;
                case "Turquoise":
                    workSheet_range.Interior.Color =
            System.Drawing.Color.Turquoise.ToArgb();
                    break;
                case "PeachPuff":
                    workSheet_range.Interior.Color =
            System.Drawing.Color.PeachPuff.ToArgb();
                    break;
                default:
                    //  workSheet_range.Interior.Color = System.Drawing.Color..ToArgb();
                    break;
            }

           // workSheet_range.Borders.Color = System.Drawing.Color.Black.ToArgb();
           // workSheet_range.Font.Bold = font;
            workSheet_range.ColumnWidth = size;
            if (fcolor.Equals(""))
            {
                workSheet_range.Font.Color = System.Drawing.Color.White.ToArgb();
            }
            else
            {
                workSheet_range.Font.Color = System.Drawing.Color.Black.ToArgb();
            }
        }

        internal void addData(int row, int col, string data,
            string cell1, string cell2)
        {
            worksheet.Cells[row, col] = data;
            workSheet_range = worksheet.get_Range(cell1, cell2);
            //workSheet_range.Borders.Color = System.Drawing.Color.Black.ToArgb();
            //workSheet_range.NumberFormat = format;
        }

        internal void ReleaseObject()
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workSheet_range);
            }
            catch (Exception ex)
            {
                app = null;
                workbook = null;
                worksheet = null;
                workSheet_range = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
