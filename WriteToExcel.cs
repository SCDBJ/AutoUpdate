using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace AutoUpdate
{
    public class WriteToExcel
    {
        public static void SaveWrokExcel(string templateFileName, string outFileName,string dateRange,IList<WeekModel> weekModels)
        {
            //需要添加 Microsoft.Office.Interop.Excel引用 
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            //Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.ApplicationClass();

            if (app == null)
            {
                throw new Exception("这台计算机上缺少Excel组件，需要安装Office2003及以上软件。");
            }
            app.Visible = false;
            app.UserControl = true;
            Microsoft.Office.Interop.Excel.Workbooks workbooks = app.Workbooks;
            Microsoft.Office.Interop.Excel._Workbook workbook = workbooks.Add(templateFileName); //加载模板

            Microsoft.Office.Interop.Excel.Sheets sheets = workbook.Sheets;
            Microsoft.Office.Interop.Excel._Worksheet worksheet = (Microsoft.Office.Interop.Excel._Worksheet)sheets.get_Item(1); //第一个工作薄。
            if (worksheet == null)
            {
                throw new Exception("工作薄模板中没有工作表!");  //工作薄中没有工作表.
            }

            worksheet.Cells[2, 3] = dateRange;

            for (int i = 0; i < weekModels.Count; i++)
            {
                int row_ = 4 + i;  //Excel模板上表头和标题行占了3行,根据实际模板需要修改;
                worksheet.Cells[row_, 3] = weekModels[i].workContent;
                worksheet.Cells[row_, 4] = weekModels[i].workTarget;
                worksheet.Cells[row_, 6] = weekModels[i].completion;
            }

            workbook.SaveAs(outFileName, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);

            //4、按顺序释放资源。
            NAR(worksheet);
            NAR(sheets);
            NAR(workbook);
            NAR(workbooks);
            app.Quit();
            NAR(app);
        }

        private static void NAR(object o)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(o);
            }
            catch
            {
            }
            finally
            {
                o = null;
            }
        }
    }
}
