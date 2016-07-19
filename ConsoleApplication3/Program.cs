using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication3
{
    class Program
    {
        static void Main(string[] args)
        {
            using (var db=new Model1())
            {
                MyEntity entry = db.MyEntities.FirstOrDefault();
                ImportExcel(db.Books.ToList());
            }
               
        }
        private static void ImportExcel(List<教材库> books)
        {
            string sql = "INSERT [dbo].[教材库] ([名称], [作者], [单位], [类型], [地址], [知识体系编号], [格式], [电子书编号], [系统编号], [工种], [岗位]) VALUES ";
            FileStream file = new FileStream(@"E:\Projects\JCK.xlsx", FileMode.Open, FileAccess.Read);
            XSSFWorkbook wk = new XSSFWorkbook(file);
            ISheet sheet = wk.GetSheetAt(0);
            List<string> list = new List<string>();
            int rowNum = 0;
            for (int i = 1; i <= sheet.LastRowNum; i++)
            {
                rowNum = i;
                IRow row = sheet.GetRow(i);
                if (row != null)
                {
                   var name=row.GetCell(0).ToString();
                    var endname = name.Substring(name.IndexOf("_")+1).Trim();
                    var urlname= name.Substring(name.IndexOf("_") -14).Trim();
                    if (books.Where(m=>m.地址.Contains(endname.Trim())).Count()==0)
                    {
                        list.Add(endname);
                        var filetype = endname.ToLower().Contains(".pdf") ? "PDF文档" : "视频";
                        sql += "(N'"+endname.Split('.')[0]+"', N'', N'综合', N'"+ filetype + "', N'http://10.194.5.102/NetExam/upload/"+ urlname + "', -1, N'"+endname.Substring(endname.LastIndexOf('.')+1).ToUpper()+"', NULL, 0, N'测试工种', N'测试岗位'),"+ System.Environment.NewLine;
                    }
                }

           }
            File.WriteAllText(@"e:\1.txt", sql) ;

        }
    }
}
