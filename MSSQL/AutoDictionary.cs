using Dapper;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using MSSQL.Models;
using MSSQL.Util;

namespace MSSQL
{
    public class AutoDictionary
    {
        private static string _path = ConfigurationManager.AppSettings["path"];
        /// <summary>
        /// 程序启动
        /// </summary>
        public static void Start()
        {
            try
            {
                var db = new SqlConnection(ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString);
                var tableNames = db.Query<string>("select name from sys.tables go;");
                var dict = new Dictionary<string, List<Dict>>();
                foreach (var tableName in tableNames)
                {
                    const string sql = "SELECT 表名 = d.name,表说明 = isnull(f.value,''),字段序号 = a.colorder,字段名 = a.name,标识 = case when COLUMNPROPERTY( a.id,a.name,'IsIdentity')=1 then '√'else '' end,主键 = case when exists(SELECT 1 FROM sysobjects where xtype='PK' and parent_obj=a.id and name in (SELECT name FROM sysindexes WHERE indid in( SELECT indid FROM sysindexkeys WHERE id = a.id AND colid=a.colid))) then '√' else '' end,类型 = b.name,占用字节数 = a.length,长度 = COLUMNPROPERTY(a.id,a.name,'PRECISION'),小数位数   = isnull(COLUMNPROPERTY(a.id,a.name,'Scale'),0),允许空 = case when a.isnullable=1 then '√'else '' end,默认值 = isnull(e.text,''),字段说明 = isnull(g.[value],'') FROM syscolumns a left join systypes b on a.xusertype=b.xusertype inner join sysobjects d on a.id=d.id  and d.xtype='U' and  d.name<>'dtproperties' left join syscomments e on a.cdefault=e.id left join sys.extended_properties g on a.id=G.major_id and a.colid=g.minor_id left join sys.extended_properties f on d.id=f.major_id and f.minor_id=0 where d.name= @tableName order by a.id,a.colorder;";
                    var columns = db.Query<Dict>(sql, new { tableName = tableName }).ToList();
                    dict.Add(tableName, columns);
                }
                #region 设置路径      
                if (string.IsNullOrEmpty(_path) || !BaseTool.IsValidPath(_path))
                {
                    _path = AppDomain.CurrentDomain.BaseDirectory;
                }
                _path= Path.Combine(_path, $"{db.Database}.xlsx");
                #endregion
                GeneratedForm(dict);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }
        private static void GeneratedForm(Dictionary<string, List<Dict>> dict)
        {
            Application app = null;
            Workbook workBook = null;
            Worksheet worksheet;
            Worksheet sheet = null;
            Range range;
            try
            {
                app = new Application()
                {
                    Visible = true,
                    DisplayAlerts = false
                };
                app.Workbooks.Add(true);
                if (File.Exists(_path))
                {
                    File.Delete(_path);
                }
                workBook = app.Workbooks.Add(Missing.Value);
                app.Sheets.Add(Missing.Value, Missing.Value, dict.Count);
                worksheet = (Worksheet)workBook.Sheets[1];//数据库字典汇总表
                worksheet.Name = "数据库字典汇总表";
                worksheet.Cells[1, 1] = "数据库字典汇总表";
                worksheet.Cells[2, 1] = "编号";
                worksheet.Cells[2, 2] = "表英文名称";
                worksheet.Cells[2, 3] = "表中文名称";
                worksheet.Cells[2, 4] = "数据说明";
                worksheet.Cells[2, 5] = "表结构描述(页号)";
                var type = typeof(Dict);
                var properties = (from p in type.GetProperties()
                                  where p.Name != "表名" && p.Name != "表说明"
                                  select p).ToArray();
                for (var i = 0; i < dict.Count; i++)
                {
                    var list = dict.ElementAt(i).Value;
                    sheet = (Worksheet)workBook.Sheets[i + 2];//数据表
                    sheet.Name = $"{(101d + i) / 100:F}";
                    sheet.Cells[1, 1] = "数据库表结构设计明细";
                    sheet.Cells[2, 1] = $"表名：{dict.ElementAt(i).Key}";
                    sheet.Cells[3, 1] = list[0].表说明;
                    for (var j = 0; j < properties.Count(); j++)
                    {
                        if (properties[j].Name != "表名" && properties[j].Name != "表说明")
                        {
                            sheet.Cells[4, j + 1] = properties[j].Name;
                            for (var k = 0; k < list.Count; k++)
                            {
                                sheet.Cells[k + 5, j + 1] = type.GetProperty(properties[j].Name).GetValue(list[k], null);
                            }
                        }
                    }
                    worksheet.Cells[i + 3, 1] = i + 1;
                    worksheet.Cells[i + 3, 2] = dict.ElementAt(i).Key;
                    worksheet.Cells[i + 3, 3] = dict.ElementAt(i).Value[0].表说明;
                    worksheet.Cells[i + 3, 4] = string.Empty;
                    worksheet.Cells[i + 3, 5] = $"表{sheet.Name}";
                    #region  数据表样式 
                    range = sheet.Range[sheet.Cells[1, 1], sheet.Cells[list.Count + 4, properties.Count()]];//选取单元格
                    range.VerticalAlignment = XlVAlign.xlVAlignCenter;//垂直居中设置 
                    range.EntireColumn.AutoFit();//自动调整列宽
                    range.Borders.LineStyle = XlLineStyle.xlContinuous;//所有框线 
                    range.Borders.Weight = XlBorderWeight.xlMedium;//边框常规粗细
                    range.Font.Name = "宋体";//设置字体 
                    range.Font.Size = 14;//字体大小  
                    range.NumberFormatLocal = "@";
                    range = sheet.Range[sheet.Cells[4, 1], sheet.Cells[list.Count + 4, properties.Count()]];//选取单元格
                    range.HorizontalAlignment = XlHAlign.xlHAlignCenter;//水平居中设置                   
                    range = sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, properties.Count()]];//选取单元格
                    range.Merge(Missing.Value);
                    range.HorizontalAlignment = XlHAlign.xlHAlignCenter;//水平居中设置
                    range.Font.Bold = true;//字体加粗
                    range.Font.Size = 24;//字体大小                           
                    range = sheet.Range[sheet.Cells[2, 1], sheet.Cells[2, properties.Count()]];//选取单元格
                    range.Merge(Missing.Value);
                    range = sheet.Range[sheet.Cells[3, 1], sheet.Cells[3, properties.Count()]];//选取单元格
                    range.Merge(Missing.Value);
                    #endregion                  
                }
                #region  汇总表样式             
                range = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[dict.Count + 2, 5]];//选取单元格
                range.ColumnWidth = 30;//设置列宽
                range.HorizontalAlignment = XlHAlign.xlHAlignCenter;//水平居中设置 
                range.VerticalAlignment = XlVAlign.xlVAlignCenter;//垂直居中设置 
                range.Borders.LineStyle = XlLineStyle.xlContinuous;//所有框线
                range.Borders.Weight = XlBorderWeight.xlMedium;//边框常规粗细 
                range.Font.Name = "宋体";//设置字体 
                range.Font.Size = 14;//字体大小 
                range.NumberFormatLocal = "@";
                range = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, 5]];//选取单元格
                range.Merge(Missing.Value);
                range.Font.Bold = true;//字体加粗
                range.Font.Size = 24;//字体大小 
                #endregion           
                sheet?.SaveAs(_path);
                worksheet.SaveAs(_path);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            finally
            {
                workBook?.Close();
                app?.Quit();
                range = null;
                sheet = null;
                worksheet = null;
                workBook = null;
                app = null;
                GC.Collect();
            }
        }
    }
}
