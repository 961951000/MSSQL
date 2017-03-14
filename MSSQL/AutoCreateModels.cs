using Dapper;
using MSSQL.Util;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;

namespace MSSQL
{

    public class AutoCreateModels
    {
        /// <summary>
        /// 程序启动
        /// </summary>
        public static int Start()
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
                var count = CreateModel(dict);
                Console.WriteLine("程序执行成功，共创建{0}个模型", count);
                return count;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                return 0;
            }
        }

        private static int CreateModel(Dictionary<string, List<Dict>> dict)
        {
            var space = ConfigurationManager.AppSettings["modelnamespace"];
            var modelsPath = ConfigurationManager.AppSettings["modelpath"];
            if (!BaseTool.IsValidPath(modelsPath))
            {
                #region 应用程序路径
                var appPath = AppDomain.CurrentDomain.BaseDirectory;
                for (var i = 0; i < 2; i++)
                {
                    var sb = new StringBuilder();
                    var list = appPath.Split('\\').ToList();
                    list.Remove("");
                    list.RemoveAt(list.Count - 1);
                    foreach (var str in list)
                    {
                        sb.Append(str).Append("\\");
                    }
                    appPath = sb.ToString();
                }
                #endregion
                var arr = appPath.Split('\\');
                space = $"{arr[arr.Length - 2]}.Models";
                modelsPath = Path.Combine(@appPath, "Models");
            }
            if (!Directory.Exists(modelsPath))
            {
                Directory.CreateDirectory(modelsPath);
            }
            var count = 0;
            foreach (var tableName in dict)
            {
                var sb = new StringBuilder();
                var sb1 = new StringBuilder();
                var className = string.Empty;
                if (tableName.Key.LastIndexOf('_') != -1)
                {
                    foreach (var str in tableName.Key.Split('_'))
                    {
                        if (!string.IsNullOrEmpty(str))
                        {
                            className += str.Substring(0, 1).ToUpper() + str.Substring(1);
                        }
                    }
                }
                else
                {
                    className = tableName.Key.Substring(0, 1).ToUpper() + tableName.Key.Substring(1);
                }
                sb.Append("using System;\r\nusing System.ComponentModel.DataAnnotations;\r\nusing System.ComponentModel.DataAnnotations.Schema;\r\n\r\nnamespace ");
                sb.Append(space);
                sb.Append("\r\n{\r\n");
                var columns = tableName.Value;
                if (columns.Count > 0)
                {
                    sb.Append("\t/// <summary>\r\n");
                    sb.Append("\t/// ").Append(tableName.Value[0].表说明).Append("\r\n");
                    sb.Append("\t/// </summary>\r\n");
                }
                sb.Append("\t[Table(\"").Append(tableName.Key).Append("\")]\r\n");  //数据标记
                sb.Append("\tpublic class ");
                sb.Append(className);
                sb.Append("\r\n\t{\r\n");
                sb.Append("\t\t#region Model\r\n");
                foreach (var column in columns)
                {
                    var propertieName = string.Empty;
                    if (column.字段名.LastIndexOf('_') != -1)
                    {
                        foreach (var str in column.字段名.Split('_'))
                        {
                            if (!string.IsNullOrEmpty(str))
                            {
                                propertieName += str.Substring(0, 1).ToUpper() + str.Substring(1);
                            }
                        }
                    }
                    else
                    {
                        propertieName = column.字段名.Substring(0, 1).ToUpper() + column.字段名.Substring(1);
                    }
                    if (!string.IsNullOrEmpty(column.字段说明))
                    {
                        sb.Append("\t\t/// <summary>\r\n");
                        sb.Append("\t\t/// ").Append(column.字段说明).Append("\r\n");
                        sb.Append("\t\t/// </summary>\r\n");
                    }
                    if (!string.IsNullOrEmpty(column.主键))
                    {
                        sb.Append("\t\t[Key, Column(\"").Append(column.字段名).Append("\")]\r\n");
                    }
                    else
                    {
                        sb.Append("\t\t[Column(\"").Append(column.字段名).Append("\")]\r\n");  //数据标记
                    }
                    if (string.IsNullOrEmpty(column.类型))
                    {
                        sb.Append("\t\tpublic string " + propertieName + "\r\n\t\t{\r\n");
                    }
                    else
                    {
                        switch (column.类型.ToLower())
                        {
                            case "tinyint":
                                {
                                    sb.Append("\t\tpublic bool? " + propertieName + "\r\n\t\t{\r\n");
                                }
                                break;
                            case "smallint":
                                {
                                    sb.Append("\t\tpublic short? " + propertieName + "\r\n\t\t{\r\n");
                                }
                                break;
                            case "int":
                                {
                                    sb.Append("\t\tpublic int? " + propertieName + "\r\n\t\t{\r\n");
                                }
                                break;
                            case "bigint":
                                {
                                    sb.Append("\t\tpublic long? " + propertieName + "\r\n\t\t{\r\n");
                                }
                                break;
                            case "decimal":
                                {
                                    sb.Append("\t\tpublic decimal? " + propertieName + "\r\n\t\t{\r\n");
                                }
                                break;
                            case "timestamp":
                                {
                                    sb.Append("\t\tpublic DateTime? " + propertieName + "\r\n\t\t{\r\n");
                                }
                                break;
                            case "datetime":
                                {
                                    sb.Append("\t\tpublic DateTime? " + propertieName + "\r\n\t\t{\r\n");
                                }
                                break;
                            case "bit":
                                {
                                    sb.Append("\t\tpublic bool " + propertieName + "\r\n\t\t{\r\n");
                                }
                                break;
                            case "money":
                                {
                                    sb.Append("\t\tpublic decimal? " + propertieName + "\r\n\t\t{\r\n");
                                }
                                break;
                            case "image":
                                {
                                    sb.Append("\t\tpublic string " + propertieName + "\r\n\t\t{\r\n");
                                }
                                break;
                            case "nvarchar":
                                {
                                    sb.Append("\t\tpublic string " + propertieName + "\r\n\t\t{\r\n");
                                }
                                break;
                            case "varchar":
                                {
                                    sb.Append("\t\tpublic string " + propertieName + "\r\n\t\t{\r\n");
                                }
                                break;
                            case "text":
                                {
                                    sb.Append("\t\tpublic string " + propertieName + "\r\n\t\t{\r\n");
                                }
                                break;
                            default:
                                {
                                    sb.Append("\t\tpublic string " + propertieName + "\r\n\t\t{\r\n");
                                }
                                break;
                        }
                    }
                    sb.Append("\t\t\tset;\r\n");
                    sb.Append("\t\t\tget;\r\n");
                    sb.Append("\t\t}\r\n");
                    sb.Append("\r\n");
                    sb1.Append(propertieName);
                    sb1.Append("=\" + ");
                    sb1.Append(propertieName);
                    sb1.Append(" + \",");
                }
                sb1.Remove(sb1.Length - 5, 5);
                sb.Append("\t\tpublic override string ToString()\r\n");
                sb.Append("\t\t{\r\n");
                sb.Append("\t\t\treturn \"");
                sb.Append(sb1);
                sb.Append(";");
                sb.Append("\r\n");
                sb.Append("\t\t}\r\n");
                sb.Append("\t\t#endregion Model\r\n");
                sb.Append("\t}\r\n").Append("}");
                var filePath = Path.Combine(modelsPath, string.Format("{0}.cs", className));
                if (WriteFile(filePath, sb.ToString()))
                {
                    count++;
                }
            }
            return count;
        }
        /// <summary>
        /// 文件写入
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="text">文本内容</param>
        public static bool WriteFile(string filePath, string text)
        {
            var flag = false;
            FileStream fs = null;
            StreamWriter sw = null;
            try
            {
                if (!File.Exists(filePath))
                {
                    // 创建写入文件
                    fs = new FileStream(filePath, FileMode.Create, FileAccess.Write);
                    sw = new StreamWriter(fs);
                    sw.WriteLine(text);

                }
                else
                {
                    // 删除文件在创建
                    File.Delete(filePath);
                    fs = new FileStream(filePath, FileMode.Create, FileAccess.Write);
                    sw = new StreamWriter(fs);
                    sw.WriteLine(text);
                }
                flag = true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            finally
            {
                sw?.Flush();
                sw?.Close();
                fs?.Close();
            }
            return flag;
        }
    }
    internal class Dict
    {
        public string 表名 { get; set; }
        public string 表说明 { get; set; }
        public int? 字段序号 { get; set; }
        public string 字段名 { get; set; }
        public string 标识 { get; set; }
        public string 主键 { get; set; }
        public string 类型 { get; set; }
        public int? 占用字节数 { get; set; }
        public int? 长度 { get; set; }
        public int? 小数位数 { get; set; }
        public string 允许空 { get; set; }
        public string 默认值 { get; set; }
        public string 字段说明 { get; set; }
    }
}
