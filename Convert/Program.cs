using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.IO;
using System.Text;

namespace Convert
{
    class Program
    {
        public static bool writeLua(string sheetname, ISheet sheet, string path)
        {
            IRow headerRow = sheet.GetRow(1);
            int columnCount = headerRow.LastCellNum;
            int rowCount = sheet.LastRowNum;
            string heads = "head={";

            for(int i = 0; i < columnCount; ++i)
            {
                string head = headerRow.GetCell(i).StringCellValue.Replace("(", "_").Replace("]", "_").Replace("[", "_").Replace("]", "_");
                heads += "\n\t" + head + "=" + (i + 1) + ",";
            }
            heads += "\n},";

            string body = "";
            try
            {
                for(int i = 2; i < rowCount; ++i)
                {
                    IRow row = sheet.GetRow(i);
                    if(row == null)
                        continue;
                    var cell0 = row.GetCell(0);
                    if(cell0.CellType == CellType.Blank
                        || (cell0.CellType == CellType.String && cell0.StringCellValue == ""))
                        continue;

                    body += "\n[" + cell0.ToString() + "]" + "={";
                    for(int j = 0; j < row.LastCellNum; ++j)
                    {
                        var cell = row.GetCell(j) ?? row.CreateCell(j);
                        switch(cell.CellType)
                        {
                        case CellType.Unknown:
                            body += "\"\"";
                            break;
                        case CellType.Numeric:
                            body += cell.NumericCellValue;
                            break;
                        case CellType.String:
                            body += "\"" + cell.StringCellValue + "\"";
                            break;
                        case CellType.Formula:
                            switch(cell.CachedFormulaResultType)
                            {
                            case CellType.Unknown:
                                body += "\"Unknown\"";
                                break;
                            case CellType.Numeric:
                                body += cell.NumericCellValue;
                                break;
                            case CellType.String:
                                body += "\"" + cell.StringCellValue + "\"";
                                break;
                            case CellType.Blank:
                                body += "\"\"";
                                break;
                            case CellType.Boolean:
                                body += cell.BooleanCellValue;
                                break;
                            case CellType.Error:
                                body += "\"Error\"";
                                break;
                            default:
                                break;
                            }
                            break;
                        case CellType.Blank:
                            body += "\"\"";
                            break;
                        case CellType.Boolean:
                            body += cell.BooleanCellValue;
                            break;
                        case CellType.Error:
                            body += "\"Error\"";
                            break;
                        default:
                            break;
                        }
                        body += ",\t";
                    }
                    body += "},";
                }
            }
            catch(Exception e)
            {
                Console.WriteLine(e);
            }

            string tail = "\nfor k,v in pairs(Lua_Table." + sheetname + ") do\n\tif k ~= \"head\" and type(v) == \"table\" then\n\t\tsetmetatable(v,{\n\t\t__newindex=function(t,kk) print(\"warning: attempte to change a readonly table\") end,\n\t\t__index=function(t,kk)\n\t\t\tif Lua_Table." + sheetname + ".head[kk] ~= nil then\n\t\t\t\treturn t[Lua_Table." + sheetname + ".head[kk]]\n\t\t\telse\n\t\t\t\tprint(\"err: \\\"Lua_Table." + sheetname + "\\\" have no field [\"..kk..\"]\")\n\t\t\t\treturn nil\n\t\t\tend\n\t\tend})\n\tend\nend";

            string strLua = "";
            strLua += "\nLua_Table=Lua_Table or {}\nLua_Table.";

            strLua += sheetname + "={\n";
            strLua += heads;
            strLua += body;
            strLua += "\n}";
            strLua += tail;
            strLua += "\nreturn Lua_Table." + sheetname;

            path += "\\" + sheetname + ".lua";
            UTF8Encoding utf8 = new UTF8Encoding(false);
            StreamWriter sw;
            using(sw = new StreamWriter(path, false, utf8))
            {
                sw.Write(strLua);
            }
            sw.Close();
            return true;
        }
        private static void Main(string[] args)
        {
            if(args.Length < 2)
            {
                Console.Write("usage:\n\t Convert.exe excel_输入目录 lua_输出目录\n");
            }
            else
            {
                string inPath = args[0];
                string outPath = args[1];
                long lastWriteTime = 0;
                long newWriteTime = 0;

                var outDir = new DirectoryInfo(outPath);
                if(!outDir.Exists)
                {
                    outDir.Create();
                }

                if(new FileInfo(inPath + @"\lastWriteTime").Exists && args.Length == 2)
                {
                    StreamReader reader = new StreamReader(inPath + @"\lastWriteTime");
                    if(reader != null)
                    {
                        string s = reader.ReadLine();
                        reader.Close();
                        //lastWriteTime = DateTime.ParseExact(s, "yyyy_MM_dd-HH_mm_ss", CultureInfo.InvariantCulture);
                        if(!long.TryParse(s, out lastWriteTime))
                        {
                            lastWriteTime = 0;
                        }
                        newWriteTime = lastWriteTime;
                    }
                }
                Console.WriteLine("上次转换时间：" + DateTime.FromFileTimeUtc(lastWriteTime));

                DirectoryInfo info = new DirectoryInfo(inPath);
                int errno = 0;
                foreach(FileInfo info2 in info.GetFiles("*.xls*"))
                {
                    //if(info2.LastWriteTime.ToFileTimeUtc() <= lastWriteTime)
                    //{
                    //    Console.WriteLine(info2.LastWriteTime.ToString("yyyy_MM_dd-HH_mm_ss") + " 无需更新【" + info2.Name + "】");
                    //}
                    //else
                    {
                        if(info2.LastWriteTime.ToFileTimeUtc() > newWriteTime)
                        {
                            newWriteTime = info2.LastWriteTime.ToFileTimeUtc();
                        }
                        Console.WriteLine(info2.LastWriteTime.ToString("yyyy_MM_dd-HH_mm_ss") + " 开始转换【" + info2.Name + "】... ");

                        try
                        {
                            var fstream = new FileStream(info2.FullName, FileMode.Open);
                            IWorkbook workbook = null;
                            if(info2.Name.EndsWith(".xlsx"))
                                workbook = new XSSFWorkbook(fstream);
                            else
                                workbook = new HSSFWorkbook(fstream);

                            for(int i = 0; i < 1/*workbook.NumberOfSheets*/; ++i)
                            {
                                var sheet = workbook.GetSheetAt(i);
                                var sheetname = sheet.SheetName;
                                Console.Write("  " + sheetname);

                                //DataTable ds = ExcelRender.RenderFromExcel(sheet, 1);

                                if(writeLua(sheetname.Replace("$", ""), sheet, outPath))
                                {
                                    Console.Write(" 成功\n");
                                }
                                else
                                {
                                    Console.Write(" 失败!!!!!!!!!\n");
                                    ++errno;
                                }
                            }
                        }
                        catch(Exception e)
                        {
                            Console.Error.Write("error：" + e.Message);
                        }
                    }
                } // for
                StreamWriter writer = new StreamWriter(inPath + @"\lastWriteTime", false);
                writer.Write(
                    newWriteTime
                );
                writer.Close();
                if(errno > 0)
                {
                    Console.WriteLine("转表 {0} 个错误，按 Enter 退出", errno);
                    Console.ReadLine();
                }
                else
                {
                    System.Threading.Thread.Sleep(2000);
                }
            }
        }
    }
}
