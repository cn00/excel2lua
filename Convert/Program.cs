using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.IO;
using System.Text;

namespace Convert
{
    public static class CellExtension
    {
        public static string SValue(this ICell cell, CellType? FormulaResultType = null)
        {
            string svalue = "";
            var cellType = FormulaResultType ?? cell.CellType;
            switch(cellType)
            {
            case CellType.Unknown:
                svalue = "nil";
                break;
            case CellType.Numeric:
                svalue = cell.NumericCellValue.ToString();
                break;
            case CellType.String:
                svalue = "\"" + cell.StringCellValue
                    .Replace("\n", "\\n")
                    .Replace("\t", "\\t")
                    .Replace("\"", "\\\"") + "\"";
                break;
            case CellType.Formula:
                svalue = cell.SValue(cell.CachedFormulaResultType);
                break;
            case CellType.Blank:
                svalue = "nil";
                break;
            case CellType.Boolean:
                svalue = cell.BooleanCellValue.ToString();
                break;
            case CellType.Error:
                svalue = "nil";
                break;
            default:
                break;
            }
            return svalue;
        }
    }
    class Program
    {
        public static bool writeLua(string sheetname, ISheet sheet, string path)
        {
            IRow headerRow = sheet.GetRow(1);
            int columnCount = headerRow.LastCellNum;
            int rowCount = sheet.LastRowNum;
            string keys = "KEY={";

            for(int i = 0; i < columnCount; ++i)
            {
                string head = headerRow.GetCell(i).StringCellValue.Replace("(", "_").Replace("]", "_").Replace("[", "_").Replace("]", "_");
                keys += "\n\t" + head + "=" + (i + 1) + ",";
            }
            keys += "\n},";

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
                        body += cell.SValue() + ",\t";
                    }
                    body += "},";
                }
            }
            catch(Exception e)
            {
                Console.WriteLine(e);
            }

            string tail = "\nfor k,v in pairs(Lua_Table." + sheetname + ") do"
                + "\n\tif k ~= \"KEY\" and type(v) == \"table\" then"
                + "\n\t\tsetmetatable(v,{"
                + "\n\t\t__newindex=function(t,kk) print(\"warning: attempte to change a readonly table\") end,"
                + "\n\t\t__index=function(t,kk)"
                + "\n\t\t\tif Lua_Table." + sheetname + ".KEY[kk] ~= nil then"
                + "\n\t\t\t\treturn t[Lua_Table." + sheetname + ".KEY[kk]]"
                + "\n\t\t\telse"
                + "\n\t\t\t\tprint(\"err: \\\"Lua_Table." + sheetname + "\\\" have no field [\"..kk..\"]\")"
                + "\n\t\t\t\treturn nil"
                + "\n\t\t\tend"
                + "\n\t\tend})"
                + "\n\tend"
                + "\nend";

            string strLua = "-- usage: \n--\tLua_Table." + sheetname + "[id][KEY] \n--\tLua_Table." + sheetname + "[id].KEY\n";
            strLua += "\nLua_Table=Lua_Table or {}\nLua_Table.";

            strLua += sheetname + "={\n";
            strLua += keys;
            strLua += body;
            strLua += "\n}";
            strLua += tail;
            strLua += "\nreturn Lua_Table." + sheetname;

            path += "/" + sheetname + ".lua";
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
            string inPath = "test";
            string outPath = "out";
            if(args.Length < 2)
            {
                Console.Write("usage:\n\t Convert.exe excel_输入目录 lua_输出目录\n请输入:");
                inPath = Console.ReadLine();
            }
            else
            {
                inPath = args[0];
                outPath = args[1];
            }
            long lastWriteTime = 0;
            long newWriteTime = 0;

            var outDir = new DirectoryInfo(outPath);
            if(!outDir.Exists)
            {
                outDir.Create();
            }

            if(new FileInfo(inPath + @"/lastWriteTime").Exists && args.Length == 2)
            {
                StreamReader reader = new StreamReader(inPath + @"/lastWriteTime");
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

            DirectoryInfo dir = new DirectoryInfo(inPath);
            int errno = 0;
            try
            {
                foreach(FileInfo file in dir.GetFiles("*.xls*"))
                {
                    if(file.LastWriteTime.ToFileTimeUtc() <= lastWriteTime)
                    {
                        Console.WriteLine(file.LastWriteTime.ToString("yyyy_MM_dd-HH_mm_ss") + " 无需更新【" + file.Name + "】");
                    }
                    else
                    {
                        if(file.LastWriteTime.ToFileTimeUtc() > newWriteTime)
                        {
                            newWriteTime = file.LastWriteTime.ToFileTimeUtc();
                        }
                        Console.WriteLine(file.LastWriteTime.ToString("yyyy_MM_dd-HH_mm_ss") + " 开始转换【" + file.Name + "】... ");

                        var fstream = new FileStream(file.FullName, FileMode.Open);
                        IWorkbook workbook = null;
                        if(file.Name.EndsWith(".xlsx"))
                            workbook = new XSSFWorkbook(fstream);
                        else
                            workbook = new HSSFWorkbook(fstream);

                        for(int i = 0; i < workbook.NumberOfSheets; ++i)
                        {
                            var sheet = workbook.GetSheetAt(i);
                            var sheetname = sheet.SheetName;
                            Console.Write("  " + sheetname);

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
                } // for
            }
            catch(Exception e)
            {
                ++errno;
                Console.Error.Write("error：" + e.Message);
            }
            StreamWriter writer = new StreamWriter(inPath + @"/lastWriteTime", false);
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
