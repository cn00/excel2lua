using System;
using System.Data;
using System.IO;
using XlsToLua;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

namespace Convert
{
    class Program
    {
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
                    if(info2.LastWriteTime.ToFileTimeUtc() <= lastWriteTime)
                    {
                        Console.WriteLine(info2.LastWriteTime.ToString("yyyy_MM_dd-HH_mm_ss") + " 无需更新【" + info2.Name + "】");
                    }
                    else
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

                                DataTable ds = ExcelRender.RenderFromExcel(sheet, 1);

                                if(LuaWriter.writeLua(sheetname.Replace("$", ""), ds, outPath))
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
