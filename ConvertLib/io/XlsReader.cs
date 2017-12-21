using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;
using GameManager.functions;

namespace XlsToLua
{
    public class XlsReader
    {
        public static DataSet ReadXls(OleDbConnection conn, string sheetname)
        {
            DataSet ds = new DataSet();
            try
            {
                string strcom = "select * from [" + sheetname + "]";
                OleDbDataAdapter odda = new OleDbDataAdapter(strcom, conn);
                odda.Fill(ds);
            }
            catch(OleDbException e)
            {
                Console.Error.Write("错误信息：" + e.Message);
                return null;
            }

            return ds;
        }
    }
}
