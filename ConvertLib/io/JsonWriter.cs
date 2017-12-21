using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Reflection;
using GameManager.functions.io;
using System.IO;

namespace GameManager.functions
{
	public class JsonWriter
	{
		public static bool writeJson(string sheetname, DataTable dataTable, string path)
		{
			string json = TableDataHandler.convertToJSON(dataTable);
			if (json == null)
			{
				return false;
			}
			// 保存json文件
			path += "\\" + sheetname + ".json";
			UTF8Encoding utf8 = new UTF8Encoding(false);
			StreamWriter sw;
			using (sw = new StreamWriter(path, false, utf8))
			{
				sw.Write(json);
			}
			sw.Close();
			return true;
		}
	}
}
