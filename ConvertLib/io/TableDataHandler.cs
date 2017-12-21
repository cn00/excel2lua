using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace GameManager.functions.io
{
	enum CellType
	{
		SIMPLE,
		ARRAY
	}

	class HeadDefine
	{
		public CellType type;
		public string headName;
		public string subName;
		public int index;
		public int columnIndex;
	}

	class TableDataHandler
	{
		public static String convertToJSON(DataTable dt)
		{
			List<HeadDefine> lstTableDefine = new List<HeadDefine>();
			// 先获取第一行的数据作为表头
			for (int i = 0; i < dt.Columns.Count; ++i )
			{
				string[] headSplit = dt.Rows[0][i].ToString().Split('[', ']');
				HeadDefine define = new HeadDefine();
				define.headName = headSplit[0];
				define.columnIndex = i;
				// 简单类型
				if (headSplit.Length == 1)
				{
					define.type = CellType.SIMPLE;
				}
				// 数组类型
				else
				{
					define.type = CellType.ARRAY;
					define.index = Convert.ToInt32(headSplit[1]);

					// 检查是否有子内容
					string[] subSplit = define.headName.Split('<', '>');
					if (subSplit.Length > 1)
					{
						define.headName = subSplit[0];
						define.subName = subSplit[1];

						if (define.headName == "")
						{
							Console.Error.Write(" 标题第" + (i + 1) + "列为空");
							break;
						}
					}
				}

				lstTableDefine.Add(define);
			}

			SortedDictionary<long, SortedDictionary<String, Object>> tableData = new SortedDictionary<long, SortedDictionary<String, Object>>();
			// 获取表数据
			// 1.循环每一行，向Dictionary中添加数据
			for (int i = 1; i < dt.Rows.Count; ++i)
			{
				SortedDictionary<string, object> dict_Row = new SortedDictionary<string, object>();
				// 2.循环当前行的每一列数据
				for (int j = 0; j < lstTableDefine.Count; )
				{
					if(lstTableDefine[j].type == CellType.SIMPLE)
					{
						dict_Row.Add(lstTableDefine[j].headName, dt.Rows[i][lstTableDefine[j].columnIndex]);
						j++;
					}
					else if(lstTableDefine[j].type == CellType.ARRAY && lstTableDefine[j].index == 1)
					{
						string arrayName = lstTableDefine[j].headName;
						List<object> lst_Array = new List<object>();
						// 3.循环添加数组的每一项
						while (j < lstTableDefine.Count && lstTableDefine[j].headName == arrayName)
						{
							int currentIndex = lstTableDefine[j].index;
							
							// 4.添加数组的每一项中的每一个元素
							// 如果是复杂类型数组
							if (lstTableDefine[j].subName != null)
							{
								Dictionary<string, object> dict_ArrayElement = new Dictionary<string,object>();
								while (j < lstTableDefine.Count && lstTableDefine[j].headName == arrayName && lstTableDefine[j].index == currentIndex)
								{
									dict_ArrayElement.Add(lstTableDefine[j].subName, dt.Rows[i][lstTableDefine[j].columnIndex]);
									j++;
								}
								lst_Array.Add(dict_ArrayElement);
							}
							// 如果是简单类型数组
							else
							{
								List<object> lst_ArrayElement = new List<object>();
								while (j < lstTableDefine.Count && lstTableDefine[j].headName == arrayName)
								{
									lst_ArrayElement.Add(dt.Rows[i][lstTableDefine[j].columnIndex]);
									j++;
								}
								lst_Array = lst_ArrayElement;
							}
							
						}
						dict_Row.Add(arrayName, lst_Array);
					}
				}

                Int64 key = 0;
                if(Int64.TryParse(dt.Rows[i][0].ToString(), out key))
                {
                    if(tableData.ContainsKey(key))
                    {
                        Console.Error.Write(" 第" + (i + 2) + "行包含重复的主键：" + Convert.ToInt64(dt.Rows[i][0]));
                        return null;
                    }
				    else
				    {
                        tableData.Add(Convert.ToInt64(dt.Rows[i][0]), dict_Row);
                    }
                }
			}

			return JsonConvert.SerializeObject(tableData,Formatting.Indented);
		}
	}
}