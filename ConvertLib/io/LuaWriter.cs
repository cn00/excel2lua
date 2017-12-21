using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.IO;
using LuaInterface;
using GameManager.functions.io;

namespace XlsToLua
{
    public class LuaWriter
    {
        public static bool writeLua(string sheetname, DataTable dt, string path)
        {
            string json = TableDataHandler.convertToJSON(dt);
            if(json == null)
            {
                return false;
            }

            Lua luaVM = new Lua();
            luaVM.DoFile("./lua/main.lua");
            object[] obj = luaVM.GetFunction("jsonToLua").Call(json);

            string strLua = "";
            // 写表头
            //strLua += "--------------------------------------------------------\n";
            //strLua += "--info:" + sheetname + "\n";
            //strLua += "--author:" + Environment.UserName + "\n";
            //strLua += "--date:" + DateTime.Now + "\n";
            //strLua += "--------------------------------------------------------\n";
            strLua += "\nLua_Table=Lua_Table or {}\nLua_Table.";

            strLua += sheetname + "={\n";
            strLua += obj[0];
            strLua += "}\nreturn Lua_Table." + sheetname;
            // 保存lua文件
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
    }
}
