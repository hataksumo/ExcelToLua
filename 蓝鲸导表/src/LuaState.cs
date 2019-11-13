﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Lua = NLua;

namespace ExcelToLua
{
    static class LuaState
    {
        private static Lua.Lua lua_instence = null;
        public static Lua.Lua Init(string v_main)
        {
            lua_instence = new Lua.Lua();
            lua_instence.DoFile(v_main);
            return lua_instence;
        }

        public static void SetPath(string v_path)
        {
            Lua.LuaFunction fun = lua_instence.GetFunction("setPath");
            if(fun != null)
                fun.Call(v_path);
        }

        public static void DoMain()
        {
            Lua.LuaFunction fun = lua_instence.GetFunction("main");
            try
            {
                fun.Call();
            }
            catch (Exception ex)
            {
                Debug.Exception("执行main报错，信息是" + ex.ToString());
            }
            
        }

        public static Lua.Lua Get_Instenct()
        {
            return lua_instence;
        }
    }
}
