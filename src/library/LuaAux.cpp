// LuaAux.cpp: implementation of the LuaAux class.
//
//////////////////////////////////////////////////////////////////////

#include <stdio.h>
#include <assert.h>
#include <string.h>

#include "LuaAux.h"
#include "LuaCompat.h"
#include <string>


//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

LuaAux::LuaAux()
{

}

LuaAux::~LuaAux()
{

}

/*
 * Prints the lua stack
 */

void LuaAux::printLuaStack(lua_State *L)
{
  int size = lua_gettop(L);
  int i = 0;

  for(i = size; i > 0; i--)
  {
    switch(lua_type(L,i))
    {
    case LUA_TNUMBER:
      printf("%d: number = %g", i, lua_tonumber(L, i));
      break;

    case LUA_TSTRING:
      printf("%d: string = \"%s\"", i, lua_tostring(L, i));
      break;

    case LUA_TTABLE:
      printf("%d: table, tag = %p", i, luaCompat_getType2(L, i));
      break;

    case LUA_TUSERDATA:
      printf("%d: userdata = %p, tag = %p", i,
        luaCompat_getPointer(L, i), luaCompat_getType2(L, i));
      break;

    case LUA_TNIL:
      printf("%d: nil", i);
      break;

    case LUA_TBOOLEAN:
      if(lua_toboolean(L, i))
        printf("%d: boolean = true", i);
      else
        printf("%d: boolean = false", i);

      break;

    default:
      printf("%d: unknown type (%d)", i, lua_type(L, i));
      break;
    }      

    printf("\n");
  }

  printf("\n");
}

void LuaAux::printPreDump(int expected) {
  printf("STACK DUMP\n");
  printf("Expected size: %i\n",expected);
}

void LuaAux::printLuaTable(lua_State *L, stkIndex t)
{

  lua_pushnil(L);  /* first key */
  while (lua_next(L, t) != 0) {
   /* `key' is at index -2 and `value' at index -1 */
   printf("%s - %s\n",
     lua_tostring(L, -2), lua_typename(L, lua_type(L, -1)));
   lua_pop(L, 1);  /* removes `value'; keeps `index' for next iteration */
  }
}

tStringBuffer LuaAux::makeLuaErrorMessage(int return_value, const char* msg)
{
    switch (return_value)
    {
    case 0:
        return "No error";

    case LUA_ERRRUN:
    {
        std::string out = "Lua runtime error";
        if (msg && *msg)
        {
            out += ": ";
            out += msg;
        }
        return out.c_str();
    }

    case LUA_ERRMEM:
        return "Lua memory allocation error.";

    case LUA_ERRERR:
        return "Lua error: error during error handler execution.";

    default:
        return "Unknown Lua error.";
    }
}

