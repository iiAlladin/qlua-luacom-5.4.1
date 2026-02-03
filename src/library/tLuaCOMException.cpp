// tLuaCOMException.cpp: implementation of the tLuaCOMException class.
//
//////////////////////////////////////////////////////////////////////

//#include <iostream.h>
#include <string.h>
#include <stdio.h>

#include "tLuaCOMException.h"
#include "tUtil.h"
#include <string>

char const * const tLuaCOMException::messages[] =  
     {
       "Internal Error",
       "Parameter(s) out of range",
       "Type conversion error",
       "COM error",
       "COM exception",
       "Unsupported feature required",
       "Windows error",
       "LuaCOM error",
       "Memory allocation error"
     };


//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

tLuaCOMException::tLuaCOMException(Errors p_code, const char *p_file, int p_line,
                                   const char *p_usermessage)
{
  code = p_code;

  file = tStringBuffer(p_file);
  line = p_line;
  usermessage = tStringBuffer(p_usermessage);

  tUtil::log("luacom", getMessage());
}

tLuaCOMException::~tLuaCOMException()
{
}



tStringBuffer tLuaCOMException::getMessage()
{
    std::string msg = messages[code]; 
    if (file && *file) {
        msg += ":("; 
        msg += file; 
        msg += ","; 
        msg += std::to_string(line);
        msg += ")"; 
    } 
    msg += ":"; 
    if (usermessage) msg += usermessage; 
    return tStringBuffer(msg.c_str());
}

tStringBuffer tLuaCOMException::GetErrorMessage(DWORD errorcode)
{
  return tUtil::GetErrorMessage(errorcode);
}
