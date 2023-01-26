<!--#include virtual="/config/conn.asp"-->

<%
response.write("<select name='fasesor' class='anchox2'>")

    response.write("<option value=0>por asignar</option> ")

             strSQL="select * from OSLP where Active='Y' and U_selector is not null "
             'response.write  strSQL            
             rsUpdateEntry.Open strSQL, adoCon2  'DMX'

        while not rsUpdateEntry.EOF
          if rsUpdateEntry("slpcode")=13 then
              response.write("<option value='"& rsUpdateEntry("slpcode") &"' SELECTED>"& rsUpdateEntry("SlpName") &"</option> ")
          else
             response.write("<option value='"& rsUpdateEntry("slpcode") &"'>"& rsUpdateEntry("SlpName") &"</option> ")
          end if
          rsUpdateEntry.movenext
        wend
        rsUpdateEntry.close
        response.write("</select>")
%>


