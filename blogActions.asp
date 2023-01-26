<!--#include virtual="/config/conn.asp"-->

<%

  Response.Expires = -1

  select case request("control")
  case "1"
      response.write("<form id='blog' action='ShowContent.asp' method='POST'>")
      Response.Write("<input type='hidden' name='contenido' value='114'>")
      Response.Write("<input type='hidden' name='control' value='1'>")
      CreateTable(800)
      
      response.write("<tr><td class='td-r subtitulo' width='18%'>descrp breve del hilo:</td>")
      response.write("<td class='td-l'><input type='text' name='fhilo' class='anchox4' maxlenght='140'></td></tr>")
      response.write("<tr><td class='td-r subtitulo'>Fecha de inicio:</td>")
      response.write("<td class='td-l'><input type='date' name='ffechai'></td></tr>")

      response.write("<tr><td colspan=2>&nbsp;</td></tr>")
      response.write("<tr><td colspan=2 class='td-c'><input type='submit' value='agregar'></form></td></tr>")
          
      closeTable
  case "2"

      response.write("<div id='blog'>")
      ShowHilo
      response.write("<tr><td colspan=2>&nbsp;</td></tr>")
      response.write("<tr><td colspan=2 class='td-c'><a href='ShowContent.asp?contenido=114' class='button'>regresar</a></td></tr>")
      closeTable    
      response.write("</div>")

      case "3"

          response.write("<form id='blog' action='ShowContent.asp' method='POST'>")
          Response.Write("<input type='hidden' name='contenido' value='114'>")
          Response.Write("<input type='hidden' name='control' value='2'>")
          Response.Write("<input type='hidden' name='father' value='"&request("ID")&"'>")

          CreateTable(920)
          response.write("<tr><td class='td-l'><textarea name='fentrada' rows=6 cols=92></textarea></td><td class='td-c'><input type='submit' value='agregar'></form></tr></table><br><br>")

          ShowHilo
    

end select




     sub ShowHilo

     	    strSQL="select *,.dbo.fn_GetMonthName(fechaI,'spanish') as fecha  from BlogVentas where ID="& request("ID")&"  order by ID desc"
  			rsUpdateEntry.Open strSQL, adoCon

  			strSQL="select *,.dbo.fn_GetMonthName(fechaI,'spanish') as fecha  from BlogVentas where Father="& request("ID")&"  order by ID desc"
      		rsUpdateEntry2.Open strSQL, adoCon

      
      CreateTable(920)

      response.write("<tr><td class='td-r subtitulo'>Hilo:</td>")
      response.write("<td class='td-l fontbig strong'>"& rsUpdateEntry("descripcion") &"<font class='fonttiny' color=red><I> ("& rsUpdateEntry("fecha") &")</I></font> &nbsp;&nbsp;&nbsp;")
        %>
        <a href="Javascript:sendReq('blogActions.asp','control,ID','3,<%=request("ID")%>','blog')"><img src="/img/b_edit.png"> + entrada </a></td></tr> <%

      i=1
      while not rsUpdateEntry2.EOF
      	  if i=1 then
          		response.write("<tr><td class='td-l padding-l' colspan=2>"& rsUpdateEntry2("Entrada") &"<font class='fonttiny' color=red><I> ("& rsUpdateEntry("fecha") &")</I></font></td>")
          else
               response.write("<tr><td class='td-r padding-r' colspan=2>"& rsUpdateEntry2("Entrada") &"<font class='fonttiny' color=red><I> ("& rsUpdateEntry("fecha") &")</I></font></td>")
          end if
          response.write("<tr><td class='td-l' colspan=2>&nbsp;</td>")
          rsUpdateEntry2.movenext
          i=i+1
          if i=3 then i=1
      wend
      rsUpdateEntry2.close

      response.write("<tr><td class='td-l' colspan=2>&nbsp;</td>")


     end sub

  

%>