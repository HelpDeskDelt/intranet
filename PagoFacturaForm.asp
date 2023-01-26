<!--#include virtual="/config/conn.asp"-->


<%
     
           Response.Expires = -1
           vRS=trim(request("variable"))
           response.write("<input type='hidden' name='RS' value="&vRS&">")
 
          response.write("<P>&nbsp;</P>")
          CreateTable(960)
          response.write("<tr>")
             response.write("<td width='30%' class='td-r subtitulo'>Seleccione socio de negocio:</td>")
             response.write("<td width='70%' class='td-l'><select name='SN'>")
     
             strSQL="select * from OCRD where substring(cardcode,1,1) in ('P','A') and frozenFor='N' order by cardName" 
             select case vRS
                case "DMX"  rsUpdateEntry.Open strSQL, adoCon2 'DMX'
                case "DFW"  rsUpdateEntry.Open strSQL, adoCon3 'DFW'
                case "MME"  rsUpdateEntry.Open strSQL, adoCon5 'MME'
             end select
        
             while not rsUpdateEntry.EOF
                response.write("<option value='"& rsUpdateEntry("cardCode") &"'>"& rsUpdateEntry("cardName") &"("& rsUpdateEntry("cardcode") &")</option>")
                rsUpdateEntry.MoveNext
             wend
             rsUpdateEntry.close
             response.write("/select></td>")      
          response.write("</tr>")

          response.write("<tr>")
             response.write("<td class='td-r subtitulo'>Es una carga masiva? *</td>")
             response.write("<td class='td-l'><input type='checkbox' name='fmasivo'>")
          response.write("</tr>")
          closeTable()
          response.write("<P class='td-c'><font style='font-size: 10px'>* utilice esta opcion si tiene 5+ facturas por programar</font></P>")
%>