<!--#include virtual="/config/conn.asp"-->

<%
     
           Response.Expires = -1          

           strSQL="select * from cat_entidades where id_entidad='"& request("variable")  &"' "
           rsUpdateEntry5.Open strSQL, adoCon            

           strSQL="select * from cat_municipios where id_entidad='"& request("variable")  &"' order by ciudad"
           'response.write strSQL
           rsUpdateEntry6.Open strSQL, adoCon 
       
           response.write("<form id='peajes' method='post' action='ShowContent.asp'>")
           response.write("<input type='hidden' name='contenido' value='109'>")  
           response.write("<input type='hidden' name='action' value=1>")
           response.write("<input type='hidden' name='entidad_destino' value='"& request("variable") &"'>")

           CreateTable(500)
           response.write("<td class='subtitulo td-r'>Entidad federativa destino</td>")
           response.write("<td class='td-l'>" & rsUpdateEntry5("entidad") &"</td></tr>")
           response.write("<tr><td class='subtitulo td-r'>Elija la ciudad:</td><td class='td-l'><select name='fciudad_destino'>")

           while not rsUpdateEntry6.EOF
               response.write("<option value='"& rsUpdateEntry6("id_municipio")&"*"& rsUpdateEntry6("ciudad") &"'>"& rsUpdateEntry6("ciudad")&" ("& rsUpdateEntry6("poblacion") &")</option>")
               rsUpdateEntry6.MoveNext
           wend
           response.write("</select></td></tr>") 
           response.write("<tr><td class='subtitulo td-r' width='35%'>Indique el almacen origen</td>")
           response.write("<td class='td-l' width='65%'><select name='falmacen'><option value='SJR'>SJR</option><option value='MXL'>MXL</option><option value='MTY'>MTY</option></select></td></tr>")

         response.write("<tr><td class='subtitulo td-r'>Elija unidad de carga</td>")

         strSQL="select a.*,b.Sucursalshort from Flotilla a left join cat_sucursal b on a.clave_sucursal=b.clave_sucursal where a.Activado=1 and a.tipo_activo!=6 order by a.Fabricante"
         rsUpdateEntry7.Open strSQL,adoCon

         response.write("<td class='td-l'><select name='funidad'>")
         while not rsUpdateEntry7.EOF
              response.write("<option value='"&rsUpdateEntry7("id_unidad")&"'>"&rsUpdateEntry7("id_unidad") &" | "& rsUpdateEntry7("fabricante") &" | maxima carga (kgs) "& rsUpdateEntry7("MaxCarga") &" ("& rsUpdateEntry7("Sucursalshort") &") </option>")
              rsUpdateEntry7.movenext
         wend         
         response.write("</select></td></tr>")      
         response.write("<tr><td colspan=2>&nbsp;</td></tr>")
         response.write("<tr><td colspan=2 class='td-c'><input type='submit' value='continuar'></td></tr></form></table>")
         
         rsUpdateEntry5.close
         rsUpdateEntry6.close      
         rsUpdateEntry7.close     

%>
