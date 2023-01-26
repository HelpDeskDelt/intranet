<!--#include virtual="/config/conn.asp"-->

<%
     
           Response.Expires = -1
           
           titulo="Seleccione el destino de su actividad"
           DoTitulo(titulo)

           strSQL="select * from cat_entidades where id_entidad='"& request("variable")  &"' "
           'response.write strSQL
           rsUpdateEntry.Open strSQL, adoCon 
          
           strSQL="select * from cat_municipios where id_entidad='"& request("variable")  &"' order by ciudad"
           'response.write strSQL
           rsUpdateEntry2.Open strSQL, adoCon 

           response.write("<form id='envios' action='ShowContent.asp' method='POST'>")
           response.write("<input type='hidden' name='contenido' value='111'>")          

           CreateTable(600)
           response.write("<tr><td class='subtitulo td-r FontMed' width='35%'>Entidad destino:</td><td class='td-l FontMed' width='65%'><select name='fentidad'>")
           response.write("<option value='"& rsUpdateEntry("id_entidad")&"'>"& rsUpdateEntry("entidad")&"</option></select> </td></tr>") 

           response.write("<tr><td class='subtitulo td-r FontMed'>Elija la ciudad:</td><td class='td-l FontMed'><select name='fCiudadDestino'>")

           while not rsUpdateEntry2.EOF
               response.write("<option value='"& rsUpdateEntry2("id_municipio")&"*"& rsUpdateEntry2("ciudad") &"'>"& rsUpdateEntry2("ciudad")&" ("& rsUpdateEntry2("poblacion") &")</option>")
               rsUpdateEntry2.MoveNext
           wend
           response.write("</select></td></tr>")
           response.write("<tr><td colspan=2 class='td-c FontMed'><input type='submit' value='continuar'></form></td></tr>")
           response.write("<tr><td colspan=2 class='td-c FontMed'>&nbsp;</td></tr>")
           CloseTable

           rsUpdateEntry.close
           rsUpdateEntry2.close
         
 
%>
