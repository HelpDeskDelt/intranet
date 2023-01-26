<!--#include virtual="/config/conn.asp"-->


<%
     
           Response.Expires = -1
           
           titulo="Cambiando la entidad federativa de origen"
           DoTitulo(titulo)

           strSQL="select * from cat_entidades where id_entidad='"& request("variable")  &"' "
           rsUpdateEntry.Open strSQL, adoCon 

           strSQL="select a.*,b.entidad from viajes a left join cat_entidades b on a.entidad_origen=b.id_entidad where a.id_viaje='"& session("viaje")  &"' "
           rsUpdateEntry2.Open strSQL, adoCon 

           strSQL="select * from cat_municipios where id_entidad='"& request("variable")  &"' order by ciudad"
           'response.write strSQL
           rsUpdateEntry3.Open strSQL, adoCon 

           response.write("<form id='viajes' action='ShowContent.asp' method='POST'>")
           response.write("<input type='hidden' name='contenido' value='88'>")
           response.write("<input type='hidden' name='action' value='4'>")

           CreateTable(600)
           response.write("<tr><td class='subtitulo td-r FontMed' width='35%'>ID_viaje:</td><td class='td-l FontMed' width='65%'>" & session("viaje") &"</td></tr>") 
           response.write("<tr><td class='subtitulo td-r FontMed'>Entidad de origen:</td><td class='td-l FontMed'>"& rsUpdateEntry2("entidad") &"</td></tr>")        
           response.write("<tr><td class='subtitulo td-r FontMed'>Cambiando a:</td><td class='td-l FontMed'>")
           response.write("<select name='fentidad_origen'><option value='"& rsUpdateEntry("id_entidad") &"'>"& rsUpdateEntry("entidad") &"</option></select></td></tr>")

           response.write("<tr><td class='subtitulo td-r FontMed'>Elija la ciudad:</td><td class='td-l FontMed'><select name='fciudad_origen'>")

           while not rsUpdateEntry3.EOF
               response.write("<option value='"& rsUpdateEntry3("id_municipio")&"*"& rsUpdateEntry3("ciudad") &"'>"& rsUpdateEntry3("ciudad")&" ("& rsUpdateEntry3("poblacion") &")</option>")
               rsUpdateEntry3.MoveNext
           wend
           response.write("</select></td></tr>")
           response.write("<tr><td colspan=2 class='td-c FontMed'><input type='submit' value='actualizar'></form></td></tr>")
           response.write("<tr><td colspan=2 class='td-c FontMed'>&nbsp;</td></tr>")
           response.write("<tr><td colspan=2 class='td-c FontMed'><a href='ShowContent.asp?contenido=88&action=2&ID="& rsUpdateEntry2("id_viatico") &"'>cancelar</a></td></tr>")
      
      
           rsUpdateEntry.close
           rsUpdateEntry2.close
           rsUpdateEntry3.close
           CloseTable

 
%>
