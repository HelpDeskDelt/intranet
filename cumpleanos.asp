<!--#include virtual="/config/conn.asp"-->       

<%       
       
       strSQL="select a.*,b.Departamento,c.Puesto,d.sucursal from Empleados a inner join cat_departamento b on a.Clave_depto=b.Clave_depto "
       strSQL= strSQL & "inner join cat_puesto c on a.clave_puesto=c.clave_puesto inner join cat_sucursal d on a.clave_sucursal=d.clave_sucursal  where a.FechaEgreso is null order by month(a.FechaNacimiento),day(a.FechaNacimiento)"     
       'response.write strSQL  
  	   rsUpdateEntry.Open strSQL, adoCon
  	   
       response.write("<table width='100%' border=0 cellpadding=3 cellspacing=2>")
       response.write("<tr>")
       response.write("<td class=subtitulo>ID</td>")
       response.write("<td class=subtitulo>COLABORADOR</td>")
       response.write("<td class=subtitulo>CUMPLEA&Ntilde;OS</td>")
       response.write("<td class=subtitulo>AREA</td>")
       response.write("<td class=subtitulo>SUCURSAL</td>")
       response.write("<td class=subtitulo>&nbsp;</td>")
       response.write("</tr>")
       
   
        while not rsUpdateEntry.EOF 
           response.write("<tr>")  %>
           <td><a href="javascript:sendReq('/modules/colaboradores.asp','control,item','2,<%=rsUpdateEntry("ID")%>','box')"> 
           	   <img src="/img/b_edit.png" border=0 with=11px height=11px alt="editar colaborador" title="editar colaborador"><%=rsUpdateEntry("ID") %> </a></td> 
           <%
           
           if  month(date())=month(rsUpdateEntry("FechaNacimiento")) then
           	    response.write("<td class=cumple>"& rsUpdateEntry("nombres") &" "  &rsUpdateEntry("PATERNO") &" " &rsUpdateEntry("MATERNO")   &"</td>")
                response.write("<td class=cumple><B>"& day(rsUpdateEntry("FechaNacimiento")) &"-" & meses(month(rsUpdateEntry("FechaNacimiento"))) &"</B></td>")
                response.write("<td class=cumple>"& rsUpdateEntry("DEPARTAMENTO") &"</td>")
                response.write("<td class=cumple>"& rsUpdateEntry("Sucursal") &"</td>")  
           else
           	    response.write("<td>"& rsUpdateEntry("nombres") &" "  &rsUpdateEntry("PATERNO") &" " &rsUpdateEntry("MATERNO")   &"</td>")
         	      response.write("<td><B>"& day(rsUpdateEntry("FechaNacimiento")) &"-" & meses(month(rsUpdateEntry("FechaNacimiento"))) &"</B></td>")
         	      response.write("<td>"& rsUpdateEntry("DEPARTAMENTO") &"</td>")
                response.write("<td>"& rsUpdateEntry("Sucursal") &"</td>")  
           end if                    
           
           response.write("<td>") %> <a href="javascript:sendReq('/modules/colaboradores.asp','control,item','3,<%=rsUpdateEntry("ID")%>','box')">
           <img src='/img/b_drop.png' border=0 with=11px height=11px alt="baja colaborador" title="baja colaborador"></td>    <%        
           response.write("</tr>")
           rsUpdateEntry.MoveNext
       wend
       rsUpdateEntry.close     %>
       <tr><td colspan=5 class=td-c>&nbsp;</td></tr> 
       <tr><td colspan=5 class=td-c><img src="/img/indicador.jpg" border=0 alt="agregar colaborador" title="agregar colaborador">
       	<a href="javascript:sendReq('/modules/colaboradores.asp','control','1','box')">  agregar un colaborador </a></td></tr>  
             <%
       response.write("</table>")    
       response.write("<P>&nbsp;</P>") 
       response.write("<P>&nbsp;</P>") 
       response.write("<P>&nbsp;</P>") 
       
       
%>       