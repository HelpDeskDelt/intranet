<!--#include virtual="/config/conn.asp"-->

<%
Response.Expires = -1
contenido=""
contenido=trim(request("data"))
ruta="/docs/"&request("data")

if Session("destino")="/modules/ShowDoc.asp"  and contenido="" then
      contenido=6
end if

if mid(contenido,1,3)="999" then           
      vposicion=InSTr(contenido,"-")
      video=mid(contenido,vposicion+1,1000)
      contenido="999"
end if 

select case contenido
case "1"      response.write("<center><H1><B>CUMPLEA&Ntilde;OS COLABORADORES EN "& year(date()) &"</B></H1></center>")
              cumple
case "2"      response.write("<center><H1><B>COTIZACIONES vs FACTURADO</B></H1></center>")
              cotizaciones      
case "3"      response.write("<center><H1><B>VIDEO: Capacitacion CRM en SAP</B></H1></center>")
              CRM    
case "4"      response.write("<center><H1><B>CONTROL DE PROYECTOS DE VENTA</B></H1></center>")
              proyectos     
case "5"      response.write("<center><H1><B>TERRITORIOS POR ASESOR DE VENTA</B></H1></center>")
              territorios   
case "6"      response.write("<center><H1><B>AVANCE DE VENTAS MENSUAL</B></H1></center>")
              AvanceVentas 
case "7"      response.write("<center><H1><B>REPRODUCIENDO UN VIDEO</B></H1></center>")
              Session("videos")="MindFullNess.mp4"
              PlayVideo
case "8"      response.write("<center><H1><B>Crear Ruta de repositorio Cliente nuevo</B></H1></center>")
              repositorio     




case "999"    response.write("<center><H1><B>REPRODUCIENDO UN TUTORIAL</B></H1></center>")
              ShowVideo
          
Case else     response.write("<center><H1>"& contenido &"</H1></center>")
              response.write("<iframe id=marco src="& ruta &"></iframe>")
end select



sub cumple
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
           <td><a href="javascript:sendReq('modules/colaboradores.asp','control,item','2,<%=rsUpdateEntry("ID")%>','box')"> 
           	   <img src="img/b_edit.png" border=0 with=11px height=11px alt="editar colaborador" title="editar colaborador"><%=rsUpdateEntry("ID") %> </a></td> 
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
           
           response.write("<td>") %> <a href="javascript:sendReq('modules/colaboradores.asp','control,item','3,<%=rsUpdateEntry("ID")%>','box')">
           <img src='img/b_drop.png' border=0 with=11px height=11px alt="baja colaborador" title="baja colaborador"></td>    <%        
           response.write("</tr>")
           rsUpdateEntry.MoveNext
       wend
       rsUpdateEntry.close     %>
       <tr><td colspan=5 class=td-c>&nbsp;</td></tr> 
       <tr><td colspan=5 class=td-c><img src="img/indicador.jpg" border=0 alt="agregar colaborador" title="agregar colaborador">
       	<a href="javascript:sendReq('modules/colaboradores.asp','control','1','box')">  agregar un colaborador </a></td></tr>  
             <%
       response.write("</table>")    
       response.write("<P>&nbsp;</P>") 
       response.write("<P>&nbsp;</P>") 
       response.write("<P>&nbsp;</P>") 
       response.write("<P>&nbsp;</P>") 
       response.write("<P>&nbsp;</P>") 
end sub




sub cotizaciones      

          response.write("<input type='text' name='fechaini' value='"& day(date()) &"/" & month(date()) &"/" & year(date())  &"'>")
          response.write( date() )
          %>
          
           <input type="button" value="fecha inicial" onclick="pureJSCalendar.open('dd.MM.yyyy', 20, 30, 1, '2018-5-5', '2019-8-20', 'txtTest', 20)" />
           <input type="text" id="txtTest" />
        

<%
end sub

sub  CRM     %>
 
   <video width="320" height="240" controls>
   <source src="movie.mp4" type="video/mp4">
   <source src="movie.ogg" type="video/ogg">
     Your browser does not support the video tag.
   </video> 


<%
end sub

sub proyectos
     Const adCmdText = &H0001
     Const adOpenStatic = 3
     nPage=0

     strSQL="select a.*,ventas2019=( select CONVERT(VARCHAR,CONVERT(MONEY,sum(subtotal)),1) from HistoricoVentas2019 where U_PROYECTO=a.Proyecto ),"  
     strSQL= strSQL &"ventas2018=( select CONVERT(VARCHAR,CONVERT(MONEY,sum(subtotal)),1) from HistoricoVentas2018 where U_PROYECTO=a.Proyecto ) "
     strSQL= strSQL &"from proyectos a order by a.FechaRegistro desc "
     'response.write strSQL  
     rsUpdateEntry.Open strSQL, adoCon4, adOpenStatic, adCmdText
     ' Set the page size of the recordset
     rsUpdateEntry.PageSize = 30    
      ' Get the number of pages
     nPageCount = rsUpdateEntry.PageCount
     if request("vPage")<>"" then
        nPage = int(trim(request("vPage")))
     else
        nPage=1
     end if      
     rsUpdateEntry.AbsolutePage=npage
       
       response.write("<table width='70%' border=0 cellpadding=3 cellspacing=2>")
       response.write("<tr><td colspan=5>Moverse a P&aacute;ginas: &nbsp;&nbsp;&nbsp;&nbsp;<B>")
       for i=1 to nPageCount step 1
           if i<>npage then   %>
               <a href="javascript:sendReq('modules/ListarProyectos.asp','vPage','<%=i%>','box')">   <%
               response.write( i &"</A>&nbsp;&nbsp;&nbsp;&nbsp;")                           
           end if
       next
       response.write("</B></td></tr>")
       response.write("<tr>")
       response.write("<td class=subtitulo>PROYECTO</td>")
       response.write("<td class=subtitulo>FECHA REGISTRO</td>")
       response.write("<td class=subtitulo>ESTATUS</td>") 
       response.write("<td class=subtitulo>VENTAS 2019</td>")   
       response.write("<td class=subtitulo>VENTAS 2018</td>")          
       response.write("</tr>")
       
       registros=1
        while not rsUpdateEntry.EOF AND registros<=30
           response.write("<tr>")  %>
           <td><a href="javascript:sendReq('modules/proyectos.asp','control,item','2,<%=rsUpdateEntry("ID")%>','box')"> 
               <img src="img/b_edit.png" border=0 with=11px height=11px alt="actualizar proyecto" title="actualizar proyecto"><%=rsUpdateEntry("proyecto") %> </a></td> 
           <%
                response.write("<td>"& day(rsUpdateEntry("FechaRegistro")) &"-" & meses(month(rsUpdateEntry("FechaRegistro")))  &"-" & year(rsUpdateEntry("FechaRegistro"))  &"</td>")
                response.write("<td>"& rsUpdateEntry("ESTATUS") &"</td>")
                response.write("<td style='text-align: right'>")   %>
                <a href="javascript:sendReq('modules/VentasProyecto.asp','proyecto,anio','<%=rsUpdateEntry("proyecto")%>,2019','box')"> <U>  <% 
                response.write( rsUpdateEntry("ventas2019") &"</U></a></td>")
              
                response.write("<td style='text-align: right'>")   %>
                <a href="javascript:sendReq('modules/VentasProyecto.asp','proyecto,anio','<%=rsUpdateEntry("proyecto")%>,2018','box')"> <U>  <% 
                response.write( rsUpdateEntry("ventas2018") &"</td>")
                registros=registros+1
                rsUpdateEntry.MoveNext
        wend

       response.write("<tr><td colspan=4>&nbsp;</td><td><B>P&aacute;gina actual: "& npage  &" </B></td></tr>")
       response.write("</table>")    

       %>
        <tr><td colspan=3 class=td-c style='font-size: 4px'>&nbsp;</td></tr> 
        <tr><td colspan=3 class=td-c><img src="img/indicador.jpg" border=0 alt="agregar un proyecto" title="agregar un proyecto">
        <a href="javascript:sendReq('modules/proyectos.asp','control','1','box')"> agregar un proyecto </a></td></tr>  
       <%
     
       response.write("<P>&nbsp;</P>") 
       response.write("<P>&nbsp;</P>") 
       response.write("<P>&nbsp;</P>") 
       response.write("<P>&nbsp;</P>") 
       response.write("<P>&nbsp;</P>") 

        rsUpdateEntry.close         
   
end sub

sub territorios
response.write("<P align='center'><img src='/images/territorios.png' border=0></P>")

end sub



sub AvanceVentas
     Session("destino")="/modules/ShowDoc.asp"

    if request("user")<>"" and request("pass")<>"" then   
         CheckLogin2    
    end if  

    if Session("session_ventas")="" then
          Formlogin
    else
          ShowAvanceVentas
    end if  
end sub
 




 Sub ShowAvanceVentas      
     strSQL="select * from HistoricoVentas" & year(date()) & " where month(DocDate)=month(getdate()) order by asesor,RazonSocial,DocDate"
     'response.write strSQL  
     rsUpdateEntry.Open strSQL, adoCon4
       
       response.write("<table width='100%' border=0 cellpadding=3 cellspacing=2>")
       response.write("<tr>")
       response.write("<td class=subtitulo>Asesor</td>")
       response.write("<td class=subtitulo>Factura</td>")
       response.write("<td class=subtitulo>Fecha</td>")  
       response.write("<td class=subtitulo>Cuenta</td>")      
       response.write("<td class=subtitulo>Subtotal</td>")  
       response.write("<td class=subtitulo>RazonSocial</td>")  
       response.write("<td class=subtitulo>Tipo</td>") 
       response.write("<td class=subtitulo>Pedido</td>")
       response.write("</tr>")
       
       vtotal=round(0,2)
       vasesor=""

        while not rsUpdateEntry.EOF 
                vasesor=rsUpdateEntry("asesor")
                vtotal=vtotal+round( rsUpdateEntry("subtotal"),2)
                response.write("<tr>")  %>
          
           <%
                response.write("<td>"& rsUpdateEntry("asesor") &"</td>")
                response.write("<td>"& rsUpdateEntry("FolioDocumento") &"</td>")
                response.write("<td>"& day(rsUpdateEntry("DocDate")) &"-" & meses(month(rsUpdateEntry("DocDate")))  &"-" & year(rsUpdateEntry("DocDate"))  &"</td>")
                response.write("<td>"& rsUpdateEntry("Cliente") &"</td>")
                response.write("<td>"& rsUpdateEntry("subtotal") &"</td>")
                response.write("<td>"& rsUpdateEntry("RazonSocial") &"</td>")
                response.write("<td>"& rsUpdateEntry("Tipo") &"</td>")
                response.write("<td>"& rsUpdateEntry("Pedido") &"</td>")
                
                rsUpdateEntry.MoveNext
                if not rsUpdateEntry.EOF then
                      if vasesor<> rsUpdateEntry("asesor") then
                          response.write("<tr><td colspan=8 class=td-c style='font-size: 4px'><HR></td></tr>")
                          response.write("<tr><td colspan=4 style='text-align: right'><B> "& vasesor &"</B></td><td> <B>"& vtotal &"</B></td>")
                          vtotal=round(0,2)
                          response.write("<tr><td colspan=8 class=td-c style='font-size: 4px'><HR></td></tr>")
                      end if
                else
                          response.write("<tr><td colspan=8 class=td-c style='font-size: 4px'><HR></td></tr>")
                          response.write("<tr><td colspan=4 style='text-align: right'><B> "& vasesor &"</B></td><td> <B>"& vtotal &"</B></td>")
                          vtotal=round(0,2)
                          response.write("<tr><td colspan=8 class=td-c style='font-size: 4px'><HR></td></tr>")

                end if
        wend
       rsUpdateEntry.close  
       response.write("</table><P>&nbsp;</P>") 
       response.write("<P>&nbsp;</P>") 
       response.write("<P>&nbsp;</P>") 
       response.write("<P>&nbsp;</P>") 

end sub



sub ShowVideo   
   video="/docs/"&video
   'response.write video
%>

<video width="900" height="600" controls>
  <source src="<%=video%>" type="video/mp4">
Your browser does not support the video tag.
</video>

<%
end sub



sub PlayVideo   
     Session("videos")="/docs/"& Session("videos")
%>

<video width="900" height="600" controls>
  <source src="<%=Session("videos")%>" type="video/mp4">
Your browser does not support the video tag.
</video>

<%
end sub
 



sub repositorio

 

end sub
%>

