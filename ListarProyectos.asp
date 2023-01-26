<!--#include virtual="/config/conn.asp"-->

<%
Response.Expires = -1
response.write("<center><H1><B>CONTROL DE PROYECTOS DE VENTA </B></H1></center>")

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

%>        