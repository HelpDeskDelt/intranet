<!--#include virtual="/config/conn.asp"-->
<!--#include virtual="/header.asp"-->

<BR>

<div id="outer">  
   <div id="BoxRounded">   
    
    <H1>&Uacute;ltimos registros</H1>
    <form id="buscar" method="POST" action="envios.asp">
    <H2 class="td-c fontmed">Buscar una gu&iacute;a o un rastero: <input type="text" name="vstring" value="<%=request("vstring")%>"> &nbsp;&nbsp;&nbsp;<input type="submit" value="buscar"></form> </H2>
    <H2 class="td-c fontmed">&nbsp;</H2>

 <table width="1300px" cellpadding="3" cellpadding="2" align="center" border=0>
   


<%
  Const adCmdText = &H0001
  Const adOpenStatic = 3
  nPage=0
  registros=1

  if request("vstring")<>"" and len(request("vstring"))>2 then
      strSQL="select  a.*,b.Portador,dbo.fn_GetMonthName(a.DocDate,'spanish') AS [DocDate2] from envios a inner join cat_portadores b on a.id_Portador=b.id_Portador where ( guia like '%"&request("vstring")&"%' OR rastreo like '%"&request("vstring")&"%' )  order by a.ID desc" 
  else
      
      strSQL="select  a.*,b.Portador,dbo.fn_GetMonthName(a.DocDate,'spanish') AS [DocDate2] from envios a inner join cat_portadores b on a.id_Portador=b.id_Portador where DATEDIFF(DAY,a.DocDate,getdate())<90 order by a.ID desc" 
  end if
  rsUpdateEntry.Open strSQL, adoCon4, adOpenStatic, adCmdText

  rsUpdateEntry.PageSize = 10 
  nPageCount = rsUpdateEntry.PageCount

  if request("vPage")<>"" then
       nPage = int(trim(request("vPage")))
  else
       nPage=1
  end if      
  rsUpdateEntry.AbsolutePage=npage
  response.write("<tr><td colspan=5 class='td-r'><B>PAGINAS:</B></td><td colspan=10>")
        for i=1 to nPageCount step 1
            if i<>npage then  
                response.write("<a href='/modules/envios.asp?vPage="&i &"'>") 
                response.write( i &"</A>&nbsp;&nbsp;&nbsp;&nbsp;") 
            end if
        next
        response.write("</td></tr>")    %>

 <tr>
      <td class="td-c azul fontbig">ID</td>
      <td class="td-c azul fontbig">Colaborador</td>
      <td class="td-c azul fontbig" width='80px'>Fecha</td>
      <td class="td-c azul fontbig">RS</td>
      <td class="td-c azul fontbig">Dcto</td>
      <td class="td-c azul fontbig">Tipo</td>
      <td class="td-c azul fontbig">Almac&eacute;n</td>
      <td class="td-c azul fontbig">Portador</td>
      <td class="td-c azul fontbig">Gu&iacute;a</td>
      <td class="td-c azul fontbig">Rastreo</td>
      <td class="td-c azul fontbig" width='80px'>Cargo <br> cliente</td>
      <td class="td-c azul fontbig">Socio de Negocio</td>
      <td class="td-c azul fontbig">Comentarios</td>
      <td class="td-c azul fontbig">Conso-<br>lidado</td>
      <td class="td-c azul fontbig">&nbsp;&nbsp;&nbsp;&nbsp;</td>
    </tr>
    <%        


  while not rsUpdateEntry.EOF AND registros<=10
     response.write("<tr>")

     response.write("<td class='td-r fontmed'><a href='envios.asp?action=1&ID="& rsUpdateEntry("ID") &"'><img src='/img/b_edit.png' with='11px' alt='editar' title='editar' height='11px' border='0'>"& rsUpdateEntry("ID") &"</a></td>")   'mike'

     strSQL="select  Nombres+' '+paterno as Colaborador from Empleados where id=" &   rsUpdateEntry("ID_empleado")
     rsUpdateEntry4.Open strSQL, adoCon

     response.write("<td class='td-l fontmed'>"& rsUpdateEntry4("Colaborador") &"</td>")
     rsUpdateEntry4.close
     response.write("<td class='td-r fontmed'>"& rsUpdateEntry("DocDate2") &"</td>")
     response.write("<td class='td-c fontmed'>"& rsUpdateEntry("RazonSocial") &"</td>")
     response.write("<td class='td-r fontmed'>"& rsUpdateEntry("DocNum") &"</td>")
     response.write("<td class='td-c fontmed'>"& rsUpdateEntry("Tipo") &"</td>")
     response.write("<td class='td-c fontmed'>"& rsUpdateEntry("WhsCode") &"</td>")
     response.write("<td class='td-l fontmed'>"& rsUpdateEntry("Portador") &"</td>")
     response.write("<td class='td-l fontmed'>"& rsUpdateEntry("Guia") &"</td>")
     response.write("<td class='td-l fontmed'> "& rsUpdateEntry("Rastreo") &"</td>")


     if rsUpdateEntry("CargoCliente") =1 then
         response.write("<td class='td-c fontmed'> si</td>")
     else
        response.write("<td class='td-c fontmed'>no</td>")
     end if

     response.write("<td class='td-l fontmed'>"& rsUpdateEntry("Cardname") &"</td>")
     response.write("<td class='td-l fontmed'>"& rsUpdateEntry("Comentarios") &"</td>")

     if rsUpdateEntry("Consolidado") =1 then
         response.write("<td class='td-c fontmed'> si</td>")
         response.write("<td class='td-l fontmed'>&nbsp;</td>")
     else
        response.write("<td class='td-c fontmed'>no</td>")
        response.write("<td class='td-l fontmed'><a href='envios.asp?action=3&ID="& rsUpdateEntry("ID") &"'><img src='/img/b_drop.png' with='11px' alt='borrar' title='borrar' height='11px' border=0></a></td>")
     end if


      
     response.write("</tr>")
     rsUpdateEntry.MoveNext
     registros=registros+1

  wend

    response.write("</table>") 

end sub

%>


</body>  


  <!-- menu script itself. you should not modify this file -->
  <script language="JavaScript" src="/menuj/menu.js"></script>
  <script language="JavaScript" src="/menuj/menu_items.js"></script>
  <script language="JavaScript" src="/menuj/menu_tpl.js"></script>
</p>
<br>
<script language="JavaScript">
  <!--//

  // some table cell or other HTML element. Always put it before </body>
  // each menu gets two parameters (see demo files)
  // 1. items structure
  // 2. geometry structure

  new menu (MENU_ITEMS, MENU_POS);
  
  //-->
</script>

