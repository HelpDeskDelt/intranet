<!--#include virtual="/config/conn.asp"-->
<!--#include virtual="/header.asp"-->

<BR>
<%

if request("ID")<>"" and request("action")<>"" then
      select case request("action")
      case "1" ShowFormEdit
      case "2" EditUP
      case "3" ShowFormDelete
      case "4" DeleteRecord
    end select
else
      ShowForm
end if












Sub ShowForm

if request("ffecha")<>"" and request("fRS")<>""   then

    if request("fremision")<>"" then
          strSQL="select  * from envios where RazonSocial='"& request("fRS") &"' and DocNum=" & request("fremision") &" and tipo='remision'" 
          rsUpdateEntry2.Open strSQL, adoCon4

          strSQL="select  top 1 a.cardcode,a.cardname,b.WhsCode from ODLN a inner join DLN1 b on a.DocEntry=b.DocEntry where a.docnum="& request("fremision") 
          vfecha=date()
          vfecha=request("ffecha")

          if request("fRS")="DMX" then
              rsUpdateEntry.Open strSQL, adoCon2
          else 
              rsUpdateEntry.Open strSQL, adoCon3
          end if  

          strSQL="set dateformat ymd;insert into envios (ID,DocDate,ID_empleado,LogDate,RazonSocial,DocNum,Tipo,WhsCode,id_Portador,Guia,Rastreo,subTotal,Cardcode,Cardname,Comentarios,ocurre,reenvio,CargoCliente) values ( (select max(ID) from envios)+1,'"
          strSQL= strSQL & year(vfecha) &"-" & month(vfecha) &"-" & day(vfecha) &"',"
          strSQL= strSQL & request("fempleado") &",getdate(),'" & request("fRS") &"'," & request("fremision") &",'remision','" & rsUpdateEntry("WhsCode") &"',"& request("fCarrier") &",'"
          strSQL= strSQL & request("fGuia") &"','" &  request("frastreo") &"'," &  request("fTotal")  &",'" & rsUpdateEntry("cardcode") &"','"& rsUpdateEntry("cardname") &"','" & UCASE(request("fNotas")) &"','"
          strSQL= strSQL & request("focurre") &"',"
          
          if request("freenvio")="on" then
             strSQL= strSQL & "1,"
          else
             strSQL= strSQL & "0,"
          end if
           if request("fCargoCliente")="on" then
             strSQL= strSQL & "1)"
          else
             strSQL= strSQL & "0)"
          end if

          'response.write strSQL

          rsUpdateEntry.close
          if rsUpdateEntry2.EOF then
               'response.write strSQL
               rsUpdateEntry.Open strSQL, adoCon4  'se crea el envio'
               response.write("<P class='alert td-c'>La informacion ha sido registrada</P>")
          else
              if request("freenvio")="on" then
                  rsUpdateEntry.Open strSQL, adoCon4  'se crea el envio'
                  response.write("<P class='alert td-c'>La informacion ha sido registrada</P>")
              else
                 response.write("<P class='alert td-c'>NO SE REGISTR&Oacute; SU ENV&Iacute;O, YA EST&Aacute; REGISTRADO UN ENV&Iacute;O PARA ESTA REMISI&Oacute;N </P>")
              end if
          end if
    else
          
             vfecha=date()
             vfecha=request("ffecha")

              strSQL="set dateformat ymd;insert into envios (ID,DocDate,ID_empleado,LogDate,RazonSocial,id_Portador,Guia,Rastreo,subTotal,Comentarios,reenvio,CargoCliente) values ( (select max(ID) from envios)+1,'"
              strSQL= strSQL & year(vfecha) &"-" & month(vfecha) &"-" & day(vfecha) &"',"
              strSQL= strSQL & request("fempleado") &",getdate(),'" & request("fRS") &"',"& request("fCarrier") &",'"
              strSQL= strSQL & request("fGuia") &"','"  & request("frastreo") &"'," &  request("fTotal")  &",'" & UCASE(request("fNotas")) &"',"
              

          if request("freenvio")="on" then
             strSQL= strSQL & "1,"
          else
             strSQL= strSQL & "0,"
          end if
           if request("fCargoCliente")="on" then
             strSQL= strSQL & "1)"
          else
             strSQL= strSQL & "0)"
          end if

              'response.write strSQL
              rsUpdateEntry.Open strSQL, adoCon4  'se crea el envio'
              response.write("<P class='alert td-c'>La informacion ha sido registrada</P>")
          
          

    end if
else
   if request("ffecha")="" and request("fRS")<>"" and ( request("fremision")<>"" or   request("fcarrier")="6" )  then
       response.write("<P class='alert td-c'>NO SE REGISTR&Oacute; SU ENV&Iacute;O, INFORMACI&Oacute;N INCOMPLETA </P>")
    end if
end if



strSQL="select ID,Nombres+' '+paterno+' '+Materno as nombre from Empleados where FechaEgreso is null order by nombre"
 'response.write strSQL  
rsUpdateEntry.Open strSQL, adoCon

    
      doAlert
      titulo="Registro de uso de paqueter&iacute;a &oacute; mensajer&iacute;a"
      Dotitulo(titulo)


      %>
    
     <table border=0 width="600px" cellpadding="2" cellpadding="2" align="center">
     <form id="envios" method="POST" action="envios.asp"> 
     <tr><td class='fontbig subtitulo td-r' width="40%">Raz&oacute;n Social: </td><td class='fontbig td-l' width="60%"><select id="fRS" name="fRS"> <option value="DMX">DMX</option> <option value="DFW">DFW</option> </select></td></tr>
     <tr><td class='fontbig subtitulo td-r'># Remisi&oacute;n: </td><td class='fontbig td-l'><input type="number" id="fremision" name="fremision"></td></tr>
     <tr><td class='fontbig subtitulo td-r'>Colaborador que captura: </td><td class='fontbig td-l'><select name="fempleado" id="fempleado">
     <%
        while not rsUpdateEntry.EOF
            response.write("<option value='"& rsUpdateEntry("ID") &"'>"& rsUpdateEntry("nombre") &"</option>")
            rsUpdateEntry.MoveNext
        wend
        rsUpdateEntry.close
     %>
     </select></td></tr>
     
     <tr><td class='fontbig subtitulo td-r'># de gu&iacute;a: </td><td class='fontbig td-l'><input type="text" id="fguia" name="fguia"></td></tr>
     <tr><td class='fontbig subtitulo td-r'># de rastreo: </td><td class='fontbig td-l'><input type="text" id="frastreo" name="frastreo"></td></tr>

     <tr><td class='fontbig subtitulo td-r'>Subtotal sin impuestos: </td><td class='fontbig td-l'><input type="text" id="ftotal" name="ftotal" placeholder="no puede ir vac&iacute;o"> </td></tr>
   
     
     <tr><td class='fontbig subtitulo td-r'>Fecha del env&iacute;o: </td><td class='fontbig td-l'><input name="ffecha" type="date"/> </td></tr>
     <tr><td class='fontbig subtitulo td-r'>Portador: </td><td class='fontbig td-l'><select name="fcarrier" id="fcarrier" onchange="ValidaRemision()">
         <option value="0">seleccione</option> <%

          strSQL="select * from cat_portadores order by ID_portador"
          rsUpdateEntry3.Open strSQL, adoCon4 

          while not rsUpdateEntry3.EOF
              response.write("<option value="& rsUpdateEntry3("ID_portador") &">"& rsUpdateEntry3("Portador") &"</option>")
              rsUpdateEntry3.MoveNext
          wend
          rsUpdateEntry3.close  %>

          </select></td></tr>

     <tr><td class='fontbig subtitulo td-r'>Cargo a Cliente: </td><td class='fontbig td-l'><input type="checkbox" name="fCargoCliente" id="fCargoCliente"> </td></tr>
     
     <tr><td class='fontbig subtitulo td-r'>Comentarios: </td><td class='fontbig td-l'><input type="text" id="fnotas" name="fnotas" class='anchox2'></td></tr>
     <tr><td class='fontbig subtitulo td-r'>En OCURRE a otro Almac&eacute;n:</td><td class='fontbig td-l'><select name='focurre' class="anchox2">
              <option value="" SELECTED>si aplica seleccione</option>
              <option value="001-Mxl">001-Mxl</option>
              <option value="002-SJR">002-SJR</option>
              <option value="007-Esc">007-Esc</option>
             </select>  </td></tr>


     <tr><td class='fontbig subtitulo td-r'>Re-env&iacute;o: </td><td class='fontbig td-l'><input type="checkbox" name="freenvio" id="freenvio"> (para un segundo env&iacute;o)</td></tr>

     <tr><td colspan=2 class='td-c'>  <div id="box">&nbsp;</div> </td></tr>
     </form>
     </table>
   </div>
</div>

    <P>&nbsp;</P><P>&nbsp;</P>
    <H3>&Uacute;ltimos registros</H3>
    <form id="buscar" method="POST" action="envios.asp">
    <H2 class="td-c fontmed">Buscar una gu&iacute;a o un rastero: <input type="text" name="vstring" value="<%=request("vstring")%>"> &nbsp;&nbsp;&nbsp;<input type="submit" value="buscar"></form> </H2>
    <H2 class="td-c fontmed">&nbsp;</H2>

 <table width="1400px" cellpadding="3" cellpadding="2" align="center" border=0>
   


<%
  Const adCmdText = &H0001
  Const adOpenStatic = 3
  nPage=0
  registros=1

  if request("vstring")<>"" and len(request("vstring"))>2 then
      strSQL="select  a.*,b.Portador,dbo.fn_GetMonthName(a.DocDate,'spanish') AS [DocDate2] from envios a inner join cat_portadores b on a.id_Portador=b.id_Portador where ( guia like '%"&request("vstring")&"%' OR rastreo like '%"&request("vstring")&"%' )  order by a.ID desc" 
  else
      
      strSQL="select  a.*,b.Portador,dbo.fn_GetMonthName(a.DocDate,'spanish') AS [DocDate2] from envios a inner join cat_portadores b on a.id_Portador=b.id_Portador where DATEDIFF(DAY,a.DocDate,getdate())<40 order by a.ID desc" 
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
  response.write("<tr><td colspan=2 class='td-r'><B>PAGINAS:</B></td><td colspan=13>")
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
      <td class="td-c azul fontbig">Almac&eacute;n <br>origen</td>
      <td class="td-c azul fontbig">Portador</td>
      <td class="td-c azul fontbig">Gu&iacute;a</td>
      <td class="td-c azul fontbig">Rastreo</td>
      <td class="td-c azul fontbig" width='80px'>Cargo <br> cliente</td>
      <td class="td-c azul fontbig">Socio de Negocio</td>
      <td class="td-c azul fontbig">Comentarios</td>
      <td class="td-c azul fontbig">Conso-<br>lidado</td>
      <td class="td-c azul fontbig">Pedido</td>
      <td class="td-c azul fontbig">&nbsp;&nbsp;&nbsp;&nbsp;</td>
      <td class="td-c azul fontbig">Almac&eacute;n <br>ocurre</td>

    </tr>
    <%        

  vfecha1=date()-5
  vfecha2=date()


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
         response.write("<td class='td-l fontmed'> "& rsUpdateEntry("Pedido") &"</td>")
         response.write("<td class='td-l fontmed'>&nbsp;</td>")
     else
        response.write("<td class='td-c fontmed'>no</td>")
        response.write("<td class='td-l fontmed'> "& rsUpdateEntry("Pedido") &"</td>")
        vfecha2=cDate( day(rsUpdateEntry("DocDate"))&"/" & month(rsUpdateEntry("DocDate"))&"/" & year(rsUpdateEntry("DocDate")) )
        if vfecha2>vfecha1 then 
             response.write("<td class='td-l fontmed'><a href='envios.asp?action=3&ID="& rsUpdateEntry("ID") &"'><img src='/img/b_drop.png' with='11px' alt='borrar' title='borrar' height='11px' border=0></a></td>")
        else
             response.write("<td class='td-l fontmed'>&nbsp;</td>")
        end if

        response.write("<td class='td-l fontmed'> "& rsUpdateEntry("ocurre") &"</td>")

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

