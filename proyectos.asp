<!--#include virtual="/config/conn.asp"-->
<!--#include virtual="/header.asp"-->

<P>&nbsp;</P><center>
<% 

Response.Expires = -1
Session("destino")="/modules/proyectos.asp"

if request("control")<>"" then  Session("control")=request("control")  end if
if request("item")<>"" then  Session("item")=request("item")  end if


if request("user")<>"" and request("pass")<>"" then	 
	   CheckLogin2	  
end if	

'if Session("session_ventas")="" then
	    'Formlogin
'else
			select case Session("control")
			case "1"      agregar  
			case "2"      editar         
			Case else 	  response.write("error")              
			end select
'end if	

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




<%

sub agregar
      Session("item")=""
      response.write("<form id='add' name='add' action='/modules/proyectos-up.asp' method='post'>")
      response.write("<table width='500px' border=0 cellpadding=2 cellspacing=1><tr><td colspan=3 class=subtitulo>Agregar un nuevo Proyecto de <B>Venta</B></td></tr>")
      
      response.write("<tr><td class=td-r width='40%'><B>Proyecto:</B></td><td class=td-l width='60%'><input class=input1 type='text' id='vproyecto' name='vproyecto' maxlength=50 style='width:180px'></td></tr> ")
           
      response.write("<tr><td class=td-r ><B>Estatus:</B></td><td class=td-l>abierto</td> </tr>")
      response.write("<tr><td class=td-r ><B>Fecha de Registro:</B></td><td class=td-l>"& day(date()) &"-" & meses(month(date())) &"-" &year(date())  &"</td> </tr>")
          
      response.write("<tr><td class=td-c colspan=2>&nbsp;</td></tr>")
      response.write("<tr><td class=td-c colspan=2><input type='submit' value='agregar'></form></td></tr></table>")
      
end sub





Sub editar

	  strSQL="SELECT * from Proyectos where ID="& Session("item")
		'response.write strSQL 
		rsUpdateEntry.Open strSQL, adoCon4
		
		response.write("<form id='add' name='add' action='/modules/proyectos-up.asp' method='post'>")
		response.write("<input type='hidden' id='item' name='item' value='"&  Session("item") &"'>")
		
      response.write("<table width='500px' border=0 cellpadding=2 cellspacing=1><tr><td colspan=3 class=subtitulo>Editar proyecto (item:  "&  Session("item") &") <font color=white>" &   rsUpdateEntry("proyecto")    &"</font></td></tr>")
      
      response.write("<tr><td class=td-r width='40%'><B>Proyecto:</B></td><td class=td-l width='60%'><B>"& rsUpdateEntry("proyecto") &"</B></td></tr> ")
           
      response.write("<tr><td class=td-r><B>Estatus:</B></td><td class=td-l><select name='festatus'>")
      if rsUpdateEntry("estatus")="abierto" then
         response.write("<option value='abierto' SELECTED>abierto</option> ")
      else
         response.write("<option value='abierto'>abierto</option> ")
      end if
      if rsUpdateEntry("estatus")="perdido" then
          response.write("<option value='perdido' SELECTED>perdido</option> ")
      else
          response.write("<option value='perdido'>perdido</option> ")
      end if
      if rsUpdateEntry("estatus")="ganado" then
           response.write("<option value='ganado' SELECTED>ganado</option> ")
      else
           response.write("<option value='ganado'>ganado</option> ")
      end if
      if rsUpdateEntry("estatus")="suspendido" then
           response.write("<option value='suspendido' SELECTED>suspendido</option> ")
      else
           response.write("<option value='suspendido'>suspendido</option> ")
      end if
      if rsUpdateEntry("estatus")="cancelado" then
             response.write("<option value='cancelado' SELECTED>cancelado</option> ")
      else
             response.write("<option value='cancelado'>cancelado</option> ")
      end if
 
      response.write("</select></td></tr>")
      response.write("<tr><td class=td-r ><B>Fecha de Registro:</B></td><td class=td-l>"& day(date()) &"-" & meses(month(date())) &"-" &year(date())  &"</td> </tr>")
          
      response.write("<tr><td class=td-c colspan=2>&nbsp;</td></tr>")
      response.write("<tr><td class=td-c colspan=2><input type='submit' value='modificar estatus'></form></td></tr></table>")
      

    rsUpdateEntry.close


end sub



%>

