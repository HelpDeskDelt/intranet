<!--#include virtual="/config/conn.asp"-->

<%
Response.Expires = -1

Session("destino")="/modules/colaboradores.asp"
if request("control")<>"" then  Session("control")=request("control")  end if
if request("item")<>"" then  Session("item")=request("item")  end if

'response.write("control=" & Session("control")    )
'response.write("item=" & Session("item")    )

if request("user")<>"" and request("pass")<>"" then	 
	   CheckLogin	  
     'response.write("control=" & Session("control")    )
     'response.write("item=" & Session("item")    )
end if	

if Session("session_id")="" then
	  Formlogin
else
			select case Session("control")
			case "1"      agregar  
			case "2"      editar         

			case "3"       Baja

			Case else 	  response.write("error")              
			end select

end if	


sub Baja   
      strSQL="SELECT * from Empleados where ID="& Session("item") 
      'response.write strSQL 
      rsUpdateEntry.Open strSQL, adoCon
      
      response.write("<form id='add' name='add' action='/modules/colabor-up.asp' method='post'>")      
      response.write("<input type='hidden' id='control' name='control' value=3>")
      response.write("<input type='hidden' id='item' name='item' value="& Session("item")  &">")

      response.write("<table width='650px' border=0 cellpadding=2 cellspacing=1><tr><td colspan=5 class=subtitulo>Dar de baja colaborador: <B>"&  Session("item") &"</td></tr>")

      response.write("<tr><td class=td-r width='50%'><B>Nombres:</B></td><td class=td-l width='50%'>"& rsUpdateEntry("nombres")& "</td></tr>")      
      response.write("<tr><td class=td-r ><B>Paterno:</B></td><td class=td-l>"& rsUpdateEntry("paterno")& "</td><td class=td-c>&nbsp;</td> ")
      response.write("<tr><td class=td-r ><B>Paterno:</B></td><td class=td-l>"& rsUpdateEntry("materno")& "</td><td class=td-c>&nbsp;</td> ")

      response.write("<tr><td class=td-r><B>Fecha de Egreso:</B></td><td class=td-l>")  %>
      <input name="vegreso" type="date"/>    <%      
      response.write("</td></tr>")
      response.write("<tr><td class=td-r ><B>Metodo de baja:</B></td><td class=td-l><select name='vFormaBaja'><option value='finiquito'>Finiquito</option><option value='liquidacion' SELECTED>Liquidacion</option><td class=td-c>&nbsp;</td> ")
      
      response.write("<tr><td class=td-c colspan=5><input type='submit' value='dar de baja'></form></td></tr></table>")
      rsUpdateEntry.close

end sub



sub agregar
      Session("item")=""
      response.write("<form id='add' name='add' action='/modules/colabor-up.asp' method='post'>")
      response.write("<table width='800px' border=0 cellpadding=2 cellspacing=1><tr><td colspan=5 class=subtitulo>Agregar un nuevo colaborador</td></tr>")
      
      response.write("<tr><td class=td-r width='18%'><B>ID:</B></td><td class=td-l width='32%'><input class=input1 type='text' id='vid' name='vid' maxlength=3></td><td class=td-c>&nbsp;</td> ")
      response.write("<td class=td-r width='18%'><B>Nombres:</B></td><td class=td-l width='32%'><input class=input1 type='text' id='vnombres' name='vnombres'></td></tr>")
      
      response.write("<tr><td class=td-r ><B>Paterno:</B></td><td class=td-l><input class=input1 type='text' id='vpaterno' name='vpaterno'></td><td class=td-c>&nbsp;</td> ")
      response.write("<td class=td-r><B>Materno:</B></td><td class=td-l><input class=input1 type='text' id='vmaterno' name='vmaterno'></td></tr>")
      
      response.write("<tr><td class=td-r ><B>Fecha Nacimiento:</B></td><td class=td-l>")  %>
      <input name="vnacimiento" type="date"/>    <%
      
      response.write("</td><td class=td-c>&nbsp;</td>")  
      response.write("<td class=td-r><B>Fecha Ingreso:</B></td><td class=td-l>")  %>
      <input name="vingreso" type="date"/>    <%
      
      response.write("</td></tr>")
      
      response.write("<tr><td class=td-r ><B>CURP:</B></td><td class=td-l><input class=input1 type='text' id='vcurp' name='vcurp' maxlength=18></td><td class=td-c>&nbsp;</td> ")
      response.write("<td class=td-r><B>RFC:</B></td><td class=td-l><input class=input1 type='text' id='vrfc' name='vrfc' maxlength=13></td></tr>")
      
      response.write("<tr><td class=td-r ><B>NSS:</B></td><td class=td-l><input class=input1 type='text' id='vnss' name='vnss' maxlength=11></td><td class=td-c>&nbsp;</td> ") 
      response.write("<td class=td-r><B>Celular:</B></td><td class=td-l><input class=input1 type='text' id='vcelular' name='vcelular' maxlength=10>10 d&iacute;gitos</td></tr>")
      
      response.write("<tr><td class=td-r ><B>SEXO:</B></td><td class=td-l><select name=vsexo><option value='F'>F</option><option value='M'>M</option></select></td><td class=td-c>&nbsp;</td> ")
      response.write("<td class=td-r><B>email:</B></td><td class=td-l><input class=input1 type='text' id='vemail' name='vemail' style='width:200px' maxlength=30></td></tr>")
      
      response.write("<tr><td class=td-r ><B>Puesto:</B></td><td class=td-l><select name=vpuesto>")
      strSQL=  "select * from cat_puesto order by clave_puesto"     
      'response.write strSQL  
  	  rsUpdateEntry.Open strSQL, adoCon
      
      while not rsUpdateEntry.EOF
           response.write("<option value='"& rsUpdateEntry("clave_puesto")  &"'>"& rsUpdateEntry("PUESTO") &"</option>")
           rsUpdateEntry.movenext
      wend
       rsUpdateEntry.close
      response.write("</select></td><td class=td-c>&nbsp;</td>")
  
      response.write("<td class=td-r><B>Departamento:</B></td><td class=td-l><select name=vdepto>")
      strSQL=  "select * from cat_departamento order by clave_depto"     
      'response.write strSQL  
  	  rsUpdateEntry.Open strSQL, adoCon
      
      while not rsUpdateEntry.EOF
           response.write("<option value='"& rsUpdateEntry("clave_depto")  &"'>"& rsUpdateEntry("Departamento") &"</option>")
           rsUpdateEntry.movenext
      wend
       rsUpdateEntry.close
      response.write("</select> </td></tr>")
     
      response.write("<tr><td class=td-r ><B>Sucursal:</B></td><td class=td-l><select name=vsucursal>")
      strSQL=  "select * from cat_sucursal order by clave_sucursal"     
      'response.write strSQL  
  	  rsUpdateEntry.Open strSQL, adoCon
      
      while not rsUpdateEntry.EOF
           response.write("<option value='"& rsUpdateEntry("clave_sucursal")  &"'>"& rsUpdateEntry("sucursal") &"</option>")
           rsUpdateEntry.movenext
      wend
       rsUpdateEntry.close
      response.write("</select> </td><td class=td-c>&nbsp;</td> <td class=td-c colspan=2>&nbsp;</td> </tr>")
      response.write("<tr><td class=td-c colspan=5>&nbsp;</td></tr>")
      response.write("<tr><td class=td-c colspan=5><input type='submit' value='agregar'></form></td></tr></table>")
      
end sub



Sub editar

	  strSQL="SELECT * from Empleados where ID="& Session("item") 
		'response.write strSQL 
		rsUpdateEntry.Open strSQL, adoCon
		
		response.write("<form id='add' name='add' action='/modules/colabor-up.asp' method='post'>")
		response.write("<input type='hidden' id='item' name='item' value='"& TRIM(request("item")) &"'>")
		
      response.write("<table width='800px' border=0 cellpadding=2 cellspacing=1><tr><td colspan=5 class=subtitulo>Editar colaborador no. "& TRIM(request("item")) &"</td></tr>")
      
      response.write("<tr><td class=td-r width='18%'><B>ID:</B></td><td class=td-l width='32%'><B>"& TRIM(request("item")) &"</B></td><td class=td-c>&nbsp;</td> ")
      response.write("<td class=td-r width='18%'><B>Nombres:</B></td><td class=td-l width='32%'><input class=input1 type='text' id='vnombres' name='vnombres' value='"& rsUpdateEntry("nombres")&"'></td></tr>")
      
      response.write("<tr><td class=td-r ><B>Paterno:</B></td><td class=td-l><input class=input1 type='text' id='vpaterno' name='vpaterno' value='"& rsUpdateEntry("paterno")&"'></td><td class=td-c>&nbsp;</td> ")
      response.write("<td class=td-r><B>Materno:</B></td><td class=td-l><input class=input1 type='text' id='vmaterno' name='vmaterno' value='"& rsUpdateEntry("materno")&"'></td></tr>")
      
    
      
      response.write("<tr><td class=td-r ><B>CURP:</B></td><td class=td-l><input class=input1 type='text' id='vcurp' name='vcurp' maxlength=18 value='"& rsUpdateEntry("CURP")&"'></td><td class=td-c>&nbsp;</td> ")
      response.write("<td class=td-r><B>RFC:</B></td><td class=td-l><input class=input1 type='text' id='vrfc' name='vrfc' maxlength=13 value='"& rsUpdateEntry("RFC")&"'></td></tr>")
      
      response.write("<tr><td class=td-r ><B>NSS:</B></td><td class=td-l><input class=input1 type='text' id='vnss' name='vnss' maxlength=11 value='"& rsUpdateEntry("NSS")&"'></td><td class=td-c>&nbsp;</td> ") 
      response.write("<td class=td-r><B>Celular:</B></td><td class=td-l><input class=input1 type='text' id='vcelular' name='vcelular' maxlength=10 value='"& rsUpdateEntry("CELULAR_PERSONAL")&"'>10 d&iacute;gitos</td></tr>")
           
      response.write("<tr><td class=td-r ><B>SEXO:</B></td><td class=td-l><select name=vsexo>")
      if rsUpdateEntry("sexo")="M" then
      	    response.write("<option value='F'>F</option><option value='M' SELECTED>M</option></select></td><td class=td-c>&nbsp;</td> ")
      else
      	    response.write("<option value='F' SELECTED>F</option><option value='M'>M</option></select></td><td class=td-c>&nbsp;</td> ")
      end if

      response.write("<td class=td-r><B>email:</B></td><td class=td-l><input class=input1 type='text' id='vemail' name='vemail' style='width:200px' maxlength=30 value='"& rsUpdateEntry("email")&"'></td></tr>")
      
      response.write("<tr><td class=td-r ><B>Puesto:</B></td><td class=td-l><select name=vpuesto>")
      strSQL=  "select * from cat_puesto order by clave_puesto"     
      'response.write strSQL  
  	  rsUpdateEntry2.Open strSQL, adoCon
      
      while not rsUpdateEntry2.EOF
           if rsUpdateEntry("clave_puesto")=rsUpdateEntry2("clave_puesto") then
                response.write("<option value='"& rsUpdateEntry2("clave_puesto")  &"' SELECTED>"& rsUpdateEntry2("PUESTO") &"</option>")
           else
                response.write("<option value='"& rsUpdateEntry2("clave_puesto")  &"'>"& rsUpdateEntry2("PUESTO") &"</option>")
           end if
           rsUpdateEntry2.movenext
      wend
      rsUpdateEntry2.close
      response.write("</select></td><td class=td-c>&nbsp;</td>")
  
      response.write("<td class=td-r><B>Departamento:</B></td><td class=td-l><select name=vdepto>")
      strSQL=  "select * from cat_departamento order by clave_depto"     
      'response.write strSQL  
  	  rsUpdateEntry2.Open strSQL, adoCon
      
      while not rsUpdateEntry2.EOF
           if rsUpdateEntry("clave_depto")=rsUpdateEntry2("clave_depto") then
                  response.write("<option value='"& rsUpdateEntry2("clave_depto")  &"' SELECTED>"& rsUpdateEntry2("Departamento") &"</option>")
           else
           	      response.write("<option value='"& rsUpdateEntry2("clave_depto")  &"'>"& rsUpdateEntry2("Departamento") &"</option>")
           end if
           rsUpdateEntry2.movenext
      wend
      rsUpdateEntry2.close
      response.write("</select> </td></tr>")
     
      response.write("<tr><td class=td-r ><B>Sucursal:</B></td><td class=td-l><select name=vsucursal>")
      strSQL=  "select * from cat_sucursal order by clave_sucursal"     
      'response.write strSQL  
  	  rsUpdateEntry2.Open strSQL, adoCon
      
      while not rsUpdateEntry2.EOF
           if rsUpdateEntry("clave_sucursal")=rsUpdateEntry2("clave_sucursal") then
                  response.write("<option value='"& rsUpdateEntry2("clave_sucursal")  &"' SELECTED>"& rsUpdateEntry2("sucursal") &"</option>")
           else
                  response.write("<option value='"& rsUpdateEntry2("clave_sucursal")  &"'>"& rsUpdateEntry2("sucursal") &"</option>")
           end if
           rsUpdateEntry2.movenext
      wend
       rsUpdateEntry2.close
      response.write("</select> </td><td class=td-c>&nbsp;</td> <td class=td-c colspan=2>&nbsp;</td> </tr>")
      response.write("<tr><td class=td-c colspan=5>&nbsp;</td></tr>")

    response.write("<tr><td class=td-c colspan=5><input type='submit' value='actualizar'></form></td></tr></table>")
    rsUpdateEntry.close


end sub



%>

