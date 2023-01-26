<!--#include virtual="/config/conn.asp"-->

<%
Response.Expires = -1

response.write( "accion=" & Session("accion") & ",item=" &  Session("item")  &"<br>")

select case Session("accion") 
case "1" agregar
case "2" baja
case "3" editar
end select 


Sub agregar
			if  TRIM(request("vid"))<>"" then
					strSQL="SELECT * from Empleados where ID="& TRIM(request("vid"))
					'response.write strSQL 
					rsUpdateEntry.Open strSQL, adoCon

					if len(trim(request("vpaterno")))<3 or len(trim(request("vnombres")))<2 then

					            response.redirect("ShowContent.asp?contenido=1&msg=nombre o apellidos, demasiado corto")

					else
							if rsUpdateEntry.EOF then
											strSQL=  "INSERT INTO Empleados (ID,PATERNO,MATERNO,NOMBRES,FECHANACIMIENTO,FECHAINGRESO,CURP,RFC,NSS,SEXO,CLAVE_DEPTO,CLAVE_PUESTO,CLAVE_SUCURSAL,EMAIL_PERSONAL,CELULAR_PERSONAL,id_usuario,email,celular,CLABE) VALUES (" 
											strSQL= strSQL  & TRIM(request("vid")) &",'"  & TRIM(UCase(request("vpaterno"))) &"','"  & TRIM(UCase(request("vmaterno"))) &"','"  & TRIM(UCase(request("vnombres"))) &"','"  & request("vnacimiento") &"','"  
											strSQL= strSQL  & request("vingreso")  &"','"  & TRIM(UCase(request("vcurp"))) &"','"  & TRIM(UCase(request("vrfc"))) &"','"  & TRIM(request("vnss")) &"','"  & request("vsexo") &"','"   
											strSQL= strSQL  & request("vdepto") &"','"  & request("vpuesto") &"','"  & request("vsucursal") &"','" & request("vemailP") &"','" & request("vcelularP") &"','"&  request("vusuario")  &"','"&  request("vemail") & "','" & request("vcelular")   &"','"&  request("vBanco") & "')"
											response.write strSQL  
											rsUpdateEntry2.Open strSQL, adoCon
                                           
                                            if request("vusuario")<>"" then

													strSQL=  "INSERT INTO cat_usuario (id_usuario,clave) VALUES ('" & request("vusuario") &"','"  & request("vclave")   &"')"
													'response.write strSQL  
													rsUpdateEntry2.Open strSQL, adoCon

										    end if
											
											response.redirect("ShowContent.asp?contenido=40&accion=0&msg=el colaborador ha sido agregado")
							else
								      response.redirect("ShowContent.asp?contenido=40&accion=0&msg=el numero de colaborador ya se encuentra utilizado")
								
							end if

					end if
			else
				 
				   response.redirect("ShowContent.asp?contenido=40&accion=0&msg=el numero de colaborador esta en blanco")
				
			end if
end sub






sub baja
   vfecha=date()
   vfecha=request("vegreso")

   strSQL=  "UPDATE Empleados "
   strSQL= strSQL  & "SET FECHAEGRESO='"& year(vfecha) &"-" & month(vfecha) &"-"  & day(vfecha) &"',"   
   strSQL= strSQL  & "FormaBaja='"& request("vFormaBaja") &"' WHERE ID=" & Session("item")
   
   response.write strSQL 
   rsUpdateEntry.Open strSQL, adoCon
   response.redirect("ShowContent.asp?contenido=40&accion=0&msg=el colaborador no. "&  Session("item") &" ha sido dado de baja")

end sub







Sub editar

   strSQL=  "set dateformat dmy;UPDATE Empleados "
   strSQL= strSQL  & "SET clave_banco='"& TRIM(request("fBanco")) &"',PATERNO='"& TRIM(UCase(request("vpaterno")))  &"',MATERNO='"& TRIM(UCase(request("vmaterno"))) &"',NOMBRES='"& TRIM(UCase(request("vnombres"))) &"',CURP='"& TRIM(UCase(request("vcurp"))) &"',RFC='"& TRIM(UCase(request("vrfc"))) &"',NSS='"&  TRIM(request("vnss")) &"',CLABE='"&  TRIM(request("vBanco")) &"',SEXO='"& request("vsexo")  &"',CLAVE_DEPTO='"& request("vdepto") &"',CLAVE_PUESTO='"& request("vpuesto") &"',CLAVE_SUCURSAL='"& request("vsucursal") &"',"
   strSQL= strSQL  & "email='"& request("vemail")&"',ContrasenaTiendaLinea='"& request("fPassTienda") &"', CELULAR='"& request("vcelular") &"',EMAIL_PERSONAL='"& request("vemailP")&"',CELULAR_PERSONAL='"& request("vcelularP") &"', id_usuario='"&  request("vusuario") &"',FECHANACIMIENTO='"& request("vDia") &"-"& request("vMes") &"-"& request("vAnio")  &"',id_empresa='"& request("fempresa") &"' "

   if request("id_checador")<>"" then
   	    strSQL= strSQL  & ",id_checador="&  request("vchecador") &" "
   end if




   strSQL= strSQL  & "WHERE ID=" & Session("item")   
   response.write strSQL 
   rsUpdateEntry.Open strSQL, adoCon

    if request("vusuario")<>"" then
		    strSQL="SELECT * from cat_usuario where id_usuario='"& request("vusuario") &"'"
		    'response.write strSQL 
		    rsUpdateEntry2.Open strSQL, adoCon

		    if rsUpdateEntry2.EOF then

													strSQL=  "INSERT INTO cat_usuario (id_usuario,clave) VALUES ('" & request("vusuario") &"','"  & request("vclave")   &"')"
													response.write strSQL  
													rsUpdateEntry.Open strSQL, adoCon

		    else
		    										strSQL=  "UPDATE cat_usuario SET clave='"  & request("vclave")   &"' WHERE id_usuario='" &  request("vusuario") &"'"
													response.write strSQL  
													rsUpdateEntry.Open strSQL, adoCon

		    end if   
		    rsUpdateEntry2.close

    end if

   response.redirect("ShowContent.asp?contenido=40&accion=0&msg=el colaborador no. "& Session("item")  &" ha sido actualizado")
  
end sub	





%>
