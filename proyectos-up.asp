<!--#include virtual="/config/conn.asp"-->

<%
Response.Expires = -1
if trim(request("item"))<>"" then
	  editar
else
	  agregar
end if	  


Sub agregar
			
											strSQL=  "INSERT INTO proyectos (PROYECTO,ESTATUS,FECHAREGISTRO) VALUES ('" & UCase(TRIM(request("vproyecto"))) &"','abierto',GETDATE() " & ")"
											response.write strSQL  
											rsUpdateEntry2.Open strSQL, adoCon4
											
											response.redirect("/modules/ShowContent.asp?contenido=4&msg=el nuevo proyecto de venta ha sido creado.")
						
end sub



Sub editar

   strSQL=  "UPDATE proyectos set ESTATUS='" & request("festatus") &"'  WHERE ID=" & request("item")   
   response.write strSQL 
   rsUpdateEntry.Open strSQL, adoCon4
   response.redirect("/modules/ShowContent.asp?contenido=4&msg=el proyecto no. "& request("item") &" ha sido actualizado")
  
end sub	
%>
