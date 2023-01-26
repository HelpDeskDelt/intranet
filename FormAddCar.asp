<!--#include virtual="/config/conn.asp"-->

<%
          strSQL="select * from Flotilla where id_unidad='"& request("ID") &"'"
          'response.write strSQL
          rsUpdateEntry.Open strSQL, adoCon

           DoAlert
           titulo="Editando veh&iacute;culo  "& request("ID")  &" del padr&oacute;n"
           DoTitulo(titulo)

           response.write("<form id='flotilla' action='ShowContent.asp' method='POST'>")
           response.write("<input type='hidden' name='contenido' value='87'>")
           response.write("<input type='hidden' name='action' value='5'>")

           response.write("<P class='td-c FontMed'><font class='subtitulo'>Tipo de activo: </font><select name='tipo'>")
           strSQL="select * from cat_activos"
           rsUpdateEntry2.Open strSQL, adoCon
           while not rsUpdateEntry2.EOF
                  if int(trim(rsUpdateEntry("tipo_activo")))=int(trim(rsUpdateEntry2("tipo_activo"))) then
                      response.write("<option value='"& rsUpdateEntry2("tipo_activo")  &"' SELECTED>"& rsUpdateEntry2("Activo") &"</option>")
                  else
                      response.write("<option value='"& rsUpdateEntry2("tipo_activo")  &"'>"& rsUpdateEntry2("Activo") &"</option>")
                  end if
                  rsUpdateEntry2.movenext
           wend
           rsUpdateEntry2.close
           response.write("</select></P>")


           CreateTable(960)
           response.write("<tr><td class='subtitulo td-r fontmed strong' width='7%'>ID:</td><td class='td-l fontmed' width='7%'><input type='text' name='fID' style='width:60px' value='"& rsUpdateEntry("ID_unidad") &"'></td><td class='subtitulo td-r fontmed strong' width='8%'>Fabricante:</td><td class='td-l fontmed' width='14%'><input type='text' name='fVendor' style='width:160px' value='"& rsUpdateEntry("Fabricante") &"'></td>")

           response.write("<td class='subtitulo td-r fontmed strong' width='9%'>Kilometraje:</td>")

           strSQL="select substring( CONVERT(VARCHAR,CONVERT(MONEY,max(kilometraje)),1),1,len(CONVERT(VARCHAR,CONVERT(MONEY,max(kilometraje)),1))-3)  as calculo  from check_unidad where ID_unidad='"& rsUpdateEntry("ID_unidad")  &"'"
           rsUpdateEntry5.Open strSQL, adoCon

           if rsUpdateEntry5.EOF then
                response.write("<td class='td-l fontmed' width='10%'>&nbsp;</td></tr>")
           else
                response.write("<td class='td-l fontmed' width='10%'><font color=red>"& rsUpdateEntry5("calculo") &"</font></td></tr>")
           end if
           rsUpdateEntry5.close
          
            response.write("<tr><td class='subtitulo td-r fontmed strong'>Raz&oacute;n Social:</td><td class='td-l fontmed'><select name='fRS'>")
            if rsUpdateEntry("RS")="DMX" then
                  response.write("<option value='DMX' SELECTED>DMX</option><option value='DFW'>DFW</option></select></td>")
            else
                  response.write("<option value='DMX'>DMX</option><option value='DFW' SELECTED>DFW</option></select></td>")
            end if

            response.write("<td class='subtitulo td-r fontmed strong'>Modelo:</td><td class='td-l fontmed'><input type='text' name='fModelo' style='width:200px' value='"& rsUpdateEntry("modelo") &"'></td><td class='subtitulo td-r fontmed strong'>Placas:</td><td class='td-l fontmed'><input type='text' name='fPlacas' style='width:160px' value='"& rsUpdateEntry("Placas") &"'></td></tr>")
 
            response.write("<tr><td class='subtitulo td-r fontmed strong'>Sucursal:</td><td class='td-l fontmed'><select name='fSucursal'>")

              strSQL=  "select * from cat_sucursal order by clave_sucursal"     
              'response.write strSQL  
              rsUpdateEntry2.Open strSQL, adoCon
      
              while not rsUpdateEntry2.EOF
                   if trim(rsUpdateEntry("clave_sucursal"))=trim(rsUpdateEntry2("clave_sucursal")) then
                        response.write("<option value='"& rsUpdateEntry2("clave_sucursal")  &"' SELECTED>"& rsUpdateEntry2("sucursalShort") &"</option>")
                   else
                        response.write("<option value='"& rsUpdateEntry2("clave_sucursal")  &"'>"& rsUpdateEntry2("sucursalShort") &"</option>")
                   end if
                   rsUpdateEntry2.movenext
              wend
               rsUpdateEntry2.close

            response.write("</select></td><td class='subtitulo td-r fontmed strong'>Descripcion:</td><td class='td-l fontmed'><input type='text' name='fDescripcion' style='width:270px' value='"& rsUpdateEntry("Descripcion") &"'></td><td class='subtitulo td-r fontmed strong'>Fecha vence Tarj-Circulacion:</td><td class='td-l fontmed'>  <select name=fDiaTC>")
             
               for i=1 to 31
                    if day(rsUpdateEntry("vigencia_TC"))=i then
                       response.write("<option value="& i &" SELECTED>"& i &"</option>")
                   else
                       response.write("<option value="& i &" >"& i &"</option>")
                   end if
               next
               response.write("</select> / <select name=fMesTC>")
               for i=1 to 12
                    if month(rsUpdateEntry("vigencia_TC"))=i then
                       response.write("<option value="& i &" SELECTED>"& meses(i) &"</option>")
                   else
                       response.write("<option value="& i &" >"&  meses(i) &"</option>")
                   end if
               next
               response.write("</select> / <select name=fAnioTC>")
               for i=year(date())+1 to year(date())-3 step -1
                   if year(rsUpdateEntry("vigencia_TC"))=i then
                       response.write("<option value="& i &" SELECTED>"& i &"</option>")
                   else
                       response.write("<option value="& i &" >"& i &"</option>")
                   end if
               next
               response.write("</select> </td></tr>")

             strSQL="select * from cat_entidades order by id_entidad"
              'response.write strSQL  
              rsUpdateEntry2.Open strSQL, adoCon

              response.write("<tr><td class='subtitulo td-r fontmed strong'># Total llantas:</td><td class='td-l fontmed'><select name='fEjes'>")

              for i=0 to 20 step 1
                 if int(trim(rsUpdateEntry("Total_Llantas")))=i then
                    response.write("<option value='"&i&"' SELECTED>"&i&"</option>")
                 else
                    response.write("<option value='"&i&"'>"&i&"</option>")
                 end if
                 
              next
              response.write("</select></td><td class='subtitulo td-r fontmed strong'>Serie:</td><td class='td-l fontmed'><input type='text' name='fSerie' style='width:220px' value='"& rsUpdateEntry("serie") &"'></td><td class='subtitulo td-r fontmed strong'>Ent Federativa Tarj-Circulacion</td><td class='td-l fontmed'><select name='fEstadoTC'>")

              while not rsUpdateEntry2.EOF
                 if rsUpdateEntry("id_entidad_TC")=rsUpdateEntry2("id_entidad") then                  
                    response.write("<option value='"& rsUpdateEntry2("id_entidad") &"' SELECTED>"& rsUpdateEntry2("entidad") &"</option>")
                 else
                    response.write("<option value='"& rsUpdateEntry2("id_entidad") &"'>"& rsUpdateEntry2("entidad") &"</option>")
                 end if
                 rsUpdateEntry2.MoveNext
              wend
              rsUpdateEntry2.close
              response.write("</select></td></tr>")
 
              response.write("<tr><td class='subtitulo td-r fontmed strong'>Plataforma: </td><td class='td-l fontmed'><select name='fPlataforma'>")
              if trim(rsUpdateEntry("Plataforma"))="S" then
                   response.write("<option value='N'>no</option><option value='S' SELECTED>si</option></select></td>")
              else
                   response.write("<option value='N' SELECTED>no</option><option value='S'>si</option></select></td>")
              end if
              response.write("<td class='subtitulo td-r fontmed strong'>Fecha compra: </td><td class='td-l fontmed'>  <select name=fDia>")
             
               for i=1 to 31
                    if day(rsUpdateEntry("FechaCompra"))=i then
                       response.write("<option value="& i &" SELECTED>"& i &"</option>")
                   else
                       response.write("<option value="& i &" >"& i &"</option>")
                   end if
               next
               response.write("</select> / <select name=fMes>")
               for i=1 to 12
                    if month(rsUpdateEntry("FechaCompra"))=i then
                       response.write("<option value="& i &" SELECTED>"& meses(i) &"</option>")
                   else
                       response.write("<option value="& i &" >"&  meses(i) &"</option>")
                   end if
               next
               response.write("</select> / <select name=fAnio>")
               for i=year(date()) to year(date())-30 step -1
                   if year(rsUpdateEntry("FechaCompra"))=i then
                       response.write("<option value="& i &" SELECTED>"& i &"</option>")
                   else
                       response.write("<option value="& i &" >"& i &"</option>")
                   end if
               next
               response.write("</select> </td><td class='subtitulo td-r fontmed strong'>Aseguradora: </td><td class='td-l fontmed'><input type='text' class='anchox1' name='fSeguro' value='"& rsUpdateEntry("seguro") &"'></td></tr>")

               response.write("<tr><td class='subtitulo td-r fontmed strong'>Max carga (kg): </td><td class='td-l fontmed'> <input type='number' name='fMaxCarga' value='"& rsUpdateEntry("MaxCarga") &"' style='width: 70px'> <td class='subtitulo td-r fontmed strong'>Monto compra: </td><td class='td-l fontmed'><input type='text' name='fMonto' value='"& rsUpdateEntry("MontoCompra") &"'></td></td><td class='subtitulo td-r fontmed strong'>Fecha vence seguro: </td><td class='td-l fontmed'><select name=fDiaSG>")
             
               for i=1 to 31
                    if day(rsUpdateEntry("Vigencia_Seguro"))=i then
                       response.write("<option value="& i &" SELECTED>"& i &"</option>")
                   else
                       response.write("<option value="& i &" >"& i &"</option>")
                   end if
               next
               response.write("</select> / <select name=fMesSG>")
               for i=1 to 12
                    if month(rsUpdateEntry("Vigencia_Seguro"))=i then
                       response.write("<option value="& i &" SELECTED>"& meses(i) &"</option>")
                   else
                       response.write("<option value="& i &" >"&  meses(i) &"</option>")
                   end if
               next
               response.write("</select> / <select name=fAnioSG>")
               for i=year(date())+1 to year(date())-3 step -1
                   if year(rsUpdateEntry("Vigencia_Seguro"))=i then
                       response.write("<option value="& i &" SELECTED>"& i &"</option>")
                   else
                       response.write("<option value="& i &" >"& i &"</option>")
                   end if
               next
               response.write("</select> </td></tr>")
               response.write("<tr><td class='subtitulo td-r fontmed strong'>Capac Tanque (lts):</td><td class='td-l fontmed'><select name='ftanque'>")
              for i=400 to 40 step -5
                    if int(trim(rsUpdateEntry("CapacTanque")))=i then
                       response.write("<option value="& i &" SELECTED>"& i &"</option>")
                   else
                       response.write("<option value="& i &" >"& i &"</option>")
                   end if
              next
              response.write("</select><td class='subtitulo td-r fontmed strong'>Combustible:</td><td class='td-l fontmed'><select name='fcombustible'>")
           
                  if rsUpdateEntry("Combustible")="gasolina"   then
                        response.write("<option value='gasolina' SELECTED>gasolina</option>")
                        response.write("<option value='diesel'>diesel</option>")
                        response.write("<option value='gas'>gas</option>")
                  end if

                  if rsUpdateEntry("Combustible")="diesel" then
                        response.write("<option value='gasolina'>gasolina</option>")
                        response.write("<option value='diesel' SELECTED>diesel</option>")
                         response.write("<option value='gas'>gas</option>")
                  end if

                  if rsUpdateEntry("Combustible")="gas"  then
                        response.write("<option value='gasolina'>gasolina</option>")
                        response.write("<option value='diesel'>diesel</option>")
                         response.write("<option value='gas' SELECTED>gas</option>")
                  end if
            

              response.write("</select><td class='subtitulo td-r fontmed strong'>Tarj Combustible:</td><td class='td-l fontmed'><input type='text' name='ftarjeta' value='"& rsUpdateEntry("TarjetaCombustible") &"'></tr>")



              response.write("<tr><td colspan=6 class='td-c'><input type='submit' value='actualizar'></form></td></tr>")

           closeTable

 
%>
<input type="button" value="cerrar" onclick="hidediv('edit')"> </P>   