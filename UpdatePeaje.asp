<!--#include virtual="/config/conn.asp"-->

<%
     
           Response.Expires = -1   

           if request("flag")="1" then    
               vEntidadOrigen=request("var1")
               vCiudadOrigen=request("var2")
               ventidad_destino=request("var3")
               vmpio_destino=request("var4")
               vciudad_destino=request("var5")
               vEjes=request("var6")
               vPeaje=trim(request("var7"))
               vPeaje=replace(vPeaje,",","")
               vPeaje=replace(vPeaje,"..",".")

               strSQL="select * from Peajes where Entidad_Origen='"&vEntidadOrigen&"' and Ciudad_Origen='"& vCiudadOrigen &"' and Entidad_Destino='"& ventidad_destino &"' and mpio_Destino='"& vmpio_destino &"' and ciudad_destino='"& vciudad_destino &"' "
               rsUpdateEntry4.Open strSQL, adoCon

               vID=rsUpdateEntry4("ID")  'ruta ID'
               rsUpdateEntry4.close

               strSQL="UPDATE Peajes set ["&vEjes&"]="&vPeaje&" where Entidad_Origen='"&vEntidadOrigen&"' and Ciudad_Origen='"& vCiudadOrigen &"' and Entidad_Destino='"& ventidad_destino &"' and mpio_Destino='"& vmpio_destino &"' and ciudad_destino='"& vciudad_destino &"' "
               'Response.write strSQL
               rsUpdateEntry5.Open strSQL, adoCon            

               Response.write("<a href='ShowContent.asp?contenido=109&action=2&ID=" & vID &"&ejes=" &vEjes &"'><img src='/img/animated-arrow.gif' border=0 alt='' title=''> DE CLICK AQUI PARA AVANZAR</a>")
          end if


          if request("flag")="2" then    
               vRuta=request("var1")
               vUnidad=request("var2")
               vTotalPeaje=request("var3")
               if trim(request("var4"))="" then
                    vTotalKms=0
               else
                    vTotalKms=int(trim(request("var4")))*2
               end if
               vSubtotalCOTI=0
               if trim(request("var5"))="" then
                   vSubtotalCOTI=0
               else
                   vSubtotalCOTI=int(trim(request("var5")))
               end if
              
               if vTotalKms<=0 then
                    Response.write("<P class='alert td-c'>debe ingresar los kil&oacute;metros para el c&aacute;lculo</P>")
               else

                    strSQL="select a.*,b.PrecioLitro from Flotilla a inner join  Combustibles b on a.Combustible=b.Combustible where a.id_unidad='" & vUnidad &"'"
                   'response.write strSQL
                   rsUpdateEntry.Open strSQL,adoCon

                   strSQL="select  cast(  cast(" & vTotalKms &"as float)/"& rsUpdateEntry("kms_litro") &"*"& rsUpdateEntry("PrecioLitro") &"*.05 as int)  as Mantenimiento, cast( cast("& vTotalKms &" as float)/"& rsUpdateEntry("kms_litro") &"*"& rsUpdateEntry("PrecioLitro") &" as int) as Combustibles"
                   'response.write(strSQL&"<br>")
                   rsUpdateEntry2.Open strSQL,adoCon

                   factor=0
                   vstring=trim(vTotalKms/600)
                   vpos=instr(vstring,".")
                   vstring=mid(vstring,vpos-1)
                   factor=1 + int( vstring )
                   vAlimentos=0
                   vAlimentos=700*factor
                   vHospedajes=0
                   vHospedajes=1000*(factor-1)
               
                   vCostosTotales=0

                   titulo="Calculo de flete"
                   dotitulo(titulo)

                   CreateTable(300)                   
                   response.write("<tr><td class='subtitulo td-r' width='70%'>Kms del roundtrip</td><td class='td-l' width='30%'>"& vTotalKms &"</td></tr>")
                   response.write("<tr><td class='subtitulo td-r'>dias en transito</td><td class='td-l'>"& factor &"</td></tr>")
                   
                   response.write("<tr><td class='subtitulo td-r'>Combustibles</td><td class='td-l'><font color=red> - "& formatcurrency(rsUpdateEntry2("Combustibles")) &"</td></tr>")
                   response.write("<tr><td class='subtitulo td-r'>Peajes</td><td class='td-l'><font color=red> - "& formatcurrency(vTotalPeaje) &"</td></tr>")
                   response.write("<tr><td class='subtitulo td-r'>Alimentos</td><td class='td-l'><font color=red> - "& formatcurrency(vAlimentos) &"</td></tr>")
                   response.write("<tr><td class='subtitulo td-r'>Hospedajes</td><td class='td-l'><font color=red> - "& formatcurrency(vHospedajes) &"</td></tr>")
                   response.write("<tr><td class='subtitulo td-r'>Mantenimiento (5%)</td><td class='td-l'><font color=red> - "& formatcurrency(rsUpdateEntry2("Mantenimiento"))  &"</td></tr>")

             

                   strSQL="select sum(" & rsUpdateEntry2("Combustibles") &"+"& vTotalPeaje &"+"& vAlimentos &"+"& vHospedajes &"+"& rsUpdateEntry2("Mantenimiento") &" ) as calculo"
                   rsUpdateEntry.close
                   rsUpdateEntry2.close
                   rsUpdateEntry.Open strSQL,adoCon

                   strSQL="select Rate from ORTT where RateDate=cast(getdate() as Date)"
                   rsUpdateEntry2.Open strSQL,adoCon2 'DMX'

                   strSQL="select round( cast("& rsUpdateEntry("calculo") &" as float)/"&rsUpdateEntry2("rate") &",2) as Flete_USD"
                    rsUpdateEntry5.Open strSQL,adoCon 

                  
                    
                   if vSubtotalCOTI>0 then
                        strSQL="select round(  ( cast("&rsUpdateEntry("calculo") &" as float)/ "& rsUpdateEntry2("rate") &" / cast(" & vSubtotalCOTI &" as float)   ) *100,2) as calculo"
                        'response.write strSQL
                        rsUpdateEntry3.Open strSQL,adoCon
                   end if

                   response.write("<tr><td class='subtitulo td-r'>FLETE (MXN)</td><td class='td-l strong'><font color=red> - "& formatcurrency(rsUpdateEntry("calculo"))  &"</td></tr>")
                   rsUpdateEntry.close
                   response.write("<tr><td class='subtitulo td-r'>FLETE (USD TC:"& rsUpdateEntry2("rate") &")</td><td class='td-l strong'><font color=red> - "& formatcurrency(rsUpdateEntry5("Flete_USD"))  &"</td></tr>")
                   rsUpdateEntry5.close
                   closetable

                   response.write("<P>&nbsp;</P>")
                   CreateTable(400)
                   if vSubtotalCOTI>0 then
                         response.write("<tr><td class='subtitulo td-r'>Cuanto representa de la cotizacion? </td><td class='td-l'><font color=green> + "& rsUpdateEntry3("calculo")  &" %</td></tr>")
                         rsUpdateEntry3.close
                   else
                         response.write("<tr><td class='subtitulo td-r'>Cuanto representa de la cotizacion? </td><td class='td-l'><font color=green> + ingrese subtotal COTI %</td></tr>")
                   end if
                   closetable
                   rsUpdateEntry2.close
                   

               end if
          end if

%>
