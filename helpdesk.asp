<!--#include virtual="/config/conn.asp"-->
<!--#include virtual="/header.asp"-->

<%

if Session("DocType")="" then
     Session("DocType")="I"
end if


if request("DocType") <>"" then
       Session("DocType")=request("DocType")  
end if


if request("items") <> "" then
    Session("items")=request("items")   
else
    Session("items")="pedido"
    Session("Modulo")="ventas"
end if


if  Session("items")="OC"   then
       Session("Modulo")="compras"
       Session("items") ="pedido"
else
       Session("Modulo")="ventas"
end if

if  Session("items")="entrada"   then
       Session("Modulo")="compras"
       Session("items") ="entrada"      
end if



doAlert
titulo="<a href='helpdesk.asp' class='H1'>DOCUMENTOS DE MARKETING EN SAP</a>"
DoTitulo(titulo)




if request("action") ="" then
    MenuDocuments
end if





if request("action") =2 then   'form de editar'
    if request("RS")="MME" then
        vEmpresa="MEIDE"
    else
        vEmpresa=request("RS")
    end if
    strSQL="select * from PRODUCTIVA_"& vempresa &".dbo.ORDR where DocNum=" & request("DocNum")
    rsUpdateEntry.Open strSQL, adoCon4
    
    response.write("<form id='helpdesk' action='helpdesk.asp' method='POST'>")

    response.write("<input type='hidden' name='action' value=3>")
    response.write("<input type='hidden' name='RS' value='"& request("RS") &"'>")
    response.write("<input type='hidden' name='DocNum' value="& request("DocNum") &">")
    response.write("<table align='center' width='760px' border=1><tr><td class='td-l'>Raz&oacute;n Social: <B>" & request("RS") &"</B></td></tr>")

    response.write("<tr><td class='td-l'>Pedido: <B>" & request("DocNum") & "</B></td></tr>")

    response.write("<tr><td class='td-l'>Seguimiento al suministro: <input style='width: 600px'  type='text' name='seguimiento' value='"& rsUpdateEntry("U_COMINTERNO") &"'>"  &"</td></tr>")
    response.write("<tr><td class='td-l'>eMail notificacion de embarques: <input style='width: 560px'  type='text' name='emailEmbarques' value='"& rsUpdateEntry("U_eMailEmbarques") &"'>"  &"</td></tr>")

    response.write("<tr><td class='td-l'>Instrucciones ventas: <font color=red>"& rsUpdateEntry("U_SEGUIMIENTO")   &"</td></tr></table>")
    response.write("</P><P class='td-c'><input type='submit' value='actualizar'></form></P>")

    rsUpdateEntry.close
end if




if request("action") =3 then  'edit UP'
    if request("RS")="MME" then
        vEmpresa="MEIDE"
    else
        vEmpresa=request("RS")
    end if
   
    strSQL="update PRODUCTIVA_"& vempresa &".dbo.ORDR set U_COMINTERNO='"& request("seguimiento")  &"',U_eMailEmbarques='"& request("emailEmbarques") &"' where docnum="& request("DocNum")
    'response.write( strSQL &"<BR>")
    rsUpdateEntry.Open strSQL, adoCon4  
    response.redirect("helpdesk.asp?msg=se actualizo pedido "& trim(request("DocNum")) & " en " & request("RS")  )
   

end if




if request("action") =4 then
         'hcontrol=1 se ha autorizado para envio de email'

      if  request("InterEmpresas")="N" then
     
         strSQL="set dateformat ymd;UPDATE notificaciones set Hcontrol=1,control='"& UCASE(request("fnotas")) &"',detalles='"& mid(request("fdetalles"),1,8000) &"',AreaCosteo="& request("fcosteo") &" where tipo='entrada' and RazonSocial='"&request("RS") &"' and Pedimento='"& request("Pedimento") &"' and DocDate='"& request("fecha") &"' and InterEmpresas='"& request("InterEmpresas")&"'  "
        
      else

          strSQL="set dateformat ymd;UPDATE notificaciones set Hcontrol=1,control='"& UCASE(request("fnotas")) &"',detalles='"& mid(request("fdetalles"),1,8000) &"',AreaCosteo="& request("fcosteo") &" where tipo='entrada' and Pedimento='"& request("Pedimento") &"' and DocDate='"& request("fecha") &"' and InterEmpresas='"& request("InterEmpresas")&"'  "

      end if
        response.write strSQL

       rsUpdateEntry.Open strSQL, adoCon4 
       response.redirect("helpdesk.asp?items=entrada&msg=la aprobacion de entradas de mercancias (pedimento "& request("Pedimento") &"), fue registrada")   

end if     



if request("action") =5 then
     'notificar costeo'

     strSQL="set dateformat ymd;update notificaciones set Hcontrol=2, DetallesCosteo='"& request("fdetalles") &"' where tipo='entrada' and RazonSocial='"&request("RS") &"' and Pedimento='"& request("Pedimento") &"' and DocDate='"& request("fecha") &"'"
     response.write strSQL

     rsUpdateEntry7.Open strSQL, adoCon4 
     response.redirect("helpdesk.asp?items=entrada&msg=la aprobacion de costeo (pedimento "& request("Pedimento") &"), fue enviada por email")
   
end if     



'response.write("<P>&nbsp;</P>")

sub MenuDocuments

    
'separador
response.write("<center>")

        select case Session("Modulo")
          case "ventas"
              select case Session("items")
              case "pedido"   response.write("<h3>"& UCase(Session("Modulo")) &"- " & UCase(Session("items")) & "s con partidas abiertas (actualizacion: "& now() &" )  </h3>  </center>")
              case "remision" response.write("<h3>"&  UCase(Session("Modulo")) &"- " & UCase(Session("items")) & "es con partidas abiertas (actualizacion: "& now() &" ) </h3> </center>") 
              case "factura" 
                    response.write("<form id='busqueda' method='POST' action='helpdesk.asp'><input type='hidden' name='items' value='factura'>")
                    response.write("<h3>"&  UCase(Session("Modulo")) &" FACTURAS ("& now() &" ) &nbsp;&nbsp;buscar por: [factura] <input class='anchoSmall' type='text' name='ffactura' value='"& request("ffactura") &"'>&nbsp;&nbsp;[remision] <input class='anchoSmall' type='text' name='fremision' value='"& request("fremision") &"'>&nbsp;&nbsp; [pedido]  <input class='anchoSmall' type='text' name='fpedido' value='"& request("fpedido") &"'>   &nbsp;&nbsp;<input type='submit' value='buscar'> </form></h3> </center>")
              end select 

          case "compras"
               select case Session("items")
                   case "pedido"   
                       response.write("<form id='busqueda2' method='POST' action='helpdesk.asp'><input type='hidden' name='items' value='OC'>")
                       response.write("<h3>"&  UCase(Session("Modulo")) &": Ordenes de compra con Partidas Abiertas ("& now() &" )  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;buscar OC: <input class='anchoSmall' type='text' name='fOC' value='"& request("fOC") &"'>&nbsp;&nbsp;&nbsp;<input type='submit' value='buscar'> </form></h3>  </center>")
                end select 
        end select 
  
  
  flag=0



  'QUERY'
  strSQL="SELECT a.*,datediff(day,a.DocDate,getdate()) as DELAY,b.Ruta,dbo.fn_GetMonthName(a.DocDate,'spanish') AS [DocDate2] "
  strSQL= strSQL & "FROM [SBOTemp].dbo.Notificaciones a left join [SBOTemp].dbo.Repositorio b on a.RazonSocial=b.RazonSocial and a.CardCode collate SQL_Latin1_General_CP1_CI_AS=b.CardCode collate SQL_Latin1_General_CP1_CI_AS "

  if request("asesor")<>"" then
      'response.write request("asesor")
      strSQL= strSQL & "inner join PRODUCTIVA_DMX.dbo.ORDR c on a.docnum=c.docnum and c.slpcode="&request("asesor")&" "
  end if


  strSQL= strSQL & "WHERE a.CANCELED='N' AND a.tipo='"&   Session("items")  &"' "  

  if request("docnum")<>"" then
        strSQL= strSQL & " AND a.docnum=" & request("docnum") &" "
  end if

  if request("ffactura")<>"" then
      strSQL= strSQL & " AND a.docnum=" & request("ffactura") &" "
  end if

  if request("fremision")<>"" then
      strSQL= strSQL & " AND a.remision=" & request("fremision") &" "
  end if

   if request("fpedido")<>"" then
      strSQL= strSQL & " AND a.pedido=" & request("fpedido") &" "
  end if

  if request("fOC")<>"" then
      strSQL= strSQL & " AND a.docnum=" & request("fOC") &" "
      'response.write strSQL
  end if

  if request("RS")<>"" then
        strSQL= strSQL & " AND a.RazonSocial='" & request("RS") &"' "
  end if

  if Session("items")<>"factura" then   
     strSQL= strSQL & "AND a.DocStatus='O' "
  end if

  if  request("ffactura")="" and request("fremision")="" and request("fpedido")="" and Session("items")<>"pedido"  and Session("items")<>"remision"  then    
      strSQL= strSQL & "AND datediff(DAY,a.DocDate,getdate() )<5 "      
  end if
  
  strSQL= strSQL & " AND a.Modulo='"& Session("Modulo") &"'"
  if  Session("Modulo")="compras" then 
       strSQL= strSQL & " AND a.DocType='"& Session("DocType") &"'"
  end if

  strSQL= strSQL & " ORDER BY a.DocDate DESC"
  if Session("items")="factura" then
       strSQL= strSQL & ",a.RequestDate Desc" 
  end if
 
  if Session("Modulo")="compras" then
     'response.write strSQL  
  end if

  'response.write strSQL   'imprimir query'
  rsUpdateEntry.Open strSQL, adoCon4
  

  if Session("items")="entrada" then
     response.write("<P>&nbsp;</P>")
  end if

 
  if Session("items")="entrada" then
       response.write("<H1>ENTRADAS DE MERCANCIA</H1>")
       response.write("<table width='1180px' cellpadding=4 cellSpacing=2 align=center border=0>")

       response.write("<tr>")

       'Sin inter-empresas'
       strSQL="if exists (select * from INFORMATION_SCHEMA.TABLES where TABLE_NAME = '_TMP_ENTRADAS' AND TABLE_SCHEMA = 'dbo')     drop table dbo._TMP_ENTRADAS;"
       'response.write("<font color=red>" & strSQL &"</font><BR>")
       rsUpdateEntry6.Open strSQL, adoCon4


       'Con inter-empresas'
       strSQL="set dateformat ymd;select a.MargenUtil,a.AreaCosteo,a.detalles,a.Pedimento,a.DocDate,a.RazonSocial,a.Hcontrol,a.sentDate,case when sentDate is null then 0 else 1 end as estatus,a.interempresas,Entradas=substring(   (select convert(varchar,DocNum)+',' from Notificaciones where Pedimento=a.Pedimento and Docdate=a.DocDate and RazonSocial=a.RazonSocial and InterEmpresas='N' for XML PATH ('') ) ,1, len( (select convert(varchar,DocNum)+',' from Notificaciones where Pedimento=a.Pedimento and Docdate=a.DocDate and RazonSocial=a.RazonSocial and InterEmpresas='N' for XML PATH ('') )  )-1),  Control=(select count(*) from Notificaciones where Pedimento=a.Pedimento and Docdate=a.DocDate and RazonSocial=a.RazonSocial  and InterEmpresas='N' for XML PATH ('') ) INTO _TMP_ENTRADAS from Notificaciones a where a.tipo='entrada'  and datediff(DAY,a.Docdate,getdate())<30 and InterEmpresas='N' group by a.MargenUtil,a.AreaCosteo,a.detalles,a.DocDate,a.Pedimento,a.RazonSocial,a.sentdate,a.Hcontrol,a.interempresas order by a.DocDate DESC,a.RazonSocial,a.Pedimento "
       'response.write("<font color=red>" & strSQL &"</font><BR>")
       rsUpdateEntry5.Open strSQL, adoCon4


       strSQL="set dateformat ymd;INSERT INTO _TMP_ENTRADAS select a.MargenUtil,a.AreaCosteo,a.detalles,a.Pedimento,a.DocDate,a.RazonSocial,a.Hcontrol,a.sentDate,case when sentDate is null then 0 else 1 end as estatus,a.interempresas,Entradas=substring(   (select convert(varchar,DocNum)+',' from Notificaciones where Pedimento=a.Pedimento and Docdate=a.DocDate and RazonSocial=a.RazonSocial and InterEmpresas='Y' for XML PATH ('') ) ,1, len( (select convert(varchar,DocNum)+',' from Notificaciones where Pedimento=a.Pedimento and Docdate=a.DocDate and RazonSocial=a.RazonSocial and InterEmpresas='Y'  for XML PATH ('') )  )-1),  Control=(select count(*) from Notificaciones where Pedimento=a.Pedimento and Docdate=a.DocDate and RazonSocial=a.RazonSocial  and InterEmpresas='Y'  for XML PATH ('') )  from Notificaciones a where a.tipo='entrada'  and datediff(DAY,a.Docdate,getdate())<30 and InterEmpresas='Y' group by a.MargenUtil,a.AreaCosteo,a.detalles,a.DocDate,a.Pedimento,a.RazonSocial,a.sentdate,a.Hcontrol,a.interempresas order by a.DocDate DESC,a.RazonSocial,a.Pedimento "
       'response.write("<font color=red>" & strSQL &"</font><BR>")
       rsUpdateEntry4.Open strSQL, adoCon4

       strSQL="DELETE _TMP_ENTRADAS where Entradas is null or Pedimento is null"
       rsUpdateEntry4.Open strSQL, adoCon4
      
       strSQL="SELECT * FROM _TMP_ENTRADAS WHERE 1=1 "

       if request("pedimento")<>"" then
           strSQL= strSQL&" and Pedimento='"& request("pedimento") &"' and RazonSocial='"& request("RazonSocial") &"' and DocDate='"& request("fecha") &"' and InterEmpresas='"& request("InterEmpresas") &"' "
       end if

       strSQL= strSQL & " ORDER BY DOCDATE DESC "
       rsUpdateEntry3.Open strSQL, adoCon4
       'response.write("<font color=red>" & strSQL &"</font><BR>")
       
    
       response.write("<td width='10%' class='subtitulo'>Pedimento</td>")
       response.write("<td width='4%' class='subtitulo'>Folio</td>")
       response.write("<td width='10%' class='subtitulo'>Fecha</td>")
       response.write("<td width='3%' class='subtitulo'>RS</td>")
       response.write("<td width='20%' class='subtitulo'>Proveedor</td>")
       response.write("<td width='5%' class='subtitulo'>Almacen</td>")
       response.write("<td width='5%' class='subtitulo'>Inter <br>Empresas</td>")
       response.write("<td width='2%' class='subtitulo'>O/C</td>")
       response.write("<td width='6%' class='subtitulo'>costear?</td>")
       response.write("<td width='16%' class='subtitulo'>Estatus</td>")       
       response.write("<td width='14%' class='subtitulo'>Acciones disponibles</td>")
       response.write("</tr>")

       i=1
       n=1




       while not rsUpdateEntry3.EOF
           
           response.write("<tr>")
           response.write("<td class='td-r' style='vertical-align: middle;' rowspan="& rsUpdateEntry3("control") &"><B>"& rsUpdateEntry3("Pedimento") &"</B></td>")
           strSQL="select DocStatus,Pedimento,Docnum as Folio,UPPER(dbo.fn_GetMonthName(DocDate,'spanish')) as Fecha,RazonSocial,CardName as Proveedor,WhsCode,InterEmpresas from Notificaciones where tipo='entrada' and RazonSocial='"& rsUpdateEntry3("RazonSocial") &"' and Docnum in ("& rsUpdateEntry3("entradas") &") and CANCELED='N' and InterEmpresas='"& rsUpdateEntry3("InterEmpresas") &"' order by Docnum desc"
             if n<3 then
                 'response.write("<font color=blue>" & strSQL &"</font><BR>")
             end if
             n=n+1
             'response.write strSQL
             rsUpdateEntry4.Open strSQL, adoCon4
            

           for i=1 to rsUpdateEntry3("control") step 1
                 if i>1 then 
                       response.write("<tr>")
                 end if

                 response.write("<td class='td-c fonttiny'>"& rsUpdateEntry4("Folio")  &" </td>")
                 response.write("<td class='td-c fonttiny'>"& rsUpdateEntry4("Fecha") &"</td>")
                 response.write("<td class='td-c fonttiny'>"& rsUpdateEntry4("RazonSocial") &"</td>")
                 response.write("<td class='td-l fonttiny'>"& mid(rsUpdateEntry4("Proveedor"),1,30) &"</td>")
                 response.write("<td class='td-c fonttiny'>"& rsUpdateEntry4("WhsCode") &"</td>")
                 response.write("<td class='td-c fonttiny'>"& rsUpdateEntry4("InterEmpresas") &"</td>")
                 response.write("<td class='td-c fonttiny'>"& rsUpdateEntry4("DocStatus") &"</td>")


                     'cuando va a solicitar una accion'
                 if request("pedimento")<>"" then
                     if trim(i)=trim(rsUpdateEntry3("control")) then

                             response.write("<form action='helpdesk.asp' method='post'>")           
                             response.write("<input type='hidden' name='RS' value='"& rsUpdateEntry4("RazonSocial") &"'>") 
                             
                             response.write("<input type='hidden' name='pedimento' value='"& rsUpdateEntry4("pedimento") &"'>") 
                             response.write("<input type='hidden' name='fecha' value='"& request("Fecha") &"'>")
                             response.write("<input type='hidden' name='InterEmpresas' value='"& request("InterEmpresas") &"'>")
                           

                         if request("estatus")=1  then  'va aprobar las entradas'
                             response.write("<input type='hidden' name='action' value=4>") 
          
                             response.write("<td class='td-c fonttiny' colspan=2 rowspan="& rsUpdateEntry3("control") &">Instrucciones en el email: <BR><input type='text' class='anchox2' name='fnotas' value=''></td>")

                             response.write("<td class='td-c' rowspan="& rsUpdateEntry3("control") &">&nbsp;</td></tr>")
                             response.write("<tr><td class='td-c subtitulo' colspan=7>Detalles de <br>las entradas</td><td class='td-c subtitulo' colspan=4>Requiere costeo</td></tr>")  

                             response.write("<tr><td class='td-c' colspan=7><textarea name='fdetalles' rows=5 cols=80></textarea></td><td class='td-l fonttiny' style='padding-left: 30px'><input class='fonttiny' type='radio' name='fcosteo' value=0>no requiere<BR><input class='fonttiny' type='radio' name='fcosteo' value=1 checked>por Importaciones<BR><input class='fonttiny' type='radio' name='fcosteo' value=2>por traslados<BR><BR><input class='fonttiny' type='submit' value='enviar eMail'></form></td></tr>")
                         end if





                         if request("estatus")=2 then  'va aprobar COSTEO'
                             response.write("<input type='hidden' name='action' value=5>") 
          
                             response.write("<td class='td-c' rowspan="& rsUpdateEntry3("control") &" colspan=3><B>Detalles en el email: </B><BR><textarea name='fdetalles' rows=5 cols=60></textarea><BR><BR><input type='submit' value='enviar email de aprobacion'></form></td></tr>")

                             
                         end if

                        
                           
   

                     end if
                 else




                     'en listado completo'
                     if i=1 then

                          
                         response.write("<td class='td-c' rowspan="& rsUpdateEntry3("control") &">")
                         if rsUpdateEntry3("AreaCosteo")>0 then
                             response.write("si</td>")
                         else
                             response.write("no</td>")
                         end if
                      

                        select case rsUpdateEntry3("estatus")
                          case 0  response.write("<td class='td-l' rowspan="& rsUpdateEntry3("control") &">no se enviado notificacion por email</td>")'sentdate is null'                            
                          case 1  response.write("<td class='td-l' rowspan="& rsUpdateEntry3("control") &"><font color=red>se envi&oacute; email DE ENTRADAS-OK en "& rsUpdateEntry3("sentDate") &" </font></td>")
                        end select

                        select case rsUpdateEntry3("hcontrol")
                          case 0  response.write("<td class='td-l' rowspan="& rsUpdateEntry3("control") &"><a href='helpdesk.asp?items=entrada&estatus=1&RazonSocial="& rsUpdateEntry3("RazonSocial") &"&pedimento="& rsUpdateEntry3("Pedimento") &"&Fecha="& year(rsUpdateEntry3("DocDate"))&"-"& month(rsUpdateEntry3("DocDate"))&"-"& day(rsUpdateEntry3("DocDate"))  &"&InterEmpresas="& rsUpdateEntry3("InterEmpresas") &"'><U>aprobar notificaci&oacute;n de entradas</U></a></td>")
                          case 1  
                             if rsUpdateEntry3("AreaCosteo")>0 then
                               response.write("<td class='td-c' rowspan="& rsUpdateEntry3("control") &"><a href='helpdesk.asp?items=entrada&estatus=2&RazonSocial="& rsUpdateEntry3("RazonSocial") &"&pedimento="& rsUpdateEntry3("Pedimento") &"&Fecha="& year(rsUpdateEntry3("DocDate"))&"-"& month(rsUpdateEntry3("DocDate"))&"-"& day(rsUpdateEntry3("DocDate"))  &"&InterEmpresas="& rsUpdateEntry3("InterEmpresas") &"'><U>APROBAR COSTEO</U></a></td>")
                             else
                                response.write("<td class='td-l' rowspan="& rsUpdateEntry3("control") &">&nbsp;</td>")

                             end if
                          case 2  response.write("<td class='td-c' rowspan="& rsUpdateEntry3("control") &"><font color=green>se aprob&oacute; costeo en "& rsUpdateEntry3("MargenUtil") &"</font></td>")
                          end select
                          
                     end if
                  end if

                 response.write("</tr>")
                 rsUpdateEntry4.MoveNext
           next

          if request("pedimento")="" then
             if rsUpdateEntry3("detalles")<>"" then
                  response.write("<tr><td class='td-r subtitulo'>detalles en las entradas:</td><td colspan=10 class='td-l' style='background-color: red; color:white;'>" & rsUpdateEntry3("detalles") &"&nbsp;</td></tr>")
             else
                 response.write("<tr><td class='td-r subtitulo'>detalles en las entradas:</td><td colspan=10 class='td-l'>" & rsUpdateEntry3("detalles") &"&nbsp;</td></tr>")
             end if
          end if
          response.write("<tr><td colspan=20><HR></td></tr>")
          rsUpdateEntry3.MoveNext
          rsUpdateEntry4.close
          
             
       wend
       rsUpdateEntry3.close
       response.write("</table><P>&nbsp;</P>")



  else  


              response.write("<table width='1160px' cellpadding=4 cellSpacing=2 align=center border=1")
              
              response.write("<tr>")


              if Session("items")<>"factura" then
                 response.write("<td width='2%' class='subtitulo'>#</td>")
              end if
              response.write("<td width='3%' class='subtitulo'>RS</td>")
              response.write("<td width='8%' class='subtitulo'>Folio</td>")
              response.write("<td width='12%' class='subtitulo'>Socio negocio</td>")
              response.write("<td width='8%' class='subtitulo'>Fecha</td>")
              response.write("<td width='3%' class='subtitulo'>D&iacute;as abierto</td>")
              response.write("<td width='13%' class='subtitulo'>Notificaci&oacute;n eMail</td>")
              response.write("<td width='11%' class='subtitulo'>Acciones disponibles</td>")
              response.write("<td width='4%' class='subtitulo'>interE</td>")

             if Session("Modulo")="ventas" then
                  'response.write Session("items")
                  select case trim(Session("items"))
                    case "pedido"    
                           response.write("<td width='9%' class='subtitulo'>Ruta</td>")
                           response.write("<td width='9%' class='subtitulo'>Proyecto</td>")
                    case "remision"  
                           response.write("<td width='9%' class='subtitulo'>Facturar?</td>")
                           response.write("<td width='9%' class='subtitulo'>pedido</td>")
                           response.write("<td width='9%' class='subtitulo'>Proyecto</td>")
                    case "factura"   
                           response.write("<td width='9%' class='subtitulo'>remision</td>")
                           response.write("<td width='9%' class='subtitulo'>pedido</td>")
                           response.write("<td width='9%' class='subtitulo'>Proyecto</td>")
                           response.write("<td width='4%' class='subtitulo'>Mode</td>")
                  end select
              
              end if


              if Session("Modulo")="compras" then
                  select case Session("items")
                    case "pedido"    response.write("<td width='9%' class='subtitulo' style='text-align: left;'>Referencia</td>")      
                  end select
                  select case Session("items")
                    case "pedido"    response.write("<td width='9%' class='subtitulo'>Proyecto</td>")
                  
                  end select


              end if

  end if
  response.write("</tr>") 



 i=1
 vDia=0
 if not rsUpdateEntry.EOF then
     vdia=day(rsUpdateEntry("docdate"))
 end if

  while not rsUpdateEntry.EOF

      vTD="td1"
     
      if rsUpdateEntry("RazonSocial")="DMX" then
           vTD="td2"
      end if

      response.write("<tr>")
      if Session("items")<>"factura" then
           response.write("<td class="& vTD &">" & i & "</B></td>") 
      end if
      response.write("<td class="& vTD &"><B>" & rsUpdateEntry("RazonSocial")  & "</B></td>" )
      if Session("Modulo")="ventas" then
            select case Session("items") 
            case "pedido"   
                  response.write("<td class='"& vTD &"'><a href='ShowContent.asp?contenido=68&action=1&fRS="& rsUpdateEntry("RazonSocial") &"&fpedido="& rsUpdateEntry("DocNum") &"' target='pedido'> PED-"& rsUpdateEntry("DocNum") &"<img src='/img/down.gif' border=0 title='' alt=''></a></td>")
           
            case "remision"    
                    response.write("<td class='"& vTD &"'><a href='ShowContent.asp?contenido=124&action=1&fRS="& rsUpdateEntry("RazonSocial") &"&fremision="& rsUpdateEntry("DocNum") &"' target='remision'> REM-"& rsUpdateEntry("DocNum") &"<img src='/img/down.gif' border=0 title='' alt=''></a></td>")

            case "factura"    
                      response.write("<td><a href='ShowContent.asp?contenido=119&action=1&fRS="&rsUpdateEntry("RazonSocial")&"&ffactura="&rsUpdateEntry("Docnum")&"' target='factura'>F-" & rsUpdateEntry("DocNum")&"<img src='/img/down.gif' border=0 ></a></td>")
            end select            
      end if
      if Session("Modulo")="compras" then
            select case Session("items") 
            case "pedido" %>  
                      <td class="<%=vTD%>"><a href="javascript:sendReq('/modules/ShowDetails.asp','rs,OC','<%=rsUpdateEntry("RazonSocial")%>,<%=rsUpdateEntry("DocNum")%>','BoxRoundedDetail')"> OC-<%=rsUpdateEntry("DocNum")%>  <img src="/img/down.gif" border=0 title="" alt=""> </a> </td>   <%      
            end select            
      end if

      response.write("<td class='td-c fonttiny'>" & mid(rsUpdateEntry("CardName"),1,30)  & "</td> " )
      response.write("<td class='td-c fonttiny' style='width:110px'>" & rsUpdateEntry("DocDate2")  & "</td> " )
      response.write("<td class='td-c fontmed'>" & rsUpdateEntry("DELAY")  & "</td> " )


      if Session("Modulo")="ventas" then
           
          if  rsUpdateEntry("Tipo") ="pedido" then
              if rsUpdateEntry("Control") ="ALERT: " then
                   response.write("<td class="& vTD &"><font class='fonttiny'>Enviado 1 vez: " & rsUpdateEntry("sentdate")  & "</font></td> " )
              else
                   response.write("<td class="& vTD &" '><font class='fonttiny'><B>[ACTUALIZADO] " & rsUpdateEntry("sentdate")  & "</font></B></td> " )
              end if

              response.write("<td class="& vTD &"><a href='helpdesk.asp?items=" & Session("items") &"&docnum=" & rsUpdateEntry("docnum") &"&RS=" & rsUpdateEntry("RazonSocial") &"'> Re-enviar Email </a> </td> " )
          end if



            if  rsUpdateEntry("Tipo") ="remision" then

                select case rsUpdateEntry("facturar")
                case "Y"
                     if  trim(rsUpdateEntry("sentdate") ) <>"" then
                          response.write("<td class="& vTD &"><font class='fonttiny'><B>AUTORIZADO: " & rsUpdateEntry("RequestDate")  & "</font></B></td> " )
                     else
                          response.write("<td class="& vTD &"><font class='fonttiny'><B>AUTORIZADO, no se ha enviado email</B></font> </td> " )
                     end if

                case "N"
                      response.write("<td class="& vTD &"><font class='fonttiny'>NO AUTORIZADO</td></font>" )
                case "H"
                      response.write("<td class="& vTD &"><font class='fonttiny' style='color:red'>HOLD </font></B></td> " )

                end select

                 response.write("<td class="& vTD &"><a href='helpdesk.asp?items=" & Session("items") &"&docnum=" & rsUpdateEntry("docnum") &"&RS=" & rsUpdateEntry("RazonSocial") &"'> AUTORIZAR </a> </td> " )
            end if
          


              if  rsUpdateEntry("Tipo") ="factura" then

                  select case rsUpdateEntry("facturar")
                  case "Y"
                       if  trim(rsUpdateEntry("sentdate") ) <>"" then
                            response.write("<td class="& vTD &"><font class='fonttiny'><B>APROBADA, email OK en: " & rsUpdateEntry("sentDate")  & "</td> " )
                       else
                            response.write("<td class="& vTD &"><B>APROBADA, eMail PDTE</B> </td> " )
                       end if
                       response.write("<td>-</td> " )

                  case "N"
                        response.write("<td class="& vTD &">NO APROBADA </td> " )
                        response.write("<td class="& vTD &"><a href='helpdesk.asp?items=" & Session("items") &"&docnum=" & rsUpdateEntry("docnum") &"&RS=" & rsUpdateEntry("RazonSocial") &"'> Autorizar notificaci&oacute;n interna</a> </td> " )

                  end select

                  end if

     end if




      if Session("Modulo")="compras" then     

              if trim(rsUpdateEntry("sentdate"))<>"" then
                  if mid(trim(rsUpdateEntry("control")),1,6)="UPDATE" then
                     response.write("<td class="& vTD &"><B>Se actualiz&oacute; OC al Proveedor en: " & rsUpdateEntry("sentDate")  & "</B></td> " )
                     response.write("<td class="& vTD &"><a href='helpdesk.asp?items=OC&docnum=" & rsUpdateEntry("docnum") &"&RS=" & rsUpdateEntry("RazonSocial") &"'> Re-env&iacute;o de OC a Proveedor por Actualizaci&oacute;n</a> </td> " )
                  else
                     response.write("<td class="& vTD &">Ya se ha enviado la OC al Proveedor en: " & rsUpdateEntry("sentDate")  & "</td> " )
                     response.write("<td class="& vTD &"><a href='helpdesk.asp?items=OC&docnum=" & rsUpdateEntry("docnum") &"&RS=" & rsUpdateEntry("RazonSocial") &"'> Re-env&iacute;o de OC a Proveedor por Actualizaci&oacute;n</a> </td> " )
                  end if
              else

                   response.write("<td class="& vTD &">Ya se cre&oacute; la OC, a&uacute;n no se ha env&iacute;ado a Proveedor</td> " )
                   response.write("<td class="& vTD &">&nbsp;</td> " )

              end if

      end if





      if rsUpdateEntry("InterEmpresas") ="Y" then
          response.write("<td style='color: red'><B>SI</B></td> " )
      else
          response.write("<td class="& vTD &" style='background-color: white!important'>no</td> " )
      end if



      if  rsUpdateEntry("Tipo") ="pedido" and rsUpdateEntry("modulo") ="compras"  then 

            if rsUpdateEntry("RazonSocial") ="MME" then
                vEmpresa="MEIDE"
            else
                vEmpresa=rsUpdateEntry("RazonSocial")
            end if

             strSQL="SELECT numAtCard,U_PROYECTO from  PRODUCTIVA_"& vempresa &".dbo.OPOR  where Docnum="& rsUpdateEntry("DocNum")
             'response.write strSQL  
             rsUpdateEntry2.Open strSQL, adoCon4
            

             vString=trim(rsUpdateEntry2("numAtCard"))
             vPos=inStr(vString,"P")             
             response.write("<td class=" & vTD &">" )

             while vPos>0
                 flag=0
                 vPos=inStr(vString,"PD")   
                 if vPos>0 then   %>
                     <input type="button" value="PD <%=mid(vString,vPos+2,4)%>" onclick="sendReq('/modules/ShowDetails.asp','rs,PD,control','<%=rsUpdateEntry("RazonSocial")%>,<%=mid(vString,vPos+2,4)%>,2','BoxRoundedDetail2')"> <%                                                         
                     vstring=mid(vString,vPos+6,300)
                     flag=1
                 end if
                 if vPos=0 then
                     vPos=inStr(vString,"P")
                 end if
                 if vPos>0 and flag=0 then
                     %>
                     <input type="button" value="P <%=mid(vString,vPos+1,4)%>" onclick="sendReq('/modules/ShowDetails.asp','rs,P,control','<%=rsUpdateEntry("RazonSocial")%>,<%=mid(vString,vPos+1,4)%>,2','BoxRoundedDetail2')">   <%                                             
                     vstring=mid(vString,vPos+5,300)                     
                 end if
                 vPos=inStr(vString,"P")
             wend
            
             response.write("</td>")
             response.write("<td class="& vTD &">" & rsUpdateEntry2("U_PROYECTO")  &"</td>" )
             rsUpdateEntry2.close

       end if



      if  rsUpdateEntry("Tipo") ="pedido" and rsUpdateEntry("modulo") ="ventas"  then  

                if rsUpdateEntry("RazonSocial")="MME" then
                    vEmpresa="MEIDE"
                else
                    vEmpresa=rsUpdateEntry("RazonSocial")
                end if


             strSQL="SELECT b.SlpName,a.U_COMINTERNO,a.U_SEGUIMIENTO,a.U_PROYECTO,a.U_eMailEmbarques from  PRODUCTIVA_"& vempresa &".dbo.ORDR a inner join PRODUCTIVA_"& vempresa &".dbo.OSLP b on a.slpcode=b.slpcode where a.Docnum="&rsUpdateEntry("DocNum")
             'response.write strSQL  
             rsUpdateEntry2.Open strSQL, adoCon4
            
             response.write("<td class="& vTD &">")  
             response.write("<a href='http://192.168.0.206/repositorio/" & rsUpdateEntry("RazonSocial") & "/" & rsUpdateEntry("Ruta")  & "/" & rsUpdateEntry("DocNum") &"' target='repositorio'>abrir repositorio</a></td>")

             response.write( "<td class="& vTD &"><B>" & rsUpdateEntry2("U_PROYECTO")  &" </B></td> ")



             response.write("<tr><td colspan=3 class='td-r'><a style='font-color=red' href='helpdesk.asp?action=2&RS="& rsUpdateEntry("RazonSocial") &"&docnum="& rsUpdateEntry("DocNum") & "'>[Seguimiento al suministro:] <br>-coment interno-</A></td>")

             response.write("<td colspan=4 class="& vTD &"><font class='fonttiny'>" &rsUpdateEntry2("U_COMINTERNO") & "<a style='font-color=red' href='helpdesk.asp?action=2&RS="& rsUpdateEntry("RazonSocial") &"&docnum="& rsUpdateEntry("DocNum") & "'> <img src='/img/Edicion.jpg' border=0 alt='editar coment interno' title='editar coment interno' height='16px' width='16px'></font></a></td>")

             response.write("<td class='td-r' colspan=2>[asesor:] </td><td class='td-l' colspan=2>"& rsUpdateEntry2("SlpName")  &"</td></tr>")

             response.write("</tr><tr><td colspan=3 class='td-r'>[Instrucciones Ventas:]<br>-seguimiento-</td><td colspan=4 class="& vTD &" style='color: red;'><font class='fonttiny'>"& rsUpdateEntry2("U_SEGUIMIENTO")  &"</font></td>")

             response.write("<td class='td-r fontmed' colspan=2><a style='font-color=red' href='helpdesk.asp?action=2&RS="& rsUpdateEntry("RazonSocial") &"&docnum="& rsUpdateEntry("DocNum") & "'>[emails:]</a></td><td class='td-l fontmed' colspan=2> "& rsUpdateEntry2("U_eMailEmbarques") &"</td>       </tr><tr>")

               separador   'mike'
           
             rsUpdateEntry2.close
      end if



      if  rsUpdateEntry("Tipo") ="remision" then     

                if rsUpdateEntry("RazonSocial")="MME" then
                    vEmpresa="MEIDE"
                else
                    vEmpresa=rsUpdateEntry("RazonSocial")
                end if


             strSQL="SELECT U_COMINTERNO,U_SEGUIMIENTO,U_PROYECTO from  PRODUCTIVA_"& vempresa &".dbo.ODLN  where Docnum="&rsUpdateEntry("DocNum")
             'response.write strSQL  
             rsUpdateEntry2.Open strSQL, adoCon4
            

            response.write("<td class="& vTD &" style='background-color: white!important'><font color='red'> "  & rsUpdateEntry("facturar")  & "</font> </td> " )
            response.write("<td class="& vTD &" style='background-color: white!important'> "  & rsUpdateEntry("pedido")  & " </td> " )
            
             
             response.write( "<td class="& vTD &"><B>" & rsUpdateEntry2("U_PROYECTO")  &" </B></td> ")

             response.write("<tr><td colspan=3 class='td-r'>Comentario interno:</td><td colspan=6 class="& vTD &"><font class='fonttiny'>") 



             strSQL="SELECT a.*,b.portador from envios a inner join cat_portadores b on a.id_portador=b.id_portador where a.Docnum="& rsUpdateEntry("DocNum") &" and a.tipo='remision' "
             'response.write strSQL
             rsUpdateEntry3.Open strSQL, adoCon4

             if rsUpdateEntry3.EOF then                   
                   vString=""
             else
                   vString=rsUpdateEntry3("portador")  & " rastreo: "& rsUpdateEntry3("rastreo") &" Fecha: "& rsUpdateEntry3("DocDate") & " comentarios: "& rsUpdateEntry3("comentarios")
             end if


             response.write( rsUpdateEntry2("U_COMINTERNO") & "</td><td colspan=3 class='td-l' rowspan=2><img src='/img/envios.png' border=0 alt='' title='' width='30px' height='30px'>&nbsp;&nbsp;<B><font class='fonttiny'>"& vString &"</font></B></td></tr>")
             response.write("<tr><td colspan=3 class='td-r'>Seguimiento:</td><td colspan=6 class="& vTD &" style='color: red;'><font class='fonttiny'>" & rsUpdateEntry2("U_SEGUIMIENTO") &" </font></td> </tr>")

             if rsUpdateEntry("facturar")="H" then
                 response.write("<tr><td colspan=3 class='td-r' style='color:white;background-color:gray'>HOLD por:</td><td class='td-l' colspan=9 style='color:white;background-color:gray'>" & rsUpdateEntry("RazonHold") &"</td></tr>")
              end if

            response.write("<tr><td colspan=12 class='separador'>&nbsp;</td></tr>")

             rsUpdateEntry2.close
             rsUpdateEntry3.close

           

      end if


      if  rsUpdateEntry("Tipo") ="factura" then

            response.write("<td class="& vTD &" style='background-color: white!important'>" & rsUpdateEntry("remision")  & "</td>" )

            response.write("<td class="& vTD &" style='background-color: white!important'><a href='ShowContent.asp?contenido=68&action=1&fRS="&rsUpdateEntry("RazonSocial")&"&fpedido="&rsUpdateEntry("pedido")&"' target='pedido'>" & rsUpdateEntry("pedido")  & "</a></td>" )

                 if rsUpdateEntry("RazonSocial")="MME" then
                    vEmpresa="MEIDE"
                else
                    vEmpresa=rsUpdateEntry("RazonSocial")
                end if


             strSQL="SELECT U_COMINTERNO,U_SEGUIMIENTO,U_PROYECTO,U_CFDi_MetodoPago,CONVERT(VARCHAR,CONVERT(MONEY,DocTotalSy-VatSumSy),1)  as subtotal from  PRODUCTIVA_"& vempresa &".dbo.OINV  where Docnum="&rsUpdateEntry("DocNum")
             'response.write strSQL  
             rsUpdateEntry2.Open strSQL, adoCon4
            
             if rsUpdateEntry("remision")<>"" then
                 strSQL="SELECT a.*,b.portador from envios a inner join cat_portadores b on a.id_portador=b.id_portador where a.RazonSocial='" & rsUpdateEntry("RazonSocial") &"' and a.Docnum="& rsUpdateEntry("remision")  &" and a.tipo='remision' "
                 'response.write (strSQL&"<br>")
                 rsUpdateEntry3.Open strSQL, adoCon4

                 if rsUpdateEntry3.EOF then                   
                       vString=""
                 else
                       vString=rsUpdateEntry3("portador")  & " rastreo: "& rsUpdateEntry3("rastreo") &" Fecha: "& rsUpdateEntry3("DocDate") & " comentarios: "& rsUpdateEntry3("comentarios") 
                 end if

                 rsUpdateEntry3.close
                 
             end if

             response.write( "<td class="& vTD &"><B>" & rsUpdateEntry2("U_PROYECTO")  &" </B></td>")
              if  rsUpdateEntry2("U_CFDi_MetodoPago")="PPD" then
                 response.write("<td class='td-c'>CREDITO  <BR> "& rsUpdateEntry2("subtotal") &"</td></tr>")
              else 
                 response.write("<td class='td-c' style='color:white; background-color:red;'>CONTADO!!  <BR> "& rsUpdateEntry2("subtotal")  &"</td></tr>")
              end if

             '---cambio de renglon'    
             response.write("<tr><td colspan=2 class='td-r'>Comentario interno:</td><td colspan=6 class="& vTD &"><font class='fonttiny'>")
             
             response.write( rsUpdateEntry2("U_COMINTERNO") & "</td><td colspan=4 class='td-l' rowspan=2><img src='/img/envios.png' border=0 alt='' title='' width='30px' height='30px'>&nbsp;&nbsp;<B><font class='fonttiny'>"& vString &"</font></B></font></td></tr>")
             response.write("<tr><td colspan=2 class='td-r'>Seguimiento:</td><td colspan=6 class="& vTD &" style='color: red;'><font class='fonttiny'>" & rsUpdateEntry2("U_SEGUIMIENTO") &" </font></td> </tr><tr><td colspan=12 class='separador'>&nbsp;</td></tr>")




             rsUpdateEntry2.close
             

             
      end if

      response.write("</tr>") 
      rsUpdateEntry.MoveNext

      if not rsUpdateEntry.EOF then
           if vDia<>day(rsUpdateEntry("docdate")) then
               'separador
               vDia=day(rsUpdateEntry("docdate"))
           end if
      end if


      i=i+1
  wend
  






if Session("Modulo")="ventas" and Session("items")="pedido" and request("DocNum")<>"" then 
      rsUpdateEntry.MoveFirst
      vmsg="Se actualizo Pedido No. "& trim(rsUpdateEntry("RazonSocial")) &"-" & trim(request("DocNum")) 

     %>

     <form id=update action="up-items.asp" method="post">
         <input type=hidden id=control  name=control value=1>  <!-- Se actualizara envio de Pedido de Venta -->
         <input type=hidden id=fRS  name=fRS value="<%=rsUpdateEntry("RazonSocial") %>">
         <input type=hidden id=fItems  name=fItems value="<%=rsUpdateEntry("tipo") %>">
         <input type=hidden id=fDocNum  name=fDocNum value="<%=rsUpdateEntry("DocNum") %>">
         <input type=hidden id=msg  name=msg value="<%=vmsg%>">
         <tr><td colspan=4 class=subtitulo style='text-align: right; padding-right: 10px'>Razon por el re-envio email Helpdesk: </td> <td colspan=2>
                 <input type="text" id="fnotas" name="fnotas" value="Actualizacion Partidas: " style="width: 450px">       </td> <td> <input type="submit" value="enviar">
         </tr>
     </form>
<%       
end if





if Session("Modulo")="ventas" and Session("items")="remision" and request("DocNum")<>"" then 
      rsUpdateEntry.MoveFirst
      if request("fautoriza")="Y" then
           vmsg="Se solicito a facturar Remision No. "& trim(rsUpdateEntry("RazonSocial")) &"-" & trim(request("DocNum")) 
      end if
      if request("fautoriza")="H" then
           vmsg="Se puso en HOLD Remision No. "& trim(rsUpdateEntry("RazonSocial")) &"-" & trim(request("DocNum")) 
      end if

      'response.write rsUpdateEntry("facturar")
     %>

     <form id=update action="up-items.asp" method="post">
         <input type=hidden id=control  name=control value=2>  <!-- Se autorizara facturacion -->
         <input type=hidden id=fRS  name=fRS value="<%=rsUpdateEntry("RazonSocial") %>">
         <input type=hidden id=fItems  name=fItems value="<%=rsUpdateEntry("tipo") %>">
         <input type=hidden id=fDocNum  name=fDocNum value="<%=rsUpdateEntry("DocNum") %>">
         <input type=hidden id=msg  name=msg value="<%=vmsg%>">  <%


         strSQL="select b.* from ODLN a inner join OCRD b on a.cardcode=b.CardCode where DocNum="&request("DocNum")
                'response.write strSQL
         select case request("RS")
                   case "DMX"   rsUpdateEntry5.Open strSQL, adoCon2
                   case "DFW"   rsUpdateEntry5.Open strSQL, adoCon3
                   case "MME"   rsUpdateEntry5.Open strSQL, adoCon5
         end select

         if trim(rsUpdateEntry5("LicTradNum"))="XEXX010101000" then                  
                 response.write("<tr><td colspan=12 class='td-c strong'><font color=red>FACTURACION PARA COMERCIO EXTERIOR </font></td></tr>") 
                  response.write("<tr><td colspan=12 class='td-c strong'><font color=red>Indique IncoTerm (pregunte al asesor de ventas):</font> <select name='fIncoTerm'>")

                  strSQL="SELECT * FROM [@INCOTERMS]"
                  rsUpdateEntry6.Open strSQL, adoCon2

                  while not rsUpdateEntry6.EOF
                      response.write("<option value='"&rsUpdateEntry6("code")&"'>"&rsUpdateEntry6("code")&"</option>")
                      rsUpdateEntry6.MoveNext
                  wend
                  rsUpdateEntry6.close
                  response.write("</select><a href='/docs/incoterms_2020.pdf'> descargue gu&iacute;a INCOTERMS</a></td></tr>") 
                  response.write("<tr><td colspan=12 class='td-c strong'>&nbsp;</td></tr>")

         end if
         rsUpdateEntry5.close

      %>


         <tr><td colspan=3 class=subtitulo style='text-align: right; padding-right: 10px'> Solicitar facturar: </td> <td colspan=2> <%
              


                if trim(rsUpdateEntry("CardCode"))="C000001" then
                   response.write("<input type='text' id='fnotas' name='fnotas' value='notas: Para efectos de facturacion, INTEREMPRESAS ' style='width: 450px'></td>")
                else
                   response.write("<input type='text' id='fnotas' name='fnotas' value='notas: ' style='width: 450px'></td>")
                end if
             
              response.write("<td colspan=2 tyle='text-align: center'> <select name='fautoriza' id='fautoriza'> ") 

              select case trim(rsUpdateEntry("facturar"))
              case "N"
                  response.write("<option value='Y'>Autorizar Facturacion </option> ")   
                  response.write("<option value='H'>Poner en HOLD</option>")      
              case "Y"
                  response.write("<option value='N'>Revertir Solicitud</option>")
                  response.write("<option value='H'>Poner en HOLD</option>")
              case "H"
                   response.write("<option value='Y'>Autorizar Facturacion </option> ")                   
              end select
              response.write("</select></td><td colspan=5>&nbsp;</td><tr>")

              if trim(rsUpdateEntry("facturar"))="N" or trim(rsUpdateEntry("facturar"))="Y" then
                 response.write("<tr><td colspan=6 class='td-l strong fonttiny'>Notas para remision en hold: <input type='text' name='fhold' class='anchox3'></td>")
                 response.write("<td class='td-l strong fonttiny'>email asesor de ventas (en caso de Hold)</td>")
                 response.write("<td class='td-l strong fonttiny' colspan=4><input type='text' name='fSlpemail' value='' style='width: 280px'></td>")
              else
                 response.write("<tr><td colspan=10 class='td-l strong fonttiny'>&nbsp;</td>")
              end if

              response.write("<td class='td-c'><input type='submit' value='actualizar'>")


               
               %>

               
         </tr>
     </form>
<%       
end if





if Session("Modulo")="ventas" and Session("items")="factura" and request("DocNum")<>"" then 
      rsUpdateEntry.MoveFirst
      vmsg="Se autoriz&oacute; a VENTAS enviar documentos de: <B>"&rsUpdateEntry("RazonSocial") &"</B> [ PED-" & rsUpdateEntry("pedido") &" ] [ REM-"& rsUpdateEntry("remision") &" ] [ FACT-" & request("DocNum") &" ]"
     'response.write rsUpdateEntry("facturar")
     %>

     <form id=update action="up-items.asp" method="post">
         <input type=hidden id=control  name=control value=3>  <!-- Se autorizara facturacion -->
         <input type=hidden id=fRS  name=fRS value="<%=rsUpdateEntry("RazonSocial") %>">
         <input type=hidden id=fItems  name=fItems value="<%=rsUpdateEntry("tipo") %>">
         <input type=hidden id=fDocNum  name=fDocNum value="<%=rsUpdateEntry("DocNum") %>">
         <input type=hidden id=msg  name=msg value="<%=vmsg%>">
         <tr><td colspan=4 class=subtitulo style='text-align: right; padding-right: 10px'> Autorizar env&iacute;o de docs a VENTAS v&iacute;a Helpdesk: </td> <td colspan=2>
                 <input type="text" id="fnotas" name="fnotas" value="INSTRUCCIONES ADMINISTRACION: ninguna" style="width: 450px">                        
               </td>  <%

               if trim(rsUpdateEntry("facturar"))="Y" then
                   response.write("<td colspan=2 tyle='text-align: center'> <select name='fautoriza' id='fautoriza'> <option value='N'>Revertir Solicitud</option></select>")
                   response.write("</td><td><input type='submit' value='revertir'>")
               else
                    response.write("<td colspan=2 tyle='text-align: center'> <select name='fautoriza' id='fautoriza'> <option value='Y'>Autorizar Solicitud</option> </select>")
                    response.write("</td><td><input type='submit' value='autorizar'>")
               end if
               %>

               
         </tr>
     </form>
<%       
end if






if Session("Modulo")="compras" and Session("items")="pedido" and request("DocNum")<>"" then 
      rsUpdateEntry.MoveFirst
      vmsg="Se actualizo re-envio via Helpdesk de OC "& trim(rsUpdateEntry("RazonSocial")) &"-" & trim(request("DocNum")) 

     %>

     <form id=update action="up-items.asp" method="post">
         <input type=hidden id=control  name=control value=4>  <!-- Se actualizara envio de OC -->
         <input type=hidden id=fRS  name=fRS value="<%=rsUpdateEntry("RazonSocial") %>">
         <input type=hidden id=fItems  name=fItems value="<%=rsUpdateEntry("tipo") %>">
         <input type=hidden id=fDocNum  name=fDocNum value="<%=rsUpdateEntry("DocNum") %>">
         <input type=hidden id=msg  name=msg value="<%=vmsg%>">
         <tr><td colspan=4 class=subtitulo style='text-align: right; padding-right: 10px'>Razon por el re-envio email Helpdesk: </td> <td colspan=2>
                 <input type="text" id="fnotas" name="fnotas" value="PD. We are updating this PO" style="width: 450px">       </td> <td> <input type="submit" value="enviar">
         </tr>
     </form>
<%       
end if
rsUpdateEntry.close

end sub





%>

</table><P>&nbsp;</P><P>&nbsp;</P><P>&nbsp;</P><P>&nbsp;</P><P>&nbsp;</P><P>&nbsp;</P><P>&nbsp;</P><P>&nbsp;</P><P>&nbsp;</P><P>&nbsp;</P><P>&nbsp;</P><P>&nbsp;</P><P>&nbsp;</P><P>&nbsp;</P><P>&nbsp;</P><P>&nbsp;</P><P>&nbsp;</P><P>&nbsp;</P><P>&nbsp;</P><P>&nbsp;</P><P>&nbsp;</P><P>&nbsp;</P><P>&nbsp;</P><P>&nbsp;</P><P>&nbsp;</P>

<div id="BoxRoundedDetail">&nbsp;</div>
<div id="BoxRoundedDetail2">&nbsp;</div>

</div>














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

