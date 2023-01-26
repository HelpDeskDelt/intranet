<%
sub contenido_selector


			contenido="0"
			if request("contenido")<>"" then
			    contenido=trim(request("contenido"))
			end if

			if request("fcontenido")<>"" then
			    contenido=request("fcontenido")
			end if


			if request("modulo")<>"" then
			     Session("modulo")=request("modulo")
			end if


			if request("accion")<>"" then
			     Session("accion")=request("accion")
			end if

			if Session("accion")="4" then  'cerrar sesion'
			      Session("session_modulo")=""
			end if

           

			'START'
			select case contenido
			case "0"      response.redirect("/home.asp?msg="&request("msg") )
			case "1"      cumple
			case "2"      Dotitulo("COTIZACIONES vs FACTURADO")                  
			              cotizaciones      
			case "3"      Dotitulo("VIDEO: Capacitacion CRM en SAP")  
			              CRM    
			case "4"      proyectos     
			case "5"      Dotitulo("TERRITORIOS POR ASESOR DE VENTA")   
			              territorios   
			case "6"      ShowAvanceVentas 
			case "7"      Dotitulo("REPRODUCIENDO UN VIDEO")       
			              Session("videos")="MindFullNess.mp4"
			              PlayVideo
			case "8"      CrearRuta     
			case "9"      if request("accion")<>"" then
			                    Session("accion")=request("accion")
			              end if			             
			              if  Session("session_modulo") ="humres" then
			                   select case Session("accion")
			                   case 1 ShowAddEmployer
			                   case 2 ShowDelEmployer
			                   case 3 ShowEditEmployer
			                   end select
			              else
			                   ShowLogin
			              end if

			case "10"     CheckLogin
			case "11"     
			               select case request("control")
			               case 1 ShowViaticos
			               case 2 CheckViaticos
			               case 3 SolicitarViaticos
			               end select             
			case "12"      ControlViaticos
			case "13"      if  Session("session_modulo") ="viaticos" then
			                   AutorizaViatico
			               else
			                   ShowLogin
			               end if
			                
			case "14"      UpViatico

			case "15"      select case request("control")
			               case 1 ShowReembolso
			               case 2 CheckReembolso
			               case 3 ConsolidarReembolso
			               case 4 SolicitarReembolso
			               end select  

			case "16"      if  Session("session_modulo") ="reembolsos" then
			                   AutorizaReembolso
			               else
			                   ShowLogin
			               end if    

			case "17"      UpReembolso
			case "18"      DetalleReembolso
			case "19"      ShowLogin
			case "20"      ControlReembolsos
			case "21"      ComprobarViatico
			case "22"      ComprobarViaticoUP
			case "23"      ShowCargas
			case "24"      CrearRutaCargas
			case "25"      XmlCargaBrowser
			case "26"      XmlCargarProcesar
			case "29"      DetalleViatico
			case "30"      DetalleXML
			case "31"      DetalleProyecto

			case "32"     Session("modulo")="H_Incidencias"
			              CambioModulo
			              if  Session("session_modulo") ="H_Incidencias" then
			                   H_Incidencias
			              else                   
			                   ShowLogin
			              end if
			case "33"     UP_incidencias
			case "34"     edit_incidencias
			case "35"     select case request("opt")
			               case 1 FormOmision
			               case 2 OmisionRequest
			               case 3 OmisionJustifique
			               case 4 UPomision
			               end select  
			case "36"     Session("modulo")="ControlEnvios"
			              CambioModulo
			              if  Session("session_modulo") ="ControlEnvios" then
			                   ControlEnvios
			              else                   
			                   ShowLogin
			              end if
			case "37"     Session("modulo")="H_Incidencias" 
			              if  Session("session_modulo") ="H_Incidencias" then
			                   DiaInhabil
			              else                   
			                   ShowLogin
			              end if
			case "38"     DiaInhabilM
			case "39"     CatHorarios
			case "40"     Session("modulo")="H_Incidencias" 
			              CambioModulo
			              if  Session("session_modulo") ="H_Incidencias" then
			                   select case request("accion")
			                   case "0" Empleados
			                   case "1" ShowAddEmployer
			                   case "2" ShowDelEmployer
			                   case "3" ShowEditEmployer
			                   end select                   
			              else                   
			                   ShowLogin
			              end if

			case "41"     Session("modulo")="checada" 
			              CambioModulo
			              if  Session("session_modulo") ="checada" then
			                   checada
			              else                   
			                   ShowLogin
			              end if    

			 case "42"    Session("modulo")="ModuloADD" 
			              CambioModulo
			              if  Session("session_modulo") ="ModuloADD" then
			                   ModuloADD
			              else                   
			                   ShowLogin
			              end if 

			case "43"     PermisoADD 

			case "44"     Session("modulo")="ChecarIntranet" 
			              CambioModulo
			              if  Session("session_modulo") ="ChecarIntranet" then
			                   ChecarIntranet
			              else                   
			                   ShowLogin
			              end if   

			case "45"     if  Session("session_modulo") ="H_Incidencias" then
			                   vacaciones
			              else                   
			                   response.redirect("/home.asp?msg=vuelva a iniciar sesion")
			              end if  

			 case "46"    Session("modulo")="SolicVacaciones" 
			              CambioModulo
			              if  Session("session_modulo") ="SolicVacaciones" then
			                   SolicVacaciones
			              else
			                   ShowLogin
			              end if

			 case "47"    
			              if  Session("session_modulo") ="SolicVacaciones" then
			                   SolicVacacionesForm
			              else                   
			                   response.redirect("/home.asp?msg=vuelva a iniciar sesion")
			              end if  

			 case "48"    
			              if  Session("session_modulo") ="SolicVacaciones" then
			                   RegistrarVacaciones
			              else                   
			                   response.redirect("/home.asp?msg=vuelva a iniciar sesion")
			              end if  
			 case "49"    
			              if  Session("session_modulo") ="SolicVacaciones" then
			                   CancelarVacaciones
			              else                   
			                   response.redirect("/home.asp?msg=vuelva a iniciar sesion")
			              end if
			case "50"     
			              if  Session("session_modulo") ="H_Incidencias" then
			                   EjercerVacaciones
			              else                   
			                   response.redirect("/home.asp?msg=vuelva a iniciar sesion")
			              end if  
			case "51"     
			              if  Session("session_modulo") ="H_Incidencias" then
			                   EjercerVacacionesUP
			              else                   
			                   response.redirect("/home.asp?msg=vuelva a iniciar sesion")
			              end if              

			case "52"     Session("modulo")="cobranza" 
			              CambioModulo
			              if  Session("session_modulo") ="cobranza" then
			                   FormEnviarCartera
			              else                   
			                   ShowLogin
			              end if   

			case "53"     
			              if  Session("session_modulo") ="cobranza" then
			                   ExecuteCartera
			              else                   
			                   response.redirect("/home.asp?msg=vuelva a iniciar sesion")
			              end if      

			case "54"     Session("modulo")="Alarmas"  'configuracion del servicio de alarmas ADT alarmex etc'
			              CambioModulo
			              if  Session("session_modulo") ="Alarmas" then
			                   ShowAlarmas
			              else                   
			                   ShowLogin
			              end if   

			case "55"     Session("modulo")="PuntoEquilibrio" 
			              CambioModulo
			              if  Session("session_modulo") ="PuntoEquilibrio" then
			                   FormPointEquilibrio
			              else                   
			                   ShowLogin
			              end if 

			case "56"     if  Session("session_modulo") ="PuntoEquilibrio" then
			                   ExecutePE
			              else                   
			                   response.redirect("/home.asp?msg=vuelva a iniciar sesion")
			              end if    

			case "57"     Session("modulo")="TipoCambio"  'configuracion del servicio de alarmas ADT alarmex etc'
			              CambioModulo
			              if  Session("session_modulo") ="TipoCambio" then
			                   ExecuteTipoCambio
			              else                   
			                   ShowLogin
			              end if   

			              
			case "58"     Session("modulo")="SolicitarCumple"  
			              CambioModulo 
			              if  Session("session_modulo") ="SolicitarCumple" then
			                   SolicitarCumple
			              else                   
			                   ShowLogin
			              end if   


			case "59"     CostosStock 
			case "60"     Mantenimiento 
			case "61"     if request("ID")<>"" then
			                  Session("ID")=request("ID") 
			              end if
			              Session("modulo")="ControlMantenimiento" 
			              CambioModulo
			              if  Session("session_modulo") ="ControlMantenimiento" then                   
			                   ControlMantenimiento
			              else    
			                   ShowLogin
			              end if   

			case "62"     Session("modulo")="TaskManager" 
			              CambioModulo
			              if  Session("session_modulo") ="TaskManager" then
			                   TaskManager
			              else                   
			                   ShowLogin
			              end if          

			case "63"     ShowRecordsCOVID
			case "64"     ShowStockTubos
			case "65"     interTransaction 
			case "66"     DellateOCTaller 
			case "67"     Session("modulo")="CorteInventario"
		                  CambioModulo
		                  if  Session("session_modulo") ="CorteInventario" then
		                     CorteInventario
		                  else
		                     ShowLogin
		                  end if 
			              
			case "68"     PedidoInventario 

			case "69"     Session("modulo")="Inventario" 
			              'CambioModulo
			              if  Session("session_modulo") ="Inventario" then
			                   Inventario
			              else                   
			                   ShowLogin
			              end if   

			case "70"     PartidasAbiertas1  'ventas'  

			case "71"     Session("modulo")="InformeEnvios"   
			              CambioModulo
			              if  Session("session_modulo") ="InformeEnvios" then
			                   InformeEnvios
			              else                   
			                   ShowLogin
			              end if        

			case "72"     Contactos
			            
			case "73"     pdfRepositorio

			case "74"     



			case "75"     ItemsCotizados
			case "76"     ShowContacts	
			
			case "77"     Session("mode")=""
			              Session("modulo")="SolicitudCompra"  
			              CambioModulo 
			              if  Session("session_modulo") ="SolicitudCompra" then
			                   SolicitudCompra
			              else                   
			                   ShowLogin
			              end if   
			case "78"     
			              if  Session("session_id") ="MCABANILLAS" or Session("session_id") ="AJAUREGUI" or Session("session_id") ="HESCOBAR" or Session("session_id") ="AROGRIGUEZ" or Session("session_id") ="AGANDAR"  then
			                   AprobarSolicitudCompra
			              else                   
			                   ShowLogin
			              end if   

			case "79"     Session("modulo")="ShowSolicitudCompras"   
			              CambioModulo
			              Session("mode")=2
			              if  Session("session_modulo") ="ShowSolicitudCompras" then
			                   ShowSolicitudCompras 
			              else                   
			                   ShowLogin
			              end if 

            case "80"     Session("modulo")="ABCIngresos"   
                          CambioModulo			              
			              if  Session("session_modulo") ="ABCIngresos" then
			                   ABCIngresos 
			              else                   
			                   ShowLogin
			              end if 

			case "81"     EstatusCliente
			case "82"     Coti_NoPartes
			case "83"     InformeClientesNuevos

			case "84"     InformeCotisNoClosed

			case "85"     Session("modulo")="AnalisisUtilidadPedido"  			             
			              if  Session("session_modulo") ="AnalisisUtilidadPedido" or Session("session_modulo")="AnalisisUtilidadHistorico" then
			                   AnalisisUtilidadPedido 
			              else                   
			                   ShowLogin
			              end if 


     case "86"     MinimosDeStock
     case "87"     Session("modulo")="Flotilla"
                   CambioModulo
                   if  Session("session_modulo") ="Flotilla" then
                         Flotilla
                   else
                         ShowLogin
                   end if

     case "88"     Session("modulo")="ControlViajes"
                   CambioModulo
                   if  Session("session_modulo") ="ControlViajes" then
                         ControlViajes
                   else
                         ShowLogin
                   end if

      case "89"    Check_unidad 			           
      case "90"    Session("modulo")="RegistroEmbarques"
                   CambioModulo
                   if  Session("session_modulo") ="RegistroEmbarques" then
                         RegistroEmbarques
                   else
                         ShowLogin
                   end if
                   
      case "91"    Session("modulo")="EnvioConsolidado"
                   CambioModulo
                   if  Session("session_modulo") ="EnvioConsolidado" then
                         EnvioConsolidado
                   else
                         ShowLogin
                   end if         

      case "92"    StockSinPeso 
      case "93"    CargarXMLEfectivale
      case "94"    DetalleXMLEfectivale

      case "95"    Session("modulo")="voucherGas"
                   CambioModulo
                   if  Session("session_modulo") ="voucherGas" then
                         voucherGas
                   else
                         ShowLogin
                   end if                    
      case "96"    Check_unidad_edit
      
      case "98"    notificaciones
      case "99"    anticipos
      case "100"   ViaticoChange
      case "101"   OpenLinesPurchasing
      case "102"   VentasAnioAsesor
      case "103"   Session("modulo")="DIOT"
                   CambioModulo
                   if  Session("session_modulo") ="DIOT" then
                         DIOT
                   else
                         ShowLogin
                   end if  
      case "104"   Session("modulo")="PAGOS"
                   CambioModulo
                   if  Session("session_modulo") ="PAGOS" then
                         PAGOS
                   else
                         ShowLogin
                   end if  
       case "105"  informeCombustibles
       case "106"  Resguardos
       case "107"  Session("modulo")="EliminaFechaEmbarque"
                   CambioModulo
                   if  Session("session_modulo") ="EliminaFechaEmbarque" then
                         EliminaFechaEmbarque
                   else
                         ShowLogin
                   end if  
       case "108"  Session("modulo")="AnalisisUtilidadHistorico"  
			              if  Session("session_modulo") ="AnalisisUtilidadPedido" or Session("session_modulo")="AnalisisUtilidadHistorico"  then
			                   AnalisisUtilidadHistorico 
			              else                   
			                   ShowLogin
			              end if 

       case "109" EstimarFlete
       case "110" Peajes
       case "111" VoucherGasOtros
       case "112" Codigos
       case "113" CotizacionesAbiertas
       case "114" ComportamientoPagos
       case "115" CostoInterE
       case "116" ShowXmlsCombustibles
       case "117" DirectorioClientes

       case "118" Session("modulo")="feeds"
                  CambioModulo
                  if  Session("session_modulo") ="feeds" then
                     feeds
                  else
                     ShowLogin
                  end if  
                  
       case "119" DetalleFactura
       case "120" ChangeStatusPO
       case "121" TranfStock
       case "122" ViaticoEdit
       case "123" EstatuSuministro
       case "124" RemisionDetalle
       case "125" EstatusProveedores
       case "126" Session("modulo")="PotterRoemer"
                  CambioModulo
                  if  Session("session_modulo") ="PotterRoemer" then
                     PotterRoemer
                  else
                     ShowLogin
                  end if  


       case "127" ItemsCotizadosHistoricos
	 case "128" Session("modulo")="ControlPagos"
                  CambioModulo
                  if  Session("session_modulo") ="ControlPagos" then
                     ControlPagos
                  else
                     ShowLogin
                  end if  

       case "129" SendEmailVendorDraft
       case "130" SendEmailVendor
       '======================================================================'
       case "131" Session("modulo")="SolicitarPagar"
                  CambioModulo
                  if  Session("session_modulo") ="SolicitarPagar" then
                       SolicitarPagar                      
                  else
                       ShowLogin
                  end if  
       case "132" ControlPagosFacturas
       case "133" ComplementosRecibidos
       case "134" ControlPagosEfectuados
       '======================================================================'
       case "135" IngresosMesSinEmail
       case "136" ProgramacionPagos

       
       case "139" FechasEmbarque
       case "140" CostoVentaTuberias
       case "199" CostoContableOnStock
	 case "200" IntroduccionStock
 	 end select

end sub




'==============================================================================================================================================='

sub IntroduccionStock  'contenido 200'
      Doalert
      titulo="INTRODUCCION AL STOCK<BR><font style='font-size: 12px'>1.- Cuarto de m&aacute;quinas</font>"
      DoTitulo(titulo)

   

      CreateTable(900)



      response.write("<tr><td class='td-c subtitulo' width='40%'>elemento</td><td class='td-c subtitulo' width='60%'>referencia</td></tr>")

      response.write("<tr><td class='td-c' colspan=2><img src='/images/Molino_Tuberia.png' border=0 alt='Molino de tuberia' title='Molino de tuberia'></tr>")

      response.write("<tr><td class='td-c'><img src='/images/Valvula_OSY.png' border=0></td><td class='td-c fontsmall'>NRS significa -husillo no ascendente- (Non-Rising Stem). OS&Y significa -husillo y horquilla externos- (Outside Stem and Yoke). En una v&aacute;lvula NRS, el husillo gira para abrir y cerrar la compuerta, pero no se mueve hacia arriba y abajo cuando gira</td></tr>")
      response.write("<tr><td class='td-c'><img src='/images/coples.png' border=0></td><td class='td-c'>&nbsp;</td></tr>")
      response.write("<tr><td class='td-c'><img src='/images/brida.png' border=0></td><td class='td-c'>&nbsp;</td></tr>")
      response.write("<tr><td class='td-c'><img src='/images/check_alivio.png' border=0></td><td class='td-c'>&nbsp;</td></tr>")
      response.write("<tr><td class='td-c'><img src='/images/compuerta.png' border=0></td><td class='td-c'>&nbsp;</td></tr>")

      response.write("<tr><td class='td-c'><img src='/images/check_succion.png' border=0></td><td class='td-c'>&nbsp;</td></tr>")
      response.write("<tr><td class='td-c'><img src='/images/alivio_bomba.png' border=0></td><td class='td-c'>&nbsp;</td></tr>")
      response.write("<tr><td class='td-c'><img src='/images/mariposa.png' border=0></td><td class='td-c'>&nbsp;</td></tr>")

      response.write("<tr><td class='td-c'><img src='/images/presion.png' border=0></td><td class='td-c'>&nbsp;</td></tr>")

      response.write("</table><P>&nbsp;</P><P>&nbsp;</P>")

end sub





sub cumple  'contenido 1'
      
       strSQL="select count(*) as calculo from Empleados where FechaEgreso is null"
       rsUpdateEntry.Open strSQL, adoCon

       titulo="CUMPLEA&Ntilde;OS COLABORADORES ("& rsUpdateEntry("calculo") &") EN "& year(date())
       DoTitulo(titulo)

       rsUpdateEntry.close

       if request("msg")<>"" then
           response.write("<P class=alert>"&  request("msg") &"</P>")
        end if

       strSQL="select a.*,b.Departamento,c.Puesto,d.sucursal from Empleados a inner join cat_departamento b on a.Clave_depto=b.Clave_depto "
       strSQL= strSQL & "inner join cat_puesto c on a.clave_puesto=c.clave_puesto inner join cat_sucursal d on a.clave_sucursal=d.clave_sucursal  where a.FechaEgreso is null order by month(a.FechaNacimiento),day(a.FechaNacimiento)"     
       'response.write strSQL  
  	   rsUpdateEntry.Open strSQL, adoCon
  	   
       CreateTable(1100)
       response.write("<tr>")
       response.write("<td class=subtitulo>ID</td>")
       response.write("<td class=subtitulo>COLABORADOR</td>")
       response.write("<td class=subtitulo>CUMPLEA&Ntilde;OS</td>")
       response.write("<td class=subtitulo>AREA</td>")
       response.write("<td class=subtitulo>SUCURSAL</td>")
       response.write("</tr>")
       
   
        while not rsUpdateEntry.EOF 
           response.write("<tr>")  %>
           <td>  <%=rsUpdateEntry("ID")%> </td> 
           <%
           
           if  month(date())=month(rsUpdateEntry("FechaNacimiento")) then
           	    response.write("<td class=cumple>"& rsUpdateEntry("nombres") &" "  &rsUpdateEntry("PATERNO") &" " &rsUpdateEntry("MATERNO")   &"</td>")
                response.write("<td class=cumple><B>"& day(rsUpdateEntry("FechaNacimiento")) &"-" & meses(month(rsUpdateEntry("FechaNacimiento"))) &"</B></td>")
                response.write("<td class=cumple>"& rsUpdateEntry("DEPARTAMENTO") &"</td>")
                response.write("<td class=cumple>"& rsUpdateEntry("Sucursal") &"</td>")  
           else
           	    response.write("<td>"& rsUpdateEntry("nombres") &" "  &rsUpdateEntry("PATERNO") &" " &rsUpdateEntry("MATERNO")   &"</td>")
         	      response.write("<td><B>"& day(rsUpdateEntry("FechaNacimiento")) &"-" & meses(month(rsUpdateEntry("FechaNacimiento"))) &"</B></td>")
         	      response.write("<td>"& rsUpdateEntry("DEPARTAMENTO") &"</td>")
                response.write("<td>"& rsUpdateEntry("Sucursal") &"</td>")  
           end if                    
        
           response.write("</tr>")
           rsUpdateEntry.MoveNext
       wend
       rsUpdateEntry.close     %>
       <tr><td colspan=6 class=separador>&nbsp;</td></tr> 
       
             <%
       response.write("</table>")    
       response.write("<P>&nbsp;</P>") 
       response.write("<P>&nbsp;</P>") 
       response.write("<P>&nbsp;</P>") 
       response.write("<P>&nbsp;</P>") 
       response.write("<P>&nbsp;</P>") 
end sub








sub ShowViaticos     'contenido 11  solicitud de viaticos

         response.write("<P>&nbsp;</p>")
         if request("msg")<>"" then         	    
         	    if int(trim(request("msg")))=1 then
                     vMsg="Ls remisiones no pertenecen a esa RAZON SOCIAL o alguna remision ingresada NO EXISTE, vuelva a intentar"
         	    else
                     vMsg="Formulario debe tener monto <100,000 motivo, cuenta bancaria y Ciudad destino. Corrija y vuelve a presionar continuar"
         	    end if
                response.write("<P class='td-c alert'>"& vMsg &"</p>")
         end if

        
         response.write("<center><table width='900px' border=1 cellpadding=2 align=center class=whiteBorders>")
         response.write("<tr><td colspan=7 class='td-c subject BoldBorders'  style='padding-top: 20px;padding-bottom: 20px;'>SOLICITUD DE VIATICOS </td></tr>")

         response.write("<form id='viaticos' method='post' action='ShowContent.asp'>")
         response.write("<input type='hidden' name='contenido' value='11'>")
         response.write("<input type='hidden' name='control' value='2'>")   'va a ir a CheckViaticos'

         response.write("<tr><td colspan=7 style='font-size: 8px;' class=whiteBorders>&nbsp;</td></tr>")
         response.write("<tr><td class='td-r bold whiteBorders'>FECHA:</td><td class='td-r BoldBorders'>  "& date() &"   </td><td colspan=4 class='td-r bold whiteBorders'>CANTIDAD:</td><td class='td-r UnderBorder azul'>$ <input class='td-r ancho' type='number'  name='vcantidad' id='vcantidad' value='"& request("vcantidad") &"'> </td></tr>")
         response.write("<tr><td colspan=7 style='font-size: 6px;' class=whiteBorders >&nbsp;</td></tr>")

         response.write("<tr><td class=whiteBorders>&nbsp;</td><td class='td-r bold whiteBorders'>EMPRESA DEL COLABORADOR:</td><td class='td-l UnderBorder' colspan=4><select class=anchox4  name='vempresa'><option value='Servicios Profesionales, Legales, Contables, Financieros, Recursos Humanos y Gestoria SR de CV'>Servicios Profesionales, Legales, Contables, Financieros, Recursos Humanos y Gestoria SR de CV</option><option value='DELTAFLOW DE MEXICO S DE RL DE CV' SELECTED>DELTAFLOW DE MEXICO S DE RL DE CV</option><option value='DELTAFLOW S DE RL DE CV'>DELTAFLOW S DE RL DE CV</option></select> </td><td class=whiteBorders>&nbsp;</td></tr>")

         response.write("<tr><td colspan=7 style='font-size: 6px;' class=whiteBorders>&nbsp;</td></tr>")
         response.write("<tr><td colspan=2 class='td-r bold whiteBorders'>NOMBRE DEL BENEFICIARIO:</td><td class='td-l UnderBorder' colspan=4 ><select class=anchox4  name='vbeneficiario'> ")
           strSQL="select * from Empleados where FechaEgreso is null order by ID desc "
           'response.write strSQL  
           rsUpdateEntry.Open strSQL, adoCon

           while not  rsUpdateEntry.EOF
              response.write("<option value='"& rsUpdateEntry("nombres") &" " & rsUpdateEntry("PATERNO") &" " & rsUpdateEntry("MATERNO") &"'>"& rsUpdateEntry("nombres") &" " & rsUpdateEntry("PATERNO") &" " & rsUpdateEntry("MATERNO") &"</option>")
              rsUpdateEntry.MoveNext
           wend
            rsUpdateEntry.close
            response.write("</select></td><td class=whiteBorders>&nbsp;</td></tr>")
         response.write("<tr><td colspan=7 style='font-size: 6px;' class=whiteBorders>&nbsp;</td></tr>")

         response.write("<tr><td colspan=2  class='td-r bold whiteBorders'>IMPORTE DE LETRA:</td><td class='td-l UnderBorder' colspan=4> &nbsp;</td><td class=whiteBorders>&nbsp;</td></tr>")
         response.write("<tr><td colspan=7 style='font-size: 10px;' class=whiteBorders>&nbsp;</td></tr>")

         response.write("<tr><td colspan=2 class=whiteBorders>&nbsp;</td><td class='td-l bold under whiteBorders'>BANCO</td><td colspan=4 class=whiteBorders>&nbsp;</td></tr>")

         response.write("<tr><td colspan=2 class='td-r bold whiteBorders'>BANCO (1):</td><td class='td-l UnderBorder'>BANCOMER</td><td class='td-r bold UnderBorder'>CUENTA:</td><td class='td-l BoldBorders azul'> <input class='td-r anchox1' type='text' name='vCuenta1' id='vCuenta1' value='"& request("vCuenta1") &"' ></td><td class='td-r bold whiteBorders'>MONTO:</td><td class='td-l UnderBorder'>&nbsp;</td></tr>")

         response.write("<tr><td colspan=2 class='td-r bold whiteBorders'>BANCO (2):</td><td class='td-l UnderBorder'>BANAMEX</td><td class='td-r bold UnderBorder'>CUENTA CLABE:</td><td class='td-l BoldBorders azul'> <input class='td-r anchox1' type='text' name='vCuenta2' id='vCuenta2' value='"& request("vCuenta2") &"' ></td><td class='td-r bold whiteBorders'>MONTO:</td><td class='td-l UnderBorder'>&nbsp;</td></tr>")

         response.write("<tr><td colspan=2 class='td-r bold whiteBorders'>BANCO (3):</td><td class='td-l UnderBorder'>&nbsp;</td><td class='td-r bold UnderBorder'>&nbsp;</td><td class='td-l whiteBorders'>&nbsp;</td><td class='td-r bold whiteBorders'>MONTO:</td><td class='td-l UnderBorder'>&nbsp;</td></tr>")
         response.write("<tr><td colspan=7 style='font-size: 4px;' class=whiteBorders>&nbsp;</td></tr>")

         response.write("<tr><td class='td-r bold whiteBorders'>N&Oacute;MINA:</td><td colspan=2 class='td-l bold UnderBorder'>DELTAFLOW DE MEXICO </td><td colspan=2 class=whiteBorders>&nbsp;</td><td class='td-r bold whiteBorders'>&nbsp;</td><td class='td-l whiteBorders'>&nbsp;</td></tr>")
         response.write("<tr><td colspan=7 style='font-size: 12px;' class=whiteBorders>&nbsp;</td></tr>")
         response.write("<tr><td class='td-r bold whiteBorders' Colspan=4># Total de remisiones a entregar:</td><td class='td-l bold'><select name='vremisiones'> ")

         for i=0 to 20 step 1
              if i=0 then
                  response.write("<option value="&i&" SELECTED>"&i&"</option>")
              else
                  response.write("<option value="&i&">"&i&"</option>")
            end if
         next

         response.write("</select></td><td colspan=2 rowspan=2 class='td-c bold BoldBorders'>OBSERVACIONES <BR> <input class='td-r anchox1' type='text' name='vNotas' id='vNotas' value='' ></td></tr>")
         response.write("<tr><td class='td-r bold whiteBorders' Colspan=4># Indique unidad a utilizar:</td><td class='td-l bold'><select name='vUnidad'> ")

         strSQL="select * from Flotilla where activado=1 and tipo_activo in (1,2,3) order by tipo_activo"
         'response.write strSQL  
         rsUpdateEntry.Open strSQL, adoCon
         response.write("<option value='000' SELECTED>N/A</option>")
         while not  rsUpdateEntry.EOF
              response.write("<option value='"& rsUpdateEntry("id_unidad") &"'>"& rsUpdateEntry("id_unidad") &"</option>")
              rsUpdateEntry.MoveNext
         wend
         rsUpdateEntry.close
         response.write("</select></td></tr>")
    

         response.write("<tr><td width='10%' class=whiteBorders style='font-size: 14px;'>&nbsp;</td><td width='12%' class=whiteBorders style='font-size: 14px;'>&nbsp;</td><td width='17%' class=whiteBorders style='font-size: 14px;'>&nbsp;</td><td width='12%' class=whiteBorders style='font-size: 14px;'>&nbsp;</td><td width='22%' class=whiteBorders style='font-size: 14px;'>&nbsp;</td><td width='12%' class=whiteBorders style='font-size: 14px;'>&nbsp;</td><td width='15%' class=whiteBorders style='font-size: 14px;'>&nbsp;</td></tr>")

         response.write("<tr><td class=whiteBorders>&nbsp;</td><td colspan=2 class='td-c UnderBorder'><select class=anchox2  name='vElabora'> ")
           strSQL="select * from Empleados  where FechaEgreso is null order by ID desc "
           'response.write strSQL  
           rsUpdateEntry.Open strSQL, adoCon

           while not  rsUpdateEntry.EOF
              response.write("<option value='"& rsUpdateEntry("nombres") &" " & rsUpdateEntry("PATERNO") &" " & rsUpdateEntry("MATERNO") &"'>"& rsUpdateEntry("nombres") &" " & rsUpdateEntry("PATERNO") &" " & rsUpdateEntry("MATERNO") &"</option>")
              rsUpdateEntry.MoveNext
           wend
            rsUpdateEntry.close
            response.write("</select></td><td colspan=2 class=whiteBorders>&nbsp;</td><td colspan=2 class='td-c UnderBorder'>ALICIA JAUREGUI RUIZ</td></td></tr>")
         response.write("<tr><td class=whiteBorders>&nbsp;</td><td colspan=2 class='td-c whiteBorders'>ELABORA</td><td colspan=2 class=whiteBorders>&nbsp;</td><td colspan=2 class='td-c whiteBorders'>AUTORIZA</td></td></tr>")

           separador
          response.write("<td  colspan='3' class='td-c whiteBorders azul'> Fecha Estimada de viaje:  <input name='vViaje' id='vViaje' type='date'/>  </td>")
         response.write("<td  colspan='2' class='td-c whiteBorders azul'>Cantidad d&iacute;as vi&aacute;ticos: <select  name='vdias' id='vdias'>") 
         for i=1 to 31
              response.write("<option value='"& i & "'>" &  i  & "</option> ")
         next 
         response.write("</select></td><td  colspan='2' class='td-c whiteBorders azul'><input type='submit' value='continuar'> </form> </td></tr><table>")

end sub




Sub AutorizaViatico  'contenido 13'  
     if request("item")<>"" then
             Session("item")=request("item")
     end if
     strSQL="select * from VIATICOS_R1  where ID=" & Session("item")
     'response.write strSQL  
     rsUpdateEntry.Open strSQL, adoCon


     response.write("<center><P>&nbsp;</P><table width='640px' border=0 cellpadding=2 align=center class=whiteBorders>")
     response.write("<form id='viaticos' method='post' action='ShowContent.asp'>")
     response.write("<input type='hidden' name='contenido' value='14'>")        

     response.write("<tr><td colspan=2 class='td-c subject BoldBorders'  style='padding-top: 20px;padding-bottom: 20px;'>SOLICITUD DE VIATICOS </td></tr>")
     response.write("<tr><td class='td-r azul'>Folio:</td><td class='td-l'>" & rsUpdateEntry("ID") &"</td></tr>") 
     response.write("<tr><td class='td-r azul'>Fecha Solicitud:</td><td class='td-l'>" & rsUpdateEntry("DocDate") &"</td></tr>")
     response.write("<tr><td class='td-r azul'>Monto:</td><td class='td-l'>$" & rsUpdateEntry("Total") &"</td></tr>") 
     response.write("<tr><td class='td-r azul'>Empresa:</td><td class='td-l'><B>" & rsUpdateEntry("Empresa") &"</B></td></tr>") 
     response.write("<tr><td class='td-r azul'>Beneficiario:</td><td class='td-l'>" & rsUpdateEntry("Beneficiario") &"</td></tr>")       
     response.write("<tr><td class='td-r azul'>Motivo:</td><td class='td-l'>" & rsUpdateEntry("Motivo") &"</td></tr>") 
     response.write("<tr><td class='td-r azul'>Elabor&oacute;:</td><td class='td-l'>" & rsUpdateEntry("Elaborador") &"</td></tr>") 
     response.write("<tr><td class='td-r azul'>Fecha Viaje:</td><td class='td-l'>" & day(rsUpdateEntry("TravelDate1")) &"/" & month(rsUpdateEntry("TravelDate1")) &"/" & year(rsUpdateEntry("TravelDate1"))  &"</td></tr>") 
     response.write("<tr><td class='td-r azul'>Dias de vi&aacute;ticos:</td><td class='td-l'>" & rsUpdateEntry("dias") &"</td></tr>")    
     response.write("<tr><td class='td-r azul'>Estatus Actual:</td><td class='td-l'><B>" & rsUpdateEntry("ESTATUS") &"</B></td></tr>") 

     response.write("<tr><td class='td-r azul'>Acci&oacute;n disponible:</td><td class='td-l'><select name='vHelpdesk' id='vHelpdesk'>")

     select case  rsUpdateEntry("helpdesk")
     case 0   response.write("<option value='1'>Autorizar: mandar a outsourcing</option>")
              response.write("<option value='10'>Cancelar solicitud</option>")
     case 1   response.write("<option value='3'>Comprobado por DFW</option>")
              response.write("<option value='10'>Cancelar solicitud</option>")
     case 2   response.write("<option value='3'>Comprobado por DFW</option>")
              response.write("<option value='10'>Cancelar solicitud</option>")
     case 3   response.write("<option value='4'>Finalizado</option>")
           

     case else response.write("<option value='"& rsUpdateEntry("helpdesk") &"'>no existen mas opciones</option>")
     end select

     response.write("</select></td></tr>")    
     response.write("<tr><td colspan=2 class='td-c'><input type='submit' value='actualizar'> </form>  </td> </tr> ")
    
     response.write("</table>")
     rsUpdateEntry.close


end sub




Sub UpViatico  'contenido 14'
     separador
     response.write( request("vHelpdesk") &"<BR>")     
     select case request("vHelpdesk")
     case "1"   strSQL="UPDATE  VIATICOS_R1 SET helpdesk=1,ESTATUS='Autorizado',sentdate=null  where ID=" & Session("item")                'Mandar a outsourcing'
     case "3"   strSQL="UPDATE  VIATICOS_R1 SET helpdesk=3,ESTATUS='Comprobado por DFW',sentdate=null  where ID=" & Session("item")        
     case "4"   strSQL="UPDATE  VIATICOS_R1 SET helpdesk=4,ESTATUS='Finalizado'  where ID=" & Session("item")                              
     case "10"  strSQL="UPDATE  VIATICOS_R1 SET helpdesk=10,ESTATUS='cancelado'  where ID=" & Session("item")            
     end select
     response.write strSQL

     rsUpdateEntry.Open strSQL, adoCon
     response.redirect("ShowContent.asp?contenido=12&msg=La solicitud "& Session("item") &", ha sido actualizada" )

end sub  





sub CheckViaticos  'contenido 11'

         'VALIDACIONES'
         'response.write ("Fecha=" & request("vViaje") & "<BR>")
         cantidad=trim(request("vcantidad"))
         cantidad=replace(cantidad,",","")
         cantidad=replace(cantidad,"$","")
         cantidad=replace(cantidad,".00","")
         if cantidad<>"" then
            entero=int(cantidad)
         end if

          if  cantidad=""  or entero>99999  or (  request("vCuenta1")="" and request("vCuenta2")=""  ) or (  request("vCuenta1")<>"" and request("vCuenta2")<>""  )  or request("vViaje")="" then
               vstring="ShowContent.asp?contenido=11&msg=error&control=1"
               if cantidad<>"" then
                   vstring=vstring & "&vcantidad=" & cantidad
                end if
               if request("vCuenta1")<>"" then
                   vstring=vstring & "&vCuenta1=" & request("vCuenta1")
                end if
                if request("vCuenta2")<>"" then
                   vstring=vstring & "&vCuenta2=" & request("vCuenta2")
                end if
                if request("vViaje")<>"" then
                   vstring=vstring & "&vViaje=" & request("vViaje")
                end if              
                response.redirect(vstring)
          end if



         response.write("<P>&nbsp;</p><center><table width='1000px' border=1 cellpadding=2 align=center class=whiteBorders>")
         response.write("<tr><td colspan=7 class='td-c subject BoldBorders'  style='padding-top: 20px;padding-bottom: 20px;'>SOLICITUD DE VIATICOS </td></tr>")

         response.write("<form id='viaticos' method='post' action='ShowContent.asp'>")

         response.write("<input type='hidden' name='contenido' value='11'>")
         response.write("<input type='hidden' name='control' value='3'>")   'va a ir: SolicitarViaticos '


         response.write("<input type='hidden' name='vremisiones' value="& request("vremisiones") &">")
         response.write("<input type='hidden' name='vUnidad' value="& request("vUnidad") &">")

         response.write("<input type='hidden' name='vfecha' value='"& date()&"'>")
         response.write("<input type='hidden' name='vcantidad' value='"& cantidad &"'>")
         response.write("<input type='hidden' name='vempresa' value='"& request("vempresa") &"'>")

         response.write("<input type='hidden' name='vbeneficiario' value='"& request("vbeneficiario") &"'>")

         vAlmacen=""
         select case trim(request("vbeneficiario"))
          case "CESAR NOE BALDERRAMA GONZALEZ" 
                vAlmacen="MXL"
          case "ISMAEL HERNANDEZ TOVAR"
                vAlmacen="SJR"          
          case "SANTIAGO JESUS MARTINEZ FIGUEROA"
               vAlmacen="MXL"
          case "VICTOR MANUEL PAZ DE JESUS"
              vAlmacen="SJR"
          case else  vAlmacen="N/A"
        end select
         
         response.write("<input type='hidden' name='vAlmacen' value='"& vAlmacen &"'>")

         response.write("<input type='hidden' name='vCuenta1' value='"& request("vCuenta1") &"'>")
         if request("vCuenta1") <>"" then
                response.write("<input type='hidden' name='vMonto1' value='"& cantidad &"'>")
         else
                response.write("<input type='hidden' name='vMonto1' value=''>")
         end if
         response.write("<input type='hidden' name='vCuenta2' value='"& request("vCuenta2") &"'>")
         if request("vCuenta2") <>"" then
                response.write("<input type='hidden' name='vMonto2' value='"& request("vcantidad")  &"'>")
         else
                response.write("<input type='hidden' name='vMonto2' value=''>")
         end if

         response.write("<input type='hidden' name='vNotas' value='"& request("vNotas") &"'>")
         response.write("<input type='hidden' name='vElabora' value='"& request("vElabora") &"'>")
         response.write("<input type='hidden' name='vViaje' value='"& request("vViaje") &"'>")
         response.write("<input type='hidden' name='vdias' value='"& request("vdias") &"'>")
         
         

         response.write("<tr><td colspan=7 style='font-size: 8px;' class=whiteBorders>&nbsp;</td></tr>")
         response.write("<tr><td class='td-r bold whiteBorders'>FECHA:</td><td class='td-r BoldBorders'>  "& date()  &"   </td><td colspan=4 class='td-r bold whiteBorders'>CANTIDAD:</td><td class='td-r UnderBorder'> "& FormatCurrency(cantidad)  & "</td></tr>")
         response.write("<tr><td colspan=7 style='font-size: 6px;' class=whiteBorders >&nbsp;</td></tr>")

         response.write("<tr><td class=whiteBorders>&nbsp;</td><td class='td-r bold whiteBorders'>EMPRESA:</td><td class='td-l UnderBorder' colspan=4>"&  request("vEmpresa") & " </td><td class=whiteBorders>&nbsp;</td></tr>")

         response.write("<tr><td colspan=7 style='font-size: 6px;' class=whiteBorders>&nbsp;</td></tr>")
         response.write("<tr><td colspan=2 class='td-r bold whiteBorders'>NOMBRE DEL BENEFICIARIO:</td><td class='td-l UnderBorder' colspan=4 >"&  request("vbeneficiario") & "</td><td class=whiteBorders>&nbsp;</td></tr>")
         response.write("<tr><td colspan=7 style='font-size: 6px;' class=whiteBorders>&nbsp;</td></tr>")

         vstring=""
         
         response.write("<tr><td colspan=2  class='td-r bold whiteBorders'>IMPORTE DE LETRA:</td><td class='td-l UnderBorder azul' colspan=4> ") '--99,999
         longitud=len(cantidad)
         if longitud=5 then
               select case mid(cantidad,1,2)
               case "10" vstring="diez mil "
               case "11" vstring="once mil "
               case "12" vstring="doce mil "
               case "13" vstring="trece mil "
               case "14" vstring="catorce mil "
               case "15" vstring="quince mil "  
               case "16" vstring="diezciseis mil " 
               case "17" vstring="diezcisiete mil " 
               case "18" vstring="diezciocho mil " 
               case "19" vstring="diezcinueve mil "         
               case else
                   Select case mid(cantidad,1,1)
                   case "2" vstring="veinte "
                   case "3" vstring="treinta "
                   case "4" vstring="cuarenta "
                   case "5" vstring="cincuenta "
                   case "6" vstring="sesenta "
                   case "7" vstring="setenta "
                   case "8" vstring="ochenta "
                   case "9" vstring="noventa "
                   end select
                   
               end select
               if mid(cantidad,2,1)<>"0" and entero>=20000  then vstring=vstring &"y " end if
          end if
          'response.write("entero=" & trim(entero) &"<BR>")

          if longitud>3 and (entero<10000 or entero>=20000) then
                  Select case mid(cantidad,longitud-3,1)
                   case "0" vstring=vstring &" mil "
                   case "1" vstring=vstring &"un mil "
                   case "2" vstring=vstring &"dos mil "
                   case "3" vstring=vstring &"tres mil "
                   case "4" vstring=vstring &"cuatro mil "
                   case "5" vstring=vstring &"cinco mil "
                   case "6" vstring=vstring &"seis mil "
                   case "7" vstring=vstring &"siete mil "
                   case "8" vstring=vstring &"ocho mil "
                   case "9" vstring=vstring &"nueve mil "
                   end select
            end if

           if longitud>2 then            
                   Select case mid(cantidad,longitud-2,1)
                   case "1" vstring=vstring &"ciento "
                   case "2" vstring=vstring &"dos cientos "
                   case "3" vstring=vstring &"tres cientos "
                   case "4" vstring=vstring &"cuatro cientos "
                   case "5" vstring=vstring &"quinientos "
                   case "6" vstring=vstring &"seis cientos "
                   case "7" vstring=vstring &"siete cientos "
                   case "8" vstring=vstring &"ocho cientos "
                   case "9" vstring=vstring &"nueve cientos "
                   end select
                                    
            end if

            if longitud>1 then            
                   Select case mid(cantidad,longitud-1,2)
                   case "10" vstring=vstring &"diez "
                   case "11" vstring=vstring &"once "
                   case "12" vstring=vstring &"doce "
                   case "13" vstring=vstring &"trece "
                   case "14" vstring=vstring &"catorce "
                   case "15" vstring=vstring &"quince "  
                   case "16" vstring=vstring &"dieciseis " 
                   case "17" vstring=vstring &"diecisiete " 
                   case "18" vstring=vstring &"dieciocho " 
                   case "19" vstring=vstring &"diezcinueve "         
                   case else

                         Select case mid(cantidad,longitud-1,1)
                         case "2" vstring=vstring &"veinte "
                         case "3" vstring=vstring &"treinta "
                         case "4" vstring=vstring &"cuarenta "
                         case "5" vstring=vstring &"cincuenta "
                         case "6" vstring=vstring &"sesenta "
                         case "7" vstring=vstring &"setenta "
                         case "8" vstring=vstring &"ochenta "
                         case "9" vstring=vstring &"noventa "
                         end select   

                         Select case mid(cantidad,longitud,1)
                         case "1" vstring=vstring &"uno "
                         case "2" vstring=vstring &"dos "
                         case "3" vstring=vstring &"tres "
                         case "4" vstring=vstring &"cuatro "
                         case "5" vstring=vstring &"cinco "
                         case "6" vstring=vstring &"seis "
                         case "7" vstring=vstring &"siete "
                         case "8" vstring=vstring &"ocho "
                         case "9" vstring=vstring &"nueve "
                         end select       


                    end select
               end if

         vstring=vstring&"en MN"

         response.write("<input class='td-l anchox4' type='text' name='vLetras' id='vLetras' value='"& vstring &"'></td><td class=whiteBorders>&nbsp;</td></tr>")
         response.write("<tr><td colspan=7 style='font-size: 10px;' class=whiteBorders>&nbsp;</td></tr>")

         response.write("<tr><td colspan=2 class=whiteBorders>&nbsp;</td><td class='td-l bold under whiteBorders'>BANCO</td><td colspan=4 class=whiteBorders>&nbsp;</td></tr>")

         response.write("<tr><td colspan=2 class='td-r bold whiteBorders'>BANCO (1):</td><td class='td-l UnderBorder'>BANCOMER</td><td class='td-r bold UnderBorder'>CUENTA:</td><td class='td-l BoldBorders'> "&  request("vCuenta1") & "</td><td class='td-r bold whiteBorders'>MONTO:</td><td class='td-l UnderBorder'>")

          if request("vCuenta1") <>"" then
                response.write( cantidad &"</td></tr>")
         else
                response.write("&nbsp;</td></tr>")
         end if

         response.write("<tr><td colspan=2 class='td-r bold whiteBorders'>BANCO (2):</td><td class='td-l UnderBorder'>BANAMEX</td><td class='td-r bold UnderBorder'>CUENTA CLABE:</td><td class='td-l BoldBorders'> "&  request("vCuenta2") & "</td><td class='td-r bold whiteBorders'>MONTO:</td><td class='td-l UnderBorder'>")
         if request("vCuenta2") <>"" then
                response.write( cantidad &"</td></tr>")
         else
                response.write("&nbsp;</td></tr>")
         end if



         response.write("<tr><td colspan=2 class='td-r bold whiteBorders'>BANCO (3):</td><td class='td-l UnderBorder'>&nbsp;</td><td class='td-r bold UnderBorder'>&nbsp;</td><td class='td-l whiteBorders'>&nbsp;</td><td class='td-r bold whiteBorders'>MONTO:</td><td class='td-l UnderBorder'>&nbsp;</td></tr>")
         response.write("<tr><td colspan=7 style='font-size: 4px;' class=whiteBorders>&nbsp;</td></tr>")

         response.write("<tr><td class='td-r bold whiteBorders'>N&Oacute;MINA:</td><td colspan=2 class='td-l bold UnderBorder'>DELTAFLOW DE MEXICO </td><td colspan=2 class=whiteBorders>&nbsp;</td><td class='td-r bold whiteBorders'>&nbsp;</td><td class='td-l whiteBorders'>&nbsp;</td></tr>")
         response.write("<tr><td colspan=7 style='font-size: 12px;' class=whiteBorders>&nbsp;</td></tr>")

         
         response.write("<tr><td class='td-r bold whiteBorders'>MOTIVO:</td><td colspan=4 class='td-l bold UnderBorder'>")
         if request("vremisiones")>0 then
             response.write(" entrega de : <table border=1>")
             for i=1 to request("vremisiones") step 1
                 response.write("<tr>")
                 response.write("<td with='90px' class='subtitulo td-r'>remision"&i&"</td>")
                 response.write("<td width='60px'><input class='anchoSmall' type='number' name='remision"&i&"'></td>")
                 response.write("<td width='60px'><select name='RS"&i&"'><option value='DMX'>DMX</option><option value='MME'>MME</option><option value='DFW' SELECTED>DFW</option></select></td>")
                 response.write("<td width='60px'><input type='checkbox' name='provisional"&i&"'>provisional")
                 response.write("</tr>")
             next
         end if
         response.write("</table>")
         

         response.write("</td><td colspan=2 rowspan=2 class='td-c bold BoldBorders'>OBSERVACIONES <BR>  "&  request("vNotas") & "</td></tr>")
         response.write("<tr><td class='td-r bold whiteBorders' td colspan=5 >&nbsp;</td></tr>")

         response.write("<tr><td width='10%' class=whiteBorders style='font-size: 14px;'>&nbsp;</td><td width='12%' class=whiteBorders style='font-size: 14px;'>&nbsp;</td><td width='17%' class=whiteBorders style='font-size: 14px;'>&nbsp;</td><td width='12%' class=whiteBorders style='font-size: 14px;'>&nbsp;</td><td width='22%' class=whiteBorders style='font-size: 14px;'>&nbsp;</td><td width='12%' class=whiteBorders style='font-size: 14px;'>&nbsp;</td><td width='15%' class=whiteBorders style='font-size: 14px;'>&nbsp;</td></tr>")

         response.write("<tr><td class=whiteBorders>&nbsp;</td><td colspan=2 class='td-c UnderBorder'>"&  request("vElabora") & "</td><td colspan=2 class=whiteBorders>&nbsp;</td><td colspan=2 class='td-c UnderBorder'>ALICIA JAUREGUI RUIZ</td></td></tr>")
         response.write("<tr><td class=whiteBorders>&nbsp;</td><td colspan=2 class='td-c whiteBorders'>ELABORA</td><td colspan=2 class=whiteBorders>&nbsp;</td><td colspan=2 class='td-c whiteBorders'>AUTORIZA</td></td></tr>")

           separador

            response.write("<tr><td  colspan='3' class='td-c whiteBorders'> Fecha Estimada de viaje:  "&  request("vViaje") & "  </td>")
         response.write("<td  colspan='2' class='td-c whiteBorders'>Cantidad d&iacute;as vi&aacute;ticos: "&  request("vdias") & "</td><td  colspan='2' class='td-c whiteBorders'><input class=anchox2 type='submit' value='Enviar solicitud'> </form> </td></tr><table>")



end sub  





sub SolicitarViaticos 'contenido 11'
   
    response.write("<P>&nbsp;</P>")  

    if request("vremisiones")>0 then

    	  remisioneSolicitadas=0          
    	  vRemisionesValidas=0
          vMotivo="Entrega de remisiones: "

          for i=1 to int(trim(request("vremisiones"))) step 1
              vString1="remision"&trim(i)
              vString2="RS"& trim(i)
              vString3="provisional"& trim(i)

              if request(vString3)="true" or request(vString3)="on" then
                  vRemisionesValidas=vRemisionesValidas+1
              else
	              strSQL="select * from Notificaciones where RazonSocial='"& request(vString2) &"' and  tipo='remision' and DocNum="& request(vString1)
	              response.write strSQL
	              rsUpdateEntry5.Open strSQL, adoCon4

	              if not rsUpdateEntry5.EOF then vRemisionesValidas=vRemisionesValidas+1
                  rsUpdateEntry5.close	
	          end if
              vMotivo=vMotivo & request(vString1) & " "              
          next
          if int(trim(request("vremisiones")))<>vRemisionesValidas then response.redirect("ShowContent.asp?contenido=11&control=1&msg=1")
    else
          vMotivo=request("vNotas")
    end if
    
     strSQL="insert into VIATICOS_R1 (ID,DocDate,Total,Empresa,Letras,Beneficiario,Cuenta1,Monto1,Cuenta2,Monto2,Motivo,Notas,Elaborador,TravelDate1,Dias,Estatus,helpdesk,ID_unidad,almacen) values ((select max(ID) from VIATICOS_R1)+1,getdate()," & request("vcantidad") & ",'"  & request("vEmpresa") & "','"   & request("vLetras") & "','" & request("vbeneficiario") & "','"   & request("vCuenta1") & "','"  & request("vMonto1") & "','"  & request("vCuenta2") & "','" & request("vMonto2")  & "','" & vMotivo  & "','" & mid( request("vNotas"),1,150)   & "','" & request("vElabora") & "','" & request("vViaje")  & "'," & request("vdias")  & ",'Por autorizar',0,'"& request("vUnidad") &"','"& request("vAlmacen") &"') "

     response.write(strSQL &"<BR>")
     rsUpdateEntry.Open strSQL, adoCon

     strSQL="select TOP 1 max(ID) as ID from VIATICOS_R1 where Total="& request("vcantidad") & " and day(docdate)=day(getdate()) and month(docdate)=month(getdate()) and year(docdate)=year(getdate()) and Beneficiario='" & request("vbeneficiario") & "' and Elaborador='"   &  request("vElabora") & "' "
     response.write(strSQL &"<BR>")
      rsUpdateEntry2.Open strSQL, adoCon

   
      vID=""
      if  not rsUpdateEntry2.EOF then
            vID=rsUpdateEntry2("ID")
            ruta1="MD d:\fileserver\wwwroot\repositorio\ADMON\Viaticos\" & vID            
            ruta3="R:/ADMON/Viaticos/" & vID

            strSQL="EXEC xp_cmdshell '"& ruta1 &"'"
            'response.write strSQL
            rsUpdateEntry3.Open strSQL, adoCon


            strSQL="UPDATE VIATICOS_R1 set TravelDate2=DATEADD(DAY,Dias-1,TravelDate1) WHERE ID=" & vID
            rsUpdateEntry5.Open strSQL, adoCon     

            for i=1 to request("vremisiones") step 1
                vString1="remision"&trim(i)
                vString2="RS"& trim(i)
                strSQL="insert into ViaticoRemision (ID,Remision,RS,DocDate,LogDate,ID_USUARIO) values ("& vID &","& request(vString1) &",'"& request(vString2)&"',getdate(),getdate(),'"& session("session_id") &"')"
                rsUpdateEntry6.Open strSQL, adoCon     
            next              
          
      end if

    


      if vID<>"" then
           response.redirect("ShowContent.asp?contenido=12&ID=" & vID &"&msg=su solicitud de viatico fue recibida&ruta1=" & ruta3 )

      else
            response.redirect("ShowContent.asp?contenido=12&msg=su solicitud de viatico fue recibida")

      end if
          

end sub





sub ShowReembolso 'contenido 15'

         response.write("<P>&nbsp;</p>")
         if request("msg")<>"" then         	    
         	    if int(trim(request("msg")))=1 then
                     vMsg="las remisiones no perteneces a esa Razon Social o alguna remision ingresada no existe"
         	    else
                     vMsg="Formulario debe tener monto <100,000 motivo, cuenta bancaria y Ciudad destino. Corrija y vuelve a presionar continuar"
         	    end if
                response.write("<P class='td-c alert'>"& vMsg &"</p>")
         end if
         response.write("<center><table width='1000px' border=1 cellpadding=2 align=center class=whiteBorders>")
         response.write("<tr><td colspan=7 class='td-c subject BoldBorders'  style='padding-top: 20px;padding-bottom: 20px;'>SOLICITUD DE REEMBOLSO / COMPROBACION </td></tr>")

         response.write("<form id='reembolso' method='post' action='ShowContent.asp'>")
         response.write("<input type='hidden' name='contenido' value='15'>")
         response.write("<input type='hidden' name='control' value='2'>")

         response.write("<tr><td colspan=7 style='font-size: 8px;' class=whiteBorders>&nbsp;</td></tr>")
         response.write("<tr><td class='td-r bold whiteBorders'>FECHA:</td><td class='td-r BoldBorders'>  "& date() &"   </td><td colspan=4 class='td-r bold whiteBorders'>CANTIDAD:</td><td class='td-r UnderBorder azul'>$ XXX.XX </td></tr>")
         response.write("<tr><td colspan=7 style='font-size: 6px;' class=whiteBorders >&nbsp;</td></tr>")

         response.write("<tr><td class=whiteBorders>&nbsp;</td><td class='td-r bold whiteBorders'>EMPRESA:</td><td class='td-l UnderBorder' colspan=4>")

         response.write("<select class=anchox4  name='vempresa'>")
         response.write("<option value='Corporativo Empresarial Izcalli SA de CV' SELECTED>Corporativo Empresarial Izcalli SA de CV </option>")
         response.write("<option value='Servicios Profesionales, Legales, Contables, Financieros, Recursos Humanos y Gestoria SR de CV'>Servicios Profesionales, Legales, Contables, Financieros, Recursos Humanos y Gestoria SR de CV</option>")
         response.write("<option value='DELTAFLOW S DE RL DE CV'>DELTAFLOW S DE RL DE CV</option>")
         response.write("<option value='DELTAFLOW DE MEXICO S DE RL DE CV' SELECTED>DELTAFLOW DE MEXICO S DE RL DE CV</option>")
         response.write("<option value='MEIDE DE MEXICO S DE RL DE CV'>MEIDE DE MEXICO S DE RL DE CV</option>")

         response.write("</select> </td><td class=whiteBorders>&nbsp;</td></tr>")

         response.write("<tr><td colspan=7 style='font-size: 6px;' class=whiteBorders>&nbsp;</td></tr>")
         response.write("<tr><td colspan=2 class='td-r bold whiteBorders'>NOMBRE DEL BENEFICIARIO:</td><td class='td-l UnderBorder' colspan=4 ><select class=anchox4  name='vbeneficiario'> ")


           strSQL="select * from Empleados where FechaEgreso is null order by ID desc "
           'response.write strSQL  
           rsUpdateEntry.Open strSQL, adoCon

           while not  rsUpdateEntry.EOF
           	  if UCASE(trim(rsUpdateEntry("ID_USUARIO"))) = UCASE(trim(session("session_id"))) then
           	  	   response.write("<option value='"& rsUpdateEntry("nombres") &" " & rsUpdateEntry("PATERNO") &" " & rsUpdateEntry("MATERNO") &"' SELECTED>"& rsUpdateEntry("nombres") &" " & rsUpdateEntry("PATERNO") &" " & rsUpdateEntry("MATERNO") &"</option>")
           	  	   vCuentaReembolsos=rsUpdateEntry("Reembolsos")
           	  	   vDestino=rsUpdateEntry("U_Destino")
           	  else
           	       response.write("<option value='"& rsUpdateEntry("nombres") &" " & rsUpdateEntry("PATERNO") &" " & rsUpdateEntry("MATERNO") &"'>"& rsUpdateEntry("nombres") &" " & rsUpdateEntry("PATERNO") &" " & rsUpdateEntry("MATERNO") &"</option>")

           	  end if
              
              rsUpdateEntry.MoveNext
           wend
            rsUpdateEntry.close
            response.write("</select></td><td class=whiteBorders>&nbsp;</td></tr>")
         response.write("<tr><td colspan=7 style='font-size: 6px;' class=whiteBorders>&nbsp;</td></tr>")

         response.write("<tr><td colspan=2  class='td-r bold whiteBorders'>IMPORTE DE LETRA:</td><td class='td-l UnderBorder' colspan=4> &nbsp;</td><td class=whiteBorders>&nbsp;</td></tr>")
         response.write("<tr><td colspan=7 style='font-size: 10px;' class=whiteBorders>&nbsp;</td></tr>")

         response.write("<tr><td colspan=2 class=whiteBorders>&nbsp;</td><td class='td-l bold under whiteBorders'>BANCO</td><td colspan=4 class=whiteBorders>&nbsp;</td></tr>")

         response.write("<tr><td colspan=2 class='td-r bold whiteBorders'>BANCO (1):</td><td class='td-l UnderBorder'>BANCOMER</td><td class='td-r bold UnderBorder'>CUENTA:</td><td class='td-l BoldBorders azul'> <input class='td-r anchox1' type='text' name='vCuenta1' id='vCuenta1' value='"& request("vCuenta1") &"' ></td><td class='td-r bold whiteBorders'>MONTO:</td><td class='td-l UnderBorder'>&nbsp;</td></tr>")

         response.write("<tr><td colspan=2 class='td-r bold whiteBorders'>BANCO (2):</td><td class='td-l UnderBorder'>BANAMEX</td><td class='td-r bold UnderBorder'>CUENTA CLABE:</td><td class='td-l BoldBorders azul'> <input class='td-r anchox1' type='text' name='vCuenta2' id='vCuenta2' value='"& vCuentaReembolsos &"' ></td><td class='td-r bold whiteBorders'>MONTO:</td><td class='td-l UnderBorder'>&nbsp;</td></tr>")

         response.write("<tr><td colspan=2 class='td-r bold whiteBorders'>BANCO (3):</td><td class='td-l UnderBorder'>&nbsp;</td><td class='td-r bold UnderBorder'>&nbsp;</td><td class='td-l whiteBorders'>&nbsp;</td><td class='td-r bold whiteBorders'>MONTO:</td><td class='td-l UnderBorder'>&nbsp;</td></tr>")
         response.write("<tr><td colspan=7 style='font-size: 4px;' class=whiteBorders>&nbsp;</td></tr>")

         response.write("<tr><td class='td-r bold whiteBorders'>N&Oacute;MINA:</td><td colspan=2 class='td-l bold UnderBorder'>DELTAFLOW DE MEXICO </td><td colspan=2 class=whiteBorders>&nbsp;</td><td class='td-r bold whiteBorders'>&nbsp;</td><td class='td-l whiteBorders'>&nbsp;</td></tr>")
         response.write("<tr><td colspan=7 style='font-size: 12px;' class=whiteBorders>&nbsp;</td></tr>")
         response.write("<tr><td class='td-r bold whiteBorders'>MOTIVO:</td><td colspan=4 class='td-l bold UnderBorder azul'><input class='td-l anchox4' type='text' name='vMotivo' id='vMotivo' value='' ></td><td colspan=2 rowspan=2 class='td-c bold BoldBorders'>OBSERVACIONES <BR> <input class='td-r anchox1' type='text' name='vNotas' id='vNotas' value='' ></td></tr>")
         response.write("<tr><td class='td-r bold whiteBorders' colspan=2>CIUDAD DESTINO:</td><td colspan=3 class='td-l bold UnderBorder azul'><input class='td-l anchox3' type='text' name='vCiudad' id='vCiudad' value='"&vDestino&"' ></td></tr>")

         response.write("<tr><td width='10%' class=whiteBorders style='font-size: 14px;'>&nbsp;</td><td width='12%' class=whiteBorders style='font-size: 14px;'>&nbsp;</td><td width='17%' class=whiteBorders style='font-size: 14px;'>&nbsp;</td><td width='12%' class=whiteBorders style='font-size: 14px;'>&nbsp;</td><td width='22%' class=whiteBorders style='font-size: 14px;'>&nbsp;</td><td width='12%' class=whiteBorders style='font-size: 14px;'>&nbsp;</td><td width='15%' class=whiteBorders style='font-size: 14px;'>&nbsp;</td></tr>")

         response.write("<tr><td class=whiteBorders>&nbsp;</td><td colspan=2 class='td-c UnderBorder'><select class=anchox2  name='vElabora'> ")
           strSQL="select * from Empleados  where FechaEgreso is null order by ID desc "
           'response.write strSQL  
           rsUpdateEntry.Open strSQL, adoCon

           while not  rsUpdateEntry.EOF
           	  if UCASE(trim(rsUpdateEntry("ID_USUARIO"))) = UCASE(trim(session("session_id"))) then
           	  	   response.write("<option value='"& rsUpdateEntry("nombres") &" " & rsUpdateEntry("PATERNO") &" " & rsUpdateEntry("MATERNO") &"' SELECTED>"& rsUpdateEntry("nombres") &" " & rsUpdateEntry("PATERNO") &" " & rsUpdateEntry("MATERNO") &"</option>")

           	  else
           	      response.write("<option value='"& rsUpdateEntry("nombres") &" " & rsUpdateEntry("PATERNO") &" " & rsUpdateEntry("MATERNO") &"'>"& rsUpdateEntry("nombres") &" " & rsUpdateEntry("PATERNO") &" " & rsUpdateEntry("MATERNO") &"</option>")

           	  end if
              
              rsUpdateEntry.MoveNext
           wend
            rsUpdateEntry.close
            response.write("</select></td><td colspan=2 class=whiteBorders>&nbsp;</td><td colspan=2 class='td-c UnderBorder'>ALICIA JAUREGUI RUIZ</td></td></tr>")
         response.write("<tr><td class=whiteBorders>&nbsp;</td><td colspan=2 class='td-c whiteBorders'>ELABORA</td><td colspan=2 class=whiteBorders>&nbsp;</td><td colspan=2 class='td-c whiteBorders'>AUTORIZA</td></td></tr>")

           response.write("<tr></td><td  colspan='10' class='td-c whiteBorders azul'><input type='submit' value='continuar'> </form> </td></tr><table>")

end sub





sub CheckReembolso   'contenido 15 CONTROL=2 '
         'VALIDACIONES'
         'response.write ("Fecha=" & request("vViaje") & "<BR>")
      
          if   request("vMotivo")="" or  (  request("vCuenta1")="" and request("vCuenta2")=""  ) or (  request("vCuenta1")<>"" and request("vCuenta2")<>""  )   then
                vstring="ShowContent.asp?contenido=15&msg=error&control=1"
              
               if request("vCuenta1")<>"" then
                   vstring=vstring & "&vCuenta1=" & request("vCuenta1")
                end if
                if request("vCuenta2")<>"" then
                   vstring=vstring & "&vCuenta2=" & request("vCuenta2")
                end if
                response.redirect(vstring)
          end if

        strSQL="insert into Reembolsos (ID,DocDate,Empresa,Beneficiario,Cuenta1,Monto1,Cuenta2,Monto2,Motivo,CiudadDestino,Notas,Elaborador,Estatus,helpdesk) values ((select max(ID)+1 from Reembolsos),getdate(),'"  & request("vEmpresa") &  "','" & request("vbeneficiario") & "','"   & request("vCuenta1") & "','"  & request("vMonto1") & "','"  & request("vCuenta2") & "','" & request("vMonto2")  & "','" & request("vMotivo")  & "','"  & request("vCiudad")  & "','" & request("vNotas")   & "','" & request("vElabora") & "','Por confirmar',0) "

     'response.write(strSQL)
     rsUpdateEntry.Open strSQL, adoCon

      strSQL="select TOP 1 * from Reembolsos where day(docdate)=day(getdate()) and month(docdate)=month(getdate()) and year(docdate)=year(getdate()) and Beneficiario='" & request("vbeneficiario") & "' and Elaborador='"   &  request("vElabora") & "' and Motivo='" & request("vMotivo")  & "' and empresa='" & request("vEmpresa") &"' order by DocDate desc"
      rsUpdateEntry.Open strSQL, adoCon

      
      Session("ID")=""
      if  not rsUpdateEntry.EOF then
            Session("ID")=rsUpdateEntry("ID")                              
      end if
      rsUpdateEntry.close

      Ruta1="MD d:\fileserver\wwwroot\repositorio\ADMON\Reembolsos\" & Session("ID")   
      strSQL="EXEC xp_cmdshell '"& ruta1 &"'"
      'response.write strSQL
      rsUpdateEntry2.Open strSQL, adoCon


         response.write("<P>&nbsp;</p><center><table width='1000px' border=1 cellpadding=2 align=center class=whiteBorders>")
         response.write("<tr><td colspan=7 class='td-c subject BoldBorders'  style='padding-top: 20px;padding-bottom: 20px;'>SOLICITUD DE REEMBOLSO / COMPROBACION "& Session("ID") &" &nbsp;&nbsp;&nbsp;<font class=alert>(no confirmado)</font></td></tr>")

  

         response.write("<tr><td colspan=7 style='font-size: 8px;' class=whiteBorders>&nbsp;</td></tr>")
         response.write("<tr><td class='td-r bold whiteBorders'>FECHA:</td><td class='td-r BoldBorders'>  "& date()  &"   </td><td colspan=4 class='td-r bold whiteBorders'>CANTIDAD:</td><td class='td-r UnderBorder'> $ XXX.XX</td></tr>")
         response.write("<tr><td colspan=7 style='font-size: 6px;' class=whiteBorders >&nbsp;</td></tr>")

         response.write("<tr><td class=whiteBorders>&nbsp;</td><td class='td-r bold whiteBorders'>EMPRESA:</td><td class='td-l UnderBorder' colspan=4>"&  request("vEmpresa") & " </td><td class=whiteBorders>&nbsp;</td></tr>")

         response.write("<tr><td colspan=7 style='font-size: 6px;' class=whiteBorders>&nbsp;</td></tr>")
         response.write("<tr><td colspan=2 class='td-r bold whiteBorders'>NOMBRE DEL BENEFICIARIO:</td><td class='td-l UnderBorder' colspan=4 >"&  request("vbeneficiario") & "</td><td class=whiteBorders>&nbsp;</td></tr>")
         response.write("<tr><td colspan=7 style='font-size: 6px;' class=whiteBorders>&nbsp;</td></tr>")
         response.write("<tr><td colspan=2  class='td-r bold whiteBorders'>IMPORTE DE LETRA:</td><td class='td-l UnderBorder azul' colspan=4> $ XXX.XX </td><td class=whiteBorders>&nbsp;</td></tr>")
         response.write("<tr><td colspan=7 style='font-size: 10px;' class=whiteBorders>&nbsp;</td></tr>")

         response.write("<tr><td colspan=2 class=whiteBorders>&nbsp;</td><td class='td-l bold under whiteBorders'>BANCO</td><td colspan=4 class=whiteBorders>&nbsp;</td></tr>")

         response.write("<tr><td colspan=2 class='td-r bold whiteBorders'>BANCO (1):</td><td class='td-l UnderBorder'>BANCOMER</td><td class='td-r bold UnderBorder'>CUENTA:</td><td class='td-l BoldBorders'> "&  request("vCuenta1") & "</td><td class='td-r bold whiteBorders'>MONTO:</td><td class='td-l UnderBorder'>")

          if request("vCuenta1") <>"" then
                response.write( cantidad &"</td></tr>")
         else
                response.write("&nbsp;</td></tr>")
         end if

         response.write("<tr><td colspan=2 class='td-r bold whiteBorders'>BANCO (2):</td><td class='td-l UnderBorder'>BANAMEX</td><td class='td-r bold UnderBorder'>CUENTA CLABE:</td><td class='td-l BoldBorders'> "&  request("vCuenta2") & "</td><td class='td-r bold whiteBorders'>MONTO:</td><td class='td-l UnderBorder'>")
         if request("vCuenta2") <>"" then
                response.write( cantidad &"</td></tr>")
         else
                response.write("&nbsp;</td></tr>")
         end if



         response.write("<tr><td colspan=2 class='td-r bold whiteBorders'>BANCO (3):</td><td class='td-l UnderBorder'>&nbsp;</td><td class='td-r bold UnderBorder'>&nbsp;</td><td class='td-l whiteBorders'>&nbsp;</td><td class='td-r bold whiteBorders'>MONTO:</td><td class='td-l UnderBorder'>&nbsp;</td></tr>")
         response.write("<tr><td colspan=7 style='font-size: 4px;' class=whiteBorders>&nbsp;</td></tr>")

         response.write("<tr><td class='td-r bold whiteBorders'>N&Oacute;MINA:</td><td colspan=2 class='td-l bold UnderBorder'>DELTAFLOW DE MEXICO </td><td colspan=2 class=whiteBorders>&nbsp;</td><td class='td-r bold whiteBorders'>&nbsp;</td><td class='td-l whiteBorders'>&nbsp;</td></tr>")
         response.write("<tr><td colspan=7 style='font-size: 12px;' class=whiteBorders>&nbsp;</td></tr>")
         response.write("<tr><td class='td-r bold whiteBorders'>MOTIVO:</td><td colspan=4 class='td-l bold UnderBorder'>"&  request("vMotivo") & "</td><td colspan=2 rowspan=2 class='td-c bold BoldBorders'>OBSERVACIONES <BR>  "&  request("vNotas") & "</td></tr>")
         response.write("<tr><td class='td-r bold whiteBorders' colspan=2>CIUDAD DESTINO:</td><td colspan=3 class='td-l bold UnderBorder'>"& request("vCiudad") &"</td></tr>")

         response.write("<tr><td width='10%' class=whiteBorders style='font-size: 14px;'>&nbsp;</td><td width='12%' class=whiteBorders style='font-size: 14px;'>&nbsp;</td><td width='17%' class=whiteBorders style='font-size: 14px;'>&nbsp;</td><td width='12%' class=whiteBorders style='font-size: 14px;'>&nbsp;</td><td width='22%' class=whiteBorders style='font-size: 14px;'>&nbsp;</td><td width='12%' class=whiteBorders style='font-size: 14px;'>&nbsp;</td><td width='15%' class=whiteBorders style='font-size: 14px;'>&nbsp;</td></tr>")

         response.write("<tr><td class=whiteBorders>&nbsp;</td><td colspan=2 class='td-c UnderBorder'>"&  request("vElabora") & "</td><td colspan=2 class=whiteBorders>&nbsp;</td><td colspan=2 class='td-c UnderBorder'>ALICIA JAUREGUI RUIZ</td></td></tr>")
         response.write("<tr><td class=whiteBorders>&nbsp;</td><td colspan=2 class='td-c whiteBorders'>ELABORA</td><td colspan=2 class=whiteBorders>&nbsp;</td><td colspan=2 class='td-c whiteBorders'>AUTORIZA</td></td></tr>")

         strSQL="select * from Reembolsos where ID=" & Session("ID")
         rsUpdateEntry.Open strSQL, adoCon
         vRFCReceptor=rsUpdateEntry("Empresa")
         rsUpdateEntry.close

         select case vRFCReceptor
          case "DELTAFLOW S DE RL DE CV"                   vRFCReceptor="DEL080208G25"
          case "DELTAFLOW DE MEXICO S DE RL DE CV"         vRFCReceptor="DME140422SS7"
          case "Corporativo Empresarial Izcalli SA de CV"  vRFCReceptor="IDA0606215S0"
          case "Servicios Profesionales, Legales, Contables, Financieros, Recursos Humanos y Gestoria SR de CV"  vRFCReceptor="SPL130121G98"
          case "MEIDE DE MEXICO S DE RL DE CV"  vRFCReceptor="MME2109143Y2"
          end select

         response.write("<tr><td  colspan='7' class='td-c whiteBorders'><HR></td></tr>")
         'response.write("<tr><td  colspan='7' class='td-c whiteBorders'>Por favor, coloque el total de los xmls y comprobantes no deducibles que va a reembolsar en la ubicaci&oacute;n <font class=alert>R:\ADMON\Reembolsos\" & Session("ID") &"</font>,<BR> cuando haya FINALIZADO la copia de xmls y archivos, presione el siguiente but&oacute;n: <BR> <a href='http://192.168.0.206/xmls_reembolso.asp?ID="& Session("ID")&"&Receptor=" & vRFCReceptor &"'><img src='/img/continuar.png' border=0></a></td></tr>")   

         response.write("<tr><td  colspan='7' class='td-c whiteBorders'>Por favor, coloque el total de los xmls y comprobantes no deducibles que va a reembolsar en la ubicaci&oacute;n: <button><a target='_self' href ='localexplorer:R:\ADMON\Reembolsos\" & Session("ID") &"' style='text-decoration: none;'> abrir Rembolsos</a></button><BR><font color=red>R:\ADMON\Reembolsos\" & Session("ID") &"</font><BR><BR> cuando haya FINALIZADO la copia de xmls y archivos, presione el siguiente but&oacute;n: <BR> <a href='http://192.168.0.206/xmls_reembolso.asp?ID="& Session("ID")&"&Receptor=" & vRFCReceptor &"'><img src='/img/continuar.png' border=0></a></td></tr>")      

         if Session("ID")<>"" then
              strSQL="delete xml where reembolso=" & Session("ID")
              rsUpdateEntry.Open strSQL, adoCon4
         end if
         
         response.write("<table>")


end sub  





Sub AutorizaReembolso  ' contenido 16'


     if request("item")<>"" then
             Session("item")=request("item")
     end if
     strSQL="select * from Reembolsos  where ID=" & Session("item")
     'response.write strSQL  
     rsUpdateEntry.Open strSQL, adoCon


     response.write("<center><P>&nbsp;</P><table width='640px' border=0 cellpadding=2 align=center class=whiteBorders>")

     response.write("<form id='reembolsos' method='post' action='ShowContent.asp'>")

         response.write("<input type='hidden' name='contenido' value='17'>")        

     response.write("<tr><td colspan=2 class='td-c subject BoldBorders'  style='padding-top: 20px;padding-bottom: 20px;'>SOLICITUD DE REEMBOLSO </td></tr>")
     response.write("<tr><td class='td-r azul'>Folio:</td><td class='td-l'>" & rsUpdateEntry("ID") &"</td></tr>") 
     response.write("<tr><td class='td-r azul'>Fecha Solicitud:</td><td class='td-l'>" & rsUpdateEntry("DocDate") &"</td></tr>")
     response.write("<tr><td class='td-r azul'>Monto:</td><td class='td-l'>$" & rsUpdateEntry("Total") &"</td></tr>") 
     response.write("<tr><td class='td-r azul'>Empresa:</td><td class='td-l'><B>" & rsUpdateEntry("Empresa") &"</B></td></tr>") 
     response.write("<tr><td class='td-r azul'>Beneficiario:</td><td class='td-l'>" & rsUpdateEntry("Beneficiario") &"</td></tr>")       
     response.write("<tr><td class='td-r azul'>Motivo:</td><td class='td-l'>" & rsUpdateEntry("Motivo") &"</td></tr>") 
     response.write("<tr><td class='td-r azul'>Elabor&oacute;:</td><td class='td-l'>" & rsUpdateEntry("Elaborador") &"</td></tr>") 
     'response.write("<tr><td class='td-r azul'>Fecha Viaje:</td><td class='td-l'>" & day(rsUpdateEntry("TravelDate1")) &"/" & month(rsUpdateEntry("TravelDate1")) &"/" & year(rsUpdateEntry("TravelDate1"))  &"</td></tr>") 
     
     response.write("<tr><td class='td-r azul'>Estatus Actual:</td><td class='td-l'><B>" & rsUpdateEntry("ESTATUS") &"</B></td></tr>") 

     response.write("<tr><td class='td-r azul'>Acci&oacute;n disponible:</td><td class='td-l'><select name='vHelpdesk' id='vHelpdesk'>")

     select case  rsUpdateEntry("helpdesk")
    
     case 1   response.write("<option value='2'>Autorizado</option>")
              response.write("<option value='10'>Cancelado</option>")

     case 2   response.write("<option value='3'>Comprobado por DFW</option>")
              response.write("<option value='10'>Cancelar solicitud</option>")
     case 3   response.write("<option value='4'>Finalizado</option>")
           

     case else response.write("<option value='"& rsUpdateEntry("helpdesk") &"'>no existen mas opciones</option>")
     end select

     response.write("</select></td></tr>")    
     response.write("<tr><td colspan=2 class='td-c'><input type='submit' value='actualizar'> </form>  </td> </tr> ")
    
     response.write("</table>")
     rsUpdateEntry.close


end sub





Sub UpReembolso  'contenido 17'

     '==============se solicto no copiar los xmls  04/feb/21 yorozco'

     'strSQL="select count(*) as calculo from xml  where reembolso=" & Session("item")
     'response.write strSQL  
     'rsUpdateEntry.Open strSQL, adoCon4

     'strSQL="select * from xml  where reembolso=" & Session("item") &" order by ID"
     'response.write strSQL  
     'rsUpdateEntry2.Open strSQL, adoCon4

     'for i=1 to rsUpdateEntry("calculo")

         'ruta="COPY d:\fileserver\wwwroot\repositorio\ADMON\Reembolsos\" &  Session("item") & "\" &  chr(34) & rsUpdateEntry2("NombreXML") &  chr(34) &" d:\fileserver\wwwroot\repositorio\ADMON\XML_Por_Procesar\"
         'strSQL="EXEC xp_cmdshell '"& ruta &"'"
         'response.write(strSQL &"<BR>")
         'rsUpdateEntry3.Open strSQL, adoCon4
         'Sleep(1) 
         'if rsUpdateEntry3.State=1 then
              'rsUpdateEntry3.close
         'end if
         'rsUpdateEntry2.MoveNext
     'next

     'rsUpdateEntry.close
     'rsUpdateEntry2.close


     select case request("vHelpdesk")
     case "2"   strSQL="UPDATE  Reembolsos SET helpdesk=2,ESTATUS='Autorizado',sentdate=null  where ID=" & Session("item")               'no se manda a outsourcing'     
     
     case "10"  strSQL="UPDATE  Reembolsos SET helpdesk=10,ESTATUS='cancelado',empresa='cancelado',Beneficiario='cancelado',Elaborador='cancelado'   where ID=" & Session("item")            
     end select
     'response.write strSQL  
     rsUpdateEntry.Open strSQL, adoCon


    response.redirect("ShowContent.asp?contenido=20&msg=La solicitud "& Session("item") &", ha sido actualizada" )
end sub  





Sub ConsolidarReembolso  '15'

   Session("ID")=request("ID")

   response.write("<P>&nbsp;</P>")
   if request("msg")<>"" then  
           response.write("<P class='alert td-c'>" & request("msg") &"</P>")
   end if
   response.write("<H1>Facturas del Reembolso No. "& request("ID") &"</H1>")
   CreateTable(960)

   Response.Write("<tr><td class='td-c azul fonttiny'>#</td><td class='td-c azul fonttiny'>UUID</td><td class='td-c azul fonttiny'>Fecha</td><td class='td-c azul fonttiny'>Tipo</td><td class='td-c azul fonttiny'>Metodo</td><td class='td-c azul fonttiny'>Forma</td><td class='td-c azul fonttiny'>?</td><td class='td-c azul fonttiny'>Moneda</td><td class='td-c azul fonttiny'>Folio</td><td class='td-c azul fonttiny'>RFC Emisor</td><td class='td-c azul fonttiny'>Emisor</td><td class='td-c azul fonttiny'>Receptor</td><td class='td-c azul fonttiny'>UsoCFDi</td><td class='td-c azul fonttiny'>Subtotal</td><td class='td-c azul fonttiny'>Impuestos</td><td class='td-c azul fonttiny'>Total</td><td class='td-c azul fonttiny'>Concepto</td></tr>")

     response.write("<form id='viaticos' method='post' action='ShowContent.asp'>")


                strSQL="select * from xml where reembolso="& request("ID") &" order by ID "
                'response.write strSQL  
                rsUpdateEntry.Open strSQL, adoCon4     

                while not rsUpdateEntry.EOF
                       Response.Write("<tr>")  %>
                        <td class='td-l fonttiny'> <%=rsUpdateEntry("ID")%></td>
                       <td class='td-l fonttiny'> <%=rsUpdateEntry("UUID")%></td>
                       <%
                       Response.Write("<td class='td-l fonttiny'>"& rsUpdateEntry("Fecha") &"</td>")
                       Response.Write("<td class='td-l fonttiny'>"& rsUpdateEntry("TipoComprobante") &"</td>")
                       Response.Write("<td class='td-l fonttiny'>"& rsUpdateEntry("MetodoPago") &"</td>")
                       Response.Write("<td class='td-l fonttiny'>"& rsUpdateEntry("FormaPago") &"</td>")        

                       response.Write("<td class='td-c fonttiny'>")

                       if rsUpdateEntry("TipoComprobante")="I" then
                        if rsUpdateEntry("MetodoPago")="PPD"  then
                          if rsUpdateEntry("FormaPago") ="99" then
                             response.write("&nbsp;&nbsp;&nbsp;<font class=alert>  ok </font>")
                            else
                                  response.write("<a href='Xml-up.asp?action=1&xml="& rsUpdateEntry("UUID") &"'><img src='/img/b_drop.png' border='0' with='11px' height='11px' alt='' title=''></a>")
                            end if 
                        else
                            if rsUpdateEntry("FormaPago") ="99" then
                                response.write("<a href='Xml-up.asp?action=1&xml="& rsUpdateEntry("UUID") &"'><img src='/img/b_drop.png' border='0' with='11px' height='11px' alt='' title=''></a>")
                            else
                                  response.write("&nbsp;&nbsp;&nbsp;<font class=alert> ok </font>")
                            end if 

                        end if
                     end if
                      response.Write("</td>") 

                       Response.Write("<td class='td-l fonttiny'>"& rsUpdateEntry("Moneda") &"</td>")
                       Response.Write("<td class='td-l fonttiny'>"& rsUpdateEntry("Folio") &"</td>")
                       Response.Write("<td class='td-l fonttiny'>"& rsUpdateEntry("RFCEmisor") &"</td>")
                       Response.Write("<td class='td-l fonttiny'>"& mid(rsUpdateEntry("Emisor"),1,45) &"</td>")
                       Response.Write("<td class='td-l fonttiny'>"& rsUpdateEntry("RFCReceptor") &"</td>")
                       Response.Write("<td class='td-l fonttiny'>"& rsUpdateEntry("UsoCFDi") &"</td>")
                       Response.Write("<td class='td-l fonttiny'>"&  FormatCurrency(rsUpdateEntry("subtotal")) &"</td>")
                       Response.Write("<td class='td-l fonttiny'>"&  FormatCurrency(rsUpdateEntry("impuestos")) &"</td>")
                       Response.Write("<td class='td-l fonttiny'>"&  FormatCurrency(rsUpdateEntry("total")) &"</td>")

                       Response.Write("<td class='td-l fonttiny'><input class='td-l fonttiny anchox2' type='text' name='concepto"& rsUpdateEntry("ID") &"' placeholder='breve descripci&oacute;n concepto'></td>")

                       Response.Write("</tr>")
                       rsUpdateEntry.movenext
                wend
                rsUpdateEntry.close
                strSQL="select isnull(sum(subtotal),0) as calculo from xml where reembolso="& request("ID") &" and TipoComprobante='I' "
                'response.write strSQL  
                rsUpdateEntry.Open strSQL, adoCon4    

                strSQL="select isnull(sum(impuestos),0) as calculo from xml where reembolso="& request("ID") &" and TipoComprobante='I' "
                'response.write strSQL  
                rsUpdateEntry2.Open strSQL, adoCon4    

                strSQL="select isnull(sum(total),0) as calculo from xml where reembolso="& request("ID") &" and TipoComprobante='I' "
                'response.write strSQL  
                rsUpdateEntry3.Open strSQL, adoCon4    

                separador

                response.write("<tr><td colspan=12 class='td-r'>&nbsp;</td><td class='td-r'>"& FormatCurrency(rsUpdateEntry("calculo")) &"</td><td class='td-r'>"& FormatCurrency(rsUpdateEntry2("calculo")) &"</td><td class='td-r'><font class=alert> "& FormatCurrency(rsUpdateEntry3("calculo")) &"</font></td></tr>")

                vTotal=rsUpdateEntry3("calculo")
                rsUpdateEntry.close
                rsUpdateEntry2.close
                rsUpdateEntry3.close
                
 

                      response.write("<tr><td colspan=14 class='td-r'><B>Ingrese el total por gastos no deducibles (sin factura): </td><td class='td-r'></B>$ <input class=anchoSmall type='text' name='nodeducibles' value='0.00'></td></tr>")

                     response.write("<input type='hidden' name='contenido' value='15'>")
                     response.write("<input type='hidden' name='control' value='4'>")
                     response.write("<input type='hidden' name='ID' value='"&  request("ID") &"'>")


               response.write("<tr><td colspan=15 class='td-r'><input type='submit' value='aplicar para reembolso'></form></td></tr>")

                response.write("</table><P>&nbsp;</P>")


end sub





sub SolicitarReembolso 'contenido 15'
     
     vNoDeducible=trim(request("nodeducibles")   )
     response.write("no deducibles=" &  vNoDeducible &"<BR>")
     vNoDeducible=replace(vNoDeducible,"$","") 
     response.write("no deducibles=" &  vNoDeducible &"<BR>")

     strSQL="select sum(total) as calculo from xml where reembolso="& request("ID") &" and TipoComprobante='I' "
     rsUpdateEntry.Open strSQL, adoCon4

     vTotal=trim(rsUpdateEntry("calculo"))
     rsUpdateEntry.close

     if vNoDeducible="" then
        vNoDeducible="0"
     end if

     flag=1
   
          strSQL="select isnull(count(*),0) as calculo from xml where reembolso="& request("ID") 
          'response.write strSQL  
          rsUpdateEntry.Open strSQL, adoCon4  

          for i=1 to rsUpdateEntry("calculo")  'hasta el numero de xml encontrados
               vstring="concepto"&i
               strSQL="UPDATE xml set Concepto='"& request(vstring) &"' where reembolso="& request("ID") &" and ID="& i
               if trim(request(vstring))="" then
                     flag=0
               end if
               response.write(strSQL & "<BR>")
               rsUpdateEntry2.Open strSQL, adoCon4
          next
          rsUpdateEntry.close


     strSQL="UPDATE reembolsos SET Total=("&  vTotal &"+" & vNoDeducible &"),helpdesk=1,Estatus='recibido',NoDeducible="& vNoDeducible &" where ID=" & request("ID") 
     response.write strSQL
     rsUpdateEntry.Open strSQL, adoCon

     if flag=1 then
          response.redirect("ShowContent.asp?contenido=20&msg=El reembolso no." & request("ID") &", ha sido recibido." ) 
     else
          response.redirect("ShowContent.asp?contenido=15&control=3&ID="& request("ID")  &"&msg=no puede dejar un concepto vacio" )  
     end if

end sub




sub  ShowHeaderTableReembolso
             if request("contenido")<>18 then 
                       response.write("<center><table width='1200px' border=1 cellpadding=2 align=center class=whiteBorders>")
              else
                       response.write("<center><table width='800px' border=1 cellpadding=2 align=center class=whiteBorders>")
            end if
             response.write("<tr><td colspan=7 class='td-c subject BoldBorders'  style='padding-top: 10px;padding-bottom: 10px;'>COMPROBACION DE REEMBOLSO No. "& Session("ID") &"</td></tr>")
            response.write("<tr><td colspan=2  class='whiteBorders'>&nbsp;</td><td class='td-r azul bold'>Beneficiario: </td><td colspan=3 class='td-l UnderBorder'>" & rsUpdateEntry("beneficiario") &"</td><td class='whiteBorders'>&nbsp;</td></tr>")
            response.write("<tr><td  colspan=2 class='whiteBorders'>&nbsp;</td><td class='td-r azul bold'>Elabor&oacute;: </td><td colspan=3 class='td-l UnderBorder'>" & rsUpdateEntry("Elaborador") &"</td><td class='whiteBorders'>&nbsp;</td></tr>")
            response.write("<tr><td  colspan=2 class='whiteBorders'>&nbsp;</td><td class='td-r azul bold'>Motivo: </td><td colspan=3 class='td-l UnderBorder'>" & rsUpdateEntry("Motivo") &"</td><td class='whiteBorders'>&nbsp;</td></tr>")
            response.write("<tr><td  colspan=2 class='whiteBorders'>&nbsp;</td><td class='td-r azul bold'>Ciudad destino: </td><td colspan=3 class='td-l UnderBorder'>" & rsUpdateEntry("CiudadDestino") &"</td><td class='whiteBorders'>&nbsp;</td></tr>")   
            response.write("<tr><td  colspan=2 class='whiteBorders'>&nbsp;</td><td class='td-r azul bold'>Facturas a reembolsar: </td><td colspan=3 class='td-l UnderBorder'>" & rsUpdateEntry("Facturas") &"</td><td class='whiteBorders'>&nbsp;</td></tr>")    
            response.write("<tr><td  colspan=2 class='whiteBorders'>&nbsp;</td><td class='td-r azul bold'>Monto comprobado: </td><td colspan=3 class='td-l UnderBorder'>" & FormatCurrency(rsUpdateEntry3("calculo")) &"</td><td class='whiteBorders'>&nbsp;</td></tr>")  
             response.write("<tr><td colspan=7 class='whiteBorders'>&nbsp;</td></tr>")
            if request("contenido")<>18 then 
            response.write("<tr><td colspan=2 class='whiteBorders'><HR></td><td colspan=3 class='whiteBorders td-c'><B>Factura "& rsUpdateEntry2("calculo")+1 &" de "& rsUpdateEntry("Facturas") &"</B></td><td class='whiteBorders' colspan=2><HR></td></tr>")
            end if

             if request("contenido")=18 then 
                   response.write("</table>")
             end if
end sub  







Sub ConfirmarReembolso  'contenido 19 '     

     strSQL="select * from Reembolsos where ID=" & Session("ID")   
     rsUpdateEntry.Open strSQL, adoCon

     ruta1="MD d:\fileserver\wwwroot\repositorio\ADMON\Reembolsos\" & Session("ID") 
     ruta2="http://192.168.0.206/repositorio/ADMON/Reembolsos/" & Session("ID") 
     ruta3="R:/ADMON/Reembolsos/" & Session("ID") & "-" 

     strSQL="UPDATE Reembolsos set Estatus='Por Autorizar',helpdesk=1,Repositorio='"&  ruta2  &"' where ID=" & Session("ID")    
     rsUpdateEntry2.Open strSQL, adoCon

     strSQL="EXEC xp_cmdshell '"& ruta1 &"'"
     'response.write strSQL
     rsUpdateEntry2.Open strSQL, adoCon

    
     cantidad=trim(rsUpdateEntry("total"))
     pos=InStr(cantidad,".")
     entero1=0
     decimales1=0
     if pos>0 then
         entero1=int(trim(mid(cantidad,1,pos-1)))
         decimales1=int(trim(mid(cantidad,pos+1,300)))
     else
         entero1=int(trim(rsUpdateEntry("total"))) 
         decimales1=0
     end if

     strSQL="select convert(decimal(18,2),sum(subtotal)) as calculo from LineReembolsos where ID=" & Session("ID")           
     rsUpdateEntry3.Open strSQL, adoCon
     cantidad=trim(rsUpdateEntry3("calculo"))
     pos=InStr(cantidad,".")
     entero2=0
     decimales2=0
     if pos>0 then
         entero2=int(trim(mid(cantidad,1,pos-1)))
         decimales2=int(trim(mid(cantidad,pos+1,300)))
     else
         entero2=rsUpdateEntry3("calculo")
         decimales2=0
     end if

     if trim(entero1)<>trim(entero2) or trim(decimales1)<>trim(decimales2) then
             cantidad=trim(rsUpdateEntry3("calculo"))
             pos=InStr(cantidad,".")
             if pos>0 then
                 session("entero")=trim(mid(cantidad,1,pos-1))
                 session("decimales")=trim(mid(cantidad,pos+1,300))
               else
                 session("entero")=rsUpdateEntry3("calculo")
                 session("decimales")="00"
             end if

             EnterosDecimales
             strSQL="UPDATE Reembolsos set Total=" & session("entero") &"."& session("decimales") &",Letras='"& session("vstring") &"',Monto1=case when Cuenta1<>'' then '"& session("entero") &"."& session("decimales")&"' else '' end,Monto2=case when Cuenta2<>'' then '"& session("entero") &"."& session("decimales")&"' else '' end  where ID="  & Session("ID")  
             'response.write strSQL
             rsUpdateEntry4.Open strSQL, adoCon             

     end if
     response.redirect("ShowContent.asp?contenido=12&msg=Se recibio su solicitud de reembolso No. " &  Session("ID") &"&ruta2=" & ruta3  )

end Sub  





sub ControlViaticos   'contenido 12'
   
    Doalert    
    titulo="CONTROL DE VIATICOS &nbsp;&nbsp;&nbsp;"

     if  Session("session_modulo") <>"viaticos" then          
          titulo=titulo & "<a href='ShowContent.asp?contenido=19&modulo=Controlviaticos'><img src='/images/login.png' width='29px' height='32px' border=0 alt='login' title='login'></a>"
     end if

     DoTitulo(titulo)
     CreateTable(1200)
     
      
       Const adCmdText = &H0001
       Const adOpenStatic = 3
       nPage=0
       registros=1

     strSQL="select year(a.docdate) as ANIO,*,convert(varchar,a.comprobar) flag,case Empresa when 'Corporativo Empresarial Izcalli SA de CV' then 'IZCALLI'  when 'DELTAFLOW DE MEXICO S DE RL DE CV' then 'DMX'  when 'Servicios Profesionales, Legales, Contables, Financieros, Recursos Humanos y Gestoria SR de CV' then 'SPL' else 'OTROS' END as empresa2,case when substring(Motivo,1,22)='Entrega de remisiones:' then 1 else 0 end as suministro  from VIATICOS_R1 a  where datediff(DAY,a.docdate,getdate() )<395  order by a.ID desc"  
     'response.write strSQL
     rsUpdateEntry.Open strSQL, adoCon, adOpenStatic, adCmdText

     rsUpdateEntry.PageSize = 20 
     nPageCount = rsUpdateEntry.PageCount

       if request("vPage")<>"" then
            nPage = int(trim(request("vPage")))
        else
            nPage=1
        end if
      rsUpdateEntry.AbsolutePage=npage
      response.write("<tr><td colspan=3 class='td-r'>PAGINAS:</td><td colspan=10>")
       for i=1 to nPageCount step 1
            if i<>npage then  
                response.write("<a href='/modules/ShowContent.asp?contenido=12&vPage="&i &"'>") 
                response.write( i &"</A>&nbsp;&nbsp;&nbsp;&nbsp;") 
            end if
        next
        response.write("</td></tr>")

       response.write("<tr>")
       response.write("<td class='subtitulo fonttiny'>ANIO</td>")
       response.write("<td class='subtitulo fonttiny'>ID</td>")
       response.write("<td class='subtitulo fonttiny'>EMPRESA</td>")
       response.write("<td class='subtitulo fonttiny'>FECHA </br>REGISTRO</td>")
       response.write("<td class='subtitulo fonttiny'>TOTAL</td>") 
       response.write("<td class='subtitulo fonttiny'>BENEFICIARIO</td>")   
          
       response.write("<td class='subtitulo fonttiny'>Entrega de remisiones: <br>[MOTIVO] </td>")     
       response.write("<td class='subtitulo fonttiny'>UNIDAD</td>")  
       response.write("<td class='subtitulo fonttiny'>ALMACEN</td>")

       'response.write("<td class='subtitulo fonttiny'>ELABORO</td>")    
       response.write("<td class='subtitulo fonttiny' width='90px'>FECHA <br>INICIO</td>")    
       response.write("<td class='subtitulo fonttiny' width='90px'>FECHA <br>FIN</td>")  
       response.write("<td class='subtitulo fonttiny'>DIAS</td>")      
       response.write("<td class='subtitulo fonttiny' width='110px'>ESTATUS</td>") 
       'response.write("<td class='subtitulo'>REPOSITORIO</td>") 
       response.write("<td class='subtitulo fonttiny'>Comprobar</td>") 
       response.write("</tr>")


     if rsUpdateEntry.EOF then
          response.write("<tr><td colspan=20 class='td-c'>no hay solicitudes de viaticos, por el momento </td></tr>")
     end if

     while not rsUpdateEntry.EOF AND registros<=20
            response.write("<tr>")
            response.write("<td class='td-c fonttiny' style='font-size: 10px'>"& rsUpdateEntry("ANIO") &"</td>")

            if Session("session_modulo") ="viaticos" and rsUpdateEntry("helpdesk")>=3  then
                  response.write("<td class='td-r'><a href='ShowContent.asp?contenido=29&ID=" &rsUpdateEntry("ID") &"'><U>"& rsUpdateEntry("ID") &"</U></a></td>")
            else
                 response.write("<td class='td-r'>"& rsUpdateEntry("ID") &"</td>")
            end if
            response.write("<td class='td-l fonttiny'>"& rsUpdateEntry("Empresa2") &"</td>")
            response.write("<td class='td-c fonttiny' style='font-size: 10px'>"& FormatDateTime(rsUpdateEntry("DocDate"),2) &"</td>")

            if Session("session_modulo") ="viaticos" OR  trim(rsUpdateEntry("ID"))=trim(request("ID")) then 
                   response.write("<td class='td-r fonttiny'><B>"& FormatCurrency(rsUpdateEntry("Total")) &"</B></td>")
            else
                    response.write("<td class='td-r fonttiny'>****</td>")
            end if

            response.write("<td class='td-l fonttiny'>"& mid(rsUpdateEntry("beneficiario"),1,20) &"</td>")

           
            vMotivo=trim(rsUpdateEntry("Motivo") )
            vMotivo=replace(vMotivo,"Entrega de remisiones:","") 
           
            if rsUpdateEntry("suministro")=1 then
                response.write("<td class='td-l  fonttiny'><a href='ShowContent.asp?contenido=100&ID="& rsUpdateEntry("ID")  &"'>"& vMotivo &"</a></td>")
            else
                response.write("<td class='td-l  fonttiny'>"& vMotivo &"</td>")
            end if

            response.write("<td class='td-l  fonttiny'>"& rsUpdateEntry("ID_unidad") &"</td>")
            response.write("<td class='td-l  fonttiny'>"& rsUpdateEntry("almacen") &"</td>")
            response.write("<td class='td-c  fonttiny'>"& rsUpdateEntry("TravelDate1") &"</td>")
            response.write("<td class='td-c  fonttiny'>"& rsUpdateEntry("TravelDate2") &"</td>")
            response.write("<td class='td-c  fonttiny'>"& rsUpdateEntry("Dias") &"</td>")

            if rsUpdateEntry("helpdesk") <>4 then      
                  if rsUpdateEntry("helpdesk") =3 then 
                         response.write("<td class='td-r  fonttiny' style='background-color: #BCFABC'>")
                  else
                         response.write("<td class='td-r fonttiny'>")
                  end if       
                  response.write("<a href='ShowContent.asp?contenido=13&modulo=viaticos&item="& rsUpdateEntry("ID") &"'><U>"& rsUpdateEntry("Estatus") &"</U></a></td>")
            else
                  response.write("<td class='td-r fonttiny' style='background-color: #BCFABC'>"& rsUpdateEntry("Estatus") & "</td>")
            end if
        
            if rsUpdateEntry("helpdesk")<3 AND trim(rsUpdateEntry("flag"))="0"   then                   
                   vRFCReceptor=rsUpdateEntry("Empresa")
                   select case vRFCReceptor
                    case "DELTAFLOW S DE RL DE CV"                   vRFCReceptor="DEL080208G25"
                    case "DELTAFLOW DE MEXICO S DE RL DE CV"         vRFCReceptor="DME140422SS7"
                    case "Corporativo Empresarial Izcalli SA de CV"  vRFCReceptor="IDA0606215S0"
                    case "Servicios Profesionales, Legales, Contables, Financieros, Recursos Humanos y Gestoria SR de CV"  vRFCReceptor="SPL130121G98"
                   end select
                   
                   response.write("<td class='td-c fonttiny'><a href='http://192.168.0.206/xmls_viatico.asp?ID="&  rsUpdateEntry("ID")  &"&Receptor=" & vRFCReceptor &"'><img src='/images/alert.gif' border=0 alt='comprobar' title='comprobar' width='14px' height='14px'></a></td>")
            else
                  response.write("<td class='td-c fonttiny'>&nbsp;</td>")
            end if
           response.write("</tr>")
           rsUpdateEntry.MoveNext
           registros=registros+1
     wend
     rsUpdateEntry.close

       response.write("</table><P>&nbsp;</P>")

       CreateTable(1200)
       response.write("<tr><td class='td-c'><img src='/images/alert.gif' border=0 alt='comprobar' title='comprobar' width='14px' height='14px'>&nbsp;&nbsp;&nbsp; Antes de iniciar el proceso de comprobaci&oacute;n, aseg&uacute;rese que todos los archivos PDF y Xml se encuentran al inicio de la carpeta repositorio del viatico</td></tr></table>")

       'se desactive este anuncio
       'if request("ruta1")<>"" then
          'response.write("<P class='alert td-c'>Por favor, coloque sus archivos de facturas (pdf/xml) en repositorio: " & request("ruta1") & "</P>")
       'end if

end sub




sub ControlReembolsos  'contenido 20'
       strSQL="delete reembolsos where helpdesk=0"   
       rsUpdateEntry.Open strSQL, adoCon

       if request("msg")<>"" then
          response.write("<P class='alert td-c'>" & request("msg") & "</P>")
       end if

       titulo="CONTROL DE REEMBOLSOS &nbsp;&nbsp;&nbsp;"
    
     if  Session("session_modulo") <>"reembolsos" then          
          titulo=titulo &"<a href='ShowContent.asp?contenido=19&modulo=ControlReembolsos'><img src='/images/login.png' width='29px' height='32px' border=0 alt='login' title='login'></a>"     
     end if

        DoTitulo(titulo)
        CreateTable(1200)        
     
        Const adCmdText = &H0001
        Const adOpenStatic = 3
        nPage=0
        registros=1

        strSQL="select year(docdate) as ANIO,* from Reembolsos where helpdesk>=1 order by ID desc"  
        rsUpdateEntry.Open strSQL, adoCon, adOpenStatic, adCmdText

        rsUpdateEntry.PageSize = 25 
        nPageCount = rsUpdateEntry.PageCount

        if request("vPage")<>"" then
            nPage = int(trim(request("vPage")))
        else
            nPage=1
        end if      
        rsUpdateEntry.AbsolutePage=npage
        response.write("<tr><td colspan=3 class='td-r'>PAGINAS:</td><td colspan=8>")
        for i=1 to nPageCount step 1
            if i<>npage then  
                response.write("<a href='/modules/ShowContent.asp?contenido=20&vPage="&i &"'>") 
                response.write( i &"</A>&nbsp;&nbsp;&nbsp;&nbsp;") 
            end if
        next
        response.write("</td></tr>")

       
       response.write("<tr>")
       response.write("<td class=subtitulo>ANIO</td>")
       response.write("<td class=subtitulo>ID</td>")
       response.write("<td class=subtitulo>FECHA </br>REGISTRO</td>")
       response.write("<td class=subtitulo>TOTAL</td>") 
       response.write("<td class=subtitulo>EMPRESA</td>")
       response.write("<td class=subtitulo>BENEFICIARIO</td>")   
       'response.write("<td class=subtitulo>CUENTA <br> BANCO</td>")      
       response.write("<td class=subtitulo>MOTIVO</td>")     
       response.write("<td class=subtitulo>ELABORO</td>")    
     
       response.write("<td class=subtitulo>ESTATUS</td>") 
       'response.write("<td class=subtitulo>REPOSITORIO</td>") 
       response.write("</tr>")


        while not rsUpdateEntry.EOF AND registros<=25
            response.write("<tr>")
            response.write("<td class='td-c fonttiny'>"& rsUpdateEntry("ANIO") &"</td>")
            if Session("session_modulo") ="reembolsos"  then
                  response.write("<td class='td-r'><a href='ShowContent.asp?contenido=18&ID=" &rsUpdateEntry("ID") &"'><U>"& rsUpdateEntry("ID") &"</U></a></td>")
            else
                 response.write("<td class='td-r'>"& rsUpdateEntry("ID") &"</td>")
            end if
            response.write("<td class='td-c fonttiny'>"& FormatDateTime(rsUpdateEntry("DocDate"),2) &"</td>")
            if Session("session_modulo") ="reembolsos" OR  trim(rsUpdateEntry("ID"))=trim(request("ID")) then 
                   response.write("<td class='td-r fonttiny'>"& FormatCurrency(rsUpdateEntry("Total")) &"</td>")
            else
                   response.write("<td class='td-r'>****</td>")
            end if
            response.write("<td class='td-l fonttiny'>"& mid(rsUpdateEntry("Empresa"),1,30) &"</td>")
            response.write("<td class='td-l fonttiny'>"& rsUpdateEntry("beneficiario") &"</td>")

            'if rsUpdateEntry("Cuenta1") <>"" then 
                  'response.write("<td class='td-l fonttiny'>"& rsUpdateEntry("Cuenta1") &"</td>")
            'else
                  'response.write("<td class='td-l fonttiny'>"& rsUpdateEntry("Cuenta2") &"</td>")
            'end if
            response.write("<td class='td-l fonttiny'>"& mid(rsUpdateEntry("Motivo"),1,30) &"</td>")
            response.write("<td class='td-l fonttiny'>"& rsUpdateEntry("Elaborador") &"</td>")
                       
            if rsUpdateEntry("helpdesk")=1 and Session("modulo")="reembolsos"  then  
                  response.write("<td class='td-c fonttiny'><a class='fonttiny' href='ShowContent.asp?contenido=16&modulo=reembolsos&item="& rsUpdateEntry("ID") &"'><U>"& rsUpdateEntry("ESTATUS") &"</U></A></td>")  
            else
                if rsUpdateEntry("helpdesk")=2 then
                  response.write("<td class='td-c fonttiny' style='background-color: #BCFABC'>"& rsUpdateEntry("ESTATUS") &"</td>")  
                else
                  response.write("<td class='td-c fonttiny'>"& rsUpdateEntry("ESTATUS") &"</td>")  
                end if
            end if
            'response.write("<td class='td-l fonttiny'><B>R:\ADMON\Reembolsos\"& rsUpdateEntry("ID") &"</B></td>")
            rsUpdateEntry.MoveNext
            registros=registros+1
       wend
       rsUpdateEntry.close

       response.write("</table>")
       

end sub  



sub DetalleReembolso   '18'
   'request("ID")'
    strSQL="select * from xml where reembolso="  & request("ID")
    rsUpdateEntry.Open strSQL, adoCon4

    strSQL="select isnull(sum(total),0) as calculo from xml where reembolso="  & request("ID")
    rsUpdateEntry2.Open strSQL, adoCon4

    strSQL="select isnull(NoDeducible,0) as calculo from Reembolsos where ID="  & request("ID")
    rsUpdateEntry3.Open strSQL, adoCon

   response.write("<P>&nbsp;</P><H1>Detalle Reembolso No. "&  request("ID") &"</H1>")
   CreateTable(1440)
   %>
   <tr><td  class="td-c whiteBorders azul">ID</td><td  class="td-c whiteBorders azul" width='9%'>Concepto</td><td  class="td-c whiteBorders azul">UUID</td><td  class="td-c whiteBorders azul">Fecha</td><td  class="td-c whiteBorders azul">RFC Emisor</td><td  class="td-c whiteBorders azul">Emisor</td><td  class="td-c whiteBorders azul">RFC Receptor</td><td  class="td-c whiteBorders azul">Metodo</td>
                   <td  class="td-c whiteBorders azul">Forma</td><td  class="td-c whiteBorders azul">Moneda</td><td  class="td-c whiteBorders azul">Folio</td><td  class="td-c whiteBorders azul">Subtotal</td><td  class="td-c whiteBorders azul">Impuestos</td><td  class="td-c whiteBorders azul">Total</td><td  class="td-c whiteBorders azul">Orden Compra</td></tr>
   <%
    while not rsUpdateEntry.EOF
      response.write("<tr>")
      response.write("<td class='td-l fonttiny'>"& rsUpdateEntry("ID") & "</td>")
      response.write("<td class='td-l fonttiny'>"& mid(rsUpdateEntry("concepto"),1,30) & "</td>")
      response.write("<td class='td-l fonttiny'>"& rsUpdateEntry("UUID") & "</td>")
      response.write("<td class='td-c fonttiny'>"& FormatDateTime(rsUpdateEntry("Fecha"),2)  & "</td>")
      response.write("<td class='td-l fonttiny'>"& rsUpdateEntry("RFCEmisor") & "</td>")
      response.write("<td class='td-l fonttiny'>"& mid(rsUpdateEntry("EMisor"),1,40) & "</td>")
      response.write("<td class='td-c fonttiny'>"& rsUpdateEntry("RFCReceptor") & "</td>")
      response.write("<td class='td-c fonttiny'>"& rsUpdateEntry("MetodoPago") & "</td>")
      response.write("<td class='td-c fonttiny'>"& rsUpdateEntry("FormaPago") & "</td>")
      response.write("<td class='td-c fonttiny'>"& rsUpdateEntry("Moneda") & "</td>")
      response.write("<td class='td-r fonttiny'>"& rsUpdateEntry("Folio") & "</td>")
      response.write("<td class='td-r fonttiny'>"& FormatCurrency(rsUpdateEntry("subtotal")) & "</td>")
      response.write("<td class='td-r fonttiny'>"& FormatCurrency(rsUpdateEntry("impuestos")) & "</td>")
      response.write("<td class='td-r fonttiny'>"& FormatCurrency(rsUpdateEntry("total")) & "</td>")
      response.write("<td class='td-r fonttiny'>")

      strSQL="select * from OPOR where U_CFDi_TimbreUUID='"  & rsUpdateEntry("UUID") &"' and CANCELED='N' "

      if rsUpdateEntry("RFCReceptor")="DME140422SS7" then
         vString="DMX-"
         rsUpdateEntry4.Open strSQL, adoCon2 'DMX'
      end if
  
      if rsUpdateEntry("RFCReceptor")="DEL080208G25" then
         vString="DFW-"
         rsUpdateEntry4.Open strSQL, adoCon3 'DFW'
      end if

      if rsUpdateEntry4.State=1 then     
           if not rsUpdateEntry4.EOF  then              
               response.write (vString & rsUpdateEntry4("Docnum") )
           end if
           rsUpdateEntry4.close
      end if

      response.write("</td></tr>")      
    
      rsUpdateEntry.movenext
    wend
    rsUpdateEntry.close    

    separador
    response.write("<tr><td colspan=11 class='td-r'>Subotal:<td><td class='td-r'><B>" &       FormatCurrency(rsUpdateEntry2("calculo"))  & "</B></td></tr>")
    response.write("<tr><td colspan=11 class='td-r'>No deducible:<td><td class='td-r'><B>" &  FormatCurrency(rsUpdateEntry3("calculo"))  & "</B></td></tr>")

    strSQL="select ("& rsUpdateEntry2("calculo") &"+" &rsUpdateEntry3("calculo") &") as calculo"
    rsUpdateEntry4.Open strSQL, adoCon

    response.write("<tr><td colspan=11 class='td-r'>TOTAL: <td><td class='td-r'><B>" &   FormatCurrency(rsUpdateEntry4("calculo"))   & "</B></td></tr>")
   
    rsUpdateEntry2.close
    rsUpdateEntry3.close
    rsUpdateEntry4.close
end sub





sub DetalleViatico   'contenido 29'  
   'request("ID")'
    strSQL="select * from xml where viatico="  & request("ID") &" order by ID"
    'response.write(strSQL&"<BR>")
    rsUpdateEntry.Open strSQL, adoCon4

    strSQL="select isnull(sum(total),0) as calculo from xml where viatico="  & request("ID")
    'response.write(strSQL&"<BR>")
    rsUpdateEntry2.Open strSQL, adoCon4

    strSQL="select isnull(NoDeducible,0) as calculo from VIATICOS_R1 where ID="  & request("ID")
    'response.write(strSQL&"<BR>")
    rsUpdateEntry3.Open strSQL, adoCon

    strSQL="SELECT STUFF( (SELECT ','+convert(varchar,Remision) from ViaticoRemision where ID="& request("ID")&" FOR XML PATH (''))   , 1, 1, '')  as remisiones"
    'response.write strSQL
    rsUpdateEntry4.Open strSQL, adoCon

   titulo="Detalle Viatico No. "&  request("ID") &" (remisiones: "&rsUpdateEntry4("remisiones")&")"
   rsUpdateEntry4.close
   doTitulo(titulo)
   CreateTable(1360)
   %>
   <tr><td  class="td-c whiteBorders azul">ID</td><td  class="td-c whiteBorders azul" width='9%'>Concepto</td><td  class="td-c whiteBorders azul">UUID</td><td  class="td-c whiteBorders azul">Fecha</td><td  class="td-c whiteBorders azul">RFC Emisor</td><td  class="td-c whiteBorders azul">Emisor</td><td  class="td-c whiteBorders azul">RFC Receptor</td><td  class="td-c whiteBorders azul">Metodo Pago</td>
                   <td  class="td-c whiteBorders azul">Forma Pago</td><td  class="td-c whiteBorders azul">Moneda</td><td  class="td-c whiteBorders azul">Folio</td><td  class="td-c whiteBorders azul">Subtotal</td><td  class="td-c whiteBorders azul">Impuestos</td><td  class="td-c whiteBorders azul">Total</td><td  class="td-c whiteBorders azul">Orden Compra</td></tr>
   <%
    while not rsUpdateEntry.EOF
      response.write("<tr>")
      response.write("<td class='td-l fonttiny'>"& rsUpdateEntry("ID") & "</td>")
      response.write("<td class='td-l fonttiny'>"& mid(rsUpdateEntry("concepto"),1,30) & "</td>")
      response.write("<td class='td-l fonttiny'>"& rsUpdateEntry("UUID") & "</td>")
      response.write("<td class='td-c fonttiny'>"& FormatDateTime(rsUpdateEntry("Fecha"),2)  & "</td>")
      response.write("<td class='td-l fonttiny'>"& rsUpdateEntry("RFCEmisor") & "</td>")
      response.write("<td class='td-l fonttiny'>"& mid(rsUpdateEntry("EMisor"),1,40) & "</td>")
      response.write("<td class='td-c fonttiny'>"& rsUpdateEntry("RFCReceptor") & "</td>")
      response.write("<td class='td-c fonttiny'>"& rsUpdateEntry("MetodoPago") & "</td>")
      response.write("<td class='td-c fonttiny'>"& rsUpdateEntry("FormaPago") & "</td>")
      response.write("<td class='td-c fonttiny'>"& rsUpdateEntry("Moneda") & "</td>")
      response.write("<td class='td-r fonttiny'>"& rsUpdateEntry("Folio") & "</td>")
      response.write("<td class='td-r fonttiny'>"& FormatCurrency(rsUpdateEntry("subtotal")) & "</td>")
      response.write("<td class='td-r fonttiny'>"& FormatCurrency(rsUpdateEntry("impuestos")) & "</td>")
      response.write("<td class='td-r fonttiny'>"& FormatCurrency(rsUpdateEntry("total")) & "</td>")
      response.write("<td class='td-r fonttiny'>")

       strSQL="select * from OPOR where U_CFDi_TimbreUUID='"  & rsUpdateEntry("UUID") &"' and CANCELED='N' "

      if rsUpdateEntry("RFCReceptor")="DME140422SS7" then
         vString="DMX-"
         rsUpdateEntry4.Open strSQL, adoCon2 'DMX'
      end if
  
       if rsUpdateEntry("RFCReceptor")="DEL080208G25" then
         vString="DFW-"
         rsUpdateEntry4.Open strSQL, adoCon3 'DFW'
      end if

      if rsUpdateEntry4.State=1 then        
          if not rsUpdateEntry4.EOF  then            
               response.write (vString & rsUpdateEntry4("Docnum") )
          end if
           rsUpdateEntry4.close
      end if

      response.write("</td></tr>")    

      
      response.write("</tr>")  
      rsUpdateEntry.movenext
    wend
    rsUpdateEntry.close    
   
    separador 
    response.write("<tr><td colspan=11 class='td-r'>Subotal:<td><td class='td-r'><B>" &       FormatCurrency(rsUpdateEntry2("calculo"))  & "</B></td></tr>")
    response.write("<tr><td colspan=11 class='td-r'>No deducible:<td><td class='td-r'><B>" &  FormatCurrency(rsUpdateEntry3("calculo"))  & "</B></td></tr>")

    strSQL="select ("& rsUpdateEntry2("calculo") &"+" &rsUpdateEntry3("calculo") &") as calculo"
    rsUpdateEntry4.Open strSQL, adoCon

    response.write("<tr><td colspan=11 class='td-r'>TOTAL: <td><td class='td-r'><B>" &   FormatCurrency(rsUpdateEntry4("calculo"))   & "</B></td></tr>")
   
    rsUpdateEntry2.close
    rsUpdateEntry3.close
    rsUpdateEntry4.close
end sub







Sub  ComprobarViatico 'contenido 21'

  Response.Expires = -1

  Session("ID")=request("ID")

   response.write("<P>&nbsp;</P><H1>COMPROBACION DEL VIATICO No. "&  Session("ID") &"</H1>")
   CreateTable(960) 

   Response.Write("<tr><td class='td-c azul fonttiny'>#</td><td class='td-c azul fonttiny'>UUID</td><td class='td-c azul fonttiny'>Fecha</td><td class='td-c azul fonttiny'>Tipo</td><td class='td-c azul fonttiny'>Metodo</td><td class='td-c azul fonttiny'>Forma</td><td class='td-c azul fonttiny'>?</td><td class='td-c azul fonttiny'>Moneda</td><td class='td-c azul fonttiny'>Folio</td><td class='td-c azul fonttiny'>RFC Emisor</td><td class='td-c azul fonttiny'>Emisor</td><td class='td-c azul fonttiny'>Receptor</td><td class='td-c azul fonttiny'>UsoCFDi</td><td class='td-c azul fonttiny'>Subtotal</td><td class='td-c azul fonttiny'>Impuestos</td><td class='td-c azul fonttiny'>Total</td><td class='td-c azul fonttiny'>Concepto</td></tr>")

                strSQL="select * from xml where viatico="& request("ID") & " order by ID"
                'response.write strSQL  
                rsUpdateEntry.Open strSQL, adoCon4    
                i=1 


     response.write("<form id='viaticos' method='post' action='ShowContent.asp'>")


          If  rsUpdateEntry.EOF then
               Response.Write("<tr><td colspan=20  class='td-c'>No ha colocado ningun xml</td></tr>")

          else
                while not rsUpdateEntry.EOF
                       Response.Write("<tr>")  
                         Response.Write("<td class='td-c fonttiny'>"& i &"</td>")
                       %>
                       <td class='td-l fonttiny'> <%=rsUpdateEntry("UUID")%></td>
                       <%
                       Response.Write("<td class='td-l fonttiny'>"& rsUpdateEntry("Fecha") &"</td>")
                       Response.Write("<td class='td-l fonttiny'>"& rsUpdateEntry("TipoComprobante") &"</td>")
                       Response.Write("<td class='td-l fonttiny'>"& rsUpdateEntry("MetodoPago") &"</td>")
                       Response.Write("<td class='td-l fonttiny'>"& rsUpdateEntry("FormaPago") &"</td>")        

                       response.Write("<td class='td-c fonttiny'>")

                       if rsUpdateEntry("TipoComprobante")="I" then
                        if rsUpdateEntry("MetodoPago")="PPD"  then
                          if rsUpdateEntry("FormaPago") ="99" then
                             response.write("&nbsp;&nbsp;&nbsp;<font class=alert>ok</font>")
                            else
                                  response.write("<a href='Xml-up.asp?action=1&xml="& rsUpdateEntry("UUID") &"'><img src='/img/b_drop.png' border='0' with='11px' height='11px' alt='' title=''></a>")
                            end if 
                        else
                            if rsUpdateEntry("FormaPago") ="99" then
                                response.write("<a href='Xml-up.asp?action=1&xml="& rsUpdateEntry("UUID") &"'><img src='/img/b_drop.png' border='0' with='11px' height='11px' alt='' title=''></a>")
                            else
                                  response.write("&nbsp;&nbsp;&nbsp;<font class=alert>ok</font>")
                            end if 

                        end if
                     end if
                      response.Write("</td>") 

                       Response.Write("<td class='td-l fonttiny'>"& rsUpdateEntry("Moneda") &"</td>")
                       Response.Write("<td class='td-l fonttiny'>"& rsUpdateEntry("Folio") &"</td>")
                       Response.Write("<td class='td-l fonttiny'>"& rsUpdateEntry("RFCEmisor") &"</td>")
                       Response.Write("<td class='td-l fonttiny'>"& mid(rsUpdateEntry("Emisor"),1,50) &"</td>")
                       Response.Write("<td class='td-l fonttiny'>"& rsUpdateEntry("RFCReceptor") &"</td>")
                       Response.Write("<td class='td-l fonttiny'>"& rsUpdateEntry("UsoCFDi") &"</td>")
                       Response.Write("<td class='td-r fonttiny'>"&  FormatCurrency(rsUpdateEntry("subtotal")) &"</td>")
                       Response.Write("<td class='td-r fonttiny'>"&  FormatCurrency(rsUpdateEntry("impuestos")) &"</td>")
                       Response.Write("<td class='td-r fonttiny'>"&  FormatCurrency(rsUpdateEntry("total")) &"</td>")                     
                       Response.Write("<td class='td-l fonttiny' ><input type='text' name='concepto"& rsUpdateEntry("ID") &"' class='td-l anchox1 fonttiny'></td>")
                       Response.Write("</tr>")

                       rsUpdateEntry.movenext
                       i=i+1
                wend
          end if
          rsUpdateEntry.close

                strSQL="select isnull(sum(subtotal),0) as calculo from xml where viatico="& request("ID") &" and TipoComprobante='I' "
                'response.write strSQL  
                rsUpdateEntry.Open strSQL, adoCon4    

                strSQL="select isnull(sum(impuestos),0) as calculo from xml where viatico="& request("ID") &" and TipoComprobante='I' "
                'response.write strSQL  
                rsUpdateEntry2.Open strSQL, adoCon4    

                strSQL="select isnull(sum(total),0) as calculo from xml where viatico="& request("ID") &" and TipoComprobante='I' "
                'response.write strSQL  
                rsUpdateEntry3.Open strSQL, adoCon4    

                separador
                response.write("<tr><td colspan=13 class='td-r'>&nbsp;</td><td class='td-r'>"& FormatCurrency(rsUpdateEntry("calculo")) &"</td>")
                response.write("<td class='td-r'>" & FormatCurrency(rsUpdateEntry2("calculo")) &"</td>")
                response.write("<td class='td-r'><font class=alert> "& FormatCurrency(rsUpdateEntry3("calculo")) &"</font></td></tr>")

                vTotal=rsUpdateEntry3("calculo")
                rsUpdateEntry.close
                rsUpdateEntry2.close
                rsUpdateEntry3.close
                
  

                     response.write("<tr><td colspan=15 class='td-r'><B>Ingrese el total por gastos no deducibles (sin factura): </td><td class='td-r'></B>$ <input class=anchoSmall type='text' name='nodeducibles' value='0.00'></td></tr>")


                     response.write("<input type='hidden' name='contenido' value='22'>")                   
                     response.write("<input type='hidden' name='ID' value="&  request("ID") &">")


               response.write("<tr><td colspan=16 class='td-r'><input type='submit' value='aplicar para comprobacion'></form></td></tr>")

                response.write("</table><P>&nbsp;</P>")


                
end sub




Sub ComprobarViaticoUP   '22'
         Response.Expires = -1
         vCalculo=0

         vNoDeducible=trim(request("nodeducibles")   )         
         vNoDeducible=replace(vNoDeducible,"$","") 
         response.write("no deducibles=" &  vNoDeducible &"<BR>")

         if vNoDeducible="" then
              vNoDeducible="0"
         end if

          strSQL="select isnull(count(*),0) as calculo from xml where viatico="& request("ID") 
          rsUpdateEntry.Open strSQL, adoCon4   
          response.write(strSQL &" calculo=" & rsUpdateEntry("calculo")  &"<BR>") 

          vCalculo=rsUpdateEntry("calculo")
          'if vCalculo>25 then vCalculo=25 end if
          rsUpdateEntry.close

           for i=1 to vCalculo
               vstring="concepto"&i
               strSQL="UPDATE xml set Concepto='"& request(vstring) &"' where viatico="& request("ID") &" and ID="& i
               response.write(strSQL&"<BR>")
               rsUpdateEntry2.Open strSQL, adoCon4
          next
          
          strSQL="UPDATE VIATICOS_R1 SET comprobar=1,NoDeducible="& vNoDeducible &" where ID=" & request("ID") 
          response.write(strSQL&"<BR>")
          rsUpdateEntry.Open strSQL, adoCon

          '-- PRORRATEA VIATICOS nuevamente'
          strSQL="exec SP_Prorratea_viaticos"
          rsUpdateEntry3.Open strSQL, adoCon

          strSQL="Select top 1 RS from ViaticoRemision where ID=" & request("ID")
          rsUpdateEntry2.Open strSQL, adoCon
          if not rsUpdateEntry2.EOF  then
               vRS=rsUpdateEntry2("RS")
          end if
          rsUpdateEntry2.close

          strSQL="Select distinct(Pedido) from ViaticoRemision where ID=" & request("ID")
          rsUpdateEntry2.Open strSQL, adoCon

          if not rsUpdateEntry2.EOF then
          	  while not rsUpdateEntry2.EOF 

          	  	    strSQL="exec [dbo].[SP_AnalisisUtilidad] "& rsUpdateEntry2("Pedido") &",'"& vRS &"','"& session("session_id") &"'"
                    rsUpdateEntry4.Open strSQL, adoCon4
                    
                   rsUpdateEntry2.movenext
          	  wend            
          end if
          rsUpdateEntry2.close

          response.redirect("ShowContent.asp?contenido=12&msg=Administracion ha recibido la comprobacion del viatico no." & request("ID")  )  

end sub




sub territorios
    response.write("<P align='center'><img src='/images/territorios2.png' border=0></P>")

end sub



sub ShowAddEmployer   'contenido 9'
      'response.write("modulo=" & request("modulo") &" accion="  & request("accion") )
      strSQL = "SELECT max(ID)+1 as calculo from Empleados"
      'response.write strSQL
      rsUpdateEntry.Open strSQL, adoCon

      response.write("<P>&nbsp;</P><form id='add' name='add' action='/modules/colabor-up.asp' method='post'>")  
      CreateTable(800)   
      response.write("<tr><td colspan=5 class=subtitulo>Agregar un nuevo colaborador</td></tr>")
      
      response.write("<tr><td class=td-r width='18%'><B>ID:</B></td><td class=td-l width='32%'><input class=input1 type='text' id='vid' name='vid' maxlength=3 value='"& rsUpdateEntry("calculo") &"'></td><td class=td-c>&nbsp;</td> ")
      rsUpdateEntry.close

      response.write("<td class=td-r width='18%'><B>Nombres:</B></td><td class=td-l width='32%'><input class=input1 type='text' id='vnombres' name='vnombres'></td></tr>")
      
      response.write("<tr><td class=td-r ><B>Paterno:</B></td><td class=td-l><input class=input1 type='text' id='vpaterno' name='vpaterno'></td><td class=td-c>&nbsp;</td> ")
      response.write("<td class=td-r><B>Materno:</B></td><td class=td-l><input class=input1 type='text' id='vmaterno' name='vmaterno'></td></tr>")
      
      response.write("<tr><td class=td-r ><B>Fecha Nacimiento:</B></td><td class=td-l>")  %>
      <input name="vnacimiento" type="date"/>    <%
      
      response.write("</td><td class=td-c>&nbsp;</td>")  
      response.write("<td class=td-r><B>Fecha Ingreso:</B></td><td class=td-l>")  %>
      <input name="vingreso" type="date"/>    <%
      
      response.write("</td></tr>")
      
      response.write("<tr><td class=td-r ><B>CURP:</B></td><td class=td-l><input class=input1 type='text' id='vcurp' name='vcurp' maxlength=18 placeholder='sin espacios'></td><td class=td-c>&nbsp;</td> ")
      response.write("<td class=td-r><B>RFC:</B></td><td class=td-l><input class=input1 type='text' id='vrfc' name='vrfc' maxlength=13></td></tr>")
      
      response.write("<tr><td class=td-r ><B>NSS:</B></td><td class=td-l><input class='input1' type='number' id='vnss' name='vnss' maxlength=11 placeholder='seguridad social'></td><td class=td-c>&nbsp;</td> ")   
      response.write("<td class=td-r><B>Celular Empresa:</B></td><td class=td-l><input class=input1 type='number' id='vcelular' name='vcelular' maxlength=10 placeholder='10 digitos'>10 d&iacute;gitos</td></tr>")
      
      response.write("<tr><td class=td-r ><B>SEXO:</B></td><td class=td-l><select name=vsexo><option value='F'>F</option><option value='M'>M</option></select></td><td class=td-c>&nbsp;</td> ")
      response.write("<td class=td-r><B>email Empresa:</B></td><td class=td-l><input class=input1 type='text' id='vemail' name='vemail' style='width:200px' maxlength=40 placeholder='sin espacios'></td></tr>")

      response.write("<tr><td class=td-r><B>email personal:</B></td><td class=td-l><input class=input1 type='text' id='vemailP' name='vemailP' style='width:200px' maxlength=40 placeholder='sin espacios'></td><td class=td-c>&nbsp;</td>")
       response.write("<td class=td-r><B>Celular personal:</B></td><td class=td-l><input class=input1 type='number' id='vcelularP' name='vcelularP' maxlength=10 placeholder='10 digitos'>10 d&iacute;gitos</td></tr> ")
       
      
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
      response.write("</select> </td><td class=td-c>&nbsp;</td> <td class=td-r> <B> Cuenta Bancaria: </B></td><td class=td-l><input class='input1 ancho' type='text' id='vBanco' name='vBanco' maxlength=18></td> </tr>")
      response.write("<tr><td class=td-r ><B>Id usuario:</B></td><td class='td-l azul'><input class='ancho' type='text' id='vusuario' name='vusuario'><td>&nbsp;</td><td class=td-r ><B>Contrase&ntilde;a:</B></td><td class='td-l azul'><input class='ancho' type='text' id='vclave' name='vclave'></tr>")
      response.write("<tr><td class=td-c colspan=5>&nbsp;</td></tr>")
      response.write("<tr><td class=td-c colspan=5><input type='submit' value='agregar'></form></td></tr></table>")



end sub


sub ShowDelEmployer    'contenido 9'
      'response.write( "accion=" & Session("accion") & ",item=" &  Session("item") )
      if request("item") <>"" then
             Session("item") = request("item") 
      end if

      strSQL="SELECT * from Empleados where ID="& Session("item") 
      'response.write strSQL 
      rsUpdateEntry.Open strSQL, adoCon
      
      response.write("<form id='add' name='add' action='/modules/colabor-up.asp' method='post'>")               
      response.write("<P>&nbsp;</P><center><table width='650px' border=0 cellpadding=2 cellspacing=1 align=center><tr><td colspan=5 class=subtitulo>Dar de baja colaborador: <B>"&  Session("item") &"</td></tr>")

      response.write("<tr><td class=td-r width='50%'><B>Nombres:</B></td><td class=td-l width='50%'>"& rsUpdateEntry("nombres")& "</td></tr>")      
      response.write("<tr><td class=td-r ><B>Paterno:</B></td><td class=td-l>"& rsUpdateEntry("paterno")& "</td><td class=td-c>&nbsp;</td> ")
      response.write("<tr><td class=td-r ><B>Paterno:</B></td><td class=td-l>"& rsUpdateEntry("materno")& "</td><td class=td-c>&nbsp;</td> ")

      response.write("<tr><td class=td-r><B>Fecha de Egreso:</B></td><td class=td-l>")  %>
      <input name="vegreso" type="date"/>    <%      
      response.write("</td></tr>")
      response.write("<tr><td class=td-r ><B>Metodo de baja:</B></td><td class=td-l><select name='vFormaBaja'><option value='finiquito' SELECTED>Finiquito</option><option value='liquidacion'>Liquidacion</option><td class=td-c>&nbsp;</td> ")
      
      response.write("<tr><td class=td-c colspan=5><input type='submit' value='dar de baja'></form></td></tr></table>")
      rsUpdateEntry.close

end sub




sub ShowEditEmployer    'contenido 40 edit'


   if trim(request("item")) <>"" then
             Session("item") = request("item")              
   end if


    strSQL="SELECT * from Empleados where ID="& Session("item") 
    'response.write strSQL 
    rsUpdateEntry.Open strSQL, adoCon

      
    response.write("<P>&nbsp;</P><form id='add' name='add' action='/modules/colabor-up.asp' method='post'>")   
    
      response.write("<center><table width='1000px' border=1 cellpadding=2 cellspacing=1 align=center><tr><td colspan=5 class=subtitulo>Editar colaborador no. "& Session("item")  &"</td></tr>")
      
      response.write("<tr><td class=td-r width='18%'><B>ID:</B></td><td class=td-l width='32%'><B>"& Session("item")  &"</B></td><td class=td-c>&nbsp;</td> ")
      response.write("<td class=td-r width='18%'><B>Nombres:</B></td><td class=td-l width='32%'><input class=input1 type='text' id='vnombres' name='vnombres' value='"& rsUpdateEntry("nombres")&"'></td></tr>")
      
      response.write("<tr><td class=td-r ><B>Paterno:</B></td><td class=td-l><input class=input1 type='text' id='vpaterno' name='vpaterno' value='"& rsUpdateEntry("paterno")&"'></td><td class=td-c>&nbsp;</td> ")
      response.write("<td class=td-r><B>Materno:</B></td><td class=td-l><input class=input1 type='text' id='vmaterno' name='vmaterno' value='"& rsUpdateEntry("materno")&"'></td></tr>")
      
    
      
      response.write("<tr><td class=td-r ><B>CURP:</B></td><td class=td-l><input class=input1 type='text' id='vcurp' name='vcurp' maxlength=18 value='"& rsUpdateEntry("CURP")&"'></td><td class=td-c>&nbsp;</td> ")
      response.write("<td class=td-r><B>RFC:</B></td><td class=td-l><input class=input1 type='text' id='vrfc' name='vrfc' maxlength=13 value='"& rsUpdateEntry("RFC")&"'></td></tr>")
      
      response.write("<tr><td class=td-r ><B>NSS:</B></td><td class=td-l><input class=input1 type='text' id='vnss' name='vnss' maxlength=11 value='"& rsUpdateEntry("NSS")&"'></td><td class=td-c>&nbsp;</td> ") 
      response.write("<td class=td-r><B>Celular Personal:</B></td><td class=td-l><input class=input1 type='text' id='vcelularP' name='vcelularP' maxlength=10 value='"& rsUpdateEntry("CELULAR_PERSONAL")&"'>10 d&iacute;gitos</td></tr>")
           
      response.write("<tr><td class=td-r ><B>SEXO:</B></td><td class=td-l><select name=vsexo>")
      if rsUpdateEntry("sexo")="M" then
            response.write("<option value='F'>F</option><option value='M' SELECTED>M</option></select></td><td class=td-c>&nbsp;</td> ")
      else
            response.write("<option value='F' SELECTED>F</option><option value='M'>M</option></select></td><td class=td-c>&nbsp;</td> ")
      end if

      response.write("<td class=td-r><B>email personal:</B></td><td class=td-l><input class=input1 type='text' id='vemailP' name='vemailP' style='width:200px' maxlength=40 value='"& rsUpdateEntry("email_Personal")&"'></td></tr>")
      
       response.write("<tr><td class=td-r ><B>celular Empresa:</B></td><td class=td-l><input class=input1 type='text' id='vcelular' name='vcelular' maxlength=10 value='"& rsUpdateEntry("celular")&"'></td><td class=td-c>&nbsp;</td> ") 
      response.write("<td class=td-r><B>email Empresa:</B></td><td class=td-l><input class=input1 type='text' id='vemail' name='vemail' style='width:200px' value='"& rsUpdateEntry("email") &"' maxlength=40 ></td></tr>")

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
      response.write("</select> </td><td class=td-c>&nbsp;</td> <td class=td-r> <B> CLABE: </B></td><td class=td-l><input class='input1 ancho' type='text' id='vBanco' name='vBanco' maxlength=18 value='"& rsUpdateEntry("CLABE") &"'></td> </tr>")

      if  rsUpdateEntry("id_usuario")<>"" then
            strSQL="SELECT * from cat_usuario where id_usuario='"& rsUpdateEntry("id_usuario") &"'"
            'response.write strSQL 
            rsUpdateEntry2.Open strSQL, adoCon
            if  not rsUpdateEntry2.EOF then
                 vclave=rsUpdateEntry2("clave")
            end if
            rsUpdateEntry2.close
      else
           vclave=""
      end if
    
       response.write("<tr><td class=td-r ><B>Id usuario:</B></td><td class='td-l azul'><input class='ancho' type='text' id='vusuario' name='vusuario' value='"& rsUpdateEntry("id_usuario") &"'><td>&nbsp;</td><td class=td-r ><B>Banco:</B></td><td class='td-l azul'><select name='fbanco'><option value=''>sin asignar</option>")

           strSQL="SELECT * from cat_bancos"
            'response.write strSQL 
            rsUpdateEntry2.Open strSQL, adoCon

          while not rsUpdateEntry2.EOF
              if rsUpdateEntry2("clave_banco")=rsUpdateEntry("clave_banco") then
                   response.write("<option value='"& rsUpdateEntry2("clave_banco") &"' SELECTED>"& rsUpdateEntry2("banco") &"</option>")
              else
                   response.write("<option value='"& rsUpdateEntry2("clave_banco") &"'>"& rsUpdateEntry2("banco") &"</option>")
              end if
              rsUpdateEntry2.MoveNext
          wend
          rsUpdateEntry2.close

       response.write("</select></td></tr>")

      

       response.write("<tr><td class=td-r ><B>Contrase&ntilde;a:</B></td><td class='td-l azul'><input class='ancho' type='text' id='vclave' name='vclave' value='"& vclave &"'><td>&nbsp;</td><td class=td-r><B>Fecha Nacimiento</B></td><td class='td-l'>Dia: <select name=VDia>")

                
               for i=1 to 31
                    if day(rsUpdateEntry("FechaNacimiento"))=i then
                       response.write("<option value="& i &" SELECTED>"& i &"</option>")
                   else
                       response.write("<option value="& i &" >"& i &"</option>")
                   end if
               next
               response.write("</select> Mes: <select name=vMes>")
               for i=1 to 12
                    if month(rsUpdateEntry("FechaNacimiento"))=i then
                       response.write("<option value="& i &" SELECTED>"& meses(i) &"</option>")
                   else
                       response.write("<option value="& i &" >"&  meses(i) &"</option>")
                   end if
               next
               response.write("</select> A&ntilde;o: <select name=vAnio>")
               for i=year(date())-70 to year(date())-16
                   if year(rsUpdateEntry("FechaNacimiento"))=i then
                       response.write("<option value="& i &" SELECTED>"& i &"</option>")
                   else
                       response.write("<option value="& i &" >"& i &"</option>")
                   end if
               next
               response.write("</select></tr>")

      strSQL="select * from Empresas"
      rsUpdateEntry3.Open strSQL, adoCon

      response.write("<tr><td class='td-r'><B>Empresa:</B></td>")
      response.write("<td class='td-l'><select name='fempresa'>")

      while not rsUpdateEntry3.EOF
      	   if rsUpdateEntry3("id_empresa")=rsUpdateEntry("id_empresa") then
                response.write("<option value='"& rsUpdateEntry3("id_empresa")  &"' SELECTED>"& rsUpdateEntry3("id_empresa") &"</option>")
      	   else      	   
                response.write("<option value='"& rsUpdateEntry3("id_empresa")  &"'>"& rsUpdateEntry3("id_empresa") &"</option>")
           end if
           rsUpdateEntry3.MoveNext
      wend
      rsUpdateEntry3.close


      response.write("</select></td><td>&nbsp;</td> <td class=td-r ><B>ID Checador:</B></td> <td class='td-l'><input class='ancho' type='int' id='vchecador' name='vchecador' value='"& rsUpdateEntry("id_checador") &"'></td>   </tr>")


      response.write("<tr><td class=td-r ><B>&nbsp;</B></td> <td class='td-l'>&nbsp;</td><td class='td-l'>&nbsp;</td><td class=td-r ><B>Contrase&ntilde;a tienda en linea:</B></td><td class='td-l'><input class='anchox2' type='text' name='fPassTienda' value='"& rsUpdateEntry("ContrasenaTiendaLinea") &"'></td></tr>")

      response.write("<tr><td class=td-c colspan=5>&nbsp;</td></tr>")

    response.write("<tr><td class=td-c colspan=5><input type='submit' value='actualizar'></form></td></tr></table>")
    rsUpdateEntry.close



end sub


sub CrearTablaAsesor
	   response.write("<table width='900px' border=0 cellpadding=3 cellspacing=2 align=center>")
       response.write("<tr>")
       'response.write("<td class=subtitulo>Asesor</td>")
       response.write("<td class=subtitulo>Factura</td>")
       response.write("<td class=subtitulo>D&iacute;a</td>")  
       response.write("<td class=subtitulo>Cuenta</td>")      
       response.write("<td class=subtitulo>Subtotal</td>")  
       response.write("<td class=subtitulo>RazonSocial</td>")  
       response.write("<td class=subtitulo>Tipo</td>") 
       response.write("<td class=subtitulo>Pedido</td>")
       response.write("<td class=subtitulo>Proyecto</td>")
       response.write("<td class=subtitulo>Operacion</td>")
       response.write("</tr>")
end sub


Sub GraficoAsesor   	%>

    <script type="text/javascript">      
      var chartData=[]; 
      chartData.push({date: new Date(2021, 0, 31, 0, 0, 0, 0), val: document.getElementById('alicia1').value });
      chartData.push({date: new Date(2021, 1, 28, 0, 0, 0, 0), val: document.getElementById('alicia2').value });
      chartData.push({date: new Date(2021, 2, 31, 0, 0, 0, 0), val: document.getElementById('alicia3').value });
      chartData.push({date: new Date(2021, 3, 30, 0, 0, 0, 0), val: document.getElementById('alicia4').value });
      chartData.push({date: new Date(2021, 4, 31, 0, 0, 0, 0), val: document.getElementById('alicia5').value });
      chartData.push({date: new Date(2021, 5, 30, 0, 0, 0, 0), val: document.getElementById('alicia6').value });
      chartData.push({date: new Date(2021, 6, 31, 0, 0, 0, 0), val: document.getElementById('alicia7').value });
      chartData.push({date: new Date(2021, 7, 31, 0, 0, 0, 0), val: document.getElementById('alicia8').value });
      chartData.push({date: new Date(2021, 8, 30, 0, 0, 0, 0), val: document.getElementById('alicia9').value });
      chartData.push({date: new Date(2021, 9, 31, 0, 0, 0, 0), val: document.getElementById('alicia10').value });
      chartData.push({date: new Date(2021, 10, 30, 0, 0, 0, 0), val: document.getElementById('alicia11').value });
      chartData.push({date: new Date(2021, 11, 31, 0, 0, 0, 0), val: document.getElementById('alicia12').value });
           
                      
      AmCharts.ready(function() {
          var chart = new AmCharts.AmStockChart();
          chart.pathToImages = "/amcharts/images/";

          var dataSet = new AmCharts.DataSet();
          dataSet.dataProvider = chartData;
          dataSet.fieldMappings = [{fromField:"val", toField:"value"}];   
          dataSet.categoryField = "date";          
          chart.dataSets = [dataSet];

          var stockPanel = new AmCharts.StockPanel();
          chart.panels = [stockPanel];

          var panelsSettings = new AmCharts.PanelsSettings();
          panelsSettings.startDuration = 1;
          chart.panelsSettings = panelsSettings;   

          var graph = new AmCharts.StockGraph();
          graph.valueField = "value";
          graph.type = "column";
          graph.fillAlphas = 1;
          graph.title = "MyGraph"; 
          stockPanel.addStockGraph(graph);

          chart.write('alicia');
      });        
          </script>    


	    <%	
end sub


 Sub ShowAvanceVentas       'contenido 6' 

     if request("fmes")<>"" then
     	   vmes=request("fmes")
           vanio=request("fanio")
           Titulo="AVANCE MENSUAL DE VENTAS: "& meses(vmes) &", "&vanio
           strSQL="select * from HistoricoVentas" & vanio & " where month(DocDate)="& vmes &" order by asesor,DocDate desc"	 
     else
           vmes=month(date())
           vanio=year(date())

	     Titulo="AVANCE MENSUAL DE VENTAS: "& meses(vmes) &", "& vanio &"<BR style='font-size:3px;height:3px'><form id='ventas' method='POST' action='ShowContent.asp'><input type='hidden' name='contenido' value='6'><select name='fmes'>"

	     for i=1 to 12 step 1
	     	  if month(date())=i then
	          Titulo=Titulo&"<option value='"&i&"' SELECTED>"&meses(i)&"</option>"
	        else
                Titulo=Titulo&"<option value='"&i&"'>"&meses(i)&"</option>"
	        end if
	     next
	     
	     Titulo=Titulo&"</select>&nbsp;&nbsp;&nbsp;<select name='fanio'>"
	     for i=year(date()) to year(date())-2 step -1
	        Titulo=Titulo&"<option value='"&i&"'>"& i &"</option>"
	     next
	     Titulo=Titulo&"</select>&nbsp;&nbsp;<input type='submit' value='mostrar'></form>"

	     strSQL="select * from HistoricoVentas" & year(date()) & " where month(DocDate)=month(getdate()) order by asesor,DocDate desc"	 	     
	end if

     DoAlert
     Dotitulo(Titulo)  
     'response.write strSQL  
     flag=1
     rsUpdateEntry.Open strSQL, adoCon4
     if rsUpdateEntry.EOF then flag=0
       
       response.write("<table width='1300px' border=1 cellpadding=3 cellspacing=2 align=center>")
       response.write("<tr><td width='900px'>&nbsp;</td><td width='10px'>&nbsp;</td>")
       response.write("<td width='390px'>&nbsp;</td></tr>")
       response.write("<tr><td class='td-c' style='vertical-align:top'>")


       vtotal=round(0,2)
       vasesor=""

       if flag=0 then NoRegistros
       inicio=1
        

        while not rsUpdateEntry.EOF 
                if inicio=1 then 
                	CrearTablaAsesor
                	response.write("<tr><td colspan=9 class='td-c strong'>" & rsUpdateEntry("asesor") &"</td><tr>")
                end if
                vasesor=rsUpdateEntry("asesor")
                vtotal=vtotal+round( rsUpdateEntry("subtotal"),2)

                if vasesor="Alicia Mayorquin" or vasesor="Martin Gibrann Parra Leal" or vasesor="Reynaldo Cardona" then
                	strSQL="select month(docdate) as mes,cast(sum(subtotal) as int) as subtotal  from HistoricoVentas2021 where asesor='"&vasesor&"' group by month(docdate) order by month(docdate)"
                	rsUpdateEntry6.Open strSQL, adoCon4
                	select case vasesor
                	case "Alicia Mayorquin"
                	    for z=1 to 12 
                	    	response.write("<input type='hidden' id='alicia"&z&"' value="& trim(rsUpdateEntry6("subTotal")) &">")
                	    	rsUpdateEntry6.MoveNext
                	    next	
                	case "Martin Gibrann Parra Leal"
                	        response.write("<input type='hidden' id='martin"&z&"' value="& trim(rsUpdateEntry6("subTotal")) &">")
                	    	rsUpdateEntry6.MoveNext
                	case "Reynaldo Cardona"
                	        response.write("<input type='hidden' id='rey"&z&"' value="& trim(rsUpdateEntry6("subTotal")) &">")
                	    	rsUpdateEntry6.MoveNext
                    end select
                    rsUpdateEntry6.close
                end if
      
                response.write("<tr>")  
                'response.write("<td class='td-l fonttiny'>"& rsUpdateEntry("asesor") &"</td>")
                if trim(rsUpdateEntry("Tipo"))="Ingresos" then
                response.write("<td class='td-r fonttiny'><a href='ShowContent.asp?contenido=119&action=1&fRS="&rsUpdateEntry("RazonSocial")&"&ffactura="&rsUpdateEntry("Docnum")&"' target='factura'>"& rsUpdateEntry("Docnum") &"</a></td>")
                else
                response.write("<td class='td-r fonttiny'>"& rsUpdateEntry("Docnum") &"</td>")

                end if
                response.write("<td class='td-c fonttiny'>"& day(rsUpdateEntry("DocDate")) &"</td>")
                response.write("<td class='td-l fonttiny'>"& mid(rsUpdateEntry("Cardname"),1,30) &"</td>")

                if trim(rsUpdateEntry("Tipo"))="NotaCreditos" then
                     response.write("<td class='td-r fonttiny' style='color:red'>"& FormatCurrency(rsUpdateEntry("subtotal")) &"</td>")                     
                else
                     response.write("<td class='td-r fonttiny' style='color:green'> + "& FormatCurrency(rsUpdateEntry("subtotal")) &"</td>")  
                end if
                
                response.write("<td class='td-c fonttiny'>"& rsUpdateEntry("RazonSocial") &"</td>")
                response.write("<td class='fonttiny'>"& rsUpdateEntry("Tipo") &"</td>")
                response.write("<td class='td-c fonttiny'><a href='ShowContent.asp?contenido=68&action=1&fRS="&rsUpdateEntry("RazonSocial")&"&fpedido="&rsUpdateEntry("Pedido")&"' target='pedido'>"& rsUpdateEntry("Pedido") &"</a></td>")
                response.write("<td class='td-l fonttiny'>"& mid(rsUpdateEntry("U_PROYECTO"),1,10) &"</td>")

                if rsUpdateEntry("U_CFDi_MetodoPago")="PPD" then
                       response.write("<td class='td-c fonttiny'>CREDITO</td>")
                else
                       response.write("<td class='td-c alert fonttiny'>CONTADO</td>")
                end if
                
                rsUpdateEntry.MoveNext
                inicio=inicio+1

                if not rsUpdateEntry.EOF then
                      if vasesor<> rsUpdateEntry("asesor") then
                          response.write("<tr><td colspan=10 class=td-c style='font-size: 4px'><HR></td></tr>")
                          strSQL="select isnull(U_MetaMensual,1) as Meta from OSLP where SlpName='"& vasesor&"' "
                          'response.write strSQL
                          rsUpdateEntry6.Open strSQL, adoCon2
                          response.write("<tr><td colspan=2>&nbsp;</td><td class='td-l strong fontsmall'> "& vasesor &" | META: <font color=red>"& rsUpdateEntry6("Meta") &"</font></td><td class='td-r fontsmall'> <B>"& FormatCurrency(vtotal) &"</B></td>")
                          vMeta=trim(rsUpdateEntry6("Meta"))
                          rsUpdateEntry6.close

                          'if day(date())=1 then   
                              'strSQL="select round( cast( (select sum(subTotal) as calculo from HistoricoVentas" & year(date()) & " where month(DocDate)=month(getdate())-1 and asesor='"& vasesor &"' and U_CFDi_MetodoPago='PUE' ) as float )/ cast( (select sum(subTotal) as calculo from HistoricoVentas" & year(date()) & " where month(DocDate)=month(getdate())-1 and asesor='"& vasesor &"'  ) as float ) * 100 ,1 ) as calculo"
                          'else
                              strSQL="select round( cast( (select sum(subTotal) as calculo from HistoricoVentas" & year(date()) & " where month(DocDate)=month(getdate()) and asesor='"& vasesor &"' and U_CFDi_MetodoPago='PUE' ) as float )/ cast( (select case when sum(subTotal)=0 then 1 else sum(subTotal) end  as calculo from HistoricoVentas" & year(date()) & " where month(DocDate)=month(getdate()) and asesor='"& vasesor &"'  ) as float ) * 100 ,1 ) as calculo"

                          'end if
                          'response.write("<td colspan=2 style='text-align: right'> "& strSQL &"</td>")
                          'response.write strSQL
                          rsUpdateEntry7.Open strSQL, adoCon4

                          strSQL="select round(cast("&vTotal&" as float)/cast("&vMeta&" as float)*100,1) as calculo"
                           rsUpdateEntry5.Open strSQL, adoCon4
                   
                          if int(vMeta)>1 then
                              response.write("<td colspan=2 class='td-r strong fontsmall'>AVANCE: "& rsUpdateEntry5("calculo") &" % </td>")
                          else
                              response.write("<td colspan=2 class='td-r strong fontsmall'>&nbsp;</td>")
                          end if
                          rsUpdateEntry5.close

                          response.write("<td colspan=3 class='td-c strong fontsmall'> "& rsUpdateEntry7("calculo") &" % vendido al contado</td>")

                          rsUpdateEntry7.close
                          vtotal=round(0,2)
                          response.write("<tr><td colspan=10 class=td-c style='font-size: 4px'><HR></td></tr>")
                          response.write("</table>")
                          response.write("<td>&nbsp;</td>")
                          response.write("<td class='td-c' style='vertical-align:top'>")
                          select case vasesor
                          	  case "Alicia Mayorquin"     %>
                          	      <div id="alicia" style="width:340px; height:220px;"></div> <%                       
                          	  case "Martin Gibrann Parra Leal"  %>
                          	      <div id="martin" style="width:340px; height:220px;"></div> <%               
                          	  case "Reynaldo Cardona"        %>   
                          	      <div id="rey" style="width:340px; height:220px;"></div>    <%
                          	  case else response.write("&nbsp;")
                          end select

                          GraficoAsesor
     
                          response.write("</td></tr>")
                          response.write("<tr><td colspan=2 style='height:30px'>&nbsp;</td></tr>")
                          response.write("<tr><td class='td-c' style='vertical-align:top'>")
                          inicio=1
                      end if
                else
                          response.write("<tr><td colspan=10 class=td-c style='font-size: 4px'><HR></td></tr>")
                           strSQL="select isnull(U_MetaMensual,1) as Meta from OSLP where SlpName='"& vasesor&"' "
                          'response.write strSQL
                          rsUpdateEntry6.Open strSQL, adoCon2
                          response.write("<tr><td colspan=2>&nbsp;</td><td class='td-l strong'> "& vasesor &" | META MENSUAL: <font color=red>"& rsUpdateEntry6("Meta") &"</font></td><td class='td-r'> <B>"& FormatCurrency(vtotal) &"</B></td>")
                           vMeta=trim(rsUpdateEntry6("Meta"))
                          rsUpdateEntry6.close
                        
                          if day(date())=1 then
                               strSQL="select round( cast( (select sum(subTotal) as calculo from HistoricoVentas" & year(date()) & " where month(DocDate)=month(getdate())-1 and asesor='"& vasesor &"' and U_CFDi_MetodoPago='PUE' ) as float )/ cast( (select sum(subTotal) as calculo from HistoricoVentas" & year(date()) & " where month(DocDate)=month(getdate())-1 and asesor='"& vasesor &"'  ) as float ) * 100 ,1 ) as calculo"
                          else
                               strSQL="select case when (select sum(subTotal) as calculo from HistoricoVentas" & year(date()) & " where month(DocDate)=month(getdate()) and asesor='"& vasesor &"' )>0 then round( cast( (select sum(subTotal) as calculo  from HistoricoVentas" & year(date()) & " where month(DocDate)=month(getdate()) and asesor='"& vasesor &"' and U_CFDi_MetodoPago='PUE' ) as float )/   cast( (select sum(subTotal) as calculo from HistoricoVentas" & year(date()) & " where month(DocDate)=month(getdate()) and asesor='"& vasesor &"' ) as float ) * 100 ,1 ) else 0 end    as calculo"

                               'strSQL="select round( cast( (select sum(subTotal) as calculo from HistoricoVentas" & year(date()) & " where month(DocDate)=month(getdate()) and asesor='"& vasesor &"' and U_CFDi_MetodoPago='PUE' ) as float )/ cast(  (select sum(subTotal) as calculo from HistoricoVentas" & year(date()) & " where month(DocDate)=month(getdate()) and asesor='"& vasesor &"'  ) as float ) * 100 ,1 ) as calculo"

                          end if
                          'response.write("<td colspan=2 style='text-align: right'> "& strSQL &"</td>")
                          rsUpdateEntry7.Open strSQL, adoCon4

                          strSQL="select round(cast("&vTotal&" as float)/cast("&vMeta&" as float)*100,1) as calculo"
                           rsUpdateEntry5.Open strSQL, adoCon4
                   
                          if int(vMeta)>1 then
                              response.write("<td colspan=2 class='td-r strong'>AVANCE: "& rsUpdateEntry5("calculo") &" % </td>")
                          else
                              response.write("<td colspan=2 class='td-r strong'>&nbsp;</td>")
                          end if
                          rsUpdateEntry5.close

                          response.write("<td colspan=3 class='td-c strong'> "& rsUpdateEntry7("calculo") &" % vendido al contado</td>")

                          rsUpdateEntry7.close
                          vtotal=round(0,2)
                          response.write("<tr><td colspan=10 class=td-c style='font-size: 4px'><HR></td></tr></table>")
                          response.write("<td>&nbsp;</td>")
                          response.write("<td class='td-c' style='vertical-align:top'>ultimo cuadro</td></tr>")

                end if
        wend
       rsUpdateEntry.close  
  
      if flag=1 then
		      strSQL="select sum(subTotal) as calculo from HistoricoVentas" & vanio & " where month(DocDate)="&vmes
		      strSQL2="select round( cast( (select sum(subTotal) as calculo from HistoricoVentas" & vanio & " where month(DocDate)="& vmes &"  and U_CFDi_MetodoPago='PUE' ) as float )/ cast( (select sum(subTotal) as calculo from HistoricoVentas" & vanio & " where month(DocDate)="& vmes &"   ) as float ) * 100 ,1 ) as calculo"

		      'response.write(strSQL &"<br>")
		      'response.write(strSQL2 &"<br>")

		      rsUpdateEntry2.Open strSQL, adoCon4
		      rsUpdateEntry7.Open strSQL2, adoCon4


		       response.write("<tr><td colspan=2>&nbsp;</td></tr>")
		      
		        
		                   

		           response.write("<tr><td class='td-r strong' style='padding-right:20px' colspan=2>GRAN TOTAL: "& FormatCurrency(rsUpdateEntry2("calculo")) &"</td><td class='td-l strong' style='padding-left:20px'> "& rsUpdateEntry7("calculo") &" % vendido al contado</td></tr>")
		           rsUpdateEntry2.close

		       response.write("</table><P>&nbsp;</P>") 
		       'response.write("</table><P class='td-c'><a href='ShowContent.asp?contenido=6&control=-1'>Mostrar avance del mes anterior</a></P>") 
		       response.write("<P>&nbsp;</P>")  
		       %> <P><button onclick="go_display_div('#titulo')"> ir arriba   </button> </P> <%
	   end if
       response.write("<P>&nbsp;</P>") 

end sub




 



sub CrearRuta    'contenido 8'
     response.write("<form id='CrearRuta' action='ShowContent.asp' method='post'><input type='hidden' id='fcontenido' name='fcontenido' value='8'>")

     titulo="CREAR RUTA EN REPOSITORIO PARA NUEVO SN"
     DoTitulo(titulo)
     response.write("<H3>El proposito de este formulario es crear una ruta de almacenamiento de archivos para un nuevo socio de negocio,<BR> actividad que debe ser realizado en una sola ocasion para un socio de negocio reci&eacute;n creado </H3> ")

     strSQL="select a.cardcode,a.RazonSocial from Notificaciones a where a.tipo='pedido' and a.modulo='ventas' and ( a.CardCode like 'C0%' OR a.CardCode='L000001' ) and a.RazonSocial is not null and a.CardCode not in (select cardcode from Repositorio where RazonSocial=a.RazonSocial)"  
 
     'response.write strSQL
     rsUpdateEntry.Open strSQL, adoCon4

if trim(request("fpath"))="" then

     if rsUpdateEntry.EOF then
         response.write("<BR><font class='alert'>Por el momento, no hay ning&uacute;n Socio de negocio sin ruta de repositorio</font>")
         
     else
         if rsUpdateEntry("RazonSocial")="MME" then
            vempresa="MEIDE"
         else
            vEmpresa=rsUpdateEntry("RazonSocial")
         end if
         strSQL="select replace(replace(cardname,' ',''),'.','') as cardname from PRODUCTIVA_"& vempresa &".dbo.OCRD where cardcode='" & rsUpdateEntry("cardcode") &"'" 
         'response.write strSQL
         rsUpdateEntry2.Open strSQL, adoCon4              
        
         response.write("<P class=alert2>Sugerencia: al crear rutas de almacenamiento de archivos, utilice nombres cortos sin espacios y sin simbolos. Utilice el siguiente cuadro de texto para modificar la ruta que va hacer creada: </P>")
         response.write("<input style='width: 500px' type='text' id='fpath' name='fpath' value='"& trim(rsUpdateEntry("cardcode")) &"-" & trim(rsUpdateEntry2("cardname")) &"'>")
         response.write("<input type='hidden' id='msg' name='msg' value='Se ha creado la ruta de repositorio'>")
         response.write("<input type='hidden' id='CardCode' name='CardCode' value='"& rsUpdateEntry("CardCode") &"'>")
         response.write("<input type='hidden' id='RS' name='RS' value='"& rsUpdateEntry("RazonSocial") &"'>")
         response.write("<input type='submit' value='Crear Ruta'></H1></center></form>")
         rsUpdateEntry2.close
        
     end if
     rsUpdateEntry.close

else
    
     if request("msg")<>"" then
        response.write("<P class=alert>"& request("msg") &"</P>") 
     end if
      vRuta=trim( request("fpath") )
      vRuta=replace(vRuta,"'","")
      vRuta=replace(vRuta,"&","")
      vRuta=replace(vRuta,"%","")
      vRuta=replace(vRuta,"\","")
      vRuta=replace(vRuta,"/","")
      vRuta=replace(vRuta,"*","")
      vRuta=replace(vRuta,",","")
      vRuta=replace(vRuta,".","")


      strSQL="insert into Repositorio values ('"& request("CardCode")  &"','"& vRuta  &"','"& request("RS")  &"',getdate() )"
      'response.write( strSQL &"<BR>")
      rsUpdateEntry2.Open strSQL, adoCon4


end if

    

end sub




sub ShowCargas   '23'

       response.write("<form id='Cargas1' action='ShowContent.asp' method='post'><input type='hidden' id='fcontenido' name='fcontenido' value='24'>")

        response.write("<center><H1><B>CARGAR XML EN SAP</B> </H1><P class='td-c'>&nbsp;</P> <P class='subtitulo fonttiny'>El proposito de este formulario es crear un espacio en repositorio de archivos,<BR> donde se puedan colocar una serie de xml que seran utilizados para crear ordenes de compra (de servicio) en SAP <BR> y posteriormente poder provisionar. Si esta listo para cargar xml a SAP de click en el boton continuar... </P> ")

        response.write("<P class='td-c'><input type='submit' value='continuar'></form></P>")


 end sub





sub CrearRutaCargas   '24'


     strSQL="insert into CargasNimloader (DocDate,Helpdesk,Estatus) values (getdate(),0,'no confirmado') "
     'response.write(strSQL)
     rsUpdateEntry.Open strSQL, adoCon4


     strSQL="select max(ID) as calculo from CargasNimloader"
     rsUpdateEntry.Open strSQL, adoCon4

     vID=""
     vID=rsUpdateEntry("calculo")
     rsUpdateEntry.close

     ruta1="MD d:\fileserver\wwwroot\repositorio\ADMON\Cargas\" & vID
     ruta2="R:/ADMON/Cargas/" & vID
     

     strSQL="update CargasNimloader set Repositorio='" & ruta2 &"' where ID=" & vID
     rsUpdateEntry.Open strSQL, adoCon4

     strSQL="EXEC xp_cmdshell '"& ruta1 &"'"
     'response.write strSQL
     rsUpdateEntry.Open strSQL, adoCon4
   



    response.write("<P>&nbsp;</P><P>&nbsp;</P><P class='td-c fonttiny'>Por favor, coloque el total de los xmls a cargar en SAP en la ubicaci&oacute;n <font class=alert>"& ruta2 &"</font>,<BR> cuando haya FINALIZADO la copia de xmls, de click el siguiente link <a href='http://192.168.0.206/nimloader.asp?ID="& vID &"'> <BR> <U> para continuar...   </U> </a> </td></tr>")      


 end sub





sub XmlCargaBrowser   '25' 

      response.write("<div id='main'>")
    
       if request("msg")<>"" then
          response.write("<P class='td-c alert'>"& request("msg")&"</P>")   
       end if
  
           response.write("<div id='data'> &nbsp; </div>")


          response.write("<P>&nbsp;</P><H1>Carga de xml No. "& request("ID") &"</H1><table border=0 cellspacing=2 cellpadding=2 align=center width='94%'>")

               response.write("<form id='XmlBrowser' action='ShowContent.asp' method='POST'>") 
               response.write("<input type='hidden' name='contenido' value='26'>")
               response.write("<input type='hidden' name='ID' value='"& request("ID") &"'>")


               Response.Write("<tr><td class='td-c azul fonttiny'>UUID</td><td class='td-c azul fonttiny'>Fecha</td><td class='td-c azul fonttiny'>Tipo</td><td class='td-c azul fonttiny'>Metodo</td><td class='td-c azul fonttiny'>Forma</td><td class='td-c azul fonttiny'>?</td><td class='td-c azul fonttiny'>Moneda</td><td class='td-c azul fonttiny'>Folio</td><td class='td-c azul fonttiny'>RFC Emisor</td><td class='td-c azul fonttiny'>Emisor</td><td class='td-c azul fonttiny'>Cliente SAP</td><td class='td-c azul fonttiny'><B>Procesar</B></td><td class='td-c azul fonttiny'>Receptor</td><td class='td-c azul fonttiny'>UsoCFDi</td><td class='td-c azul fonttiny'>Subtotal</td><td class='td-c azul fonttiny'>Impuestos</td><td class='td-c azul fonttiny'>Total</td></tr>")

                strSQL="select * from xml where carga="& request("ID") &" order by ID"
                'response.write strSQL  
                rsUpdateEntry.Open strSQL, adoCon4     

                while not rsUpdateEntry.EOF
                       Response.Write("<tr>")  %>
                       <td class='td-l fonttiny'><A href="Javascript:sendReq('/modules/xmlDetail.asp','UUID','<%=rsUpdateEntry("UUID")%>','data')"> <%=rsUpdateEntry("UUID")%></A></td>
                       <%
                       Response.Write("<td class='td-l fonttiny'>"& rsUpdateEntry("Fecha") &"</td>")
                       Response.Write("<td class='td-l fonttiny'>"& rsUpdateEntry("TipoComprobante") &"</td>")
                       Response.Write("<td class='td-l fonttiny'>"& rsUpdateEntry("MetodoPago") &"</td>")
                       Response.Write("<td class='td-l fonttiny'>"& rsUpdateEntry("FormaPago") &"</td>")        

                       response.Write("<td class='td-c fonttiny'>")
                       bandera=1

                       if rsUpdateEntry("TipoComprobante")="I" then
                        if rsUpdateEntry("MetodoPago")="PPD"  then
                          if rsUpdateEntry("FormaPago") ="99" then
                             response.write("&nbsp;&nbsp;&nbsp;<font class=alert>  ok </font>")
                            else
                                  response.write("<a href='Xml-up.asp?action=1&xml="& rsUpdateEntry("UUID") &"'><img src='/img/b_drop.png' border='0' with='11px' height='11px' alt='' title=''></a>")
                                  bandera=0
                            end if 
                        else
                            if rsUpdateEntry("FormaPago") ="99" then
                                response.write("<a href='Xml-up.asp?action=1&xml="& rsUpdateEntry("UUID") &"'><img src='/img/b_drop.png' border='0' with='11px' height='11px' alt='' title=''></a>")
                                bandera=0
                            else
                                  response.write("&nbsp;&nbsp;&nbsp;<font class=alert> ok </font>")
                            end if 

                        end if
                     end if
                      response.Write("</td>") 

                       Response.Write("<td class='td-l fonttiny'>"& rsUpdateEntry("Moneda") &"</td>")
                       Response.Write("<td class='td-l fonttiny'>"& rsUpdateEntry("Folio") &"</td>")
                       Response.Write("<td class='td-l fonttiny'>"& rsUpdateEntry("RFCEmisor") &"</td>")
                       Response.Write("<td class='td-l fonttiny'>"& mid(rsUpdateEntry("Emisor"),1,60) &"</td>")
                       Response.Write("<td class='td-l fonttiny'>")

                       strSQL="select * from OCRD where LicTradNum='" &  rsUpdateEntry("RFCEmisor")  &"'"
                       flag=0
                       'response.write strSQL  

                      if  rsUpdateEntry("RFCReceptor") ="DEL080208G25" then 'DFW'
                        rsUpdateEntry2.Open strSQL, adoCon3
                        if not rsUpdateEntry2.EOF then
                            response.write("<font class=alert>" & rsUpdateEntry2("CardCode") &"</font>")
                            flag=1
                        else
                              response.write("&nbsp;")
                        end if
                      else
                        rsUpdateEntry2.Open strSQL, adoCon2
                         if not rsUpdateEntry2.EOF then
                            response.write("<font class=alert>" &   rsUpdateEntry2("CardCode") &"</font>")
                            flag=1
                        else
                              response.write("&nbsp;")
                        end if
                     end if
                      rsUpdateEntry2.close

                       Response.Write("</td>")
                       response.Write("<td class='td-c fonttiny'>")
                       if (flag=1 and bandera=1) then
                          response.write("<input style='td-c fonttiny' type='checkbox' name='box"& rsUpdateEntry("ID") &"' checked>") 
                       else
                          response.write("<input style='td-c fonttiny' type='checkbox' name='box"& rsUpdateEntry("ID") &"' >") 
                       end if
                       Response.Write("</td>")


                       Response.Write("<td class='td-l fonttiny'>"& rsUpdateEntry("RFCReceptor") &"</td>")
                       Response.Write("<td class='td-l fonttiny'>"& rsUpdateEntry("UsoCFDi") &"</td>")
                       Response.Write("<td class='td-l fonttiny'>"&  FormatCurrency(rsUpdateEntry("subtotal")) &"</td>")
                       Response.Write("<td class='td-l fonttiny'>"&  FormatCurrency(rsUpdateEntry("impuestos")) &"</td>")
                       Response.Write("<td class='td-l fonttiny'>"&  FormatCurrency(rsUpdateEntry("total")) &"</td>")

                       Response.Write("</tr>")
                       rsUpdateEntry.movenext
                wend
                rsUpdateEntry.close

                response.write("<tr><td colspan=20 class='td-r'><input type='submit' value='cargar a SAP xml seleccionados'></form>")
                response.write("</table><P>&nbsp;</P>")





      response.write("</div>")
end sub




sub XmlCargarProcesar  '26'
         
     response.write( "<P>&nbsp; </P><table width='450px' cellpadding=2 cellspacing=2 align='center'>")

     strSQL="select count(*) as calculo from xml where carga=" &  request("ID")
     rsUpdateEntry.Open strSQL, adoCon4
     total=rsUpdateEntry("calculo")
     rsUpdateEntry.close

     strSQL="select * from xml where carga=" &  request("ID") &" order by ID"
     rsUpdateEntry2.Open strSQL, adoCon4        

     ruta="MD d:\fileserver\wwwroot\repositorio\ADMON\Cargas\" & request("ID") & "\xml_cargados"
     strSQL="EXEC xp_cmdshell '"& ruta &"'"
     'response.write strSQL
     rsUpdateEntry.Open strSQL, adoCon4


     for i=1 to  total
         vstring="box" & trim(rsUpdateEntry2("ID"))
         if rsUpdateEntry3.State=1 then
                    rsUpdateEntry3.close
         end if
         if rsUpdateEntry4.State=1 then
                    rsUpdateEntry4.close
         end if

         if request(vstring)="true" or  request(vstring)="on" then
               strSQL="update xml set cargado=1 where carga=" & request("ID")  &" and  NombreXML='" & rsUpdateEntry2("NombreXML") &"'"
               rsUpdateEntry3.Open strSQL, adoCon4

               ruta="COPY d:\fileserver\wwwroot\repositorio\ADMON\cargas\" &  request("ID") & "\" &  chr(34) & rsUpdateEntry2("NombreXML") &  chr(34) &" d:\fileserver\wwwroot\repositorio\ADMON\XML_Por_Procesar\"
               strSQL="EXEC xp_cmdshell '"& ruta &"'"
               'response.write( "<tr><td class='td-l azul fonttiny'>" & rsUpdateEntry2("ID") &"</td><td class='td-l fonttiny'>" & rsUpdateEntry2("NombreXML") &"</td></tr>")         
               rsUpdateEntry3.Open strSQL, adoCon4

               ruta="MOVE d:\fileserver\wwwroot\repositorio\ADMON\cargas\" &  request("ID") & "\" &  chr(34) & rsUpdateEntry2("NombreXML") &  chr(34) & " d:\fileserver\wwwroot\repositorio\ADMON\cargas\" & request("ID") & "\xml_cargados\"
               strSQL="EXEC xp_cmdshell '"& ruta &"'"                      
               rsUpdateEntry4.Open strSQL, adoCon4


               Sleep(1) 
               'response.write( vstring & "=" & request(vstring) & "<BR>" )
         end if

         rsUpdateEntry2.MoveNext
     next
         if rsUpdateEntry.State=1 then
                    rsUpdateEntry.close
         end if
           if rsUpdateEntry2.State=1 then
                    rsUpdateEntry2.close
         end if

         if rsUpdateEntry3.State=1 then
                    rsUpdateEntry3.close
         end if
         if rsUpdateEntry4.State=1 then
                    rsUpdateEntry4.close
         end if

         strSQL="EXEC SP_numerar_cargas "& request("ID")              
         rsUpdateEntry3.Open strSQL, adoCon4


         strSQL="select count(*) as calculo from xml where carga= "& request("ID")              
         rsUpdateEntry.Open strSQL, adoCon4

         strSQL="update CargasNimloader  set  NumXml=" & rsUpdateEntry("calculo")   &",helpdesk=1,Estatus='confirmado'  where ID= "& request("ID")              
         rsUpdateEntry2.Open strSQL, adoCon4

         strSQL="select *  from xml where carga= "& request("ID")  &" order by ID"            
         rsUpdateEntry2.Open strSQL, adoCon4

         while not rsUpdateEntry2.EOF
              response.write( "<tr><td class='td-l azul fonttiny'>" & rsUpdateEntry2("ID") &"</td><td class='td-l fonttiny'>" & rsUpdateEntry2("NombreXML") &"</td></tr>")
              rsUpdateEntry2.MoveNext
         wend
         rsUpdateEntry2.close


      response.write("</table><P>&nbsp; </P><P class='subtitulo tc-c fonttiny'>La carga de " & rsUpdateEntry("calculo") & " xmls fue realizada con &eacute;xito, en unos minutos quedar&aacute;n creadas sus ordenes de compra en SAP</P>")
    
 end sub






sub DetalleXML
      response.write("<P>&nbsp;</P><H1>Buscar un XML cargado</H1>")

               response.write("<form id='XmlBrowser' action='ShowContent.asp' method='POST'>") 
               response.write("<input type='hidden' name='contenido' value='30'>")
               response.write("<P class='td-c'>Indique UUID o nombre de archivo: <input type='text' class='anchox2' name='xml' value='"& request("xml") &"'></P>")
               response.write("<P class='td-c'><input type='submit' value='buscar'> </P> </form>") 

   

        if len(trim(request("xml")))>2 then

         strSQL="select count(*) as calculo from xml where UUID like '%"& trim(request("xml")) &"%' or NombreXML like '%"& trim(request("xml")) &"%'"
         rsUpdateEntry.Open strSQL, adoCon4

         strSQL="select * from xml where UUID like '%"& trim(request("xml")) &"%' or NombreXML like '%"& trim(request("xml")) &"%'"
         rsUpdateEntry2.Open strSQL, adoCon4

         tipo=""
         vID=0

         
         if rsUpdateEntry("calculo")=1 or request("UUID")<>"" then

                  if trim(rsUpdateEntry2("viatico"))<>"" then
                     tipo="VIATICO"
                     vID=rsUpdateEntry2("viatico")

                      strSQL="select * from xml where viatico="  & vID &" order by ID"
                      'response.write(strSQL&"<BR>")
                      rsUpdateEntry3.Open strSQL, adoCon4

                  end if 

                  if trim(rsUpdateEntry2("Reembolso"))<>"" then
                     tipo="REEMBOLSO"
                     vID=rsUpdateEntry2("Reembolso")

                     strSQL="select * from xml where Reembolso="  & vID  &" order by ID"
                     'response.write(strSQL&"<BR>")
                     rsUpdateEntry3.Open strSQL, adoCon4

                  end if 

                   if trim(rsUpdateEntry2("carga"))<>"" then
                     tipo="CARGA"
                     vID=rsUpdateEntry2("Carga")

                     strSQL="select * from xml where Carga="  & vID  &" order by ID"
                     'response.write(strSQL&"<BR>")
                     rsUpdateEntry3.Open strSQL, adoCon4

                  end if 
           
                  response.write("<P class='td-c'>El XML fue cargado a trav&eacute;s de <B>"& tipo &"</B> No. <B>"& vID &"</B> en fecha: <B>"&  rsUpdateEntry2("fecha") &"</B> con nombre de archivo: <B>"& rsUpdateEntry2("NombreXML")&"</B></P>")

         end if

           
               if request("carga")<>"" then
                      response.write("<H1>Total de XML subidos a traves de carga no. "& request("carga") &"</H1>")


                     strSQL="select * from xml where Carga="  & request("carga") &" order by ID"
                     'response.write(strSQL&"<BR>")
                     rsUpdateEntry3.Open strSQL, adoCon4

               end if


         
               if rsUpdateEntry3.State=1 then  

                                  response.write("<table width='1360px' border=0 cellpadding=2 cellspacing=2 align=center >" ) %>
                     <tr><td  class="td-c whiteBorders azul">ID</td><td  class="td-c whiteBorders azul" width='9%'>Nombre Archivo</td><td  class="td-c whiteBorders azul">UUID</td><td  class="td-c whiteBorders azul">Fecha</td><td  class="td-c whiteBorders azul">RFC Emisor</td><td  class="td-c whiteBorders azul">Emisor</td><td  class="td-c whiteBorders azul">RFC Receptor</td><td  class="td-c whiteBorders azul">Metodo Pago</td>
                                     <td  class="td-c whiteBorders azul">Forma Pago</td><td  class="td-c whiteBorders azul">Moneda</td><td  class="td-c whiteBorders azul">Folio</td><td  class="td-c whiteBorders azul">Subtotal</td><td  class="td-c whiteBorders azul">Impuestos</td><td  class="td-c whiteBorders azul">Total</td><td  class="td-c whiteBorders azul" width='6%'>Orden Compra</td></tr>
                     <%
                      while not rsUpdateEntry3.EOF
                        response.write("<tr>")
                        response.write("<td class='td-l fonttiny'>"& rsUpdateEntry3("ID") & "</td>")
                        if rsUpdateEntry2("NombreXML")=rsUpdateEntry3("NombreXML") then
                            response.write("<td class='td-l fonttiny alert'>"& rsUpdateEntry3("NombreXML") & "</td>")
                        else
                           response.write("<td class='td-l fonttiny'>"& rsUpdateEntry3("NombreXML") & "</td>")
                        end if
                        if trim(rsUpdateEntry3("UUID"))=trim(request("xml")) then
                           response.write("<td class='td-l fonttiny alert'>"& rsUpdateEntry3("UUID") & "</td>")
                        else
                           response.write("<td class='td-l fonttiny'>"& rsUpdateEntry3("UUID") & "</td>")
                        end if
                        response.write("<td class='td-c fonttiny'>"& FormatDateTime(rsUpdateEntry3("Fecha"),2)  & "</td>")
                        response.write("<td class='td-l fonttiny'>"& rsUpdateEntry3("RFCEmisor") & "</td>")
                        response.write("<td class='td-l fonttiny'>"& mid(rsUpdateEntry3("EMisor"),1,40) & "</td>")
                        response.write("<td class='td-c fonttiny'>"& rsUpdateEntry3("RFCReceptor") & "</td>")
                        response.write("<td class='td-c fonttiny'>"& rsUpdateEntry3("MetodoPago") & "</td>")
                        response.write("<td class='td-c fonttiny'>"& rsUpdateEntry3("FormaPago") & "</td>")
                        response.write("<td class='td-c fonttiny'>"& rsUpdateEntry3("Moneda") & "</td>")
                        response.write("<td class='td-r fonttiny'>"& rsUpdateEntry3("Folio") & "</td>")
                        response.write("<td class='td-r fonttiny'>"& FormatCurrency(rsUpdateEntry3("subtotal")) & "</td>")
                        response.write("<td class='td-r fonttiny'>"& FormatCurrency(rsUpdateEntry3("impuestos")) & "</td>")
                        response.write("<td class='td-r fonttiny'>"& FormatCurrency(rsUpdateEntry3("total")) & "</td>")
                        response.write("<td class='td-r fonttiny'>")

                         strSQL="select * from OPOR where U_CFDi_TimbreUUID='"  & rsUpdateEntry3("UUID") &"' and CANCELED='N' "
                         'response.write strSQL

                        if rsUpdateEntry3("RFCReceptor")="DME140422SS7" then
                           vString="DMX-"
                           rsUpdateEntry4.Open strSQL, adoCon2 'DMX'
                        end if
                    
                         if rsUpdateEntry3("RFCReceptor")="DEL080208G25" then
                           vString="DFW-"
                           rsUpdateEntry4.Open strSQL, adoCon3 'DFW'
                        end if

                        if rsUpdateEntry4.State=1 then        
                            if not rsUpdateEntry4.EOF  then            
                                 response.write (vString & rsUpdateEntry4("Docnum") & " (" & rsUpdateEntry4("DocStatus") &")" )
                            end if
                             rsUpdateEntry4.close
                        end if

                        response.write("</td></tr>")    

                        
                        response.write("</tr>")  
                        rsUpdateEntry3.movenext
                      wend
                      rsUpdateEntry3.close    
                    rsUpdateEntry.close
         rsUpdateEntry2.close



               end if



        else
            if request("xml")<>"" then
                 response.write("<P class='td-c alert'>Par&aacute;metro a buscar debe tener tres caracyeres por lo menos</P>")
            end if
        end if

    
     if rsUpdateEntry.State=1 then        
         rsUpdateEntry.close
     end if


     strSQL="select distinct(carga) from xml where carga is not null and cargado=1 order by carga"
     'response.write strSQL
     rsUpdateEntry.Open strSQL, adoCon4 'temp'


     response.write("</table>")  
     response.write("<P>&nbsp;</P><H1>Buscar por numero de cargas</H1><table width='860px' border=0 cellpadding=2 cellspacing=2 align=center >")
     response.write("<tr>")
    
     for i=1 to 20
          if not rsUpdateEntry.EOF then
              response.write("<td style='td-c'><a href='ShowContent.asp?contenido=30&xml=123&carga="& rsUpdateEntry("carga")  &"' class='fonttiny'>"& rsUpdateEntry("carga") &"</a></td>")
              rsUpdateEntry.MoveNext
          else
              response.write("<td style='td-c'>&nbsp;</td>")
          end if

     next
     response.write("</tr>") 
     rsUpdateEntry.close
     
end sub  




sub H_Incidencias  'contenido 32'


   strSQL="select Nombres+' '+paterno+' '+Materno as nombre,ID,ID_checador from Empleados where FechaEgreso is null order by nombre "
   'response.write strSQL  
   rsUpdateEntry.Open strSQL, adoCon

   
   if request("msg")<>"" then
       response.write("<P class='alert td-c'>"& request("msg") &"</P>")
   end if
      
   modulosHumRes

   response.write("<P style='font-size:10px'>&nbsp;</P>")
   response.write("<table width='1160px' border=1 cellpadding=0 cellspacing=0 align=center><tr><td class='td-c' style='vertical-align: top'>") 


   response.write("<H1>Registro de incidencias de personal</H1>")      
   response.write("<table width='600px' border=0 cellpadding=5 cellspacing=2 align=center>") 
   response.write("<form id='humres' action='ShowContent.asp' method='POST'>") 
   response.write("<input type='hidden' name='contenido' value='33'>")

   response.write("<tr><td class='td-r azul'>Colaborador:</td><td colspan=3 width='70%'><select name=vempleado>")

   response.write("<option value='' SELECTED>sin seleccionar</option>")

   while not rsUpdateEntry.EOF
       response.write("<option value='"& rsUpdateEntry("ID") &"'>"& rsUpdateEntry("nombre") &"</option>")
       rsUpdateEntry.MoveNext
   wend
   rsUpdateEntry.close
   response.write("</select></td></tr>")

   strSQL="select * from cat_incidencias order by ID "
   'response.write strSQL  
   rsUpdateEntry.Open strSQL, adoCon

   response.write("<tr><td class='td-r azul'>Incidencia:</td><td class=tdl><select name=vincidencia>")
   while not rsUpdateEntry.EOF
       response.write("<option value='"& rsUpdateEntry("ID") &"'>"& rsUpdateEntry("incidencia") &"</option>")
       rsUpdateEntry.MoveNext
   wend
   rsUpdateEntry.close
   response.write("</select></td>")
   response.write("<td class='td-r azul'>Fecha:</td><td class=tdl>")    %>
      <input name="vFecha" type="date"/>    </td></tr> <% 

   response.write("<tr><td class='td-r azul'>D&iacute;as consecutivos vacaciones</td><td class='td-l'><select name=vDias>")
   response.write("<option value=0 SELECTED>no aplica</option>")
   for i=1 to 10
       response.write("<option value="& i &">"& i &"</option>")
   next
   response.write("</select></td><td class='td-r azul'>Comentarios:</td><td class=tdl><input type='text' name='vcomentarios' class='anchox1'></td></tr>")
      
   response.write("<tr><td class='td-r azul'>Horas Permiso <BR> (si aplica):</td><td class=tdl><select name=vHoras>")
   response.write("<option value=0 SELECTED>no aplica</option>")
   for i=1 to 8
       response.write("<option value="& i &">"& i &"</option>")
   next
   response.write("</select></td>")
   response.write("<td class='td-r'>&nbsp;</td><td class=tdl>&nbsp;</td></tr>")
   response.write("<tr><td class='td-c' colspan=4><input type=submit value=registrar></td></form></table>") 

   dim vfecha1 : vfecha1=date()
   dim vfecha2 : vfecha2=date()

   if day(date())<16 then
        vfecha1=cdate("01/" & month(date()) & "/"&  year(date()) )
        vfecha2=cdate("16/" & month(date()) & "/"&  year(date()) )
   else
        vfecha1=cdate("16/" & month(date()) & "/"&  year(date()) )
        vfecha2=vfecha1+20
        vfecha2=cdate( "01/" & month(vfecha2) & "/"&   year(vfecha2) )
   end if

   if request("fecha1")<>"" then
        vfecha1=cdate(trim(request("fecha1") ))
        vfecha2=cdate(trim(request("fecha2") ))
   end if

   response.write("<P>&nbsp;</P><table width='660px' border=1 cellpadding=5 cellspacing=2 align=center><tr><td rowspan=2 class='td-r azul'>Quincenas:</td>") 
   response.write("<td class='td-c'><a href='ShowContent.asp?contenido=32&fecha1=01/01/" &year(date) &"&fecha2=16/01/" &year(date) &"'>Ene-1 </a></td>")
   response.write("<td class='td-c'><a href='ShowContent.asp?contenido=32&fecha1=16/01/" &year(date) &"&fecha2=01/02/" &year(date) &"'>Ene-2 </a></td>")
   response.write("<td class='td-c'><a href='ShowContent.asp?contenido=32&fecha1=01/02/" &year(date) &"&fecha2=16/02/" &year(date) &"'>Feb-1 </a></td>")
   response.write("<td class='td-c'><a href='ShowContent.asp?contenido=32&fecha1=16/02/" &year(date) &"&fecha2=01/03/" &year(date) &"'>Feb-2 </a></td>")
   response.write("<td class='td-c'><a href='ShowContent.asp?contenido=32&fecha1=01/03/" &year(date) &"&fecha2=16/03/" &year(date) &"'>Mar-1 </a></td>")
   response.write("<td class='td-c'><a href='ShowContent.asp?contenido=32&fecha1=16/03/" &year(date) &"&fecha2=01/04/" &year(date) &"'>Mar-2 </a></td>")
   response.write("<td class='td-c'><a href='ShowContent.asp?contenido=32&fecha1=01/04/" &year(date) &"&fecha2=16/04/" &year(date) &"'>Abr-1 </a></td>")
   response.write("<td class='td-c'><a href='ShowContent.asp?contenido=32&fecha1=16/04/" &year(date) &"&fecha2=01/05/" &year(date) &"'>Abr-2 </a></td>")
   response.write("<td class='td-c'><a href='ShowContent.asp?contenido=32&fecha1=01/05/" &year(date) &"&fecha2=16/05/" &year(date) &"'>May-1 </a></td>")
   response.write("<td class='td-c'><a href='ShowContent.asp?contenido=32&fecha1=16/05/" &year(date) &"&fecha2=01/06/" &year(date) &"'>May-2 </a></td>")
   response.write("<td class='td-c'><a href='ShowContent.asp?contenido=32&fecha1=01/06/" &year(date) &"&fecha2=16/06/" &year(date) &"'>Jun-1 </a></td>")
   response.write("<td class='td-c'><a href='ShowContent.asp?contenido=32&fecha1=16/06/" &year(date) &"&fecha2=01/07/" &year(date) &"'>Jun-2 </a></td></tr><tr>")

   response.write("<td class='td-c'><a href='ShowContent.asp?contenido=32&fecha1=01/07/" &year(date) &"&fecha2=16/07/" &year(date) &"'>Jul-1 </a></td>")
   response.write("<td class='td-c'><a href='ShowContent.asp?contenido=32&fecha1=16/07/" &year(date) &"&fecha2=01/08/" &year(date) &"'>Jul-2 </a></td>")
   response.write("<td class='td-c'><a href='ShowContent.asp?contenido=32&fecha1=01/08/" &year(date) &"&fecha2=16/08/" &year(date) &"'>Ago-1 </a></td>")
   response.write("<td class='td-c'><a href='ShowContent.asp?contenido=32&fecha1=16/08/" &year(date) &"&fecha2=01/09/" &year(date) &"'>Ago-2 </a></td>")
   response.write("<td class='td-c'><a href='ShowContent.asp?contenido=32&fecha1=01/09/" &year(date) &"&fecha2=16/09/" &year(date) &"'>Sep-1 </a></td>")
   response.write("<td class='td-c'><a href='ShowContent.asp?contenido=32&fecha1=16/09/" &year(date) &"&fecha2=01/10/" &year(date) &"'>Sep-2 </a></td>")
   response.write("<td class='td-c'><a href='ShowContent.asp?contenido=32&fecha1=01/10/" &year(date) &"&fecha2=16/10/" &year(date) &"'>Oct-1 </a></td>")
   response.write("<td class='td-c'><a href='ShowContent.asp?contenido=32&fecha1=16/10/" &year(date) &"&fecha2=01/11/" &year(date) &"'>Oct-2 </a></td>")
   response.write("<td class='td-c'><a href='ShowContent.asp?contenido=32&fecha1=01/11/" &year(date) &"&fecha2=16/11/" &year(date) &"'>Nov-1 </a></td>")
   response.write("<td class='td-c'><a href='ShowContent.asp?contenido=32&fecha1=16/11/" &year(date) &"&fecha2=01/12/" &year(date) &"'>Nov-2 </a></td>")
   response.write("<td class='td-c'><a href='ShowContent.asp?contenido=32&fecha1=01/12/" &year(date) &"&fecha2=16/12/" &year(date) &"'>Dic-1 </a></td>")
   response.write("<td class='td-c'><a href='ShowContent.asp?contenido=32&fecha1=16/12/" &year(date) &"&fecha2=01/01/" &year(date)+1 &"'>Dic-2 </a></td></tr>")

    
   response.write("</table>")

   strSQL="set dateformat dmy;select b.ID,b.nombres+' '+b.paterno+' '+b.materno as colaborador,Fecha as DocDate,dbo.fn_GetMonthName(a.fecha,'spanish') as Fecha,a.INCIDENCIA as ID_Incidencia,c.Incidencia,a.Periodo,a.Horas,a.Notas from Incidencias a inner join Empleados b on a.ID=b.ID inner join cat_incidencias c on a.incidencia=c.ID where Fecha>='" & vfecha1 &"' and Fecha<'" & vfecha2 &"'    order by a.Fecha" 
   'response.write strSQL
   rsUpdateEntry.Open strSQL, adoCon
 
   

   response.write("<P style='font-size: 2px; height: 2px'>&nbsp;</P><table width='820px' border=0 cellpadding=5 cellspacing=2 align=center  class='table2excel table2excel_with_colors' data-tableName='table1'>") 
   response.write("<tr><td colspan=6 class='td-c strong fontbig'>Incidencias Registradas Per&iacute;odo "& vfecha1 &" - " &vfecha2-1 &"</td></tr>")


   response.write("<tr><td class='subtitulo' width='30%'>Colaborador</td>")
   response.write("<td class='subtitulo' width='11%'>Fecha</td>")
   response.write("<td class='subtitulo' width='22%'>Incidencia</td>")
   response.write("<td class='subtitulo' width='5%'>Horas</td>")
   response.write("<td class='subtitulo' width='25%'>Comentarios</td>")
   response.write("<td class='subtitulo'>Controles</td></tr>")
   vfecha=date()
   i=1
  if rsUpdateEntry.EOF then
          response.write("<tr><td colspan=6 class='td-c'>no hay registros en este periodo </td></tr>")

  else  
       while not rsUpdateEntry.EOF
           if i=1 then               
                   response.write("<tr><td colspan=6 class='td-c'><P class='td-c bold'>--------------------<B>"& rsUpdateEntry("Fecha") &"--------------------</B></td></tr>")
           end if
           vfecha=rsUpdateEntry("Fecha") 
           response.write("<tr><td class='td-l fonttiny'>"& rsUpdateEntry("colaborador") &"</td>")
           response.write("<td class='td-c fonttiny' width='80px'>"& rsUpdateEntry("Fecha") &"</td>")

           if int(trim(rsUpdateEntry("id_Incidencia")))=4 then
           	   'RESTANTES'
           	   strSQL="select (select DiasDerecho from PeriodosVacacionales where ID="&rsUpdateEntry("ID") &" and Periodo="&rsUpdateEntry("Periodo") &") - (select count(*) from incidencias where Incidencia=4 and ID="&rsUpdateEntry("ID") &" and Periodo="&rsUpdateEntry("Periodo") &" and Fecha<='"& rsUpdateEntry("DocDate") &"') as calculo"
           	    rsUpdateEntry7.Open strSQL, adoCon

               response.write("<td class='td-l'>"& rsUpdateEntry("Incidencia") &" <br>(<font color=red>dias + en "& rsUpdateEntry("Periodo") &": <b>"& rsUpdateEntry7("calculo") &"</b></font>) </td>")
               rsUpdateEntry7.close
           else
               response.write("<td class='td-l'>"& rsUpdateEntry("Incidencia") &"</td>")
           end if

           if rsUpdateEntry("Horas")=0 then
               response.write("<td class='td-l'>N/A</td>")
           else
               response.write("<td class='td-l'>"& rsUpdateEntry("Horas") &"</td>")
           end if
           response.write("<td class='td-l'>"& rsUpdateEntry("Notas") &"</td>")
           response.write("<td class='td-l'>")
           
        

           response.write("<a href='ShowContent.asp?contenido=34&accion=3&ID="& rsUpdateEntry("ID") &"&ID_incidencia=" & rsUpdateEntry("ID_incidencia") & "&fecha=" & day(rsUpdateEntry("Fecha")) &"-" & month(rsUpdateEntry("Fecha")) &"-" & year(rsUpdateEntry("Fecha"))  &"'> <img src='/img/b_drop.png' with='11px' alt='eliminar' title='eliminar' height='11px' border=0> </a> </td>")

           rsUpdateEntry.MoveNext
           i=i+1

          if not rsUpdateEntry.EOF then
               if vfecha<> rsUpdateEntry("Fecha") then
                   response.write("<tr><td colspan=6 style='font-size:3px'>&nbsp;</td></tr>")
                   response.write("<tr><td colspan=6 class='td-c'><P class='td-c bold'>--------------------<B>"& rsUpdateEntry("Fecha") &"</B>--------------------</td></tr>")
               end if
           end if
          
       wend
  end if
   rsUpdateEntry.close

   %>

   </table><button class="exportToExcel">Exportar a excel</button>  &nbsp;&nbsp;&nbsp;    

         
    <script>
      $(function() {
        $(".exportToExcel").click(function(e){
          var table = $(this).prev('.table2excel');
          if(table && table.length){
            var preserveColors = (table.hasClass('table2excel_with_colors') ? true : false);
            $(table).table2excel({
              exclude: ".noExl",
              name: "Excel Document Name",
              filename: "myFileName" + new Date().toISOString().replace(/[\-\:\.]/g, "") + ".xls",
              fileext: ".xls",
              exclude_img: true,
              exclude_links: true,
              exclude_inputs: true,
              preserveColors: preserveColors
            });
          }
        });
        
      });
    </script>

    <%


   response.write("</td><td width='430px' style='vertical-align: top'> <H1>Solicitud activas</H1> ")
   response.write("<table width='425px' border=1 cellpadding=5 cellspacing=2 align=center>")
   strSQL="select a.*,b.Nombres+' '+b.Paterno+' '+b.Materno as colaborador,dbo.fn_GetMonthName(a.fecha,'spanish') AS fecha2, case when fecha<getdate() then 1 else 0 end as estatus,c.incidencia as descripcion, dbo.fn_GetMonthName(a.Logdate,'spanish') as FechaIntro  from  SolicitudVacaciones a inner join Empleados b on a.ID=b.ID inner join cat_incidencias c on a.incidencia=c.ID where a.Activo=1 order by a.ID,a.FECHA"
   rsUpdateEntry.Open strSQL, adoCon

   i=1
   flag=0

   while not rsUpdateEntry.EOF
      if i>1 then
          if vcolaborador<>rsUpdateEntry("colaborador") then
            response.write("<tr><td colspan=4 class='subtitulo td-l'>"& rsUpdateEntry("colaborador") &"</td>")
            response.write("<td class='subtitulo td-c'>fecha solic.</td></tr>")
          end if
      else
           response.write("<tr><td colspan=4 class='subtitulo td-l'>"& rsUpdateEntry("colaborador") &"</td>")
           response.write("<td class='subtitulo td-c'>fecha solic.</td></tr>")
      end if

      vcolaborador=rsUpdateEntry("colaborador") 

      
      if rsUpdateEntry("estatus")=1 then 
          response.write("<tr><td class='td-c alert Fontsmall' width='23%'>"& rsUpdateEntry("Fecha2") &"</td>")
          flag=1
      else
           response.write("<tr><td class='td-c Fontsmall'  width='23%'>"& rsUpdateEntry("Fecha2") &"</td>")
      end if

      response.write("<td class='td-c Fontsmall' width='23%'>"& rsUpdateEntry("descripcion") &"</td>")

      response.write("<td class='td-c'>"& rsUpdateEntry("periodo") &"</td>")
      response.write("<td class='td-c' width='25%'><a href='ShowContent.asp?contenido=50&ID="& rsUpdateEntry("ID") &"&fecha="& rsUpdateEntry("Fecha") &"&action=1'><img src='/img/ejercer.png' border=0 height='20px' width='20px' alt='ejercer la solicitud' title='ejercer la solicitud'> </a> &nbsp; &nbsp; &nbsp; <a href='ShowContent.asp?contenido=50&ID="& rsUpdateEntry("ID") &"&fecha="& rsUpdateEntry("Fecha") &"&action=2'><img src='/img/b_drop.png' border=0 with=18px height=18px alt='cancelar solicitud' title='cancelar solicitud'> </a> </td>")  

      response.write("<td class='td-c fonttiny' width='21%'>"& rsUpdateEntry("FechaIntro") &"</td></tr>")

      rsUpdateEntry.movenext
      i=i+1
   wend
   rsUpdateEntry.close
   

   if flag=1 then
       response.write("</tr><tr><td colspan=4 class='td-c alert'>la solicitud debe ejercerse o cancelarse </td></tr></table>")
   else
       response.write("</tr><tr><td colspan=4 >&nbsp; </td></tr></table>")
   end if
   response.write("</td></tr></table><P>&nbsp; </P>")
end sub



sub UP_incidencias

    if request("vfecha") <>"" then
        vfecha=date()
        vfecha=cdate(request("vfecha") )
        Dias=int(trim(request("vDias")))    
    end if

    if request("vincidencia")="4"  and  Dias>0  then   'Incidencia 4 = vacaciones'
          
           insert_records=0
           DiasRegistrados=1
           while DiasRegistrados<=Dias
                 flag=1
                 'si ya hay una incidencia registrada'
                 strSQL="set dateformat ymd;select * from Incidencias where ID="& request("vempleado") &" and INCIDENCIA=" & request("vincidencia") &" and FECHA='" & year(vfecha) &"-"& month(vfecha) &"-"& day(vfecha) &"'"
                 'response.write( strSQL  &"<BR>")
                 rsUpdateEntry.Open strSQL, adoCon

                 if not rsUpdateEntry.EOF then
                        flag=0 'si hay una incidencia de este tipo no puede registrarse la incidencia en este dia'
                        'response.write( vfecha &"-1.flag-" & flag &"<BR>")
                 end if
                 rsUpdateEntry.close
              
                 strSQL="set dateformat dmy;select DATEPART(dw,'"& vfecha &"') as numero"
                 'response.write (strSQL &"<BR>")
                 rsUpdateEntry.Open strSQL, adoCon

                 vNum=rsUpdateEntry("numero")
                 rsUpdateEntry.close

                 if Vnum=7 or VNum=1 then  'es sabado o domindo
                     flag=0
                     'response.write( vfecha &"-2.flag-" & flag &"<BR>")
                 end if

                 strSQL="set dateformat ymd;select * from H_DiaInhabil where  FECHA='" & year(vfecha) &"-"& month(vfecha) &"-"& day(vfecha) &"'"
                 'response.write strSQL
                 rsUpdateEntry.Open strSQL, adoCon
                 'response.write( strSQL  &"<BR>")
                 if not rsUpdateEntry.EOF then
                        flag=0 'es dia festivo'
                        'response.write( vfecha &"-3.flag-" & flag &"<BR>")
                 end if
                 rsUpdateEntry.close


                 strSQL="set dateformat ymd;select * from SolicitudVacaciones where ID="& request("vempleado") &" and FECHA='" & year(vfecha) &"-"& month(vfecha) &"-"& day(vfecha) &"' and Activo=1 "
                 'response.write strSQL
                 rsUpdateEntry.Open strSQL, adoCon

                 if not rsUpdateEntry.EOF then
                        flag=0  'tiene registrado una solicitud de vacaciones que debe ejercerse'
                        'response.write( vfecha &"-4.flag-" & flag &"<BR>")
                 end if
                 rsUpdateEntry.close


                 strSQL="set dateformat ymd;select * from Incidencias where ID="& request("vempleado") &" and INCIDENCIA=3  and FECHA='" & year(vfecha) &"-"& month(vfecha) &"-"& day(vfecha) &"'"
                 rsUpdateEntry.Open strSQL, adoCon
                                
                 if flag=1 then
                       'response.write (vfecha &"###<BR>")
                          strSQL="select top 1 Periodo from PeriodosVacacionales where (DiasDerecho-DiasTomados)>0 and ID="& request("vempleado") &" order by Periodo" 
                          rsUpdateEntry2.Open strSQL, adoCon
                          vPeriodo=rsUpdateEntry2("Periodo")
                          rsUpdateEntry2.close

                          if request("vincidencia")="4"  then  'si es vacaciones insertar periodo'

                        strSQL="set dateformat ymd;insert into Incidencias (ID,FECHA,INCIDENCIA,NOTAS,ID_USUARIO,LogDate,Periodo) values ("& request("vempleado") &",'"& year(vfecha) &"-"& month(vfecha) &"-"& day(vfecha) &"',"& request("vincidencia") &",'"& request("vcomentarios") &"','" & Session("session_id")  & "',getdate(),"& vPeriodo &"  )"
                          rsUpdateEntry2.Open strSQL, adoCon
                          insert_records=1

                        else
                            strSQL="set dateformat ymd;insert into Incidencias (ID,FECHA,INCIDENCIA,NOTAS,ID_USUARIO,LogDate) values ("& request("vempleado") &",'"& year(vfecha) &"-"& month(vfecha) &"-"& day(vfecha) &"',"& request("vincidencia") &",'"& request("vcomentarios") &"','" & Session("session_id")  & "',getdate()  )"
                          rsUpdateEntry2.Open strSQL, adoCon
                            insert_records=1

                         end if

                          

                           strSQL="update PeriodosVacacionales set DiasTomados=(select DiasTomados from PeriodosVacacionales where ID="& request("vempleado") &" and Periodo="& vPeriodo &")+1 where ID="&request("vempleado") &" and Periodo="& vPeriodo
                           rsUpdateEntry2.Open strSQL, adoCon

                                           
                 end if
                 vfecha=vfecha+1
                 DiasRegistrados=DiasRegistrados+1  
                 rsUpdateEntry.close

           wend
         if insert_records=1 then
              response.redirect("ShowContent.asp?contenido=32&msg=las vacaciones quedaron registradas")
         else
              response.redirect("ShowContent.asp?contenido=32&msg=no se registraron vacaciones, es probable que haya solicitudes que deban ejercerse")
         end if

    else
    
        if request("vempleado")<>"" then

          strSQL="set dateformat ymd;select * from Incidencias where ID="& request("vempleado") &" and INCIDENCIA=" & request("vincidencia") &" and FECHA='" & year(vfecha) &"-"& month(vfecha) &"-"& day(vfecha) &"'"
          rsUpdateEntry.Open strSQL, adoCon

          if not rsUpdateEntry.EOF then
              response.redirect("ShowContent.asp?contenido=32&msg=esta intentando duplicar una incidencia")
          end if
          rsUpdateEntry.close


          strSQL="set dateformat ymd;select * from SolicitudVacaciones where ID="& request("vempleado") &" and FECHA='" & year(vfecha) &"-"& month(vfecha) &"-"& day(vfecha) &"' and Activo=1 "
          'response.write strSQL
          rsUpdateEntry.Open strSQL, adoCon

          if not rsUpdateEntry.EOF then
              response.redirect("ShowContent.asp?contenido=32&msg=no puede registrar vacaciones en esta fecha, revise si hay una solicitud de vacaciones que puede ejercer (lado derecho)")
          end if
          rsUpdateEntry.close


          if request("vincidencia")="4" then
                          strSQL="select top 1 Periodo from PeriodosVacacionales where (DiasDerecho-DiasTomados)>0 and ID="& request("vempleado") &" order by Periodo" 
                          rsUpdateEntry2.Open strSQL, adoCon
                          vPeriodo=rsUpdateEntry2("Periodo")
                          rsUpdateEntry2.close

                          strSQL="set dateformat ymd;insert into Incidencias (ID,FECHA,INCIDENCIA,HORAS,NOTAS,ID_USUARIO,LogDate,Periodo) values ("& request("vempleado") &",'"& year(vfecha) &"-"& month(vfecha) &"-"& day(vfecha) &"',"& request("vincidencia") &","& request("vHoras") &",'"& request("vcomentarios") &"','" & Session("session_id")  & "',getdate(), "& vPeriodo &")"
                          rsUpdateEntry.Open strSQL, adoCon


                           strSQL="update PeriodosVacacionales set DiasTomados=(select DiasTomados from PeriodosVacacionales where ID="& request("vempleado") &" and Periodo="& vPeriodo &")+1 where ID="&request("vempleado") &" and Periodo="& vPeriodo
                           rsUpdateEntry2.Open strSQL, adoCon
     
          else


                  strSQL="set dateformat ymd;insert into Incidencias (ID,FECHA,INCIDENCIA,HORAS,NOTAS,ID_USUARIO,LogDate) values ("& request("vempleado") &",'"& year(vfecha) &"-"& month(vfecha) &"-"& day(vfecha) &"',"& request("vincidencia") &","& request("vHoras") &",'"& request("vcomentarios") &"','" & Session("session_id")  & "',getdate()  )"
                  rsUpdateEntry.Open strSQL, adoCon


          end if          
          response.redirect("ShowContent.asp?contenido=32&msg=la incidencia ha quedado registrada")

        else

            response.redirect("ShowContent.asp?contenido=32&msg=solicitud incompleta")
        end if

    end if


end sub




sub edit_incidencias  'contenido 34'
     
     
     if trim(request("accion"))=1 then 'editar incidencia'
        vfecha1=cdate(request("Fecha") )

        strSQL="set dateformat dmy;select b.ID,b.nombres+' '+b.paterno+' '+b.materno as colaborador,dbo.fn_GetMonthName(a.fecha,'spanish') as Fecha,a.INCIDENCIA as ID_Incidencia,c.Incidencia,a.Horas,a.Notas from Incidencias a inner join Empleados b on a.ID=b.ID inner join cat_incidencias c on a.incidencia=c.ID where Fecha='" & vfecha1 &"' and b.id=" & request("ID") &" and a.INCIDENCIA=" & request("ID_incidencia")
          'response.write strSQL
          rsUpdateEntry.Open strSQL, adoCon

          if rsUpdateEntry.EOF then
               response.write("<P class='td-c alert'>Ocurrio un error <BR> "& strSQL &"</P>")
          else

              'response.write("<P class='td-c alert'>OK <BR> "& strSQL &"</P>")

               response.write("<P>&nbsp</P><H1>Edici&oacute;n de incidencia de personal</H1><table width='800px' border=0 cellpadding=5 cellspacing=2 align=center>")

               response.write("<form id='humres' action='ShowContent.asp' method='POST'>") 

               response.write("<input type='hidden' name='contenido' value='34'>")
               response.write("<input type='hidden' name='accion' value='2'>")
               response.write("<input type='hidden' name='colaborador' value='"& rsUpdateEntry("colaborador") &"'>")
               response.write("<input type='hidden' name='ID' value='"& request("ID") &"'>")
               response.write("<input type='hidden' name='ID_incidencia' value='"& request("ID_incidencia") &"'>")
               response.write("<input type='hidden' name='fecha' value='"& trim(request("fecha")) &"'>")

               response.write("<tr><td class='td-r azul' width='30%'>Colaborador:</td><td colspan=3>"& rsUpdateEntry("colaborador") &"</td></tr>")

               strSQL="select * from cat_incidencias order by ID "
               'response.write strSQL  
               rsUpdateEntry2.Open strSQL, adoCon

               response.write("<tr><td class='td-r azul'>Incidencia:</td><td class=tdl><select name=vincidencia>")
                while not rsUpdateEntry2.EOF
                   if int(trim(rsUpdateEntry("ID_incidencia")))=int(trim(rsUpdateEntry2("ID"))) then
                       response.write("<option value='"& rsUpdateEntry2("ID") &"' SELECTED>"& rsUpdateEntry2("incidencia") &"</option>")
                   else
                       response.write("<option value='"& rsUpdateEntry2("ID") &"'>"& rsUpdateEntry2("incidencia") &"</option>")
                   end if
                   rsUpdateEntry2.MoveNext
                wend
                rsUpdateEntry2.close

               response.write("</select></td>")
               response.write("<td class='td-r azul'>Fecha:</td><td class=tdl>"& rsUpdateEntry("Fecha") &"</td></tr>")  

               response.write("<tr><td class='td-r azul'>Horas Permiso <BR> (si aplica):</td><td class=tdl><select name=vHoras>")
               if rsUpdateEntry("Horas")=0 then
                   response.write("<option value=0 SELECTED>no aplica</option>")
               else
                   response.write("<option value=0>no aplica</option>")
               end if

               for i=1 to 8
                   if rsUpdateEntry("Horas")=i then
                      response.write("<option value="& i &" SELECTED>"& i &"</option>")
                   else
                      response.write("<option value="& i &">"& i &"</option>")
                   end if
               next

               response.write("</select></td>")
               response.write("<td class='td-r azul'>Comentarios:</td><td class=tdl><input type='text' name='vcomentarios' class='anchox3' value='"& rsUpdateEntry("notas") &"'></td></tr><tr><td class='td-c' colspan=4>&nbsp;</td></tr>")
               response.write("<tr><td class='td-c' colspan=4><input type=submit value=actualizar></td></form></table><P>&nbsp</P>") 

     
          end if
          rsUpdateEntry.close

     end if



     if trim(request("accion"))=2 then 'actualizar incidencia'

          vfecha=cdate( request("fecha"))

          strSQL="set dateformat ymd;UPDATE Incidencias set INCIDENCIA=" & request("vincidencia") &",HORAS=" & request("vHoras") &", Notas='" & request("vcomentarios") &"' where ID="& request("ID") &" and INCIDENCIA=" & request("ID_incidencia") &" and FECHA='" & year(vfecha) &"-"& month(vfecha) &"-"& day(vfecha) &"'"
          'response.write strSQL
          rsUpdateEntry.Open strSQL, adoCon

          response.redirect("ShowContent.asp?contenido=32&msg=Se actualizo incidencia de "& request("colaborador") & " en " & vfecha ) 

     end if





     if trim(request("accion"))=3 then 'form borar incidencia'


     vfecha1=cdate(request("Fecha") )

        strSQL="set dateformat dmy;select b.ID,b.nombres+' '+b.paterno+' '+b.materno as colaborador,dbo.fn_GetMonthName(a.fecha,'spanish') as Fecha,a.INCIDENCIA as ID_Incidencia,c.Incidencia,a.Horas,a.Notas from Incidencias a inner join Empleados b on a.ID=b.ID inner join cat_incidencias c on a.incidencia=c.ID where Fecha='" & vfecha1 &"' and b.id=" & request("ID") &" and a.INCIDENCIA=" & request("ID_incidencia")
          'response.write strSQL
          rsUpdateEntry.Open strSQL, adoCon

          if rsUpdateEntry.EOF then
               response.write("<P class='td-c alert'>Ocurrio un error <BR> "& strSQL &"</P>")
          else

              'response.write("<P class='td-c alert'>OK <BR> "& strSQL &"</P>")

               response.write("<P>&nbsp</P><H1>Eliminar una incidencia de personal</H1><table width='800px' border=0 cellpadding=5 cellspacing=2 align=center>")

               response.write("<form id='humres' action='ShowContent.asp' method='POST'>") 

               response.write("<input type='hidden' name='contenido' value='34'>")
               response.write("<input type='hidden' name='accion' value='5'>")  'no usar accion 4'
               response.write("<input type='hidden' name='colaborador' value='"& rsUpdateEntry("colaborador") &"'>")
               response.write("<input type='hidden' name='ID' value='"& request("ID") &"'>")
               response.write("<input type='hidden' name='ID_incidencia' value='"& request("ID_incidencia") &"'>")
               response.write("<input type='hidden' name='fecha' value='"& trim(request("fecha")) &"'>")

               response.write("<tr><td class='td-r azul'>Colaborador:</td><td colspan=3>"& rsUpdateEntry("colaborador") &"</td></tr>")

               strSQL="select * from cat_incidencias order by ID "
               'response.write strSQL  
               rsUpdateEntry2.Open strSQL, adoCon

               response.write("<tr><td class='td-r azul'>Incidencia:</td><td class=tdl>"& rsUpdateEntry("incidencia") &"</td>")
               response.write("<td class='td-r azul'>Fecha:</td><td class=tdl>"& rsUpdateEntry("Fecha") &"</td></tr>")  

               response.write("<tr><td class='td-r azul'>Horas Permiso <BR> (si aplica):</td><td class=tdl>")
               if rsUpdateEntry("Horas")=0 then
                   response.write("no aplica</td>")
               else
                   response.write( rsUpdateEntry("Horas") &"</td>")
               end if
               
               response.write("<td class='td-r azul'>Comentarios:</td><td class=tdl>"& rsUpdateEntry("notas") &"</td></tr><tr><td class='td-c' colspan=4>&nbsp;</td></tr>")
               response.write("<tr><td class='td-c' colspan=4><input type=submit value=eliminar></td></form></table><P>&nbsp</P>") 

     
          end if
          rsUpdateEntry.close

     

     end if



     if trim(request("accion"))=5 then 'borar incidencia'

      vfecha1=cdate(request("Fecha") )
      response.write (vfecha1 & "<BR>")

      if request("ID_incidencia")=4 then 'borro una vacacion'
           strSQL="set dateformat ymd;select * from Incidencias where incidencia=4 and fecha='"& year(vfecha1) &"-" & month(vfecha1) &"-" & day(vfecha1)  &"' and id=" & request("ID") 

           rsUpdateEntry.Open strSQL, adoCon 

           strSQL="UPDATE PeriodosVacacionales set DiasTomados=DiasTomados-1 where id=" & request("ID") &" and Periodo=" & rsUpdateEntry("Periodo")
           rsUpdateEntry.close

           rsUpdateEntry.Open strSQL, adoCon 

      end if



      strSQL="set dateformat ymd;DELETE incidencias  where Fecha='" & year(vfecha1) &"-" & month(vfecha1) &"-" & day(vfecha1)  &"' and id=" & request("ID") &" and INCIDENCIA=" & request("ID_incidencia")
      response.write (strSQL & "<BR>")

     

      rsUpdateEntry.Open strSQL, adoCon      
      response.redirect("ShowContent.asp?contenido=32&msg=Se elimino una incidencia seleccionada") 

     end if


end sub  



sub FormOmision    'contenido 35'
  
  strSQL="select Nombres+' '+paterno+' '+Materno as nombre,ID,ID_checador from Empleados where FechaEgreso is null order by nombre "
   'response.write strSQL  
   rsUpdateEntry.Open strSQL, adoCon

   response.write("<P>&nbsp;</P>")
   if request("msg")<>"" then
       response.write("<P class='alert td-c'>"& request("msg") &"</P>")
   end if
   response.write("<H1>Registro de omisi&oacute;n en checador (salida)</H1><table width='470px' border=0 cellpadding=5 cellspacing=2 align=center>")

   response.write("<form id='humres' action='ShowContent.asp' method='POST'>") 
   response.write("<input type='hidden' name='contenido' value='35'>")
   response.write("<input type='hidden' name='opt' value='2'>")


   response.write("<tr><td class='td-r azul'>Colaborador:</td><td><select name=vempleado>")
   while not rsUpdateEntry.EOF
       response.write("<option value='"& rsUpdateEntry("ID") &"'>"& rsUpdateEntry("nombre") &"</option>")
       rsUpdateEntry.MoveNext
   wend
   rsUpdateEntry.close
   response.write("</select></td></tr>")
   response.write("<tr><td class='td-c' colspan=2>&nbsp;</td></tr>")
   response.write("<tr><td class='td-c' colspan=2><input type='submit' value='Mostras checadas omitidas'></form></td></tr>")



end sub


Sub OmisionRequest  'contenido 35'
   strSQL="select Nombres+' '+paterno+' '+Materno as nombre,ID,ID_checador from Empleados where id=" & request("vempleado")
   'response.write strSQL  
   rsUpdateEntry.Open strSQL, adoCon

   response.write("<P>&nbsp;</P>")
   if request("msg")<>"" then
       response.write("<P class='alert td-c'>"& request("msg") &"</P>")
   end if
   response.write("<H1>Registro de omisi&oacute;n en checador (salida)</H1><table width='470px' border=0 cellpadding=5 cellspacing=2 align=center>")

   response.write("<tr><td class='td-r azul' >Colaborador:</td><td>"& rsUpdateEntry("nombre") &"</td></tr>")
   response.write("<tr><td class='td-c azul' colspan=2>Checadas omitidas en el per&iacute;odo disponible a justificar</td></tr>")

   if day(date())<16 then
        vfecha1=cdate("01/" & month(date()) & "/"&  year(date()) )
        vfecha1=cdate("16/" & month(date()) & "/"&  year(date()) )
   else
        vfecha1=cdate("16/" & month(date()) & "/"&  year(date()) )
        vfecha2=vfecha1+20
        vfecha2=cdate( "01/" & month(vfecha2) & "/"&   year(vfecha2) )
   end if

    strSQL="set dateformat ymd;select a.ID,a.nombres+' '+a.paterno+' '+a.materno as colaborador,dbo.fn_GetMonthName(b.Entrada,'spanish') as Entrada,dbo.fn_GetMonthName(b.salida,'spanish') as Salida from Empleados a inner join H_Evento b on a.ID_checador=b.IDEmpleado where b.Entrada is not null and b.Salida is null and  b.Entrada>='" & year(vfecha1) &"-" & month(vfecha1) &"-" &day(vfecha1)  &"' and b.Entrada<'" & year(vfecha2) &"-" &month(vfecha2) &"-" &day(vfecha2)  &"' and a.ID=" & request("vempleado") &" and cast(b.Entrada as date)<cast(getdate() as date)   and cast(b.Entrada as date) not in ( select fecha from incidencias where incidencia=11 and ID="& request("vempleado") &" )   order by b.Entrada"
    'response.write strSQL  
    rsUpdateEntry2.Open strSQL, adoCon    

    response.write("<tr><td class='td-c azul'>Entrada</td><td class='td-c azul'>Salida</td><tr>")
    if rsUpdateEntry2.EOF then         
         response.write("<tr><td class='td-c' colspan=2>no hay registros con estos criterios</td><tr>")
    else

        while not rsUpdateEntry2.EOF           
           response.write("<tr><td class='td-c'><a href='ShowContent.asp?contenido=35&opt=3&id=" & rsUpdateEntry2("ID")  &"&fecha=" & day(rsUpdateEntry2("entrada"))&"-"& month(rsUpdateEntry2("entrada"))&"-"& year(rsUpdateEntry2("entrada")) &"'>" & rsUpdateEntry2("Entrada") &"</a></td><td class='td-c'><a href='ShowContent.asp?contenido=35&opt=3&id=" & rsUpdateEntry2("ID")  &"&fecha=" & day(rsUpdateEntry2("entrada"))&"-"& month(rsUpdateEntry2("entrada"))&"-"& year(rsUpdateEntry2("entrada")) &"'>no hay registro de salida</a> </tr>")
      
           rsUpdateEntry2.MoveNext
        wend
    end if
   
   rsUpdateEntry2.close
   rsUpdateEntry.close

end sub




sub OmisionJustifique    'contenido 35'
   'response.write( request("ID") & "<BR>" )
   'response.write( request("fecha") & "<BR>" )


   strSQL="select Nombres+' '+paterno+' '+Materno as nombre,ID,ID_checador from Empleados where id=" & request("ID")
   'response.write strSQL  
   rsUpdateEntry.Open strSQL, adoCon

   response.write("<P>&nbsp;</P>")
   if request("msg")<>"" then
       response.write("<P class='alert td-c'>"& request("msg") &"</P>")
   end if
   response.write("<H1>Registro de omisi&oacute;n en checador (salida)</H1><table width='470px' border=0 cellpadding=5 cellspacing=2 align=center>")

   response.write("<form id='humres' action='ShowContent.asp' method='POST'>") 
   response.write("<input type='hidden' name='contenido' value='35'>")
   response.write("<input type='hidden' name='id' value='"&  request("ID")  &"'>")
   response.write("<input type='hidden' name='fecha' value='"&  request("fecha")  &"'>")
   response.write("<input type='hidden' name='opt' value='4'>")

   response.write("<tr><td class='td-r azul' >Colaborador:</td><td>"& rsUpdateEntry("nombre") &"</td></tr>")
   response.write("<tr><td class='td-r azul' >Omision de checada al salir:</td><td>"& request("fecha") &"</td></tr>")
   response.write("<tr><td class='td-r azul' >Justificaci&oacute;n:</td><td><input type='text' name='Vcausa' class=anchox2></td></tr>")

   response.write("<tr><td class='td-c' colspan=2>&nbsp;</td></tr>")
   response.write("<tr><td class='td-c' colspan=2><input type='submit' value='Registrar'></form></td></tr>")

end sub  




Sub  UPomision   'contenido 35'
   response.write( request("ID") & "<BR>" )
   response.write( request("fecha")  & "<BR>" )
   response.write( request("Vcausa")  & "<BR>" )

   strSQL="set dateformat dmy;INSERT into  incidencias (ID,FECHA,INCIDENCIA,NOTAS)   values ("& request("ID") &",'"& request("fecha") &"',11,'"& request("Vcausa") &"')"         
   response.write strSQL  
   rsUpdateEntry.Open strSQL, adoCon

   response.redirect("ShowContent.asp?contenido=35&opt=1&msg=se registro la omision")

end sub  




       

sub ControlEnvios   'contenido 36'     
 flag=0  
    strSQL="select WhsCode,count(whsCode) as calculo from envios where DocDate>='2021-04-01' and tipo='remision' group by WhsCode"
    rsUpdateEntry.Open strSQL, adoCon4
    titulo="Control de env&iacute;os y mensajer&iacute;as <br>(desde 1-abr-21 | "& rsUpdateEntry("WhsCode") &": " & rsUpdateEntry("calculo") &" | "
    rsUpdateEntry.MoveNext
    titulo= titulo & rsUpdateEntry("WhsCode") &": " & rsUpdateEntry("calculo") &")"
    rsUpdateEntry.close
    DoTitulo(titulo)

    response.write("<table border=0 align=center cellspacing=2 cellpadding=5>")
    'CreateTable(1420)

    response.write("<form id='buscar' method='POST' action='ShowContent.asp'>")
    response.write("<input type='hidden' name='contenido' value=36>")
 
 
    if flag=0 then
       titulos
    end if

    response.write("<tr>")
    response.write("<td><input type='number' style='width:60px' name='fID' value='"& request("fID") &"'></td>")
    response.write("<td><select name='fcolaborador' style='width: 100px'><option value=0 selected>seleccione</option>")
    strSQL="select ID,Nombres+' '+paterno+' '+Materno as nombre from Empleados where FechaEgreso is null order by nombre" 
    rsUpdateEntry2.Open strSQL, adoCon
    while not rsUpdateEntry2.EOF
      response.write("<option value="& rsUpdateEntry2("ID") &">"& mid(rsUpdateEntry2("nombre"),1,80) &"</option>")
      rsUpdateEntry2.MoveNext
    wend
    rsUpdateEntry2.close
    response.write("</select></td>")
    response.write("<td><input type='date' style='width:90px' name='ffecha'></td>")
    response.write("<td colspan=3>&nbsp;</td>")
    response.write("<td><select name='fportador' style='width: 100px'><option value=0 selected>seleccione</option>")
    strSQL="select * from cat_portadores" 
    rsUpdateEntry2.Open strSQL, adoCon4
    while not rsUpdateEntry2.EOF
      response.write("<option value="& rsUpdateEntry2("id_Portador") &">"& mid(rsUpdateEntry2("Portador"),1,80) &"</option>")
      rsUpdateEntry2.MoveNext
    wend
    rsUpdateEntry2.close
    response.write("</select></td>")
    response.write("<td colspan=2 class='td-c'><input type='text' style='width:180px' name='vstring'></td>")
    response.write("<td >&nbsp;</td>")
    response.write("<td class='td-c'><input type='text' style='width:60px' name='fsubtotal'></td>")
    'response.write("<td colspan=2 class='td-c'>&nbsp;</td>")
    response.write("<td><input type='number' style='width:60px' name='fPedido' value='"& request("fPedido") &"'></td>")
    response.write("<td class='td-c'>&nbsp;</td>")
    response.write("<td><input type='number' style='width:60px' name='fsubtotalc' value='"& request("fsubtotalc") &"'></td>")
      response.write("<td><input type='number' style='width:60px' name='ffactura' value='"& request("ffactura") &"'></td>")
    response.write("</tr>")
    response.write("<tr><td colspan=15 class='td-c'> <input type='submit' value='buscar'></form> </td></tr>")



  Const adCmdText = &H0001
  Const adOpenStatic = 3
  nPage=0
  registros=1
  flag=0
  Flag_Pedido=0

  strSQL="select  a.*,b.Portador,dbo.fn_GetMonthName(a.DocDate,'spanish') AS [DocDate2],c.DocNum as Factura,d.Ruta  from envios a inner join cat_portadores b on a.id_Portador=b.id_Portador "
  strSQL= strSQL &"left join Notificaciones c on c.tipo='factura' and c.CANCELED='N' and a.RazonSocial=c.RazonSocial and a.DocNum=c.remision left join Repositorio d on a.RazonSocial=d.RazonSocial collate SQL_Latin1_General_CP850_CI_AS and a.CardCode=d.CardCode collate SQL_Latin1_General_CP850_CI_AS "
  strSQL= strSQL &" where 1=1 "

  if request("fID")<>"" then 
      strSQL=strSQL & "and a.ID="& request("fID") &" "
      flag=1
  end if
  if request("fcolaborador")>0 then 
      strSQL=strSQL & "and a.ID_empleado="& request("fcolaborador") &" "
      flag=1
  end if
  if request("ffecha")<>"" then 
      vfecha=date()
      vfecha=request("ffecha")
      'response.write( vFecha )
      strSQL=strSQL & "and year(a.DocDate)=" & year(vfecha) &" and month(a.DocDate)=" & month(vfecha) &" and day(a.DocDate)=" & day(vfecha) &" "
      flag=1
  end if
  if request("fportador")>0 then 
      strSQL=strSQL & "and a.id_Portador=" & request("fportador")  &" "
      flag=1
  end if
  if request("vstring")<>"" and len(request("vstring"))>2 then
      strSQL=strSQL & "and ( guia like '%"&request("vstring")&"%' OR rastreo like '%"&request("vstring")&"%' )  " 
      flag=1
  end if
   if request("fsubtotal")<>"" then 
      vString=trim(request("fsubtotal"))
      vPos=inStr(vString,".")
      if vPos>0 then
          vString=mid(vString,1,vpos-1)
      end if
      strSQL=strSQL & "and a.subtotal>=" & int(vString)-3  &" and a.subtotal<=" & int(vString)+3 &" "
      flag=1
  end if
  if request("fPedido")<>"" then 
      strSQL=strSQL & "and a.Pedido=" & request("fPedido")  &" "
      flag=1
  end if
   if request("fsubtotalc")<>"" then 
      vString=trim(request("fsubtotalc"))
      vPos=inStr(vString,".")
      if vPos>0 then
          vString=mid(vString,1,vpos-1)
      end if
      strSQL=strSQL & "and a.subtotalconsolidado>=" & int(vString)-3  &" and a.subtotalconsolidado<=" & int(vString)+3 &" "
      flag=1
  end if

  if request("ffactura")<>"" then 
     strSQL=strSQL & "and c.DocNum=" &request("ffactura")
     flag=1
  end if

  if flag=0 then
           strSQL=strSQL & "and datediff(MONTH,a.DocDate,getdate())<16 "
  end if
  strSQL=strSQL & "order by a.ID desc"
  'response.write strSQL

  rsUpdateEntry.Open strSQL, adoCon4, adOpenStatic, adCmdText
  nPageCount=0

  if not rsUpdateEntry.EOF then
      rsUpdateEntry.PageSize = 50
      nPageCount = rsUpdateEntry.PageCount

      if request("vPage")<>"" then
           nPage = int(trim(request("vPage")))
      else
           nPage=1
      end if      
      rsUpdateEntry.AbsolutePage=npage
   end if

   if flag=0 then
      separador 
      response.write("<tr><td colspan=2 class='td-r'><B>PAGINAS:</B></td><td colspan=13>")
            for i=1 to nPageCount step 1
                if i<>npage then  
                    response.write("<a href='/modules/ShowContent.asp?contenido=36&vPage="&i &"'>") 
                    response.write( i &"</A>&nbsp;&nbsp;&nbsp;&nbsp;") 
                end if
            next
            response.write("</td></tr>")   
          separador
    end if


if  rsUpdateEntry.EOF then
    NoRegistros

else


   
     if request("fPedido")<>"" then
        strSQL="select top 1 * from Notificaciones where Docnum=" &  request("fPedido")  &" and tipo='pedido' and modulo='ventas' order by DocDate desc" 
        rsUpdateEntry3.Open strSQL, adoCon4

        RazonSocial=rsUpdateEntry3("RazonSocial")
        rsUpdateEntry3.close

        strSQL="select *,DocTotalFc-VatSumFc as 'subtotal' from ORDR where DocNum=" & request("fPedido")        
        'response.write strSQL

        if RazonSocial="DMX" then
             rsUpdateEntry3.Open strSQL, adoCon2
             'response.write("DMX=" & strSQL )
        else
             rsUpdateEntry3.Open strSQL, adoCon3
             'response.write("DFW=" & strSQL )
        end if
 
        response.write("<tr><td colspan=15 >&nbsp;</td></tr>")
        response.write("<tr><td colspan=15 >&nbsp;</td></tr>")

        if rsUpdateEntry3.EOF then
              response.write("error: es fin de archivo , "& RazonSocial &" en query: " & strSQL )
        end if
       
        
        response.write("<tr><td colspan=3 class='td-r subtitulo fonttiny'>C&oacute;digo Cliente: </td><td colspan=5 class='td-l'>"& rsUpdateEntry3("CardCode") &"</td><td colspan=9>&nbsp;</td></tr>")

        response.write("<tr><td colspan=3 class='td-r subtitulo fonttiny'>Cliente: </td><td colspan=5 class='td-l'>"& rsUpdateEntry3("cardname") &"</td><td colspan=9>&nbsp;</td></tr>")
        response.write("<tr><td colspan=3 class='td-r subtitulo fonttiny'>Proyecto: </td><td colspan=5 class='td-l'>"& rsUpdateEntry3("U_PROYECTO") &"</td><td colspan=9>&nbsp;</td></tr>")

        response.write("<tr><td colspan=3 class='td-r subtitulo fonttiny'>Valor total del pedido sin IVA: </td><td colspan=5 class='td-l'>$ "& rsUpdateEntry3("subtotal") &" USD </td><td colspan=9>&nbsp;</td></tr>")
        response.write("<tr><td colspan=3 class='td-r subtitulo fonttiny'>Flete total Estimado: </td><td colspan=5 class='td-l'>$ <font color=green>"& rsUpdateEntry3("U_FleteTotal") &" MXN </font></td><td colspan=9>&nbsp;</td></tr>")
        vFleteTotal=rsUpdateEntry3("U_FleteTotal")
        response.write("<tr><td colspan=3 class='td-r subtitulo fonttiny'>Estatus del pedido: </td><td colspan=5 class='td-l'>")

        if rsUpdateEntry3("DocStatus")="C" then
           response.write("<font color=green>CERRADO </font></td><td colspan=9>&nbsp;</td></tr>")
        else
           response.write("<font color=red>ABIERTO &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font><U><a href='ShowContent.asp?contenido=68&action=1&fRS="&RazonSocial &"&fPedido="& request("fPedido")  &"' target='_pedido'>mostrar partidas abiertas</a></U></td><td colspan=9>&nbsp;</td></tr>")
        end if

       
        rsUpdateEntry3.close
        response.write("<tr><td colspan=15 >&nbsp;</td></tr>")  
        response.write("<tr><td colspan=15 >&nbsp;</td></tr>")
        response.write("<tr><td class='td-c' colspan=15><H1>GASTOS DE FLETES/PAQUETERIA EJERCIDOS</H1></td></tr>")

      end if






  while not rsUpdateEntry.EOF AND registros<=50

     if request("fPedido")<>"" then       
        RazonSocial=rsUpdateEntry("RazonSocial")
     end if

     response.write("<tr>")

     response.write("<td class='td-r fonttiny'>"& rsUpdateEntry("ID") &"</td>") 

     strSQL="select  Nombres+' '+paterno as Colaborador from Empleados where id=" &   rsUpdateEntry("ID_empleado")
     rsUpdateEntry4.Open strSQL, adoCon

     response.write("<td class='td-l fonttiny'>"& left(rsUpdateEntry4("Colaborador"),14) &"</td>")
     rsUpdateEntry4.close
     response.write("<td class='td-r fonttiny'>"& rsUpdateEntry("DocDate2") &"</td>")
     response.write("<td class='td-c fonttiny'>"& rsUpdateEntry("RazonSocial") &"</td>")
     response.write("<td class='td-r fonttiny'><a href='/modules/ShowContent.asp?contenido=124&action=1&fRS="&rsUpdateEntry("RazonSocial")&"&fremision="&rsUpdateEntry("DocNum")&"' target='remision'>"& rsUpdateEntry("DocNum") &"</a></td>")
     'response.write("<td class='td-c fonttiny'>"& rsUpdateEntry("Tipo") &"</td>")
     response.write("<td class='td-c fonttiny'>"& rsUpdateEntry("WhsCode") &"</td>")
     response.write("<td class='td-l fonttiny'>"& rsUpdateEntry("Portador") &"</td>")
     response.write("<td class='td-l fonttiny'>"& rsUpdateEntry("Guia") &"</td>")

     select case rsUpdateEntry("id_Portador")
          case 1 response.write("<td class='td-l fonttiny'><a href='https://www.paquetexpress.com.mx/rastreo/"& rsUpdateEntry("Rastreo") &"' target='rastreo'>"& rsUpdateEntry("Rastreo") &"&nbsp;<img src='/images/tracking.png' width='30px' border=0 title='rastree'></a></td>")  
           case 2 response.write("<td class='td-l fonttiny'><a href='https://www.castores.com.mx/rastreo/"& rsUpdateEntry("Rastreo") &"' target='rastreo'>"& rsUpdateEntry("Rastreo") &"&nbsp;<img src='/images/tracking.png' width='30px' border=0 title='rastree'></a></td>")  
          case 3 response.write("<td class='td-l fonttiny'> "& rsUpdateEntry("Rastreo") &"</td>")  
          case 4 response.write("<td class='td-l fonttiny'> "& rsUpdateEntry("Rastreo") &"</td>")  
          case 5 response.write("<td class='td-l fonttiny'> "& rsUpdateEntry("Rastreo") &"</td>")  
          case 6 response.write("<td class='td-l fonttiny'> "& rsUpdateEntry("Rastreo") &"</td>")  
          case 8 response.write("<td class='td-l fonttiny'> "& rsUpdateEntry("Rastreo") &"</td>")  
          case else  response.write("<td class='td-l fonttiny'> "& rsUpdateEntry("Rastreo") &"</td>")  
      end select
     


     if rsUpdateEntry("CargoCliente") =1 then
         response.write("<td class='td-c fonttiny'> si</td>")
     else
        response.write("<td class='td-c fonttiny'>no</td>")
     end if

     response.write("<td class='td-r fonttiny'>$ "& rsUpdateEntry("subtotal") &"</td>")
     response.write("<td class='td-c fonttiny'>"& rsUpdateEntry("Pedido") &"<a href='localexplorer:R:\"&rsUpdateEntry("RazonSocial")&"\"& rsUpdateEntry("Ruta") &"\" & rsUpdateEntry("Pedido")  &"'><img src='/img/repositorio.png' border=0 alt='repositorio' title='repositorio'></a></td>")

     if rsUpdateEntry("Consolidado") =1 then         
         response.write("<td class='td-c fonttiny'> si</td>")         
         response.write("<td class='td-r fonttiny'>$ "& rsUpdateEntry("subtotalconsolidado") &"</td>")
     else
        response.write("<td class='td-c fonttiny'>no</td>")        
        response.write("<td class='td-r fonttiny'>-</td>")
     end if

      response.write("<td class='td-c fonttiny'>"& rsUpdateEntry("factura") &"</td></tr>")
      response.write("<tr><td>&nbsp;</td>")
      response.write("<td class='td-l fonttiny' colspan=6><B>[Socio negocio:] </B>"& rsUpdateEntry("Cardname") &"</td>")
      response.write("<td class='td-l fonttiny' colspan=6><B>[Comentarios:] </B>"& rsUpdateEntry("Comentarios") &"</td><td colspan=2>&nbsp;</td>")
      response.write("</tr><tr><td colspan=15 height='30px'>&nbsp;</td></tr>")
    
     rsUpdateEntry.MoveNext
     registros=registros+1

  wend
end if

    if request("fPedido")<>"" then

        strSQL="select top 1 * from Notificaciones where Docnum=" &  request("fPedido")  &" and tipo='pedido' and modulo='ventas' order by DocDate desc" 
        rsUpdateEntry3.Open strSQL, adoCon4

        if not rsUpdateEntry3.EOF then
            RazonSocial=rsUpdateEntry3("RazonSocial")
        end if
        rsUpdateEntry3.close

        if vFleteTotal="" then
              vFleteTotal=0
        end if


      strSQL="select isnull(SUM(subtotal),0) as calculo,"& vFleteTotal &"-isnull(SUM(subtotal),0) as calculo2,case when SUM(subtotal)<"& vFleteTotal &" then 1 else 0 end as estatus  from envios where Pedido=" & request("fPedido")  &" and RazonSocial='" & RazonSocial &"' and CargoCliente=0 "
      'response.write strSQL
      rsUpdateEntry3.Open strSQL, adoCon4

      response.write("<tr><td colspan=10 class='td-r subtitulo'>Flete Gastado actual: </td><td class='td-l'>$ "& rsUpdateEntry3("calculo") &" MXN </td><td colspan=6>&nbsp;</td></tr>") 
        if rsUpdateEntry3("estatus")=0 then
          response.write("<tr><td colspan=10 class='td-r subtitulo'>Gasto de flete excedido por: </td><td class='td-l' width='130px'>$ <font color=red>"& rsUpdateEntry3("calculo2")&" MXN </font></td><td colspan=6>&nbsp;</td></tr>") 
        else
          response.write("<tr><td colspan=10 class='td-r subtitulo'>Remanente en flete por: </td><td class='td-l' width='130px'>$ <font color=green>"& rsUpdateEntry3("calculo2")&" MXN </font></td><td colspan=6>&nbsp;</td></tr>")
        end if

        strSQL="select 'viatico-'+convert(varchar,a.ID) as viatico,dbo.fn_GetMonthName(a.DocDate,'spanish') as Fecha,a.Beneficiario,a.MontoComprobado,a.Motivo,b.Pedido from VIATICOS_R1 a inner join ViaticoRemision b on a.ID=b.ID where b.Pedido=" & request("fPedido") & " group by a.ID,a.DocDate,a.Beneficiario,a.MontoComprobado,a.Motivo,b.Pedido"
        'response.write   strSQL       
        rsUpdateEntry4.Open strSQL, adoCon

        vMontoComprobado="0"
        while not rsUpdateEntry4.EOF 
             'response.write strSQL
             response.write("<tr><td class='td-r' colspan=2>" &  rsUpdateEntry4("viatico") &"</td>")
             
             response.write("<td class='td-c'>" &  rsUpdateEntry4("Fecha") &"</td>")
            
             response.write("<td class='td-r' colspan=7>" &  rsUpdateEntry4("beneficiario") &"</td>")
             response.write("<td class='td-r'>$ " &  rsUpdateEntry4("MontoComprobado") &"</td>")
             vMontoComprobado=vMontoComprobado & "+"& trim(rsUpdateEntry4("MontoComprobado") )
         
             response.write("<td class='td-c' colspan=2>" &  rsUpdateEntry4("Motivo") &"</td>")
             response.write("<td class='td-r'>" &  rsUpdateEntry4("Pedido") &"</td>")
             response.write("<td class='td-r colspan=3'>&nbsp;</td></tr>")
             rsUpdateEntry4.MoveNext
        wend
        rsUpdateEntry4.close
        'response.write("string|"& vMontoComprobado & "|<br>")
        

        strSQL="SELECT "& rsUpdateEntry3("calculo") &"+(" & vMontoComprobado &") as calculo"
        strSQL=replace(strSQL,"+)","+0)")
        'response.write   strSQL   
        rsUpdateEntry4.Open strSQL, adoCon

        response.write("<tr><td colspan=10 class='td-r subtitulo'>Fletes y vi&aacute;ticos: </td><td class='td-l'>$ "& rsUpdateEntry4("calculo") &" MXN </td><td colspan=6>&nbsp;</td></tr>")

        rsUpdateEntry4.close
        rsUpdateEntry3.close


     end if
 
    response.write("</table>") 

end sub




sub titulos %>

   <tr>
      <td class="td-c azul fonttiny">ID</td>
      <td class="td-c azul fonttiny">Colaborador</td>
      <td class="td-c azul fonttiny" width='80px'>Fecha</td>
      <td class="td-c azul fonttiny">RS</td>
      <td class="td-c azul fonttiny">Remision</td>     
      <td class="td-c azul fonttiny">Almac&eacute;n</td>
      <td class="td-c azul fonttiny">Portador</td>
      <td class="td-c azul fonttiny">Gu&iacute;a</td>
      <td class="td-c azul fonttiny">Rastreo</td>
      <td class="td-c azul fonttiny" width="80px">Cargo <br> cliente</td>
      <td class="td-c azul fonttiny">Subtotal</td>      
      <td class="td-c azul fonttiny">Pedido</td>
      <td class="td-c azul fonttiny">Consolidado</td>
      <td class="td-c azul fonttiny">Subtotal Consolidado</td>
      <td class="td-c azul fonttiny">Factura</td>
    </tr>

<%
end sub


sub DiaInhabil   'contenido 37' 
   if request("accion")<>"" then
      select case request("accion")
        case "1"
             msg="se ha agregado un dia inhabil"
             vfecha=date()
             vfecha=request("ffecha")
        
             strSQL="set dateformat dmy;select * from cat_DiaInhabil where fecha='" & day(vfecha)&"-"&month(vfecha)&"-"&year(vfecha) &"'"   
             'response.write strSQL  
             rsUpdateEntry.Open strSQL, adoCon

             if rsUpdateEntry.EOF then
                 if request("fcomentario")<>"" then
                     vcomentario=trim(request("fcomentario"))
                 else
                     vcomentario="no se indico comentario"
                 end if
                 strSQL="set dateformat dmy;insert into cat_DiaInhabil (fecha,comentario) values ('"& day(vfecha)&"-"&month(vfecha)&"-"&year(vfecha) &"','"& vcomentario &"')"     
                  'response.write( strSQL &"<br>")      
                  rsUpdateEntry2.Open strSQL, adoCon
             else
                 msg="no se registro, ya hay un registro parecido"
             end if
             rsUpdateEntry.close   

        case "2"
             msg="se ha editado un dia inhabil"
             vfecha=date()
             vfecha=request("ffecha")

             if request("fcomentario")<>"" then
                 strSQL="set dateformat dmy;update cat_DiaInhabil set comentario='"& request("fcomentario") &"' where fecha='"& day(vfecha)&"-"&month(vfecha)&"-"&year(vfecha) &"'"     
                 'response.write( strSQL &"<br>")      
                 rsUpdateEntry2.Open strSQL, adoCon
             else
                  msg="no se realizo ningun cambio"
             end if


        case "3"
             msg="se ha borrado un dia inhabil"
             vfecha=date()
             vfecha=request("ffecha")

             strSQL="set dateformat dmy;delete cat_DiaInhabil where fecha='"& day(vfecha)&"-"&month(vfecha)&"-"&year(vfecha) &"'"     
             'response.write( strSQL &"<br>")      
             rsUpdateEntry2.Open strSQL, adoCon

      end select      

   end if


   modulosHumRes
   response.write("<P style='font-size:12px'>&nbsp;</P>")   
   if request("msg")<>"" or msg<>"" then
           response.write("<P class='alert td-c'>"&  request("msg") & msg  &"</P>")
   end if    
   response.write("<H1>D&iacute;as inh&aacute;biles "& year(date()) &" + </H1>")   
   response.write("<P style='font-size:4px'>&nbsp;</P>")  

    

       strSQL="select *,dbo.fn_GetMonthName(fecha,'spanish') as Fecha2 from cat_DiaInhabil where year(fecha)>=year(getdate()) order by fecha"     
       'response.write strSQL  
       rsUpdateEntry.Open strSQL, adoCon
       
       response.write("<center><table width='360px' border=0 cellpadding=7 cellspacing=2 align=center>")
       response.write("<tr>")
       response.write("<td class=subtitulo>FECHA</td>")
       response.write("<td class=subtitulo>COMENTARIO</td>")     
       response.write("<td class=subtitulo>&nbsp;</td>")
       response.write("</tr>")
       
   
        while not rsUpdateEntry.EOF 
           response.write("<tr>")  
           response.write("<td class='td-r fonttiny'>"& rsUpdateEntry("fecha2") &"</td>")
           response.write("<td class='td-l fonttiny'><a href='/modules/ShowContent.asp?contenido=38&accion=2&fecha=" & day(rsUpdateEntry("fecha"))&"-"&month(rsUpdateEntry("fecha"))&"-"&year(rsUpdateEntry("fecha")) &"'>"& rsUpdateEntry("comentario") &"</a></td>")
           response.write("<td class='td-c fonttiny'><a href='/modules/ShowContent.asp?contenido=38&accion=3&fecha=" & day(rsUpdateEntry("fecha"))&"-"&month(rsUpdateEntry("fecha"))&"-"&year(rsUpdateEntry("fecha")) &"'><img src='/img/b_drop.png' border=0 with=11px height=11px alt=borrar title=borrar></td>")
           response.write("</tr>") 
           rsUpdateEntry.MoveNext
       wend
       rsUpdateEntry.close     %>

       <tr><td colspan=3 class=td-c>&nbsp;</td></tr> 
       <tr><td colspan=3 class=td-c><img src="/img/indicador.jpg" border=0 alt="agregar" title="agregar">
        <a href="/modules/ShowContent.asp?contenido=38&accion=1">  agregar un d&iacute;a inh&aacute;bil </a></td>
         
             <%
       response.write("</table>")    
       response.write("<P>&nbsp;</P>") 
       response.write("<P>&nbsp;</P>") 

end sub



sub DiaInhabilM
   

    if request("accion")="1" then
           response.write("<P style='font-size:12px'>&nbsp;</P>")  
           response.write("<H1>Agregar un d&iacute;a inh&aacute;bil</H1>")   
           response.write("<P style='font-size:8px'>&nbsp;</P>")  
           response.write("<center><table width='360px' border=0 cellpadding=7 cellspacing=2 align=center><form id='add' method='POST' action='ShowContent.asp'>")
           response.write("<input type='hidden' name='contenido' value=37><input type='hidden' name='accion' value=1><tr>")
           response.write("<td class=subtitulo>FECHA</td>")
           response.write("<td class=subtitulo>COMENTARIO</td>")              
           response.write("</tr>")

           response.write("<tr><td class='td-c fonttiny'><input type=date name='ffecha'></td>")
           response.write("<td class='td-l fonttiny'><input type='text' class='anchox2' name='fcomentario' value=''></td></tr>")
           response.write("<tr><td colspan=2 class='td-c fonttiny'><input type=submit value='agregar'></td>")


           response.write("</form></table>")

    end if

    if request("accion")="2" then
           vfecha=date()
           vfecha=request("fecha")

           strSQL="set dateformat dmy;select * from cat_DiaInhabil where fecha='" & day(vfecha)&"-"&month(vfecha)&"-"&year(vfecha) &"'"   
           'response.write strSQL  
           rsUpdateEntry.Open strSQL, adoCon

           response.write("<P style='font-size:12px'>&nbsp;</P>")  
           response.write("<H1>Editar comentario de un d&iacute;a inh&aacute;bil</H1>")   
           response.write("<P style='font-size:8px'>&nbsp;</P>")  
           response.write("<center><table width='360px' border=0 cellpadding=7 cellspacing=2 align=center><form id='add' method='POST' action='ShowContent.asp'>")
           response.write("<input type='hidden' name='contenido' value=37><input type='hidden' name='ffecha' value='"& day(vfecha)&"-"&month(vfecha)&"-"&year(vfecha) &"'><input type='hidden' name='accion' value=2><tr>")

           response.write("<td class=subtitulo>FECHA</td>")
           response.write("<td class=subtitulo>COMENTARIO</td>")              
           response.write("</tr>")

           response.write("<tr><td class='td-c fonttiny'>"& vfecha &"</td>")
           response.write("<td class='td-l fonttiny'><input type='text' class='anchox2' name='fcomentario' value='"& rsUpdateEntry("comentario")&"'></td></tr>")
           response.write("<tr><td colspan=2 class='td-c fonttiny'><input type=submit value='editar'></td>")


           response.write("</form></table>")

    end if


 if request("accion")="3" then
           vfecha=date()
           vfecha=request("fecha")

           strSQL="set dateformat dmy;select * from cat_DiaInhabil where fecha='" & day(vfecha)&"-"&month(vfecha)&"-"&year(vfecha) &"'"   
           'response.write strSQL  
           rsUpdateEntry.Open strSQL, adoCon

           response.write("<P style='font-size:12px'>&nbsp;</P>")  
           response.write("<H1>Borrar un d&iacute;a inh&aacute;bil</H1>")   
           response.write("<P style='font-size:8px'>&nbsp;</P>")  
           response.write("<center><table width='360px' border=0 cellpadding=7 cellspacing=2 align=center><form id='add' method='POST' action='ShowContent.asp'>")
           response.write("<input type='hidden' name='contenido' value=37><input type='hidden' name='ffecha' value='"& day(vfecha)&"-"&month(vfecha)&"-"&year(vfecha) &"'><input type='hidden' name='accion' value=3><tr>")

           response.write("<td class=subtitulo>FECHA</td>")
           response.write("<td class=subtitulo>COMENTARIO</td>")              
           response.write("</tr>")

           response.write("<tr><td class='td-c fonttiny'>"& vfecha &"</td>")
           response.write("<td class='td-l fonttiny'>"& rsUpdateEntry("comentario")&" </td></tr>")
           response.write("<tr><td colspan=2 class='td-c fonttiny'><input type=submit value='borrar'></td>")


           response.write("</form></table>")

    end if

end sub









sub CatHorarios  'contenido=39'

  modulosHumRes

  response.write("<P style='font-size:12px'>&nbsp;</P>")   
   if request("msg")<>"" or msg<>"" then
           response.write("<P class='alert td-c'>"&  request("msg") & msg  &"</P>")
   end if    
   response.write("<H1>Cat&aacute;logo de Horarios</H1>")   
   response.write("<P style='font-size:4px'>&nbsp;</P>")  

   strSQL="select *,substring( convert(varchar, CONVERT(TIME, inicio) ),1,8) as inicia,substring( convert(varchar, CONVERT(TIME,fin) ),1,8) as finaliza from Cat_Horarios order by ID"   
   'response.write strSQL  
   rsUpdateEntry.Open strSQL, adoCon

   response.write("<center><table width='700' border=0 cellpadding=5 cellspacing=2 align=center>")
    response.write("<tr>")
   response.write("<td class=subtitulo width='2%'>ID</td>")
   response.write("<td class=subtitulo width='30%'>Nombre del horario</td>")         
   response.write("<td class=subtitulo>Tiempo Total</td>")           
   response.write("<td class=subtitulo>Horas de comida</td>")              
   response.write("<td class=subtitulo>Inicio</td>")              
   response.write("<td class=subtitulo>Fin</td>")              
   response.write("</tr>")

   while not rsUpdateEntry.EOF
           response.write("<tr>")
           response.write("<td class='td-c fonttiny'>"& rsUpdateEntry("ID")  &"</td>")
           response.write("<td class='td-l fonttiny'>"& rsUpdateEntry("NombreHorario")&" </td>")
           response.write("<td class='td-c fonttiny'>"& rsUpdateEntry("TiempoTotal")&" </td>")
           response.write("<td class='td-c fonttiny'>"& rsUpdateEntry("Comida")&" </td>")
           response.write("<td class='td-c fonttiny'>"& rsUpdateEntry("inicia")&" </td>")
           response.write("<td class='td-c fonttiny'>"& rsUpdateEntry("finaliza")&" </td>")        
           response.write("</tr>")

       rsUpdateEntry.MoveNext
   wend
   rsUpdateEntry.close 
   response.write("</table>")



end sub  




sub Empleados   'contenido 40   ELY'
   
        modulosHumRes

       response.write("<P style='font-size:12px'>&nbsp;</P>")   
       if request("msg")<>"" or msg<>"" then
               response.write("<P class='alert td-c'>"&  request("msg") & msg  &"</P>")
       end if    
 
       strSQL="select count(*) as calculo from Empleados where FechaEgreso is null"
       rsUpdateEntry.Open strSQL, adoCon

       response.write("<center><H1><B>CATALOGO DE COLABORADORES ACTIVOS ("& rsUpdateEntry("calculo") &")</B></H1>")
       response.write("<P style='font-size:1px'>&nbsp;</P>") 
       rsUpdateEntry.close

       
       strSQL="select a.*,b.Departamento,c.Puesto,d.sucursalShort,round(cast( DATEDIFF(day,a.FechaNacimiento,getdate() ) as float)/365.242199,1) as edad,round(cast( DATEDIFF(day,a.FechaIngreso,getdate() ) as float)/365.242199,1) as antiguedad from Empleados a inner join cat_departamento b on a.Clave_depto=b.Clave_depto "
       strSQL= strSQL & "inner join cat_puesto c on a.clave_puesto=c.clave_puesto inner join cat_sucursal d on a.clave_sucursal=d.clave_sucursal  where a.FechaEgreso is null order by b.Clave_depto,d.clave_sucursal"     
       'response.write strSQL  
       rsUpdateEntry.Open strSQL, adoCon
       
       response.write("<center><table width='1200px' border=0 cellpadding=5 cellspacing=2 align=center>")
       response.write("<tr>")
       response.write("<th class=subtitulo>ID</td>")
       response.write("<th class=subtitulo>COLABORADOR</td>")
       response.write("<th class=subtitulo>AREA</td>")
       response.write("<th class=subtitulo>SUC</td>")
       response.write("<th class=subtitulo>NACIMIENTO</td>")
       response.write("<th class=subtitulo>EDAD</td>")
       response.write("<th class=subtitulo>INGRESO</td>")
       response.write("<th class=subtitulo>ANTIGUEDAD</td>")
       response.write("<th class=subtitulo>NSS</td>")
       response.write("<th class=subtitulo>&nbsp;</td>")
       response.write("</tr>")
       
   
        while not rsUpdateEntry.EOF 
           response.write("<tr>")  %>
           <td class='td-r fonttiny'><a href="/modules/ShowContent.asp?contenido=40&accion=3&item=<%=rsUpdateEntry("ID")%>">  <%=rsUpdateEntry("ID")%> </a></td> 
           <%
                     
                response.write("<td class='fonttiny'>"& rsUpdateEntry("nombres") &" "  &rsUpdateEntry("PATERNO") &" " &rsUpdateEntry("MATERNO")   &"</td>")
                response.write("<td class='fonttiny'>"& rsUpdateEntry("departamento") &"</td>")
                response.write("<td class='fonttiny'>"& rsUpdateEntry("SucursalShort") &"</td>")  
                response.write("<td  class='td-c fonttiny'>"& day(rsUpdateEntry("FechaNacimiento")) &"-" & meses(month(rsUpdateEntry("FechaNacimiento"))) &"-"&  year(rsUpdateEntry("FechaNacimiento"))  &"</td>")
                response.write("<td class='fonttiny'>"& rsUpdateEntry("edad") &"</td>")  
                response.write("<td  class='td-c fonttiny'>"& day(rsUpdateEntry("FechaIngreso")) &"-" & meses(month(rsUpdateEntry("FechaIngreso"))) &"-"&  year(rsUpdateEntry("FechaIngreso"))  &"</td>")
                response.write("<td class='td-c fonttiny'>"& rsUpdateEntry("antiguedad") &"</td>") 
                response.write("<td  class='td-c fonttiny'>"& rsUpdateEntry("NSS") &"</td>") 
                
                             
           
           response.write("<td>") %> <a href="/modules/ShowContent.asp?contenido=40&accion=2&item=<%=rsUpdateEntry("ID")%>">
           <img src='/img/b_drop.png' border=0 with=11px height=11px alt="baja colaborador" title="baja colaborador"></td>    <%        
           response.write("</tr>")
           rsUpdateEntry.MoveNext
       wend
       rsUpdateEntry.close     %>
       <tr><td colspan=10 class=separador>&nbsp;</td></tr> 
       <tr><td colspan=10 class=td-c><img src="/img/indicador.jpg" border=0 alt="agregar colaborador" title="agregar colaborador">
        <a href="/modules/ShowContent.asp?contenido=40&accion=1">  agregar un colaborador </a></td> </tr>  
             <%
       response.write("</table>")    
       response.write("<P>&nbsp;</P>") 
       response.write("<P>&nbsp;</P>") 
       response.write("<P>&nbsp;</P>") 
       response.write("<P>&nbsp;</P>") 
       response.write("<P>&nbsp;</P>") 


end sub 






sub checada   'contenido 41'


if request("ID")<>"" then
    strSQL="set dateformat dmy;select * from checadas where day(fecha)=day(getdate()) and month(fecha)=month(getdate()) and year(fecha)=year(getdate()) and id="& request("ID")
    'response.write strSQL
    rsUpdateEntry.Open strSQL, adoCon

    if rsUpdateEntry.EOF then
       strSQL="set dateformat dmy;insert into checadas (ID,fecha) values ("&  request("ID") &",'"& day(date()) &"-"& month(date()) &"-"& year(date())  &"')"
       rsUpdateEntry2.Open strSQL, adoCon
    end if
    rsUpdateEntry.close
end if
  

if request("ftipo")<>"" then

    if request("ftipo")="entrada" then
         strSQL="set dateformat dmy;select * from checadas where ID=" & request("ID") &" and fecha='"& day(date()) &"-"& month(date()) &"-"& year(date()) &"' and entrada is not null "
         rsUpdateEntry3.Open strSQL, adoCon
         if rsUpdateEntry3.EOF then
                if session("session_id")="RCARDONA" or session("session_id")="DVAZQUEZ" then
                   strSQL="set dateformat dmy;update checadas set entrada=dateadd(HOUR,2, dateadd(Mi,-2,GETDATE()) ) where fecha='"& day(date()) &"-"& month(date()) &"-"& year(date()) &"' and id="& request("ID")
                else
                   strSQL="set dateformat dmy;update checadas set entrada=dateadd(Mi,-2,GETDATE())  where fecha='"& day(date()) &"-"& month(date()) &"-"& year(date()) &"' and id="& request("ID")
                end if
                rsUpdateEntry4.Open strSQL, adoCon
                response.write("<P class='td-c alert'>se registro entrada</P>")
         else
                response.write("<P class='td-c alert'>ya existe una entrada de personal existente, NO SE REGISTRO</P>")
         end if
         rsUpdateEntry3.close
    end if


    if request("ftipo")="salida" then
         strSQL="set dateformat dmy;select * from checadas where ID=" & request("ID") &" and fecha='"& day(date()) &"-"& month(date()) &"-"& year(date()) &"' and salida is not null "
         rsUpdateEntry3.Open strSQL, adoCon

         if rsUpdateEntry3.EOF then
         	    if session("session_id")="RCARDONA" or session("session_id")="DVAZQUEZ" then
         	    	  strSQL="set dateformat dmy;update checadas set salida=dateadd(HOUR,2, dateadd(Mi,-2,GETDATE()) ) where fecha='"& day(date()) &"-"& month(date()) &"-"& year(date()) &"' and id="& request("ID")
         	    else
                      strSQL="set dateformat dmy;update checadas set salida=  dateadd(Mi,-2,GETDATE())   where fecha='"& day(date()) &"-"& month(date()) &"-"& year(date()) &"' and id="& request("ID")
         	    end if
                rsUpdateEntry4.Open strSQL, adoCon
                response.write("<P class='td-c alert'>se registro salida</P>")
         else
                response.write("<P class='td-c alert'>ya existe una salida de personal existente, NO SE REGISTRO</P>")
         end if
         rsUpdateEntry3.close

    end if
    'response.write strSQL
end if



strSQL="select nombres+' '+paterno as colaborador,* from Empleados where FechaEgreso is null and id_usuario='" & Session("session_id") &"'"
 'response.write strSQL  
rsUpdateEntry.Open strSQL, adoCon
response.write ("<P>&nbsp;</P>")

        Doalert
        titulo="REGISTRO DE PERSONAL DE ENTRADA/SALIDA<BR> Colaborador: " & rsUpdateEntry("colaborador")
        DoTitulo(titulo)

        CreateTable(360)
      %>

   
     <form id="checada" method="POST" action="ShowContent.asp"> 
     <input type="hidden" name="contenido" value="41">
     <input type="hidden" name="ID" value="<%=rsUpdateEntry("ID")%>">

     <tr><td class='fonttiny subtitulo td-r'>Fecha: </td><td class='fonttiny td-l'><%=date()%></td></tr>
     <tr><td class='fonttiny subtitulo td-r' width="40%">Tipo: </td><td class='fonttiny td-l' width="60%"><select id="ftipo" name="ftipo"> <option value="entrada">entrada</option> <option value="salida">salida</option> </select></td></tr>
     
     <tr><td class='fonttiny' colspan=2>&nbsp; </td></tr>
     <tr><td class='fonttiny td-c' colspan=2><input type="submit" value="registrar"> </td></tr>

     </form>
     </table>
<%
     response.write ("<P>&nbsp;</P>")
     titulo="Registros &uacute;ltima quincena"
     DoTitulo(titulo)
     
     strSQL="select top 17 .dbo.fn_GetMonthName(b.Fecha,'spanish') as Fecha,b.Entrada,b.Salida,case when DATEDIFF(Mi,DATEADD(HOUR,8,cast( cast(Entrada as Date) as DateTime)   ), Entrada  )>0 then 'R' else ' ' end as Estatus  from Empleados a inner join checadas b on a.ID=b.ID where a.FechaEgreso is null and a.id_usuario='"& Session("session_id") &"'  order by b.fecha desc"     
     'response.write strSQL
     rsUpdateEntry2.Open strSQL, adoCon
     rsUpdateEntry3.Open strSQL, adoCon
     CreateTable(420)
     
         Dim Campos(4) 
         i=1
         RowIn
         For Each rsUpdateEntry2 in rsUpdateEntry2.Fields            
            Response.Write "<td class='subtitulo td-c'>" & rsUpdateEntry2.Name & "</td>"
            campos(i)=rsUpdateEntry2.Name
            i=i+1           
         Next        
         RowOut

         while not rsUpdateEntry3.EOF
             rowin
             for i=1 to 4
                   Response.write("<td class='td-r fonttiny'>"& rsUpdateEntry3(campos(i)) &"</td>")
             next             
             rsUpdateEntry3.MoveNext
             RowOut
         wend         
         rsUpdateEntry3.close

      CloseTable
      

     


if request("ID")<>"" then
     response.write("<P>&nbsp;</P><P>&nbsp;</P><P>&nbsp;</P><P>&nbsp;</P><P>&nbsp;</P><P>&nbsp;</P><P>&nbsp;</P>")
     response.write("<P>&nbsp;</P><H1>&Uacute;ltimas checadas de "& rsUpdateEntry("colaborador") &"</H1><table border=0 width='600px' cellpadding=5 cellpadding=2 align='center'>")

     response.write("<tr><td class='td-c azul fonttiny'>Fecha</td><td class='td-c azul fonttiny'>Entrada</td><td class='td-c azul fonttiny'>Salida</td></tr>")


     strSQL="select dbo.fn_GetMonthName(fecha,'spanish') as Fecha2,* from checadas where datediff(DAY,FECHA,getdate())<20 and id="& request("ID") &" order by fecha desc"
     rsUpdateEntry3.Open strSQL, adoCon

     while not rsUpdateEntry3.EOF
        response.write("<tr><td class='td-c fonttiny'>" &   rsUpdateEntry3("fecha2") &"</td>")
        response.write("<td class='td-c fonttiny'>" &   rsUpdateEntry3("entrada") &"</td><td class='td-c fonttiny'>" &   rsUpdateEntry3("salida") &"</td></tr>")
        rsUpdateEntry3.MoveNext
     wend
     rsUpdateEntry3.close

     response.write("</table>")
end if
 

     rsUpdateEntry.close
end sub




sub ModuloADD   'contenido 42'   

    msg=""
    if request("borrar")<>"" and request("modulo") <>"" then
        strSQL="delete permisos where modulo='"& request("modulo") &"' and id_usuario='"& request("borrar") &"'  "  
        rsUpdateEntry.Open strSQL, adoCon              
        msg="se elimino a "& request("borrar") &" permiso de modulo: " & request("modulo")
    end if


    if request("fusuario") <>"" then  
        strSQL="insert into permisos (id_usuario,modulo) values ('" &  request("fusuario") &"','"& request("modulo")   &"')"
        rsUpdateEntry.Open strSQL, adoCon
  
        msg="se registro un permiso para un modulo de intranet"
    end if

    if request("fmodulo") <>"" then  
        strSQL="insert into modulos (modulo,LogDate) values ('" &  request("fmodulo") &"',getdate() )"
        rsUpdateEntry.Open strSQL, adoCon
  
        msg="se registro un modulo de intranet"
    end if

    response.write("<P>&nbsp;</P>")
    if msg<>"" then
           response.write("<P class='alert td-c'>"&  msg &"</P>")
    end if

%>

     
     <H1> REGISTRO DE UN MODULO DE INTRANET</H1>
     <P>&nbsp;</P>

     <table border=1 width="900px" cellpadding="5" cellpadding="2" align="center">
     <tr><td width="38%" class="td-c">

     <form id="modulos" method="POST" action="ShowContent.asp"> 
     <input type="hidden" name="contenido" value="42">
   
     Nombre del m&oacute;dulo: &nbsp;&nbsp;&nbsp; <input type="text" name="fmodulo" value=""> <P>&nbsp;</P> <input type="submit" value="registrar"> </td>

     </form>

     </td><td width="62%" class="td-l"> <B> M&Oacute;DULOS REGISTRADOS  </b> <br><br> <%

      strSQL="delete permisos where id_usuario in (select id_usuario from Empleados where FechaEgreso is not null)"
      rsUpdateEntry.Open strSQL, adoCon

      strSQL="select *,dbo.fn_GetMonthName(LogDate,'spanish') as Fecha from modulos order by ID" 
      rsUpdateEntry.Open strSQL, adoCon

      while not rsUpdateEntry.EOF
        response.write( "* <font class='subtitulo fonttiny'>" & rsUpdateEntry("modulo") &":</font> <a href='ShowContent.asp?contenido=43&modulo=" & rsUpdateEntry("modulo") &"'><img src='/img/mas.png' border=0 height='9px' width='9px' alt='' title=''></a> &nbsp;&nbsp;" )
        strSQL="select * from permisos where modulo='"& rsUpdateEntry("modulo") &"'"
        rsUpdateEntry2.Open strSQL, adoCon

        while not rsUpdateEntry2.EOF
            response.write( rsUpdateEntry2("id_usuario") & "<a href='ShowContent.asp?contenido=42&modulo="& rsUpdateEntry("modulo")&"&borrar="& rsUpdateEntry2("id_usuario") &" '><img src='/img/b_drop.png' border=0 with=11px height=11px alt=eliminar title=eliminar> </a>, " )
            rsUpdateEntry2.MoveNext
        wend
        rsUpdateEntry2.close

        response.write( "<br>")
        rsUpdateEntry.movenext
      wend
      rsUpdateEntry.close

      response.write("</td></tr> </table>")
  

end sub



sub PermisoADD  'contenido 43'
    if request("modulo") <>"" then
  
      %>
     <P>&nbsp;</P>
     
     <H1> REGISTRO DE UN PERMISO PARA MODULO: <%=request("modulo")%> </H1>
     <P>&nbsp;</P>

     <table border=1 width="900px" cellpadding="5" cellpadding="2" align="center">
     <tr><td width="45%" class="td-c">

     <form id="modulos" method="POST" action="ShowContent.asp"> 
     <input type="hidden" name="contenido" value="42">
     <input type="hidden" name="modulo" value="<%=request("modulo")%>">
   
     Elija un colaborador para permitir el uso del <BR>
     m&oacute;dulo: &nbsp;&nbsp;&nbsp; <select name="fusuario" class="anchox1"> 
     <%
       strSQL="select id_usuario,nombres+' '+paterno as colaborador from Empleados where FechaEgreso is null and id_usuario<>'' and id_usuario not in (select id_usuario from permisos where modulo='"& request("modulo") &"')"
       rsUpdateEntry.Open strSQL, adoCon

       while not rsUpdateEntry.EOF
        response.write("<option value='"& rsUpdateEntry("id_usuario") &"'>"& rsUpdateEntry("colaborador")  &"</option>")
        rsUpdateEntry.movenext
       wend
       rsUpdateEntry.close

     %>
      </select>
      <P>&nbsp;</P> <input type="submit" value="registrar"> </td>

     </form>

     </td><td width="55%" class="td-l"> <B> USUARIOS CON PERMISO  </b> <br><br> <%

      strSQL="select *,dbo.fn_GetMonthName(LogDate,'spanish') as Fecha from modulos where modulo='"& request("modulo") &"' " 
      rsUpdateEntry.Open strSQL, adoCon

      while not rsUpdateEntry.EOF
        response.write( "* <font class='subtitulo'>" & rsUpdateEntry("modulo") &":</font> <a href='ShowContent.asp?contenido=43&modulo=" & rsUpdateEntry("modulo") &"'><img src='/img/mas.png' border=0 height='9px' width='9px' alt='' title=''></a> &nbsp;&nbsp;" )
        strSQL="select * from permisos where modulo='"& rsUpdateEntry("modulo") &"'"
        rsUpdateEntry2.Open strSQL, adoCon

        while not rsUpdateEntry2.EOF
            response.write( rsUpdateEntry2("id_usuario")&", " )
            rsUpdateEntry2.MoveNext
        wend
        rsUpdateEntry2.close

        response.write( "<br>")
        rsUpdateEntry.movenext
      wend
      rsUpdateEntry.close

      response.write("</td></tr> </table>")
  




    end if
end sub




sub ChecarIntranet 'contenido 44
        
      
        titulo="Muestra las ultimas 100 checadas HomeOffice"
        DoTitulo(titulo)

        %>
    

     <table border=1 width="700px" cellpadding="5" cellpadding="2" align="center">
     <%
       vfecha=date()

       strSQL="select top 100 a.*,b.Nombres+' '+b.Paterno as colaborador from checadas a inner join Empleados b on a.ID=b.ID order by a.Entrada desc"
       rsUpdateEntry.Open strSQL, adoCon

       while not rsUpdateEntry.EOF
        vfecha=rsUpdateEntry("Fecha")
        response.write("<tr>")
        response.write("<td class='td-r'>" & rsUpdateEntry("ID") & "</td>")
        response.write("<td class='td-c'>" & rsUpdateEntry("Fecha") & "</td>")
        response.write("<td class='td-c'>" & rsUpdateEntry("Entrada") & "</td>")
        response.write("<td class='td-c'>" & rsUpdateEntry("Salida") & "</td>")
        response.write("<td class='td-l'>" & rsUpdateEntry("colaborador") & "</td>")
        response.write("</tr>")
        rsUpdateEntry.movenext
        if not rsUpdateEntry.EOF then
             if vfecha<>rsUpdateEntry("Fecha") then
                  response.write("<tr><td colspan=5 class='separador'> &nbsp; </td></tr>")
             end if
        end if
       wend
       rsUpdateEntry.close

       response.write("</table>")
  
end sub



sub titulosPeriodosVacacion
	   response.write("<tr>")
       response.write("<th class='subtitulo fonttiny'>ID</td>")
       response.write("<th class='subtitulo fonttiny'>Colaborador</td>")    'a 16 caracteres'
       response.write("<th class='subtitulo fonttiny'>SUC</td>")
       response.write("<th class='subtitulo fonttiny'>Ingreso</td>")
       response.write("<th class='subtitulo fonttiny'>A&ntilde;os</td>")
          
       response.write("<th class='subtitulo fonttiny'>A&ntilde;os <br> x periodo</td>")
       response.write("<th class='subtitulo fonttiny'>Dias <br> Derecho</td>")
       response.write("<th class='subtitulo fonttiny'>Dias <br> Tomados</td>")
       response.write("<th class='subtitulo fonttiny'>Dias <br> Pdtes </td>")
       response.write("<th class='subtitulo fonttiny'>Dias <br> Progr</td>")

       response.write("<th class='subtitulo fonttiny'>A&ntilde;os <br> x periodo</td>")
       response.write("<th class='subtitulo fonttiny'>Dias <br> Derecho</td>")
       response.write("<th class='subtitulo fonttiny'>Dias <br> Tomados</td>")
       response.write("<th class='subtitulo fonttiny'>Dias <br> Pdtes </td>")
       response.write("<th class='subtitulo fonttiny'>Dias <br> Progr</td>")

       response.write("<th class='subtitulo fonttiny'>A&ntilde;os <br> x periodo</td>")
       response.write("<th class='subtitulo fonttiny'>Dias <br> Derecho</td>")
       response.write("<th class='subtitulo fonttiny'>Dias <br> Tomados</td>")
       response.write("<th class='subtitulo fonttiny'>Dias <br> Pdtes </td>")
       response.write("<th class='subtitulo fonttiny'>Dias <br> Progr</td>")
       response.write("</tr>")

end sub


sub vacaciones 'contenido 45'

       modulosHumRes

       response.write("<P style='font-size:12px'>&nbsp;</P>")   
       if request("msg")<>"" or msg<>"" then
               response.write("<P class='alert td-c'>"&  request("msg") & msg  &"</P>")
       end if    
 
       strSQL="select count(*) as calculo from Empleados where FechaEgreso is null"
       rsUpdateEntry.Open strSQL, adoCon

       titulo="PERIODOS VACACIONALES COLABORADORES ACTIVOS ("& rsUpdateEntry("calculo")  &")"
       DoTitulo(titulo)

       rsUpdateEntry.close

       
       strSQL="select a.*,b.sucursalShort,round(cast( DATEDIFF(day,a.FechaIngreso,getdate() ) as float)/365.242199,1) as antiguedad from Empleados a inner join cat_sucursal b on a.clave_sucursal=b.clave_sucursal where a.FechaEgreso is null order by month(a.FechaIngreso),day(a.FechaIngreso)"     
       'response.write strSQL  
       rsUpdateEntry.Open strSQL, adoCon
       
       response.write("<center><table width='1200px' border=1 cellpadding=2 cellspacing=2 align=center>")
       if month(date())>=10 then
           response.write("<tr><td colspan=5>&nbsp;</td><td colspan=5 class='td-c subtitulo'>"& year(date())-1 &"</td><td colspan=5 class='td-c subtitulo'>"& year(date()) &"</td><td colspan=5 class='td-c subtitulo'>"& year(date())+1 &"</td></tr>") 
       else
          response.write("<tr><td colspan=5>&nbsp;</td><td colspan=5 class='td-c subtitulo'>"& year(date())-2 &"</td><td colspan=5 class='td-c subtitulo'>"& year(date())-1 &"</td><td colspan=5 class='td-c subtitulo'>"& year(date()) &"</td></tr>")
       end if

      
       titulosPeriodosVacacion
       
   
        while not rsUpdateEntry.EOF      

                response.write("<tr>")  
                response.write("<td class='fonttiny separador td-c strong'>"& rsUpdateEntry("ID") &"</td>")
                vcolaborador=  rsUpdateEntry("nombres") &" "  &rsUpdateEntry("PATERNO") &" " &rsUpdateEntry("MATERNO")              
                response.write("<td class='Fonttiny separador'>"&  vcolaborador &"</td>")                
                response.write("<td class='Fonttiny separador'>"& rsUpdateEntry("SucursalShort") &"</td>")                 
                response.write("<td  class='td-r fonttiny separador'>"& day(rsUpdateEntry("FechaIngreso")) &"-" & meses(month(rsUpdateEntry("FechaIngreso"))) &"-"&  year(rsUpdateEntry("FechaIngreso"))  &"</td>")
                response.write("<td class='td-c fonttiny separador'>"& rsUpdateEntry("antiguedad") &"</td>") 

                if month(date())>10 then
                    FOR vanio=year(date())-1 to year(date())+1      
                        strSQL="select *,DiasDerecho-DiasTomados as Pendientes from PeriodosVacacionales where ID=" &  rsUpdateEntry("ID") &" and Periodo=" & vanio
	                    'response.write strSQL
                        rsUpdateEntry2.Open strSQL, adoCon              	
                        showPeriodosVacacioneles
                        rsUpdateEntry2.close  
                    NEXT
                else
                    FOR vanio=year(date())-2 to year(date())
                    	strSQL="select *,DiasDerecho-DiasTomados as Pendientes from PeriodosVacacionales where ID=" &  rsUpdateEntry("ID") &" and Periodo=" & vanio
	                    'response.write strSQL
                        rsUpdateEntry2.Open strSQL, adoCon              	
                        showPeriodosVacacioneles
                        rsUpdateEntry2.close  
                    NEXT
                end if
        

                response.write("</tr>")
                'response.write("<tr><td colspan=15 class='separador'>&nbsp;</td></tr>") 
           rsUpdateEntry.MoveNext
           
       wend
       rsUpdateEntry.close    
       response.write("</table>")    

       response.write("<div id='BoxRoundedDetailTop'>&nbsp;</div>")
       response.write("<div id='BoxRoundedDetail2'>&nbsp;</div>")

       response.write("<P>&nbsp;</P>")         

       
       response.write("<P>&nbsp;</P>") 

end sub


sub showPeriodosVacacioneles	                

                    if rsUpdateEntry2.EOF then
                        response.write("<td class='td-c fonttiny separador'>-</td>") 
                        response.write("<td class='td-c fonttiny separador'>-</td>") 
                        response.write("<td class='td-c fonttiny separador'>-</td>") 
                        response.write("<td class='td-c fonttiny separador'>-</td>")                         
                        response.write("<td class='td-c fonttiny separador'>-</td>")  
                    else
                        if (vanio=year(date())-1 and month(rsUpdateEntry("FechaIngreso"))>month(date())  )  OR (vanio=year(date()) and month(rsUpdateEntry("FechaIngreso"))<=month(date())  )  then
                            response.write("<td class='td-c fonttiny separador'>"& rsUpdateEntry2("AniosAntiguedad") &"</td>") 
                            response.write("<td class='td-c fonttiny naranja separador'>"& rsUpdateEntry2("DiasDerecho") &"</td>") 
                            response.write("<td class='td-c fonttiny separador'>")   %>
                            <a href="javascript:sendReq('/modules/VacacionDetails.asp','ID,periodo,accion','<%=rsUpdateEntry("ID")%>,<%=rsUpdateEntry2("Periodo")%>,1','BoxRoundedDetailTop')"><U>
                            <%  
                            response.write( rsUpdateEntry2("DiasTomados") &"</U></a></td>") 
                            response.write("<td class='td-c fonttiny separador'><B>"& rsUpdateEntry2("Pendientes") &"</B></td>") 
                            response.write("<td class='td-c fonttiny separador'>"& rsUpdateEntry2("DiasProgramados") &"</td>") 
                        else
                            response.write("<td class='td-c fonttiny separador'>"& rsUpdateEntry2("AniosAntiguedad") &"</td>") 
                            response.write("<td class='td-c fonttiny separador'>"& rsUpdateEntry2("DiasDerecho") &"</td>") 
                            response.write("<td class='td-c fonttiny separador'>")   %>
                            <a href="javascript:sendReq('/modules/VacacionDetails.asp','ID,periodo,accion','<%=rsUpdateEntry("ID")%>,<%=rsUpdateEntry2("Periodo")%>,1','BoxRoundedDetailTop')"><U>
                            <%  
                            response.write( rsUpdateEntry2("DiasTomados") &"</td>") 
                            if vanio=year(date())-1 and rsUpdateEntry2("Pendientes")>0  then
                                   response.write("<td class='td-c fonttiny separador'><B><font color=red>"& rsUpdateEntry2("Pendientes") &"</font></B></td>") 
                            else
                                   response.write("<td class='td-c fonttiny separador'><B>"& rsUpdateEntry2("Pendientes") &"</B></td>") 
                            end if
                            response.write("<td class='td-c fonttiny separador'>"& rsUpdateEntry2("DiasProgramados") &"</td>") 

                         end if
                    end if                    


end sub


sub SolicVacaciones   'contenido 46'

Doalert

if request("action")=1 then
    vMsg="su solicitud de vacaciones se cancelo, ya se aviso a capital humano"
    Msg="su solicitud de vacaciones se cancelo, ya se aviso a capital humano"

    strSQL="exec SP_CancelarVacaciones "& request("ID") &","& request("periodo") &",'"& request("fecha")  &"'"
    rsUpdateEntry.Open strSQL, adoCon

    strSQL="set dateformat ymd;update SolicitudVacaciones set Activo=0,CancelDate=getdate() where ID="&request("ID") &" and fecha='"& request("fecha") &"' and periodo=" &request("periodo") 
    rsUpdateEntry.Open strSQL, adoCon

    strSQL="update PeriodosVacacionales set DiasProgramados=(select DiasProgramados from PeriodosVacacionales where ID="&request("ID") &" and Periodo="&request("periodo") &")-1 where ID="&request("ID") &" and Periodo="&request("periodo") 
    rsUpdateEntry.Open strSQL, adoCon

end if


strSQL="select nombres+' '+paterno as colaborador,*,dbo.fn_GetMonthName(FechaIngreso,'spanish') as Fecha2 from Empleados where FechaEgreso is null and id_usuario='" & Session("session_id") &"'"
'response.write("<BR>"& strSQL  )
rsUpdateEntry.Open strSQL, adoCon


strSQL="select *,DiasDerecho-DiasTomados-DiasProgramados as disponibles from PeriodosVacacionales where ID="& rsUpdateEntry("ID") &" and DiasTomados<(DiasDerecho+DiasProgramados) order by Periodo" 
'response.write("<BR>"& strSQL  ) 
rsUpdateEntry2.Open strSQL, adoCon


%>

<P>&nbsp;</P> 
  <%
  Doalert
  Titulo="Solicitar vacaciones para: "& rsUpdateEntry("colaborador") &"<BR><font class='fonttiny'>Ingreso: " & rsUpdateEntry("fecha2") &"</font>"   
  DoTitulo(titulo)   
%>


     <P>&nbsp;</P>

     <table border=0 width="638px" cellpadding="5" cellpadding="2" align="center">
     <form id="checada" method="POST" action="ShowContent.asp"> 
     <input type="hidden" name="contenido" value="47">
     <input type="hidden" name="ID" value="<%=rsUpdateEntry("ID")%>">

     <tr>
         <td class='fonttiny subtitulo td-c'>Per&iacute;odo </td>
         <td class='fonttiny subtitulo td-c'>Antig&utilde;edad </td>
         <td class='fonttiny subtitulo td-c'>D&iacute;as <br> derecho </td>
         <td class='fonttiny subtitulo td-c'>D&iacute;as  <br> tomados </td>
         <td class='fonttiny subtitulo td-c'>D&iacute;as  <br> programados </td>
         <td class='fonttiny subtitulo td-c'>D&iacute;as  <br> disponibles </td>
     </tr>
   
   <%
    vDisponible=0
    i=1
    while not rsUpdateEntry2.EOF        
        response.write("<tr><td class='td-c fonttiny'>" &   rsUpdateEntry2("Periodo") &"</td>")
        response.write("<td class='td-c fonttiny'>" &   rsUpdateEntry2("AniosAntiguedad") &"</td>")
        response.write("<td class='td-c fonttiny'>" &   rsUpdateEntry2("DiasDerecho") &"</td>")
        response.write("<td class='td-c fonttiny'>" &   rsUpdateEntry2("DiasTomados") &"</td>")
        response.write("<td class='td-c fonttiny'>" &   rsUpdateEntry2("DiasProgramados") &"</td>")
        response.write("<td class='td-c fonttiny'>" &   rsUpdateEntry2("disponibles") &"</td></tr>")
        if i=1 and rsUpdateEntry2("disponibles")>0 then
             vDisponible=rsUpdateEntry2("disponibles") 
             vPeriodo=rsUpdateEntry2("Periodo") 
             i=2
        end if
        rsUpdateEntry2.MoveNext
    wend
    rsUpdateEntry2.close

    response.write("<input type='hidden' name='Diasdisponibles' value="& vDisponible &">")
    response.write("<input type='hidden' name='Periodosdisponibles' value="& i &">")
    response.write("<input type='hidden' name='Periodo' value="& vPeriodo &">")

    response.write("<tr><td class='fonttiny' colspan=6>&nbsp; </td></tr>")
    response.write("<tr><td class='td-r fonttiny subtitulo' colspan=3>D&iacute;as disponibles periodo m&aacute;s antiguo, seleccione cuantos necesita: </td><td class='td-l fonttiny' colspan=3><select name='fdias'>")

    for n=vDisponible to 1 step -1     
          if  n=1  then    
              response.write("<option value="& n & " SELECTED>" & n &"</option>")  
          else
              response.write("<option value="& n & ">" & n &"</option>")
          end if
    next  
    response.write("</select></td></tr>")

   %>  
     
     <tr><td class='fonttiny td-c' colspan=6><font class='alert'>Solicite los dias de vacaciones en la misma quincena <br>(si su dia esta en una quincena posterior, vuelva a hacer otra solicitud)</font> </td></tr>
     <tr><td class='fonttiny td-c' colspan=6><input type="submit" value="solicitar"> </td></tr>

     </form>
     </table>
     <h1>Re-imprima formatos de solicitudes</h1>  <%
     strSQL="select top 6 (identificador),cast(Logdate as Date) as Fecha from SolicitudVacaciones  where ID="& rsUpdateEntry("ID") &" and identificador is not null group by identificador,cast(Logdate as Date)  order by identificador desc"
     'response.write strSQL
     rsUpdateEntry5.Open strSQL, adoCon

     while not rsUpdateEntry5.EOF
         response.write("<a href='FormatoVacaciones.asp?ID="& rsUpdateEntry("ID") &"&identificador="& rsUpdateEntry5("identificador")&"' target='formato'>Imprima solicitud hecha en: " & rsUpdateEntry5("fecha") &"</a><br>")
         rsUpdateEntry5.MoveNext
     wend
     rsUpdateEntry5.close   
      %>


   <P>&nbsp;</P><P>&nbsp;</P>
   <table border=1 width="1400px">
    <tr><td class="td-c" style="vertical-align: top;">
    <H1> Solicitud(s) de vacaciones recibidas </H1><br>

     <table border=0 width="600px" cellpadding="5" cellpadding="2" align="center">
      <tr>
         <td class='fonttiny subtitulo td-c'>Fecha de vaciones </td>
         <td class='fonttiny subtitulo td-c'>Per&iacute;odo </td>
         <td class='fonttiny subtitulo td-c'>Fecha de la solicitud </td>
         <td class='fonttiny subtitulo td-c'>Estatus </td>
         <td class='fonttiny subtitulo td-c'>Tipo</td>
         <td class='fonttiny subtitulo td-c'>&nbsp;</td>
         
     </tr>

<%

 strSQL="select a.LogDate,a.periodo,a.activo,dbo.fn_GetMonthName(a.fecha,'spanish') as Fecha2,CASE when a.Activo=1  then 'programada' else case when b.fecha is not null then 'ejercida' else 'cancelada' end  END as estatus,    case when a.fecha>getdate() then 1 else 0 end as cancela,a.fecha,b.FECHA as Fecha_incidencia,c.Incidencia  from SolicitudVacaciones a left join Incidencias b on b.INCIDENCIA=4 and a.id=b.ID and a.FECHA=b.FECHA  left join cat_incidencias c on a.INCIDENCIA=c.ID where a.id="& rsUpdateEntry("ID") &" AND A.PERIODO>=year(getdate())-1 order by a.fecha desc"
 'response.write strSQL  
 rsUpdateEntry2.Open strSQL, adoCon

if rsUpdateEntry2.EOF then
    response.write("<tr>")
    response.write("<td colspan=4 class='td-c fonttiny'>no hay solicitudes recibidas</td>" )    
    response.write("</tr>")

else  
 while  not rsUpdateEntry2.EOF
    response.write("<tr>")
    response.write("<td class='td-c fonttiny'>" & rsUpdateEntry2("fecha2") &"</td>" )
    response.write("<td class='td-c fonttiny'>" & rsUpdateEntry2("periodo") &"</td>" )
    response.write("<td class='td-c fonttiny'>" & rsUpdateEntry2("LogDate") &"</td>" )
    response.write("<td class='td-c fonttiny'>" & rsUpdateEntry2("estatus") &"</td>" )
     response.write("<td class='td-c fonttiny'>" & rsUpdateEntry2("incidencia") &"</td>" )

    if rsUpdateEntry2("activo")=1 and rsUpdateEntry2("cancela")=1  then
        response.write("<td class='td-c fonttiny'><a href='ShowContent.asp?contenido=49&ID="&rsUpdateEntry("ID") &"&fecha="& rsUpdateEntry2("fecha") &"&periodo="& rsUpdateEntry2("periodo") & "'><img src='/img/b_drop.png' with=11px alt='' title='' height=11px border=0></a></td>" ) 
    else
        response.write("<td class='td-c fonttiny'>&nbsp;</td>" ) 
    end if
    response.write("</tr>")
    rsUpdateEntry2.MoveNext
 wend
end if
     rsUpdateEntry2.close
     response.write("</table></td><td class=td-c style='vertical-align: top;'><H1>Programadas del area</H1><BR>")

     if rsUpdateEntry("clave_puesto")="027" then
     	     strSQL="select d.Nombres+' '+d.paterno as colaborador,a.LogDate,a.periodo,a.activo,dbo.fn_GetMonthName(a.fecha,'spanish') as Fecha2,CASE when a.Activo=1  then 'programada' else case when b.fecha is not null then 'ejercida' else 'cancelada' end  END as estatus,    case when a.fecha>getdate() then 1 else 0 end as cancela,a.fecha,b.FECHA as Fecha_incidencia,c.Incidencia  from SolicitudVacaciones a left join Incidencias b on b.INCIDENCIA=4 and a.id=b.ID and a.FECHA=b.FECHA  left join cat_incidencias c on a.INCIDENCIA=c.ID left join Empleados d on a.ID=d.ID where a.id!="&  rsUpdateEntry("ID") &" and a.Activo=1  AND A.PERIODO>=year(getdate())-1 and d.Clave_depto='"& rsUpdateEntry("clave_depto") &"' order by a.ID,a.fecha desc"
     	           rsUpdateEntry5.Open strSQL, adoCon

           
	     CreateTable(600)	     
	     response.write("<tr><td class='td-c subtitulo'>colaborador</td>")
	     response.write("<td class='td-c subtitulo'>fecha</td>")
	     response.write("<td class='td-c subtitulo'>Periodo</td></tr>")

	     while not rsUpdateEntry5.EOF
		     rowin
		       response.write("<td class='td-c fonttiny'>"&rsUpdateEntry5("colaborador")&"</td>")
		       response.write("<td class='td-c fonttiny'>"&rsUpdateEntry5("fecha2")&"</td>")
		       response.write("<td class='td-c fonttiny'>"&rsUpdateEntry5("periodo")&"</td>")
		     rowout
		     rsUpdateEntry5.MoveNext
	     wend
	     rsUpdateEntry5.close
	     closeTable()
     else
        CreateTable(600)
        rowin
           response.write("<td>&nbsp;</td>")
        rowout
        closeTable()
      end if

     response.write("</td></tr><tr><td class=td-c style='vertical-align: top;'>")
      %>  

      <P>&nbsp;</P>
      <H1> Vacaciones ejercidas </H1>

     <table border=0 width="360px" cellpadding="5" cellpadding="2" align="center">
      <tr>
         <td class='fonttiny subtitulo td-c'>Fecha de vaciones </td>
         <td class='fonttiny subtitulo td-c'>Per&iacute;odo </td>        
         <td class='fonttiny subtitulo td-c'>Comentarios</td>        
     </tr>
     <%

      strSQL="select *,dbo.fn_GetMonthName(fecha,'spanish') as Fecha2 from Incidencias where id=" & rsUpdateEntry("ID") &" and INCIDENCIA=4  and PERIODO>=year(getdate())-1 order by FECHA desc"
      'response.write strSQL  
       rsUpdateEntry2.Open strSQL, adoCon

 while  not rsUpdateEntry2.EOF
    response.write("<tr>")
    response.write("<td class='td-c fonttiny'>" & rsUpdateEntry2("fecha2") &"</td>" )
    response.write("<td class='td-c fonttiny'>" & rsUpdateEntry2("periodo") &"</td>" )
    response.write("<td class='td-c fonttiny'>" & rsUpdateEntry2("notas") &"&nbsp;</td>" )
    response.write("</tr>")
    rsUpdateEntry2.MoveNext
 wend
     rsUpdateEntry2.close
     response.write("</table></table><P>&nbsp;</P><P>&nbsp;</P>") 

 end sub 





sub CancelarVacaciones  'contenido 49'
  
   strSQL="select nombres+' '+paterno as colaborador,*,dbo.fn_GetMonthName(FechaIngreso,'spanish') as Fecha2 from Empleados where ID=" & request("ID")
   'response.write strSQL  
   rsUpdateEntry.Open strSQL, adoCon

   strSQL="select top 1 * from SolicitudVacaciones where ID=" & request("ID") &" and activo=1 order by Periodo desc"
   rsUpdateEntry2.Open strSQL, adoCon
   vPeriodo=int(trim(rsUpdateEntry2("Periodo")))
   rsUpdateEntry2.close
 

   response.write("<P>&nbsp;</P><P>&nbsp;</P>")
   titulo="Cancelar solicitud de vacaciones para: "& rsUpdateEntry("colaborador") 
   doTitulo(titulo)

     if int(trim(request("Periodo")))>=vPeriodo then
   %>

     <P class=alert> Est&aacute; seguro de cancelar la solicitud de vacaciones para el d&iacute;a: <%=request("fecha")%>   </P>
     <P>&nbsp;</P>
     <P>   <%
       response.write("<a href='ShowContent.asp?contenido=46&action=1&ID="& request("ID") &"&fecha="&request("fecha") &"&periodo=" &request("periodo")  &"'><U>CANCELAR  </U> </a></P>")   %>
     <P>&nbsp;</P>
     <P> <a href="ShowContent.asp?contenido=46"><U>REGRESAR </U> </a></P> 
<%
     else
         
           CreateTable(500)
           RowIn
           response.write("<td class='alert td-c'>NO ES POSIBLE CANCELAR UN DIA DE VACACIONES DE UN PERIODO ANTERIOR A TRAVES DE INTRANET, DEBIDO QUE ES INFERIOR AL PERIODO VACACIONAL ACTUAL, POR FAVOR ENVIE SU SOLICITUD DE CANCELACION POR CORREO ELECTRONICO, PARA AJUSTAR LOS DIAS DE SU PERIODO VACACIONAL ACTUAL, O SI LE ES POSIBLE REALICE PRIMERO LA CANCELACION DE LOS SIGUIENTES DIAS DE SU PERIODO ACTUAL, PARA POSTERIORMENTE REALIZAR LA CANCELACION DE ESTE DIA (DESPUES DE LA CANCELACION PUEDE REGISTRAR LOS DIAS PROGRAMADOS A FUTURO)</td>")
           RowOut
           RowIn
           response.write("<td class='td-l'>")
           strSQL="select * from SolicitudVacaciones where ID=" & request("ID") &" and activo=1 and periodo="& vPeriodo &" order by FECHA desc"
           rsUpdateEntry2.Open strSQL, adoCon

           while not rsUpdateEntry2.EOF
               response.write(" * " & rsUpdateEntry2("fecha") &" <a href='ShowContent.asp?contenido=46&action=1&ID="& request("ID") &"&fecha="&rsUpdateEntry2("fecha") &"&periodo=" &rsUpdateEntry2("periodo")  &"'><U>[Procesar cancelacion de este d&iacutea]</a><br>"   )   
               rsUpdateEntry2.MoveNext
           wend
           rsUpdateEntry2.close

     end if
end sub



sub SolicVacacionesForm  'contenido 47'

   
   strSQL="select a.*,a.DiasDerecho-a.DiasTomados-a.DiasProgramados as disponibles,b.nombres+' '+b.paterno+' '+b.materno as colaborador from PeriodosVacacionales a inner join Empleados b on a.id=b.id where a.ID="& request("ID")  &" and a.DiasTomados<(a.DiasDerecho+a.DiasProgramados) order by a.Periodo"
   'response.write strSQL  
   rsUpdateEntry.Open strSQL, adoCon

   strSQL="select count(distinct(Periodo)) as Periodos from PeriodosVacacionales  where ID="& request("ID") &" and DiasTomados<(DiasDerecho+DiasProgramados) "
   rsUpdateEntry2.Open strSQL, adoCon

   titulo="Solicitar " & request("fdias") &" d&iacute;as de vacacion(s) para: "&rsUpdateEntry("colaborador") &" <br>(per&iacute;odos disponibles " & rsUpdateEntry2("periodos") &")"
   doTitulo(titulo)
    %>
    

     <table border=0 width="440px" cellpadding="5" cellpadding="3" align="center">
     <tr>
      <td class='subtitulo td-c fonttiny'># D&iacute;a </td>
      <td class='subtitulo td-c fonttiny'>Fecha </td>    
      <td class='subtitulo td-c fonttiny'>Per&iacute;odo </td>
      
     </tr>

    <%
          rsUpdateEntry2.close
          response.write("<form id='dias' method='POST' action='ShowContent.asp'>")
           response.write("<input type='hidden' name='contenido' value=48>")
           response.write("<input type='hidden' name='ID' value="& request("ID") &">")
           response.write("<input type='hidden' name='dias' value="& request("fdias") &">")
           response.write("<input type='hidden' name='periodo' value="& request("periodo") &">")

      programados=0

      for i=1 to request("fdias") 
          response.write("<tr>")
          programados=programados+1
          if request("Periodosdisponibles")>1 and  programados>rsUpdateEntry("disponibles") then                          
                rsUpdateEntry.MoveNext
                programados=1              
          end if
          response.write("<td class='td-r fonttiny'>"& i & "</td>")         
          response.write("<td class='td-l fonttiny'><input type='date' name='dia"&i&"'></td>")         
          response.write("<td class='td-l fonttiny'>"& rsUpdateEntry("periodo") &"</td>")
          response.write("<input type='hidden' name='periodo"&i&"' value="& rsUpdateEntry("periodo")  &">")      
          response.write("</tr>")
      next
      rsUpdateEntry.close

      response.write("<tr><td colspan=3 class='td-c'>&nbsp; </td></tr>")
      response.write("<tr><td colspan=3 class='td-l fonttiny' style='padding-left: 100px'><font color=red>* Revise las fechas introducidas, en caso de enviar una fecha no v&aacute;lida su entera solicitud ser&aacute; descartada...<br>** La fecha planeada debe cumplir con 21 d&iacute;as de anticipacion<br>***Su planeaci&oacute;n a futuro no puede exceder m&aacute;s all&aacute; de 4 meses  </font></td></tr>")

      response.write("<tr><td colspan=3 class='td-l fonttiny' style='padding-left: 100px'>&nbsp;</td></tr>")

      response.write("<tr><td class='td-r'> Observaciones: </td><td colspan=2 class='td-l'><input type='text' name='fnotas' class='anchox4' placeholder='comparta algun requerimiento especial'></td></tr>")

     
      response.write("<tr><td colspan=3 class='td-c'><input type='submit' value='solicitar'></form></td></tr>")

      response.write("</table><P>&nbsp;</P>")
      strSQL="select a.*, b.puesto from Empleados a left join cat_puesto b on a.Clave_Puesto=b.Clave_Puesto where a.ID="&request("ID")
      'response.write strSQL
      rsUpdateEntry.Open strSQL, adoCon
      vdepto=rsUpdateEntry("Clave_Puesto")

      if trim(rsUpdateEntry("puesto"))="GERENTE" then   '027'
      	  rsUpdateEntry.close
      	  strSQL="select c.Incidencia,.dbo.fn_GetMonthName(a.FECHA,'spanish') Fecha,a.LogDate,b.Nombres+' '+b.Paterno as Colaborador 		from SolicitudVacaciones a inner join Empleados b on a.ID=b.ID inner join cat_incidencias c on a.INCIDENCIA=c.ID 		where a.FECHA>getdate() and a.Activo=1 and a.ID<>"& request("ID")&" and ( b.Clave_Puesto='027' or b.Clave_depto='"& vdepto &"') 		UNION 		select c.Incidencia,.dbo.fn_GetMonthName(a.FECHA,'spanish') Fecha,a.LogDate,b.Nombres+' '+b.Paterno as Colaborador 		from Incidencias a inner join Empleados b on a.ID=b.ID inner join cat_incidencias c on a.INCIDENCIA=c.ID 		where a.FECHA>getdate() and a.ID<>"& request("ID")&" and ( b.Clave_Puesto='027' or b.Clave_depto='"& vdepto &"') "
      	   
      else
          rsUpdateEntry.close
          strSQL="select c.Incidencia,.dbo.fn_GetMonthName(a.FECHA,'spanish') Fecha,a.LogDate,b.Nombres+' '+b.Paterno as Colaborador 		from SolicitudVacaciones a inner join Empleados b on a.ID=b.ID inner join cat_incidencias c on a.INCIDENCIA=c.ID 		where a.FECHA>getdate() and a.Activo=1 and a.ID<>"& request("ID")&" and  b.Clave_depto='"& vdepto &"' 		UNION 		select c.Incidencia,.dbo.fn_GetMonthName(a.FECHA,'spanish') Fecha,a.LogDate,b.Nombres+' '+b.Paterno as Colaborador 		from Incidencias a inner join Empleados b on a.ID=b.ID inner join cat_incidencias c on a.INCIDENCIA=c.ID 		where a.FECHA>getdate() and a.ID<>"& request("ID")&" and  b.Clave_depto='"& vdepto &"' "
  
      end if
      'response.write strSQL
      rsUpdateEntry.Open strSQL, adoCon
      rsUpdateEntry2.Open strSQL, adoCon

      titulo="SOLICITUDES RECIBIDAS DE OTROS COLABORADORES"
      DoTitulo(titulo)

      CreateTable(700)
      CamposRS1
      ShowValoresRS2
      CloseTable
 
end sub




sub  RegistrarVacaciones  'contenido 48' 
  
     vString=""
     for i=1 to request("dias")
         vString="dia"&i
         if request(vString)="" then
             response.redirect("ShowContent.asp?contenido=46&msg=envio alguna fecha en blanco")
         end if
     next

    invalido=0
    flag=0

    for i=1 to request("dias")   'retrayendo las fechas introducidas'
        vString="dia"&i
        vString2="periodo"&i

        strSQL="select DATEDIFF(DAY,getdate(),'"& request(vString) &"') as calculo "
        rsUpdateEntry.Open strSQL, adoCon
        flag=int(trim(rsUpdateEntry("calculo")))
        rsUpdateEntry.close

        'response.write("la diferencia en dias de la fecha solicitada a la fecha de hoy es: " & flag )

        if flag>=21 and flag<124   then  'candado'

        	strSQL="select * from incidencias where incidencia=4 and ID="& request("ID")  &" and FECHA='"& request(vString) &"'"
        	rsUpdateEntry.Open strSQL, adoCon

        	if rsUpdateEntry.EOF then   'quiere decir no hay una incidencia de RecHum registrada antes'

        	            rsUpdateEntry.close

		            'datos de su periodo activo'
		            strSQL="select a.*,a.DiasDerecho-a.DiasTomados-a.DiasProgramados as disponibles,b.nombres+' '+b.paterno+' '+b.materno as colaborador,b.id_usuario from PeriodosVacacionales a inner join Empleados b on a.id=b.id where a.ID="& request("ID")  &" and a.Periodo="& request("periodo")
		            'response.write(strSQL &"<BR>")
		            rsUpdateEntry.Open strSQL, adoCon

		            strSQL="select max(identificador)+1 as calculo from SolicitudVacaciones"
		            rsUpdateEntry5.Open strSQL, adoCon

		            vIdentificador=trim(rsUpdateEntry5("calculo"))
		            rsUpdateEntry5.close


		            strSQL="set dateformat ymd;insert into SolicitudVacaciones (ID,FECHA,INCIDENCIA,id_usuario,LogDate,Periodo,Activo,Identificador) values ("& request("ID")  &",'"& request(vString) &"',4,'"& rsUpdateEntry("id_usuario") &"',getdate(),"& request("periodo") &",1,"&vIdentificador&")"
		            'response.write(strSQL &"<BR>")
		            rsUpdateEntry2.Open strSQL, adoCon
		     
		            strSQL="update PeriodosVacacionales set DiasProgramados=" & rsUpdateEntry("DiasProgramados")+1 &" where ID="& request("ID") &" and Periodo="&request(vString2)
		            'response.write(strSQL &"<BR>")
		            rsUpdateEntry2.Open strSQL, adoCon

		            rsUpdateEntry.close
		    else

                   rsUpdateEntry.close
	        end if

        else
            if invalido=0 then
                invalido=1
            end if
        end if

     next

     strSQL="exec SP_SolicitaVacaciones "& request("ID") &","& request("dias")
     'response.write(strSQL &"<BR>")
     rsUpdateEntry2.Open strSQL, adoCon

     if invalido=1 then
             response.redirect("ShowContent.asp?contenido=46&msg=una de las fechas no supera los 21 dias de anticipacion o excede 4 meses a futuro")
     else
             response.redirect("FormatoVacaciones.asp?ID=" & request("ID") &"&notas="&request("fnotas") )
             'response.redirect("ShowContent.asp?contenido=46&msg=las solicitudes de vacaciones fueron registradas")
     end if
    

end sub



sub EjercerVacaciones    'contenido 50  

  strSQL2="select case when cast('" & request("FECHA") &"' as Date)<=cast(getdate() as date) then 1 else case when DateDiff(DAY,cast(getdate() as date),cast('" & request("FECHA") &"' as Date))<=15 then 1 else 0 end end as flag"
  'response.write strSQL2
  rsUpdateEntry7.Open strSQL2, adoCon
  

  if trim(rsUpdateEntry7("flag"))="0" then
      rsUpdateEntry7.close
      response.redirect("ShowContent.asp?contenido=32&msg=convierta la solicitud en incidencia cuando concuerde con la quincena respectiva")
  else      
      rsUpdateEntry7.close
      strSQL2="select case when day(GETDATE())<=15 then convert(varchar,year(GETDATE()))+'-'+convert(varchar,month(GETDATE()))+'-'+'16' else convert(varchar,year(DATEADD(DAY,28,getdate())))+'-'+convert(varchar,month(DATEADD(DAY,28,getdate())))+'-'+'01' end as vfecha"
      rsUpdateEntry6.Open strSQL2, adoCon
      strSQL2="select case when cast('" & request("FECHA") &"' as Date)<=cast('"&rsUpdateEntry6("vfecha")&"' as date) then 1 else 0 end as flag"
      'response.write strSQL2
      rsUpdateEntry6.close
      rsUpdateEntry7.Open strSQL2, adoCon
      if trim(rsUpdateEntry7("flag"))="0" then
	      rsUpdateEntry7.close
	      response.redirect("ShowContent.asp?contenido=32&msg=convierta la solicitud en incidencia cuando concuerde con la quincena respectiva")
      else      
      	rsUpdateEntry7.close
      end if

  end if
  

  if request("action")="1" then 'ejercer vacaciones'
        strSQL="set dateformat ymd;select a.*,dbo.fn_GetMonthName(fecha,'spanish') as Fecha2,b.nombres+' '+b.paterno+' '+b.materno as colaborador,c.incidencia as descripcion  from SolicitudVacaciones a inner join Empleados b on a.ID=b.ID inner join cat_incidencias c on a.incidencia=c.ID where a.ID="& request("ID") & " and a.FECHA='" & request("FECHA") &"'"
   'response.write strSQL  
   rsUpdateEntry.Open strSQL, adoCon


    %>
     <P>&nbsp;</P>
     <H1> Ejercer una solucitud de <%=rsUpdateEntry("colaborador")%> </H1>
     <P>&nbsp;</P>

     <table border=0 width="440px" cellpadding="5" cellpadding="3" align="center">
     <tr>     
      <td class='subtitulo td-c fontmed'>Fecha </td>    
      <td class='subtitulo td-c fontmed'>Incidencia</td>
      <td class='subtitulo td-c fontmed'>Per&iacute;odo </td>      
      <td class='subtitulo td-c fontmed'>Comentarios </td>
     </tr>

    <%          
          response.write("<tr>")
          response.write("<td class='td-c fontmed'>"& rsUpdateEntry("Fecha2") & "</td>")        
          response.write("<td class='td-c fontmed'>"& rsUpdateEntry("descripcion") & "</td>")         
          response.write("<td class='td-c fontmed'>"& rsUpdateEntry("Periodo") & "</td>")         
          response.write("<td class='td-c fontmed'>"& rsUpdateEntry("Notas") &"&nbsp;</td>")  
          response.write("</tr>")
      rsUpdateEntry.close

      response.write("<tr><td colspan=3 class='td-c'>&nbsp; </td></tr>")
      response.write("<tr><td colspan=3 class='td-c'><a href='ShowContent.asp?contenido=51&ID="&request("ID")&"&fecha="&request("fecha")&"'><U>CONVERTIR SOLICITUD EN INCIDENCIA</U></A></td></tr>")
      response.write("</table>")
  end if




    if request("action")="2" then 'cancelar vacaciones'
        
        response.write("<P>&nbsp;</P><H1>Desea cancelar una solicitud de vacaciones?</H1> <P>&nbsp;</P>")
        response.write("<CENTER><table width='300px' border=1 cellpadding=5 cellspacing=2 align=center>")

       strSQL="set dateformat ymd;select a.*,b.Nombres+' '+b.Paterno+' '+b.Materno as colaborador,dbo.fn_GetMonthName(a.fecha,'spanish') AS fecha2 from  SolicitudVacaciones a inner join Empleados b on a.ID=b.ID where b.id="& request("ID") &" and Fecha='"& request("Fecha") &"'"
       'response.write strSQL
       rsUpdateEntry.Open strSQL, adoCon

         response.write("<tr><td colspan=3 class='subtitulo td-c'>"& rsUpdateEntry("colaborador") &"</td></tr>")
         response.write("<tr><td class='td-c'>"& rsUpdateEntry("Fecha2") &"</td>")
         response.write("<tr><td class='td-c'><a href='ShowContent.asp?contenido=50&action=3&ID="& request("ID") &"&fecha="&Request("fecha") &"'> cancelar solicitud </a></td></tr></table>")
       
      rsUpdateEntry.close

    end if




     if request("action")="3" then 'cancelar vacaciones UP'

          response.write(request("ID") &"<BR>") 
          response.write(request("FECHA") &"<BR>") 
          vPeriodo=year(date())

          strSQL="set dateformat ymd;select * from  SolicitudVacaciones where Activo=1 and ID="& request("ID")& " and FECHA='"&request("FECHA") &"' and INCIDENCIA in (3,4) "
          response.write strSQL
          rsUpdateEntry.Open strSQL, adoCon

          vPeriodo=rsUpdateEntry("Periodo")


          strSQL="set dateformat ymd;UPDATE  SolicitudVacaciones set CancelDate=getdate(),Activo=2 where ID="& request("ID")& " and FECHA='"&request("FECHA") &"' and INCIDENCIA="& rsUpdateEntry("INCIDENCIA")
          response.write strSQL
          rsUpdateEntry2.Open strSQL, adoCon

          if rsUpdateEntry("INCIDENCIA")=4 then
              strSQL="set dateformat ymd;update PeriodosVacacionales set DiasProgramados=DiasProgramados-1  where ID="& request("ID")& " and PERIODO="&vPeriodo
              response.write strSQL
              rsUpdateEntry2.Open strSQL, adoCon
          end if

          response.redirect("ShowContent.asp?contenido=32&msg=la solicitud fue cancelada")


    end if


end sub




Sub EjercerVacacionesUP  'contenido 51'

   strSQL="set dateformat ymd;select * from SolicitudVacaciones where ID="& request("ID") & " and FECHA='" & request("FECHA") &"'"
   'response.write strSQL  
   rsUpdateEntry.Open strSQL, adoCon

   strSQL="set dateformat ymd;select * from INCIDENCIAS where ID="& request("ID") & " and FECHA='" & request("FECHA") &"' and incidencia=4"
   response.write("<P>&nbsp;</P><P>&nbsp;</P><P>&nbsp;</P>" & strSQL ) 
   rsUpdateEntry2.Open strSQL, adoCon

   if rsUpdateEntry2.EOF then
       rsUpdateEntry2.close
       strSQL="set dateformat ymd;insert into INCIDENCIAS (ID,FECHA,INCIDENCIA,id_usuario,LogDate,PERIODO,NOTAS) values ("& request("ID") &",'"& request("FECHA") &"',"& rsUpdateEntry("incidencia") &",'"& Session("session_id")  &"',getdate()," & rsUpdateEntry("PERIODO") &",'" & rsUpdateEntry("Notas") &"')"
       rsUpdateEntry2.Open strSQL, adoCon

       strSQL="set dateformat ymd;update SolicitudVacaciones  set Activo=2,CancelDate=getdate()  where ID="& request("ID") &" and FECHA='"& request("FECHA") &"' and PERIODO=" & rsUpdateEntry("PERIODO")
       rsUpdateEntry2.Open strSQL, adoCon

      if rsUpdateEntry("incidencia")=4 then   'es una vacacion'
         strSQL="update PeriodosVacacionales set DiasProgramados=DiasProgramados-1,DiasTomados=DiasTomados+1 where ID="& request("ID") &" and PERIODO=" & rsUpdateEntry("PERIODO")
         rsUpdateEntry2.Open strSQL, adoCon
      end if

      rsUpdateEntry.close
      response.redirect("ShowContent.asp?contenido=32&msg=la solicitud se ha convertido en incidencia")
   else
      rsUpdateEntry.close
      rsUpdateEntry2.close
      response.redirect("ShowContent.asp?contenido=32&msg=ya hay una vacacion en esa fecha, proceda a eliminar la solicitud")


   end if


end sub  



sub FormEnviarCartera 'contenido 52'

    response.write("<P>&nbsp;</P><CENTER>")
    Doalert

    TITULO="ENVIAR INFORME: CARTERA DE CLIENTES"
    DoTitulo(TITULO)

    dim vfecha1 : vfecha1=date()
    dim vfecha2 : vfecha2=date()

      strSQL="select CASE  when DATEPART(dw,getdate())=2 then  CONVERT(VARCHAR(10), getdate(), 105)   when DATEPART(dw,getdate())=3 then  CONVERT(VARCHAR(10), dateadd(day,-1,getdate() ), 105) when DATEPART(dw,getdate())=4 then  CONVERT(VARCHAR(10), dateadd(day,-2,getdate() ), 105) when DATEPART(dw,getdate())=5 then  CONVERT(VARCHAR(10), dateadd(day,-3,getdate() ), 105) when DATEPART(dw,getdate())=6 then  CONVERT(VARCHAR(10), dateadd(day,-4,getdate() ), 105) when DATEPART(dw,getdate())=7 then  CONVERT(VARCHAR(10), dateadd(day,-5,getdate() ), 105) when DATEPART(dw,getdate())=1 then  CONVERT(VARCHAR(10), dateadd(day,-6,getdate() ), 105)  END as FechaCorte"
    rsUpdateEntry.Open strSQL, adoCon4
    'response.write(StrSQL&"<BR>")
    'response.write("Fecha de corte en SQL=" & rsUpdateEntry("FechaCorte") )
   
    
    vfecha1=cdate( day(rsUpdateEntry("FechaCorte"))  & "/"&    month(rsUpdateEntry("FechaCorte")) & "/"&  year(rsUpdateEntry("FechaCorte")) )
    vfecha2=vfecha1+7
    rsUpdateEntry.close

    'response.write( vfecha1 &"-" &vfecha2 )

    %>

     <form id="cobranza" method="POST" action="ShowContent.asp"> 
     <input type="hidden" name="contenido" value="53">

     Fecha de Corte: <select name="ffecha"><option value="<%=vfecha1%>"><%=vfecha1%></option><option value="<%=vfecha2%>"><%=vfecha2%> </option></select><br>
   
     Destinadarios:  <br><input style="width: 1200px" type="text" name="femail" value="mescobar@deltaflow.com.mx;ajauregui@deltaflow.com.mx;yorozco@deltaflow.com.mx;miguel.cabanillas@deltaflow.com.mx;aflores@deltaflow.com.mx;" ><br><br>
     <input type="submit" value="enviar informe"></form>

     <%

end sub




sub ExecuteCartera 'contenido 53'

  response.write( request("ffecha") &"<BR>")
  response.write( request("femail") &"<BR>")
  strSQL="exec SP_cobranza "& day(request("ffecha")) &"," &month(request("ffecha")) &"," & year(request("ffecha")) &",'"& request("femail")  &"'"
  response.write(  strSQL &"<BR>")
  rsUpdateEntry.Open strSQL, adoCon4
  response.redirect("ShowContent.asp?contenido=52&msg=se ha enviado la cartera via email helpdesk, revise su bandeja")

end sub



sub ShowAlarmas  'contenido 54'

    response.write("<P>&nbsp;</P><H1>CONFIGURACION PARA SISTEMAS DE ALARMAS</H1><p>&nbsp;</P>")
    response.write("<P class='td-c fontmed'>Clave Maestra: <font color=red>2837 </font> (Almacen y Arrayan) Palabra secreta:  <font color=red>flowmax </font><BR>Alarmex (telefono 686 104 2000)  4509 (Almacen)</P>")
    response.write("<P class='td-c fontmed'>Como crear o borrar usuarios. Presione *5 [maestra]<BR></P>")

    response.write("<P>&nbsp;</P><P class='td-c fontmed'>Ejemplo: *5 [maestra]  03  (borra el usuario en el canal 03) </P><P>&nbsp;</P>")

    response.write("<P class='td-c fontmed'>ALMACEN<BR>04-ajauregui<BR>01- aflores<BR>02- cesar<BR>04- mescobar<BR>05- mrangel <BR>09- akgonzalez</P><P>&nbsp;</P>")

    response.write("<P class='td-c fontmed'><a href='https://alarmex.com.mx' target='_new'><U>SITIO WEB ALARMEX</U></a> usuario: delta123  pass: delta0107 </P>")

    response.write("<P class='td-c fontmed'>ADT telefono 800 2025 238,  <a href='https://totalconnect2.com/' target='_new2'><U>SITIO WEB ADT</U></a>   usuario: 632292  pass: Mexico2020*<BR> Arrayan: contrato no. 704208719  SJR:  705116994  /// recomensas ADT usuario: sistemas@deltaflow.com.mx pass: Delta.mxl.</P>")

    CreateTable(600)
    rowin
    response.write("<th class='subtitulo td-c' colspan=2><a href='https://printing.winstondata.com/adt_clientes/login.aspx'>descargar facturas</a> </td>")
    RowOut
    rowin
    response.write("<td class='td-c'>Arrayan # 704208719 </td><td class='td-c'>pass: 82373020 </td>")
    RowOut
    rowin
    response.write("<td class='td-c'>SJR # 705116994 </td><td class='td-c'>pass: 21868550 </td>")
    RowOut
    rowin
    response.write("<td class='td-c'>Mty # 705276467 </td><td class='td-c'>pass: 62628842 </td>")
    RowOut
    rowin
    response.write("<td class='td-c'>New Mty # 705391969 </td><td class='td-c'>pass: 23537201 </td>")
    RowOut
    closeTable

    response.write("<P class='td-c fontmed'><a href='https://www.mitotalplay.com.mx/' target='total'>TOTAL PLAY MTY  </a><br>cuenta: 0107659961 pass: $S3id0r$</P>")
    

    response.write("<P class='td-c fontmed'>ADD: Poner Maestra, elejir ranura ej 811 </P>") 
    response.write("<P class='td-c fontmed'>DEL: Poner Maestra, elejir ranura ej 811 + #0 </P>") 

end sub




sub FormPointEquilibrio   'contenido 55'

    DoAlert
    titulo="ENVIAR PUNTO DE EQUILIBRIO"
    DoTitulo(titulo)

    %>

     <form id="cobranza" method="POST" action="ShowContent.asp"> 
     <input type="hidden" name="contenido" value="56">

     Mes: <select name="fmes">
    <% for i=1 to 12
          if i=month(date())-1 then
             response.write("<option value="& i &" selected>"&meses(i) &"</option>")
          else
             response.write("<option value="& i &">"&meses(i) &"</option>")
          end if
       next 
        response.write("</select><br><br>A&ntilde;o: <select name='fanio'>")

      if month(date())=1 then
           response.write("<option value="& year(date())-1 &">"& year(date())-1 &"</option>")
           response.write("<option value="& year(date()) &">"& year(date()) &"</option>")
      else
           response.write("<option value="& year(date()) &">"& year(date()) &"</option>")
      end if  
       response.write("</select><br><br>")   %>

     Informar a detalle:  <select name="fcontrol"><option value=0 selected>no</option><option value=1>si</option></select> -->Elija a detalle si va a revisar calculos<br><br>

      
     Destinadarios:  <br><input style="width: 900px" type="text" name="femail" value="mescobar@deltaflow.com.mx;ajauregui@deltaflow.com.mx;miguel.cabanillas@deltaflow.com.mx;hescobar@deltaflow.com.mx;" ><br><br>
     <input type="submit" value="enviar informe"></form>

     <%


end sub




sub ExecutePE  'contenido 56'

 if  request("fcontrol")=0 then    'pidio el informe sin detalles'
     strSQL="SELECT * from PuntoEquilibrio where ANIO= "&  request("fanio") & " and MES=" & request("fmes") 
     rsUpdateEntry.Open strSQL, adoCon4

     if rsUpdateEntry.EOF then
          strSQL="INSERT into PuntoEquilibrio (ANIO,MES,ID_USUARIO) values ("& request("fanio") &","& request("fmes")  &",'"& Session("session_id") &"')"  
     else
          strSQL="UPDATE PuntoEquilibrio SET ID_USUARIO='"& Session("session_id") &"'"      
     end if
     rsUpdateEntry2.Open strSQL, adoCon4
     rsUpdateEntry.close
 end if


  strSQL="exec [SP_PuntoEquilibrio] '"& request("femail") &"'," & request("fmes") &"," & request("fanio")  &"," & request("fcontrol") 
  response.write(  strSQL &"<BR>")
  rsUpdateEntry.Open strSQL, adoCon4
  response.redirect("ShowContent.asp?contenido=55&msg=se ha enviado Punto de Equilibrio via email helpdesk, revise su bandeja")

end sub



sub ExecuteTipoCambio  'contenido 57'

     strSQL="exec SP_TC"  
     rsUpdateEntry.Open strSQL, adoCon4

     response.redirect ("/home.asp?msg=se envio el Tipo de cambio del dia via Helpdesk")

end sub







sub SolicitarCumple 'contenido 58'

    if request("action")="" then
        response.write("<P>&nbsp;</p>")
        if request("msg")<>"" then
           response.write("<P class='alert td-c'>"&  request("msg") &"</P>")
         end if
        response.write("<Form id='FormDiaCumple' method='POST' action='ShowContent.asp'><table align='center' border=0 width='150px' cellspacing=2 cellpadding=3><tr><td class='subtitulo td-c'>Elija la empresa:</td></tr><tr><td class='td-c'>")
        response.write("<input type='hidden' name='contenido' value=58>")
        response.write("<input type='hidden' name='action' value=1>")
        response.write("<select name='fempresa' class='anchox3'><option value=1>Servicios Profesionales</option><option value=3>Deltaflow de Mexico</option></select></td></tr><tr><td>&nbsp;</td></tr><tr><td class='td-c'><input type='submit' value='continuar'></td></tr>")
        response.write("</table><P>&nbsp;</p></form>")
    end if


    if request("action")="1" then        
        vEmpresa=""
        vLogo=""
        select case request("fempresa")
        case 1 
            vEmpresa="Servicios Profesionales, legales, financieros <BR> Recursos humanos y Gestoria, SR de CV"
            vLogo="logo-servPro.png"
            response.write("<input type='hidden' name='fempresa' value=1>")
        case 2
            vEmpresa="Corporativo Empresarial <BR> Izcalli SA de CV"
            vLogo="logo-Izcalli.png"
            response.write("<input type='hidden' name='fempresa' value=2>")
        case 3
            vEmpresa="Deltaflow de Mexico S de RL de CV"
            response.write("<input type='hidden' name='fempresa' value=3>")
            vLogo="logo-DMX.png"
        end select

        response.write("<P>&nbsp;</p><P>&nbsp;</p><Form id='FormDiaCumple' method='POST' action='ShowContent.asp'><table width='900px' border=1 cellpadding=2 align=center class=whiteBorders>")

        response.write("<input type='hidden' name='contenido' value=58>")
        response.write("<input type='hidden' name='action' value=2>")

        response.write("<input type='hidden' name='empresa' value='"& vEmpresa &"'>")
        response.write("<input type='hidden' name='logo' value='"& vLogo &"'>")

        response.write("<tr><td class='td-c whiteBorders' colspan=3 rowspan=2><img src='/images/"& vLogo &"' border=0></td><td class='td-c' colspan=3><B>FECHA DE LA SOLICITUD</B> <BR>  "& date() &" </td></tr>")
        response.write("<tr><td class='td-c whiteBorders' style='height: 4px; font-size: 4px'>&nbsp;</td><td class='td-c whiteBorders' style='height: 4px; font-size: 4px'>&nbsp;</td><td class='td-c whiteBorders' style='height: 4px; font-size: 4px'>&nbsp;</td></tr>")
        response.write("<tr><td class='td-c BoldBorders fontmed' colspan=3 style='height: 40px'><B>PERMISO CON GOCE DE SUELDO</td><td colspan=3 class='td-c whiteBorders'>&nbsp;</td></tr>")
        response.write("<tr><td class='td-c whiteBorders' colspan=6>&nbsp;</td></tr>")
        response.write("<tr><td class='td-l leftrighttopBorders' colspan=6 style='padding-left: 100px'>NOMBRE DEL SOLICITANTE:</td></tr>")
        response.write("<tr><td class='td-l leftrightBorders' colspan=6 style='padding-left: 100px'><select class=anchox4  name='fID'> ")

           strSQL="select * from Empleados where FechaEgreso is null and id_usuario='"& session("session_id") &"' "
           'response.write strSQL  
           rsUpdateEntry.Open strSQL, adoCon

           while not  rsUpdateEntry.EOF
              response.write("<option value="& rsUpdateEntry("ID") &">"& rsUpdateEntry("nombres") &" " & rsUpdateEntry("PATERNO") &" " & rsUpdateEntry("MATERNO") &"</option>")
              rsUpdateEntry.MoveNext
           wend
            rsUpdateEntry.close
            response.write("</select></td></tr>")

        response.write("<tr><td class='td-l fontmed leftrightBorders' colspan=6 style='padding-left: 100px'>&nbsp;</td></tr>")
        response.write("<tr><td class='td-l fontmed leftrightBorders' colspan=6 style='padding-left: 100px'>Por medio de la presente solicito permiso con goce de sueldo:</td></tr>")
        response.write("<tr><td class='td-l fontmed leftrightBorders' colspan=6 style='padding-left: 100px'>Especificar motivo:</td></tr>")
        response.write("<tr><td class='td-l fontmed leftrightBorders' colspan=6 style='padding-left: 100px'><select class=anchox2  name='fincidencia'><option value=3>D&iacute;a de cumplea&ntilde;os</select> </td></tr>")
        response.write("<tr><td class='td-l fontmed leftrightBorders' colspan=6 style='padding-left: 100px'>&nbsp;</td></tr>")
        response.write("<tr><td class='td-l fontmed leftrightBorders' colspan=6 style='padding-left: 100px'>Fecha de la ausencia:</td></tr>")
        response.write("<tr><td class='td-l fontmed leftrightBorders' colspan=6 style='padding-left: 100px'><input name='ffecha' id='ffecha' type='date'/> </td></tr>")
        response.write("<tr><td class='td-l fontmed leftrightBorders' colspan=6 style='padding-left: 100px'>&nbsp;</td></tr>")

        response.write("<tr><td class='td-l fontmed leftrightBorders' colspan=6 style='padding-left: 100px'>Comentarios:</td></tr>")
        response.write("<tr><td class='td-l fontmed leftrightBorders' colspan=6 style='padding-left: 100px'><input name='fnotas' id='fnotas' type='text'/ class='anchox4'></td></tr>")
        response.write("<tr><td class='td-l fontmed leftrightBorders' colspan=6 style='padding-left: 100px'>&nbsp;</td></tr>")
        response.write("<tr><td class='td-l fontmed leftrightbottomBorders ' colspan=6 style='padding-left: 100px'>&nbsp;</td></tr>")

        response.write("<tr><td class='td-l fontmed whiteBorders' colspan=6 style='padding-left: 100px'>&nbsp;</td></tr>")
        response.write("<tr><td class='td-l fontmed whiteBorders' colspan=6 style='padding-left: 100px'>&nbsp;</td></tr>")
        response.write("<tr><td class='td-l fontmed whiteBorders' colspan=6 style='padding-left: 100px'>&nbsp;</td></tr>")

        response.write("<tr><td class='td-l fontmed whiteBorders' width='200px'><HR></td><td class='td-l fontmed whiteBorders' width='100px'>&nbsp;</td><td class='td-l fontmed whiteBorders' width='200px'><HR></td><td class='td-l fontmed whiteBorders' width='100px'>&nbsp;</td><td class='td-l fontmed whiteBorders' colspan=2><HR></td></tr>")

         response.write("<tr><td class='td-c fontmed whiteBorders'>FIRMA SOLICITANTE </td><td class='td-l fontmed whiteBorders'>&nbsp;</td><td class='td-c fontmed whiteBorders'>FIRMA JEFE INMEDIATO </td><td class='td-l fontmed whiteBorders'>&nbsp;</td><td class='td-c fontmed whiteBorders'>FIRMA CAPITAL HUMANO </td></tr>")

        response.write("<tr><td class='td-l fontmed whiteBorders' colspan=6 style='padding-left: 100px'>&nbsp;</td></tr>")
         response.write("<tr><td class='td-l fontmed whiteBorders' colspan=6 style='padding-left: 100px'>&nbsp;</td></tr>")

        'response.write("<tr><td class='td-l fontmed whiteBorders' colspan=6><hr></td></tr>")
        response.write("<tr><td class='td-c fontmed whiteBorders' colspan=6><input type='submit' value='solicitar'></td></tr><tr><td class='td-c fontmed whiteBorders' colspan=6>&nbsp;</td></tr></table> </form> <P>&nbsp;</p><P>&nbsp;</p>")


    end if


   if request("action")="2" then  

      strSQL="set dateformat ymd;select * from INCIDENCIAS where ID="& request("fID") & " and FECHA='" & request("ffecha") &"'"
      'response.write("<P>&nbsp;</P><P>&nbsp;</P><P>&nbsp;</P>" & strSQL ) 
      rsUpdateEntry.Open strSQL, adoCon

      if not rsUpdateEntry.EOF then  'si no es fin de archivo entonces hay una incidencia en esta fecha'
            rsUpdateEntry.close
            response.redirect("ShowContent.asp?contenido=58&msg=ya existe una incidencia en esta fecha, comuniquese a capital humano para conocer mas informacion")
      end if
      rsUpdateEntry.close

      strSQL="set dateformat ymd;select * from SolicitudVacaciones where ID="& request("fID") & " and FECHA='" & request("ffecha") &"' and activo=1"
      rsUpdateEntry.Open strSQL, adoCon

      if not rsUpdateEntry.EOF then  'si no es fin de archivo entonces hay una incidencia en esta fecha'
            rsUpdateEntry.close
            response.redirect("ShowContent.asp?contenido=58&msg=ya tiene una solicitud de vacaciones / cumpleanos  en esta fecha")
      end if
      rsUpdateEntry.close

     strSQL="set dateformat ymd;select * from H_DiaInhabil where  FECHA='" & year(vfecha) &"-"& month(vfecha) &"-"& day(vfecha) &"'"
     'response.write strSQL
     rsUpdateEntry.Open strSQL, adoCon
                 
     if not rsUpdateEntry.EOF then
            rsUpdateEntry.close
            response.redirect("ShowContent.asp?contenido=58&msg=la fecha solicitada es un dia festivo, elija otro dia")
      end if
      rsUpdateEntry.close        

        strSQL="set dateformat dmy;select DATEPART(dw,'"& request("ffecha") &"') as numero"
        'response.write (strSQL &"<BR>")
         rsUpdateEntry.Open strSQL, adoCon

         vNum=rsUpdateEntry("numero")
         rsUpdateEntry.close

         if Vnum=7 or VNum=1 then  'es sabado o domindo
                response.redirect("ShowContent.asp?contenido=58&msg=la fecha solicitada es un sabado o domingo, elija otro dia")   
         end if


        strSQL="exec SP_SolicitaCumple "& request("fID") 
        'response.write("<BR><BR><BR>" & strSQL &"<BR>")
        rsUpdateEntry2.Open strSQL, adoCon

       strSQL="set dateformat ymd;insert into SolicitudVacaciones (ID,FECHA,INCIDENCIA,LogDate,Periodo,Activo,NOTAS) values ("& request("fID")  &",'"& request("ffecha") &"',3,getdate(),year(getdate()),1,'"&request("fnotas") &"')"
       'response.write("<BR><BR><BR>" & strSQL &"<BR>")
       rsUpdateEntry2.Open strSQL, adoCon

       ' response.redirect("ShowContent.asp?contenido=58&msg=la solicitada de cumpleanos quedo registrada") 
       'FORM SOLICITAR CUMPLE DESPUES DEL SUBMIT'


        strSQL="set dateformat ymd;select *,dbo.fn_GetMonthName(fecha,'spanish') as Fecha2 from SolicitudVacaciones where ID="& request("fID") & " and INCIDENCIA=3 and FECHA='"& request("ffecha") &"'"
        rsUpdateEntry2.Open strSQL, adoCon

         response.write("<P>&nbsp;</p>")
         response.write("<P class='alert td-c'>su solicitud ha sido recibida</P>")
        

        response.write("<P>&nbsp;</p><P>&nbsp;</p><table width='900px' border=1 cellpadding=2 align=center class=whiteBorders>")
        response.write("<tr><td class='td-c whiteBorders' colspan=3 rowspan=2><img src='/images/"& request("logo") &"' border=0></td><td class='td-c' colspan=3><B>FECHA DE LA SOLICITUD</B> <BR>  "& date() &" </td></tr>")
        response.write("<tr><td class='td-c whiteBorders' style='height: 4px; font-size: 4px'>&nbsp;</td><td class='td-c whiteBorders' style='height: 4px; font-size: 4px'>&nbsp;</td><td class='td-c whiteBorders' style='height: 4px; font-size: 4px'>&nbsp;</td></tr>")
        response.write("<tr><td class='td-c BoldBorders fontmed' colspan=3 style='height: 40px'><B>PERMISO CON GOCE DE SUELDO</td><td colspan=3 class='td-c whiteBorders'>&nbsp;</td></tr>")
        response.write("<tr><td class='td-c whiteBorders' colspan=6>&nbsp;</td></tr>")
        response.write("<tr><td class='td-l leftrighttopBorders' colspan=6 style='padding-left: 100px'>NOMBRE DEL SOLICITANTE:</td></tr>")
        strSQL="select * from Empleados where ID=" & request("fID")
        'response.write strSQL  
        rsUpdateEntry.Open strSQL, adoCon
        response.write("<tr><td class='td-l leftrightBorders' colspan=6 style='padding-left: 100px'>" &  rsUpdateEntry("nombres") &" " & rsUpdateEntry("PATERNO") &" " & rsUpdateEntry("MATERNO") &"</td></tr>")
        rsUpdateEntry.close
        response.write("<tr><td class='td-l fontmed leftrightBorders' colspan=6 style='padding-left: 100px'>&nbsp;</td></tr>")
        response.write("<tr><td class='td-l fontmed leftrightBorders' colspan=6 style='padding-left: 100px'>Por medio de la presente solicito permiso con goce de sueldo:</td></tr>")
        response.write("<tr><td class='td-l fontmed leftrightBorders' colspan=6 style='padding-left: 100px'>Especificar motivo:</td></tr>")
        response.write("<tr><td class='td-l fontmed leftrightBorders' colspan=6 style='padding-left: 100px'>DIA DE CUMPLEA&Ntilde;OS</td></tr>")
        response.write("<tr><td class='td-l fontmed leftrightBorders' colspan=6 style='padding-left: 100px'>&nbsp;</td></tr>")
        response.write("<tr><td class='td-l fontmed leftrightBorders' colspan=6 style='padding-left: 100px'>Fecha de la ausencia:</td></tr>")
        response.write("<tr><td class='td-l fontmed leftrightBorders' colspan=6 style='padding-left: 100px'>"& rsUpdateEntry2("Fecha2") &"</td></tr>")
        response.write("<tr><td class='td-l fontmed leftrightBorders' colspan=6 style='padding-left: 100px'>&nbsp;</td></tr>")

        response.write("<tr><td class='td-l fontmed leftrightBorders' colspan=6 style='padding-left: 100px'>Comentarios:</td></tr>")
        response.write("<tr><td class='td-l fontmed leftrightBorders' colspan=6 style='padding-left: 100px'>"& rsUpdateEntry2("NOTAS") &"</td></tr>")
        response.write("<tr><td class='td-l fontmed leftrightBorders' colspan=6 style='padding-left: 100px'>&nbsp;</td></tr>")
        response.write("<tr><td class='td-l fontmed leftrightbottomBorders ' colspan=6 style='padding-left: 100px'>&nbsp;</td></tr>")

        response.write("<tr><td class='td-l fontmed whiteBorders' colspan=6 style='padding-left: 100px'>&nbsp;</td></tr>")
        response.write("<tr><td class='td-l fontmed whiteBorders' colspan=6 style='padding-left: 100px'>&nbsp;</td></tr>")
        response.write("<tr><td class='td-l fontmed whiteBorders' colspan=6 style='padding-left: 100px'>&nbsp;</td></tr>")

        response.write("<tr><td class='td-l fontmed whiteBorders' width='200px'><HR></td><td class='td-l fontmed whiteBorders' width='100px'>&nbsp;</td><td class='td-l fontmed whiteBorders' width='200px'><HR></td><td class='td-l fontmed whiteBorders' width='100px'>&nbsp;</td><td class='td-l fontmed whiteBorders' colspan=2><HR></td></tr>")

         response.write("<tr><td class='td-c fontmed whiteBorders'>FIRMA SOLICITANTE </td><td class='td-l fontmed whiteBorders'>&nbsp;</td><td class='td-c fontmed whiteBorders'>FIRMA JEFE INMEDIATO </td><td class='td-l fontmed whiteBorders'>&nbsp;</td><td class='td-c fontmed whiteBorders'>FIRMA CAPITAL HUMANO </td></tr>")

        response.write("<tr><td class='td-l fontmed whiteBorders' colspan=6 style='padding-left: 100px'>&nbsp;</td></tr>")
         response.write("<tr><td class='td-l fontmed whiteBorders' colspan=6 style='padding-left: 100px'>&nbsp;</td></tr>")

        'response.write("<tr><td class='td-l fontmed whiteBorders' colspan=6><hr></td></tr>")
        response.write("<tr><td class='td-c fontmed whiteBorders' colspan=6>&nbsp;</td></tr><tr><td class='td-c fontmed whiteBorders' colspan=6>(imprima esta pantalla para presentarla como solicitud firmada)</td></tr></table>  <P>&nbsp;</p><P>&nbsp;</p>")




   end if 


end sub




sub CostosStock  'contenido 59'
     TITULO="COSTO CONTABLE DEL ALMACEN"
     DoTitulo(TITULO)
     response.write("<table width=900px>")

    
             
              strSQL="select Rate,RateDate from ORTT where year(RateDate)=year(getdate()) and month(RateDate)=month(getdate()) and day(RateDate)=day(getdate())"
              'response.write strSQL
              rsUpdateEntry.Open strSQL, adoCon2 'DMX'

              vRate=""
              if request("frate") <>"" then
                      vRate=request("frate")
              else
                      vRate=trim(rsUpdateEntry("Rate"))
              end if

              response.write ("<form action='ShowContent.asp' method='post'>")
              response.write("<input type='hidden' name='contenido' value=59>") 
              response.write("<input type='hidden' name='control' value=2>") 
              response.write("<tr><td>El tipo de cambio para el d&iacute;a de hoy es: <input type='text' name='frate' value='"& vRate &"'></td></tr>")
              rsUpdateEntry.close 

     if request("control")=1 then

              response.write("<tr><td>&nbsp;</td></tr>")
              response.write("<tr><td><B>Pegue el listado de sus c&oacute;digos en esta &aacute;rea:</B></td></tr>")
              'response.write ("<input type='hidden' name='frate' id='frate' value='"&  trim(rsUpdateEntry("Rate")) &"'>")
                 %>
              
              <tr><td>
              <textarea rows="20" cols="25" id="fpartidas" name="fpartidas"><%=request("fpartidas")%></textarea></td></tr>

              <tr><td>
              <input type="submit" value="mostrar costos"></td></tr>
              </form></table>
              <%

    end if

    Dim partidasX(500)

    response.write("</table><P>&nbsp;</P>")

    if request("control")=2 then

      rate=""
      rate=request("frate")
      'response.write(rate &"<BR>")

      partidas=""
      partidas=trim(request("fpartidas"))
      partida=""
      Pos=1
      longitud=0
      strSQL=""

      x=0

      while Pos<>0
        pos=InStr(partidas,chr(13))
        
        if pos=0 then
              partida=trim(partidas) 
        else
              partida=trim(mid(Partidas,1,Pos-1))
              Partidas=mid(Partidas,Pos+2,2000)
        end if
         
        if len(trim(partida))>1 then
             x=x+1
             'strSQL=strSQL & "'"& trim(partida) &"',"   
             partidasX(x)= trim(partida)
        end if    
          'response.write( strSQL &"<BR>")
      wend
      
      flag=0

      for i=1 to x      

          if flag=0 then
              response.write("<H1><B>COSTO CONTABLE</B> </H1>")
              response.write("<table width='730px' cellspacing=2 cellpadding=2 border=0>")
              response.write("<tr><td class=subtitulo>Cogido</td><td class=subtitulo>Descripcion</td><td class=subtitulo>Costo MXN</td><td class=subtitulo>Almacen</td><td class=subtitulo>Razon Social</td></tr>")
              flag=1
          end if

          'REVISA SI HAY STOCK EN UNA RAZON SOCIAL'
          strSQL="SELECT isnull( cast( SUM(Y.InQty)-SUM(Y.OutQty) as int ),0) as calculo from  OINM Y  WITH (NOLOCK) where y.ItemCode='"&  partidasX(i) &"' "
          'response.write(strSQL&"<BR>")
          rsUpdateEntry.Open strSQL, adoCon3   'DFW

          vCalculo=0
          vCalculo=rsUpdateEntry("calculo")
          rsUpdateEntry.close
              
             

          if vCalculo>0 then
                strSQL2="select  '$ ' + CONVERT(VARCHAR,CONVERT(MONEY,a.AvgPrice),1)  as AvgPrice,a.WhsCode,b.ItemCode,b.ItemName from OITW a inner join OITM b on a.ItemCode=b.ItemCode where a.WhsCode='001-Mxl' and a.ItemCode='"&  partidasX(i) &"' "
                 'response.write(strSQL2&"<BR>")
                rsUpdateEntry.Open strSQL2, adoCon3   'DFW

                  strSQL3="select  '$ ' + CONVERT(VARCHAR,CONVERT(MONEY,a.AvgPrice),1)  as AvgPrice,a.WhsCode,b.ItemCode,b.ItemName from OITW a inner join OITM b on a.ItemCode=b.ItemCode where a.WhsCode='002-SJR' and a.ItemCode='"&  partidasX(i) &"' "
                '' response.write(strSQL3&"<BR>")
                rsUpdateEntry2.Open strSQL3, adoCon3   'DFW

                   strSQL4="select  '$ ' + CONVERT(VARCHAR,CONVERT(MONEY,a.AvgPrice),1)  as AvgPrice,a.WhsCode,b.ItemCode,b.ItemName from OITW a inner join OITM b on a.ItemCode=b.ItemCode where a.WhsCode='003-MTY' and a.ItemCode='"&  partidasX(i) &"' "
                ' response.write(strSQL3&"<BR>")
                rsUpdateEntry3.Open strSQL4, adoCon3   'DFW
                       
                response.write("<tr><td>"& rsUpdateEntry("ItemCode") &"</td><td>"& rsUpdateEntry("ItemName") &"</td><td class='td-r'>" & rsUpdateEntry("AvgPrice") &"</td><td class='td-c'>MXL</td><td class='td-c'>DFW</td></tr>")

                  response.write("<tr><td>"& rsUpdateEntry2("ItemCode") &"</td><td>"& rsUpdateEntry2("ItemName") &"</td><td class='td-r'>" & rsUpdateEntry2("AvgPrice") &"</td><td class='td-c'>SJR</td><td class='td-c'>DFW</td></tr>")

                   response.write("<tr><td>"& rsUpdateEntry3("ItemCode") &"</td><td>"& rsUpdateEntry3("ItemName") &"</td><td class='td-r'>" & rsUpdateEntry3("AvgPrice") &"</td><td class='td-c'>MTY</td><td class='td-c'>DFW</td></tr>")


                rsUpdateEntry.close
                rsUpdateEntry2.close
                rsUpdateEntry3.close
          else
                response.write("<tr><td>"&  partidasX(i) &"</td><td colspan=3 class='td-c'>sin stock</td><td class='td-c'>DFW</td></tr>")


          end if

          rsUpdateEntry.Open strSQL, adoCon2   'DMX
          vCalculo=rsUpdateEntry("calculo")
          rsUpdateEntry.close

          if vCalculo>0 then

                 strSQL2="select  '$ ' + CONVERT(VARCHAR,CONVERT(MONEY,a.AvgPrice),1)  as AvgPrice,a.WhsCode,b.ItemCode,b.ItemName from OITW a inner join OITM b on a.ItemCode=b.ItemCode where a.WhsCode='001-Mxl' and a.ItemCode='"&  partidasX(i) &"' "
                 'response.write(strSQL2&"<BR>")
                rsUpdateEntry.Open strSQL2, adoCon2   'DMX   

                  strSQL3="select  '$ ' + CONVERT(VARCHAR,CONVERT(MONEY,a.AvgPrice),1)  as AvgPrice,a.WhsCode,b.ItemCode,b.ItemName from OITW a inner join OITM b on a.ItemCode=b.ItemCode where a.WhsCode='002-SJR' and a.ItemCode='"&  partidasX(i) &"' "
                '' response.write(strSQL3&"<BR>")
                rsUpdateEntry2.Open strSQL3, adoCon2   'DMX   

                 strSQL4="select  '$ ' + CONVERT(VARCHAR,CONVERT(MONEY,a.AvgPrice),1)  as AvgPrice,a.WhsCode,b.ItemCode,b.ItemName from OITW a inner join OITM b on a.ItemCode=b.ItemCode where a.WhsCode='003-MTY' and a.ItemCode='"&  partidasX(i) &"' "
                '' response.write(strSQL3&"<BR>")
                rsUpdateEntry3.Open strSQL4, adoCon2   'DMX  
              
                       
                response.write("<tr><td>"& rsUpdateEntry("ItemCode") &"</td><td>"& rsUpdateEntry("ItemName") &"</td><td class='td-r'>" & rsUpdateEntry("AvgPrice") &"</td><td class='td-c'>MXL</td><td class='td-c'>DMX</td></tr>")

                  response.write("<tr><td>"& rsUpdateEntry2("ItemCode") &"</td><td>"& rsUpdateEntry2("ItemName") &"</td><td class='td-r'>" & rsUpdateEntry2("AvgPrice") &"</td><td class='td-c'>SJR</td><td class='td-c'>DMX</td></tr>")
                   response.write("<tr><td>"& rsUpdateEntry3("ItemCode") &"</td><td>"& rsUpdateEntry3("ItemName") &"</td><td class='td-r'>" & rsUpdateEntry3("AvgPrice") &"</td><td class='td-c'>MTY</td><td class='td-c'>DMX</td></tr>")

                rsUpdateEntry.close
                rsUpdateEntry2.close
                rsUpdateEntry3.close
          else
                 response.write("<tr><td>"&  partidasX(i) &"</td><td colspan=3 class='td-c'>sin stock</td><td class='td-c'>DMX</td></tr>")
          end if

          separador

      next
      response.write("</table><P>&nbsp;</P><P>&nbsp;</P><P>&nbsp;</P><P>&nbsp;</P>")
           

       

    end if

end sub






sub titulosTask
    response.write("<tr><td class='td-c subtitulo'>ID </td><td class='td-c subtitulo'>Responsable</td><td class='td-c subtitulo'>Relacionado1</td><td class='td-c subtitulo'>Relacionado2</td><td class='td-c subtitulo'>F. Solicitud</td><td class='td-c subtitulo'>F. Agenda</td><td class='td-c subtitulo'>Adjunto</td><td class='td-c subtitulo' width='90px'>&nbsp;</td> </tr>")  
end sub





sub ShowTasks


                if rsUpdateEntry.EOF then
                     NoRegistros
                else
                  i=1

                  while not rsUpdateEntry.EOF
                      CreateTable(1000)
                      vresponsable=rsUpdateEntry("responsable")
                      
                      titulosTask
                      response.write("<tr><td class='td-c fontmed' width='40px' rowspan=3><a href='ShowContent.asp?contenido=62&action=3&ID="& rsUpdateEntry("ID") &"'><B><U>" &  rsUpdateEntry("ID") &"</U></B></a></td>" )
                     

                      'response.write("<tr><td class='td-c fontmed' width='120px'>" &  mid(rsUpdateEntry("Beneficiario"),1,22) &"</td>" )
                      response.write("<td class='td-c fontmed' width='120px'>" &  mid(rsUpdateEntry("Ejecutor"),1,22) &"</td>" )                      
                      response.write("<td class='td-c fontmed' width='120px'>" &  mid(rsUpdateEntry("Relacion"),1,22) &"</td>" )
                      response.write("<td class='td-c fontmed' width='120px'>" &  mid(rsUpdateEntry("Relacion2"),1,22) &"</td>" )
                      response.write("<td class='td-c fontmed' width='80px'>" &  day(rsUpdateEntry("DocDate")) &"-" & meses( month(rsUpdateEntry("DocDate")) ) &"-" & year(rsUpdateEntry("DocDate")) &"</td>" )

                      if rsUpdateEntry("Urgencia")=1 then
                          response.write("<td class='td-c fontmed' width='80px' style='background-color: #FFC0C0'>" &  day(rsUpdateEntry("FechaProgramada")) &"-" & meses( month(rsUpdateEntry("FechaProgramada")) )  &"-" & year(rsUpdateEntry("FechaProgramada")) &"</td>" )
                      else
                         response.write("<td class='td-c fontmed' width='80px'>" &  day(rsUpdateEntry("FechaProgramada")) &"-" & meses( month(rsUpdateEntry("FechaProgramada")) ) &"-" & year(rsUpdateEntry("FechaProgramada")) &"</td>" )
                      end if
                      
                      'response.write("<td class='td-l fontmed' >" &  mid(rsUpdateEntry("Solicitud"),1,200) &"</td>" )

                      if  rsUpdateEntry("ArchivoRelacionado")="" then
                           response.write("<td class='td-c fontmed'>&nbsp;</td>" )
                      else
                           response.write("<td class='td-c fontmed'><a href='/TaskManager/"&  rsUpdateEntry("ArchivoRelacionado")&"' target='_xoxo'><U>" & rsUpdateEntry("ArchivoRelacionado") &"</u></a></td>")                                                      
                      end if          

                     if request("fpalabra")<>"" then
                         if rsUpdateEntry("estatus")=1 then
                          response.write("<td class='td-c fontmed' width='90px'>cerrada</td>" )
                         else
                          response.write("<td class='td-c fontmed alert' width='90px'>abierta</td>" )
                         end if
                     end if

                      response.write("</tr>")

                       response.write("<td class='td-l Fontbig' colspan=10><B>[ <font style='background-color:C6DCC6;'>Asunto</font>] " &  rsUpdateEntry("Asunto") &" </B></td></tr>" )

                       if rsUpdateEntry("Solicitud")<>"" then
                            response.write("<tr><td colspan=10 class='td-l Fontbig'><B>[ Descripci&oacute;n ] </B>"& rsUpdateEntry("Solicitud") &"</td></tr>")
                      else
                           response.write("<tr><td colspan=10 style='font-size: 3px'>&nbsp;</td></tr>")
                      end if

                      'if rsUpdateEntry("seguimiento")<>"" then
                            'response.write("<tr><td colspan=10 class='td-l fontbig'><B>[ Seguimiento ] </B>"& rsUpdateEntry("seguimiento") &"</td></tr>")
                      'else
                           'response.write("<tr><td colspan=10 style='font-size: 3px'>&nbsp;</td></tr>")
                      'end if
                      
                      rsUpdateEntry.MoveNext
                      i=i+1
                      closeTable

                  wend
                end if
                
                 response.write("</table><P>&nbsp;</P>")


end sub








sub TaskManager  'contenido 62'

   modulosTask
   
   select case request("action")
     case 0     response.write("<P style='font-size:4px'>&nbsp;</P>")  
                 if request("msg")<>"" then
                   response.write("<P class='alert td-c'>"&  request("msg") &"</P>")
                 end if 

                if session("session_id")="MCABANILLAS" then
                    TITULO="Tareas abiertas, responsable: "& session("session_id") 
                else
                    TITULO="Tareas abiertas, responsable: "& session("session_id") 
                end if
                DoTitulo(TITULO)


                 if request("responsable")<>"" then
                     if request("responsable")="none" then
                         session_filtro=""
                     else
                         session_filtro=request("responsable")
                     end if
                 end if


                 strSQL="select a.*,case when FechaProgramada<getdate() then 1 else 0 end as Urgencia,dbo.fn_GetMonthName(a.DocDate,'spanish') as FechaX,dbo.fn_GetMonthName(a.FechaProgramada,'spanish') as FechaY, case when FechaCierre is null then 0 else 1 end as Estatus,b.nombres+' '+b.paterno as Ejecutor,c.nombres+' '+c.paterno as Beneficiario,d.nombres+' '+d.paterno as Relacion,e.nombres+' '+e.paterno as relacion2 from TaskManager a left join Empleados b on b.FechaEgreso is null and  a.Responsable=b.id_usuario left join Empleados c on c.FechaEgreso is null and  a.Solicitante=c.id_usuario left join Empleados d on d.FechaEgreso is null and d.ID_USUARIO<>'' and  a.Relacionados=d.id_usuario left join Empleados e on e.FechaEgreso is null and e.ID_USUARIO<>'' and  a.relacionado2=e.id_usuario where a.FechaCierre is null "

                 if session_filtro<>"" then
                     strSQL= strSQL & "and a.Responsable='"& session_filtro &"'   order by a.FechaProgramada"
                 else
                     strSQL= strSQL & "and a.Responsable='"& session("session_id")  &"'   order by a.FechaProgramada"
                 end if

                'response.write(strSQL &"<BR>")
                rsUpdateEntry.Open strSQL, adoCon  'intranet'
                
                ShowTasks
                rsUpdateEntry.close

                if session_filtro="" then
                      separador

                      response.write("<H1>Tareas abiertas relacionadas a: "& session("session_id") &"</H1>")
                     
                     strSQL="select a.*,case when FechaProgramada<getdate() then 1 else 0 end as Urgencia,dbo.fn_GetMonthName(a.DocDate,'spanish') as FechaX,dbo.fn_GetMonthName(a.FechaProgramada,'spanish') as FechaY, b.nombres+' '+b.paterno as Ejecutor,c.nombres+' '+c.paterno as Beneficiario,d.nombres+' '+d.paterno as Relacion,e.nombres+' '+e.paterno as relacion2 from TaskManager a left join Empleados b on b.FechaEgreso is null and  a.Responsable=b.id_usuario left join Empleados c on c.FechaEgreso is null and  a.Solicitante=c.id_usuario left join Empleados d on d.FechaEgreso is null and d.ID_USUARIO<>'' and  a.Relacionados=d.id_usuario left join Empleados e on e.FechaEgreso is null and e.ID_USUARIO<>'' and  a.relacionado2=e.id_usuario where a.FechaCierre is null  and ( a.Relacionados='"& session("session_id")  &"' or  a.Relacionado2='"& session("session_id") &"' or  a.Relacionado3='"& session("session_id") &"' ) order by a.responsable"
                       'response.write(strSQL &"<BR>")

                    rsUpdateEntry.Open strSQL, adoCon  'intranet'
                    
                    ShowTasks
                    rsUpdateEntry.close
                end if
                    
    

     case 1    'ADD TASK' 
         response.write("<P>&nbsp;</P>")
         if request("msg")<>"" then
           response.write("<P class='alert td-c'>"&  request("msg") &"</P>")
         end if
         response.write("<H1>Nueva Tarea No. ? </H1><P>&nbsp;</P>")

         CreateTable(960)
        
         response.write ("<form action='ShowContent.asp' method='post'>")
         response.write("<input type='hidden' name='contenido' value=62>") 
         response.write("<input type='hidden' name='action' value=2>") 

         strSQL="select * from Empleados where FechaEgreso is null and ID_USUARIO='"& session("session_id") &"'"
         'response.write strSQL  
         rsUpdateEntry.Open strSQL, adoCon

         response.write("<tr><td class='td-l fontmed bold whiteBorders'>Solicitante: </td><td  class='td-l fontmed bold whiteBorders'>" &  rsUpdateEntry("nombres") &" " & rsUpdateEntry("PATERNO") &" " & rsUpdateEntry("MATERNO") &"</td></tr>")

         rsUpdateEntry.close


         response.write("<tr><td class='td-l fontmed bold whiteBorders'>Asunto de la Tarea: </td><td  class='td-l fontmed bold whiteBorders'><input type='text' name='fasunto' class='anchox3'></td></tr>")
         response.write("<tr><td class='td-l fontmed bold whiteBorders'>Fecha de la solicitud: </td><td  class='td-l fontmed whiteBorders'>"& date() &"</td></tr>")
         response.write("<tr><td class='td-l fontmed bold whiteBorders'>Fecha sugerida para la realizaci&oacute;n: </td><td  class='td-l fontmed whiteBorders'><input type='date' name='fprogramada'></td></tr>")
          response.write("<tr><td class='td-l fontmed bold whiteBorders'>Desea adjuntar una foto/pdf? </td><td  class='td-l fontmed whiteBorders'><input type='checkbox' name='ffile'></td></tr>")
         response.write ("<tr><td class='td-l fontmed bold' colspan=2>Descripci&oacute;n de la Tarea: <BR> <textarea name='fsolicitud' rows=16 cols=118></textarea></td></tr>")
                  
         response.write ("<tr><td class='td-l fontmed bold' colspan=2><font style='background-color:red; color: white'>APLICA A PROCESO DE PINTADO </font> <br>Orden del taller: <input type='text' name='fOC' class='anchoSmall' value='OC-'><br> Actualmente pint&aacute;ndose en taller: <input type='checkbox' name='fpintando'><P>&nbsp;</P></td></tr>")


         strSQL="select * from Empleados where FechaEgreso is null and ID_USUARIO is not null and ID_USUARIO<>'' "
         'response.write strSQL  
         rsUpdateEntry.Open strSQL, adoCon    

         response.write("<tr><td class='td-r fontmed bold whiteBorders'><font color='red'>Responsable: </font></td><td  class='td-l fontmed bold whiteBorders'><select name='fresponsable'>")

         while not  rsUpdateEntry.EOF
              if session("session_id") ="MCABANILLAS" and rsUpdateEntry("ID_USUARIO")="gvazquez" then
                  
                  response.write("<option value="& rsUpdateEntry("ID_USUARIO") &" SELECTED>"& rsUpdateEntry("nombres") &" " & rsUpdateEntry("PATERNO") &" " & rsUpdateEntry("MATERNO") &"</option>")   
                  
              else

                  response.write("<option value="& rsUpdateEntry("ID_USUARIO") &">"& rsUpdateEntry("nombres") &" " & rsUpdateEntry("PATERNO") &" " & rsUpdateEntry("MATERNO") &"</option>")   
                  
              end if
              rsUpdateEntry.MoveNext
         wend         
         response.write("</select></td></tr>")


          response.write("<tr><td class='td-r fontmed bold whiteBorders'>Colaborador relacionado 1: </td><td  class='td-l fontmed bold whiteBorders'><select name='frelacionado'>")
          response.write("<option value='' selected>ninguno</option>")

          rsUpdateEntry.MoveFirst
          while not  rsUpdateEntry.EOF
               'response.write("<option value="& rsUpdateEntry("ID_USUARIO") &">"& rsUpdateEntry("nombres") &" " & rsUpdateEntry("PATERNO") &" " & rsUpdateEntry("MATERNO") &"</option>")   
              if UCASE(trim(rsUpdateEntry("ID_USUARIO"))) = trim(session("session_id")) then
                  response.write("<option value="& rsUpdateEntry("ID_USUARIO") &" SELECTED>"& rsUpdateEntry("nombres") &" " & rsUpdateEntry("PATERNO") &" " & rsUpdateEntry("MATERNO") &"</option>")
              else
                  response.write("<option value="& rsUpdateEntry("ID_USUARIO") &">"& rsUpdateEntry("nombres") &" " & rsUpdateEntry("PATERNO") &" " & rsUpdateEntry("MATERNO") &"</option>")
              end if
              rsUpdateEntry.MoveNext
         wend
         rsUpdateEntry.MoveFirst
         response.write("</select></td></tr>")

       
         response.write("<tr><td class='td-r fontmed bold whiteBorders'>Colaborador relacionado 2: </td><td  class='td-l fontmed bold whiteBorders'><select name='frelacionado2'>")
          response.write("<option value='' selected>ninguno</option>")
          rsUpdateEntry.MoveFirst
          while not  rsUpdateEntry.EOF
               response.write("<option value="& rsUpdateEntry("ID_USUARIO") &">"& rsUpdateEntry("nombres") &" " & rsUpdateEntry("PATERNO") &" " & rsUpdateEntry("MATERNO") &"</option>")              
              rsUpdateEntry.MoveNext
         wend
         rsUpdateEntry.MoveFirst
         response.write("</select></td></tr>")

          response.write("<tr><td class='td-r fontmed bold whiteBorders'>Colaborador relacionado 3: </td><td  class='td-l fontmed bold whiteBorders'><select name='frelacionado3'>")
          response.write("<option value='' selected>ninguno</option>")
          rsUpdateEntry.MoveFirst
          while not  rsUpdateEntry.EOF
               response.write("<option value="& rsUpdateEntry("ID_USUARIO") &">"& rsUpdateEntry("nombres") &" " & rsUpdateEntry("PATERNO") &" " & rsUpdateEntry("MATERNO") &"</option>")              
              rsUpdateEntry.MoveNext
         wend
         rsUpdateEntry.MoveFirst
         response.write("</select></td></tr>")


         rsUpdateEntry.Close
         
         response.write ("<tr><td class='td-c fontmed bold' colspan=2><input type=submit value='agregar tarea'></td></tr></form></table><P>&nbsp;</P>")





     
     case 2  'UP ADD TASK' 
        vOCpintado=request("fOC")
        if vOCpintado="OC-" then
             vOCpintado=""
        end if

        if  request("fprogramada")="" then
             strSQL="set dateformat ymd;insert into TaskManager (ID,Solicitante,Responsable,Docdate,FechaProgramada,LogDate,Asunto,Solicitud,ArchivoRelacionado,Relacionados,relacionado2,relacionado3,OC_Pintado) values ((select max(ID) from TaskManager)+1,'"& UCASE(session("session_id")) &"','"& UCASE(request("fresponsable")) &"',getdate(),getdate(),getdate(),'"& request("fasunto") &"','"& request("fsolicitud") &"','','"& UCASE(request("frelacionado")) &"','"& UCASE(request("frelacionado3")) &"','"& UCASE(request("frelacionado2")) &"','"& vOCpintado&"')"
        else  
             strSQL="set dateformat ymd;insert into TaskManager (ID,Solicitante,Responsable,Docdate,FechaProgramada,LogDate,Asunto,Solicitud,ArchivoRelacionado,Relacionados,relacionado2,relacionado3,OC_Pintado) values ((select max(ID) from TaskManager)+1,'"& UCASE(session("session_id")) &"','"& UCASE(request("fresponsable")) &"',getdate(),'" & request("fprogramada") &"',getdate(),'"& request("fasunto") &"','"& request("fsolicitud") &"','','"& UCASE(request("frelacionado")) &"','"& UCASE(request("frelacionado2")) &"','"& UCASE(request("frelacionado3")) &"','"& vOCpintado&"')"
           
        end if
        'response.write strSQL
        rsUpdateEntry.Open strSQL, adoCon   

        strSQL="set dateformat ymd;select * from TaskManager where Solicitante='"& UCASE(session("session_id")) &"' and  responsable='"& UCASE(request("fresponsable")) &"' and DocDate='"& year(date()) &"-" &month(date()) &"-" &day(date())  &"' and asunto='"& request("fasunto")  &"'" 
        response.write strSQL
        rsUpdateEntry2.Open strSQL, adoCon

        if not rsUpdateEntry2.EOF then
              Session("ID")=rsUpdateEntry2("ID")
              if request("ffile")="on" or request("ffile")="true"  then   'va a subir archivo'
                    response.redirect("/upload/uploadform2.asp?ID="& rsUpdateEntry2("ID") )
              else 
                    response.redirect("ShowContent.asp?contenido=62&action=0&msg=se agrego nueva tarea no. " & rsUpdateEntry2("ID") )
              end if
            
        else
            response.redirect("ShowContent.asp?contenido=62&action=0&msg=se agrego una nueva tarea")
        end if

     


     case 3  'EDIT TASK' 
         strSQL="select replace(a.seguimiento,'<BR />','<P>&nbsp;</P>') as Seguimientox,a.*,b.Nombres+' '+b.Paterno+' '+b.Materno as Beneficiario,c.Nombres+' '+c.Paterno+' '+c.Materno as Ejecutor,dbo.fn_GetMonthName(a.DocDate,'spanish') as FechaX,dbo.fn_GetMonthName(a.FechaProgramada,'spanish') as FechaY,case when FechaCierre is null then 0 else 1 end as 'closed' from TaskManager a inner join Empleados b on a.Solicitante=b.id_usuario inner join Empleados c on a.Responsable=c.id_usuario where a.ID=" & request("ID") 
         'response.write(strSQL &"<BR>")
       rsUpdateEntry.Open strSQL, adoCon  'intranet'

         'response.write("<P style>&nbsp;</P>")
         if request("msg")<>"" then
           response.write("<P class='alert td-c'>"&  request("msg") &"</P>")
         end if
         response.write("<H3>Editar Tarea No. "& request("ID") &" </H3><P>&nbsp;</P>")

         CreateTable(800)


         response.write ("<form action='ShowContent.asp' method='post'>")
         response.write("<input type='hidden' name='contenido' value=62>") 
         response.write("<input type='hidden' name='action' value=4>") 
         response.write("<input type='hidden' name='ID' value="& request("ID") &">") 

         if rsUpdateEntry("ArchivoRelacionado")<>"" then
             response.write("<tr><td class='td-l fontmed bold whiteBorders'>Solicitante: </td><td  class='td-l fontmed bold whiteBorders'>" &  rsUpdateEntry("beneficiario") &"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href='/TaskManager/"& rsUpdateEntry("ArchivoRelacionado") &"' target='adjunto'>Archivo adjunto: <U>"& rsUpdateEntry("ArchivoRelacionado") &"</U></a></td></tr>")
         else
             response.write("<tr><td class='td-r fontmed subtitulo whiteBorders' width='180px'>Solicitante: </td><td  class='td-l fontmed bold whiteBorders'>" &  rsUpdateEntry("beneficiario") &"</td></tr>")
          end if

         response.write("<tr><td class='td-r fontmed subtitulo whiteBorders'>Asunto de la Tarea: </td><td  class='td-l fontmed bold whiteBorders'><input type='text' name='fasunto' class='anchox4' value='"& rsUpdateEntry("asunto") &"'></td></tr>")
         response.write("<tr><td class='td-r fontmed subtitulo whiteBorders'>Fecha de la solicitud: </td><td  class='td-l fontmed whiteBorders'>"& rsUpdateEntry("FechaX") &"</td></tr>")
         response.write("<tr><td class='td-r fontmed subtitulo whiteBorders'>Fecha de agenda: &nbsp;&nbsp;("& rsUpdateEntry("FechaY") &")</td><td  class='td-l fontmed whiteBorders'><input type='date' name='fprogramada'> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;recordar por e-mail en fecha de agenda ")

         if  rsUpdateEntry("EnviarRecordatorio")=true   then
               response.write(  "<input type='checkbox' name='recordar' checked> </td></tr>")
         else
               response.write(  "<input type='checkbox' name='recordar'> </td></tr>")
         end if
          
          response.write ("<tr><td class='td-r fontmed subtitulo whiteBorders'>Descripci&oacute;n de la Tarea:</td><td  class='td-l fontmed whiteBorders'>"& rsUpdateEntry("Solicitud") &"</td></tr>") 

        closeTable
        CreateTable(980)

         response.write ("<tr><td class='td-l Fontbig' colspan=2><B>Seguimiento: </B><BR>"& rsUpdateEntry("Seguimiento")  &" <textarea name='fseguimiento' rows=8 cols=120></textarea></td></tr>")
                  
         
        response.write ("<tr><td class='td-l fontmed bold' colspan=2><font style='background-color:red; color: white'>APLICA A PROCESO DE PINTADO </font> <br>Orden del taller: <input type='text' name='fOC' class='anchoSmall' value='"& rsUpdateEntry("OC_Pintado") &"'><br> Actualmente pint&aacute;ndose en taller: ")

         if rsUpdateEntry("Pintandose")=1 then  
            response.write("<input type='checkbox' name='fpintando' checked><P>&nbsp;</P></td></tr>")
         else
            response.write("<input type='checkbox' name='fpintando'><P>&nbsp;</P></td></tr>")
         end if


         strSQL="select * from Empleados where FechaEgreso is null and ID_USUARIO is not null and ID_USUARIO<>'' "
         'response.write strSQL  
         rsUpdateEntry2.Open strSQL, adoCon    

         response.write("<tr><td class='td-R fontmed bold whiteBorders'><font color='red'>RESPONSABLE: </font></td><td  class='td-l fontmed bold whiteBorders'><select name='fresponsable'>")
         
         rsUpdateEntry2.MoveFirst
         while not  rsUpdateEntry2.EOF
              if UCASE(trim(rsUpdateEntry("Responsable"))) = UCASE(rsUpdateEntry2("ID_USUARIO")) then
                  response.write("<option value="& rsUpdateEntry2("ID_USUARIO") &" SELECTED>"& rsUpdateEntry2("nombres") &" " & rsUpdateEntry2("PATERNO") &" " & rsUpdateEntry2("MATERNO") &"</option>")
              else
                  response.write("<option value="& rsUpdateEntry2("ID_USUARIO") &">"& rsUpdateEntry2("nombres") &" " & rsUpdateEntry2("PATERNO") &" " & rsUpdateEntry2("MATERNO") &"</option>")
              end if
              rsUpdateEntry2.MoveNext
         wend

       
         response.write("</select></td></tr>")

          response.write("<tr><td class='td-r fontmed bold whiteBorders'>Colaborador relacionado 1: </td><td  class='td-l fontmed bold whiteBorders'><select name='frelacionado'>")

          if rsUpdateEntry("relacionados")="" then
             response.write("<option value='' selected>ninguno</option>")
          else
             response.write("<option value=''>ninguno</option>")
          end if

          rsUpdateEntry2.MoveFirst
          while not  rsUpdateEntry2.EOF
              if UCASE(rsUpdateEntry("relacionados"))=UCASE(rsUpdateEntry2("ID_USUARIO")) then
                    response.write("<option value="& rsUpdateEntry2("ID_USUARIO") &" SELECTED>"& rsUpdateEntry2("nombres") &" " & rsUpdateEntry2("PATERNO") &" " & rsUpdateEntry2("MATERNO") &"</option>")     
              else
                   response.write("<option value="& rsUpdateEntry2("ID_USUARIO") &">"& rsUpdateEntry2("nombres") &" " & rsUpdateEntry2("PATERNO") &" " & rsUpdateEntry2("MATERNO") &"</option>")    
              end if         
              rsUpdateEntry2.MoveNext
         wend         
         response.write("</select></td></tr>")


         response.write("<tr><td class='td-r fontmed bold whiteBorders'>Colaborador relacionado 2: </td><td  class='td-l fontmed bold whiteBorders'><select name='frelacionado2'>")

          if rsUpdateEntry("relacionado2")="" then
             response.write("<option value='' selected>ninguno</option>")
          else
             response.write("<option value=''>ninguno</option>")
          end if

          rsUpdateEntry2.MoveFirst
          while not  rsUpdateEntry2.EOF
              if UCASE(rsUpdateEntry("relacionado2"))=UCASE(rsUpdateEntry2("ID_USUARIO")) then
                    response.write("<option value="& rsUpdateEntry2("ID_USUARIO") &" SELECTED>"& rsUpdateEntry2("nombres") &" " & rsUpdateEntry2("PATERNO") &" " & rsUpdateEntry2("MATERNO") &"</option>")     
              else
                   response.write("<option value="& rsUpdateEntry2("ID_USUARIO") &">"& rsUpdateEntry2("nombres") &" " & rsUpdateEntry2("PATERNO") &" " & rsUpdateEntry2("MATERNO") &"</option>")    
              end if         
              rsUpdateEntry2.MoveNext
         wend         
         response.write("</select></td></tr>")


          response.write("<tr><td class='td-r fontmed bold whiteBorders'>Colaborador relacionado 3: </td><td  class='td-l fontmed bold whiteBorders'><select name='frelacionado3'>")

          if rsUpdateEntry("relacionado3")="" then
             response.write("<option value='' selected>ninguno</option>")
          else
             response.write("<option value=''>ninguno</option>")
          end if

          rsUpdateEntry2.MoveFirst
          while not  rsUpdateEntry2.EOF
              if UCASE(rsUpdateEntry("relacionado3"))=UCASE(rsUpdateEntry2("ID_USUARIO")) then
                    response.write("<option value="& rsUpdateEntry2("ID_USUARIO") &" SELECTED>"& rsUpdateEntry2("nombres") &" " & rsUpdateEntry2("PATERNO") &" " & rsUpdateEntry2("MATERNO") &"</option>")     
              else
                   response.write("<option value="& rsUpdateEntry2("ID_USUARIO") &">"& rsUpdateEntry2("nombres") &" " & rsUpdateEntry2("PATERNO") &" " & rsUpdateEntry2("MATERNO") &"</option>")    
              end if         
              rsUpdateEntry2.MoveNext
         wend         
         response.write("</select></td></tr>")

      
         response.write ("<tr><td class='td-l fontmed bold whiteBorders'>Estatus: </td><td  class='td-l fontmed whiteBorders'><select name='festatus'><option value=0>dejar abierta</option><option value=1>cerrar tarea</option></select>&nbsp;&nbsp;&nbsp;&nbsp; check si quiere reenviar actualizacion por email <input type='checkbox' name='fenviar'> </td></tr>")
         
         if rsUpdateEntry("closed")=0 then
              response.write ("<tr><td class='td-c fontmed bold' colspan=2><input type=submit value='actualizar tarea'></td></tr></form>")
         end if

         rsUpdateEntry2.Close
         rsUpdateEntry.Close


         closeTable
              

     case 4  'UP EDIT TASK' 

         if request("fOC")="OC-" then
              vOCpintado=""
         end if
         if request("fOC")<>"" then
             vOCpintado=request("fOC")
         end if
         if request("fpintando")="true" or request("fpintando")="on" then
              vPintandose=1
         else
              vPintandose=0
         end if

         vRecordar=0
         if  request("recordar")="true" or request("recordar")="on"  then
               vRecordar=1
         end if

         if request("fprogramada")="" then
             if request("fseguimiento")<>"" then
                if request("festatus")=1 then   'elijio cerrar la tarea'
                   strSQL="set dateformat ymd;update TaskManager set EnviarRecordatorio="& vRecordar &",OC_Pintado='"& vOCpintado &"',Pintandose="& vPintandose&", asunto='"&request("fasunto") &"',relacionados='"& UCASE(request("frelacionado")) &"',relacionado2='"& UCASE(request("frelacionado2")) &"',relacionado3='"& UCASE(request("frelacionado3")) &"',Responsable='"& UCASE(request("fresponsable")) &"',FechaCierre=getdate(),LogDate=getdate(),Seguimiento=isnull(Seguimiento,'')+'" &request("fseguimiento") &"("& UCASE(session("session_id")) &" '+ convert(varchar, getdate(), 113) +')<BR />' where id=" &request("ID")
                else
                   strSQL="set dateformat ymd;update TaskManager set EnviarRecordatorio="& vRecordar &",OC_Pintado='"& vOCpintado &"',Pintandose="& vPintandose&", asunto='"&request("fasunto") &"',relacionados='"& UCASE(request("frelacionado")) &"',relacionado2='"& UCASE(request("frelacionado2")) &"',relacionado3='"& UCASE(request("frelacionado3")) &"',Responsable='"& UCASE(request("fresponsable")) &"',Seguimiento=isnull(Seguimiento,'')+'" &request("fseguimiento") &"("& UCASE(session("session_id")) &" '+ convert(varchar, getdate(), 113) +')<BR />' where id=" &request("ID")

                end if
             else 'no hay seguimiento'
             if request("festatus")=1 then   'elijio cerrar la tarea'
                   strSQL="set dateformat ymd;update TaskManager set EnviarRecordatorio="& vRecordar &",OC_Pintado='"& vOCpintado &"',Pintandose="& vPintandose&", asunto='"&request("fasunto") &"',relacionados='"& UCASE(request("frelacionado")) &"',relacionado2='"& UCASE(request("frelacionado2")) &"',relacionado3='"& UCASE(request("frelacionado3")) &"',Responsable='"& UCASE(request("fresponsable")) &"',FechaCierre=getdate(),LogDate=getdate() where id=" &request("ID")
              else
                   strSQL="set dateformat ymd;update TaskManager set EnviarRecordatorio="& vRecordar &",OC_Pintado='"& vOCpintado &"',Pintandose="& vPintandose&", asunto='"&request("fasunto") &"',relacionados='"& UCASE(request("frelacionado")) &"',relacionado2='"& UCASE(request("frelacionado2")) &"',Responsable='"& UCASE(request("fresponsable")) &"' where id=" &request("ID")

              end if
             end if
         else  'se cambia fecha de agenda'

            if request("fseguimiento")<>"" then
                if request("festatus")=1 then   'elijio cerrar la tarea'
                   strSQL="set dateformat ymd;update TaskManager set EnviarRecordatorio="& vRecordar &",OC_Pintado='"& vOCpintado &"',Pintandose="& vPintandose&", asunto='"&request("fasunto") &"',relacionados='"& UCASE(request("frelacionado")) &"',relacionado2='"& UCASE(request("frelacionado2")) &"',relacionado3='"& UCASE(request("frelacionado3")) &"',Responsable='"& request("fresponsable") &"',FechaCierre=getdate(),LogDate=getdate(),FechaProgramada='"&request("fprogramada") &"',Seguimiento=isnull(Seguimiento,'')+'" &request("fseguimiento") &"("& UCASE(session("session_id")) &" '+ convert(varchar, getdate(), 113) +')<BR />' where id=" &request("ID")
                else
                   strSQL="set dateformat ymd;update TaskManager set EnviarRecordatorio="& vRecordar &",OC_Pintado='"& vOCpintado &"',Pintandose="& vPintandose&", asunto='"&request("fasunto") &"',relacionados='"& UCASE(request("frelacionado")) &"',relacionado2='"& UCASE(request("frelacionado2")) &"',relacionado3='"& UCASE(request("frelacionado3")) &"',Responsable='"& request("fresponsable") &"',FechaProgramada='"&request("fprogramada") &"',Seguimiento=isnull(Seguimiento,'')+'" &request("fseguimiento") &"( "& UCASE(session("session_id")) &" '+ convert(varchar, getdate(), 113) +')<BR />' where id=" &request("ID")

                end if
             else 'no hay seguimiento'
             if request("festatus")=1 then   'elijio cerrar la tarea'
                   strSQL="set dateformat ymd;update TaskManager set EnviarRecordatorio="& vRecordar &",OC_Pintado='"& vOCpintado &"',Pintandose="& vPintandose&", asunto='"&request("fasunto") &"',relacionados='"& UCASE(request("frelacionado")) &"',relacionado2='"& UCASE(request("frelacionado2")) &"',relacionado3='"& UCASE(request("frelacionado3")) &"',Responsable='"& UCASE(request("fresponsable")) &"',FechaCierre=getdate(),LogDate=getdate(),FechaProgramada='"&request("fprogramada") &"' where id=" &request("ID")
              else
                   strSQL="set dateformat ymd;update TaskManager set EnviarRecordatorio="& vRecordar &",OC_Pintado='"& vOCpintado &"',Pintandose="& vPintandose&", asunto='"&request("fasunto") &"',relacionados='"& UCASE(request("frelacionado")) &"',relacionado2='"& UCASE(request("frelacionado2")) &"',relacionado3='"& UCASE(request("frelacionado3")) &"',Responsable='"& UCASE(request("fresponsable")) &"',FechaProgramada='"&request("fprogramada") &"' where id=" &request("ID")

              end if
             end if



         end if
         response.write strSQL
         rsUpdateEntry.Open strSQL, adoCon

         if request("fenviar")="on" or request("fenviar")="true" then
             strSQL="update TaskManager set sentdate=null where ID=" &request("ID")
             rsUpdateEntry2.Open strSQL, adoCon
         end if
         response.redirect("ShowContent.asp?contenido=62&action=0&msg=se actualizo tarea no. "& request("ID") )




      case 9  'busqueda' 
                               
                 Titulo="B&uacute;squeda en tareas"
                 DoTitulo(Titulo)
                 
                 if len(trim(request("fpalabra")))>=3 then

                 strSQL="select a.*,case when FechaProgramada<getdate() then 1 else 0 end as Urgencia,dbo.fn_GetMonthName(a.DocDate,'spanish') as FechaX,dbo.fn_GetMonthName(a.FechaProgramada,'spanish') as FechaY, case when FechaCierre is null then 0 else 1 end as Estatus,b.nombres+' '+b.paterno as Ejecutor,c.nombres+' '+c.paterno as Beneficiario,d.nombres+' '+d.paterno as Relacion,e.nombres+' '+e.paterno as relacion2 from TaskManager a left join Empleados b on b.FechaEgreso is null and  a.Responsable=b.id_usuario left join Empleados c on c.FechaEgreso is null and  a.Solicitante=c.id_usuario left join Empleados d on d.FechaEgreso is null and d.ID_USUARIO<>'' and  a.Relacionados=d.id_usuario left join Empleados e on e.FechaEgreso is null and e.ID_USUARIO<>'' and  a.relacionado2=e.id_usuario where ( a.ID like '%"& trim(request("fpalabra")) &"%' OR  Asunto like '%"& trim(request("fpalabra")) &"%' OR Solicitud like '%"& trim(request("fpalabra")) &"%' OR Seguimiento like '%"& trim(request("fpalabra")) &"%' )  and ( a.Relacionados='"& session("usuario_id") &"' OR a.Relacionado2='"& session("usuario_id") &"' OR a.Relacionado3='"& session("usuario_id") &"' ) order by a.Responsable" 

                'response.write(strSQL &"<BR>")
                rsUpdateEntry.Open strSQL, adoCon  'intranet'
                
                ShowTasks
                rsUpdateEntry.close
               
                   else
                        NoRegistros
                   end if



     case 10     response.write("<P style='font-size:4px'>&nbsp;</P>")  

                 if request("msg")<>"" then
                   response.write("<P class='alert td-c'>"&  request("msg") &"</P>")                   
                 end if 
                 response.write("<form id='tareas' method='POST' action='ShowContent.asp'>") 
                 response.write("<input type='hidden' name='contenido' value='62'>")
                 response.write("<input type='hidden' name='action' value='10'>")

                 response.write("<P class='td-c'>Filtrar por [fecha inicial] <input type='date' name='ffechai'> a [fecha final] <input type='date' name='ffechaf'>")
                 response.write("<input type='submit' value='filtrar'></form></P><P>&nbsp;</P>")
                



                 response.write("<H1>Tareas cerradas, responsable/relacionado: "& session("session_id") &"</H1>")

                 strSQL="select a.*,case when FechaProgramada<getdate() then 1 else 0 end as Urgencia,dbo.fn_GetMonthName(a.DocDate,'spanish') as FechaX,dbo.fn_GetMonthName(a.FechaProgramada,'spanish') as FechaY, case when FechaCierre is null then 0 else 1 end as Estatus,b.nombres+' '+b.paterno as Ejecutor,c.nombres+' '+c.paterno as Beneficiario,d.nombres+' '+d.paterno as Relacion,e.nombres+' '+e.paterno as relacion2 from TaskManager a left join Empleados b on b.FechaEgreso is null and  a.Responsable=b.id_usuario left join Empleados c on c.FechaEgreso is null and  a.Solicitante=c.id_usuario left join Empleados d on d.FechaEgreso is null and d.ID_USUARIO<>'' and  a.Relacionados=d.id_usuario left join Empleados e on e.FechaEgreso is null and e.ID_USUARIO<>'' and  a.relacionado2=e.id_usuario where a.FechaCierre is not null and ( a.Responsable='"& session("session_id")  &"' OR a.Relacionados='"& session("session_id")  &"' OR a.Relacionado2='"& session("session_id")  &"' OR a.Relacionado3='"& session("session_id")  &"' )   " 

                 if request("ffechai")<>"" and request("ffechaf")<>"" then
                 	 strSQL= strSQL &"and FechaCierre>='"& request("ffechai") &"' and FechaCierre<'"& request("ffechaf") &"' "

                 end if

                 strSQL= strSQL &" order by a.ID desc "

                'response.write(strSQL &"<BR>")
                rsUpdateEntry.Open strSQL, adoCon  'intranet'

                
                ShowTasks
                rsUpdateEntry.close



           case else response.write("error")

   end select


end sub


sub modulosTask
   response.write("<center><P style='font-size:4px'>&nbsp;</P>")  
   titulo="M&oacute;dulos disponibles:" 
   doTitulo(titulo)      %>

<table width="900px" border=0 cellpadding="1" cellspacing="1" align="center">
<tr><td width="75%">

    <table width="500px" border=0 cellpadding="1" cellspacing="1" align="center">

             <tr><td width="18%" class="td-c submenu"><a href='ShowContent.asp?contenido=62&action=0'>Tareas abiertas</a></td>   
                 <td width="18%" class="td-c submenu"><a href='ShowContent.asp?contenido=62&action=1'>Crear tarea</a></td> 
                 <td width="18%" class="td-c submenu"><a href='ShowContent.asp?contenido=62&action=10'>Tareas cerradas</a></td>           
                 <td width="18%" class="submenu td-r"><a href='#'>Buscar:</a> </td> 

                 <form id="buscar" action="ShowContent.asp" method="POST">         
                 <input type="hidden" name="contenido" value="62">
                 <input type="hidden" name="action" value="9">

                 <td width="18%" class="td-c submenu"><input type="text" name="fpalabra" value="<%=request("fpalabra")%>" style='width: 140px'></td>
                 <td width="10%" class="td-r submenu"><input type="submit" value="buscar"></form></td>  
                            
    </table>  
    </td>
    <td><span class="fontmed strong"> Filtrar por responsable: </span><BR> <%

      strSQL="select b.id_usuario,b.Nombres+' '+b.Paterno  as colaborador  from Empleados b inner join (  select distinct(a.Responsable)   from TaskManager a  where a.FechaCierre is null and ( a.Responsable='"& session("session_id") &"' OR a.Relacionados='"& session("session_id") &"' OR relacionado2='"& session("session_id") &"' OR relacionado3='"& session("session_id") &"')  ) as aux  on b.id_usuario=aux.Responsable"

      rsUpdateEntry.Open strSQL, adoCon

      while not  rsUpdateEntry.EOF
          response.write( "<a href='ShowContent.asp?contenido=62&responsable=" & rsUpdateEntry("id_usuario") &"'>" & rsUpdateEntry("colaborador") &"</a><br>" )
          rsUpdateEntry.MoveNext
      wend
      response.write( "<a href='ShowContent.asp?contenido=62&responsable=none'>remover filtro</a><br>" )
      rsUpdateEntry.close


    %>

     </td>
  </tr>
</table>
<%
end sub  



sub ShowRecordsCOVID   'contenido 63'
   response.write("<P>&nbsp;</P>")
         if request("msg")<>"" then
           response.write("<P class='alert td-c'>"&  request("msg") &"</P>")
         end if

         
   response.write("<H1>Registros de colaboradores contra el COVID19</H1><P>&nbsp;</P><center><table width=920px align=center cellpadding=3 cellspacing=2 border=1>")

   if request("estatus")="1" then
      strSQL="select UPPER(dbo.fn_GetMonthName(DOCDATE,'spanish')) as Fecha,COLABORADOR,TOS,RESPIRAR,FIEBRE,TEMPERATURA,OXIGENO,SINTOMAS,CONTACTO from COVID where month(docdate)=month(getdate())-1 and year(docdate)=year(getdate()) order by day(DOCDATE)"
   else
      strSQL="select UPPER(dbo.fn_GetMonthName(DOCDATE,'spanish')) as FECHA,COLABORADOR,TOS,RESPIRAR,FIEBRE,TEMPERATURA,OXIGENO,SINTOMAS,CONTACTO from COVID where month(docdate)=month(getdate()) and year(docdate)=year(getdate()) order by day(DOCDATE)"
   end if

   rsUpdateEntry.Open strSQL, adoCon
   rsUpdateEntry2.Open strSQL, adoCon

   CamposRS1
   ShowValoresRS2

   response.write("</table><P>&nbsp;</p><P class='td-c'><a href='ShowContent.asp?contenido=63&estatus=1'>Mostrar mes anterior</a></p>")

end sub




sub ShowStockTubos  'contenido 64'
    response.write("<P>&nbsp;</P>")
         if request("msg")<>"" then
           response.write("<P class='alert td-c'>"&  request("msg") &"</P>")
         end if
   response.write("<H1>TUBERIA SOLICITADA Y SU INVENTARIO</H1><center><table width=1100px align=center cellpadding=3 cellspacing=2 border=1>")

   strSQL="EXEC SP_PlaneacionPintado"
   rsUpdateEntry5.Open strSQL, adoCon4

   strSQL="SELECT x.ItemCode,substring(Y.ItemName,1,30) as Descripcion,cast( SUM(x.OpenQty) as int) as Requerido,Pedidos=( STUFF(            ( SELECT ',' + RazonSocial+'-'+CONVERT(varchar,Docnum) FROM TEMP_TUBOS where ItemCode=x.ItemCode    FOR XML PATH ('')), 1, 1, ''    ) ),   MXL=cast( isnull( ( select SUM(Stock) from TEMP_INV where WhsCode='001-Mxl' and ItemCode=x.ItemCode  ) , 0 ) as int),      SJR=cast( isnull( ( select SUM(Stock) from TEMP_INV where WhsCode='002-SJR' and ItemCode=x.ItemCode  ) , 0 ) as int),      TRAS=cast( isnull( ( select SUM(Stock) from TEMP_INV where WhsCode='TR' and ItemCode=x.ItemCode  ) , 0 ) as int),      PR=cast( isnull( ( select SUM(Stock) from TEMP_INV where WhsCode='PR' and ItemCode=x.ItemCode  ) , 0 ) as int),   AD=cast( isnull( ( select SUM(Stock) from TEMP_INV where WhsCode='AD' and ItemCode=x.ItemCode  ) , 0 ) as int), VIRT=cast( isnull( ( select SUM(Stock) from TEMP_INV where WhsCode='006-Virt' and ItemCode=x.ItemCode  ) , 0 ) as int),  TOTAL=cast( isnull( ( select SUM(Stock) from TEMP_INV where ItemCode=x.ItemCode  ) , 0 ) as int) , x.Otro_color,      MXL2=cast( isnull( ( select SUM(Stock) from TEMP_INV where WhsCode='001-Mxl' and ItemCode collate SQL_Latin1_General_CP1_CI_AS=x.Otro_color  collate SQL_Latin1_General_CP1_CI_AS ) , 0 ) as int),      SJR2=cast( isnull( ( select SUM(Stock) from TEMP_INV where WhsCode='002-SJR' and ItemCode collate SQL_Latin1_General_CP1_CI_AS=x.Otro_color  collate SQL_Latin1_General_CP1_CI_AS   ) , 0 ) as int),      TRAS2=cast( isnull( ( select SUM(Stock) from TEMP_INV where WhsCode='TR' and ItemCode collate SQL_Latin1_General_CP1_CI_AS=x.Otro_color  collate SQL_Latin1_General_CP1_CI_AS   ) , 0 ) as int),      PR2=cast( isnull( ( select SUM(Stock) from TEMP_INV where WhsCode='PR' and ItemCode collate SQL_Latin1_General_CP1_CI_AS=x.Otro_color  collate SQL_Latin1_General_CP1_CI_AS   ) , 0 ) as int),     AD2=cast( isnull( ( select SUM(Stock) from TEMP_INV where WhsCode='AD' and ItemCode collate SQL_Latin1_General_CP1_CI_AS=x.Otro_color  collate SQL_Latin1_General_CP1_CI_AS   ) , 0 ) as int),  VIRT2=cast( isnull( ( select SUM(Stock) from TEMP_INV where WhsCode='006-Virt' and ItemCode collate SQL_Latin1_General_CP1_CI_AS=x.Otro_color  collate SQL_Latin1_General_CP1_CI_AS   ) , 0 ) as int),   TOTAL2=cast( isnull( ( select SUM(Stock) from TEMP_INV where ItemCode collate SQL_Latin1_General_CP1_CI_AS =x.Otro_color collate SQL_Latin1_General_CP1_CI_AS ) , 0 ) as int)           FROM TEMP_TUBOS x left join PRODUCTIVA_DMX.dbo.OITM Y on x.ItemCode=Y.ItemCode  where Y.ItemName not like '%chapet%'  Group by x.ItemCode,Y.ItemName,x.Otro_color  Order by x.ItemCode "

   'response.Write strSQL
   rsUpdateEntry.Open strSQL, adoCon4
   rsUpdateEntry2.Open strSQL, adoCon4

   CamposRS1
   ShowValoresRS2

   response.write("</table><P>&nbsp;</P>")   
   response.write("<H1>INVENTARIO DE PVC</H1><center><table width=1100px align=center cellpadding=3 cellspacing=2 border=1>")

   strSQL="select distinct(x.ItemCode) ,y.itemname, MXL=cast( isnull( ( select SUM(Stock) from TEMP_INV where WhsCode='001-Mxl' and ItemCode=x.ItemCode  ) , 0 ) as int), SJR=cast( isnull( ( select SUM(Stock) from TEMP_INV where WhsCode='002-SJR' and ItemCode=x.ItemCode  ) , 0 ) as int),      TRAS=cast( isnull( ( select SUM(Stock) from TEMP_INV where WhsCode='TR' and ItemCode=x.ItemCode  ) , 0 ) as int),       PR=cast( isnull( ( select SUM(Stock) from TEMP_INV where WhsCode='PR' and ItemCode=x.ItemCode  ) , 0 ) as int),   AD=cast( isnull( ( select SUM(Stock) from TEMP_INV where WhsCode='AD' and ItemCode=x.ItemCode  ) , 0 ) as int),          TOTAL=cast( isnull( ( select SUM(Stock) from TEMP_INV where ItemCode=x.ItemCode  ) , 0 ) as int)   from TEMP_INV x left join PRODUCTIVA_DMX.dbo.OITM y on x.ItemCode=y.ItemCode  where x.ItemCode COLLATE SQL_Latin1_General_CP1_CI_AS not in ( SELECT itemcode COLLATE SQL_Latin1_General_CP1_CI_AS FROM TEMP_TUBOS ) and x.ItemCode COLLATE SQL_Latin1_General_CP1_CI_AS not in ( SELECT Otro_color COLLATE SQL_Latin1_General_CP1_CI_AS FROM TEMP_TUBOS ) and x.ItemName like 'TUB%' AND x.ItemName like '%PVC%'  order by x.ItemCode DESC"

   'response.Write strSQL
   rsUpdateEntry5.Open strSQL, adoCon4
   rsUpdateEntry6.Open strSQL, adoCon4

   CamposRS5
   ShowValoresRS6


   response.write("</table><P>&nbsp;</P>") 
   response.write("<H1>INVENTARIO DE OTROS CODIGOS DE TUBERIA</H1><center><table width=1100px align=center cellpadding=3 cellspacing=2 border=1>")

   strSQL="select distinct(x.ItemCode) ,y.itemname, MXL=cast( isnull( ( select SUM(Stock) from TEMP_INV where WhsCode='001-Mxl' and ItemCode=x.ItemCode  ) , 0 ) as int), SJR=cast( isnull( ( select SUM(Stock) from TEMP_INV where WhsCode='002-SJR' and ItemCode=x.ItemCode  ) , 0 ) as int),      TRAS=cast( isnull( ( select SUM(Stock) from TEMP_INV where WhsCode='TR' and ItemCode=x.ItemCode  ) , 0 ) as int),       PR=cast( isnull( ( select SUM(Stock) from TEMP_INV where WhsCode='PR' and ItemCode=x.ItemCode  ) , 0 ) as int),   AD=cast( isnull( ( select SUM(Stock) from TEMP_INV where WhsCode='AD' and ItemCode=x.ItemCode  ) , 0 ) as int),          TOTAL=cast( isnull( ( select SUM(Stock) from TEMP_INV where ItemCode=x.ItemCode  ) , 0 ) as int)   from TEMP_INV x left join PRODUCTIVA_DMX.dbo.OITM y on x.ItemCode=y.ItemCode  where x.ItemCode COLLATE SQL_Latin1_General_CP1_CI_AS not in ( SELECT itemcode COLLATE SQL_Latin1_General_CP1_CI_AS FROM TEMP_TUBOS ) and x.ItemCode COLLATE SQL_Latin1_General_CP1_CI_AS not in ( SELECT Otro_color COLLATE SQL_Latin1_General_CP1_CI_AS FROM TEMP_TUBOS ) and x.ItemName like 'TUB%' AND x.ItemName not like '%PVC%'  order by x.ItemCode"

   'response.Write strSQL
   rsUpdateEntry3.Open strSQL, adoCon4
   rsUpdateEntry4.Open strSQL, adoCon4

   CamposRS3
   ShowValoresRS4

   response.write("</table><P>&nbsp;</P><P CLASS='td-c fontmed'>* ambas razones sociales</P>")

  

   strSQL="if exists (select * from INFORMATION_SCHEMA.TABLES where TABLE_NAME = 'TEMP_TUBOS' AND TABLE_SCHEMA = 'dbo')     drop table dbo.TEMP_TUBOS;"
   rsUpdateEntry7.Open strSQL, adoCon4

   strSQL="if exists (select * from INFORMATION_SCHEMA.TABLES where TABLE_NAME = 'TEMP_INV' AND TABLE_SCHEMA = 'dbo')    drop table dbo.TEMP_INV;"
   rsUpdateEntry7.Open strSQL, adoCon4

end sub  





sub interTransaction    'contenido 65 COMPRAS'

    DoAlert

    if request("fRS")<>"" then     
        select case request("action")
        case 1    Titulo="Precio Interempresa (USD) de "& request("fRS")
        case 2    Titulo="Precio Interempresa (USD) de "& request("fRS") &" (TC: "& request("fRate") &")"
        case 3    titulo="Precio Interempresa (USD) de "& request("fRS") &" (TC: "& request("fRate") &") "
                 
        end select
        
    else
         titulo="Costos de Transferencia InterEmpresas" 
     end if

     DoTitulo(titulo)

    if request("action")="" then
        response.write("<form id='RS' action='ShowContent.asp' method='POST'><input type='hidden' name='contenido' value=65><input type='hidden' name='action' value=1>")
        response.write("<P class='td-c'>RS de d&oacute;nde saldr&aacute; el stock o quien es proveedor: <select name='fRS'><option value='DMX'>DMX</option><option value='DFW'>DFW</option><option value='MME'>MME</option></select><BR><BR><BR><input type='submit' value='continuar'></P></form>")
    end if



    if request("action")=1 then 

        response.write("<form id='RS' action='ShowContent.asp' method='POST'>")
        response.write("<input type='hidden' name='contenido' value=65>")
        response.write("<input type='hidden' name='action' value=2>") 
        response.write("<input type='hidden' name='fRS' value='"& request("fRS") &"'>")
        response.write("<table border=1 cellspacing=2 cellpadding=3 align=center width='330px'>")
        response.write("<tr><td class='td-r subtitulo'>Razon Social Origen: </td><td class='td-l'>" & request("fRS") &"</td></tr>")     
        response.write("</table>")
        response.write("<P style='font-size: 4px'>&nbsp;</P><center><table border=1 cellspacing=2 cellpadding=3 align=center width='500px'>")
        strSQL="select Rate,RateDate from ORTT where year(RateDate)=year(getdate()) and month(RateDate)=month(getdate()) and day(RateDate)=day(getdate())"
        'response.write strSQL
        rsUpdateEntry.Open strSQL, adoCon2 'DMX'

        vRate=""
        if rsUpdateEntry.EOF then
                vRate="20"
        else        
                vRate=trim(rsUpdateEntry("Rate"))
        end if
        rsUpdateEntry.close
        response.write("<tr><td class='td-r subtitulo'>El TC para el d&iacute;a es: <input type='text' name='frate' value='"& vRate &"'></td></tr></table>")
        response.write("<P class='td-c'>Pegue el listado de sus partidas en esta &aacute;rea:<BR><textarea rows=30 cols=25 name='fpartidas'></textarea></P>")


        response.write("<P class='td-c'><BR><input type='submit' value='continuar'></P></form>")
   end if




   if request("action")=2 then   'usuario ya pegó el listado de partidas'
      
        partidas=""
        partidas=trim(request("fpartidas"))
        partida=""
        Pos=1
        longitud=0
        strSQL=""

        while Pos<>0
          pos=InStr(partidas,chr(13))
          
          if pos=0 then
                partida=trim(partidas) 
          else
              partida=trim(mid(Partidas,1,Pos-1))
                Partidas=mid(Partidas,Pos+2,2000)
            end if
           
            if len(trim(partida))>1 then
               'strSQL=strSQL & "'"& trim(partida) &"',"  
               strSQL=strSQL & trim(partida) &"*"  
            end if    
            'response.write( strSQL &"<BR>")
        wend
        vPos=len(strSQL)

        strSQL="exec [dbo].[SP_MUB_InterEmpresa]  '"&request("fRS")&"','"&  mid(strSql,1,vPos-1) &"'"
        'response.write strSQL
        rsUpdateEntry.Open strSQL, adoCon4

        strSQL="select itemcode,stock_mxl,round(costo_mxl,2) as costo_mxl,stock_sjr,round(costo_sjr,2) as costo_sjr,stock_mty,round(costo_mty,2) as costo_mty,stock_mxl+stock_sjr+stock_mty as STOCK_TOTAL  from  _MUB  order by itemcode"
        'response.write strSQL
        rsUpdateEntry.Open strSQL, adoCon4
        rsUpdateEntry2.Open strSQL, adoCon4
        
        response.write("<P style='font-size: 4px'>&nbsp;</P><center><table border=1 cellspacing=2 cellpadding=3 align=center width='900px'>")
   
       response.write("<tr><td class='td-c'><table border=0 cellspacing=2 cellpadding=3 align=center width='600px' style='vertical-align: top'>")
        CamposRS1
        ShowValoresRS2
      
        strSQL="select *,stock_mxl+stock_sjr+stock_mty as STOCK_TOTAL from _MUB order by itemcode"
        'response.write strSQL
        rsUpdateEntry4.Open strSQL, adoCon4

        i=1
        response.write("<form id='partidas' action='ShowContent.asp' method='POST'>")
        response.write("<input type='hidden' name='contenido' value=65>")
        response.write("<input type='hidden' name='action' value=3>")
        response.write("<input type='hidden' name='fRS' value='"& request("fRS") &"'>")
        response.write("<input type='hidden' name='fRate' value='"& request("fRate") &"'>") 
        
        response.write("</table></td><td class='td-c'><table border=0 cellspacing=2 cellpadding=3 align=left width='180px' style='vertical-align: top'>") 

        response.write("<tr><td class='subtitulo'>Elija almac&eacute;n </td></tr>")

        while not rsUpdateEntry4.EOF  
        	response.write("<tr>")

            'cuando no hay stock'            
        	if int(trim(rsUpdateEntry4("stock_mxl")))=0 and int(trim(rsUpdateEntry4("stock_sjr")))=0 and int(trim(rsUpdateEntry4("stock_mty")))=0 then
        		                  
                       response.write("<td class='td-c'><select name='item"& i &"' style='width:50px; font-size: 11px' ><option value='MXL' SELECTED>MXL</option><option value='SJR'>SJR</option><option value='MTY'>MTY</option></select></td></tr>")
                
                       
                  
            else
                     response.write("<td class='td-c'><select name='item"& i &"' style='width:50px; font-size: 11px' ><option value='MXL' SELECTED>MXL</option><option value='SJR'>SJR</option><option value='MTY'>MTY</option></select></td></tr>")

            end if


            rsUpdateEntry4.MoveNext
            i=i+1
        wend  
        rsUpdateEntry4.close
        response.write("</table>")

        response.write("<tr><td colspan=10 class='td-c'><input type='submit' value='Finalizar'> </td></tr> </table> ")

   end if


   

   if request("action")=3 then  'paso final ya se ha elejido almacen'

       'si se mueve % cambiar en SAP -DMX- ID17-Precio inter-E o Tras descuentos'
       ' y cambiar ID22-Precios Inter-E que es para las PO'

      strSQL="select *,CASE when substring(ItemCode,1,5)='MEC10' OR substring(ItemCode,1,5)='MEC40' THEN round(costo_mxl/.85,2) ELSE round(costo_mxl/.91,2) END as Al5_mxl,CASE when substring(ItemCode,1,5)='MEC10' OR substring(ItemCode,1,5)='MEC40' THEN round(costo_sjr/.85,2) ELSE round(costo_sjr/.91,2) END as Al5_sjr,CASE when substring(ItemCode,1,5)='MEC10' OR substring(ItemCode,1,5)='MEC40' THEN round(costo_mty/.85,2) ELSE round(costo_mty/.91,2)  END as Al5_mty from  _MUB order by itemcode"
      'response.write strSQL
      rsUpdateEntry.Open strSQL, adoCon4 

       response.write("<tr><td class='td-c'><table border=0 cellspacing=2 cellpadding=3 align=center width='400px'>")
       i=1

             response.write("<tr><td class='subtitulo'>itemcode </td><td class='subtitulo'>Almacen</td><td class='subtitulo'>Precio Inter-E</td></tr>")  

                 while not rsUpdateEntry.EOF

                 	    vstring="item"&trim(i)  'se elijio almacen'
                      rowin       
                      response.write("<td class='subtitulo td-r'>" & rsUpdateEntry("ItemCode") &"</td>")
                      response.write("<td class='td-c'>" & request(vstring)  &"</td>")
                      select case request(vstring)
                            case "MXL"  response.write("<td class='td-r'>" & rsUpdateEntry("Al5_mxl") &"</td>")
                            case "SJR"  response.write("<td class='td-r'>" & rsUpdateEntry("Al5_sjr") &"</td>")
                            case "MTY"  response.write("<td class='td-r'>" & rsUpdateEntry("Al5_mty") &"</td>")
                      end select                         
                      rowout
                      rsUpdateEntry.MoveNext
                      i=i+1
                 wend
                 rsUpdateEntry.close 
  
        response.write("</table><P>&nbsp;</P>")


        'strSQL="if exists (select * from INFORMATION_SCHEMA.TABLES where TABLE_NAME = '_MUB' AND TABLE_SCHEMA = 'dbo')     drop table dbo._MUB;"
        'rsUpdateEntry7.Open strSQL, adoCon4 

   end if

end sub




sub DellateOCTaller 'contenido 66'

      strSQL="select top 1 * from TaskManager where FechaCierre is null and Pintandose=1 order by DocDate desc"
      'response.write(strSQL &"<BR>")
      rsUpdateEntry.Open strSQL, adoCon  'intranet'
      response.write("<P>&nbsp;</P><center><table width=750px align=center cellpadding=3 cellspacing=2 border=1>")

      if not rsUpdateEntry.EOF then
       
         response.write("<tr><td class='td-l fontmed bold whiteBorders'>Asunto de la Tarea No. "& rsUpdateEntry("ID")& " </td><td  class='td-l fontmed bold whiteBorders'>"& rsUpdateEntry("asunto") &"</td></tr>")
        
         response.write("<tr><td class='td-l fontmed bold whiteBorders'>Fecha de agenda: </td><td  class='td-l fontmed whiteBorders'>"& rsUpdateEntry("FechaProgramada") &"</td></tr>")
          
        response.write ("<tr><td class='td-l fontmed bold' colspan=2><font style='background-color:red; color: white'>APLICA A PROCESO DE PINTADO </font> <br>Orden del taller: <input type='text' name='fOC' class='anchoSmall' value='"& rsUpdateEntry("OC_Pintado") &"'><br> VA A PINTARSE &oacute;  YA ESTA EN TALLER: ")

         if rsUpdateEntry("Pintandose")=1 then  
            response.write("<input type='checkbox' name='fpintando' checked><P>&nbsp;</P></td></tr>")
         else
            response.write("<input type='checkbox' name='fpintando'><P>&nbsp;</P></td></tr>")
         end if

         response.write ("<tr><td class='td-l fontmed bold' colspan=2>SEGUIMIENTO: <BR>"& rsUpdateEntry("SEGUIMIENTO") &"</TD></TR>")


      else

       response.write("<tr><td class='td-l fontmed bold whiteBorders'>No existen &oacute;rdenes de producci&oacute;n en taller actualmente</td></tr>")
             

      end if

      response.write("</table><P>&nbsp;</P>")
      rsUpdateEntry.close



    response.write("<P>&nbsp;</P>")
         if request("msg")<>"" then
           response.write("<P class='alert td-c'>"&  request("msg") &"</P>")
         end if
   if request("fOC")<>"" then
       response.write("<H1>Partidas de la "& request("fOC") &" del Taller de Pintado</H1><P style='font-size: 4px;'>&nbsp;</P>")
   else
       response.write("<H1>Partidas de una OC del Taller de Pintado</H1><P style='font-size: 4px;'>&nbsp;</P>")
   end if

   if request("action")="" then
      response.write("<form id='RS' action='ShowContent.asp' method='POST'><input type='hidden' name='contenido' value=66><input type='hidden' name='action' value=1>")
      response.write("<P class='td-c'>Indique la OC: <input type='text' name='fOC' value='OC-'><BR><BR><BR><input type='submit' value='Mostrar partidas'></P></form>")

   end if

    if request("action")=1 then
      strSQL="select x.DocNum as 'OF',y.ItemCode as Otro_color,c.ItemName as Descripcion_negro,cast(y.PlannedQty as int) as Cantidad,x.ItemCode as Codigo_rojo,d.ItemName as Descripcion_rojo from OWOR x inner join WOR1 y on x.DocEntry=y.DocEntry left join OITM c on y.ItemCode=c.ItemCode left join OITM d on x.ItemCode=d.ItemCode where x.Comments like '%"& request("fOC") &"%' and y.ItemCode not like 'MAT%' order by x.DocNum "

      'response.write strSQL
      rsUpdateEntry.Open strSQL, adoCon2 'DMX'
      rsUpdateEntry2.Open strSQL, adoCon2 'DMX'

       response.write("<center><table border=1 cellspacing=2 cellpadding=3 align=center width='1060x'>")

       CamposRS1
       ShowValoresRS2

       response.write("</table>")

   end if


end sub


sub CorteInventario 'contenido 67 

    doAlert
    titulo="CORTE EN INVENTARIO"
    DoTitulo(titulo)

    if request("action")="" then

    dim vfecha1 : vfecha1=date()
    dim vfecha2 : vfecha2=date()
    dim vfecha3 : vfecha3=date()
    dim vfecha4 : vfecha4=date()


    vfecha1=cdate("01/" & month(date()) & "/"&  year(date())  )
    vfecha1=vfecha1-1

    vfecha2=vfecha2-30    
    vfecha2=cdate("01/" & month(vfecha2) & "/"&  year(vfecha2)  )
    vfecha2=vfecha2-1

    vfecha3=vfecha3-90
    vfecha3=cdate("01/" & month(vfecha3) & "/"&  year(vfecha3)  )
    vfecha3=vfecha3-1

    vfecha4=vfecha4-365
    vfecha4=cdate("01/" & month(vfecha4) & "/"&  year(vfecha4)  )
    vfecha4=vfecha4-1

    vFecha1Text=""
    vFecha1Text=trim(year(vfecha1))&"-"&trim(month(vfecha1))&"-"&trim(day(vfecha1))
    vFecha2Text=""
    vFecha2Text=trim(year(vfecha2))&"-"&trim(month(vfecha2))&"-"&trim(day(vfecha2))
    vFecha3Text=""
    vFecha3Text=trim(year(vfecha3))&"-"&trim(month(vfecha3))&"-"&trim(day(vfecha3))
    vFecha4Text=""
    vFecha4Text=trim(year(vfecha4))&"-"&trim(month(vfecha4))&"-"&trim(day(vfecha4))

    %>
    <br>
    <button role="link" onclick="window.location='ShowContent.asp?contenido=67&action=1&contable=1&vfecha=<%=vFecha1Text%>'">Inventario al <%=vfecha1%></button>
    <br><br>
    <button role="link" onclick="window.location='ShowContent.asp?contenido=67&action=1&vfecha=<%=vFecha2Text%>'">Inventario al <%=vfecha2%></button>
    <br><br>
    <button role="link" onclick="window.location='ShowContent.asp?contenido=67&action=1&vfecha=<%=vFecha3Text%>'">Inventario al <%=vfecha3%></button>
    <br><br>
    <button role="link" onclick="window.location='ShowContent.asp?contenido=67&action=1&vfecha=<%=vFecha4Text%>'">Inventario al <%=vfecha4%></button>

    <P>&nbsp;</P>
    <P class="td-c fonttiny">el proceso de generar el inventario a una fecha especifica puede demorar unos segundos</P>
    
    <%
    end if

    if request("action")<>"" then    
         strSQL="if exists (select * from INFORMATION_SCHEMA.TABLES where TABLE_NAME = 'TEMP_INV' AND TABLE_SCHEMA = 'dbo')  drop table dbo.TEMP_INV;" 
         rsUpdateEntry7.Open strSQL, adoCon4

         strSQL="exec SP_CorteInventario '"&request("vfecha")&"'"
         response.write("<P class='fonttiny'>query: "&strSQL&"</P>")
         rsUpdateEntry7.Open strSQL, adoCon4
            

         if request("contable")="1" then
         	 response.write("<P class='fontmed'><I>*Basado en informe de autitor&iacute;a de stock (VALOR CONTABLE) de SAP al "&request("vfecha") &"</I></P>") 

         	 strSQL="select CONVERT(VARCHAR,CONVERT(MONEY,sum(Costo)),1) as TotalMXN2,sum(Costo) as TotalMXN,CONVERT(VARCHAR,CONVERT(MONEY,sum(Costo_USD)),1) as TotalUSD2,sum(Costo_USD) as TotalUSD from TEMP_INV"
             rsUpdateEntry3.Open strSQL, adoCon4 

         	 strSQL="select RazonSocial,CONVERT(VARCHAR,CONVERT(MONEY,sum(Costo)),1) as CostoMXN,(select Rate from PRODUCTIVA_DMX.dbo.ORTT where RateDate='"&request("vfecha")&"') as TC,CONVERT(VARCHAR,CONVERT(MONEY,sum(Costo_USD)),1) as CostoUSD,round(sum(Costo_USD)/"& rsUpdateEntry3("TotalUSD") &"*100,2) as Pctje  from TEMP_INV group by RazonSocial"         

         	 'response.write("<BR>"&strSQL)	  

         	 rsUpdateEntry.Open strSQL, adoCon4
             rsUpdateEntry2.Open strSQL, adoCon4
         
             Dim Campos(20) 
             i=1

             response.write("<table border=1 cellspacing=2 cellpadding=3 width='500px' class='table2excel table2excel_with_colors' data-tableName='table1'>")  

             RowIn
             for Each rsUpdateEntry in rsUpdateEntry.Fields  
                 campos(i)=rsUpdateEntry.Name     
                 Response.Write("<th class='subtitulo td-c'>" & campos(i) & "</th>")                
                 i=i+1           
             Next        
             RowOut             

             while not rsUpdateEntry2.EOF 
                RowIn
                for x=1 to 5
                     response.write("<td class='td-r fontbig'>" & rsUpdateEntry2(campos(x)) &"</td>")
                next
                RowOut
                rsUpdateEntry2.MoveNext
             wend   
             rsUpdateEntry2.close
             separador
             
             rowin
                   response.write("<td class='td-r fontbig strong'>TOTAL</td><td class='td-r fontbig strong'>"&rsUpdateEntry3("TotalMXN2") &"</td><td class='td-r fontbig strong'>&nbsp;</td><td class='td-r fontbig strong'>"&rsUpdateEntry3("TotalUSD2") &"</td>")
             RowOut
             %>
             </table><button class="exportToExcel">Exportar a excel</button>  <P>&nbsp;</P> <%  
             

             strSQL="select distinct(a.WhsCode) as Almacen,CONVERT(VARCHAR,CONVERT(MONEY,isnull(DFW.CostoMXN,0)),1) as DFW_CostoMXN,CONVERT(VARCHAR,CONVERT(MONEY,isnull(DMX.CostoMXN,0)),1) as DMX_CostoMXN,CONVERT(VARCHAR,CONVERT(MONEY,isnull(Ambas.CostoMXN,0)),1) as Ambas,CONVERT(VARCHAR,CONVERT(MONEY,isnull(USD.CostoUSD,0)),1) as En_USD, round(USD.CostoUSD/"&rsUpdateEntry3("TotalUSD")&"*100,2) as Prctj from TEMP_INV a left join (    select WhsCode,sum(Costo) as CostoMXN from TEMP_INV where RazonSocial='DFW' group by WhsCode     ) as DFW on a.WhsCode=DFW.WhsCode left join (     select WhsCode,sum(Costo) as CostoMXN from TEMP_INV where RazonSocial='DMX' group by WhsCode     ) as DMX on a.WhsCode=DMX.WhsCode left join (      select WhsCode,sum(Costo) as CostoMXN from TEMP_INV group by WhsCode 	) as Ambas on a.WhsCode=Ambas.WhsCode left join (select WhsCode,sum(Costo_USD) as CostoUSD from TEMP_INV group by WhsCode ) as USD on a.WhsCode=USD.WhsCode order by a.WhsCode "

             'response.write("<BR>"&strSQL)	  

         	 rsUpdateEntry2.Open strSQL, adoCon4
             rsUpdateEntry4.Open strSQL, adoCon4
 
            response.write("<table border=1 cellspacing=2 cellpadding=3 width='750px' class='table2excel table2excel_with_colors' data-tableName='table1'>")

             i=1             
             RowIn
             for Each rsUpdateEntry2 in rsUpdateEntry2.Fields  
                 campos(i)=rsUpdateEntry2.Name     
                 Response.Write("<th class='subtitulo td-c'>" & campos(i) & "</th>")                
                 i=i+1           
             Next        
             RowOut   
             while not rsUpdateEntry4.EOF 
                RowIn
                for x=1 to 6
                     response.write("<td class='td-r fontbig'>" & rsUpdateEntry4(campos(x)) &"</td>")
                next
                RowOut
                rsUpdateEntry4.MoveNext
             wend   
             rsUpdateEntry4.close
             separador
             rowin
             response.write("<td colspan=3 class='td-r strong fontbig'>TOTAL</td><td class='td-r strong fontbig'>"&rsUpdateEntry3("TotalMXN2")&"</td><td class='td-r strong fontbig'>"&rsUpdateEntry3("TotalUSD2")&"</td><td class='td-r strong fontbig'>&nbsp;</td>")
             RowOut
             %>
             </table><button class="exportToExcel">Exportar a excel</button>  <P>&nbsp;</P>  <%

             rsUpdateEntry3.close

         end if
         
         response.write("<P>&nbsp;</P>")
         response.write("<P class='fontmed'><I>*Basado en listas de precio SAP (VALOR DEL MERCADO) al "&request("vfecha") &"</I></P>") 

         strSQL="SELECT distinct(a.ItmsGrpNam) as Familia,       CONVERT(VARCHAR,CONVERT(MONEY,isnull(DFW_MXL.ValorComercial,0)),1) as DFW_MXL,	   CONVERT(VARCHAR,CONVERT(MONEY,isnull(DMX_MXL.ValorComercial,0)),1) as DMX_MXL,	   CONVERT(VARCHAR,CONVERT(MONEY,isnull(DFW_SJR.ValorComercial,0)),1) as DFW_SJR,	   CONVERT(VARCHAR,CONVERT(MONEY,isnull(DMX_SJR.ValorComercial,0)),1) as DMX_SJR,	   CONVERT(VARCHAR,CONVERT(MONEY,isnull(DFW_MTY.ValorComercial,0)),1) as DFW_MTY,	   CONVERT(VARCHAR,CONVERT(MONEY,isnull(DMX_MTY.ValorComercial,0)),1) as DMX_MTY,	   CONVERT(VARCHAR,CONVERT(MONEY,isnull(DFW_VIRT.ValorComercial,0)),1) as DFW_VIRT,	   CONVERT(VARCHAR,CONVERT(MONEY,isnull(DMX_VIRT.ValorComercial,0)),1) as DMX_VIRT,	   CONVERT(VARCHAR,CONVERT(MONEY,isnull(DFW_PR.ValorComercial,0)),1) as DFW_PR,	   CONVERT(VARCHAR,CONVERT(MONEY,isnull(DMX_PR.ValorComercial,0)),1) as DMX_PR,	   CONVERT(VARCHAR,CONVERT(MONEY,isnull(DFW_TR.ValorComercial,0)),1) as DFW_TR,	   CONVERT(VARCHAR,CONVERT(MONEY,isnull(DMX_TR.ValorComercial,0)),1) as DMX_TR, CONVERT(VARCHAR,CONVERT(MONEY,isnull(RESG.ValorComercial,0)),1) as RESGR,	   CONVERT(VARCHAR,CONVERT(MONEY,isnull(DFW_TOT.ValorComercial,0)),1) as DFW_TOT,	   CONVERT(VARCHAR,CONVERT(MONEY,isnull(DMX_TOT.ValorComercial,0)),1) as DMX_TOT,	   CONVERT(VARCHAR,CONVERT(MONEY,isnull(TOTAL.ValorComercial,0)),1) as TOTAL,	   round(TOTAL.ValorComercial/(select sum(Valor_Comercial) from TEMP_INV )*100,2) as Prctj "

         strSQL=strSQL&"FROM TEMP_INV a left join (   select ItmsGrpNam,sum(Valor_Comercial) as ValorComercial from TEMP_INV where RazonSocial='DFW' and WhsCode in ('001-Mxl','AD','RESG-MXL') group by ItmsGrpNam    ) as DFW_MXL on a.ItmsGrpNam=DFW_MXL.ItmsGrpNam left join ( select ItmsGrpNam,sum(Valor_Comercial) as ValorComercial from TEMP_INV where RazonSocial='DMX' and WhsCode in ('001-Mxl','AD','RESG-MXL') group by ItmsGrpNam    ) as DMX_MXL on a.ItmsGrpNam=DMX_MXL.ItmsGrpNam    left join (    select ItmsGrpNam,sum(Valor_Comercial) as ValorComercial from TEMP_INV where RazonSocial='DFW' and WhsCode in ('002-SJR','RESG-SJR') group by ItmsGrpNam    ) as DFW_SJR on a.ItmsGrpNam=DFW_SJR.ItmsGrpNam left join ( select ItmsGrpNam,sum(Valor_Comercial) as ValorComercial from TEMP_INV where RazonSocial='DMX' and WhsCode in ('002-SJR','RESG-SJR') group by ItmsGrpNam    ) as DMX_SJR on a.ItmsGrpNam=DMX_SJR.ItmsGrpNam left join (    select ItmsGrpNam,sum(Valor_Comercial) as ValorComercial from TEMP_INV where RazonSocial='DFW' and WhsCode in ('003-Mty','RESG-MTY') group by ItmsGrpNam    ) as DFW_MTY on a.ItmsGrpNam=DFW_MTY.ItmsGrpNam left join ( select ItmsGrpNam,sum(Valor_Comercial) as ValorComercial from TEMP_INV where RazonSocial='DMX' and WhsCode in ('003-Mty','RESG-MTY') group by ItmsGrpNam    ) as DMX_MTY on a.ItmsGrpNam=DMX_MTY.ItmsGrpNam left join (  select ItmsGrpNam,sum(Valor_Comercial) as ValorComercial from TEMP_INV where RazonSocial='DFW' and WhsCode='006-Virt' group by ItmsGrpNam    ) as DFW_VIRT on a.ItmsGrpNam=DFW_VIRT.ItmsGrpNam left join (select ItmsGrpNam,sum(Valor_Comercial) as ValorComercial from TEMP_INV where RazonSocial='DMX' and WhsCode='006-Virt' group by ItmsGrpNam   ) as DMX_VIRT on a.ItmsGrpNam=DMX_VIRT.ItmsGrpNam left join (   select ItmsGrpNam,sum(Valor_Comercial) as ValorComercial from TEMP_INV where RazonSocial='DFW' and WhsCode='PR' group by ItmsGrpNam   ) as DFW_PR on a.ItmsGrpNam=DFW_PR.ItmsGrpNam left join (	select ItmsGrpNam,sum(Valor_Comercial) as ValorComercial from TEMP_INV where RazonSocial='DMX' and WhsCode='PR' group by ItmsGrpNam   ) as DMX_PR on a.ItmsGrpNam=DMX_PR.ItmsGrpNam left join (   select ItmsGrpNam,sum(Valor_Comercial) as ValorComercial from TEMP_INV where RazonSocial='DFW' and WhsCode='TR' group by ItmsGrpNam   ) as DFW_TR on a.ItmsGrpNam=DFW_TR.ItmsGrpNam left join (	select ItmsGrpNam,sum(Valor_Comercial) as ValorComercial from TEMP_INV where RazonSocial='DMX' and WhsCode='TR' group by ItmsGrpNam   ) as DMX_TR on a.ItmsGrpNam=DMX_TR.ItmsGrpNam left join ( select ItmsGrpNam,sum(Valor_Comercial) as ValorComercial from TEMP_INV where WhsCode in ('RESG-MXL','RESG-SJR','RESG-MTY') group by ItmsGrpNam ) as RESG on a.ItmsGrpNam=RESG.ItmsGrpNam   left join (   select ItmsGrpNam,sum(Valor_Comercial) as ValorComercial from TEMP_INV where RazonSocial='DFW' group by ItmsGrpNam   ) as DFW_TOT on a.ItmsGrpNam=DFW_TOT.ItmsGrpNam left join (	select ItmsGrpNam,sum(Valor_Comercial) as ValorComercial from TEMP_INV where RazonSocial='DMX' group by ItmsGrpNam   ) as DMX_TOT on a.ItmsGrpNam=DMX_TOT.ItmsGrpNam left join (    select ItmsGrpNam,sum(Valor_Comercial) as ValorComercial from TEMP_INV group by ItmsGrpNam   ) as TOTAL on a.ItmsGrpNam=TOTAL.ItmsGrpNam "

         strSQL=strSQL&"order by a.ItmsGrpNam"
         'response.write strSQL

         rsUpdateEntry5.Open strSQL, adoCon4
         rsUpdateEntry6.Open strSQL, adoCon4

         
         response.write("<table border=1 cellspacing=2 cellpadding=3 width='1200px' class='table2excel table2excel_with_colors' data-tableName='table1'>")  

             RowIn
             for Each rsUpdateEntry5 in rsUpdateEntry5.Fields  
                 Response.Write("<th class='subtitulo td-c fonttiny'>" & rsUpdateEntry5.Name & "</th>")                  
             Next        
             RowOut 
             Columnas = Split("Familia,DFW_MXL,DMX_MXL,DFW_SJR,DMX_SJR,DFW_MTY,DMX_MTY,DFW_VIRT,DMX_VIRT,DFW_PR,DMX_PR,DFW_TR,DMX_TR,RESGR,DFW_TOT,DMX_TOT,TOTAL,Prctj",",") 

             while not rsUpdateEntry6.EOF 
                RowIn
                for x=0 to 17
                     response.write("<td class='td-r fonttiny'>" & rsUpdateEntry6(Columnas(x)) &"</td>")
                next
                RowOut
                rsUpdateEntry6.MoveNext
             wend   
             rsUpdateEntry6.close   %>
              </table><button class="exportToExcel">Exportar a excel</button>  <P>&nbsp;</P> <%

             strSQL="select CONVERT(VARCHAR,CONVERT(MONEY,sum(Valor_Comercial)),1) as ValorComercial from TEMP_INV"
             rsUpdateEntry6.Open strSQL, adoCon4

             CreateTable(350)
             RowIn
             response.write("<td class='td-r subtitulo fontbig'>VALOR TOTAL DE MERCADO</td><td class='td-r fontbig'>"&rsUpdateEntry6("ValorComercial")&"</td>")
             RowOut             
             rsUpdateEntry6.close
             closeTable

             strSQL="if exists (select * from INFORMATION_SCHEMA.TABLES where TABLE_NAME = 'TEMP_INV' AND TABLE_SCHEMA = 'dbo')  drop table dbo.TEMP_INV;" 
             rsUpdateEntry7.Open strSQL, adoCon4

             %>
              

         
    <script>
      $(function() {
        $(".exportToExcel").click(function(e){
          var table = $(this).prev('.table2excel');
          if(table && table.length){
            var preserveColors = (table.hasClass('table2excel_with_colors') ? true : false);
            $(table).table2excel({
              exclude: ".noExl",
              name: "Excel Document Name",
              filename: "myFileName" + new Date().toISOString().replace(/[\-\:\.]/g, "") + ".xls",
              fileext: ".xls",
              exclude_img: true,
              exclude_links: true,
              exclude_inputs: true,
              preserveColors: preserveColors
            });
          }
        });
        
      });
    </script>
    <%
    end if



end sub





sub PedidoInventario 'contenido 68'
    response.write("<P>&nbsp;</P>")
    if request("msg")<>"" then
           response.write("<P class='alert td-c'>"&  request("msg") &"</P>")
    end if

    


    if request("action")="" then
    	   titulo="Mostrar stock disponible de un pedido"
         DoTitulo(titulo)

         response.write("<form id='inv' method='POST' action='ShowContent.asp'>")
         response.write("<input type='hidden' name='contenido' value=68>")
         response.write("<input type='hidden' name='action' value=1>")


         response.write("<P class='td-c'>Raz&oacute;n Social: <select name='fRS'><option value='DFW'>DFW</option><option value='DMX'>DMX</option><option value='MEIDE'>MME</option></select>")
         response.write("&nbsp;&nbsp;&nbsp;Pedido: <input class='anchoSmall' type='numeric' name='fpedido' value="&request("fpedido") &"></P>")

         response.write("<P class='td-c'>&nbsp;</P><P class='td-c'><input type='submit' value='continuar'></form></P>")

    else
    
         titulo="Mostrar stock disponible del pedido "&request("fpedido")
         DoTitulo(titulo)

         HeaderPedido
         
         CreateTable(1200)

          response.write("<tr><th colspan=9>&nbsp;</td><th colspan=6 class='td-c subtitulo'>STOCK DMX </td><th colspan=6 class='td-c subtitulo'>STOCK DFW </td><td>&nbsp</td></tr>")
         response.write("<tr><th class='subtitulo td-c'>#</td><th class='subtitulo td-c'>Codigo</td><th class='subtitulo td-c'>Descripcion</td><th class='subtitulo td-c'>Estatus</td><th class='subtitulo td-c'>Pedida</td><th class='subtitulo td-c'>Pendiente</td><th class='subtitulo td-c'>Almacen </td><th class='subtitulo td-c'>Fecha embarque</td><th class='subtitulo td-c'>MXL</td><th class='subtitulo td-c'>SJR</td><th class='subtitulo td-c'>TRAS</td><th class='subtitulo td-c'>PR</td><th class='subtitulo td-c'>AD</td><th class='subtitulo td-c'>Total</td><th class='subtitulo td-c'>MXL</td><th class='subtitulo td-c'>SJR</td><th class='subtitulo td-c'>TRAS</td><th class='subtitulo td-c'>PR</td><th class='subtitulo td-c'>AD</td><th class='subtitulo td-c'>Total</td></tr>")
          
         n=1 
         while not rsUpdateEntry.EOF
             if n=1 then
                  vString=" style='background-color: white;' "
             else
                  vString=" style='background-color: #E1DFDF;' "
             end if
             response.write("<tr>")
             response.write("<td class='td-c' "& vString &">" & rsUpdateEntry("lineNum")+1 &"</td>")
             response.write("<td class='td-l fontmed' "& vString &">" & rsUpdateEntry("ItemCode") &"</td>")
             response.write("<td class='td-l fontmed' "& vString &">" & left(rsUpdateEntry("Dscription"),80) &"</td>")
             if trim(rsUpdateEntry("LineStatus"))="O" then
                  response.write("<td class='td-c fontmed' "& vString &"><font style='color:white;background-color:red'>abierta</font</td>")
             else
                  response.write("<td class='td-c fontmed' "& vString &">cerrada</td>")
             end if
             response.write("<td class='td-r' "& vString &">" & rsUpdateEntry("Quantity") &"</td>")

             if  int(trim(rsUpdateEntry("OpenQty")))>0 then
                 response.write("<td class='td-c' "& vString &"><font color='red'>" & rsUpdateEntry("OpenQty") &"</font></td>")
             else
                 response.write("<td class='td-c' "& vString &"><font color='green'>" & rsUpdateEntry("OpenQty") &"</font></td>")
             end if
             response.write("<td class='td-c fontmed' "& vString &">" & rsUpdateEntry("WhsCode") &"</td>")
             'response.write("<td class='td-c fontmed' "& vString &">" & rsUpdateEntry("Shipdate") &"</td>")

             if rsUpdateEntry("flag")=1 then
                  response.write("<td class='td-c fontmed alert'>" & rsUpdateEntry("U_FECHALMACEN") &"</td>")
             else
                  response.write("<td class='td-c fontmed' "& vString &">" & rsUpdateEntry("U_FECHALMACEN") &"</td>")
             end if

             if rsUpdateEntry("LineStatus")="O" then

             strSQL="select MXL=isnull((select CAST( SUM(InQty)-SUM(OutQty) as int) from OINM WITH (NOLOCK)  where itemcode='"& rsUpdateEntry("ItemCode") &"' and Warehouse='001-Mxl'),0),SJR=isnull((select CAST( SUM(InQty)-SUM(OutQty) as int) from OINM WITH (NOLOCK) where itemcode='"& rsUpdateEntry("ItemCode") &"' and Warehouse='002-SJR'),0),TRAS=isnull((select CAST( SUM(InQty)-SUM(OutQty) as int) from OINM WITH (NOLOCK) where itemcode='"& rsUpdateEntry("ItemCode") &"' and Warehouse='TR'),0),PR=isnull((select CAST( SUM(InQty)-SUM(OutQty) as int) from OINM WITH (NOLOCK) where itemcode='"& rsUpdateEntry("ItemCode") &"' and Warehouse='PR'),0),AD=isnull((select CAST( SUM(InQty)-SUM(OutQty) as int) from OINM WITH (NOLOCK) where itemcode='"& rsUpdateEntry("ItemCode") &"' and Warehouse='AD'),0),TOTAL=isnull((select CAST( SUM(InQty)-SUM(OutQty) as int) from OINM WITH (NOLOCK) where itemcode='"& rsUpdateEntry("ItemCode") &"' and Warehouse in ('001-Mxl','002-SJR','TR','PR','AD') ),0)"
             'response.write strSQL

             rsUpdateEntry2.Open strSQL, adoCon2 'DMX'
             response.write("<td class='td-r' "& vString &">" & rsUpdateEntry2("MXL") &"</td>")
             response.write("<td class='td-r' "& vString &">" & rsUpdateEntry2("SJR") &"</td>")
             response.write("<td class='td-r' "& vString &">" & rsUpdateEntry2("TRAS") &"</td>")
             response.write("<td class='td-r' "& vString &">" & rsUpdateEntry2("PR") &"</td>")
             response.write("<td class='td-r' "& vString &">" & rsUpdateEntry2("AD") &"</td>")
             response.write("<td class='td-r fontmed' "& vString &"><B>" & rsUpdateEntry2("Total") &"</B></td>")
             rsUpdateEntry2.close

             rsUpdateEntry2.Open strSQL, adoCon3 'DFW'
             response.write("<td class='td-r' "& vString &">" & rsUpdateEntry2("MXL") &"</td>")
             response.write("<td class='td-r' "& vString &">" & rsUpdateEntry2("SJR") &"</td>")
             response.write("<td class='td-r' "& vString &">" & rsUpdateEntry2("TRAS") &"</td>")
             response.write("<td class='td-r' "& vString &">" & rsUpdateEntry2("PR") &"</td>")
             response.write("<td class='td-r' "& vString &">" & rsUpdateEntry2("AD") &"</td>")
             response.write("<td class='td-r fontmed' "& vString &"><B>" & rsUpdateEntry2("Total") &"</B></td>")
             rsUpdateEntry2.close

             else
                response.write("<td class='td-r' "& vString &">-</td><td class='td-r' "& vString &">-</td><td class='td-r' "& vString &">-</td><td class='td-r' "& vString &">-</td><td class='td-r' "& vString &">-</td><td class='td-r' "& vString &">-</td><td class='td-r' "& vString &">-</td><td class='td-r' "& vString &">-</td><td class='td-r' "& vString &">-</td><td class='td-r' "& vString &">-</td><td class='td-r' "& vString &">-</td><td class='td-r' "& vString &">-</td>")

             end if

             response.write("</tr>")
             rsUpdateEntry.MoveNext
             n=n+1
             if n=3 then
                 n=1
             end if
         wend
         rsUpdateEntry.close
         closeTable

    end if

end sub


sub HeaderPedido
    strSQL="select c.SlpName as asesor,x.numAtCard,x.U_COMINTERNO,x.U_PESO_GLOBAL,x.Project,x.docnum,x.docdate,x.cardcode,x.cardname,x.Comments,(x.DocTotalSy-x.VatSumSy) as 'subtotal',x.DocTotalSy,x.VatSumSy as IVA,y.*,case when U_FECHALMACEN is not null THEN 1 else 0 end as 'flag','R:\"&request("fRS")&"\'+R.Ruta+'\'+convert(varchar,x.DocNum) as repositorio,x.CANCELED from ORDR x inner join RDR1 y on x.docentry=y.docentry left join OSLP c on x.SlpCode=c.slpcode left join SBOTemp.dbo.Repositorio R on R.RazonSocial='"& request("fRS") &"' and x.cardcode=R.CardCode where x.DocNum=" &request("fpedido") 

     if request("lineNum") <>"" then strSQL=strSQL&" and y.lineNum="&request("lineNum") 
                                             
         strSQL= strSQL &" order by y.lineNum"
         'response.write strSQL

         if request("fRS")="DMX" then
                 rsUpdateEntry.Open strSQL, adoCon2 'DMX'
         else
                 rsUpdateEntry.Open strSQL, adoCon3 'DFW'
         end if
      
         CreateTable(900)
         rowin
         response.write("<td class='td-c' width='50%'>")
           

        if not rsUpdateEntry.EOF then

         response.write("<table border=1 width='430px' cellpadding=3 cellspacing=2>")        
         
         response.write("<tr><td class='subtitulo td-r' width='100px'>Pedido:</td><td class='td-l'>"& request("fRS") &"-" &request("fPedido") &"</td></tr>")
         response.write("<tr><td class='subtitulo td-r'>Socio negocio:</td><td class='td-l'>"& rsUpdateEntry("cardname") &"</td></tr>")
         response.write("<tr><td class='subtitulo td-r'>Fecha:</td><td class='td-l'>"& rsUpdateEntry("DocDate") &"</td></tr>")
         response.write("<tr><td class='subtitulo td-r'>Referencia:</td><td class='td-l'>"& rsUpdateEntry("numAtCard") &"</td></tr>")
         response.write("<tr><td class='subtitulo td-r'>Proyecto:</td><td class='td-l'>"& rsUpdateEntry("Project") &"</td></tr>")
         response.write("<tr><td class='subtitulo td-r'>Peso Total:</td><td class='td-l'>"& rsUpdateEntry("U_PESO_GLOBAL") &" Kg</td></tr>")
          response.write("<tr><td class='subtitulo td-r'>Seguimiento:</td><td class='td-l'>"& rsUpdateEntry("U_COMINTERNO") &"</td></tr>")


          response.write("<tr><td class='subtitulo td-r'>Estatus:</td>")
          if rsUpdateEntry("CANCELED")="Y" then
              response.write("<td class='td-l' style='background-color:red;color:white'>PEDIDO CANCELADO</td></tr>")
          else
              response.write("<td class='td-l'>NO CANCELADO</td></tr>")
          end if

         response.write("</table></td><td class='td-c' width='50%'>")
         response.write("<table border=1 width='430px' cellpadding=3 cellspacing=2>")
               response.write("<tr><td class='subtitulo td-r'>Comentarios:</td><td class='td-l'>"& rsUpdateEntry("Comments") &"</td></tr>")
               response.write("<tr><td class='subtitulo td-r'>Subtotal:</td><td class='td-l'>"&  rsUpdateEntry("subTotal") &"</td></tr>")
               response.write("<tr><td class='subtitulo td-r'>IVA:</td><td class='td-l'>"& rsUpdateEntry("IVA") &"</td></tr>")
               response.write("<tr><td class='subtitulo td-r'>Total:</td><td class='td-l'>"& rsUpdateEntry("DocTotalSy") &"</td></tr>")
               response.write("<tr><td class='subtitulo td-r'>Asesor:</td><td class='td-l'>"& rsUpdateEntry("asesor") &"</td></tr>")
              
               response.write("<tr><td class='subtitulo td-r'>Repositorio:</td><td class='td-l'>"& rsUpdateEntry("repositorio") &"</td></tr>")
               'response.write("<tr><td class='subtitulo td-r'>Repositorio:</td><td class='td-l'><a href='file:"&rsUpdateEntry("repositorio")&"' target=repositorio>" & rsUpdateEntry("repositorio") &"</a></td></tr>")
         response.write("</table></td></tr> ")      

         else
               response.write("<tr><td colspan=8>no existen pedidos con esos criterios</td></tr>")

         end if

         response.write("</table>")
         response.write("<P>&nbsp;</P>")


end sub	



sub inventario  'contenido 69'

    if  not Session("session_modulo") ="Inventario" then
          response.redirect("ShowContent.asp?msg=vuelva a iniciar sesion")
    end if

   
    response.write("<P>&nbsp;</P>")
    if request("msg")<>"" then
           response.write("<P class='alert td-c'>"&  request("msg") &"</P>")
    end if


   response.write("<div id='stock'>&nbsp;</div>")
   titulo="Mostrar inventario y valor comercial<BR><font class=fontmed>Puede elegir uno o mas criterios</font>"
   DoTitulo(titulo)


         response.write("<form id='inventario' method='POST' action='ShowContent.asp'>")
         response.write("<input type='hidden' name='contenido' value=69>")
         response.write("<input type='hidden' name='action' value=2>")

      
         response.write("<table width='800px' cellpadding=4 cellspacing=3 align='center'><tr><td class='td-r'>Familia: </td><td class='td-l'><select name='ffamilia'>")
         strSQL="select * from OITB order by ItmsGrpNam"
         rsUpdateEntry.Open strSQL, adoCon2 'DMX'

         response.write("<option value='0'>sin seleccionar</option>")
         while not rsUpdateEntry.EOF
              if trim(request("ffamilia"))=trim(rsUpdateEntry("ItmsGrpNam")) then
                 response.write("<option value='"& rsUpdateEntry("ItmsGrpNam") &"' SELECTED>"&  rsUpdateEntry("ItmsGrpNam") &"</option>")
              else
                  response.write("<option value='"& rsUpdateEntry("ItmsGrpNam") &"'>"&  rsUpdateEntry("ItmsGrpNam") &"</option>")
              end if
              rsUpdateEntry.MoveNext
         wend
         rsUpdateEntry.close

         response.write("</select> &nbsp;&nbsp;")
         if request("enlistar")="on" then
             response.write("<input type='checkbox' name='enlistar' checked> muestre toda la familia aun sin inventario </td></tr>")
         else
             response.write("<input type='checkbox' name='enlistar'> muestre toda la familia aun sin inventario </td></tr>")
         end if

         response.write("<tr><td class='td-r'>Descripci&oacute;n comienza con palabra: </td><td class='td-l'> <input class='anchox2' type='text' name='fpalabra1' value='"& request("fpalabra1") &"'></td></tr>")
         response.write("<tr><td class='td-r'>Descripci&oacute;n incluye otra palabra: </td><td class='td-l'> <input class='anchox2' type='text' name='fpalabra2' value='"& request("fpalabra2") &"'></td></tr>")
         response.write("<tr><td class='td-r'>Descripci&oacute;n incluye otra palabra: </td><td class='td-l'> <input class='anchox2' type='text' name='fpalabra3' value='"& request("fpalabra3") &"'></td></tr>")

         response.write("<tr><td class='td-r'>Descripci&oacute;n <B>NO INCLUYE </B>palabra:  </td><td class='td-l'><input class='anchox2' type='text' name='fpalabra4' value='"& request("fpalabra4") &"'></td></tr>")
         response.write("<tr><td class='td-r'>Descripci&oacute;n <B>NO INCLUYE </B>palabra: </td><td class='td-l'><input class='anchox2' type='text' name='fpalabra5' value='"& request("fpalabra5") &"'></td></tr>")
         response.write("<tr><td class='td-r'>Descripci&oacute;n <B>NO INCLUYE </B>palabra: </td><td class='td-l'><input class='anchox2' type='text' name='fpalabra6' value='"& request("fpalabra6") &"'></td></tr>")

         response.write("<tr><td class='td-r'>C&oacute;digo de venta o de compra:</td><td class='td-l'><input class='anchox2' type='text' name='fcodigo' value='"& request("fcodigo") &"'></td></tr>")
          
          
         response.write("<tr><td class='td-c' colspan=2><br><input type='submit' value='Mostrar inventario'></form></td></tr></table>")


  



    if request("action")=2 then

          
          if request("ffamilia")="0" and request("fpalabra1")="" and request("fpalabra2")="" and request("fpalabra3")="" and request("fcodigo")="" then
                response.redirect("ShowContent.asp?contenido=69&msg=no indico ningun criterio de busqueda")
          end if

          if request("ffamilia")="0" and len(request("fpalabra1"))<3 and len(request("fpalabra2"))<3 and len(request("fpalabra3"))<3 and len(request("fcodigo"))<4 then
                response.redirect("ShowContent.asp?contenido=69&msg=el criterio de busqueda es muy ambiguo")
          end if


          '========================================CALCULE PRIMERO LOS TOTALES==================================================='
          strSQL="select '$ ' + CONVERT(VARCHAR,CONVERT(MONEY,sum(COSTO_VENTA)),1)  as calculo,sum(COSTO_VENTA) as calculo2 from _INV where 1=1 "

          if request("ffamilia")<>"0" then
                strSQL= strSQL &" AND ItmsGrpNam='"& request("ffamilia")  &"' "             
          end if     

          if len(request("fpalabra1"))>2 then
                strSQL= strSQL &" AND ItemName like '"& request("fpalabra1") &"%' "                
          end if 
          if len(request("fpalabra2"))>=1 then
                strSQL= strSQL &" AND ItemName like '%"& request("fpalabra2") &"%' "                
          end if 
          if len(request("fpalabra3"))>=1 then
                strSQL= strSQL &" AND ItemName like '%"& request("fpalabra3") &"%' "   
          end if 

          if len(request("fpalabra4"))>2 then
                strSQL= strSQL &" AND ItemName not like '%"& request("fpalabra4") &"%' "                
          end if 
           if len(request("fpalabra5"))>2 then
                strSQL= strSQL &" AND ItemName not like '%"& request("fpalabra5") &"%' "                
          end if 
           if len(request("fpalabra6"))>2 then
                strSQL= strSQL &" AND ItemName not like '%"& request("fpalabra6") &"%' "                
          end if 

        

         'response.write(strSQL &"<br>")
          rsUpdateEntry3.Open strSQL, adoCon4
          vCalculo2=1
          
          if int(trim(rsUpdateEntry3("calculo2")))>0 then
                 vCalculo2=int(trim(rsUpdateEntry3("calculo2")))
          end if
           '========================================EXTRAIGA LOS REGISTROS==================================================='
          
         
            if  request("enlistar")="on" then
               strSQL="select a.ItemCode as Codigo,a.itemname as Descripcion,a.ItmsGrpNam as Familia,a.Cod_Proveed,a.MXL,a.SJR,a.MTY,a.TRAS,a.VIRT,a.STOCK,a.PDTE as Comprometido,a.RESG,a.AD,TOTAL, '$ ' + CONVERT(VARCHAR,CONVERT(MONEY,a.COSTO_VENTA),1) as COSTO_VENTA, Pareto=round( CAST(a.COSTO_VENTA AS FLOAT)/CAST("& vCalculo2 &" AS FLOAT)*100,2) from _INV a  where 1=1 " 
            else
                 strSQL="select a.ItemCode as Codigo,a.itemname as Descripcion,a.ItmsGrpNam as Familia,a.Cod_Proveed,a.MXL,a.SJR,a.MTY,a.TRAS,a.VIRT,a.STOCK,a.PDTE as Comprometido,a.RESG,a.AD,TOTAL, '$ ' + CONVERT(VARCHAR,CONVERT(MONEY,a.COSTO_VENTA),1) as COSTO_VENTA, Pareto=round( CAST(a.COSTO_VENTA AS FLOAT)/CAST("& vCalculo2 &" AS FLOAT)*100,2) from _INV a where Stock!=0 " 

            end if
         


         

          if request("ffamilia")<>"0" then
                strSQL= strSQL &" AND a.ItmsGrpNam='"& request("ffamilia")  &"' "             
          end if 

        

          if len(request("fpalabra1"))>2 then
                strSQL= strSQL &" AND a.ItemName like '"& request("fpalabra1") &"%' "                
          end if 
          if len(request("fpalabra2"))>=1 then
                strSQL= strSQL &" AND a.ItemName like '%"& request("fpalabra2") &"%' "                
          end if 
          if len(request("fpalabra3"))>=1 then
                strSQL= strSQL &" AND a.ItemName like '%"& request("fpalabra3") &"%' "                
          end if 

          if len(request("fpalabra4"))>2 then
                strSQL= strSQL &" AND a.ItemName not like '%"& request("fpalabra4") &"%' "                
          end if 
          if len(request("fpalabra5"))>2 then
                strSQL= strSQL &" AND a.ItemName not like '%"& request("fpalabra5") &"%' "                
          end if 
          if len(request("fpalabra6"))>2 then
                strSQL= strSQL &" AND a.ItemName not like '%"& request("fpalabra6") &"%' "                
          end if 

          if len(request("fcodigo"))>2 and request("fcodigo")<>"sistemas" then    
                   strSQL= strSQL &" AND ( a.ItemCode in ('"& request("fcodigo") &"') OR  Cod_Proveed in ('"& request("fcodigo") &"')  ) "                      
          end if 

            if  request("enlistar")="on" then
                strSQL= strSQL &"   " 
            else
                strSQL= strSQL &" AND a.Stock!=0  " 
            end if

          strSQL= strSQL &" order by cast(a.COSTO_VENTA as int) desc"    
          'response.write(strSQL &"<br><br><br>")

         response.write("<table border=1 cellspacing=2 cellpadding=3 width='1200px' class='table2excel table2excel_with_colors' data-tableName='table1'>")        

         rsUpdateEntry.Open strSQL, adoCon4
         rsUpdateEntry2.Open strSQL, adoCon4
         
         Dim Campos(16) 
         i=1
         RowIn
         For Each rsUpdateEntry in rsUpdateEntry.Fields       
            Response.Write("<th class='subtitulo td-c fonttiny'>" & rsUpdateEntry.Name & "</th>")
            campos(i)=rsUpdateEntry.Name
            i=i+1           
         Next        
         RowOut

         if rsUpdateEntry2.EOF then
                 NoRegistros
         else         
          
            while not rsUpdateEntry2.EOF 
                response.write("<tr>")
                response.write("<td class='td-l fonttiny'>" & rsUpdateEntry2(campos(1)) &"</td>")
                response.write("<td class='td-l fonttiny'>" & mid(rsUpdateEntry2(campos(2)),1,56) &"<BR>"&mid(rsUpdateEntry2(campos(2)),57,100) &"</td>")
                response.write("<td class='td-l fonttiny'>" & rsUpdateEntry2(campos(3)) &"</td>")
                response.write("<td class='td-l fonttiny'>" & rsUpdateEntry2(campos(4)) &"</td>")
               
                               
                if int(trim(rsUpdateEntry2("MXL")))>0 then
                    response.write("<td class='td-r fontmed'>")      %>
                    <a href="Javascript:sendReq('DetalleStock.asp','codigo,almacen','<%=rsUpdateEntry2("Codigo") %>,001-Mxl','stock')">  <% response.write( rsUpdateEntry2("MXL") &"</a></td>")
                else
                    response.write("<td class='td-r fontmed'>" & rsUpdateEntry2("MXL") &"</td>")
                end if

                if int(trim(rsUpdateEntry2("SJR")))>0 then
                    response.write("<td class='td-r fontmed'>")      %>
                    <a href="Javascript:sendReq('DetalleStock.asp','codigo,almacen','<%=rsUpdateEntry2("Codigo") %>,002-SJR','stock')">  <% response.write( rsUpdateEntry2("SJR") &"</a></td>")
                else
                    response.write("<td class='td-r fontmed'>" & rsUpdateEntry2("SJR") &"</td>")
                end if

                if int(trim(rsUpdateEntry2("MTY")))>0 then
                    response.write("<td class='td-r fontmed'>")      %>
                    <a href="Javascript:sendReq('DetalleStock.asp','codigo,almacen','<%=rsUpdateEntry2("Codigo") %>,003-MTY','stock')">  <% response.write( rsUpdateEntry2("MTY") &"</a></td>")
                else
                    response.write("<td class='td-r fontmed'>" & rsUpdateEntry2("MTY") &"</td>")
                end if
                
                if int(trim(rsUpdateEntry2("TRAS")))>0 then
                    response.write("<td class='td-r fontmed'>")      %>
                    <a href="Javascript:sendReq('DetalleStock.asp','codigo,almacen','<%=rsUpdateEntry2("Codigo") %>,TR','stock')">  <% response.write( rsUpdateEntry2("TRAS") &"</a></td>")
                else
                    response.write("<td class='td-r fontmed'>" & rsUpdateEntry2("TRAS") &"</td>")
                end if

                if int(trim(rsUpdateEntry2("VIRT")))>0 then
                    response.write("<td class='td-r fontmed'>" & rsUpdateEntry2("VIRT") &"</td>")
                else
                    response.write("<td class='td-r fontmed'>" & rsUpdateEntry2("VIRT") &"</td>")
                end if

                if int(trim(rsUpdateEntry2("Stock")))<0 then
                    response.write("<td class='td-r fontmed alert'>" & rsUpdateEntry2("Stock") &"</td>")
                else
                    response.write("<td class='td-r fontmed'>" & rsUpdateEntry2("Stock") &"</td>")
                end if


                if int(trim(rsUpdateEntry2("Comprometido")))<>0 then
                    response.write("<td class='td-r fontmed alert'>" & rsUpdateEntry2("Comprometido") &"</td>")
                else
                    response.write("<td class='td-r fontmed'>" & rsUpdateEntry2("Comprometido") &"</td>")
                end if

                response.write("<td class='td-r fontmed'>" & rsUpdateEntry2("RESG") &"</td>")

                 if int(trim(rsUpdateEntry2("AD")))>0 then
                    response.write("<td class='td-r fontmed'>")      %>
                    <a href="Javascript:sendReq('DetalleStock.asp','codigo,almacen','<%=rsUpdateEntry2("Codigo") %>,AD','stock')">  <% response.write( rsUpdateEntry2("AD") &"</a></td>")
                else
                    response.write("<td class='td-r fontmed'>" & rsUpdateEntry2("AD") &"</td>")
                end if

                if int(trim(rsUpdateEntry2("TOTAL")))<0 then
                    response.write("<td class='td-r fontmed alert'>" & rsUpdateEntry2("TOTAL") &"</td>")
                else
                    response.write("<td class='td-r fontmed'>" & rsUpdateEntry2("TOTAL") &"</td>")
                end if

                
            
               
                response.write("<td class='td-r fonttiny'>" & rsUpdateEntry2("COSTO_VENTA") &"</td>")
                response.write("<td class='td-r fontmed'>" & rsUpdateEntry2("Pareto") &"</td>")
                response.write("</tr>")
                rsUpdateEntry2.MoveNext
            wend
            rsUpdateEntry2.close

           response.write("<tr><td class='td-r subtitulo' colspan=9>TOTAL</td><td class='td-r fontsmall' width='100px'><B>"& rsUpdateEntry3("calculo") &"</B></td><td class='td-r fontmed' width='100px' colspan=6>&nbsp;</td></tr>")
            rsUpdateEntry3.close

         end if
         %>
         </table><button class="exportToExcel">Exportar a excel</button>  &nbsp;&nbsp;&nbsp; <button onclick="go_display_div('#titulo')"> ir arriba   </button>     

         
    <script>
      $(function() {
        $(".exportToExcel").click(function(e){
          var table = $(this).prev('.table2excel');
          if(table && table.length){
            var preserveColors = (table.hasClass('table2excel_with_colors') ? true : false);
            $(table).table2excel({
              exclude: ".noExl",
              name: "Excel Document Name",
              filename: "myFileName" + new Date().toISOString().replace(/[\-\:\.]/g, "") + ".xls",
              fileext: ".xls",
              exclude_img: true,
              exclude_links: true,
              exclude_inputs: true,
              preserveColors: preserveColors
            });
          }
        });
        
      });
    </script>

    <%

    end if



end sub

 



sub  PartidasAbiertas1 'ventas   contenido 70'   

   strSQL="SELECT 'DMX' as RS,T0.CardName as 'SOCIO NEGOCIO',sum( round(T1.OpenQty*T1.Price,2) ) as SUBTOTAL,CONVERT(VARCHAR,CONVERT(MONEY,  sum(round(T1.OpenQty*T1.Price,2))   ),1) as subtotal2 FROM [PRODUCTIVA_DMX].dbo.ORDR T0 WITH (NOLOCK) inner join [PRODUCTIVA_DMX].dbo.RDR1 T1 WITH (NOLOCK) on T0.DocEntry=T1.DocEntry WHERE T0.CANCELED='N' and T0.DocStatus='O' and T1.LineStatus='O'  GROUP BY T0.CardName UNION SELECT 'DFW' as RS,T0.CardName as 'SOCIO NEGOCIO',sum(round(T1.OpenQty*T1.Price,2)) as SUBTOTAL, CONVERT(VARCHAR,CONVERT(MONEY, cast( sum(round(T1.OpenQty*T1.Price,2)) as int)  ),1) as subtotal2 FROM [PRODUCTIVA_DFW].dbo.ORDR T0 WITH (NOLOCK) inner join [PRODUCTIVA_DFW].dbo.RDR1 T1 WITH (NOLOCK) on T0.DocEntry=T1.DocEntry WHERE T0.CANCELED='N' and T0.DocStatus='O' and T1.LineStatus='O'  GROUP BY T0.CardName order by subtotal desc"

  'response.write strSQL

   strSQL2="SELECT CONVERT(VARCHAR,CONVERT(MONEY, (SELECT ISNULL( SUM( round(T1.OpenQty*T1.Price,2) ),0) FROM [PRODUCTIVA_DMX].dbo.ORDR T0 WITH (NOLOCK) inner join [PRODUCTIVA_DMX].dbo.RDR1 T1 WITH (NOLOCK) on T0.DocEntry=T1.DocEntry WHERE T0.CANCELED='N' and T0.DocStatus='O' and T1.LineStatus='O' ) +   (SELECT ISNULL( SUM( round(T1.OpenQty*T1.Price,2) ),0)  FROM [PRODUCTIVA_DFW].dbo.ORDR T0 WITH (NOLOCK) inner join [PRODUCTIVA_DFW].dbo.RDR1 T1 WITH (NOLOCK) on T0.DocEntry=T1.DocEntry WHERE T0.CANCELED='N' and T0.DocStatus='O' and T1.LineStatus='O' )   ),1)   AS CALCULO"

   'response.write strSQL

   titulo="PARTIDAS ABIERTAS PEDIDOS DE VENTA"
   DoTitulo(titulo)
   
 
   response.write("<center><table border=1 cellspacing=2 cellpadding=3 align=center width='640px' class='table2excel table2excel_with_colors' data-tableName='table1'>")

   rsUpdateEntry.Open strSQL, adoCon4
   
   rsUpdateEntry3.Open strSQL2, adoCon4

   response.write("<tr><td class='subtitulo'>Razon Social</td>")
   response.write("<td class='subtitulo'>Socio de negocio</td>")
   response.write("<td class='subtitulo'>subtotal</td></tr>")
  

    if rsUpdateEntry.EOF then
           NoRegistros
    else         
           while not rsUpdateEntry.EOF
                response.write("<tr><td class='td-c fontmed'>"& rsUpdateEntry("RS") &"</td>")
                response.write("<td class='td-r fontmed'>"& rsUpdateEntry("SOCIO NEGOCIO") &"</td>")
                response.write("<td class='td-r fontmed'>"& rsUpdateEntry("subtotal2") &"</td></tr>")
                rsUpdateEntry.MoveNext
           wend
           rsUpdateEntry.close
           response.write("<tr><td colspan=2 class='td-r subtitulo'>GRAN TOTAL</td><td class='td-r fontmed strong'>" & rsUpdateEntry3("calculo") &"</td></tr>")
           rsUpdateEntry3.close           
           %>
           </table><button class="exportToExcel">Exportar a excel</button> </center>
      <%

    end if 
 
   strSQL="if exists (select * from INFORMATION_SCHEMA.TABLES where TABLE_NAME = 'TEMP_VENTAS' AND TABLE_SCHEMA = 'dbo')     drop table dbo.TEMP_VENTAS;"
   rsUpdateEntry3.Open strSQL, adoCon4

   strSQL="SELECT 'DMX' as RS,T0.CardName as 'SOCIO_NEGOCIO',T0.DocNum as Pedido,sum( round(T1.OpenQty*T1.Price,2) ) as SUBTOTAL INTO TEMP_VENTAS FROM [PRODUCTIVA_DMX].dbo.ORDR T0 inner join [PRODUCTIVA_DMX].dbo.RDR1 T1 on T0.DocEntry=T1.DocEntry WHERE T0.CANCELED='N' and T0.DocStatus='O' and T1.LineStatus='O' GROUP BY T0.CardName,T0.DocNum"
   rsUpdateEntry3.Open strSQL, adoCon4

   strSQL="INSERT INTO TEMP_VENTAS   SELECT 'DFW' as RS,T0.CardName as 'SOCIO_NEGOCIO',T0.DocNum as Pedido,sum( round(T1.OpenQty*T1.Price,2) ) as SUBTOTAL FROM [PRODUCTIVA_DFW].dbo.ORDR T0 inner join [PRODUCTIVA_DFW].dbo.RDR1 T1 on T0.DocEntry=T1.DocEntry WHERE T0.CANCELED='N' and T0.DocStatus='O' and T1.LineStatus='O' GROUP BY T0.CardName,T0.DocNum"
   rsUpdateEntry3.Open strSQL, adoCon4

   vSocio=""
   strSQL="SELECT RS,SOCIO_NEGOCIO,PEDIDO,CONVERT(VARCHAR,CONVERT(MONEY, subtotal   ),1)  AS SUBTOTAL FROM TEMP_VENTAS ORDER BY RS,SOCIO_NEGOCIO,PEDIDO"
   rsUpdateEntry3.Open strSQL, adoCon4
   rsUpdateEntry4.Open strSQL, adoCon4

   response.write("<P>&nbsp;</P><H1>DETALLE</H1><center><table border=1 cellspacing=2 cellpadding=3 align=center width='640px' class='table2excel table2excel_with_colors' data-tableName='table1'>")

   CamposRS3
   while not rsUpdateEntry4.EOF
                response.write("<tr><td class='td-c fontmed'>"& rsUpdateEntry4("RS") &"</td>")
                response.write("<td class='td-r fontmed'>"& rsUpdateEntry4("SOCIO_NEGOCIO") &"</td>")
                 response.write("<td class='td-r fontmed'>"& rsUpdateEntry4("pedido") &"</td>")
                response.write("<td class='td-r fontmed'>"& rsUpdateEntry4("subtotal") &"</td></tr>")
                rsUpdateEntry4.MoveNext
   wend
   rsUpdateEntry4.close

   %>
   </table><button class="exportToExcel">Exportar a excel</button> </center>

   <script>
              $(function() {
                $(".exportToExcel").click(function(e){
                  var table = $(this).prev('.table2excel');
                  if(table && table.length){
                    var preserveColors = (table.hasClass('table2excel_with_colors') ? true : false);
                    $(table).table2excel({
                      exclude: ".noExl",
                      name: "Excel Document Name",
                      filename: "myFileName" + new Date().toISOString().replace(/[\-\:\.]/g, "") + ".xls",
                      fileext: ".xls",
                      exclude_img: true,
                      exclude_links: true,
                      exclude_inputs: true,
                      preserveColors: preserveColors
                    });
                  }
                });
                
              });
    </script>

   <%

   strSQL="if exists (select * from INFORMATION_SCHEMA.TABLES where TABLE_NAME = 'TEMP_VENTAS' AND TABLE_SCHEMA = 'dbo')     drop table dbo.TEMP_VENTAS;"
   rsUpdateEntry4.Open strSQL, adoCon4


end sub



sub InformeEnvios   'contenido 71'

    response.write("<P>&nbsp;</P>")
    if request("msg")<>"" then
           response.write("<P class='alert td-c'>"&  request("msg") &"</P>")
    end if

    response.write("<center><H1><B>EXTRAE LA INFORMACION DE ENVIOS</B> </H1><P>&nbsp;</P>")  


    %>

     <form id="cobranza" method="POST" action="ShowContent.asp"> 
     <input type="hidden" name="contenido" value="71">   <%

     response.write("A&ntilde;o: <select name='fanio'>")

      
           response.write("<option value="& year(date())-1 &">"& year(date())-1 &"</option>")
           response.write("<option value="& year(date()) &" selected>"& year(date()) &"</option>")
           response.write("</select><br><br>")   %>

     
     <br><br>
     <input type="submit" value="Generar informe"></form>

     <%

     if request("fanio")<>"" then    
         strSQL="set dateformat ymd;select a.ID,dbo.fn_GetMonthName(a.DocDate,'spanish') as Fecha,a.RazonSocial as 'RS',a.Cardname,a.DocNum as Remision,a.pedido,CASE when a.pedido is null THEN '' else case when c.DocStatus='C' THEN 'cerrado' else 'abierto' END END as Estatus,a.Guia, CONVERT(VARCHAR,CONVERT(MONEY,a.subTotal),1)  as Costo,CONVERT(VARCHAR,CONVERT(MONEY, e.subTotal ),1)   as 'subtotal <br> remision',a.ocurre,a.comentarios,b.Portador,case when a.CargoCliente=1 then 'si' else 'no' end as 'Cargo <br> cliente' from envios a inner join cat_portadores b on a.id_portador=b.id_portador left join Notificaciones c on  c.modulo='ventas' and c.tipo='pedido' and a.pedido=c.docnum left join PRODUCTIVA_DMX.dbo.ORTT d on c.DocDate=d.RateDate left join Notificaciones e on e.tipo='remision' and a.RazonSocial=e.RazonSocial and a.Docnum=e.DocNum where year(a.docdate)="& request("fanio") & " order by a.Pedido"         
         'response.write strSQL
         
          response.write("<table width='1200px' cellpadding=5 cellspacing=2 border=1 class='table2excel table2excel_with_colors' data-tableName='table1'>")


          rsUpdateEntry.Open strSQL, adoCon4  
          rsUpdateEntry2.Open strSQL, adoCon4

         Dim Campos(20) 
         i=1
        
         For Each rsUpdateEntry in rsUpdateEntry.Fields                        
            campos(i)=rsUpdateEntry.Name
            i=i+1           
         Next   
         i=i-1 

         RowIn  
         for n=1 to i  
            Response.Write("<th class='subtitulo td-c'>" & campos(n) & "</th>")
         Next
         RowOut

         while not rsUpdateEntry2.EOF
             rowin
             Response.Write("<td class='fonttiny td-r' width='40px'>" & rsUpdateEntry2(campos(1)) & "</td>")
             Response.Write("<td class='fonttiny td-r' width='90px'>" & rsUpdateEntry2(campos(2)) & "</td>")
             Response.Write("<td class='fonttiny td-r' width='40px'>" & rsUpdateEntry2(campos(3)) & "</td>")
             Response.Write("<td class='fonttiny td-r' width='200px'>" & rsUpdateEntry2(campos(4)) & "</td>")
             Response.Write("<td class='fonttiny td-r' width='65px'>" & rsUpdateEntry2(campos(5)) & "</td>")
             Response.Write("<td class='fonttiny td-r' width='65px'>" & rsUpdateEntry2(campos(6)) & "</td>")
             Response.Write("<td class='fonttiny td-r' width='65px'>" & rsUpdateEntry2(campos(7)) & "</td>")
             for n=8 to 11  
                  Response.Write("<td class='fonttiny td-r'>" & rsUpdateEntry2(campos(n)) & "</td>")
             Next
             Response.Write("<td class='fonttiny td-r' width='120px'>" & rsUpdateEntry2(campos(12)) & "</td>")
             Response.Write("<td class='fonttiny td-r' width='80px'>" & rsUpdateEntry2(campos(13)) & "</td>")
             Response.Write("<td class='fonttiny td-r' width='50px'>" & rsUpdateEntry2(campos(14)) & "</td>")
             rsUpdateEntry2.MoveNext
             rowout
         wend
         rsUpdateEntry2.close


         %>
         </table><button class="exportToExcel">Exportar a excel</button> </center>

   <script>
              $(function() {
                $(".exportToExcel").click(function(e){
                  var table = $(this).prev('.table2excel');
                  if(table && table.length){
                    var preserveColors = (table.hasClass('table2excel_with_colors') ? true : false);
                    $(table).table2excel({
                      exclude: ".noExl",
                      name: "Excel Document Name",
                      filename: "myFileName" + new Date().toISOString().replace(/[\-\:\.]/g, "") + ".xls",
                      fileext: ".xls",
                      exclude_img: true,
                      exclude_links: true,
                      exclude_inputs: true,
                      preserveColors: preserveColors
                    });
                  }
                });
                
              });
    </script>

   <%
  
         response.write("<P>&nbsp;</P>")

     end if

end sub





sub Contactos   'contenido 72'
 
 if request("action")=4 then
     titulo="Registrar contacto no. "& request("ID") &"<br>se convirtio en Socio de negocio"
     DoAlert
     DoTitulo(titulo)
     'response.write request("ID")
     response.write("<form id='lead' method='POST' action='ShowContent.asp'>")
     response.write("<input type='hidden' name='contenido' value='72'>")
     response.write("<input type='hidden' name='action' value=5>")
     response.write("<input type='hidden' name='ID' value='"& request("ID") &"'>")

     strSQL="select *,.dbo.fn_GetMonthName(fecha_solicitud,'spanish') as fecha from contactos where id="&request("ID")
     rsUpdateEntry.Open strSQL, adoCon

     CreateTable(600)
     response.write("<tr><td width='40%' class='td-r subtitulo'>Empresa:</td><td class='td-l' width='60%'>" & rsUpdateEntry("empresa") &"</td></tr>")
     response.write("<tr><td class='td-r subtitulo'>Contacto:</td><td class='td-l'>" & rsUpdateEntry("contacto") &"</td></tr>")
     response.write("<tr><td class='td-r subtitulo'>Fecha:</td><td class='td-l'>" & rsUpdateEntry("fecha") &"</td></tr>")
     response.write("<tr><td class='td-r subtitulo'>Socio de negocio en DMX:</td><td class='td-l'><select name='fSocioNegocio'><option value='0'>sin seleccionar</option>")
   
     strSQL="select cardcode,cardname from PRODUCTIVA_DMX.dbo.OCRD where substring(cardcode,1,1)='C' and CardName not like '%deltaflow%' order by CardCode desc" 
     rsUpdateEntry2.Open strSQL, adoCon2

     while not rsUpdateEntry2.EOF
      response.write("<option value='"&rsUpdateEntry2("cardCode")&"'>"& rsUpdateEntry2("cardCode") &"-" & mid(rsUpdateEntry2("cardname"),1,26) &"</option>")
      rsUpdateEntry2.MoveNext
     wend
     rsUpdateEntry2.close
     response.write("</select> </td></tr>")
     response.write("</table><br><input type='submit' value='registrar'></form>")
     response.write("<P>&nbsp;</P>")


end if


if request("action")=5 then
    'response.write(request("ID") &"<br>")
    'response.write(request("fSocioNegocio") &"<br>")
    strSQL="update contactos set cardCode='"& request("fSocioNegocio") &"' where id="&request("ID")
    response.write strSQL
    rsUpdateEntry.Open strSQL, adoCon

    response.redirect("ShowContent.asp?contenido=72&msg=se actualizo la informacion de contactos")
end if   


 if request("action")="" then

    if trim(request("duplicar"))<>"" then
          strSQL="select * from contactos where id=" & trim(request("duplicar"))
          rsUpdateEntry.Open strSQL, adoCon

          vcontacto= rsUpdateEntry("contacto")
          vEmpresa= rsUpdateEntry("empresa")
          vcelular= rsUpdateEntry("celular")
          vGiro= rsUpdateEntry("giro")
          vEmail= rsUpdateEntry("email_contacto")
          vTelEmpresa= rsUpdateEntry("telefono_empresa")
          vPortal= rsUpdateEntry("PORTAL_EMPRESA")
          vCiudad= rsUpdateEntry("ciudad")

          rsUpdateEntry.close

    end if

    DoAlert
    titulo="CONTACTOS DE MARKETING"
    DoTitulo(titulo)    %>
  
    <!--
    response.write("<form id='form1' method='POST' action='ShowContent.asp'>")
    response.write("<input type='hidden' name='contenido' value='72'>")
    response.write("<input type='hidden' name='action' value=1>")

    CreateTable(1200)
    response.write("<tr><td style='vertical-align:top; text-align:center' width='1000px'>")
    CreateTable(1000)
    
    'response.write("<tr><td colspan=2 class='subtitulo td-c'>CONTACTO</td><td colspan=2 class='subtitulo td-c'>EMPRESA</td></tr>")    

    response.write("<tr>")
      response.write("<td class='subtitulo mix td-r' width='110px'>Contacto:</td>")
      response.write("<td class='mix td-l'><input type='text' name='fcontacto' class='anchox3' value='"& vcontacto &"'></td>")
      response.write("<td class='subtitulo mix td-r' width='110px'>Empresa:</td>")
      response.write("<td class='mix td-l'><input type='text' name='fempresa' class='anchox2' value='"& vempresa &"'></td>")
    response.write("</tr>")

    response.write("<tr>")
      response.write("<td class='subtitulo mix td-r'>Celular: </td>")

      if vcelular<>"" then
         response.write("<td class='mix td-l'><input type='text' name='fcelular' class='anchox1' maxlength=10  value='"& vcelular &"' ></td>")
      else
         response.write("<td class='mix td-l'><input type='text' name='fcelular' class='anchox1' maxlength=10  placeholder='10 digitos'></td>")
      end if
      response.write("<td class='subtitulo mix td-r'>Giro:</td>")
      response.write("<td class='mix td-l'><input type='text' name='fgiro' class='anchox2' value='"& vGiro &"' ></td>")
    response.write("</tr>")

    response.write("<tr>")
      response.write("<td class='subtitulo mix td-r'>eMail: </td>")
      response.write("<td class='mix td-l'><input type='text' name='femail' class='anchox3' value='"& vEmail &"'></td>")
      response.write("<td class='subtitulo mix td-r'>Telefono:</td>")

      if request("fTelEmpresa")<>"" then
         response.write("<td class='mix td-l'><input type='text' name='fClavePais' maxlength=5  value='+52' style='width:40px'>&nbsp;&nbsp;&nbsp;<input type='text' name='fTelEmpresa' class='anchox2' maxlength=10 value='"& vTelEmpresa &"' ></td>")
      else       
         response.write("<td class='mix td-l'> <input type='text' name='fClavePais' maxlength=5  value='+52' style='width:40px'>&nbsp;&nbsp;&nbsp; <input type='text' name='fTelEmpresa' class='anchox2' maxlength=10 placeholder='10 digitos'></td>")
      end if
    response.write("</tr>")

    response.write("<tr>")
      response.write("<td class='subtitulo mix td-r'>Fecha solicitud: </td>")
      response.write("<td class='mix td-l'><input type='date' name='ffecha'></td>")
      response.write("<td class='subtitulo mix td-r'>Portal Web:</td>")
      response.write("<td class='mix td-l'><input type='text' name='fportal' class='anchox2' value='"& vPortal &"'></td>")
    response.write("</tr>")

    response.write("<tr><td class='subtitulo mix td-r'>Asesor: </td><td class='mix td-l'>")

      response.write("<select name='fasesor' class='anchox2'>")

    response.write("<option value=0>por asignar</option> ")

             strSQL="select * from OSLP where Active='Y' and U_selector is not null "
             'response.write  strSQL            
             rsUpdateEntry.Open strSQL, adoCon2  'DMX'

        while not rsUpdateEntry.EOF
          if rsUpdateEntry("slpcode")=13 then
              response.write("<option value='"& rsUpdateEntry("slpcode") &"' SELECTED>"& rsUpdateEntry("SlpName") &"</option> ")
          else
             response.write("<option value='"& rsUpdateEntry("slpcode") &"'>"& rsUpdateEntry("SlpName") &"</option> ")
          end if
          rsUpdateEntry.movenext
        wend
        rsUpdateEntry.close
        response.write("</select></td>")

    
      response.write("<td class='subtitulo mix td-r'>Ent. federativa:</td>")
      response.write("<td class='mix td-l'><select name='fentidad' class='anchox2'>")
         

        strSQL="select * from cat_entidades order by id_entidad"
        rsUpdateEntry.Open strSQL, adoCon

        while not rsUpdateEntry.EOF
          response.write("<option value='"& rsUpdateEntry("id_entidad") &"'>"& rsUpdateEntry("entidad") &"</option> ")
          rsUpdateEntry.movenext
        wend
        rsUpdateEntry.close
        response.write("</select></td>")
    response.write("</tr>")

     response.write("<tr>")
      response.write("<td class='subtitulo mix td-r'>adjuntar?</td>")
      response.write("<td class='mix td-l'><input type='checkbox' name='ffile'> </td>")
      response.write("<td class='subtitulo mix td-r'>Ciudad:</td>")
      response.write("<td class='mix td-l'><input type='text' name='fciudad' class='anchox2' value='"& vCiudad &"'></td>")
    response.write("</tr>")


    response.write("<tr>")
      response.write("<td class='subtitulo mix td-r'>Solicitud: </td>")
      response.write("<td class='mix td-l' colspan=3><input type='text' name='fsolicitud' style='width:800px'></td>")     
    response.write("</tr>")
   
    %>
     <tr><td class='mix td-l' colspan=4>
           <DIV id="cuerpo">
                  <a href="Javascript:sendReq('cuerpo.htm','var','1','cuerpo')">   &nbsp;&nbsp;&nbsp;<U><font color=red>introducir cuerpo de email</font></U></a>
           </DIV>
    </td></tr>
    
    '<font color=red><img src='/img/animated-arrow.gif'> METODO DE ENTRADA </font>&nbsp;&nbsp;&nbsp; <select name='fmode'> ")
           
        'response.write("<option value=0>por asignar</option> ")

             'strSQL="select * from cat_mode order by id_mode"
             'rsUpdateEntry.Open strSQL, adoCon

        'while not rsUpdateEntry.EOF
          'if trim(rsUpdateEntry("id_mode"))="4" then
              'response.write("<option value='"& rsUpdateEntry("id_mode") &"' SELECTED>"& rsUpdateEntry("mode") &"</option> ")
          'else
              'response.write("<option value='"& rsUpdateEntry("id_mode") &"'>"& rsUpdateEntry("mode") &"</option> ")
          'end if
          'rsUpdateEntry.movenext
        'wend
        'rsUpdateEntry.close
        'response.write("</select>&nbsp;&nbsp;<input type='submit' value='registrar'></form></td>")
    -->
    <%
        CreateTable(800)
        rowin
        response.write("<td class='td-c' width='650px'><form id='contactos' action='ShowContent.asp' method='POST'>")
        response.write("Buscar:  &nbsp;&nbsp;&nbsp;<input type='text' name='fpalabra' value='"& request("fpalabra") &"'>&nbsp;&nbsp;&nbsp;<input type='submit' value='buscar'>")

	    response.write("<input type='hidden' value="& request("contenido") &" name='contenido'>")
	    response.write("<input type='hidden' value="& request("mode")&" name='mode'>")
	    response.write("</form></td><td width='150px' style='vertical-align:top; text-align:right'><b>APLICAR FILTRO:</b><BR>")
	  


    strSQL="select distinct(a.asesor),b.SlpName from contactos a inner join PRODUCTIVA_DMX.dbo.OSLP b on a.asesor=b.SlpCode"
    rsUpdateEntry2.Open strSQL, adoCon

    while not  rsUpdateEntry2.EOF
          response.write( "<a href='ShowContent.asp?contenido="& request("contenido") &"&asesor=" & rsUpdateEntry2("asesor") &"'>" & rsUpdateEntry2("SlpName") &"</a><br>" )
          rsUpdateEntry2.MoveNext
    wend
    response.write( "<a href='ShowContent.asp?contenido="& request("contenido") &"&asesor=none'><b>remover filtro</b></a><br>" )
    rsUpdateEntry2.close
    response.write("</td></tr>")
    closeTable
    separador
    ShowContacts
   

  end if
   

    if request("action")=1 then

        vAsesor=0

        if request("fmode")=0 then  'no eligio un metodo de entrada'
             response.redirect("ShowContent.asp?contenido=72&msg=no selecciono un metodo de entrada")
        end if

        if request("fasesor")="" then  'se eligira al asesor de forma automatica'
             strSQL="select TOP 1 * from OSLP where Active='Y' and slpcode>0 and ( memo like '%ventas%' OR SlpName like '%Héctor%' OR SlpName like '%julio%' ) order by U_selector, SlpName "
             rsUpdateEntry.Open strSQL, adoCon2

             vAsesor=rsUpdateEntry("SlpCode")
             rsUpdateEntry.close 
        else
             vAsesor=int(trim(request("fasesor")))
        end if

        'ya que sabemos que asesor es, subire el contador pues tomo una solicitud'
        strSQL="UPDATE OSLP set U_selector=U_selector+1 where SlpCode=" & vAsesor
        rsUpdateEntry5.Open strSQL, adoCon2  

   
        vfecha=""
        vfecha=trim(year(date())) &"-" & trim(month(date())) &"-" &trim(day(date())) 
        if request("ffecha") <>"" then
             vfecha=trim(request("ffecha"))
        end if

        if request("fcuerpo") <>"" then
             vCuerpo=mid(trim(request("fcuerpo")),1,8000)
             vcuerpo=replace(vCuerpo,"'","''")
        end if

        vSolicitud=trim(request("fsolicitud"))
        vSolicitud=replace(vSolicitud,"'","''")


        strSQL="set dateformat ymd;insert into contactos (ID,ID_MODE,EMPRESA,GIRO,PORTAL_EMPRESA,TELEFONO_EMPRESA,ID_ENTIDAD,CONTACTO,CELULAR,EMAIL_CONTACTO,SOLICITUD,ASESOR,FECHA_SOLICITUD,LOGDATE,ID_USUARIO,CUERPO_EMAIL,CIUDAD,ID_ESTATUS,CodigoInternacional) values ((select max(ID)+1 from contactos),"& request("fmode") &",'"& request("fempresa") &"','"& request("fgiro") &"','"& request("fportal") &"','"& request("fTelEmpresa") &"','"& request("fentidad") &"','"& request("fcontacto") &"','"& request("fcelular") &"','"& request("femail") &"','"& request("fsolicitud") &"',"& vAsesor &",'"& vFecha &"',getdate(),'"& session("session_id") &"','"& vCuerpo &"','"& request("fciudad") &"','01','"& request("fClavePais") &"')"

        response.write (   strSQL  &"<br>" )
        rsUpdateEntry3.Open strSQL, adoCon

        strSQL="select top 1 * from contactos order by LOGDATE desc"
        rsUpdateEntry3.Open strSQL, adoCon

        Session("ID")=rsUpdateEntry3("ID")
        rsUpdateEntry3.close

        if request("ffile")="on" or request("ffile")="true" then
             response.redirect("/upload/uploadform3.asp?ID="& Session("ID") &"&msg=el contacto fue agregado, adjunte el archivo a la solicitud")
        else 
             response.redirect("ShowContent.asp?contenido=72&msg=el contacto fue agregado")
        end if

    end if




    if request("action")=2 then
         titulo="EDITAR UN CONTACTO DE MARKETING No. "& REQUEST("ID")
         DoTitulo(titulo)
       

    response.write("<form id='form1' method='POST' action='ShowContent.asp'>")
        response.write("<input type='hidden' name='contenido' value='72'>")
        response.write("<input type='hidden' name='action' value=3>")
        response.write("<input type='hidden' name='id' value="& REQUEST("ID") &">")


    StrSQL="SELECT a.*,.dbo.fn_GetMonthName(a.fecha_solicitud,'spanish') as FECHA FROM CONTACTOS a left join PRODUCTIVA_DMX.dbo.OCRD b on a.cardCode collate Modern_Spanish_CI_AS =b.cardcode collate Modern_Spanish_CI_AS WHERE ID=" & REQUEST("ID")
    rsUpdateEntry4.Open strSQL, adoCon

    CreateTable(960)
    
    response.write("<tr><td colspan=2 class='subtitulo td-c'>CONTACTO</td><td colspan=2 class='subtitulo td-c'>EMPRESA</td></tr>")
    response.write("<tr><td colspan=4 class='td-c'><font color=red>METODO DE ENTRADA </font>&nbsp;&nbsp;&nbsp; <select name='fmode'> ")
           
        response.write("<option value=0>por asignar</option> ")

             strSQL="select * from cat_mode order by id_mode"
             rsUpdateEntry.Open strSQL, adoCon

        while not rsUpdateEntry.EOF
          if rsUpdateEntry4("ID_MODE")=rsUpdateEntry("id_mode") then
              response.write("<option value='"& rsUpdateEntry("id_mode") &"' SELECTED>"& rsUpdateEntry("mode") &"</option> ")
          else
              response.write("<option value='"& rsUpdateEntry("id_mode") &"'>"& rsUpdateEntry("mode") &"</option> ")
          end if
          rsUpdateEntry.movenext
        wend
        rsUpdateEntry.close
        response.write("</select></td></tr>")

    response.write("<tr>")
      response.write("<td class='subtitulo mix td-r' width='80px'>Contacto:</td>")
      response.write("<td class='mix td-l'><input type='text' name='fcontacto' class='anchox3' value='"& rsUpdateEntry4("contacto") &"'></td>")
      response.write("<td class='subtitulo mix td-r' width='90px'>Empresa:</td>")
      response.write("<td class='mix td-l'><input type='text' name='fempresa' class='anchox3' value='"& rsUpdateEntry4("empresa") &"'></td>")
    response.write("</tr>")

    response.write("<tr>")
      response.write("<td class='subtitulo mix td-r'>Celular: </td>")
      response.write("<td class='mix td-l'><input type='text' name='fcelular' class='anchox1' value='"& rsUpdateEntry4("celular") &"' maxlength=10></td>")
      response.write("<td class='subtitulo mix td-r'>Giro:</td>")
      response.write("<td class='mix td-l'><input type='text' name='fgiro' class='anchox2' value='"& rsUpdateEntry4("giro") &"'></td>")
    response.write("</tr>")

    response.write("<tr>")
      response.write("<td class='subtitulo mix td-r'>eMail: </td>")
      response.write("<td class='mix td-l'><input type='text' name='femail' class='anchox3' value='"& rsUpdateEntry4("email_contacto") &"'></td>")
      response.write("<td class='subtitulo mix td-r'>Telefono:</td>")
      response.write("<td class='mix td-l'> <B> "& rsUpdateEntry4("CodigoInternacional") &"</B> &nbsp;&nbsp; <input type='text' name='fTelEmpresa' class='anchox2' value='"& rsUpdateEntry4("telefono_empresa") &"' maxlength=10></td>")
    response.write("</tr>")

    response.write("<tr>")
      response.write("<td class='subtitulo mix td-r'>Fecha solicitud: </td>")
      response.write("<td class='mix td-l'>"& rsUpdateEntry4("FECHA") &"</td>")
      response.write("<td class='subtitulo mix td-r'>Portal Web:</td>")
      response.write("<td class='mix td-l'><input type='text' name='fportal' class='anchox2' value='"& rsUpdateEntry4("PORTAL_EMPRESA") &"'></td>")
    response.write("</tr>")

    response.write("<tr>")
      response.write("<td class='subtitulo mix td-r'>Asesor: </td>")
      response.write("<td class='mix td-l'><select name='fasesor' class='anchox2'>")

             response.write("<option value=0>por asignar</option> ")

             strSQL="select * from OSLP where Active='Y' and slpcode>0"
             rsUpdateEntry.Open strSQL, adoCon2

        while not rsUpdateEntry.EOF
          if rsUpdateEntry4("asesor")=rsUpdateEntry("slpcode") then
              response.write("<option value='"& rsUpdateEntry("slpcode") &"' SELECTED>"& rsUpdateEntry("SlpName") &"</option> ")
          else
              response.write("<option value='"& rsUpdateEntry("slpcode") &"'>"& rsUpdateEntry("SlpName") &"</option> ")
          end if
          rsUpdateEntry.movenext
        wend
        rsUpdateEntry.close
        response.write("</select></td>")
     
      response.write("<td class='subtitulo mix td-r'>Ent. federativa:</td>")
      response.write("<td class='mix td-l'><select name='fentidad' class='anchox2'>")
         

        strSQL="select * from cat_entidades order by id_entidad"
        rsUpdateEntry.Open strSQL, adoCon

        while not rsUpdateEntry.EOF
          if rsUpdateEntry4("id_entidad")=rsUpdateEntry("id_entidad") then
              response.write("<option value='"& rsUpdateEntry("id_entidad") &"' SELECTED>"& rsUpdateEntry("entidad") &"</option> ")
          else
              response.write("<option value='"& rsUpdateEntry("id_entidad") &"'>"& rsUpdateEntry("entidad") &"</option> ")
          end if
          rsUpdateEntry.movenext
        wend
        rsUpdateEntry.close
        response.write("</select></td>")
    response.write("</tr>")

     response.write("<tr>")
      response.write("<td class='subtitulo mix td-r'>&nbsp; </td>")
      response.write("<td class='mix td-l'>&nbsp;</td>")
      response.write("<td class='subtitulo mix td-r'>Ciudad:</td>")
      response.write("<td class='mix td-l'><input type='text' name='fciudad' class='anchox2' value='"& rsUpdateEntry4("ciudad") &"'></td>")
    response.write("</tr>")


    response.write("<tr>")
      response.write("<td class='subtitulo mix td-r'>Solicitud: </td>")
      response.write("<td class='mix td-l' colspan=3><input type='text' name='fsolicitud' style='width:800px' value='"& rsUpdateEntry4("solicitud") &"'></td>")     
    response.write("</tr>")
    

    


    response.write("<tr>")
      response.write("<td class='subtitulo mix td-r'>Cuerpo del email: </td>")
      response.write("<td class='mix td-l' colspan=3>"& rsUpdateEntry4("cuerpo_email") &"</td>")     
    response.write("</tr>")

    response.write("<tr>")
      response.write("<td class='subtitulo mix td-r'>Archivo adjuntado: </td>")
      response.write("<td class='mix td-l' colspan=3><a href='/contactos/"&rsUpdateEntry4("nombre_archivo")&"' target='adjunto'>"& rsUpdateEntry4("nombre_archivo") &"</a></td>")     
    response.write("</tr>")



    if int(trim(rsUpdateEntry4("id_estatus")))<4 then
        response.write("<tr><td class='subtitulo mix td-c' colspan=4>Estatus: <select name='festatus'>")

        strSQL="select * from cat_estatus where id_estatus not in ('00') order by id_estatus"
        rsUpdateEntry.Open strSQL, adoCon

            while not rsUpdateEntry.EOF
              if rsUpdateEntry4("id_estatus")=rsUpdateEntry("id_estatus") then
                  response.write("<option value='"& rsUpdateEntry("id_estatus") &"' SELECTED>"& rsUpdateEntry("estatus") &"</option> ")
              else
                  response.write("<option value='"& rsUpdateEntry("id_estatus") &"'>"& rsUpdateEntry("estatus") &"</option> ")
              end if
              rsUpdateEntry.movenext
            wend
            rsUpdateEntry.close
            response.write("</select>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")
             response.write("<input type='submit' value='actualizar'></form>")

             response.write("<tr><td class='mix td-c' colspan=4><a href='ShowContent.asp?contenido=72&duplicar="& rsUpdateEntry4("ID") &"'>duplicar contacto con otra solicitud </a></td></tr>")
    else
            response.write("<tr><td class='subtitulo mix td-c' colspan=4>Estatus: CERRADO </td></tr>")
            response.write("<tr><td class='mix td-c' colspan=4><a href='ShowContent.asp?contenido=72&duplicar="& rsUpdateEntry4("ID") &"'>duplicar contacto con otra solicitud </a></td></tr>")
            response.write("<tr><td class='mix td-c' colspan=4><a href='ShowContent.asp?contenido=72&action=4&ID="& rsUpdateEntry4("ID") &"'>registrar que se convirtio en cliente </a></td></tr>")
    end if



    response.write("</table><br>")
    response.write("<P>&nbsp;</P>")

    end if  
    

   if request("action")=3 then
       strSQL="update contactos set id_mode="& request("fmode") &",CONTACTO='"& request("fcontacto") &"', EMPRESA='"& request("fempresa") &"', CELULAR='"& request("fcelular") &"', GIRO='"& request("fgiro") &"', EMAIL_CONTACTO='"& request("femail") &"', TELEFONO_EMPRESA='"& request("fTelEmpresa") &"', PORTAL_EMPRESA='"& request("fportal") &"', ASESOR="& request("fasesor") &" , ID_ENTIDAD='"& request("fentidad") &"' , CIUDAD='"& request("fciudad") &"', SOLICITUD='"& request("fsolicitud") &"', ID_ESTATUS='"& request("festatus") &"',LogDate=getdate()  WHERE ID=" & request("ID")

       response.write strSQL
       rsUpdateEntry.Open strSQL, adoCon

       if request("festatus")="04" or request("festatus")="05" or request("festatus")="06" then
          strSQL="update contactos set  fecha_solucion=getdate(),sentdate=null  WHERE ID=" & request("ID")
          rsUpdateEntry2.Open strSQL, adoCon
       end if

       response.redirect("ShowContent.asp?contenido=72&msg=el contacto no. "&  request("ID") &" fue actualizado ")
   end if

end sub





sub ShowContacts

   if request("contenido")=72 then
      titulo="CONTACTOS EN PROCESO"
   else
      titulo="CONTACTOS DE MARKETING ATENDIDOS"  'contenido 76'
       Dotitulo(titulo)
   end if
  
  
  
   CreateTable(1120)   

  Const adCmdText = &H0001
  Const adOpenStatic = 3
  nPage=0
  registros=1


  session("asesor")=""
  if request("asesor")<>"" then
         session("asesor")=trim(request("asesor"))
  end if
  if trim(request("asesor"))="none" then
       session("asesor")=""
  end if




  if request("contenido")=72 and trim(request("fpalabra"))="" then  'contactos atendidos'

  strSQL="select top 400 e.entidad,a.*,b.Mode,.dbo.fn_GetMonthName(a.fecha_solicitud ,'spanish') as 'fechasolicitud',c.SlpName as 'asesor',d.estatus,case when fecha_solucion is not null then .dbo.fn_GetMonthName(a.fecha_solucion ,'spanish') else '' end as fecha_solucion  from contactos A inner join cat_mode b on a.ID_mode=b.ID_mode left join PRODUCTIVA_DMX.dbo.OSLP c on a.asesor=c.SlpCode inner join cat_estatus d on a.id_estatus=d.id_estatus left join cat_entidades e on a.id_entidad=e.id_entidad  where 1=1 " 

     if session("asesor")<>"" then
         strSQL=strSQL & " and a.asesor=" & session("asesor")
     end if

     strSQL=strSQL & " order by A.fecha_solicitud DESC,A.LogDate DESC "

   end if



   if request("contenido")=72 and len(request("fpalabra"))>3 then 

  strSQL="select e.entidad,a.*,b.Mode,.dbo.fn_GetMonthName(a.fecha_solicitud ,'spanish') as 'fechasolicitud',c.SlpName as 'asesor',d.estatus,case when fecha_solucion is not null then .dbo.fn_GetMonthName(a.fecha_solucion ,'spanish') else '' end as fecha_solucion  from contactos A inner join cat_mode b on a.ID_mode=b.ID_mode left join PRODUCTIVA_DMX.dbo.OSLP c on a.asesor=c.SlpCode inner join cat_estatus d on a.id_estatus=d.id_estatus left join cat_entidades e on a.id_entidad=e.id_entidad "

  strSQL=strSQL &"where  ( Empresa like '%"& request("fpalabra") &"%' OR contacto like '%"& request("fpalabra") &"%' OR email_contacto like '%"& request("fpalabra") &"%' OR solicitud like '%"& request("fpalabra") &"%' OR cuerpo_email like '%"& request("fpalabra") &"%' ) "

     if session("asesor")<>"" then
         strSQL=strSQL & " and a.asesor=" & session("asesor")
     end if

     strSQL=strSQL & " order by A.fecha_solicitud DESC,A.LogDate DESC "

   end if




  if  request("contenido")=76    and trim(request("fpalabra"))="" then  'contactos atendidos'
   strSQL="select top 400 e.entidad,a.*,b.Mode,.dbo.fn_GetMonthName(a.fecha_solicitud ,'spanish') as 'fechasolicitud',c.SlpName as 'asesor',d.estatus,case when fecha_solucion is not null then .dbo.fn_GetMonthName(a.fecha_solucion ,'spanish') else '' end as fecha_solucion  from contactos A inner join cat_mode b on a.ID_mode=b.ID_mode left join PRODUCTIVA_DMX.dbo.OSLP c on a.asesor=c.SlpCode inner join cat_estatus d on a.id_estatus=d.id_estatus left join cat_entidades e on a.id_entidad=e.id_entidad  where convert(int,a.ID_ESTATUS) >=4  "  'contactos atendidos'

    if session("asesor")<>"" then
         strSQL=strSQL & " and a.asesor=" & session("asesor")
     end if

     strSQL=strSQL & " order by A.fecha_solicitud DESC,A.LogDate DESC "

  end if

 
  
    response.write("</form>")
    'response.write strSQL
    rsUpdateEntry.Open strSQL, adoCon, adOpenStatic, adCmdText
  

     rsUpdateEntry.PageSize = 10 
     nPageCount = rsUpdateEntry.PageCount

  if not rsUpdateEntry.EOF then
       if request("vPage")<>"" then
            nPage = int(trim(request("vPage")))
        else
            nPage=1
        end if
      rsUpdateEntry.AbsolutePage=npage

      response.write("<tr><td colspan=3 class='td-r'><B><U>PAGINAS:</U></B></td><td colspan=7 class='td-l'>")
      for i=1 to nPageCount step 1
            if i<>npage then  
                response.write("<a href='/modules/ShowContent.asp?contenido="& request("contenido") &"&vPage="&i &"'>") 
                response.write( i &"</A>&nbsp;&nbsp;&nbsp;&nbsp;") 
            end if
          if i=50 or i=100 or i=150 or i=200 then             
               response.write("<br>")
          end if
      next
      response.write("</td></tr>")
  
  else

      NoRegistros

  end if


if rsUpdateEntry.EOF then
    'NoRegistros

else


 while not rsUpdateEntry.EOF AND registros<=10
   response.write("<tr>")
   response.write("<td class=subtitulo>#</td><td class=subtitulo>Metodo <br> entrada</td><td class=subtitulo>Contacto</td><td class=subtitulo>Celular</td><td class=subtitulo>Telefono <br> empresa</td><td class=subtitulo>eMail</td><td class=subtitulo>Fecha</td><td class=subtitulo>Estatus</td><td class=subtitulo>Asesor</td>")
   response.write("</tr>")
   response.write("<tr>")
     response.write("<td class='td-c'><a href='ShowContent.asp?contenido=72&action=2&ID="& rsUpdateEntry("ID") &"'><B>"& rsUpdateEntry("ID") &"&nbsp;<img src='/img/b_edit.png' border=0 with=11px height=11px alt='modificar' title='modificar'></a></B></td>")
     response.write("<td class='td-c'>"& rsUpdateEntry("Mode") &"</td>")
     response.write("<td class='td-c' width='170px'><B>"& rsUpdateEntry("contacto") &"</B></td>")
     response.write("<td class='td-c'>"& rsUpdateEntry("celular") &"</td>")
     response.write("<td class='td-c'>"& rsUpdateEntry("telefono_empresa") &"</td>")
     response.write("<td class='td-c' width='120px' class='fontmed'><a class='fontmed' href='https://mail.google.com/mail/?view=cm&fs=1&tf=1&to="& rsUpdateEntry("email_contacto") &"' target='_blank'>"& rsUpdateEntry("email_contacto") &"</a></td>")
     response.write("<td class='td-c' width='82px'>"& rsUpdateEntry("fechasolicitud") &"</td>")
   

     if int(trim(rsUpdateEntry("id_Estatus")))<3 then
          response.write("<td class='td-c' width='110px' style='color:white;background-color:red;'><a style='color:white' href='ShowContent.asp?contenido=72&action=2&ID="& rsUpdateEntry("ID") &"'>"& rsUpdateEntry("Estatus") &"</A>&nbsp;<font class=fontmed>" & rsUpdateEntry("fecha_solucion") &"</font></td>")
     else
          response.write("<td class='td-c' width='110px' style='background-color:#E2F0E2;'><a href='ShowContent.asp?contenido=72&action=2&ID="& rsUpdateEntry("ID") &"'>"& rsUpdateEntry("Estatus") &"</A>&nbsp;<font class=fontmed>" & rsUpdateEntry("fecha_solucion") &"</font></td>")
     end if

     
     response.write("<td class='td-c' width='170px'><font color=red>"& rsUpdateEntry("asesor") &"</font></td>")
     response.write("</tr>")

     response.write("<tr>")
     response.write("<td colspan=3 class='td-l fontmed'>[Estado:] "& rsUpdateEntry("entidad") &"</td>")
     response.write("<td colspan=2 class='td-l fontmed'>[Ciudad:] "& rsUpdateEntry("ciudad") &"</td>")
     response.write("<td class='td-l fontmed' colspan=3>[Solicitud:] <B>"& rsUpdateEntry("solicitud") &"</B></td>")
     response.write("<td class='td-l fontmed' colspan=2>[Empresa:] "& rsUpdateEntry("Empresa") &"</td>")
     response.write("</tr>")

     response.write("<tr>")
     response.write("<td colspan=10 class='td-l fontmed'>[Archivo adjunto:] <U><a href='/contactos/"& rsUpdateEntry("nombre_archivo")  &"' target='exe'>"& rsUpdateEntry("nombre_archivo") &"</U></a> </td></tr>")
    
    separador
   

   rsUpdateEntry.MoveNext
   registros=registros+1
 wend
   
end if   


   rsUpdateEntry.close
   

   response.write("</table><P>&nbsp;</P>")


end sub  


sub pdfRepositorio 'contenido 73

    Titulo="INFORME DE PDF PEDIDO FALTANTE EN REPOSITORIO"
    DoTitulo(Titulo)

    strSQL="exec [dbo].[SP_RevisionRepositorioVentas]"
    rsUpdateEntry7.Open strSQL, adoCon4

    strSQL="select * from _REVISION"
    rsUpdateEntry.Open strSQL, adoCon4
    rsUpdateEntry2.Open strSQL, adoCon4

    CreateTable(500)
    CamposRS1
    ShowValoresRS2

    response.write("<table><P>&nbsp;</P>")

    strSQL="if exists (select * from INFORMATION_SCHEMA.TABLES where TABLE_NAME = '_REVISION' AND TABLE_SCHEMA = 'dbo')     drop table dbo._REVISION;"
    rsUpdateEntry7.Open strSQL, adoCon4

end sub 



sub DistribucionCuentas 'contenido 74'
    response.redirect("distribucion.asp")
end sub



sub ItemsCotizados    'contenido 75'
    
    DoAlert
    Titulo="ANALISIS EN LA COTIZACION DE ARTICULOS"
    DoTitulo(Titulo)
    response.write("<P>&nbsp;</P>")

    response.write("<form id='inventario' method='POST' action='ShowContent.asp'>")
         response.write("<input type='hidden' name='contenido' value=75>")
         response.write("<input type='hidden' name='action' value=2>")

         CreateTable(500)
         response.write("<tr><td class='td-r'>Familia: </td><td class='td-l'><select name='ffamilia'>")
         strSQL="select * from OITB order by ItmsGrpNam"
         rsUpdateEntry.Open strSQL, adoCon2    'DMX'

         response.write("<option value='0'>sin seleccionar</option>")
         while not rsUpdateEntry.EOF
              if trim(request("ffamilia"))=trim(rsUpdateEntry("ItmsGrpNam")) then
                 response.write("<option value='"& rsUpdateEntry("ItmsGrpNam") &"' SELECTED>"&  rsUpdateEntry("ItmsGrpNam") &"</option>")
              else
                  response.write("<option value='"& rsUpdateEntry("ItmsGrpNam") &"'>"&  rsUpdateEntry("ItmsGrpNam") &"</option>")
              end if
              rsUpdateEntry.MoveNext
         wend
         rsUpdateEntry.close

         response.write("</select></td></tr>")
         response.write("<tr><td class='td-r'>Descripci&oacute;n incluye palabra: </td><td class='td-l'> <input class='anchox2' type='text' name='fpalabra1' value='"& request("fpalabra1") &"'></td></tr>")
         response.write("<tr><td class='td-r'>Descripci&oacute;n incluye otra palabra: </td><td class='td-l'> <input class='anchox2' type='text' name='fpalabra2' value='"& request("fpalabra2") &"'></td></tr>")
         response.write("<tr><td class='td-r'>Descripci&oacute;on <B>NO INCLUYE </B>palabra:  </td><td class='td-l'><input class='anchox2' type='text' name='fpalabra3' value='"& request("fpalabra3") &"'></td></tr>")
         response.write("<tr><td class='td-r'>Descripci&oacute;on <B>NO INCLUYE </B>palabra: </td><td class='td-l'><input class='anchox2' type='text' name='fpalabra4' value='"& request("fpalabra4") &"'></td></tr>")
         response.write("<tr><td class='td-r'>C&oacute;digo:</td><td class='td-l'><input class='anchox2' type='text' name='fcodigo' value='"& request("fcodigo") &"'></td></tr>")
         response.write("<tr><td class='td-r'>Fecha inicio:</td><td class='td-l'><input type='date' name='ffechaI'></td></tr>")
         response.write("<tr><td class='td-r'>Fecha fin:</td><td class='td-l'><input type='date' name='ffechaF'></td></tr>")
  
         response.write("<tr><td class='td-c' colspan=2><br><input type='submit' value='Mostrar Analisis'></form></td></tr></table>")

  vItems=""
  if request("action")="2" then

        if request("ffamilia")="0" and len(trim(request("fpalabra1")))<4 and len(trim(request("fpalabra2")))<4 and len(trim(request("fpalabra3")))<4 and len(trim(request("fcodigo")))<4  then

               response.redirect("ShowContent.asp?contenido=75&msg=debe ser mas especifico con los criterios de busqueda")
        end if


        'CONCATENA TODOS LOS CODIGOS DE UNA FAMILIA SEPARADO POR COMAS'
        strSQL="SELECT CASE WHEN LEN(I.Items)>1 THEN SUBSTRING(I.Items,0,LEN(I.Items)) ELSE '' END as Items FROM (SELECT (SELECT STUFF(    (SELECT ''''+ ItemCode+''''+',' from OITM where 1=1 "
        if request("ffamilia")<>"0" then
             strSQL2="select ItmsGrpCod from OITB where ItmsGrpNam='"& request("ffamilia") &"' "
             rsUpdateEntry.Open strSQL2, adoCon2    'DMX'
             strSQL=strSQL & " and ItmsGrpCod="& rsUpdateEntry("ItmsGrpCod")             
             rsUpdateEntry.close             
        end if
        if request("fpalabra1")<>"" then
            strSQL=strSQL & " and ItemName like '%"& request("fpalabra1") &"%' "
        end if
        if request("fpalabra2")<>"" then
            strSQL=strSQL & " and ItemName like '%"& request("fpalabra2") &"%' "
        end if
        if request("fpalabra3")<>"" then
            strSQL=strSQL & " and ItemName not like '%"& request("fpalabra3") &"%' "
        end if
        if request("fpalabra4")<>"" then
            strSQL=strSQL & " and ItemName not like '%"& request("fpalabra4") &"%' "
        end if
        if request("fcodigo")<>"" then
            strSQL=strSQL & " and ItemCode='"& request("fcodigo") &"' "
        end if
        strSQL=strSQL & " for XML PATH ('')  ),1,0, '' )       ) as Items  ) as I "    
        'response.write strSQL    
        rsUpdateEntry.Open strSQL, adoCon2    'DMX'

        vItems=trim(rsUpdateEntry("Items"))
        rsUpdateEntry.close

        strSQL="if exists (select * from INFORMATION_SCHEMA.TABLES where TABLE_NAME = '_ItemsCotizados' AND TABLE_SCHEMA = 'dbo')     drop table dbo._ItemsCotizados;"
        rsUpdateEntry7.Open strSQL, adoCon4  'SBO'

        strSQL="CREATE TABLE [_ItemsCotizados](    [ItemCode] [varchar](50)  NOT NULL,   [ItemName] [varchar](300) NOT NULL,   [CantPsto] [int] NULL,   [MontPsto] decimal(18,2) NULL,   [CantCompra] [int] NULL,   [MontCompra] decimal(18,2) NULL,  [CantFact] [int] NULL,   [MontFact] decimal(18,2) NULL,  [Bateo] decimal(18,2) NULL,  [Pareto] decimal(18,2) NULL  ) ON [PRIMARY]"
        rsUpdateEntry6.Open strSQL, adoCon4  'SBO'

        
        strSQL= "set dateformat ymd;INSERT INTO _ItemsCotizados (ItemCode,ItemName,CantPsto,MontPsto,CantCompra,MontCompra,CantFact,MontFact) SELECT A.ITEMCODE,A.ItemName,ISNULL(A1.CANTDMX,0)+ISNULL(A5.CANTDFW,0) AS CantPsto,ISNULL(A1.ParaPresupuesto,0)+ISNULL(A5.ParaPresupuesto,0) AS MontPsto,ISNULL(A2.CANTDMX,0)+ISNULL(A6.CANTDFW,0) AS CantCompra,ISNULL(A2.ParaCompra,0)+ISNULL(A6.ParaCompra,0) AS MontCompra, ISNULL(A3.CANTDMX,0)+ISNULL(A4.CANTDMX,0)+ISNULL(A7.CANTDFW,0)+ISNULL(A8.CANTDFW,0) AS CantFact,ISNULL(A3.Facturado,0)+ISNULL(A4.Facturado,0)+ISNULL(A7.Facturado,0)+ISNULL(A8.Facturado,0) AS MontFact FROM PRODUCTIVA_DMX.dbo.OITM A WITH (NOLOCK) "

        'A1 cantidades DMX Para Presupuesto
        strSQL=strSQL & "LEFT JOIN (SELECT O.ITEMCODE, ISNULL(sum(C.Quantity),0) AS CantDMX, ISNULL(sum(C.TOTALFRGN),0) AS ParaPresupuesto FROM PRODUCTIVA_DMX.dbo.OITM O WITH (NOLOCK) JOIN PRODUCTIVA_DMX.dbo.QUT1 C WITH (NOLOCK) ON O.ItemCode = C.ItemCode JOIN PRODUCTIVA_DMX.dbo.OQUT C1 WITH (NOLOCK)  ON C.DOCENTRY = C1.DOCENTRY WHERE C.DOCDATE >='"&request("ffechaI") &"' AND C.DOCDATE <='"&request("ffechaF") &"'  AND C1.CANCELED ='N' AND C1.CARDNAME NOT LIKE ('%deltaflow%') AND C1.U_TIPOCOTIZACION not LIKE('%COMPRA%') AND C.ItemCode in ("& vItems  &") GROUP BY O.ItemCode) A1 ON A.ITEMCODE = A1.ITEMCODE "

        'A5 Cantidades DFW Para Presupuesto
         strSQL=strSQL & "LEFT JOIN (SELECT O.ITEMCODE, ISNULL(sum(C.Quantity),0) AS CantDFW, ISNULL(sum(C.TOTALFRGN),0) AS ParaPresupuesto  FROM PRODUCTIVA_DFW.dbo.OITM O WITH (NOLOCK) JOIN PRODUCTIVA_DFW.dbo.QUT1 C WITH (NOLOCK) ON O.ItemCode = C.ItemCode JOIN PRODUCTIVA_DFW.dbo.OQUT C1 WITH (NOLOCK) ON C.DOCENTRY = C1.DOCENTRY WHERE C.DOCDATE >='"&request("ffechaI") &"' AND C.DOCDATE <='"&request("ffechaF") &"' AND C1.CANCELED ='N' AND C1.CARDNAME NOT LIKE ('%deltaflow%') AND C1.U_TIPOCOTIZACION not LIKE('%COMPRA%')  AND C.ItemCode in ("& vItems  &")  GROUP BY O.ItemCode) A5 ON A.ITEMCODE = A5.ITEMCODE "

        'A2 cantidades DMX Para Compra
        strSQL=strSQL & "LEFT JOIN  (SELECT O.ITEMCODE, ISNULL(sum(C.Quantity),0) AS CantDMX, ISNULL(sum(C.TOTALFRGN),0) AS ParaCompra  FROM PRODUCTIVA_DMX.dbo.OITM O WITH (NOLOCK) JOIN PRODUCTIVA_DMX.dbo.QUT1 C WITH (NOLOCK) ON O.ItemCode = C.ItemCode JOIN PRODUCTIVA_DMX.dbo.OQUT C1 WITH (NOLOCK)  ON C.DOCENTRY = C1.DOCENTRY WHERE C.DOCDATE >='"&request("ffechaI") &"' AND C.DOCDATE <='"&request("ffechaF") &"'  AND C1.CANCELED ='N' AND C1.CARDNAME NOT LIKE ('%deltaflow%') AND C1.U_TIPOCOTIZACION LIKE('%COMPRA%')  AND C.ItemCode in ("& vItems  &")  GROUP BY O.ItemCode) A2 ON A.ITEMCODE = A2.ITEMCODE "

         'A6 Cantidades DFW Para Compra
         strSQL=strSQL & "LEFT JOIN (SELECT O.ITEMCODE, ISNULL(sum(C.Quantity),0) AS CantDFW, ISNULL(sum(C.TOTALFRGN),0) AS ParaCompra  FROM PRODUCTIVA_DFW.dbo.OITM O JOIN PRODUCTIVA_DFW.dbo.QUT1 C WITH (NOLOCK) ON O.ItemCode = C.ItemCode JOIN PRODUCTIVA_DFW.dbo.OQUT C1 WITH (NOLOCK) ON C.DOCENTRY = C1.DOCENTRY WHERE C.DOCDATE >='"&request("ffechaI") &"' AND C.DOCDATE <='"&request("ffechaF") &"' AND C1.CANCELED ='N' AND C1.CARDNAME NOT LIKE ('%deltaflow%') AND C1.U_TIPOCOTIZACION LIKE('%COMPRA%')  AND C.ItemCode in ("& vItems  &")  GROUP BY O.ItemCode) A6 ON A.ITEMCODE = A6.ITEMCODE "


        'A3 DMX FACTURADO
        strSQL=strSQL & "LEFT JOIN (SELECT O.ITEMCODE, ISNULL(sum(C.Quantity),0) AS CantDMX, ISNULL(sum(C.TOTALFRGN),0) AS Facturado FROM PRODUCTIVA_DMX.dbo.OITM O WITH (NOLOCK) JOIN PRODUCTIVA_DMX.dbo.INV1 C WITH (NOLOCK) ON O.ItemCode = C.ItemCode JOIN PRODUCTIVA_DMX.dbo.OINV C1 WITH (NOLOCK)  ON C.DOCENTRY = C1.DOCENTRY WHERE C.DOCDATE >='"&request("ffechaI") &"' AND C.DOCDATE <='"&request("ffechaF") &"'  AND C1.CANCELED ='N' AND C1.CARDNAME NOT LIKE ('%deltaflow%')  AND C.ItemCode in ("& vItems  &")  GROUP BY O.ItemCode) A3 ON A.ITEMCODE = A3.ITEMCODE " 


        'A4 DMX NOTA DE CREDITO
         strSQL=strSQL & "LEFT JOIN (SELECT O.ITEMCODE, -ISNULL(sum(C.Quantity),0) AS CantDMX, -ISNULL(sum(C.TOTALFRGN),0) AS Facturado FROM PRODUCTIVA_DMX.dbo.OITM O WITH (NOLOCK) JOIN PRODUCTIVA_DMX.dbo.RIN1 C WITH (NOLOCK) ON O.ItemCode = C.ItemCode JOIN PRODUCTIVA_DMX.dbo.ORIN C1 WITH (NOLOCK) ON C.DOCENTRY = C1.DOCENTRY WHERE C.DOCDATE >='"&request("ffechaI") &"' AND C.DOCDATE <='"&request("ffechaF") &"'  AND C1.CANCELED ='N' AND C1.CARDNAME NOT LIKE ('%deltaflow%')  AND C.ItemCode in ("& vItems  &")  GROUP BY O.ItemCode) A4 ON A.ITEMCODE = A4.ITEMCODE "


         'A7 DFW FACTURADO
         strSQL=strSQL & "LEFT JOIN  (SELECT  O.ITEMCODE, ISNULL(sum(C.Quantity),0) AS CantDFW, ISNULL(sum(C.TOTALFRGN),0) AS Facturado  FROM PRODUCTIVA_DFW.dbo.OITM O WITH (NOLOCK) JOIN PRODUCTIVA_DFW.dbo.INV1 C WITH (NOLOCK) ON O.ItemCode = C.ItemCode JOIN PRODUCTIVA_DFW.dbo.OINV C1 WITH (NOLOCK) ON C.DOCENTRY = C1.DOCENTRY WHERE C.DOCDATE >='"&request("ffechaI") &"' AND C.DOCDATE <='"&request("ffechaF") &"'  AND C1.CANCELED ='N' AND C1.CARDNAME NOT LIKE ('%deltaflow%')  AND C.ItemCode in ("& vItems  &")  GROUP BY O.ItemCode) A7 ON A.ITEMCODE = A7.ITEMCODE "

         'A8 DFW NOTA DE CREDITO
         strSQL=strSQL & "LEFT JOIN  (SELECT  O.ITEMCODE, -ISNULL(sum(C.Quantity),0) AS CantDFW, -ISNULL(sum(C.TOTALFRGN),0) AS Facturado  FROM PRODUCTIVA_DFW.dbo.OITM O WITH (NOLOCK)  JOIN PRODUCTIVA_DFW.dbo.RIN1 C WITH (NOLOCK) ON O.ItemCode = C.ItemCode JOIN PRODUCTIVA_DFW.dbo.ORIN C1 WITH (NOLOCK)  ON C.DOCENTRY = C1.DOCENTRY WHERE C.DOCDATE >='"&request("ffechaI") &"' AND C.DOCDATE <='"&request("ffechaF") &"'  AND C1.CANCELED ='N' AND C1.CARDNAME NOT LIKE ('%deltaflow%') AND C.ItemCode in ("& vItems  &")  GROUP BY O.ItemCode) A8 ON A.ITEMCODE = A8.ITEMCODE "
         

         strSQL=strSQL & "WHERE A.ItemCode in ("& vItems &")"
         rsUpdateEntry5.Open strSQL, adoCon4  'SBO'
         'response.write strSQL

         strSQL="delete _ItemsCotizados where CantPsto=0 and CantCompra=0"
         rsUpdateEntry4.Open strSQL, adoCon4  'SBO'


         strSQL="UPDATE _ItemsCotizados SET Bateo=CASE WHEN CantCompra>0 THEN CAST(CantFact AS FLOAT )/CAST(CantCompra AS FLOAT ) *100  ELSE 0 END, PARETO= case when (select sum(MontFact) from _ItemsCotizados) >0  then CAST(MontFact AS FLOAT ) / CAST( (select sum(MontFact) from _ItemsCotizados) AS FLOAT ) * 100 else 0 end "
         'response.write strSQL
         rsUpdateEntry3.Open strSQL, adoCon4  'SBO'

         response.write("<P>&nbsp;</P>")
         response.write("<table border=1 cellspacing=2 cellpadding=3 align=center width='1200px' class='table2excel table2excel_with_colors' data-tableName='table1'>")

         response.write("<tr><td style='height:10px;' colspan=2>&nbsp;</td>")
         response.write("<td style='height:10px;' colspan=4 class='td-c strong'> C O T I Z A D O</td>")
         response.write("<td style='height:10px;' colspan=2 class='td-c strong'> F A C T U R A D O </td>")
         response.write("<td style='height:10px;' colspan=3>&nbsp;</td></tr>")


              

         strSQL="select a.ItemCode as codigo,a.ItemName as Descripcion ,a.CantPsto, '$ ' + CONVERT(VARCHAR,CONVERT(MONEY,a.MontPsto),1) as MontPsto,a.CantCompra,'$ ' + CONVERT(VARCHAR,CONVERT(MONEY,a.MontCompra),1) as MontCompra,    a.CantFact, '$ ' + CONVERT(VARCHAR,CONVERT(MONEY,a.MontFact),1) as MontFact, convert(varchar,a.Bateo)+ ' %' as Bateo, a.Pareto,  stock=( select cast( isnull( (select SUM(Y.InQty)-SUM(Y.OutQty) from  PRODUCTIVA_DMX.dbo.OINM Y  WITH (NOLOCK) where y.ItemCode collate SQL_Latin1_General_CP1_CI_AS =a.ItemCode collate SQL_Latin1_General_CP1_CI_AS) ,0) +  isnull( (select SUM(Y.InQty)-SUM(Y.OutQty) from  PRODUCTIVA_DFW.dbo.OINM Y  WITH (NOLOCK) where y.ItemCode collate SQL_Latin1_General_CP1_CI_AS =a.ItemCode collate SQL_Latin1_General_CP1_CI_AS ) ,0) as int ) )  from _ItemsCotizados a order by a.Pareto DESC"

         'response.write strSQL
         rsUpdateEntry.Open strSQL, adoCon4  'SBO'
         rsUpdateEntry2.Open strSQL, adoCon4  'SBO'

         CamposRS1
          
         while not rsUpdateEntry2.EOF

               response.write("<tr><td class='td-l fontmed'>" & mid(rsUpdateEntry2("Codigo"),1,25) & mid(rsUpdateEntry2("Codigo"),26,50) &"</td>")
               response.write("<td class='td-r fonttiny'>" & rsUpdateEntry2("Descripcion") &"</td>")
            
               response.write("<td class='td-r fontmed'>" & replace(replace(FormatCurrency(rsUpdateEntry2("CantPsto")),"$",""),".00","") &"</td>")

               response.write("<td class='td-r fontmed'>" & rsUpdateEntry2("MontPsto") &"</td>")

               response.write("<td class='td-r fontmed' style='background-color:#E1DFDF;'>" & rsUpdateEntry2("CantCompra") &"</td>")
               response.write("<td class='td-r fontmed' style='background-color:#E1DFDF;'>" & rsUpdateEntry2("MontCompra") &"</td>")

               response.write("<td class='td-r fontmed'>" & replace(replace(FormatCurrency(rsUpdateEntry2("CantFact")),"$",""),".00","") &"</td>")

               response.write("<td class='td-r fontmed'>" & rsUpdateEntry2("MontFact") &"</td>")
               response.write("<td class='td-r fontmed'>" & rsUpdateEntry2("Bateo") &"</td>")

               response.write("<td class='td-r fontmed'>" & rsUpdateEntry2("Pareto") &"</td>") 
               response.write("<td class='td-r fontmed'>" & replace(replace(FormatCurrency(rsUpdateEntry2("stock")),"$",""),".00","")  &"</td>")


               '<td class="td-r"><A href="Javascript:sendReq('/modules/DetalleFacturacion.asp','codigo,fechai,fechaf','<%=rsUpdateEntry2("codigo")
               '  =request("ffechaI")    =request("ffechaF")    ','detalle')"> XXX </a></td>       
              

               
               response.write("</tr>")

               rsUpdateEntry2.movenext
        wend
        rsUpdateEntry2.close



         strSQL="select sum(CantPsto) as CantPsto, '$ ' + CONVERT(VARCHAR,CONVERT(MONEY,sum(MontPsto)),1) as MontPsto,sum(CantCompra) as CantCompra,'$ ' + CONVERT(VARCHAR,CONVERT(MONEY,sum(MontCompra)),1) as MontCompra, sum(CantFact) as CantFact, '$ ' + CONVERT(VARCHAR,CONVERT(MONEY,sum(MontFact)),1) as MontFact, sum(Pareto) as Pareto from _ItemsCotizados"
         'response.write strSQL
         rsUpdateEntry7.Open strSQL, adoCon4  'SBO'

         response.write("<tr><td colspan=2 class='td-r subtitulo'>TOTAL</td>")
         response.write("<td class='td-r'>" & replace(replace(FormatCurrency(rsUpdateEntry7("CantPsto")),"$",""),".00","")  &"</td>")

         response.write("<td class='td-r'>" & rsUpdateEntry7("MontPsto") &"</td>")

         response.write("<td class='td-r' style='background-color:#E1DFDF;'>" & replace(replace(FormatCurrency(rsUpdateEntry7("CantCompra")),"$",""),".00","")  &"</td>")

         response.write("<td class='td-r' style='background-color:#E1DFDF;'>" & rsUpdateEntry7("MontCompra") &"</td>")

         response.write("<td class='td-r'>" & replace(replace(FormatCurrency(rsUpdateEntry7("CantFact")),"$",""),".00","")  &"</td>")
         
         response.write("<td class='td-r'>" & rsUpdateEntry7("MontFact") &"</td>")
         response.write("<td class='td-r'>&nbsp;</td>")
        
         response.write("<td class='td-r' style='background-color:#E1DFDF;'>&nbsp;</td>")
         response.write("<td class='td-r'>" & rsUpdateEntry7("Pareto") &"</td>")
         response.write("<td class='td-r'>&nbsp;</td>")
         rsUpdateEntry7.close

         response.write("</table>")   %>  <button class="exportToExcel">Exportar a excel</button> </center>

   <script>
              $(function() {
                $(".exportToExcel").click(function(e){
                  var table = $(this).prev('.table2excel');
                  if(table && table.length){
                    var preserveColors = (table.hasClass('table2excel_with_colors') ? true : false);
                    $(table).table2excel({
                      exclude: ".noExl",
                      name: "Excel Document Name",
                      filename: "myFileName" + new Date().toISOString().replace(/[\-\:\.]/g, "") + ".xls",
                      fileext: ".xls",
                      exclude_img: true,
                      exclude_links: true,
                      exclude_inputs: true,
                      preserveColors: preserveColors
                    });
                  }
                });
                
              });
    </script>

   <%

          response.write("</table><P>&nbsp;</P><div id='detalle'>&nbsp;</div>")

          strSQL="if exists (select * from INFORMATION_SCHEMA.TABLES where TABLE_NAME = '_ItemsCotizados' AND TABLE_SCHEMA = 'dbo')     drop table dbo._ItemsCotizados;"
          rsUpdateEntry7.Open strSQL, adoCon4  'SBO'

  end if


end sub



sub SolicitudCompra  'contenido 77'

  
  
  if request("action")="" then

    response.write("<P>&nbsp;</P>")
    if request("msg")<>"" then
           response.write("<P class='alert td-c'>"&  request("msg") &"</P>")
    end if

    title="SOLICITUD DE COMPRA <BR> <font class='fontmed'> (MISCELANEAS)</font>"
    DoTitulo(title)
    response.write("<P>&nbsp;</P>")

    response.write("<form id='form1' method='POST' action='ShowContent.asp'>")
    response.write("<input type='hidden' name='contenido' value='77'>")
    response.write("<input type='hidden' name='action' value=1>")

    
    CreateTable(800)
    
    response.write("<tr>")
      response.write("<td class='subtitulo mix td-r' width='40%'>Beneficiario:</td>")
      response.write("<td class='mix td-l' width='60%'><select name='fbeneficiario'>")

        strSQL="select * from empleados where FechaEgreso is null order by nombres"
        rsUpdateEntry.Open strSQL, adoCon

        while not rsUpdateEntry.EOF
          response.write("<option value='"& rsUpdateEntry("ID") &"'>"& rsUpdateEntry("nombres") &" " &rsUpdateEntry("paterno") &" " & rsUpdateEntry("materno")  &"</option> ")
          rsUpdateEntry.movenext
        wend
        rsUpdateEntry.close
        response.write("</select></td>")
    response.write("</tr>")


    response.write("<tr>")
      response.write("<td class='subtitulo mix td-r'>Solicitante:</td>")
      response.write("<td class='mix td-l'><select name='fsolicitante'>")

      strSQL="select a.* from Empleados a inner join cat_puesto b on a.Clave_Puesto=b.Clave_Puesto where (b.puesto like '%gerente%' OR b.puesto like '%director%')  and a.FechaEgreso is null order by a.nombres"
        rsUpdateEntry.Open strSQL, adoCon

       while not rsUpdateEntry.EOF
          response.write("<option value='"& rsUpdateEntry("ID") &"'>"& rsUpdateEntry("nombres") &" " &rsUpdateEntry("paterno") &" " & rsUpdateEntry("materno")  &"</option> ")
          rsUpdateEntry.movenext
        wend
        rsUpdateEntry.close
        response.write("</select></td>")
    response.write("</tr>")

     response.write("<tr>")
      response.write("<td class='subtitulo mix td-r'>Fecha Solicitud</td>")
      response.write("<td class='mix td-l'>"& date() &"</td>")
    response.write("</tr>")

    response.write("<tr>")
      response.write("<td class='subtitulo mix td-r'>Fecha estimada para el uso del art&iacute;culo:</td>")
      response.write("<td class='mix td-l'><input type='date' name='fFechaEsperada' ></td>")
    response.write("</tr>")

    response.write("<tr>")
      response.write("<td class='subtitulo mix td-r'>Art&iacute;culo que solicita comprar:</td>")
      response.write("<td class='mix td-l'><input type='text' class='anchox4' name='fsolicitud' ></td>")
    response.write("</tr>")

    response.write("<tr>")
      response.write("<td class='subtitulo mix td-r'>Descripci&oacute;n general:</td>")
      response.write("<td class='mix td-l'><input type='text' class='anchox4' name='fdescripcion' ></td>")
    response.write("</tr>")

    response.write("<tr>")
      response.write("<td class='subtitulo mix td-r'>Posibles Proveedores:</td>")
      response.write("<td class='mix td-l'><input type='text' class='anchox4' name='fproveedor' ></td>")
    response.write("</tr>")

    response.write("<tr>")
      response.write("<td class='subtitulo mix td-r'>Presupuesto que estima:</td>")
      response.write("<td class='mix td-l'><input type='text' class='anchox1' name='fpresupuesto' placeholder='no utilice signos/comas'></td>")
    response.write("</tr>")

    response.write("<tr>")
      response.write("<td class='subtitulo mix td-r'>Va a adjuntar un archivo (<1 Mb):</td>")
      response.write("<td class='mix td-l'><input type='checkbox' name='ffile'></td>")  
    response.write("</tr>")

    response.write("</table><br><input type='submit' value='registrar'></form>")

    response.write("<P>&nbsp;</P>")
    ShowSolicitudCompras

  end if


  if request("action")="1" then   'agregar una solicitud de compra'

       response.write("<P>&nbsp;</P><P>&nbsp;</P>")

    
       strSQL="set dateformat ymd;insert into SolictudCompras (ID,Beneficiario,Solicitante,DocDate,FechaEsperada,LogDate,id_usuario,Solicitud,Descripcion,Proveedores,Presupuesto) values ( isnull( (select max(ID) from SolictudCompras),0)+1 ,"& request("fbeneficiario") &","& request("fsolicitante") &",getdate(),'"&request("fFechaEsperada") &"',getdate(),'"& UCASE(session("session_id")) &"','"& UCASE(request("fsolicitud")) &"','"& UCASE(request("fdescripcion")) &"','"& UCASE(request("fproveedor")) &"',"&request("fpresupuesto") &")"

        response.write(strSQL &"<BR>")
        rsUpdateEntry.Open strSQL, adoCon   


        strSQL="set dateformat ymd;select * from SolictudCompras where id_usuario='"& UCASE(session("session_id")) &"' and  beneficiario="& request("fbeneficiario") &" and  solicitante="& request("fsolicitante") &" and  solicitud='"& UCASE(request("fsolicitud")) &"' and DocDate='"& year(date()) &"-" &month(date()) &"-" &day(date())  &"'" 

        response.write(strSQL &"<BR>")
        rsUpdateEntry2.Open strSQL, adoCon

        if not rsUpdateEntry2.EOF then
              Session("ID")=rsUpdateEntry2("ID")
              if request("ffile")="on" or request("ffile")="true"  then   'va a subir archivo'
                    response.redirect("/upload/uploadform4.asp?ID="& rsUpdateEntry2("ID") )
              else 
                    response.redirect("ShowContent.asp?contenido=77&msg=se agrego nueva solicitud de compra no. " & rsUpdateEntry2("ID") )
              end if
            
        else
            response.redirect("ShowContent.asp?contenido=77&msg=se agrego una nueva solicitud de compra")
        end if


  end if


if request("action")="2" then   'modificar una solicitud de compra'

     strSQL="select a.*,b.Nombres+' '+b.Paterno+' '+b.Materno as NomBeneficiario,c.Nombres+' '+c.Paterno+' '+c.Materno as NomSolicitante,.dbo.fn_GetMonthName(a.DocDate ,'spanish') as 'FechaSolicitud',.dbo.fn_GetMonthName(a.DocDate ,'spanish') as 'FechaEsperada2', CASE Aprobacion WHEN 0 THEN 'en autorizacion' WHEN 1 THEN 'autorizado' WHEN 2 THEN 'finalizado' END as estatus from  SolictudCompras a inner join Empleados b on a.Beneficiario=b.ID inner join Empleados c on a.Solicitante=c.ID where a.ID="& request("ID")

     rsUpdateEntry.Open strSQL, adoCon

     response.write("<form id='SolictudCompras' action='ShowContent.asp' method='POST'>")
     response.write("<input type='hidden' name='contenido' value=77>")
     response.write("<input type='hidden' name='action' value='3'>")

     response.write("<P>&nbsp;</P>")
     response.write("<H1>EDITAR LA SOLICITUD DE COMPRA No. "& request("ID") &"</H1>")
     response.write("<P>&nbsp;</P>")
     CreateTable(800)  

    response.write("<tr>")
      response.write("<td class='subtitulo mix td-r' width='40%'>Beneficiario:</td>")
      response.write("<td class='mix td-l' width='60%'><select name='fbeneficiario'>")

        strSQL="select * from empleados where FechaEgreso is null order by nombres"
        rsUpdateEntry2.Open strSQL, adoCon

        while not rsUpdateEntry2.EOF
           if  rsUpdateEntry("beneficiario")=rsUpdateEntry2("ID") then 
               response.write("<option value='"& rsUpdateEntry2("ID") &"' SELECTED>"& rsUpdateEntry2("nombres") &" " &rsUpdateEntry2("paterno") &" " & rsUpdateEntry2("materno")  &"</option> ")
           else
               response.write("<option value='"& rsUpdateEntry2("ID") &"'>"& rsUpdateEntry2("nombres") &" " &rsUpdateEntry2("paterno") &" " & rsUpdateEntry2("materno")  &"</option> ")
           end if
          rsUpdateEntry2.movenext
        wend
        rsUpdateEntry2.close
        response.write("</select></td>")
    response.write("</tr>")


    response.write("<tr>")
      response.write("<td class='subtitulo mix td-r'>Solicitante:</td>")
      response.write("<td class='mix td-l'><select name='fsolicitante'>")

      strSQL="select a.* from Empleados a inner join cat_puesto b on a.Clave_Puesto=b.Clave_Puesto where (b.puesto like '%gerente%' OR b.puesto like '%director%')  and a.FechaEgreso is null order by a.nombres"
        rsUpdateEntry2.Open strSQL, adoCon

       while not rsUpdateEntry2.EOF
          if  rsUpdateEntry("solicitante")=rsUpdateEntry2("ID") then 
               response.write("<option value='"& rsUpdateEntry2("ID") &"' SELECTED>"& rsUpdateEntry2("nombres") &" " &rsUpdateEntry2("paterno") &" " & rsUpdateEntry2("materno")  &"</option> ")
           else
               response.write("<option value='"& rsUpdateEntry2("ID") &"'>"& rsUpdateEntry2("nombres") &" " &rsUpdateEntry2("paterno") &" " & rsUpdateEntry2("materno")  &"</option> ")
           end if
          rsUpdateEntry2.movenext
        wend
        rsUpdateEntry2.close
        response.write("</select></td>")
    response.write("</tr>")

     response.write("<tr>")
      response.write("<td class='subtitulo mix td-r'>Fecha Solicitud</td>")
      response.write("<td class='mix td-l'>"& rsUpdateEntry("FechaSolicitud") &"</td>")
    response.write("</tr>")

    response.write("<tr>")
      response.write("<td class='subtitulo mix td-r'>Fecha estimada para el uso del art&iacute;culo:</td>")
      response.write("<td class='mix td-l'>"& rsUpdateEntry("FechaEsperada2") &"&nbsp;&nbsp;&nbsp;<input type='date' name='fFechaEsperada' ></td>")
    response.write("</tr>")

    response.write("<tr>")
      response.write("<td class='subtitulo mix td-r'>Art&iacute;culo que solicita comprar:</td>")
      response.write("<td class='mix td-l'><input type='text' class='anchox4' name='fsolicitud' value='"& rsUpdateEntry("solicitud") &"'></td>")
    response.write("</tr>")

    response.write("<tr>")
      response.write("<td class='subtitulo mix td-r'>Descripci&oacute;n general:</td>")
      response.write("<td class='mix td-l'><input type='text' class='anchox4' name='fdescripcion' value='"& rsUpdateEntry("descripcion") &"'></td>")
    response.write("</tr>")

    response.write("<tr>")
      response.write("<td class='subtitulo mix td-r'>Posibles Proveedores:</td>")
      response.write("<td class='mix td-l'><input type='text' class='anchox4' name='fproveedor' value='"& rsUpdateEntry("Proveedores") &"' ></td>")
    response.write("</tr>")

    response.write("<tr>")
      response.write("<td class='subtitulo mix td-r'>Presupuesto que estima:</td>")
      response.write("<td class='mix td-l'><input type='text' class='anchox1' name='fpresupuesto' value='"& rsUpdateEntry("Presupuesto") &"'></td>")
    response.write("</tr>")


    response.write("</table><br><P>&nbsp;</P><input type='submit' value='actualizar'></form>")
    response.write("<P>&nbsp;</P>")

end if





if request("action")="3" then   'modificar una solicitud de compra UP'

   strSQL="UPDATE SolicitudCompras"

end if







end sub






sub AprobarSolicitudCompra 'contenido 78'
    
             if request("action")="1" then  'aprobar presupuesto UP'

                strSQL="set dateformat ymd;UPDATE SolictudCompras set Aprobacion=1,sentdate=null,LogDate=getdate(),FechaPresupuesto='"& request("fechaEjercer") &"'  where ID="&request("ID")
                response.write("<BR><BR><BR><BR>"& strSQL )                
                rsUpdateEntry.Open strSQL, adoCon

                response.redirect("ShowContent.asp?contenido=77&msg=se autorizo presupuesto en solicitud no. "& request("ID")  )

             end if

             if request("action")="2" then  'cerrar solicitud UP'
 
                strSQL="set dateformat ymd;UPDATE SolictudCompras set Aprobacion=2,sentdate=null,FechaSolucion=getdate(),Ejercido="& request("fejercido")&" ,Proveedores='"& request("fproveedor") &"'  where ID="&request("ID")
                response.write("<BR><BR><BR><BR>"& strSQL )                
                rsUpdateEntry.Open strSQL, adoCon

                response.redirect("ShowContent.asp?contenido=77&msg=se finalizo solicitud no. "& request("ID")  )

             end if


             strSQL="select a.*,b.Nombres+' '+b.Paterno+' '+b.Materno as NomBeneficiario,c.Nombres+' '+c.Paterno+' '+c.Materno as NomSolicitante,.dbo.fn_GetMonthName(a.DocDate ,'spanish') as 'FechaSolicitud',.dbo.fn_GetMonthName(a.DocDate ,'spanish') as 'FechaEsperada2'   from SolictudCompras a inner join Empleados b on a.Beneficiario=b.ID inner join Empleados c on a.Solicitante=c.ID where a.id=" & request("ID")
             rsUpdateEntry.Open strSQL, adoCon
             'response.write strSQL
             
             titulo="en blanco"

             select case  rsUpdateEntry("Aprobacion")
             case 0  Titulo="AUTORIZAR PRESUPUESTO COMPRA No. " &rsUpdateEntry("ID")
             case 1  Titulo="CERRAR SOLICITUD No. "&rsUpdateEntry("ID")
             end select
             
             response.write("<P>&nbsp;</P>")
             DoTitulo(Titulo)
             response.write("<P>&nbsp;</P>")
             CreateTable(540)
          
             response.write("<form id='aprobacion' method='post' action='ShowContent.asp'>")
             response.write("<input type='hidden' name='contenido' value='78'>")     
             response.write("<input type='hidden' name='ID' value="& rsUpdateEntry("ID") &">") 
             
             response.write("<tr><td class='td-r azul' width='35%'>Fecha Solicitud:</td><td class='td-l' width='65%'>" & rsUpdateEntry("FechaSolicitud") &"</td></tr>")
             response.write("<tr><td class='td-r azul'>Beneficiario:</td><td class='td-l'>" & rsUpdateEntry("NomBeneficiario") &"</td></tr>")
             response.write("<tr><td class='td-r azul'>Solicitante:</td><td class='td-l'>" & rsUpdateEntry("NomSolicitante") &"</td></tr>")
             response.write("<tr><td class='td-r azul'>Solicitud:</td><td class='td-l'>" & rsUpdateEntry("solicitud") &"</td></tr>")
             response.write("<tr><td class='td-r azul'>Presupuesto:</td><td class='td-l'>$ " & rsUpdateEntry("presupuesto") &"</td></tr>")


             select case  rsUpdateEntry("Aprobacion")
                 case 0   response.write("<tr><td class='td-r azul'>Fecha para ejercer <br>presupuesto:</td><td class='td-l'><input type='date' name='FechaEjercer'>")
                          response.write("<tr><td class='td-r azul'>Acci&oacute;n disponible:</td><td class='td-l'><select name='action'>")
                          response.write("<option value=1>aprobar</option>")
                          response.write("</select></td></tr>")   
                 case 1   response.write("<tr><td class='td-r azul'>Presupuesto ejercido:</td><td class='td-l'><input type='text' name='fejercido' placeholder='no utilice signos/comas'></td></tr>")
                          response.write("<tr><td class='td-r azul'>Proveedor Utilizado:</td><td class='td-l'><input type='text' name='fproveedor'></td></tr>")
                          response.write("<tr><td class='td-r azul'>Acci&oacute;n disponible:</td><td class='td-l'><select name='action'>")
                          response.write("<option value=2>cerrar solicitud</option>")
                          response.write("</select></td></tr>") 
             end select
    
           response.write("<tr><td colspan=2 class='td-c'>&nbsp;</td></tr>")
           response.write("<tr><td colspan=2 class='td-c'><input type='submit' value='actualizar'> </form>  </td> </tr> ")
          
           response.write("</table>")
           rsUpdateEntry.close
   
   
      

end sub 




sub  ShowSolicitudCompras
   
  vMode=""
  if Session("mode")<>"" then
       vMode=Session("mode")
  end if  
  if request("mode")<>""  then
       vMode=request("mode")
  end if

  Const adCmdText = &H0001
  Const adOpenStatic = 3
  nPage=0
  registros=1

  if vMode="" then   'elijio solicitudes abiertas'

    response.write("<H1>SOLICITUDES DE COMPRA ABIERTAS </H1>")
    CreateTable(800) 

   strSQL="select a.*,b.Nombres+' '+b.Paterno+' '+b.Materno as NomBeneficiario,c.Nombres+' '+c.Paterno+' '+c.Materno as NomSolicitante,.dbo.fn_GetMonthName(a.DocDate ,'spanish') as 'FechaSolicitud',.dbo.fn_GetMonthName(a.DocDate ,'spanish') as 'FechaEsperada2', CASE Aprobacion WHEN 0 THEN 'en autorizacion' WHEN 1 THEN 'autorizado' WHEN 2 THEN 'finalizado' END as estatus,case when FechaSolucion is null then 1 else 0 end as mode,.dbo.fn_GetMonthName(a.FechaPresupuesto ,'spanish') as 'FechaEjercer2' from  SolictudCompras a inner join Empleados b on a.Beneficiario=b.ID inner join Empleados c on a.Solicitante=c.ID where a.FechaSolucion is null order by a.FechaEsperada "
  
  else

    response.write("<P>&nbsp;</P><H1>SOLICITUDES DE COMPRA CERRADAS</H1>")
    CreateTable(800) 

   strSQL="select a.*,b.Nombres+' '+b.Paterno+' '+b.Materno as NomBeneficiario,c.Nombres+' '+c.Paterno+' '+c.Materno as NomSolicitante,.dbo.fn_GetMonthName(a.DocDate ,'spanish') as 'FechaSolicitud',.dbo.fn_GetMonthName(a.DocDate ,'spanish') as 'FechaEsperada2', CASE Aprobacion WHEN 0 THEN 'en autorizacion' WHEN 1 THEN 'autorizado' WHEN 2 THEN 'finalizado' END as estatus,case when FechaSolucion is null then 1 else 0 end as mode,.dbo.fn_GetMonthName(a.FechaPresupuesto ,'spanish') as 'FechaEjercer2' from  SolictudCompras a inner join Empleados b on a.Beneficiario=b.ID inner join Empleados c on a.Solicitante=c.ID where a.FechaSolucion is not null order by a.FechaEsperada "

     'response.write("<br><br><br><br>" & strSQL )

  end if

    rsUpdateEntry.Open strSQL, adoCon, adOpenStatic, adCmdText
    'response.write strSQL

if not rsUpdateEntry.EOF then

     rsUpdateEntry.PageSize = 10 
     nPageCount = rsUpdateEntry.PageCount

       if request("vPage")<>"" then
            nPage = int(trim(request("vPage")))
        else
            nPage=1
        end if
      rsUpdateEntry.AbsolutePage=npage

      response.write("<tr><td class='td-r'><B><U>PAGINAS:</U></B></td><td colspan=7 class='td-l'>")
      for i=1 to nPageCount step 1
            if i<>npage then  
                response.write("<a href='/modules/ShowContent.asp?contenido=77&vPage="&i &"'>") 
                response.write( i &"</A>&nbsp;&nbsp;&nbsp;&nbsp;") 
            end if
      next
      response.write("</td></tr>")


 while not rsUpdateEntry.EOF AND registros<=10
   response.write("<tr>")
   response.write("<td class='subtitulo td-r' width='15%'>ID:</td><td class='td-l' width='12%'>")


    select case  rsUpdateEntry("Aprobacion")
         'response.write("<a href='ShowContent.asp?contenido=77&action=2&ID="& rsUpdateEntry("ID") &"'><B>"& rsUpdateEntry("ID") &"&nbsp;<img src='/img/b_edit.png' border=0 with=11px height=11px alt='modificar' title='modificar'></a></td>")

         case 0   response.write("<B>"& rsUpdateEntry("ID") &"&nbsp;</td>")
         case 1   response.write("<B>"& rsUpdateEntry("ID") &"&nbsp;</td>")
         case 2   response.write("<B>"& rsUpdateEntry("ID") &"&nbsp;</td>")
    end select


   response.write("<td class='subtitulo td-r' width='12%'>Beneficiario:</td><td class='td-l fontmed'>"& rsUpdateEntry("NomBeneficiario") &"</td><td class='subtitulo td-r' width='12%'>Adjunto:</td><td class='td-l' width='16%'><a href='/compras/"& rsUpdateEntry("nombre_archivo")&"' target='_new' border=0 alt='' title=''><U>"& rsUpdateEntry("nombre_archivo") &"</U></a></td>")
   response.write("</tr><tr>")
   response.write("<td class='subtitulo td-r'>Fecha  esperada:</td><td class='td-l'>"& rsUpdateEntry("FechaEsperada2") &"</td>")
   response.write("<td class='subtitulo td-r'>Solicitante:</td><td class='td-l fontmed'>"& rsUpdateEntry("NomSolicitante") &"</td>")
   response.write("</tr>")  
   response.write("<tr><td class='subtitulo td-r'>Solicitud:</td><td class='td-l fontmed' colspan=5>"& rsUpdateEntry("solicitud") &"</td></tr>")
   response.write("<tr><td class='subtitulo td-r'>Descripci&oacute;n:</td><td class='td-l fontmed' colspan=5>"& rsUpdateEntry("descripcion") &"</td></tr>")

  
   select case  rsUpdateEntry("Aprobacion")
         case 0   response.write("<tr><td class='subtitulo td-r'>Presupuesto:</td><td class='td-c'>$ "& rsUpdateEntry("Presupuesto") &"</td>")
                  response.write("<td class='subtitulo td-r' width='15%'>Estatus:</td><td class='td-c'>")

                  if session("session_id")="MCABANILLAS" or session("session_id")="AJAUREGUI" then
                        response.write("<a href='ShowContent.asp?contenido=78&ID="& rsUpdateEntry("ID") &"'>"& rsUpdateEntry("estatus") &"&nbsp; <img src='/images/alert.gif' border=0 alt='modificar' title='modificar' width='14px' height='14px'><a></td>")  
                  else
                        response.write( rsUpdateEntry("estatus") &"</td>")  

                  end if
                 response.write("<td class='subtitulo td-r'>Fecha para ejercer:</td><td class='td-c'> &nbsp;</td>")
         case 1  response.write("<tr><td class='subtitulo td-r'>Presupuesto:</td><td class='td-c'>$ "& rsUpdateEntry("Presupuesto") &"</td>")
                  response.write("<td class='subtitulo td-r' width='15%'>Estatus:</td><td class='td-c'>")
                  if session("session_id")<>"AJAUREGUI" then
                      response.write("<a href='ShowContent.asp?contenido=78&ID="& rsUpdateEntry("ID") &"'>"& rsUpdateEntry("estatus") &"&nbsp; <img src='/images/alert.gif' border=0 alt='modificar' title='modificar' width='14px' height='14px'><a></td>")
                  else
                      response.write( rsUpdateEntry("estatus") &"</td>")
                  end if  
                 response.write("<td class='subtitulo td-r'>Fecha para ejercer:</td><td class='td-c' style='color:red'> "& rsUpdateEntry("FechaEjercer2") &"</td>")
         case 2  response.write("<tr><td class='subtitulo td-r'>Ejercido:</td><td class='td-c'>$ "& rsUpdateEntry("Ejercido") &"</td>")
                  response.write("<td class='subtitulo td-r' width='15%'>Estatus:</td><td class='td-c'>")
                  response.write( rsUpdateEntry("estatus") &"</td>")  
                 response.write("<td class='subtitulo td-r'>Provedor:</td><td class='td-c'> "& rsUpdateEntry("Proveedores") &"</td>")
    end select


   separador
   
   
   rsUpdateEntry.MoveNext
   registros=registros+1

   

 wend




else

     NoRegistros

end if
   
   rsUpdateEntry.close
   

   response.write("</table>")
   response.write("<P>&nbsp;</P>")

end sub


sub ABCIngresos   'contenido 80'
        
        

    if request("action")="" then
      
         DoAlert
         Titulo="<a href='ShowContent.asp?contenido=80' class='H1'>INGRESOS REGISTRADOS POR HELPDESK</a><BR><font class='fontmed'>Este modulo se utiliza para ajustar los registros de ingresos de helpdesk del mes que acaba de concluir</font>"
         DoTitulo(Titulo)
         'response.write("<P>&nbsp;</P>")

         response.write("<form id='inventario' method='POST' action='ShowContent.asp'>")
         response.write("<input type='hidden' name='contenido' value='80'>")        
        
         CreateTable(500)

         if month(date())=1 then
             vAnio=year(date())-1
             vMes=12
         else
             vAnio=year(date())
             vMes=month(date())-1
         end if

         response.write("<tr><td class='td-r fontmed'>a&ntilde;o fiscal:</td><td class='td-l fontmed'>"& vAnio &"</td></tr>")
         response.write("<input type='hidden' name='anio' value="& vAnio &">")
         response.write("<tr><td class='td-r fontmed'>mes fiscal:</td><td class='td-l fontmed'>"& vMes &"</td></tr>")
         response.write("<input type='hidden' name='mes' value="& vMes &">")

         response.write("<tr><td class='td-r fontmed'>acci&oacute;n a realizar:</td><td class='td-l'><select name='action'><option value=1>agregar un registro</option></td></tr>")         
         response.write("<tr><td class='td-c' colspan=2><br><input type='submit' value='continuar'></form></td></tr>")

         closeTable
         response.write("<P>&nbsp;</P>")

         ShowIngresosHelpdesk

     end if



    
    if request("action")="1" then

         response.write("<P>&nbsp;</P>")
         response.write("<form id='inventario' method='POST' action='ShowContent.asp'>")
         response.write("<input type='hidden' name='contenido' value='80'>")
         response.write("<input type='hidden' name='action' value='6'>")
        
         CreateTable(500)

         if month(date())=1 then
             vAnio=year(date())-1
             vMes=12
         else
             vAnio=year(date())
             vMes=month(date())-1
         end if

         response.write("<tr><td class='td-r fontmed'>a&ntilde;o fiscal:</td><td class='td-l fontmed'>"& vAnio &"</td></tr>")
         response.write("<input type='hidden' name='anio' value="& vAnio &">")
         response.write("<tr><td class='td-r fontmed'>mes fiscal:</td><td class='td-l fontmed'>"& vMes &"</td></tr>")
         response.write("<input type='hidden' name='mes' value="& vMes &">")

         response.write("<tr><td class='td-r fontmed'>tipo de documento a agregar:</td><td class='td-l'><select name='tipo'><option value='Ingresos'>factura</option><option value='Anticipos'>anticipo</option><option value='NotaCreditos'>nota de credito</option></select></td></tr>")

         response.write("<tr><td class='td-r fontmed'>razon social:</td><td class='td-l'><select name='RS'><option value='DMX'>DMX</option><option value='DFW'>DFW</option></select></td></tr>")

         response.write("<tr><td class='td-r fontmed'># documento:</td><td class='td-l'><input type='number' name='docnum'></td></tr>")          
         response.write("<tr><td class='td-c' colspan=2><br><input type='submit' value='agregar'></form></td></tr>")


    end if

    if request("action")="2" then

         response.write("<P>&nbsp;</P>")  

         Titulo="<a href='ShowContent.asp?contenido=80' class='H1'>INGRESOS REGISTRADOS POR HELPDESK</a>"
         doTitulo(titulo)
                            
         if month(date())=1 then
             vAnio=year(date())-1
             vMes=12
         else
             vAnio=year(date())
             vMes=month(date())-1
         end if

         strSQL="select RazonSocial,Docnum as Folio,.dbo.fn_GetMonthName(DocDate,'spanish') as Fecha,CardName as SocioNegocio, CONVERT(VARCHAR,CONVERT(MONEY,subtotal),1) as Subtotal ,DocRate as TC,Tipo,U_CFDi_MetodoPago as MetodoPago,Asesor,isnull(convert(varchar,Pedido),'') as Pedido from HistoricoVentas"&vAnio &" where RazonSocial='"& request("RS") &"' and  tipo='"& request("tipo") &"' and Docnum="&request("docnum")
         'response.write strSQL
         rsUpdateEntry.Open strSQL, adoCon4
         rsUpdateEntry2.Open strSQL, adoCon4

         CreateTable(1100)
         Dim Campos(10) 
         i=1
         RowIn
         For Each rsUpdateEntry in rsUpdateEntry.Fields            
            Response.Write "<td class='subtitulo td-c'>" & rsUpdateEntry.Name & "</td>"
            campos(i)=rsUpdateEntry.Name
            i=i+1           
         Next        
         RowOut

         while not rsUpdateEntry2.EOF
             rowin
             for i=1 to 10
                   Response.write("<td class='td-r fonttiny'>"& rsUpdateEntry2(campos(i)) &"</td>")
             next             
             rsUpdateEntry2.MoveNext
             RowOut
         wend
         
         rsUpdateEntry2.close
         closeTable

         response.write("<form id='inventario' method='POST' action='ShowContent.asp'>")
         response.write("<input type='hidden' name='contenido' value='80'>")
         response.write("<input type='hidden' name='action' value='3'>")
         response.write("<input type='hidden' name='RS' value='"&request("RS") &"'>")
         response.write("<input type='hidden' name='tipo' value='"&request("tipo") &"'>")
         response.write("<input type='hidden' name='folio' value='"&request("docnum") &"'>")

         response.write("<P class='td-c'><select name='opcion'><option value='1'>esta seguro desea borrar registro</option><select></P>")          
         response.write("<P class='td-c'><input type='submit' value='borrar'></P></form>")


    end if

    if request("action")="3" then
         if month(date())=1 then
             vAnio=year(date())-1
             vMes=12
         else
             vAnio=year(date())
             vMes=month(date())-1
         end if
         strSQL="delete from HistoricoVentas"&vAnio &" where RazonSocial='"& request("RS") &"' and tipo='"&request("tipo")&"' and docnum="& request("folio")
         response.write strSQL
         rsUpdateEntry.Open strSQL, adoCon4
         response.redirect("ShowContent.asp?contenido=80&msg=registro borrado")   

    end if



    if request("action")="6" then

         if month(date())=1 then
             vAnio=year(date())-1
             vMes=12
         else
             vAnio=year(date())
             vMes=month(date())-1
         end if

        strSQL="select * from HistoricoVentas"&vAnio &" where RazonSocial='"& request("RS") &"' and  tipo='"& request("tipo") &"' and Docnum="&request("docnum")
        'response.write strSQL
        rsUpdateEntry.Open strSQL, adoCon4

        if not rsUpdateEntry.EOF then
            rsUpdateEntry.close
            response.redirect("ShowContent.asp?contenido=80&msg=numero de documento y tipo ya se encuentra registrado")
        end if
        rsUpdateEntry.close


        select case request("tipo")

          case "Ingresos"
               select case request("RS")
                case "DMX"   
                      strSQL="select * from OINV where docnum="& request("docnum")
                      rsUpdateEntry.Open strSQL, adoCon2 'DMX'

                      if rsUpdateEntry.EOF then
                         rsUpdateEntry.close
                         response.redirect("ShowContent.asp?contenido=80&msg=factura no existe")
                      end if

                      strSQL="insert into HistoricoVentas"& vAnio &"  (DocDate,DocNum,CardCode,CardName,subtotal,Iva,total,DocCur,DocRate,RazonSocial,Tipo,asesor,U_PROYECTO,U_CFDi_MetodoPago)  Select t1.DocDate,t1.docnum,t1.CardCode,t1.CardName,t1.DocTotalSy-t1.VatSumSy as 'subtotal',t1.VatSumSy as 'Iva', t1.DocTotalSy as 'total',DocCur,DocRate,'DMX' as RazonSocial,'Ingresos' as Tipo,S.SlpName as 'asesor',t1.U_PROYECTO,T1.U_CFDi_MetodoPago FROM [PRODUCTIVA_DMX].dbo.OINV t1 inner join  [PRODUCTIVA_DMX].dbo.OSLP S ON T1.SlpCode = S.SlpCode  inner join [PRODUCTIVA_DMX].dbo.NNM1 SE on t1.Series=SE.Series  where t1.Docnum=" & request("docnum")

                         rsUpdateEntry7.Open strSQL, adoCon4 'SBO'
                         response.redirect("ShowContent.asp?contenido=80&msg=la factura DMX fue agregada")

                case "DFW"
                      strSQL="select * from OINV where docnum="& request("docnum")
                      rsUpdateEntry.Open strSQL, adoCon3

                      if rsUpdateEntry.EOF then
                         rsUpdateEntry.close
                         response.redirect("ShowContent.asp?contenido=80&msg=factura no existe")
                      end if

                      strSQL="insert into HistoricoVentas"& vAnio &"  (DocDate,DocNum,CardCode,CardName,subtotal,Iva,total,DocCur,DocRate,RazonSocial,Tipo,asesor,U_PROYECTO,U_CFDi_MetodoPago)  Select t1.DocDate,t1.docnum,t1.CardCode,t1.CardName,t1.DocTotalSy-t1.VatSumSy as 'subtotal',t1.VatSumSy as 'Iva', t1.DocTotalSy as 'total',DocCur,DocRate,'DFW' as RazonSocial,'Ingresos' as Tipo,S.SlpName as 'asesor',t1.U_PROYECTO,T1.U_CFDi_MetodoPago FROM [PRODUCTIVA_DFW].dbo.OINV t1 inner join  [PRODUCTIVA_DFW].dbo.OSLP S ON T1.SlpCode = S.SlpCode  inner join [PRODUCTIVA_DFW].dbo.NNM1 SE on t1.Series=SE.Series  where t1.Docnum=" & request("docnum")

                         rsUpdateEntry7.Open strSQL, adoCon4 'SBO'
                         response.redirect("ShowContent.asp?contenido=80&msg=la factura DFW fue agregada")

               end select

          case "Anticipos"
               select case request("RS")
                case "DMX"
                      strSQL="select * from ODPI where docnum="& request("docnum")
                      rsUpdateEntry.Open strSQL, adoCon2

                      if rsUpdateEntry.EOF then
                         rsUpdateEntry.close
                         response.redirect("ShowContent.asp?contenido=80&msg=anticipo no existe")
                      end if

                      strSQL="insert into HistoricoVentas"& vAnio &" (DocDate,DocNum,CardCode,CardName,subtotal,Iva,total,DocCur,DocRate,RazonSocial,Tipo,asesor,U_PROYECTO,U_CFDi_MetodoPago)  SELECT ODPI.DocDate,ODPI.DocNum,ODPI.CardCode,ODPI.CardName,ODPI.DocTotalSy-ODPI.VatSumSy,ODPI.VatSumSy,ODPI.DocTotalSy, ODPI.DocCur, ODPI.DocRate,'DMX' as RazonSocial,'Anticipos' as Tipo,S.SlpName as 'asesor',ODPI.U_PROYECTO,ODPI.U_CFDi_MetodoPago FROM [PRODUCTIVA_DMX].dbo.ODPI ODPI inner join  [PRODUCTIVA_DMX].dbo.OSLP S WITH (NOLOCK) ON ODPI.SlpCode = S.SlpCode inner join [PRODUCTIVA_DMX].dbo.NNM1 SE on ODPI.Series=SE.Series where ODPI.docnum="& request("docnum")
                       
                      response.write strSQL
                      rsUpdateEntry7.Open strSQL, adoCon4 'SBO'
                      response.redirect("ShowContent.asp?contenido=80&msg=el anticipo DMX fue agregado")

                case "DFW"
                      strSQL="select * from ODPI where docnum="& request("docnum")
                      rsUpdateEntry.Open strSQL, adoCon3

                      if rsUpdateEntry.EOF then
                         rsUpdateEntry.close
                         response.redirect("ShowContent.asp?contenido=80&msg=anticipo no existe")
                      end if

                       strSQL="insert into HistoricoVentas"& vAnio &" (DocDate,DocNum,CardCode,CardName,subtotal,Iva,total,DocCur,DocRate,RazonSocial,Tipo,asesor,U_PROYECTO,U_CFDi_MetodoPago)  SELECT ODPI.DocDate,ODPI.DocNum,ODPI.CardCode,ODPI.CardName,ODPI.DocTotalSy-ODPI.VatSumSy,ODPI.VatSumSy,ODPI.DocTotalSy, ODPI.DocCur, ODPI.DocRate,'DFW' as RazonSocial,'Anticipos' as Tipo,S.SlpName as 'asesor',ODPI.U_PROYECTO,ODPI.U_CFDi_MetodoPago FROM [PRODUCTIVA_DFW].dbo.ODPI ODPI inner join  [PRODUCTIVA_DFW].dbo.OSLP S WITH (NOLOCK) ON ODPI.SlpCode = S.SlpCode inner join [PRODUCTIVA_DFW].dbo.NNM1 SE on ODPI.Series=SE.Series where ODPI.docnum="& request("docnum")

                      rsUpdateEntry7.Open strSQL, adoCon4 'SBO'
                         response.redirect("ShowContent.asp?contenido=80&msg=el anticipo DFW fue agregado")

               end select

          case "NotaCreditos"
               select case request("RS")
                case "DMX"
                      strSQL="select * from ORIN where docnum="& request("docnum")
                      rsUpdateEntry.Open strSQL, adoCon2

                      if rsUpdateEntry.EOF then
                         rsUpdateEntry.close
                         response.redirect("ShowContent.asp?contenido=80&msg=nota de credito no existe")
                      end if

                       strSQL="insert into HistoricoVentas"& vAnio &"  (DocDate,DocNum,CardCode,CardName,subtotal,Iva,total,DocCur,DocRate,RazonSocial,Tipo,asesor,U_PROYECTO,U_CFDi_MetodoPago)  SELECT ORIN.DocDate,ORIN.DocNum,ORIN.CardCode,ORIN.CardName,-1*(ORIN.DocTotalSy-ORIN.VatSumSy),-1*ORIN.VatSumSy,-1*(ORIN.DocTotalSy),ORIN.DocCur,ORIN.DocRate,'DMX' as RazonSocial,'NotaCreditos' as Tipo,S.SlpName as 'asesor',ORIN.U_PROYECTO,ORIN.U_CFDi_MetodoPago FROM [PRODUCTIVA_DMX].dbo.ORIN  inner join  [PRODUCTIVA_DMX].dbo.OSLP S WITH (NOLOCK) ON ORIN.SlpCode = S.SlpCode inner join [PRODUCTIVA_DMX].dbo.NNM1 SE on ORIN.Series=SE.Series where ORIN.DocNum=" & request("docnum")

                      rsUpdateEntry7.Open strSQL, adoCon4 'SBO'
                         response.redirect("ShowContent.asp?contenido=80&msg=la nota de credito DMX fue agregada")

                case "DFW"
                      strSQL="select * from ORIN where docnum="& request("docnum")
                      rsUpdateEntry.Open strSQL, adoCon3

                      if rsUpdateEntry.EOF then
                         rsUpdateEntry.close
                         response.redirect("ShowContent.asp?contenido=80&msg=nota de credito no existe")
                      end if

                      strSQL="insert into HistoricoVentas"& vAnio &"  (DocDate,DocNum,CardCode,CardName,subtotal,Iva,total,DocCur,DocRate,RazonSocial,Tipo,asesor,U_PROYECTO,U_CFDi_MetodoPago)  SELECT ORIN.DocDate,ORIN.DocNum,ORIN.CardCode,ORIN.CardName,-1*(ORIN.DocTotalSy-ORIN.VatSumSy),-1*ORIN.VatSumSy,-1*(ORIN.DocTotalSy),ORIN.DocCur,ORIN.DocRate,'DFW' as RazonSocial,'NotaCreditos' as Tipo,S.SlpName as 'asesor',ORIN.U_PROYECTO,ORIN.U_CFDi_MetodoPago FROM [PRODUCTIVA_DFW].dbo.ORIN  inner join  [PRODUCTIVA_DFW].dbo.OSLP S WITH (NOLOCK) ON ORIN.SlpCode = S.SlpCode inner join [PRODUCTIVA_DFW].dbo.NNM1 SE on ORIN.Series=SE.Series where ORIN.DocNum=" & request("docnum")

                      rsUpdateEntry7.Open strSQL, adoCon4 'SBO'
                         response.redirect("ShowContent.asp?contenido=80&msg=la nota de credito DFW fue agregada")

               end select

        end select


    end if


end sub



sub ShowIngresosHelpdesk


         if month(date())=1 then
             vAnio=year(date())-1
             vMes=12
         else
             vAnio=year(date())
             vMes=month(date())-1
         end if
         
         response.write("<form id='inventario' method='POST' action='ShowContent.asp'>")
         response.write("<input type='hidden' name='contenido' value='80'>")
         

         response.write("<H3>Ingresos helpdesk a&ntilde;o: " & vAnio &" mes: " &meses(vMes) &"&nbsp;&nbsp;&nbsp;buscar: <input type='text' name='fpalabra'>&nbsp;&nbsp;<input type='submit' value='buscar'></form></H3>")
         
          
         response.write("<center><table width='1100px' cellpadding=4 border=0  class='mix table2excel table2excel_with_colors' data-tableName='table1' >")


         strSQL="select RazonSocial,Docnum as Folio,.dbo.fn_GetMonthName(DocDate,'spanish') as Fecha,CardName as SocioNegocio, CONVERT(VARCHAR,CONVERT(MONEY,subtotal),1) as Subtotal ,DocRate as TC,Tipo,U_CFDi_MetodoPago as MetodoPago,Asesor,isnull(convert(varchar,Pedido),'') as Pedido from historicoventas"& vAnio &"  where month(DocDate)="& vMes

         if request("fpalabra")<>"" then
               strSQL= strSQL &" and DocNum=" & request("fpalabra")
         else
               strSQL= strSQL &" order by RazonSocial,tipo,docnum"
         end if

         rsUpdateEntry.Open strSQL, adoCon4
         rsUpdateEntry2.Open strSQL, adoCon4

         Dim Campos(10) 
         i=1
         RowIn
         For Each rsUpdateEntry in rsUpdateEntry.Fields            
            Response.Write "<td class='subtitulo td-c'>" & rsUpdateEntry.Name & "</td>"
            campos(i)=rsUpdateEntry.Name
            i=i+1           
         Next 
         Response.Write "<td class='subtitulo td-c'>&nbsp;</td>"
         RowOut

         while not rsUpdateEntry2.EOF
             rowin
             for i=1 to 10
                   Response.write("<td class='td-r fonttiny'>"& rsUpdateEntry2(campos(i)) &"</td>")
             next
             Response.write("<td class='td-r fonttiny'><a href='ShowContent.asp?contenido=80&action=2&RS="&rsUpdateEntry2("RazonSocial")&"&docnum="&rsUpdateEntry2("Folio")&"&tipo="&rsUpdateEntry2("tipo")&"'><img src='/img/b_drop.png' with='8px' alt='borrar' title='borrar' height='8px' border=0></a></td>")
             rsUpdateEntry2.MoveNext
             RowOut
         wend

          strSQL="select CONVERT(VARCHAR,CONVERT(MONEY,sum(subtotal)),1) as calculo from historicoventas"& vAnio &"  where month(DocDate)="& vMes
          'response.write strSQL
          rsUpdateEntry3.Open strSQL, adoCon4

         response.write("<tr><td colspan=4 class='td-r subtitulo'>TOTAL</td><td>"& rsUpdateEntry3("calculo") &"</td><td colspan=5 class='td-r'>&nbsp;</td></tr>")
         rsUpdateEntry3.close

         %>
         </table><button class="exportToExcel">Exportar a excel</button>          

         
    <script>
      $(function() {
        $(".exportToExcel").click(function(e){
          var table = $(this).prev('.table2excel');
          if(table && table.length){
            var preserveColors = (table.hasClass('table2excel_with_colors') ? true : false);
            $(table).table2excel({
              exclude: ".noExl",
              name: "Excel Document Name",
              filename: "myFileName" + new Date().toISOString().replace(/[\-\:\.]/g, "") + ".xls",
              fileext: ".xls",
              exclude_img: true,
              exclude_links: true,
              exclude_inputs: true,
              preserveColors: preserveColors
            });
          }
        });
        
      });
    </script>

    <P>&nbsp;</P>

    <%  
end sub




sub EstatusCliente  'contenido 81'

  titulo="<a href='ShowContent.asp?contenido=81'>Pendiente Suministro versus L&iacute;nea de Cr&eacute;dito</a>"
  DoTitulo(titulo)


  strSQL="if exists (select * from INFORMATION_SCHEMA.TABLES where TABLE_NAME = '_TEMP_Cartera' AND TABLE_SCHEMA = 'dbo')     drop table dbo._TEMP_Cartera;"  
  rsUpdateEntry4.Open strSQL, adoCon4

  strSQL="if exists (select * from INFORMATION_SCHEMA.TABLES where TABLE_NAME = '_TEMP_COBRA' AND TABLE_SCHEMA = 'dbo')     drop table dbo._TEMP_COBRA;"  
  rsUpdateEntry5.Open strSQL, adoCon4


  strSQL="exec SP_Cartera_clientes_intranet"  
  rsUpdateEntry5.Open strSQL, adoCon4


  '---------------------------------------------------------------------------------------------------------------------------------------------------------------------'

if request("RS")="" then

   strSQL="SELECT a.*,Suministrar=isnull( ( select  CONVERT(VARCHAR,CONVERT(MONEY, sum( round(T1.OpenQty*T1.Price,2) )),1) as subtotal  FROM [PRODUCTIVA_DMX].dbo.ORDR T0 WITH (NOLOCK) inner join [PRODUCTIVA_DMX].dbo.RDR1 T1 WITH (NOLOCK) on T0.DocEntry=T1.DocEntry  WHERE T0.cardcode=a.CardCode  and T0.CANCELED='N' and T0.DocStatus='O' and T1.LineStatus='O' ),'0')  FROM SBOTemp.dbo._TEMP_Cartera a where RazonSocial='DMX' order by a.CardName  "   
       
   'response.write(strSQL&"<br><br>")
   rsUpdateEntry.Open strSQL, adoCon2   'DMX'

    '--DFW'
    strSQL="SELECT AUX.* FROM OCRD x INNER JOIN ( 	SELECT 'DFW' as RS,T0.CardCode,substring((select CardName from OCRD where CardCode=T0.CardCode),1,35) 'SocioNegocio', CONVERT(VARCHAR,CONVERT(MONEY, sum( round(T1.OpenQty*T1.Price,2) )),1) as subtotal,		   sum( round(T1.OpenQty*T1.Price,2) ) as subtotal2 	FROM ORDR T0 WITH (NOLOCK) inner join RDR1 T1 WITH (NOLOCK) on T0.DocEntry=T1.DocEntry 	WHERE T0.cardcode<>'C000001' and T0.CANCELED='N' and T0.DocStatus='O' and T1.LineStatus='O' GROUP BY T0.CardCode	UNION	SELECT a.RazonSocial,a.CardCode,substring((select CardName from OCRD where CardCode=a.CardCode),1,35) 'SocioNegocio','0' as subtotal,0 as subtotal2 	FROM SBOTemp.dbo._TEMP_COBRA a  	WHERE a.RazonSocial='DFW' AND a.CardCode NOT IN (SELECT CardCode FROM ORDR Where DocStatus='O') 	GROUP BY a.RazonSocial,a.CardCode  ) AS AUX ON x.CardCode=AUX.CardCode ORDER BY AUX.subtotal2 DESC  "
   
   'response.write(strSQL&"<br><br>")
   rsUpdateEntry2.Open strSQL, adoCon3   'DFW'

else
    
    strSQL="SELECT '"& request("RS") &"' as RazonSocial,T0.CardCode,T0.cardname,CONVERT(VARCHAR,CONVERT(MONEY, sum( round(T1.OpenQty*T1.Price,2) )),1) as Suministrar  FROM ORDR T0 WITH (NOLOCK) inner join RDR1 T1 WITH (NOLOCK) on T0.DocEntry=T1.DocEntry   WHERE T0.cardcode='"& request("SN") &"' and T0.CANCELED='N' and T0.DocStatus='O' and T1.LineStatus='O' GROUP BY T0.CardCode,T0.CardName   "   
    'response.write strSQL
    select case request("RS")
    case "DMX"  rsUpdateEntry.Open strSQL, adoCon2   'DMX'
    case "DFW"  rsUpdateEntry.Open strSQL, adoCon3   'DFW'
    end select
    

    if rsUpdateEntry.EOF then
           rsUpdateEntry.close
           strSQL="SELECT '"& request("RS") &"' as RazonSocial,T0.CardCode,T0.CardName,0 as Suministrar from OCRD T0 WHERE T0.cardcode='"& request("SN") &"' " 

            select case request("RS")
		    case "DMX"  rsUpdateEntry.Open strSQL, adoCon2   'DMX'
		    case "DFW"  rsUpdateEntry.Open strSQL, adoCon3   'DFW'
		    end select
             
           vSocioNegocio=rsUpdateEntry("CardName")       
    else
           vSocioNegocio=rsUpdateEntry("CardName") 
    end if
end if  

  


   'TOTAL FACTURADO EN EL MES'
   strSQL="select  CONVERT(VARCHAR,CONVERT(MONEY,  isnull( (select sum(subtotal) from HistoricoVentas"& year(date()) &" where month(DocDate)=month(getdate()) ) ,0) ),1) as calculo"  
   rsUpdateEntry4.Open strSQL, adoCon4

 

  CreateTable(1200) 
  response.write("<tr><td class='td-c subtitulo'>RS</td><td class='td-c subtitulo'>Socio negocio</td><td class='td-c subtitulo'>Monto a suministrar</td><td class='td-c subtitulo'>Limite</td><td class='td-c subtitulo'>Dias</td><td class='td-c subtitulo' width='70px'>Adeudo</td><td class='td-c subtitulo' width='70px'>Vencido</td><td class='td-c subtitulo' width='70px'>Corriente</td><td class='td-c subtitulo' width='70px'>Refactrcion</td><td class='td-c subtitulo' width='70px'>91+</td><td class='td-c subtitulo' width='70px'>61-90</td><td class='td-c subtitulo' width='70px'>31-60</td><td class='td-c subtitulo' width='70px'>01-30</td></tr>")


  vTotales=""
 
 
  while not rsUpdateEntry.EOF    'imprimiendo DMX'
      response.write ("<tr>")  
      response.write ("<td class='td-c fontmed strong'>"& rsUpdateEntry("RazonSocial") &"</td>")    
      response.write ("<td class='td-l fonttiny'>"& mid(rsUpdateEntry("CardName"),1,30) &"</td>")   
    
    
    
       response.write ("<td class='td-r fontmed strong'><a href='ShowContent.asp?contenido=81&RS="& rsUpdateEntry("RazonSocial") &"&SN="& rsUpdateEntry("CardCode") &"'>"& rsUpdateEntry("Suministrar") &"</a></td>")       

       vTotales= vTotales  &"+" & trim( rsUpdateEntry("Suministrar") )
       'response.write( vTotales &"<br>") 


      strSQL="select CONVERT(VARCHAR,CONVERT(MONEY, Limite  ),1) as Limite,ExtraDays as Dias,CONVERT(VARCHAR,CONVERT(MONEY, total ),1) as total, CONVERT(VARCHAR,CONVERT(MONEY, vencido ),1) as vencido, CONVERT(VARCHAR,CONVERT(MONEY, normal ),1) as normal, CONVERT(VARCHAR,CONVERT(MONEY, refacturacion ),1) as refacturacion, CONVERT(VARCHAR,CONVERT(MONEY, periodo4 ),1) as periodo4, CONVERT(VARCHAR,CONVERT(MONEY, periodo3 ),1) as periodo3, CONVERT(VARCHAR,CONVERT(MONEY, periodo2 ),1) as periodo2, CONVERT(VARCHAR,CONVERT(MONEY, periodo1 ),1) as periodo1 from _TEMP_Cartera where RazonSocial='"& rsUpdateEntry("RazonSocial") &"' AND CardCode='"& rsUpdateEntry("CardCode")  &"' "
      'response.write strSQL
         
      rsUpdateEntry5.Open strSQL, adoCon4
    
      if not rsUpdateEntry5.EOF then
            response.write("<td class='fontmed td-r'>"& rsUpdateEntry5("Limite") &"</td>")
            response.write("<td class='fontmed td-r'>"& rsUpdateEntry5("Dias") &"</td>")
            response.write("<td class='fontmed td-r strong'><a href='ShowContent.asp?contenido=81&RS=DMX&SN="& rsUpdateEntry("CardCode") &"'>"& rsUpdateEntry5("total") &"</a></td>")            
            response.write("<td class='fontmed td-r'><font color=red>"& rsUpdateEntry5("vencido") &"</font></td>")
            response.write("<td class='fontmed td-r'>"& rsUpdateEntry5("normal") &"</td>")
            response.write("<td class='fontmed td-r'>"& rsUpdateEntry5("refacturacion") &"</td>")
            response.write("<td class='fontmed td-r'>"& rsUpdateEntry5("periodo4") &"</td>")
            response.write("<td class='fontmed td-r'>"& rsUpdateEntry5("periodo3") &"</td>")
            response.write("<td class='fontmed td-r'>"& rsUpdateEntry5("periodo2") &"</td>")
            response.write("<td class='fontmed td-r'>"& rsUpdateEntry5("periodo1") &"</td>")
      else
            for i=1 to 8
                 response.write("<td class='fontmed td-r'>0</td>")
            next
      end if
      rsUpdateEntry5.close


      response.write ("</tr>")  
      rsUpdateEntry.MoveNext
  wend

if request("RS")="" then

  separador
  while not rsUpdateEntry2.EOF  'imprimiendo DFW'
      
      response.write ("<tr>")  
      response.write ("<td class='td-c fontmed strong'>"& rsUpdateEntry2("RS") &"</td>")    
      response.write ("<td class='td-l fontmed'>"& mid(rsUpdateEntry2("SocioNegocio"),1,30) &"</td>")   
     
      
      response.write ("<td class='td-r'><a href='ShowContent.asp?contenido=81&RS=DFW&SN="& rsUpdateEntry2("CardCode") &"'>"& rsUpdateEntry2("subtotal") &"</a></td>")  
      vTotales= vTotales  &"+" & trim( rsUpdateEntry2("subtotal") )
      vTotales=replace(vTotales,"++","+")
      vTotales=replace(vTotales,",","")
  

       strSQL="select CONVERT(VARCHAR,CONVERT(MONEY, Limite  ),1) as Limite,ExtraDays as Dias,CONVERT(VARCHAR,CONVERT(MONEY, total ),1) as total, CONVERT(VARCHAR,CONVERT(MONEY, vencido ),1) as vencido, CONVERT(VARCHAR,CONVERT(MONEY, normal ),1) as normal, CONVERT(VARCHAR,CONVERT(MONEY, refacturacion ),1) as refacturacion, CONVERT(VARCHAR,CONVERT(MONEY, periodo4 ),1) as periodo4, CONVERT(VARCHAR,CONVERT(MONEY, periodo3 ),1) as periodo3, CONVERT(VARCHAR,CONVERT(MONEY, periodo2 ),1) as periodo2, CONVERT(VARCHAR,CONVERT(MONEY, periodo1 ),1) as periodo1 from _TEMP_Cartera where RazonSocial='"& rsUpdateEntry2("RS") &"' AND CardCode='"& rsUpdateEntry2("CardCode")  &"' "
      'response.write strSQL
         
      rsUpdateEntry5.Open strSQL, adoCon4
    
      if not rsUpdateEntry5.EOF then
            response.write("<td class='fontmed td-r'>"& rsUpdateEntry5("Limite") &"</td>")
            response.write("<td class='fontmed td-r'>"& rsUpdateEntry5("Dias") &"</td>")
            response.write("<td class='fontmed td-r strong'><a href='ShowContent.asp?contenido=81&RS=DFW&SN="& rsUpdateEntry2("CardCode") &"'>"& rsUpdateEntry5("total") &"</a></td>")
            response.write("<td class='fontmed td-r'><font color=red>"& rsUpdateEntry5("vencido") &"</font></td>")
            response.write("<td class='fontmed td-r'>"& rsUpdateEntry5("normal") &"</td>")
            response.write("<td class='fontmed td-r'>"& rsUpdateEntry5("refacturacion") &"</td>")
            response.write("<td class='fontmed td-r'>"& rsUpdateEntry5("periodo4") &"</td>")
            response.write("<td class='fontmed td-r'>"& rsUpdateEntry5("periodo3") &"</td>")
            response.write("<td class='fontmed td-r'>"& rsUpdateEntry5("periodo2") &"</td>")
            response.write("<td class='fontmed td-r'>"& rsUpdateEntry5("periodo1") &"</td>")
      else
            for i=1 to 8
                 response.write("<td class='fontmed td-r'>0</td>")
            next
      end if
      rsUpdateEntry5.close


      rsUpdateEntry2.MoveNext

  wend

  rsUpdateEntry2.close
end if

  rsUpdateEntry.close
  


  response.write("<tr><td colspan=4>&nbsp;</td></tr>")
  vTotales=replace(vTotales,",","")

  vTotales="select CONVERT(VARCHAR,CONVERT(MONEY,  sum("&vTotales &")    ),1)    as calculo"
  'response.write vTotales
  rsUpdateEntry3.Open vTotales, adoCon4
 
  response.write("<tr><td colspan=2 class='td-r fontsmall strong'>[VENTAS DEL MES] "& rsUpdateEntry4("calculo") &" &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;versus </td><td class='td-r Fontsmall strong'>"& rsUpdateEntry3("calculo") &" PDTE</td></tr>")
  rsUpdateEntry3.close
   
   rsUpdateEntry4.close

  closeTable

  if request("RS")="" then

  titulo="Pendiente suministro clientes al contado"
  doTitulo(titulo)

  strSQL="SELECT 'DMX' as RS,a.CardCode,a.CardName,round(sum(T1.OpenQty*T1.Price),2) as subtotal,CONVERT(VARCHAR,CONVERT(MONEY, sum( round(T1.OpenQty*T1.Price,2) )),1) as subtotal2 FROM PRODUCTIVA_DMX.dbo.OCRD a inner join PRODUCTIVA_DMX.dbo.ORDR T0 WITH (NOLOCK) on  a.DebtLine=0 and  substring(a.Cardcode,1,1)='C' and a.CardCode=T0.CardCode and T0.CANCELED='N' and T0.DocStatus='O' inner join PRODUCTIVA_DMX.dbo.RDR1 T1 WITH (NOLOCK) on T1.LineStatus='O' and T0.DocEntry=T1.DocEntry  where a.cardname not like '%deltaflow%' and a.cardname not like '%meide%'  group by a.CardCode,a.CardName UNION SELECT 'DFW' as RS,a.CardCode,a.CardName,round(sum(T1.OpenQty*T1.Price),2) as subtotal,CONVERT(VARCHAR,CONVERT(MONEY, sum( round(T1.OpenQty*T1.Price,2) )),1) as subtotal2 FROM PRODUCTIVA_DFW.dbo.OCRD a inner join PRODUCTIVA_DFW.dbo.ORDR T0 WITH (NOLOCK) on  a.DebtLine=0 and  substring(a.Cardcode,1,1)='C' and a.CardCode=T0.CardCode and T0.CANCELED='N' and T0.DocStatus='O' inner join PRODUCTIVA_DFW.dbo.RDR1 T1 WITH (NOLOCK) on T1.LineStatus='O' and T0.DocEntry=T1.DocEntry  where a.cardname not like '%deltaflow%' and a.cardname not like '%meide%'  group by a.CardCode,a.CardName UNION SELECT 'MME' as RS,a.CardCode,a.CardName,round(sum(T1.OpenQty*T1.Price),2) as subtotal,CONVERT(VARCHAR,CONVERT(MONEY, sum( round(T1.OpenQty*T1.Price,2) )),1) as subtotal2 FROM PRODUCTIVA_MEIDE.dbo.OCRD a inner join PRODUCTIVA_MEIDE.dbo.ORDR T0 WITH (NOLOCK) on  a.DebtLine=0 and  substring(a.Cardcode,1,1)='C' and a.CardCode=T0.CardCode and T0.CANCELED='N' and T0.DocStatus='O' inner join PRODUCTIVA_MEIDE.dbo.RDR1 T1 WITH (NOLOCK) on T1.LineStatus='O' and T0.DocEntry=T1.DocEntry  where a.cardname not like '%deltaflow%' and a.cardname not like '%meide%'  group by a.CardCode,a.CardName "   
       
  'response.write(strSQL&"<br><br>")
  rsUpdateEntry.Open strSQL, adoCon4  

  CreateTable(900)
  rowin
  response.write("<td class='subtitulo td-c'>RS</td><td class='subtitulo td-c'>Socio Negocio</td><td class='subtitulo td-c'>Monto a suministrar </td>")
  response.write("<td class='subtitulo td-l strong' colspan=7>Pedidos abiertos</td>")
 
  rowout
  strSUM="select sum("

  while not rsUpdateEntry.EOF
      rowin
      response.write("<td class='td-c fonttiny'>"&rsUpdateEntry("RS")&"</td>")
      response.write("<td class='td-l fonttiny'>"& mid(rsUpdateEntry("cardname"),1,45) &"</td>")
      strSUM=strSUM & "+" & trim(rsUpdateEntry("subtotal"))
      response.write("<td class='td-r fonttiny strong'>"& rsUpdateEntry("subtotal2") &"</td>")

      strSQL="SELECT distinct(T0.Docnum) as Docnum  FROM ORDR T0 WITH (NOLOCK) where T0.CardCode='"& rsUpdateEntry("cardCode")  &"' and T0.CANCELED='N' and T0.DocStatus='O' "   
      'response.write strSQL
      select case rsUpdateEntry("RS")
      case "DMX" rsUpdateEntry2.Open strSQL, adoCon2   'DMX'
      case "DFW" rsUpdateEntry2.Open strSQL, adoCon3   'DFW'
      end select

      i=0
      while not rsUpdateEntry2.EOF
          response.write("<td class='td-c fonttiny' width='40px'><a href='ShowContent.asp?contenido=68&action=1&fRS="&rsUpdateEntry("RS")&"&fpedido="& rsUpdateEntry2("docnum") &"' target='pedido'>"& rsUpdateEntry2("docnum") &"</a></td>")
          rsUpdateEntry2.MoveNext
          i=i+1
      wend
      for n=i to 5
          response.write("<td class='td-c fonttiny' width='40px'>&nbsp;</td>")
      next
      rsUpdateEntry2.close
      RowOut
      rsUpdateEntry.MoveNext
  wend
  rsUpdateEntry.close
  strSUM=strSUM & ") as calculo"
  rsUpdateEntry4.Open strSUM, adoCon4
  separador
  'response.write strSUM
  RowIn
        response.write("<td colspan=2 class='td-r strong'>TOTAL</td>")
        response.write("<td class='td-r fonttiny strong'>"& FormatCurrency(rsUpdateEntry4("calculo")) &"</td>")
        rsUpdateEntry4.close
  rowout
  
  closeTable

  end if
  response.write("<P>&nbsp;</P>")



  response.write("<div name='detalle' id='detalle'>&nbsp;</div>")  

 

  if request("RS")<>"" then
   

    titulo="Estatus del cliente: " &vSocioNegocio &" ("&request("RS")&")"
    doTitulo(titulo)

    CreateTable(1100)  'tabla total'
      response.write("<tr><td width='25%' class='td-c' style='vertical-align:top'>")
      CreateTable(260)

      strSQL="SELECT T3.ItmsGrpNam,CONVERT(VARCHAR,CONVERT(MONEY, sum( round(T1.OpenQty*T1.Price,2) )  ),1) as subtotal  FROM ORDR T0 WITH (NOLOCK) inner join RDR1 T1 WITH (NOLOCK) on T0.DocEntry=T1.DocEntry   inner join OITM T2 WITH (NOLOCK) on T1.ItemCode=T2.ItemCode   inner join OITB T3 WITH (NOLOCK) on T2.ItmsGrpCod=T3.ItmsGrpCod    WHERE T0.CANCELED='N' and T0.DocStatus='O' and T1.LineStatus='O' and T0.cardcode='"& request("SN") &"'     group by T3.ItmsGrpNam order by ItmsGrpNam "

      'response.write(strSQL&"<br>")
      select case request("RS")
      case "DMX" rsUpdateEntry4.Open strSQL, adoCon2   'DMX'
      case "DFW" rsUpdateEntry4.Open strSQL, adoCon3   'DFW'
      end select
      

              strSQL="SELECT CONVERT(VARCHAR,CONVERT(MONEY, sum( round(T1.OpenQty*T1.Price,2) )  ),1) as calculo  FROM ORDR T0 WITH (NOLOCK) inner join RDR1 T1 WITH (NOLOCK) on T0.DocEntry=T1.DocEntry   inner join OITM T2 WITH (NOLOCK) on T1.ItemCode=T2.ItemCode   inner join OITB T3 WITH (NOLOCK) on T2.ItmsGrpCod=T3.ItmsGrpCod    WHERE T0.CANCELED='N' and T0.DocStatus='O' and T1.LineStatus='O' and T0.cardcode='"& request("SN") &"'     "

              'response.write strSQL
              select case request("RS")
		      case "DMX" rsUpdateEntry5.Open strSQL, adoCon2   'DMX'
		      case "DFW" rsUpdateEntry5.Open strSQL, adoCon3   'DFW'
		      end select

            


      response.write("<tr><td colspan=2 class='td-c subtitulo'>")  %>
      <a href="Javascript:sendReq('DetalleCliente.asp','SN,action','<%=request("SN")%>,1','ShowData')" class="strong">PENDIENTE POR SUMINISTRAR </a></td></tr>  <%

      response.write("<tr><td width='50%' class='td-r subtitulo'>Familia</td><td width='50%' class='td-r subtitulo'>Suministro</td></tr>")


         while not rsUpdateEntry4.EOF
           response.write("<tr>")
           response.write("<td class='td-r'>"&  rsUpdateEntry4("ItmsGrpNam")   &"</td>")
           response.write("<td class='td-r'>"&  rsUpdateEntry4("subTotal")   &"</td>")
           response.write("</tr>")
           rsUpdateEntry4.MoveNext
        wend
        rsUpdateEntry4.close
        response.write("<tr><td class='td-c' colspan=2><HR></td></tr>")
        response.write("<tr><td class='td-r strong'>subtotal</td><td class='td-r strong'>"&  rsUpdateEntry5("calculo")   &"</td></tr>")
        rsUpdateEntry5.close

        closeTable 

        response.write("</td><td width='25%' class='td-c' style='vertical-align:top'>")
        CreateTable(260)

        response.write("<tr><td colspan=2 class='td-c subtitulo'>")  %>
        <a href="Javascript:sendReq('DetalleCliente.asp','SN,action','<%=request("SN")%>,2','ShowData')" class="strong">REMISIONADO POR FACTURAR </a></td></tr>  <%
        
        response.write("<tr><td width='50%' class='td-r subtitulo'>Familia</td><td width='50%' class='td-r subtitulo'>Suministro</td></tr>")

        strSQL="SELECT T3.ItmsGrpNam,CONVERT(VARCHAR,CONVERT(MONEY, sum( round(T1.OpenQty*T1.Price,2) )  ),1) as subtotal  FROM ODLN T0 WITH (NOLOCK) inner join DLN1 T1 WITH (NOLOCK) on T0.DocEntry=T1.DocEntry   inner join OITM T2 WITH (NOLOCK) on T1.ItemCode=T2.ItemCode   inner join OITB T3 WITH (NOLOCK) on T2.ItmsGrpCod=T3.ItmsGrpCod    WHERE T0.CANCELED='N' and T0.DocStatus='O' and T1.LineStatus='O' and T0.cardcode='"& request("SN") &"'    group by T3.ItmsGrpNam order by ItmsGrpNam   "

          'response.write strSQL
              select case request("RS")
		      case "DMX" rsUpdateEntry4.Open strSQL, adoCon2   'DMX'
		      case "DFW" rsUpdateEntry4.Open strSQL, adoCon3   'DFW'
		      end select
        

        strSQL="SELECT CONVERT(VARCHAR,CONVERT(MONEY, sum( round(.16*T1.OpenQty*T1.Price,2) )  ),1) as IVA  FROM ODLN T0 WITH (NOLOCK) inner join DLN1 T1 WITH (NOLOCK) on T0.DocEntry=T1.DocEntry   inner join OITM T2 WITH (NOLOCK) on T1.ItemCode=T2.ItemCode   inner join dbo.OITB T3 WITH (NOLOCK) on T2.ItmsGrpCod=T3.ItmsGrpCod    WHERE T0.CANCELED='N' and T0.DocStatus='O' and T1.LineStatus='O' and T0.cardcode='"& request("SN") &"'  "
          'response.write strSQL
              select case request("RS")
		      case "DMX" rsUpdateEntry5.Open strSQL, adoCon2   'DMX'
		      case "DFW" rsUpdateEntry5.Open strSQL, adoCon3   'DFW'
		      end select

       

        strSQL="SELECT CONVERT(VARCHAR,CONVERT(MONEY, sum( round(1.16*T1.OpenQty*T1.Price,2) )  ),1) as calculo  FROM ODLN T0 WITH (NOLOCK) inner join DLN1 T1 WITH (NOLOCK) on T0.DocEntry=T1.DocEntry   inner join OITM T2 WITH (NOLOCK) on T1.ItemCode=T2.ItemCode   inner join OITB T3 WITH (NOLOCK) on T2.ItmsGrpCod=T3.ItmsGrpCod    WHERE T0.CANCELED='N' and T0.DocStatus='O' and T1.LineStatus='O' and T0.cardcode='"& request("SN") &"'  "
          'response.write strSQL
              select case request("RS")
		      case "DMX" rsUpdateEntry6.Open strSQL, adoCon2   'DMX'
		      case "DFW" rsUpdateEntry6.Open strSQL, adoCon3   'DFW'
		      end select
        


        while not rsUpdateEntry4.EOF
           response.write("<tr>")
           response.write("<td class='td-r'>"&  rsUpdateEntry4("ItmsGrpNam")   &"</td>")
           response.write("<td class='td-r'>"&  rsUpdateEntry4("subTotal")   &"</td>")
           response.write("</tr>")
           rsUpdateEntry4.MoveNext
        wend
        rsUpdateEntry4.close
        
        response.write("<tr><td class='td-r'>IVA</td><td class='td-r'>"&  rsUpdateEntry5("IVA")   &"</td></tr>")
        response.write("<tr><td class='td-c' colspan=2><HR></td></tr>")
        response.write("<tr><td class='td-r strong'>Total</td><td class='td-r strong'>"&  rsUpdateEntry6("calculo")   &"</td></tr>")
        rsUpdateEntry5.close
        rsUpdateEntry6.close
        closeTable 

        response.write("</td><td width='25%' class='td-c' style='vertical-align:top'>")
         CreateTable(260)
         response.write("<tr><td colspan=2 class='td-c subtitulo'>TOTAL A CREDITO</td></tr>")
         response.write("<tr><td width='50%' class='td-r subtitulo'>Familia</td><td width='50%' class='td-r subtitulo'>Subtotal</td></tr>")


          strSQL="select e.ItmsGrpNam,CONVERT(VARCHAR,CONVERT(MONEY,   sum( round( c.Quantity*c.Price*a.factor,2))  ),1)  as subtotal     from SBOTemp.dbo._TEMP_COBRA a inner join OINV b on a.DocNum=b.docnum inner join INV1 c on b.docentry=c.docentry    inner join OITM d on c.ItemCode=d.ItemCode inner join OITB e on d.ItmsGrpCod=e.ItmsGrpCod    where a.RazonSocial='"&request("RS")&"' and a.CardCode='"& request("SN") &"'      group by e.ItmsGrpNam   order by  e.ItmsGrpNam "

          'response.write strSQL
              select case request("RS")
		      case "DMX" rsUpdateEntry2.Open strSQL, adoCon2   'DMX'
		      case "DFW" rsUpdateEntry2.Open strSQL, adoCon3   'DFW'
		      end select

             


          strSQL="select CONVERT(VARCHAR,CONVERT(MONEY,    sum( round(.16*c.Quantity*c.Price*a.factor,2))    ),1)    as IVA    from SBOTemp.dbo._TEMP_COBRA a inner join OINV b on a.DocNum=b.docnum inner join INV1 c on b.docentry=c.docentry      where a.RazonSocial='"&request("RS")&"' and a.CardCode='"& request("SN") &"'  "
              select case request("RS")
		      case "DMX" rsUpdateEntry3.Open strSQL, adoCon2   'DMX'
		      case "DFW" rsUpdateEntry3.Open strSQL, adoCon3   'DFW'
		      end select

   
          strSQL="select CONVERT(VARCHAR,CONVERT(MONEY,    sum(total)    ),1)    as calculo  from _TEMP_COBRA  where RazonSocial='"&request("RS")&"' and CardCode='"& request("SN") &"'  "
          rsUpdateEntry4.Open strSQL, adoCon4   'Temp'

          
         
         while not rsUpdateEntry2.EOF
           response.write("<tr>")
           response.write("<td class='td-r'>"&  rsUpdateEntry2("ItmsGrpNam")   &"</td>")
           response.write("<td class='td-r'>"&  rsUpdateEntry2("subTotal")   &"</td>")
           response.write("</tr>")
           rsUpdateEntry2.MoveNext
        wend
        rsUpdateEntry2.close        

        response.write("<tr><td class='td-r'>IVA</td><td class='td-r'>"&  rsUpdateEntry3("IVA") &"</td></tr>")
        rsUpdateEntry3.close

        response.write("<tr><td class='td-c' colspan=2><HR></td></tr>")
        response.write("<tr><td class='td-r strong'>total</td><td class='td-r strong'>"&  rsUpdateEntry4("calculo")   &"</td></tr>")
        rsUpdateEntry4.close

        closeTable

        response.write("</td><td width='25%' class='td-c' style='vertical-align:top'>")
        CreateTable(260)
        

         strSQL="select e.ItmsGrpNam,CONVERT(VARCHAR,CONVERT(MONEY,  sum( round( c.Quantity*c.Price*a.factor,2)  )    ),1)  as Total    from SBOTemp.dbo._TEMP_COBRA a inner join OINV b on a.DocNum=b.DocNum inner join INV1 c on b.DocEntry=c.DocEntry    inner join OITM d on c.ItemCode=d.ItemCode inner join OITB e on d.ItmsGrpCod=e.ItmsGrpCod    where a.RazonSocial='"&request("RS")&"' and a.CardCode='"& request("SN") &"' and a.DiasAtraso>=1    group by e.ItmsGrpNam order by e.ItmsGrpNam   "
         'response.write strSQL

          select case request("RS")
	      case "DMX" rsUpdateEntry2.Open strSQL, adoCon2   'DMX'
	      case "DFW" rsUpdateEntry2.Open strSQL, adoCon3   'DFW'
	      end select

        

         strSQL="select CONVERT(VARCHAR,CONVERT(MONEY,  sum(b.VatSumFC-b.VatPaidFC)    ),1)  as IVA     from SBOTemp.dbo._TEMP_COBRA a inner join OINV b on a.docnum=b.docnum    where a.RazonSocial='"&request("RS")&"' and a.CardCode='"& request("SN") &"' and a.DiasAtraso>=1 "
         'response.write strSQL
          select case request("RS")
	      case "DMX" rsUpdateEntry3.Open strSQL, adoCon2   'DMX'
	      case "DFW" rsUpdateEntry3.Open strSQL, adoCon3   'DFW'
	      end select
        

           strSQL="select CONVERT(VARCHAR,CONVERT(MONEY,    sum(total)    ),1)    as calculo  from _TEMP_COBRA  where RazonSocial='"&request("RS")&"' and CardCode='"& request("SN") &"'  and DiasAtraso>=1 "
          rsUpdateEntry4.Open strSQL, adoCon4   'Temp'

        response.write("<tr><td colspan=2 class='td-c subtitulo'>")  %>
        <a href="Javascript:sendReq('DetalleCliente.asp','SN,action','<%=request("SN")%>,3','ShowData')" class="strong">CREDITO VENCIDO </a></td></tr>  <%
        
         response.write("<tr><td width='50%' class='td-r subtitulo'>Familia</td><td width='50%' class='td-r subtitulo'>Total</td></tr>")

         while not rsUpdateEntry2.EOF
           response.write("<tr>")
           response.write("<td class='td-r'>"&  rsUpdateEntry2("ItmsGrpNam")   &"</td>")
           response.write("<td class='td-r'><font color=red>"&  rsUpdateEntry2("Total")   &"</font></td>")
           response.write("</tr>")
           rsUpdateEntry2.MoveNext
        wend
        rsUpdateEntry2.close        

        response.write("<tr><td class='td-r'>IVA</td><td class='td-r'><font color=red>"&  rsUpdateEntry3("IVA") &"</td></tr>")
        rsUpdateEntry3.close

        response.write("<tr><td class='td-c' colspan=2><HR></td></tr>")
        response.write("<tr><td class='td-r strong'>total</td><td class='td-r strong'><font color=red>"&  rsUpdateEntry4("calculo")   &"</font></td></tr>")
        rsUpdateEntry4.close

        closeTable


        response.write("</td><tr>")
        closeTable

        response.write("<div id='ShowData'>&nbsp;</div>")


  end if


  
  'strSQL="if exists (select * from INFORMATION_SCHEMA.TABLES where TABLE_NAME = '_TEMP_Cartera' AND TABLE_SCHEMA = 'dbo')     drop table dbo._TEMP_Cartera;"  
  'rsUpdateEntry4.Open strSQL, adoCon4

  'strSQL="if exists (select * from INFORMATION_SCHEMA.TABLES where TABLE_NAME = '_TEMP_COBRA' AND TABLE_SCHEMA = 'dbo')     drop table dbo._TEMP_COBRA;"  
  'rsUpdateEntry5.Open strSQL, adoCon4


end sub






sub Coti_NoPartes  'contenido 82   

    doAlert
    titulo="Listar no. partes proveedor por cotizacion de venta"
    DoTitulo(titulo)


    if request("action")="" then

         response.write("<form id='inv' method='POST' action='ShowContent.asp'>")
         response.write("<input type='hidden' name='contenido' value=82>")
         response.write("<input type='hidden' name='action' value=1>")

         response.write("<P class='td-c'>Raz&oacute;n Social: <select name='fRS'><option value='DMX'>DMX</option><option value='DFW'>DFW</option></select>")
         response.write("&nbsp;&nbsp;&nbsp;Cotizacion: <input class='anchoSmall' type='numeric' name='fcoti' value="&request("fcoti") &"></P>")

         response.write("<P class='td-c'>&nbsp;</P><P class='td-c'><input type='submit' value='continuar'></form></P>")


    else


 'response.write("<P class='td-c'>&nbsp;</P>")

         strSQL="select a.docdate,a.CardName,b.LineNum,b.ItemCode,b.Dscription,b.Quantity,d.ItmsGrpNam,c.SuppCatNum from OQUT a inner join QUT1 b on a.docentry=b.docentry left join OITM c on b.ItemCode=c.ItemCode left join OITB d on c.ItmsGrpCod=d.ItmsGrpCod where a.docnum="& request("fcoti") &" order by b.LineNum"
         'response.write strSQL

         if request("fRS")="DMX" then
                 rsUpdateEntry.Open strSQL, adoCon2 'DMX'
         else
                 rsUpdateEntry.Open strSQL, adoCon3 'DFW'
         end if
      
         response.write("<table border=1 width='550px' cellpadding=3 cellspacing=2>")

        if not rsUpdateEntry.EOF then
        
         response.write("<tr><td class='subtitulo td-r'>Razon social:</td><td class='td-l'>"& request("fRS") &"</td></tr>")
         response.write("<tr><td class='subtitulo td-r'>Cotizacion:</td><td class='td-l'>"& request("fcoti") &"</td></tr>")
         response.write("<tr><td class='subtitulo td-r'>Socio negocio:</td><td class='td-l'>"& rsUpdateEntry("cardname") &"</td></tr>")

         else
               response.write("<tr><td colspan=8>no existen pedidos con esos criterios</td></tr>")

         end if

         response.write("</table> <P>&nbsp;</P>")

         CreateTable(840)

         
         response.write("<tr><td class='subtitulo td-c'>#</td><td class='subtitulo td-c'>codigo</td><td class='subtitulo td-c'>descripcion</td><td class='subtitulo td-c'>cantidad</td><td class='subtitulo td-c'>proveedor</td><td class='subtitulo td-c'>numero parte</td></tr>")
          
         n=1 
         while not rsUpdateEntry.EOF
             if n=1 then
                  vString=" style='background-color: white;' "
             else
                  vString=" style='background-color: #E1DFDF;' "
             end if

             response.write("<tr>")
             response.write("<td class='td-c' "& vString &">" & rsUpdateEntry("lineNum")+1 &"</td>")
             response.write("<td class='td-l fontmed' "& vString &">" & rsUpdateEntry("ItemCode") &"</td>")
             response.write("<td class='td-l fontmed' "& vString &">" & left(rsUpdateEntry("Dscription"),50) &"</td>")             
             response.write("<td class='td-r' "& vString &">" & rsUpdateEntry("Quantity") &"</td>")
             response.write("<td class='td-r' "& vString &">" & rsUpdateEntry("ItmsGrpNam") &"</td>")
             response.write("<td class='td-r' "& vString &">" & rsUpdateEntry("SuppCatNum") &"</td>")


             response.write("</tr>")
             rsUpdateEntry.MoveNext
             n=n+1
             if n=3 then
                 n=1
             end if
         wend
         rsUpdateEntry.close
         closeTable

    end if

end sub




sub InformeClientesNuevos   'contenido 83'

     titulo="HISTORIAL DE REGISTRO CLIENTES NUEVOS"
     DoTitulo(TITULO)

     strSQL="select a.RazonSocial,a.CardCode,Cardname=( select case a.RazonSocial   when 'DMX' then (select CardName from PRODUCTIVA_DMX.dbo.OCRD where CardCode=a.CardCode)   when 'DFW' then (select CardName from PRODUCTIVA_DFW.dbo.OCRD where CardCode=a.CardCode)    when 'MME' then (select CardName from PRODUCTIVA_MEIDE.dbo.OCRD where CardCode=a.CardCode) end ), LogDate Fecha,month(a.LogDate) as Mes,asesor=( select case a.RazonSocial 		       when 'DMX' then (select y.SlpName from PRODUCTIVA_DMX.dbo.OCRD x inner join PRODUCTIVA_DMX.dbo.OSLP y on x.SlpCode=y.SlpCode  where CardCode=a.CardCode)  when 'DFW' then (select y.SlpName from PRODUCTIVA_DFW.dbo.OCRD x inner join PRODUCTIVA_DFW.dbo.OSLP y on x.SlpCode=y.SlpCode  where CardCode=a.CardCode) when 'MME' then (select y.SlpName from PRODUCTIVA_MEIDE.dbo.OCRD x inner join PRODUCTIVA_MEIDE.dbo.OSLP y on x.SlpCode=y.SlpCode  where CardCode=a.CardCode) end ),contacto=(select ID from INTRANET.dbo.Contactos where CardCode collate Modern_Spanish_CI_AS =a.CardCode collate Modern_Spanish_CI_AS )   from Repositorio a   where year(a.logdate)>=2021 order by a.logdate desc  "
     'response.write strSQL
     rsUpdateEntry.Open strSQL, adoCon4 'SBO'

     'CreateTable(740)

     response.write("<table width='1000px' border=0 cellspacing=2 cellpadding=5 class='table2excel table2excel_with_colors' data-tableName='table1' >")  
     response.write("<tr>")
     response.write("<td class='td-c subtitulo fonttiny'>#</td>")
      response.write("<td class='td-c subtitulo fonttiny'>RS</td>")
     response.write("<td class='td-c subtitulo fonttiny'>Codigo SN</td>")
     response.write("<td class='td-c subtitulo fonttiny' width='270px'>Socio negocio</td>")
     response.write("<td class='td-c subtitulo fonttiny' width='210px'>Primer pedido</td>")
     response.write("<td class='td-c subtitulo fonttiny' width='210px'>Asesor</td>")
     response.write("<td class='td-c subtitulo fonttiny'># Contacto <br>Marketing</td>")
     response.write("</tr>")

     vMes=0
     i=1
     while not rsUpdateEntry.EOF
          vMes=rsUpdateEntry("mes")
          response.write("<tr>")
          response.write("<td class='td-l fonttinyfonttiny'>"& i &"</td>")
          response.write("<td class='td-c fonttiny'>"& rsUpdateEntry("RazonSocial") &"</td>")
          response.write("<td class='td-c fonttiny'>"& rsUpdateEntry("CardCode") &"</td>")
          response.write("<td class='td-l fonttiny'>"& mid(rsUpdateEntry("CardName"),1,34) &"</td>")
          response.write("<td class='td-c fonttiny'>"& rsUpdateEntry("fecha") &"</td>")
          response.write("<td class='td-l fonttiny'>"& mid(rsUpdateEntry("asesor"),1,30) &"</td>")

          if trim(rsUpdateEntry("contacto"))<>"" then
              response.write("<td class='td-l fontsmall'><a href='http://192.168.0.204/modules/ShowContent.asp?contenido=72&action=2&ID="& trim(rsUpdateEntry("contacto")) &"' target='contacto'>"& rsUpdateEntry("contacto") &"</a></td>")
          else
               response.write("<td class='td-l fontsmall'>"& rsUpdateEntry("contacto") &"</td>")
          end if
          rsUpdateEntry.movenext
          i=i+1
          if not rsUpdateEntry.EOF then
                if vMes<>rsUpdateEntry("mes") then
                  separador   
                  i=1               
                end if
          end if
     wend
     rsUpdateEntry.close
     %>
     </table><button class="exportToExcel">Exportar a excel</button>  &nbsp;&nbsp;&nbsp; <button onclick="go_display_div('#titulo')"> ir arriba   </button>     

         
    <script>
      $(function() {
        $(".exportToExcel").click(function(e){
          var table = $(this).prev('.table2excel');
          if(table && table.length){
            var preserveColors = (table.hasClass('table2excel_with_colors') ? true : false);
            $(table).table2excel({
              exclude: ".noExl",
              name: "Excel Document Name",
              filename: "myFileName" + new Date().toISOString().replace(/[\-\:\.]/g, "") + ".xls",
              fileext: ".xls",
              exclude_img: true,
              exclude_links: true,
              exclude_inputs: true,
              preserveColors: preserveColors
            });
          }
        });
        
      });
    </script>

    <%
    closeTable

end sub


sub InformeCotisNoClosed 'contenido 84'

       titulo="INFORME DE COTIZACION NO CONCLUIDAS"
       Dotitulo(TITULO)




end sub



sub AnalisisUtilidadPedido 'contenido 85' 

    response.write("<P>&nbsp;</P>")
    if request("msg")<>"" then
           response.write("<P class='alert td-c'>"&  request("msg") &"</P>")
    end if
   
    if request("fpedido")<>"" then
        titulo="An&aacute;lisis de la Utilidad del pedido " & request("fpedido")
    else
        titulo="An&aacute;lisis de la Utilidad por pedido"
    end if
    DoTitulo(titulo)
    response.write("<form id='utilidad' method='POST' action='ShowContent.asp'>")
    response.write("<input type='hidden' name='contenido' value=85>")


    if request("action")="" then
        
         response.write("<input type='hidden' name='action' value=1>")
         response.write("<P class='td-c'>Raz&oacute;n Social: <select name='fRS'><option value='DFW'>DFW</option><option value='DMX'>DMX</option><option value='MEIDE'>MME</option></select>")
         response.write("&nbsp;&nbsp;&nbsp;Pedido: <input class='anchoSmall' type='numeric' name='fpedido' value=" & request("fpedido") &"></P>")

         response.write("<P class='td-c'>&nbsp;</P><P class='td-c'><input type='submit' value='continuar'></form></P>")

    else
    
         response.write("<input type='hidden' name='action' value=2>")

         'QUERY PARA DATOS DEL PEDIDO'
         strSQL="select isnull(U_FleteTotal,0) as flete2,CONVERT(VARCHAR,CONVERT(MONEY,isnull(U_FleteTotal,0)),1) as flete,a.*,b.LogDate,c.SlpName,CONVERT(VARCHAR,CONVERT(MONEY,isnull(DocTotalSy-VatSumFc,0)),1) as subtotal,replace(a.Comments,'Basado en Ofertas de ventas ','COTI-') as notas "
 
         strSQL= strSQL &"from PRODUCTIVA_"& request("fRS") &".dbo.ORDR a inner join Notificaciones b on b.Modulo='ventas' and b.tipo='pedido' and a.DocNum=b.DocNum left join PRODUCTIVA_"& request("fRS") &".dbo.OSLP c on a.SlpCode=c.SlpCode where b.RazonSocial='"& request("fRS") &"' and a.DocNum="&  request("fpedido")
        
         'response.write strSQL
          rsUpdateEntry.Open strSQL, adoCon4 'SBOtemp'
         
      
         response.write("<table border=0 width='1200px' cellpadding=0 cellspacing=0 align='center'><td colspan=3 style='height:3px;font-size:3px;background-color:#CCCCCC'>&nbsp;</td></tr><tr><td class='td-c' width='30%' style='vertical-align:top'>")

        vflag=0
        if not rsUpdateEntry.EOF then

         response.write("<table border=1 width='95%' cellpadding=3 cellspacing=2 aling='center'>")
         response.write("<tr><td class='subtitulo td-r'>Razon social:</td><td class='td-l'>"& request("fRS") &"-" & rsUpdateEntry("CardCode") &"</td></tr>")
         response.write("<tr><td class='subtitulo td-r'>Referencia:</td><td class='td-l'>"& rsUpdateEntry("NumatCard") &"</td></tr>")         
         response.write("<tr><td class='subtitulo td-r'>Socio negocio:</td><td class='td-l'>"& rsUpdateEntry("cardname") &"</td></tr>")
         response.write("<tr><td class='subtitulo td-r'>Proyecto:</td><td class='td-l'>"& rsUpdateEntry("U_PROYECTO")  &"</td></tr>")         
           
         response.write("</table></td><td class='td-c' width='30%' style='vertical-align:top'>")  

         response.write("<table border=1 width='95%' cellpadding=3 cellspacing=2 aling='center'>")

         if rsUpdateEntry("DocStatus")="C" then
            response.write("<tr><td class='subtitulo td-r' width='110px'>Estatus:</td><td class='td-l strong'>SUMINISTRADO</td></tr>")
            vflag=1
         else
            response.write("<tr><td class='subtitulo td-r' width='110px'>Estatus:</td><td class='td-l strong'>SUMINISTRO PDTE</td></tr>")
         end if

         response.write("<tr><td class='subtitulo td-r'>SUBTOTAL SIN IVA:</td><td class='td-l strong'><font color=green>"& rsUpdateEntry("subtotal") &"</font></td></tr>")
         response.write("<tr><td class='subtitulo td-r'>Fecha de creaci&oacute;n:</td><td class='td-l'>"& rsUpdateEntry("LogDate") &" </td></tr>")
         response.write("<tr><td class='subtitulo td-r'>Asesor:</td><td class='td-l'>"& rsUpdateEntry("SlpName") &" </td></tr>")
        

         response.write("</table></td><td class='td-c' width='40%' style='vertical-align:top'>")  

         response.write("<table border=1 width='95%' cellpadding=3 cellspacing=2 aling='center'>")
         response.write("<tr><td class='subtitulo td-r'>Comentarios:</td><td class='td-l' rowspan=2>"&  rsUpdateEntry("notas")  &"</td></tr>")         
         response.write("<tr><td class='subtitulo td-r'>&nbsp;</td></tr>")
         response.write("<tr><td class='subtitulo td-r'>Moneda:</td><td class='td-l strong'>"& rsUpdateEntry("DocCur") &" - " & rsUpdateEntry("DocRate")  &" EN PEDIDO</td><tr><td colspan=2 class='td-r fonttiny strong'>* Antes de embarques, viaticos y combustibles</td></tr>")
         response.write("</table></td></tr><tr><td colspan=3 style='height:3px;font-size:3px;background-color:#CCCCCC'>&nbsp;</td></tr></table>")

        rsUpdateEntry.close
         '------------------------------------------------------------------------------------------------'

             '-- PRORRATEA VIATICOS'
         strSQL="exec SP_Prorratea_viaticos"
         rsUpdateEntry.Open strSQL, adoCon

         strSQL="exec [dbo].[SP_AnalisisUtilidad] "& request("fPedido") &",'"& request("fRS") &"','"& session("session_id") &"'"
         rsUpdateEntry.Open strSQL, adoCon4


         'QUERY PARA DETALLE DE LAS PARTIDAS Y SUS TOTALES'
         strSQL="select * from AnalisisUtilidad where RS='"& request("fRS") &"' and Pedido="&  request("fPedido") &" order by Linea"
         rsUpdateEntry2.Open strSQL, adoCon4

         response.write("<table border=1 width='1200px' cellpadding=3 cellspacing=2 aling='center'>")
         response.write("<tr><td class='td-c subtitulo'>#</td><td class='td-c subtitulo'>&nbsp;</td><td class='td-c subtitulo'>Codigo</td><td class='td-c subtitulo'>Item</td><td class='td-c subtitulo'>Almacen</td><td class='td-c subtitulo'>Cantidad</td><td class='td-c subtitulo'>Costo SAP<br>unitario</td><td class='td-c subtitulo'>Provision <br>Flete MXN</td><td class='td-c subtitulo'>Precio venta <br>unitario</td><td class='td-c subtitulo'>Subtotal utilidad <br>estimada</td><td class='td-c subtitulo'>% utilidad <br>estimada</td><td class='td-c subtitulo'>(MXN) Subtotal  <br>Ingresos</td><td class='td-c subtitulo'>(MXN) Subtotal  <br>Costo Contable</td><td class='td-c subtitulo'>(MXN) Utilidad  <br>Contable *</td><td class='td-c subtitulo'>% Utilidad  <br>contable *</td></tr>")

        

         while not rsUpdateEntry2.EOF             
             response.write("<tr>")
             response.write("<td class='td-r fontmed'>"& rsUpdateEntry2("linea") &"</td>" ) 
             response.write("<td class='td-r fontmed'>"& rsUpdateEntry2("LineaEstatus") &"</td>" )
             response.write("<td class='td-l fonttiny'>"& rsUpdateEntry2("ItemCode") &"</td>" ) 
             response.write("<td class='td-l fonttiny'>"& rsUpdateEntry2("Dscription") &"</td>" ) 
             response.write("<td class='td-c fonttiny'>"& rsUpdateEntry2("WhsCode") &"</td>" ) 
             response.write("<td class='td-r fonttiny'>"& rsUpdateEntry2("Quantity") &"</td>" ) 
             response.write("<td class='td-r fonttiny'>"& rsUpdateEntry2("U_PRECIOLISTA") &"</td>" )   
             response.write("<td class='td-r fonttiny'><font color=green>"& rsUpdateEntry2("ProvisionFleteMXN") &"</font></td>" )               
             response.write("<td class='td-r fonttiny'>"& rsUpdateEntry2("Price") &"</td>" )
             response.write("<td class='td-r fonttiny'>"& rsUpdateEntry2("SubUtilidadEstm") &"</td>" )            
             response.write("<td class='td-r fonttiny' style='background-color: #C6E0F0'>"& rsUpdateEntry2("PcrtUtilEstm") &" % </td>" )     
             'response.write("<td class='td-r fontmed'>"& rsUpdateEntry2("QtyRemisionada") &"</td>" )
             response.write("<td class='td-r fonttiny'>"& rsUpdateEntry2("SubtotalIngresosMXN") &"</td>" )
             response.write("<td class='td-r fonttiny'>"& rsUpdateEntry2("SubtotalCosto") &"</td>" )
             response.write("<td class='td-r fonttiny'>"& rsUpdateEntry2("UtilidadMXN") &"</td>" )
             response.write("<td class='td-r fonttiny' style='background-color: #C6E0F0'>"& rsUpdateEntry2("PcrtUtil") &"</td>" )
             response.write("</tr>")
             rsUpdateEntry2.MoveNext            

         wend
         separador
         rsUpdateEntry2.close

         
  if vflag=0 then
  	       response.write("<P class='alert td-c'>informacion adicional sera mostrada cuando el pedido este completamente suministrado</P>")
  end if  'si flag=1'


         strSQL="select CONVERT(VARCHAR,CONVERT(MONEY,isnull( subCostoSAP ,0)),1) as CostoSAP,CONVERT(VARCHAR,CONVERT(MONEY,isnull( ProvisionFleteMXN ,0)),1) as ProvisionFleteMXN,      CONVERT(VARCHAR,CONVERT(MONEY,isnull( SubIngresosEstm ,0)),1) as SubIngresosEstm,     CONVERT(VARCHAR,CONVERT(MONEY,isnull( SubUtilidadEstm ,0)),1) as SubUtilidadEstm,PcrtUtilEstm,       CONVERT(VARCHAR,CONVERT(MONEY,isnull( SubtotalCosto ,0)),1) as SubtotalCosto,    CONVERT(VARCHAR,CONVERT(MONEY,isnull( SubtotalIngresosMXN ,0)),1) as SubtotalIngresosMXN,     CONVERT(VARCHAR,CONVERT(MONEY,isnull( UtilidadMXN ,0)),1) as UtilidadMXN,PcrtUtil,    CONVERT(VARCHAR,CONVERT(MONEY,isnull( AjusteFlete ,0)),1) as AjusteFlete, CONVERT(VARCHAR,CONVERT(MONEY,isnull( Envios ,0)),1) as Envios, CONVERT(VARCHAR,CONVERT(MONEY,isnull( viaticos ,0)),1) as viaticos,   CONVERT(VARCHAR,CONVERT(MONEY,isnull( Combustibles ,0)),1) as Combustibles,  CONVERT(VARCHAR,CONVERT(MONEY,isnull( UtilidadAjustada ,0)),1) as UtilidadAjustada,PcrtUtilAjustada,AjusteFlete as Flete,SubtotalIngresosMXN as ingresos    from AnalisisUtilidadPedido where RS='"&request("fRS") &"' and Pedido="& request("fpedido")
         'response.write(strSQL&"<br>")
         rsUpdateEntry3.Open strSQL, adoCon4
    

           '-----------------------subtotales en la estimacion de venta'
         response.write("<tr><td colspan=6 class='td-r fonttiny'>Costo de venta en la cotizaci&oacute;n</td>")
         response.write("<td class='td-r fonttiny strong'>"& rsUpdateEntry3("CostoSAP") &"</td>")  
         response.write("<td class='td-r fonttiny'><font color=green> " & rsUpdateEntry3("ProvisionFleteMXN") &"<font></td>")   
         response.write("<td>&nbsp;</td>")
         response.write("<td class='td-r fonttiny strong'>"& rsUpdateEntry3("SubUtilidadEstm") &"</td>")
         response.write("<td class='td-r fonttiny strong' style='background-color: #C6E0F0'>"& left( trim(rsUpdateEntry3("PcrtUtilEstm")) , 5) &" %</td>")

         response.write("<td class='td-r fonttiny strong'>"& rsUpdateEntry3("SubtotalIngresosMXN") &"</td>")
         response.write("<td class='td-r fonttiny strong'>"& rsUpdateEntry3("subtotalCosto") &"</td>")
         response.write("<td class='td-r fonttiny strong'>"& rsUpdateEntry3("UtilidadMXN") &" * </td>")
         response.write("<td class='td-r fonttiny strong' style='background-color: #C6E0F0'>"& mid(rsUpdateEntry3("PcrtUtil"),1,5) &" % * </td></table>")



         '-----GASTOS EJERCIDOS---------------'
         
         response.write("<P style='font-size: 3px'>&nbsp;</P><table border=1 width='1100px' cellpadding=3 cellspacing=2 aling='center'>")
         response.write("<tr><td colspan=12 class='td-c'><H3>Gastos ejercidos por fletes / paqueter&iacute;as / vi&aacute;ticos </H3></td></tr>")


         response.write("<tr><td class='td-c subtitulo fontmed'>Tipo</td><td class='td-c subtitulo fontmed'>#</td><td class='td-c subtitulo fontmed' width='90px'>Fecha</td><td class='td-c subtitulo fontmed'>Remision</td><td class='td-c subtitulo fontmed'>Pedido</td><td class='td-c subtitulo fontmed'>Portador</td><td class='td-c subtitulo fontmed'>Guia</td><td class='td-c subtitulo fontmed'>Rastreo</td><td class='td-c subtitulo fontmed'>Cliente Paga^</td><td class='td-c subtitulo fontmed'>Consolidado</td><td class='td-c subtitulo fontmed'>Subtotal</td><td class='td-c subtitulo fontmed'>Factor</td></tr>")


           '---EMBARQUES EXTERNOS'
         strSQL="select e.ID,SBOTemp.dbo.fn_GetMonthName(e.docdate,'spanish') as fecha,e.DocNum as remision,e.pedido,carrier.portador,case when e.CargoCliente=0 then 'no' else 'si' end as CargoCliente,case when isnull(e.Consolidado,0)=0 then 'no' else 'si' end as Consolidado,CONVERT(VARCHAR,CONVERT(MONEY,subtotal),1) subtotal_MXN,subTotal,guia,rastreo,comentarios  from SBOTemp.dbo.envios e inner join SBOTemp.dbo.cat_portadores carrier on e.id_portador=carrier.id_portador where e.docnum in (select distinct(b.docnum) from DLN1 a inner join ODLN b on a.DocEntry=b.DocEntry where a.Baseref="& request("fpedido") &" and b.CANCELED='N')"   
         'response.write(strSQL&"<BR>")

         if request("fRS")="DMX" then            
                 rsUpdateEntry4.Open strSQL, adoCon2 'DMX'
         else                 
                 rsUpdateEntry4.Open strSQL, adoCon3 'DFW'
         end if

         
         while not rsUpdateEntry4.EOF
              response.write("<tr>")
              response.write("<td class='td-r fontmed'>externo</td>" )
              response.write("<td class='td-r fontmed'><a href='ShowContent.asp?contenido=90&action=10&ID="&rsUpdateEntry4("ID")&"' target='externo'>"& rsUpdateEntry4("ID") &"</a></td>" )
              response.write("<td class='td-r fonttiny'>"& rsUpdateEntry4("fecha") &"</td>")              
              response.write("<td class='td-r fonttiny'>"& rsUpdateEntry4("remision") &"</td>" )
              response.write("<td class='td-r fonttiny'>"& rsUpdateEntry4("pedido") &"</td>" )
              response.write("<td class='td-r fonttiny'>"& rsUpdateEntry4("portador") &"</td>" )
              response.write("<td class='td-r fonttiny'>"& rsUpdateEntry4("guia") &"</td>" )
              response.write("<td class='td-r fonttiny'>"& rsUpdateEntry4("rastreo") &"</td>" )

              response.write("<td class='td-r fontmed'>"& rsUpdateEntry4("CargoCliente") &"</td>" )

             
                  'FleteSUM=FleteSUM &"-"& trim(rsUpdateEntry4("subtotal") )
                   'response.write(FleteSUM&"<BR>")

                   
            
              response.write("<td class='td-r fontmed'>"& rsUpdateEntry4("Consolidado") &"</td>" )

              if rsUpdateEntry4("CargoCliente")="no" then
                 response.write("<td class='td-r fontmed'>$ <font color=red>-"& rsUpdateEntry4("subtotal_MXN") &"</font></td>" )
              else
                 response.write("<td class='td-r fontmed'>$ <font color=red>-"& rsUpdateEntry4("subtotal_MXN") &" ^</font></td>" )
              end if
              response.write("<td class='td-r fontmed'>1</td>" )  'FACTOR EN 1 PARA ENVIOS EXTERNOS'
              response.write("</tr>")
              rsUpdateEntry4.MoveNext           
         wend
         rsUpdateEntry4.close

          response.write("<tr><td colspan=11 class='td-r strong'>Gasto total por envios: </td><td class='td-r strong' style='color:red'>-"& rsUpdateEntry3("envios")  &"</td></tr>")
         

       
         'CONSULTA DE VIATICOS'
         strSQL5="select a.ID,.dbo.fn_GetMonthName(a.DocDate,'spanish') as Fecha,a.Beneficiario,b.Pedido,case when b.Consolidado=1 then 'si' else 'no' end as Consolidado,sum(b.ejercido) as Ejercido from VIATICOS_R1 a inner join ViaticoRemision b on a.ID=b.ID where b.Pedido="& request("fpedido") &" group by a.ID,a.Beneficiario,b.Pedido,b.Consolidado,a.DocDate"
         'response.write(strSQL5&"<BR>")
         rsUpdateEntry7.Open strSQL5, adoCon


          while not rsUpdateEntry7.EOF
              response.write("<tr>")
              response.write("<td class='td-r fontmed'>interno</td>" )
              response.write("<td class='td-r fontmed'>VIAT-<a href='ShowContent.asp?contenido=29&ID="& rsUpdateEntry7("ID") &"' target='viatico'>"& rsUpdateEntry7("ID") &"</a></td>" )
              response.write("<td class='td-r fontmed'>"& rsUpdateEntry7("fecha") &"</td>" )
               strSQL3="select STUFF( (select ','+convert(varchar,remision)+' ' from ViaticoRemision where ID="&  rsUpdateEntry7("ID")  &" FOR XML PATH ('')),1,1,'') as remisiones  "   
              'response.write( strSQL3 &"<br>")  
              rsUpdateEntry4.Open strSQL3, adoCon
              response.write("<td class='td-r fontmed'>"& rsUpdateEntry4("remisiones") &"</td>" )
              rsUpdateEntry4.close
              response.write("<td class='td-r fontmed'>"& rsUpdateEntry7("pedido") &"</td>" )
              response.write("<td class='td-r fontmed'>"& rsUpdateEntry7("beneficiario") &"</td>" )
              response.write("<td class='td-r fontmed'>-</td>" )
              response.write("<td class='td-r fontmed'>-</td>" )
              response.write("<td class='td-r fontmed'>-</td>" )
             
              response.write("<td class='td-r fontmed'>"& rsUpdateEntry7("Consolidado") &"</td>" )
             
              response.write("<td class='td-r fontmed'>$ <font color=red>-"& rsUpdateEntry7("ejercido") &"</font></td>" )
              response.write("<td class='td-r fontmed'>1</td>" )
              response.write("</tr>")
              rsUpdateEntry7.MoveNext           
         wend
         rsUpdateEntry7.close

          response.write("<tr><td colspan=11 class='td-r strong'>Gasto proporcional por viaticos: </td><td class='td-r strong' style='color:red'>-"& rsUpdateEntry3("viaticos")  &"</td></tr>")
         

         '-------------------------------------------------CONSUMO EN COMBUSTIBLES'

         strSQL="select x.TarjetaCombustible as Tarjeta,x.Estacion,x.ID_viatico,dbo.fn_GetMonthName(x.Docdate,'spanish') as Fecha,x.ID_unidad,substring(f.Descripcion,1,14) as vehiculo,y.kms_viaje,y.ciudad_destino,x.Litros,c.PrecioLitro,CONVERT(VARCHAR,CONVERT(MONEY, round(x.Litros*c.PrecioLitro,2)    ),1)   as Combustibles,convert(varchar,cast(d.Factor as float)) as Factor  from VoucherGas x left join viajes y on x.ID_viatico=y.ID_viatico left join flotilla f on y.id_unidad=f.id_unidad left join Combustibles C on f.Combustible=c.combustible left join ViaticoRemision d on d.Pedido="& request("fpedido") &" and x.ID_viatico=d.ID  where x.ID_viatico in (select distinct(a.ID) from VIATICOS_R1 a inner join ViaticoRemision b on a.ID=b.ID where b.Pedido="& request("fpedido") &")" 

         'response.write strSQL
         rsUpdateEntry6.Open strSQL, adoCon  
         rsUpdateEntry7.Open strSQL, adoCon

         Dim Campos(12)
         i=1
         rowin
         For Each rsUpdateEntry6 in rsUpdateEntry6.Fields
                Response.Write "<td class='subtitulo td-c'>" & rsUpdateEntry6.Name & "</td>"
                campos(i)=rsUpdateEntry6.Name
                i=i+1
         Next 
         rowout
                     
         while not rsUpdateEntry7.EOF
             rowin
             for i=1 to 11
                  response.write("<td class='fonttiny td-r'>" & rsUpdateEntry7(Campos(i)) &"</td>")
             next            
             response.write("<td class='fonttiny td-r'>" & rsUpdateEntry7(Campos(12)) &"</td>")
             RowOut
             rsUpdateEntry7.MoveNext
         wend      
         rsUpdateEntry7.close
       
         response.write("<tr><td colspan=11 class='td-r strong'>Gasto proporcional por combustible: </td><td class='td-r strong' style='color:red'>-"& rsUpdateEntry3("Combustibles")  &"</td></tr>")
         


       

         if int(trim(rsUpdateEntry3("Flete")))>0 then  
             
              'response.write("<tr><td colspan=11 class='td-r'><font color=green><I>Super&aacute;vit por traslados: </I></td><td class='td-r fontmed'><font color=green> "& rsUpdateEntry3("AjusteFlete") &" </font></td></tr>")         
              vString="<H3 class='td-c'>AJUSTE EN LA UTILIDAD POR + <font color:green>"& rsUpdateEntry3("AjusteFlete") &"</font></H3>"
         else 
          
              'response.write("<tr><td colspan=11 class='td-r'><font color=red><I>D&eacute;ficit por traslados: </I></td><td class='td-r fontmed'><font color=red> "& rsUpdateEntry3("AjusteFlete") &"  </font></td></tr>")       
              vString="<H3 class='td-c'>AJUSTE EN LA UTILIDAD POR <font color:red>"&rsUpdateEntry3("AjusteFlete")  &"</font></H3>"
         end if
   
         closeTable

 

      
       if  int(trim(rsUpdateEntry3("ingresos")))>0 then
       
         response.write(vString)   'titulo del ajuste'
           
              CreateTable(700)  
              rowin
                

                   response.write("<td class='subtitulo td-c fontmed'>(MXN) INGRESOS <br>(sin fletes)</td><td class='subtitulo td-c fontmed'>COSTO CONTABLE </td><td class='subtitulo td-c fontmed'>UTILIDAD</td><td class='subtitulo td-c fontmed'>UTILIDAD <br>AJUSTADA </td><td class='subtitulo td-c fontmed'>% UTILIDAD</td></tr>")

                   rowin
                   response.write("<td class='fontmed td-c strong'>"& rsUpdateEntry3("SubtotalIngresosMXN")  &"</td>")                 
                   response.write("<td class='fontmed td-c strong'>"& rsUpdateEntry3("SubtotalCosto") &"</td>")
                   response.write("<td class='fontmed td-c strong'>"& rsUpdateEntry3("UtilidadMXN") &"</td>")
                   response.write("<td class='fontmed td-c strong'>"& rsUpdateEntry3("UtilidadAjustada") &"</td>")                               
                   response.write("<td class='fontmed td-c strong'>"& mid(rsUpdateEntry3("PcrtUtilAjustada"),1,5) &" % </td></tr>")                      
                   rsUpdateEntry3.close
                   closeTable


       end if

  

  


         else
               response.write("<tr><td colspan=8>no existen pedidos con esos criterios</td></tr>")

         end if
         response.write("</table> <P>&nbsp;</P>")

       
    end if

end sub





sub MinimosDeStock 'contenido 86'

     DoAlert
     response.write("<P>&nbsp;</P>")
     titulo="Minimos de Stock en DMX <br><font class='fontmed'>(la configuracion de stock minimo se debe colocar en DMX en el almacen MXL)</font>"
     DoTitulo(titulo)

     response.write("<form id='MinStock' method='POST' action='ShowContent.asp'>")
     response.write("<input type='hidden' name='contenido' value=86>")
     response.write("<input type='hidden' name='action' value=1>")

     CreateTable(500)
     response.write("<tr><td class='td-r' width='40%'>Seleccione familia: </td><td class='td-l' width='60%'><select name='ffamilia'>")
     strSQL="select * from OITB order by ItmsGrpNam"
     rsUpdateEntry.Open strSQL, adoCon2    'DMX'

     response.write("<option value='0'>todas las familias</option>")
     while not rsUpdateEntry.EOF
              if trim(request("ffamilia"))=trim(rsUpdateEntry("ItmsGrpCod")) then
                 response.write("<option value='"& rsUpdateEntry("ItmsGrpCod") &"' SELECTED>"&  rsUpdateEntry("ItmsGrpNam") &"</option>")
              else
                  response.write("<option value='"& rsUpdateEntry("ItmsGrpCod") &"'>"&  rsUpdateEntry("ItmsGrpNam") &"</option>")
              end if
              rsUpdateEntry.MoveNext
         wend
         rsUpdateEntry.close

         response.write("</select></td></tr>")         
         response.write("<tr><td class='td-c' colspan=2><br><input type='submit' value='Mostrar Minimos'></form></td></tr></table><P>&nbsp;</P>")

     if request("action")="1" then
          
          if request("ffamilia")=0 then
              strSQL="select a.ItemCode as Codigo,a.ItemName as Descripcion,b.ItmsGrpNam as Familia,cast(c.MinStock as int) as MinStock,cast( (select isnull( (select cast(SUM(InQty)-SUM(OutQty) as int) from PRODUCTIVA_DMX.dbo.OINM WITH (NOLOCK)  where ItemCode=a.ItemCode),0) +  isnull( (select cast(SUM(InQty)-SUM(OutQty) as int) from PRODUCTIVA_DFW.dbo.OINM WITH (NOLOCK)  where ItemCode=a.ItemCode),0) ) as int )  AS Stock,case when (select isnull( (select cast(SUM(InQty)-SUM(OutQty) as int) from PRODUCTIVA_DMX.dbo.OINM WITH (NOLOCK) where ItemCode=a.ItemCode),0) + isnull( (select cast(SUM(InQty)-SUM(OutQty) as int) from PRODUCTIVA_DFW.dbo.OINM WITH (NOLOCK) where ItemCode=a.ItemCode),0) ) < c.MinStock then 'reordene' else '' end as Estatus   from OITM a WITH (NOLOCK) inner join OITB b WITH (NOLOCK) on a.ItmsGrpCod=b.ItmsGrpCod inner join OITW c WITH (NOLOCK) on c.WhsCode='001-Mxl' and a.ItemCode=c.ItemCode where a.InvntItem='Y' and c.MinStock>0" 
          else
               strSQL="select a.ItemCode as Codigo,a.ItemName as Descripcion,b.ItmsGrpNam as Familia,cast(c.MinStock as int) as MinStock,cast( (select isnull( (select cast(SUM(InQty)-SUM(OutQty) as int) from PRODUCTIVA_DMX.dbo.OINM WITH (NOLOCK)  where ItemCode=a.ItemCode),0) +  isnull( (select cast(SUM(InQty)-SUM(OutQty) as int) from PRODUCTIVA_DFW.dbo.OINM WITH (NOLOCK)  where ItemCode=a.ItemCode),0) ) as int )  AS Stock, case when (select isnull( (select cast(SUM(InQty)-SUM(OutQty) as int) from PRODUCTIVA_DMX.dbo.OINM WITH (NOLOCK) where ItemCode=a.ItemCode),0) + isnull( (select cast(SUM(InQty)-SUM(OutQty) as int) from PRODUCTIVA_DFW.dbo.OINM WITH (NOLOCK) where ItemCode=a.ItemCode),0) ) < c.MinStock then 'reordene' else '' end as Estatus   from OITM a WITH (NOLOCK) inner join OITB b WITH (NOLOCK) on a.ItmsGrpCod=b.ItmsGrpCod inner join OITW c WITH (NOLOCK) on c.WhsCode='001-Mxl' and a.ItemCode=c.ItemCode where a.InvntItem='Y' and c.MinStock>0 and a.ItmsGrpCod='" & request("ffamilia") &"'"
          end if
          'response.write  strSQL

           rsUpdateEntry.Open strSQL, adoCon2   'DMX'
           rsUpdateEntry2.Open strSQL, adoCon2

           CreateTable(860)
           CamposRS1
           ShowValoresRS2
           closeTable

     end if


end sub



sub Flotilla 'contenido 87'

      if request("action")="1" then
           DoAlert

           if request("control")="1" then   'asignar concepto a un archivo subido al servidor'
             strSQL="update FileManager set Concepto='"& request("fconcepto") &"' where ID=" & request("ID")
             'response.write strSQL
            rsUpdateEntry.Open strSQL, adoCon   'NET'

           end if


           response.write("<div id='edit'>&nbsp;</div>") 

           titulo="Padr&oacute;n Vehicular"
           DoTitulo(titulo)

           strSQL="select b.Activo as 'Tipo',a.ID_unidad,a.RS,a.Fabricante,a.Modelo,a.Descripcion,a.Combustible,a.TarjetaCombustible as Tarjeta,a.Placas,a.Plataforma,a.Total_Llantas as [Total llantas],dbo.fn_GetMonthName(vigencia_TC,'spanish') as [vigencia TC],kilometraje=(select isnull(max(Kilometraje),0) as kilometraje from check_unidad where ID_unidad=a.id_unidad),c.Sucursalshort as Sucursal from Flotilla a inner join cat_activos b on a.tipo_activo=b.tipo_activo  left join cat_sucursal c on a.Clave_Sucursal=c.Clave_Sucursal where a.activado=1 order by a.tipo_activo,a.ID_unidad "
           'response.write strSQL
           rsUpdateEntry.Open strSQL, adoCon   'NET'
           rsUpdateEntry2.Open strSQL, adoCon

           CreateTable(1100)
           CamposRS1

           while not rsUpdateEntry2.EOF
             response.write("<tr>")
             response.write("<td class='td-l fonttiny'>"& rsUpdateEntry2("tipo") &"</td>")    %>
                  <td class='td-r fontmed'>  
                      <a href="Javascript:sendReq('/modules/FormAddCar.asp','ID','<%=rsUpdateEntry2("ID_unidad")%>','edit')" class="strong"><%=rsUpdateEntry2("ID_unidad")%>               
                  </td>

             <%
             response.write("<td class='td-c fontmed'>"& rsUpdateEntry2("RS") &"</td>")
             response.write("<td class='td-l fonttiny'>"& rsUpdateEntry2("Fabricante") &"</td>")
             response.write("<td class='td-l fontmed'>"& rsUpdateEntry2("Modelo") &"</td>")
             response.write("<td class='td-l fonttiny'>"& mid(rsUpdateEntry2("Descripcion"),1,40) &"</td>")
             response.write("<td class='td-l fonttiny'>"& rsUpdateEntry2("Combustible") &"</td>")
             response.write("<td class='td-l fontmed' width='130px'>"& rsUpdateEntry2("Tarjeta") &"</td>")
             response.write("<td class='td-c fontmed'>"& rsUpdateEntry2("Placas") &"</td>")
             response.write("<td class='td-c fontmed'>"& rsUpdateEntry2("Plataforma") &"</td>")
             response.write("<td class='td-c fontmed'>"& rsUpdateEntry2("Total Llantas") &"</td>")
             response.write("<td class='td-r fonttiny' width='110px'>"& rsUpdateEntry2("vigencia TC") &"</td>")
             response.write("<td class='td-r fontmed'>"& rsUpdateEntry2("kilometraje") &"</td>")
             response.write("<td class='td-r fontmed'>"& rsUpdateEntry2("Sucursal") &"</td>")
             response.write("</tr>")
             rsUpdateEntry2.MoveNext
           wend
           rsUpdateEntry2.close

           closeTable 
           response.write("<P>&nbsp;</P>")      
           response.write("<P class='td-c fontmed'><a href='ShowContent.asp?contenido=87&action=2'>agregar un veh&iacute;culo</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;| <a href='/Upload/uploadform5.asp'>subir archivos al servidor</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;| <a   <a href='ShowContent.asp?contenido=88&action=1'>ir a control de viajes </a></P>")

           titulo="Archivos relacionados a la flotilla"
           DoTitulo(titulo)

           CreateTable(480)
           response.write("<tr><td class='td-c subtitulo fontmed'>Nombre archivo</td><td class='td-c subtitulo fontmed'>Subido en fecha</td><td class='td-c subtitulo fontmed'>Control</td></tr>")

           strSQL="select *,.dbo.fn_GetMonthName(LogDate,'spanish') as Fecha,case when Concepto is null then 'indique un nombre' else Concepto end as Concepto2,case when Concepto is null then 1 else 0 end as Estatus from FileManager where ID>0 order by ID Desc"
           rsUpdateEntry2.Open strSQL, adoCon   'NET'

           while not rsUpdateEntry2.EOF
                RowIn
                'rsUpdateEntry2("nombre_archivo")'
                 if rsUpdateEntry2("Estatus") =1 then
                     response.write("<td class='td-r fontmed'><a href='ShowContent.asp?contenido=87&action=6&ID="& rsUpdateEntry2("ID") &"'>"& rsUpdateEntry2("Concepto2") &"</a></td>")
                 else
                     response.write("<td class='td-r fontmed'><a href='"& rsUpdateEntry2("ruta") & rsUpdateEntry2("nombre_archivo")  &"' target='download'>"& rsUpdateEntry2("Concepto") &"</a></td>")
                 end if
                 response.write("<td class='td-r fontmed'>"& rsUpdateEntry2("Fecha") &"</td>")
                 response.write("<td class='td-r fontmed'>x</td>")
                RowOut
                rsUpdateEntry2.MoveNext
           wend
           rsUpdateEntry2.close
           closeTable
        
    end if


      if request("action")="2" then   'ADD'
           DoAlert
           titulo="Agregar un veh&iacute;culo al padr&oacute;n vehicular"
           DoTitulo(titulo)

           response.write("<form id='flotilla' action='ShowContent.asp' method='POST'>")
           response.write("<input type='hidden' name='contenido' value='87'>")
           response.write("<input type='hidden' name='action' value='3'>")

           response.write("<P class='td-c fontmed'><font class='subtitulo'>Tipo de activo: </font><select name='tipo'>")
           strSQL="select * from cat_activos"
           rsUpdateEntry.Open strSQL, adoCon
           while not rsUpdateEntry.EOF
                  response.write("<option value='"& rsUpdateEntry("tipo_activo")  &"'>"& rsUpdateEntry("Activo") &"</option>")
                   rsUpdateEntry.movenext
           wend
           rsUpdateEntry.close
           response.write("</select></P>")


           CreateTable(820)
           response.write("<tr><td class='subtitulo td-r fontmed strong' width='7%'>ID:</td><td class='td-l fontmed' width='7%'><input type='text' name='fID' style='width:60px'></td><td class='subtitulo td-r fontmed strong' width='8%'>Fabricante:</td><td class='td-l fontmed' width='14%'><input type='text' name='fVendor' style='width:160px' placeholder='ej Chevrolet Chrysler'></td><td class='subtitulo td-r fontmed strong' width='9%'>Kilometraje:</td><td class='td-l fontmed' width='10%'>---</td></tr>")
          
            response.write("<tr><td class='subtitulo td-r fontmed strong'>Raz&oacute;n Social:</td><td class='td-l fontmed'><select name='fRS'><option value='DMX'>DMX</option><option value='DFW'>DFW</option></select></td><td class='subtitulo td-r fontmed strong'>Modelo:</td><td class='td-l fontmed'><input type='text' name='fModelo' style='width:200px' placeholder='ej Ram 4000'></td><td class='subtitulo td-r fontmed strong'>Placas:</td><td class='td-l fontmed'><input type='text' name='fPlacas' style='width:160px'></td></tr>")
 
            response.write("<tr><td class='subtitulo td-r fontmed strong'>Sucursal:</td><td class='td-l fontmed'><select name='fSucursal'>")

              strSQL=  "select * from cat_sucursal order by clave_sucursal"     
              'response.write strSQL  
              rsUpdateEntry.Open strSQL, adoCon
      
              while not rsUpdateEntry.EOF
                   response.write("<option value='"& rsUpdateEntry("clave_sucursal")  &"'>"& rsUpdateEntry("sucursalShort") &"</option>")
                   rsUpdateEntry.movenext
              wend
               rsUpdateEntry.close

            response.write("</select></td><td class='subtitulo td-r fontmed strong'>Descripcion:</td><td class='td-l fontmed' colspan=3><input type='text' name='fDescripcion' style='width:370px' placeholder='ej Unidad de reparto Thorthon'></td></tr>")

             response.write("<tr><td class='subtitulo td-r fontmed strong'># Total llantas:</td><td class='td-l fontmed'><select name='fLlantas'>")

              for i=0 to 20 step 1                 
                    response.write("<option value='"&i&"'>"&i&"</option>")                              
              next
              response.write("</select></td><td class='subtitulo td-r fontmed strong'>Serie:</td><td class='td-l fontmed'><input type='text' name='fSerie' style='width:220px' ></td><td class='subtitulo td-r fontmed strong'>Ent Federativa Tarj-Circulacion</td><td class='td-l fontmed'><select name='fEstadoTC'>")

               strSQL="select * from cat_entidades order by id_entidad"
              'response.write strSQL  
              rsUpdateEntry2.Open strSQL, adoCon

              while not rsUpdateEntry2.EOF                 
                 response.write("<option value='"& rsUpdateEntry2("id_entidad") &"'>"& rsUpdateEntry2("entidad") &"</option>")                 
                 rsUpdateEntry2.MoveNext
              wend
              rsUpdateEntry2.close
              response.write("</select></td></tr>")
 
              response.write("<tr><td class='subtitulo td-r fontmed strong'>Plataforma: </td><td class='td-l fontmed'><select name='fPlataforma'>")
              response.write("<option value='N'>no</option><option value='S'>si</option></select></td>")              
              response.write("<td class='subtitulo td-r fontmed strong'>Fecha compra: </td><td class='td-l fontmed'> <input type='date' name='fFcompra'></td><td class='subtitulo td-r fontmed strong'>Aseguradora: </td><td class='td-l fontmed'><input type='text' class='anchox1' name='fSeguro'></td></tr>")

               response.write("<tr><td class='subtitulo td-r fontmed strong'>Max carga (kg): </td><td class='td-l fontmed'> <input type='number' name='fMaxCarga' style='width: 70px'> <td class='subtitulo td-r fontmed strong'>Monto compra: </td><td class='td-l fontmed'><input type='text' name='fMonto'></td></td><td class='subtitulo td-r fontmed strong'>Fecha vence seguro: </td><td class='td-l fontmed'><input type='date' name='fvencimiento'></tr><tr><td class='subtitulo td-r fontmed strong'>Combustible:</td><td class='td-l fontmed'><select name='fcombustible'>")           
              response.write("<option value='gasolina'>gasolina</option>")
              response.write("<option value='diesel'>diesel</option>")
              response.write("<option value='gas'>gas</option>")   
              response.write("</select><td class='subtitulo td-r fontmed strong'>Tarj Combustible:</td><td class='td-l fontmed'><input type='text' name='ftarjeta'></tr>")
              response.write("</select><td colspan=4>&nbsp;</td></td></tr>")
              response.write("<tr><td colspan=6 class='td-c'><input type='submit' value='agregar'></form></td></tr>")
           closeTable

      end if

       if request("action")="3" then   'ADD UP'

        if len(trim(request("fID")))>=2 then

                  strSQL="select * from Flotilla where ID_unidad='"& trim(request("fID")) &"'"
                  rsUpdateEntry.Open strSQL, adoCon

                  if not rsUpdateEntry.EOF then
                       rsUpdateEntry.close
                        response.redirect("ShowContent.asp?contenido=87&action=1&msg=error identificador duplicado")
                  else
                       rsUpdateEntry.close
                  end if


                  strSQL="set dateformat ymd;insert into Flotilla (tipo_activo,ID_unidad,RS,clave_sucursal,Fabricante,Modelo,Serie,Placas,Vigencia_Seguro,Descripcion,FechaCompra,MontoCompra,Plataforma,Total_Llantas,ID_USUARIO,LogDate,id_entidad_TC,seguro,MaxCarga,combustible,TarjetaCombustible) values ("& request("tipo") &","

                  
                  strSQL= strSQL& "'"& UCASE(trim(request("fID"))) &"',"
                  strSQL= strSQL& "'"& request("fRS")&"',"
                  strSQL= strSQL& "'"& request("fSucursal")&"',"
                  strSQL= strSQL& "'"& request("fVendor")&"',"
                  strSQL= strSQL& "'"& request("fModelo")&"',"
                  strSQL= strSQL& "'"& UCASE(trim(request("fserie"))) &"',"
                  strSQL= strSQL& "'"& UCASE(trim(request("fplacas"))) &"',"
                  strSQL= strSQL& "'"& request("fvencimiento")&"',"
                  strSQL= strSQL& "'"& request("fDescripcion")&"',"

                  if request("fFcompra")<>"" then
                      strSQL= strSQL& "'"& request("fFcompra")&"',"
                  else
                      strSQL= strSQL& "'1900-01-01',"
                  end if

                  if request("fmonto")<>"" then
                     strSQL= strSQL&  request("fMonto")&","
                  else
                     strSQL= strSQL& "0,"
                  end if

                  vCarga=trim(request("fMaxCarga"))
                  if vCarga="" then 
                      vCarga="0" 
                  end if

                  strSQL= strSQL& "'"& request("fPlataforma")&"',"
                  strSQL= strSQL&  request("fLlantas")&",'"& session("session_id")   &"',getdate(),'" & request("fEstadoTC") &"','"& request("fSeguro")&"',"& vCarga &",'"& request("fcombustible") &"','"& request("ftarjeta") &"' ) "

                  response.write strSQL
                  rsUpdateEntry.Open strSQL, adoCon

                  response.redirect("ShowContent.asp?contenido=87&action=1&msg=se ha agregado el registro")

           else

                  response.redirect("ShowContent.asp?contenido=87&action=1&msg=error al intentar registrar")

         end if

      end if



      if request("action")="4" then   'EDIT'
             'migrado a un DIV  FormAddCar.asp'          
      end if


  if request("action")="5" then   'EDIT UP'        
        vMaxCarga="0"
        if request("fMaxCarga")<>0 then
              vMaxCarga=trim(request("fMaxCarga"))
              vMaxCarga=replace(vMaxCarga,",","")
        end if

        strSQL="set dateformat ymd;UPDATE Flotilla set tipo_activo="&request("tipo") &", Fabricante='"& request("FVendor") &"', RS='"& request("fRS") &"', Modelo='"& UCASE( request("fmodelo")) &"', clave_sucursal='"&request("fSucursal")&"', Descripcion='"& request("fDescripcion")&"', Serie='"& request("fserie")&"',Placas='"& UCASE(request("fPlacas")) &"',Plataforma='"& request("fPlataforma")&"',FechaCompra='"& request("fanio")&"-"&request("fmes")&"-"&request("fdia")&"', MontoCompra=" & request("fMonto")& ",Total_Llantas=" & request("fEjes")& ",id_entidad_TC='"& request("fEstadoTC") &"',vigencia_TC='"& request("fanioTC")&"-"&request("fmesTC")&"-"&request("fdiaTC")&"',seguro='"& request("fSeguro") &"',Vigencia_Seguro='"& request("fanioSG")&"-"&request("fmesSG")&"-"&request("fdiaSG")&"',CapacTanque="& request("ftanque") &",Combustible='"& request("fcombustible") &"',TarjetaCombustible='"& request("ftarjeta") &"',ID_USUARIO='"& session("session_id") &"',LogDate=getdate(),MaxCarga="& vMaxCarga &"   where ID_unidad='"&request("fID") &"' "
        response.write strSQL

        rsUpdateEntry.Open strSQL, adoCon
        response.redirect("ShowContent.asp?contenido=87&action=1&msg=se ha actualizado correctamente")

  end if


  if request("action")="6" then   'EDIT UP'      
 
       titulo="Asignar un concepto a un archivo en servidor"
       DoTitulo(titulo)

       response.write("<form id='nombre' action='ShowContent.asp' method='Post'>")
       response.write("<input type='hidden' name='contenido' value='87'>")
       response.write("<input type='hidden' name='action' value='1'>")
       response.write("<input type='hidden' name='ID' value='"&request("ID") &"'>")
       response.write("<input type='hidden' name='control' value='1'>")
       response.write("<input type='hidden' name='msg' value='se ha renombrado el archivo de manera exitosa'>")

       CreateTable(550)
       RowIn
       response.write("<td class='td-r subtitulo'>Indique concepto por el archivo:</td><td class='td-l'><input style='width:260px' type='text' name='fconcepto' placeholder='utilice un concepto breve'></td>")
       response.write("<tr><td colspan=2>&nbsp;</td>")
       response.write("<tr><td colspan=2 class='td-c'><input type='submit' value='asignar'></td></table></form>")
 


  end if



end sub



sub ControlViajes  'contenido 88'  'control de viajes'

     
       if request("action")="1" then   'showviajes'
   
       'si falta un viatico de agregar un viaje lo agrega'
       strSQL="insert into viajes (id_viaje,ID_viatico,docdate,Chofer,TravelDate,ciudad_origen,mpio_Origen,Entidad_Origen,ID_unidad)  select TOP 1 (select max(id_viaje) from Viajes)+1,ID,DocDate,beneficiario,TravelDate1,case when almacen='SJR' then 'San Juan del Rio' else 'Mexicali' end, case when almacen='SJR' then '016' else '002' end,case when almacen='SJR' then '21' else '02' end,ID_unidad from VIATICOS_R1 where Docdate>='2021-04-01' and Helpdesk<5 and ID not in (select ID_viatico from viajes)  order by ID "
       rsUpdateEntry.Open strSQL, adoCon

       DoAlert
       response.write("<div id='edit'>&nbsp;</div>")
       strSQL="select b.entidad,count(a.Entidad_Origen) as calculo from viajes a inner join cat_entidades b on a.Entidad_Origen=b.id_entidad where a.id_viaje>0 group by b.entidad"
       'response.write strSQL
       rsUpdateEntry.Open strSQL, adoCon

       titulo="Control de viajes <br>" & rsUpdateEntry("entidad") & " ("&  rsUpdateEntry("calculo") &") / "
       rsUpdateEntry.MoveNext
       titulo= titulo &  rsUpdateEntry("entidad") & " ("&  rsUpdateEntry("calculo") &") "
       rsUpdateEntry.MoveNext
       titulo= titulo &  rsUpdateEntry("entidad") & " ("&  rsUpdateEntry("calculo") &") "
       rsUpdateEntry.close
       DoTitulo(titulo)

       Const adCmdText = &H0001
       Const adOpenStatic = 3
       nPage=0
       registros=1

       strSQL="select a.ID_viaje,a.ID_viatico as ID_viatico, .dbo.fn_GetMonthName(a.TravelDate,'spanish') as FechaViaje,a.Chofer,isnull(a.Chofer_aux,'') as Chofer_aux,a.ciudad_origen,b.entidad as entidad_origen,a.ciudad_destino,c.entidad as entidad_destino,'' as check_salida,'' as check_entrada,a.kms_viaje,id_unidad from viajes a left join cat_entidades b on a.entidad_origen=b.id_entidad  left join cat_entidades c on a.entidad_destino=c.id_entidad  where a.id_viaje>0 order by a.ID_viaje desc"
       'response.write(strSQL&"<BR>")
       rsUpdateEntry3.Open strSQL, adoCon
       rsUpdateEntry4.Open strSQL, adoCon, adOpenStatic, adCmdText

       rsUpdateEntry4.PageSize = 10 
       nPageCount = rsUpdateEntry4.PageCount


       CreateTable(1100)

       if request("vPage")<>"" then
            nPage = int(trim(request("vPage")))
        else
            nPage=1
        end if
      rsUpdateEntry4.AbsolutePage=npage
      response.write("<tr><td colspan=3 class='td-r'>PAGINAS:</td><td colspan=10>")
       for i=1 to nPageCount step 1
            if i<>npage then  
                response.write("<a href='/modules/ShowContent.asp?contenido=88&action=1&vPage="&i &"'>") 
                response.write( i &"</A>&nbsp;&nbsp;&nbsp;&nbsp;") 
            end if
        next
        response.write("</td></tr>")

       CamposRS3


       while not rsUpdateEntry4.EOF AND registros<=10
           RowIn  
           response.write("<td class='td-r fontmed'>&nbsp;</td>")

           if rsUpdateEntry4("ciudad_origen")="Mexicali" then
                   response.write("<td class='td-r fontmed gris'><a href='ShowContent.asp?contenido=88&action=2&ID="&rsUpdateEntry4("ID_viatico") &"' class='strong'>" & rsUpdateEntry4("ID_viatico") &"</A></td>")                     
                   response.write("<td class='td-r fontmed gris'>"& rsUpdateEntry4("FechaViaje") &"</td>")
                   response.write("<td class='td-r fontmed gris'>"& rsUpdateEntry4("Chofer") &"</td>")
                   response.write("<td class='td-r fontmed gris'>"& rsUpdateEntry4("Chofer_aux") &"</td>")
                   response.write("<td class='td-r fontmed gris'>"& rsUpdateEntry4("ciudad_origen") &"</td>")
                   response.write("<td class='td-r fontmed gris'>"& rsUpdateEntry4("entidad_origen") &"</td>")
                   response.write("<td class='td-r fontmed gris'>"& rsUpdateEntry4("ciudad_destino") &"</td>")
                   response.write("<td class='td-r fontmed gris'>"& rsUpdateEntry4("entidad_destino") &"</td>")
                   
                   strSQL="select * from Check_unidad where id_viaje="& rsUpdateEntry4("id_viaje") & " and mode='out' and kilometraje is not null"
                   rsUpdateEntry7.Open strSQL, adoCon

                   flag=0

                   if rsUpdateEntry7.EOF then   'si es fin de archivo quiere decir que aun no se ingresa la info'
                        response.write("<td class='td-r fontmed gris'><a href='ShowContent.asp?contenido=89&ID="& rsUpdateEntry4("id_VIAJE") &"&unidad="& rsUpdateEntry4("id_unidad") &"&action=1'><img src='/images/check-out2.png' border='' width='26px' height='30px' border=0 title='Check salida'></a></td>")
                   else
                        response.write("<td class='td-r fontmed gris'>-</td>")
                        flag=flag+1
                   end if
                   rsUpdateEntry7.close

                   strSQL="select * from Check_unidad where id_viaje="& rsUpdateEntry4("id_viaje") & " and mode='in' and kilometraje is not null"
                   rsUpdateEntry7.Open strSQL, adoCon

                   if rsUpdateEntry7.EOF then   'si es fin de archivo quiere decir que aun no se ingresa la info'
                       response.write("<td class='td-r fontmed gris'><a href='ShowContent.asp?contenido=89&ID="& rsUpdateEntry4("id_VIAJE") &"&unidad="& rsUpdateEntry4("id_unidad") & "&action=2'><img src='/images/check-in.png' border='' width='26px' height='30px' border=0 title='Check entrada'></a></td>")    
                   else
                         response.write("<td class='td-r fontmed gris'>-</td>")
                         flag=flag+1
                   end if     
                   rsUpdateEntry7.close  

                   if flag=2 and int(trim(rsUpdateEntry4("kms_viaje")))=0 then
                        strSQL="update viajes set kms_viaje=(select case when c.kilometraje>b.kilometraje then c.kilometraje-b.kilometraje else 0 end  from viajes a inner join check_unidad b on b.mode='out' and a.ID_viaje=b.id_viaje inner join check_unidad c on c.mode='in' and a.ID_viaje=c.id_viaje where a.ID_viaje="& rsUpdateEntry4("id_VIAJE") &")   where ID_viaje="& rsUpdateEntry4("id_VIAJE")
                        rsUpdateEntry5.Open strSQL, adoCon

                        response.write("<td class='td-r fontmed gris'>-</td>") 
                   else
                        response.write("<td class='td-r fontmed gris'>"& rsUpdateEntry4("kms_viaje") &"</td>")  
                   end if 
                            
                   response.write("<td class='td-r fontmed gris'>"& rsUpdateEntry4("id_unidad") &"</td>")  
            else
                   response.write("<td class='td-r fontmed'><a href='ShowContent.asp?contenido=88&action=2&ID="&rsUpdateEntry4("ID_viatico") &"' class='strong'>" & rsUpdateEntry4("ID_viatico") &"</A></td>")                     
                   response.write("<td class='td-r fontmed'>"& rsUpdateEntry4("FechaViaje") &"</td>")
                   response.write("<td class='td-r fontmed'>"& rsUpdateEntry4("Chofer") &"</td>")
                   response.write("<td class='td-r fontmed'>"& rsUpdateEntry4("Chofer_aux") &"</td>")
                   response.write("<td class='td-r fontmed'>"& rsUpdateEntry4("ciudad_origen") &"</td>")
                   response.write("<td class='td-r fontmed'>"& rsUpdateEntry4("entidad_origen") &"</td>")
                   response.write("<td class='td-r fontmed'>"& rsUpdateEntry4("ciudad_destino") &"</td>")
                   response.write("<td class='td-r fontmed'>"& rsUpdateEntry4("entidad_destino") &"</td>")

                   flag=0
                   strSQL="select * from Check_unidad where id_viaje="& rsUpdateEntry4("id_viaje") & " and mode='out' and kilometraje is not null"
                   rsUpdateEntry7.Open strSQL, adoCon

                   if rsUpdateEntry7.EOF then   'si es fin de archivo quiere decir que aun no se ingresa la info'
                        response.write("<td class='td-r fontmed'><a href='ShowContent.asp?contenido=89&ID="& rsUpdateEntry4("id_VIAJE") & "&unidad="& rsUpdateEntry4("id_unidad") & "&action=1'><img src='/images/check-out2.png' border='' width='26px' height='30px' border=0 title='Check salida'></a></td>")
                   else
                       response.write("<td class='td-r fontmed'>-</td>")
                       flag=flag+1
                   end if
                   rsUpdateEntry7.close

                   strSQL="select * from Check_unidad where id_viaje="& rsUpdateEntry4("id_viaje") & " and mode='in' and kilometraje is not null"
                   rsUpdateEntry7.Open strSQL, adoCon

                   if rsUpdateEntry7.EOF then   'si es fin de archivo quiere decir que aun no se ingresa la info'
                        response.write("<td class='td-r fontmed'><a href='ShowContent.asp?contenido=89&ID="& rsUpdateEntry4("id_VIAJE") &"&unidad="& rsUpdateEntry4("id_unidad") &"&action=2'><img src='/images/check-in.png' border='' width='26px' height='30px' border=0 title='Check entrada'></a></td>")  
                   else
                        response.write("<td class='td-r fontmed'>-</td>")  
                        flag=flag+1              
                   end if  
                   rsUpdateEntry7.close       

                    if flag=2 and int(trim(rsUpdateEntry4("kms_viaje")))=0 then
                        strSQL="update viajes set kms_viaje=(select case when c.kilometraje>b.kilometraje then c.kilometraje-b.kilometraje else 0 end  from viajes a inner join check_unidad b on b.mode='out' and a.ID_viaje=b.id_viaje inner join check_unidad c on c.mode='in' and a.ID_viaje=c.id_viaje where a.ID_viaje="& rsUpdateEntry4("id_VIAJE") &")   where ID_viaje="& rsUpdateEntry4("id_VIAJE")
                         rsUpdateEntry5.Open strSQL, adoCon
                        response.write("<td class='td-r fontmed'>-</td>") 
                   else
                        response.write("<td class='td-r fontmed'>"& rsUpdateEntry4("kms_viaje") &"</td>") 
                   end if 

                   response.write("<td class='td-r fontmed'>"& rsUpdateEntry4("id_unidad") &"</td>")  
            end if

           RowOut             
           rsUpdateEntry4.MoveNext
           registros=registros+1
       wend
       rsUpdateEntry4.close
       closeTable

       ShowRevisionUnidades

 end if



 if request("action")=2 then   'edit'

        if request("ID") <>"" then
             strSQL="select a.*,municipio_origen=isnull( (select top 1 municipio from cat_municipios where id_entidad=a.entidad_origen and id_municipio=a.mpio_origen),'error'),municipio_destino=isnull( (select top 1 municipio from cat_municipios where id_entidad=a.entidad_destino and id_municipio=a.mpio_destino),'error') from viajes a where a.id_viatico="& request("ID") 
            'response.write strSQL
       else
        
             strSQL="select a.*,municipio_origen=isnull( (select top 1 municipio from cat_municipios where id_entidad=a.entidad_origen and id_municipio=a.mpio_origen),'error'),municipio_destino=isnull( (select top 1 municipio from cat_municipios where id_entidad=a.entidad_destino and id_municipio=a.mpio_destino),'error') from viajes a where a.id_viaje="& session("viaje")  
            'response.write strSQL

        end if

           rsUpdateEntry.Open strSQL, adoCon
           
           DoAlert
           titulo="Editando viaje  "& rsUpdateEntry("id_viaje")  
           DoTitulo(titulo)

           response.write("<div id='destino'>")

           response.write("<form id='flotilla' action='ShowContent.asp' method='POST'>")
           response.write("<input type='hidden' name='contenido' value='88'>")
           response.write("<input type='hidden' name='action' value='3'>")

           session("viaje")=1
           session("viaje")=rsUpdateEntry("ID_viaje")

           response.write("<input type='hidden' name='fID' value='"& rsUpdateEntry("ID_viaje") &"'>")

           CreateTable(600)
           response.write("<tr><td class='subtitulo td-r fontmed' width='35%'>ID_viaje:</td><td class='td-l fontmed' width='65%'>"& rsUpdateEntry("ID_viaje") &"</td></tr>")
           response.write("<tr><td class='subtitulo td-r fontmed'>ID_Viatico:</td><td class='td-l fontmed'>"& rsUpdateEntry("ID_viatico") &"</td></tr>")

           response.write("<tr><td class='subtitulo td-r fontmed'>Chofer:</td><td class='td-l fontmed'><select name='fchofer'>")

           strSQL="select * from Empleados where FechaEgreso is null order by ID desc "
           'response.write strSQL  
           rsUpdateEntry2.Open strSQL, adoCon           

           while not  rsUpdateEntry2.EOF
              if rsUpdateEntry2("nombres") &" " & rsUpdateEntry2("PATERNO") &" " & rsUpdateEntry2("MATERNO") = rsUpdateEntry("chofer") then
                  response.write("<option value='"& rsUpdateEntry2("nombres") &" " & rsUpdateEntry2("PATERNO") &" " & rsUpdateEntry2("MATERNO") &"' SELECTED>"& rsUpdateEntry2("nombres") &" " & rsUpdateEntry2("PATERNO") &" " & rsUpdateEntry2("MATERNO") &"</option>")
              else
                   response.write("<option value='"& rsUpdateEntry2("nombres") &" " & rsUpdateEntry2("PATERNO") &" " & rsUpdateEntry2("MATERNO") &"'>"& rsUpdateEntry2("nombres") &" " & rsUpdateEntry2("PATERNO") &" " & rsUpdateEntry2("MATERNO") &"</option>")
              end if
              rsUpdateEntry2.MoveNext
           wend
            rsUpdateEntry2.MoveFirst
            response.write("</select></td></tr><tr><td class='subtitulo td-r fontmed'>Chofer auxiliar:</td><td class='td-l fontmed'><select name='fChoferAux'>")
            flag=0      

            while not  rsUpdateEntry2.EOF
               if rsUpdateEntry2("nombres") &" " & rsUpdateEntry2("PATERNO") &" " & rsUpdateEntry2("MATERNO") = rsUpdateEntry("chofer_aux") then
                  response.write("<option value='"& rsUpdateEntry2("nombres") &" " & rsUpdateEntry2("PATERNO") &" " & rsUpdateEntry2("MATERNO") &"' SELECTED>"& rsUpdateEntry2("nombres") &" " & rsUpdateEntry2("PATERNO") &" " & rsUpdateEntry2("MATERNO") &"</option>")
                  flag=1
              else
                   response.write("<option value='"& rsUpdateEntry2("nombres") &" " & rsUpdateEntry2("PATERNO") &" " & rsUpdateEntry2("MATERNO") &"'>"& rsUpdateEntry2("nombres") &" " & rsUpdateEntry2("PATERNO") &" " & rsUpdateEntry2("MATERNO") &"</option>")
              end if
              rsUpdateEntry2.MoveNext
           wend
           rsUpdateEntry2.close

            if flag=0 then
                 response.write("<option value='' SELECTED>no aplica</option>")
            else
                 response.write("<option value='999'>no aplica</option>")
            end if

           response.write("</select></td></tr>")
           

           

           response.write("<tr><td class='subtitulo td-r fontmed'>Entidad Federativa Origen:</td><td class='td-l fontmed'>")   %>

           <select id="entidad_origen" onchange="Javascript:PassValueChangeDiv('entidad_origen','destino','/modules/FormChangeState.asp')">   <%  'elemento,divn,modulo  

           strSQL="select * from cat_entidades order by id_entidad"
           rsUpdateEntry2.Open strSQL, adoCon 

          while not rsUpdateEntry2.EOF
              if trim(rsUpdateEntry2("id_entidad")) = trim(rsUpdateEntry("entidad_origen")) then  
                  response.write("<option value='"& rsUpdateEntry2("id_entidad") &"' SELECTED>"& rsUpdateEntry2("entidad") &"</option>")
              else
                  response.write("<option value='"& rsUpdateEntry2("id_entidad") &"'>"& rsUpdateEntry2("entidad")  &"</option>")
              end if
              rsUpdateEntry2.MoveNext
          wend
          rsUpdateEntry2.Close
          response.write("</select></td></tr>")
                
          '----------------------------------------------------------------------------------------------------------------------------------------'
          
          response.write("<tr><td class='subtitulo td-r fontmed'>Municipio:</td><td class='td-l fontmed'>"& rsUpdateEntry("municipio_origen") &"</td></tr>")
         

          strSQL="select id_entidad,id_municipio,ciudad from cat_municipios where id_entidad='"& rsUpdateEntry("entidad_origen") &"' group by id_entidad,id_municipio,ciudad  order by ciudad"
          'response.write strSQL
          rsUpdateEntry2.Open strSQL, adoCon 

          if trim(rsUpdateEntry("ciudad_origen"))<>"" then
              response.write("<tr><td class='subtitulo td-r fontmed'>Ciudad Origen:</td><td class='td-l fontmed'><select name='fCiudadOrigen'>")          
          else
              response.write("<tr><td class='subtitulo td-r fontmed'><font color=red>confirme Ciudad:</font></td><td class='td-l fontmed'><select name='fCiudadOrigen'>")
          end if

          while not rsUpdateEntry2.EOF
              if  trim(rsUpdateEntry2("id_municipio")) = trim(rsUpdateEntry("mpio_origen")) and trim(rsUpdateEntry2("ciudad")) = trim(rsUpdateEntry("ciudad_origen")) then  
                  response.write("<option value='"& rsUpdateEntry2("ciudad") &"' SELECTED>"& rsUpdateEntry2("ciudad") &"</option>")
              else
                  response.write("<option value='"& rsUpdateEntry2("ciudad") &"'>"& rsUpdateEntry2("ciudad")  &"</option>")
              end if
              rsUpdateEntry2.MoveNext
          wend
          rsUpdateEntry2.Close
          response.write("</select></td></tr>")

          '----------------------------------------------------------------------------------------------------------------------------------------'

           response.write("<tr><td class='subtitulo td-r fontmed'>Entidad Federativa Destino:</td><td class='td-l fontmed'>")   %>

           <select id="entidad_destino" onchange="Javascript:PassValueChangeDiv('entidad_destino','destino','/modules/FormChangeState2.asp')">   <%  'elemento,divn,modulo  

           strSQL="select * from cat_entidades order by id_entidad"
           rsUpdateEntry2.Open strSQL, adoCon 

          while not rsUpdateEntry2.EOF
              if trim(rsUpdateEntry2("id_entidad")) = trim(rsUpdateEntry("entidad_destino")) then  
                  response.write("<option value='"& rsUpdateEntry2("id_entidad") &"' SELECTED>"& rsUpdateEntry2("entidad") &"</option>")
              else
                  response.write("<option value='"& rsUpdateEntry2("id_entidad") &"'>"& rsUpdateEntry2("entidad")  &"</option>")
              end if
              rsUpdateEntry2.MoveNext
          wend
          rsUpdateEntry2.Close
          response.write("</select></td></tr>")


          if rsUpdateEntry("mpio_destino") <>"" then 

                response.write("<tr><td class='subtitulo td-r fontmed'>Municipio:</td><td class='td-l fontmed'>"& rsUpdateEntry("municipio_destino") &"</td></tr>")
               
                strSQL="select id_entidad,id_municipio,ciudad from cat_municipios where id_entidad='"& rsUpdateEntry("entidad_destino") &"' group by id_entidad,id_municipio,ciudad  order by ciudad"
                'response.write strSQL
                rsUpdateEntry2.Open strSQL, adoCon 

                if trim(rsUpdateEntry("ciudad_destino"))<>"" then
                    response.write("<tr><td class='subtitulo td-r fontmed'>Ciudad Destino:</td><td class='td-l fontmed'><select name='fCiudadDestino'>")          
                else
                    response.write("<tr><td class='subtitulo td-r fontmed'><font color=red>confirme Ciudad:</font></td><td class='td-l fontmed'><select name='fCiudadDestino'>")
                end if

                while not rsUpdateEntry2.EOF
                    if  trim(rsUpdateEntry2("id_municipio")) = trim(rsUpdateEntry("mpio_destino")) and trim(rsUpdateEntry2("ciudad")) = trim(rsUpdateEntry("ciudad_destino")) then  
                        response.write("<option value='"& rsUpdateEntry2("ciudad") &"' SELECTED>"& rsUpdateEntry2("ciudad") &"</option>")
                    else
                        response.write("<option value='"& rsUpdateEntry2("ciudad") &"'>"& rsUpdateEntry2("ciudad")  &"</option>")
                    end if
                    rsUpdateEntry2.MoveNext
                wend
                rsUpdateEntry2.Close
                response.write("</select></td></tr>")
            end if

          '----------------------------------------------------------------------------------------------------------------------------------------'
           
           response.write("<tr><td class='subtitulo td-r fontmed'>Fecha del viaje:</td><td class='td-l fontmed'> <select name=fDia>")

                
               for i=1 to 31
                    if day(rsUpdateEntry("TravelDate"))=i then
                       response.write("<option value="& i &" SELECTED>"& i &"</option>")
                   else
                       response.write("<option value="& i &" >"& i &"</option>")
                   end if
               next
               response.write("</select> / <select name=fMes>")
               for i=1 to 12
                    if month(rsUpdateEntry("TravelDate"))=i then
                       response.write("<option value="& i &" SELECTED>"& meses(i) &"</option>")
                   else
                       response.write("<option value="& i &" >"&  meses(i) &"</option>")
                   end if
               next
               response.write("</select> / <select name=fAnio>")

               for i=year(date())+1 to year(date())-1 step -1
                   if year(rsUpdateEntry("TravelDate"))=i then
                       response.write("<option value="& i &" SELECTED>"& i &"</option>")
                   else
                       response.write("<option value="& i &" >"& i &"</option>")
                   end if
               next
               response.write("</select></td></tr>")

           
               strSQL="select * from flotilla where tipo_activo<=3"
               'response.write strSQL
               rsUpdateEntry2.Open strSQL, adoCon 

                response.write("<tr><td class='subtitulo td-r fontmed'>ID_unidad:</td><td class='td-l fontmed'> <select name=fUnidad>")

                while not rsUpdateEntry2.EOF
                    if  trim(rsUpdateEntry2("id_unidad")) = trim(rsUpdateEntry("id_unidad"))  then  
                        response.write("<option value='"& rsUpdateEntry2("id_unidad") &"' SELECTED>"& rsUpdateEntry2("id_unidad") &"</option>")
                    else
                        response.write("<option value='"& rsUpdateEntry2("id_unidad") &"'>"& rsUpdateEntry2("id_unidad")  &"</option>")
                    end if
                    rsUpdateEntry2.MoveNext
                wend
               rsUpdateEntry2.Close
               response.write("</select></td></tr>")
            

           
           response.write("<tr><td colspan=2 class='td-c'><input type='submit' value='actualizar'></form></td></tr>")
           CloseTable

           response.write("</div>")
           rsUpdateEntry.close

     end if


   if request("action")=3 then  'EDIT VIAJE UP'

         strSQL="UPDATE viajes set chofer='"& request("fchofer") &"',"
         if request("fChoferAux")<>"" and request("fChoferAux")<>"999" then
             strSQL= strSQL& "chofer_aux='"& request("fChoferAux") &"',"
         end if
         if request("fChoferAux")="999" then
             strSQL= strSQL& "chofer_aux='',"
         end if
         strSQL= strSQL& "ciudad_origen='"& request("fCiudadOrigen") &"',ciudad_destino='"& request("fCiudadDestino") &"',DocDate='"& request("fAnio") &"-"& request("fmes") &"-"& request("fDia") &"',ID_USUARIO='"& session("session_id") &"',logdate=getdate(),ID_unidad='"& request("funidad") &"'  where id_viaje="& request("fID")

         response.write strSQL
         rsUpdateEntry2.Open strSQL, adoCon 
         response.redirect("ShowContent.asp?contenido=88&action=1&msg=datos del viaje actualizado")

   end if


   if request("action")=4 then  'UP ENTIDAD ORIGEN'
        vString=request("fciudad_origen")
        vPos=inStr(vString,"*")
        vString1=mid(vString,1,Vpos-1)
        vString2=mid(vString,Vpos+1,30)

        strSQL="UPDATE viajes set entidad_origen='"& request("fentidad_origen") &"',mpio_Origen='"& vstring1 &"',ciudad_origen='"& vString2 &"' where id_viaje="& session("viaje") 
        response.write strSQL
        rsUpdateEntry2.Open strSQL, adoCon 
        response.redirect("ShowContent.asp?contenido=88&action=2&msg=entidad de origen actualizada")

   end if


   if request("action")=5 then  'UP ENTIDAD DESTINO'
        vString=request("fciudad_destino")
        vPos=inStr(vString,"*")
        vString1=mid(vString,1,Vpos-1)
        vString2=mid(vString,Vpos+1,30)

        strSQL="UPDATE viajes set entidad_destino='"& request("fentidad_destino") &"',mpio_Destino='"& vstring1 &"',ciudad_destino='"& vString2 &"' where id_viaje="& session("viaje") 
        response.write strSQL
        rsUpdateEntry2.Open strSQL, adoCon 
        response.redirect("ShowContent.asp?contenido=88&action=2&msg=entidad destino actualizada")

   end if


end sub


sub ShowRevisionUnidades

    titulo="Resultados revision de unidades" 
    DoTitulo(titulo)

    strSQL="select top 15 ID_UNIDAD,.dbo.fn_GetMonthName(DocDate,'spanish') as Fecha,[IF],FF,II,DI,II2,IT,DT,IT2,DT2,Cuerpo,Vidrios,wippers,FrontLamp,BackLamp,Hazzard,Direccionales,Cabina,AC,       Bateria,Aceite,Frenos,Direccion,Agua,Cinturones,Antifreeze,Filtro,Combustible,TanqueAire,Electrico,ID_check from check_unidad where Mode='in' AND ( [IF]='A' OR FF='A' OR II='A' OR DI='A' OR II2='A' OR IT='A' OR DT='A' OR IT2='A' OR DT2='A' OR Cuerpo='A' OR Vidrios='A' OR wippers='A' OR FrontLamp='A'  OR BackLamp='A' OR Hazzard='A' OR Direccionales='A' OR Cabina='A' OR AC='A' OR Bateria='A' OR Aceite='A' OR Frenos='A' OR Direccion='A' OR Agua='A'   OR Cinturones='A' OR Antifreeze='A' OR Filtro='A' OR Combustible='A' OR TanqueAire='A' OR Electrico='A' ) order by LogDate desc"
      'response.write strSQL

      rsUpdateEntry.Open strSQL, adoCon 
     
      CreateTable(1400)
       response.write("<td class='subtitulo td-c fontmed'>ID</td>")
       response.write("<td class='subtitulo td-c fontmed'>Fecha</td>")
       response.write("<td class='subtitulo td-c fontmed'>IF</td>")
       response.write("<td class='subtitulo td-c fontmed'>FF</td>")
       response.write("<td class='subtitulo td-c fontmed'>II</td>")
       response.write("<td class='subtitulo td-c fontmed'>DI</td>")
       response.write("<td class='subtitulo td-c fontmed'>II2</td>")
       response.write("<td class='subtitulo td-c fontmed'>IT</td>")
       response.write("<td class='subtitulo td-c fontmed'>DT</td>")
       response.write("<td class='subtitulo td-c fontmed'>IT2</td>")
       response.write("<td class='subtitulo td-c fontmed'>cuerpo</td>")
       response.write("<td class='subtitulo td-c fontmed'>vidrios</td>")
       response.write("<td class='subtitulo td-c fontmed'>wippers</td>")
       response.write("<td class='subtitulo td-c fontmed'>FrontLamp</td>")
       response.write("<td class='subtitulo td-c fontmed'>BackLamp</td>")
       response.write("<td class='subtitulo td-c fontmed'>Hazzars</td>")
       response.write("<td class='subtitulo td-c fontmed'>direcs</td>")
       response.write("<td class='subtitulo td-c fontmed'>cabina</td>")
       response.write("<td class='subtitulo td-c fontmed'>A/C</td>")
       response.write("<td class='subtitulo td-c fontmed'>Bat</td>")
       response.write("<td class='subtitulo td-c fontmed'>Oil</td>")
       response.write("<td class='subtitulo td-c fontmed'>brake</td>")
       response.write("<td class='subtitulo td-c fontmed'>direcc</td>")

       response.write("<td class='subtitulo td-c fontmed'>agua</td>")
       response.write("<td class='subtitulo td-c fontmed'>cints</td>")
       response.write("<td class='subtitulo td-c fontmed'>afreez</td>")
       response.write("<td class='subtitulo td-c fontmed'>filtr</td>")
       response.write("<td class='subtitulo td-c fontmed'>combust</td>")
       response.write("<td class='subtitulo td-c fontmed'>TanqAir</td>")
       response.write("<td class='subtitulo td-c fontmed'>Electrc</td>")
      

      while not rsUpdateEntry.EOF
         RowIn
         response.write("<td class=' td-c fontmed'><a href='ShowContent.asp?contenido=96&ID="& rsUpdateEntry("ID_check") &"'>"&rsUpdateEntry("ID_unidad")&"</a></td>")   
         response.write("<td class=' td-c fontmed'>"&rsUpdateEntry("Fecha")&"</td>")
         response.write("<td class=' td-c fontmed'>"&rsUpdateEntry("IF")&"</td>")
         response.write("<td class=' td-c fontmed'>"&rsUpdateEntry("FF")&"</td>")
         response.write("<td class=' td-c fontmed'>"&rsUpdateEntry("II")&"</td>")
         response.write("<td class=' td-c fontmed'>"&rsUpdateEntry("DI")&"</td>")
         response.write("<td class=' td-c fontmed'>"&rsUpdateEntry("II2")&"</td>")
         response.write("<td class=' td-c fontmed'>"&rsUpdateEntry("IT")&"</td>")
         response.write("<td class=' td-c fontmed'>"&rsUpdateEntry("DT")&"</td>")
         response.write("<td class=' td-c fontmed'>"&rsUpdateEntry("IT2")&"</td>")

         response.write("<td class=' td-c fontmed'>"&rsUpdateEntry("cuerpo")&"</td>")
         response.write("<td class=' td-c fontmed'>"&rsUpdateEntry("vidrios")&"</td>")
         response.write("<td class=' td-c fontmed'>"&rsUpdateEntry("wippers")&"</td>")
         response.write("<td class=' td-c fontmed'>"&rsUpdateEntry("FrontLamp")&"</td>")
         response.write("<td class=' td-c fontmed'>"&rsUpdateEntry("BackLamp")&"</td>")
         response.write("<td class=' td-c fontmed'>"&rsUpdateEntry("Hazzard")&"</td>")
         response.write("<td class=' td-c fontmed'>"&rsUpdateEntry("Direccionales")&"</td>")
         response.write("<td class=' td-c fontmed'>"&rsUpdateEntry("cabina")&"</td>")
         response.write("<td class=' td-c fontmed'>"&rsUpdateEntry("AC")&"</td>")
         response.write("<td class=' td-c fontmed'>"&rsUpdateEntry("bateria")&"</td>")
         response.write("<td class=' td-c fontmed'>"&rsUpdateEntry("aceite")&"</td>")
         response.write("<td class=' td-c fontmed'>"&rsUpdateEntry("frenos")&"</td>")
         response.write("<td class=' td-c fontmed'>"&rsUpdateEntry("direccion")&"</td>")
         response.write("<td class=' td-c fontmed'>"&rsUpdateEntry("agua")&"</td>")
         response.write("<td class=' td-c fontmed'>"&rsUpdateEntry("cinturones")&"</td>")
         response.write("<td class=' td-c fontmed'>"&rsUpdateEntry("antifreeze")&"</td>")
         response.write("<td class=' td-c fontmed'>"&rsUpdateEntry("filtro")&"</td>")
         response.write("<td class=' td-c fontmed'>"&rsUpdateEntry("combustible")&"</td>")
         response.write("<td class=' td-c fontmed'>"&rsUpdateEntry("TanqueAire")&"</td>")
         response.write("<td class=' td-c fontmed'>"&rsUpdateEntry("Electrico")&"</td>")

         RowOut
         rsUpdateEntry.MoveNext
      wend
      closeTable
      rsUpdateEntry.close

end sub


sub Check_unidad  'contenido 89'

    if request("control")="" then
 
        strSQL="select a.*,b.tipo_activo,b.*,c.activo,.dbo.fn_GetMonthName(b.Vigencia_Seguro,'spanish') as vigencia from viajes a inner join flotilla b on a.id_unidad=b.id_unidad inner join cat_activos c on b.tipo_activo=c.tipo_activo where a.id_viaje=" &request("ID")
        'response.write strSQL
        rsUpdateEntry.Open strSQL, adoCon 
      
        vMode=""
        select case request("action")
        case "1"  
                  vmode="out"     
                  TITULO="REVISION DE SALIDA DE UNA UNIDAD: "& rsUpdateEntry("ACTIVO")
        case "2"  
                  vmode="in"
                  TITULO="REVISION DE ENTRADA DE UNA UNIDAD: "& rsUpdateEntry("ACTIVO")
        case "3"  
                  vmode="mto"
                  TITULO="REVISION DE MANTENIMIENTO DE UNA UNIDAD: "& rsUpdateEntry("ACTIVO")
        end select

        strSQL="select * from Check_unidad where id_viaje="&request("ID") &" and mode='"& vmode&"' " 
        'response.write (strSQL &"<BR>")
        rsUpdateEntry2.Open strSQL, adoCon 
        
        if rsUpdateEntry2.EOF then
           strSQL="insert into Check_unidad (ID_Check,ID_viaje,Mode,DocDate,Logdate,id_usuario,id_unidad) values ( (select max(ID_Check) from Check_unidad)+1,"& request("ID") &",'"& vmode &"',getdate(),getdate(),'"& session("session_id") &"','"& request("unidad") &"')" 
           'response.write (strSQL &"<BR>")
           rsUpdateEntry3.Open strSQL, adoCon 
        end if

        rsUpdateEntry2.close

         strSQL="select * from Check_unidad where mode='"& vmode &"' and id_viaje="& request("ID")
         rsUpdateEntry2.Open strSQL, adoCon 


           DoAlert
           DoTitulo(TITULO)

           response.write("<form id='checkUnidad' method='post' action='ShowContent.asp'>")
           response.write("<input type='hidden' name='contenido' value='89'>")
           response.write("<input type='hidden' name='control' value='2'>")                  'va a ir a CheckChechUnidad'
           response.write("<input type='hidden' name='ID' value='"& rsUpdateEntry2("id_check") &"'>") 

           CreateTable(900)
           response.write("<tr><td with='40%' class='subtitulo strong'>INTERIOR / EXTERIOR</td><td with='60%' class='subtitulo strong'>DATOS GENERALES</td></tr>")
           
                 select case rsUpdateEntry("tipo_activo")  
                  case  "3"  response.write("<tr><td><img src='/images/camion.png' border=1 alt='' title='' width='300px'>")        'camion'                 
                  case  "2"  response.write("<tr><td><img src='/images/DeCarga.png' border=1 alt='' title='' width='300px'>")       'vehiculo de carga'
                  case  "1"  response.write("<tr><td><img src='/images/utilitario.png' border=1 alt='' title='' width='300px'>")    'utilitario'
                  end select

                 response.write("<td style='vertical-align:top;text-align:center'>")

                      CreateTable(540)
                      response.write("<tr><td class='td-r subtitulo fontmed' width='25%'>ID_UNIDAD:</td><td class='td-l fontmed strong' width='25%'>"& rsUpdateEntry("id_unidad") &"</td>")
                      response.write("<td class='td-r subtitulo fontmed strong' width='25%'>KILOMETRAJE:</td><td class='td-l fontmed' width='25%'><input type='text' name='fkilometraje' placeholder='introduzca kms sin comas'></td></tr>")

                      response.write("<tr><td class='td-r subtitulo fontmed'>Fabricante:</td><td class='td-l fontmed'>"& rsUpdateEntry("Fabricante") &"</td>")
                      response.write("<td class='td-r subtitulo fontmed'>Fecha captura:</td><td class='td-l fontmed'>"& date() &"</td></tr>")

                      response.write("<tr><td class='td-r subtitulo fontmed'>Modelo:</td><td class='td-l fontmed'>"& rsUpdateEntry("Modelo") &"</td>")
                      response.write("<td class='td-r subtitulo fontmed'>Placas:</td><td class='td-l fontmed'>"& rsUpdateEntry("Placas") &"</td></tr>")

                      response.write("<tr><td class='td-r subtitulo fontmed'>Seguro:</td><td class='td-l fontmed'>"& rsUpdateEntry("seguro") &"</td>")
                      response.write("<td class='td-r subtitulo fontmed'>Plataforma:</td><td class='td-l fontmed'>"& rsUpdateEntry("Plataforma") &"</td></tr>")

                      response.write("<tr><td class='td-r subtitulo fontmed'>Vigencia seguro:</td><td class='td-l fontmed'>"& rsUpdateEntry("vigencia") &"</td>")
                      response.write("<td class='td-r subtitulo fontmed'>Total llantas:</td><td class='td-l fontmed'>"& rsUpdateEntry("Total_Llantas") &"</td></tr>")

                      response.write("<tr><td class='td-r subtitulo fontmed'>Combustible:</td><td class='td-l fontmed'>"& rsUpdateEntry("Combustible") &"</td>")
                      response.write("<td class='td-r subtitulo fontmed'>Max carga (kg):</td><td class='td-l fontmed'>"& rsUpdateEntry("MaxCarga") &"</td></tr>")
                      closeTable

                 response.write("</td></tr>")
                 response.write("<tr><td style='vertical-align:top;text-align:center'>")


                  CreateTable(280)
                  response.write("<tr><td class='td-c subtitulo fontmed strong' colspan=2>NEUMATICOS</td></tr>")

                  response.write("<tr><td class='td-r subtitulo fontmed'>IF:</td><td class='td-l fontmed'><select name='fIF'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td></tr>")
                  response.write("<tr><td class='td-r subtitulo fontmed'>FF:</td><td class='td-l fontmed'><select name='fFF'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td></tr>")
                  response.write("<tr><td class='td-r subtitulo fontmed'>II:</td><td class='td-l fontmed'><select name='fII'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td></tr>")      
                   response.write("<tr><td class='td-r subtitulo fontmed'>DI:</td><td class='td-l fontmed'><select name='fDI'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td></tr>")
                  response.write("<tr><td class='td-r subtitulo fontmed'>II2:</td><td class='td-l fontmed'><select name='fII2'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td></tr>")
                  response.write("<tr><td class='td-r subtitulo fontmed'>IT:</td><td class='td-l fontmed'><select name='fIT'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td></tr>")            
                   response.write("<tr><td class='td-r subtitulo fontmed'>DT:</td><td class='td-l fontmed'><select name='fDT'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td></tr>")
                  response.write("<tr><td class='td-r subtitulo fontmed'>IT2:</td><td class='td-l fontmed'><select name='fIT2'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td></tr>")
                  response.write("<tr><td class='td-r subtitulo fontmed'>DT2:</td><td class='td-l fontmed'><select name='fDT2'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td></tr>")                     
                    
                  closeTable
                  response.write("</td><td style='vertical-align:top;text-align:center'>")
                  CreateTable(540)

                   response.write("<tr><td class='td-c subtitulo fontmed strong' colspan=2>VISIBLES</td><td class='td-c subtitulo fontmed strong' colspan=4>FUNCIONES</td></tr>")

                  response.write("<tr><td class='td-r subtitulo fontmed'>Cuerpo exterior:</td><td class='td-l fontmed'><select name='fcuerpo'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td>")                     
                  response.write("<td class='td-r subtitulo fontmed'>Aceite de motor:</td><td class='td-l fontmed'><select name='faceite'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td></tr>")

                  response.write("<tr><td class='td-r subtitulo fontmed'>Vidrios:</td><td class='td-l fontmed'><select name='fvidrios'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td>")                     
                  response.write("<td class='td-r subtitulo fontmed'>L&iacute;quido de frenos:</td><td class='td-l fontmed'><select name='ffrenos'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td></tr>")                           

                  response.write("<tr><td class='td-r subtitulo fontmed'>Wippers:</td><td class='td-l fontmed'><select name='fwippers'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td>")                     
                  response.write("<td class='td-r subtitulo fontmed'>Aceite direccion hidraulica:</td><td class='td-l fontmed'><select name='fhidraulica'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td></tr>")   
 
                  response.write("<tr><td class='td-r subtitulo fontmed'>Luces frontales:</td><td class='td-l fontmed'><select name='fFrontLamp'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td>")                     
                  response.write("<td class='td-r subtitulo fontmed'>Agua wippers:</td><td class='td-l fontmed'><select name='fagua'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td></tr>")                             

                  response.write("<tr><td class='td-r subtitulo fontmed'>Luces traseras:</td><td class='td-l fontmed'><select name='fBackLamp'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td>")                     
                  response.write("<td class='td-r subtitulo fontmed'>Cinturones:</td><td class='td-l fontmed'><select name='fcinturon'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td></tr>")                             

                  response.write("<tr><td class='td-r subtitulo fontmed'>Luces intermitentes:</td><td class='td-l fontmed'><select name='fHazzard'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td>")                     
                  response.write("<td class='td-r subtitulo fontmed'>Antifreeze:</td><td class='td-l fontmed'><select name='fAntifreeze'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td></tr>")    

                  response.write("<tr><td class='td-r subtitulo fontmed'>Direccionales:</td><td class='td-l fontmed'><select name='fDireccionales'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td>")                     
                  response.write("<td class='td-r subtitulo fontmed'>Filtro de aire:</td><td class='td-l fontmed'><select name='fAire'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td></tr>")   

                  response.write("<tr><td class='td-r subtitulo fontmed'>Luces interior cabina:</td><td class='td-l fontmed'><select name='fCabina'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td>")                     
                  response.write("<td class='td-r subtitulo fontmed'>Nivel Combustible:</td><td class='td-l fontmed'><select name='fcombustible'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td></tr>")   

                  response.write("<tr><td class='td-r subtitulo fontmed'>Aire acondicionado:</td><td class='td-l fontmed'><select name='fAC'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td>")                     
                  response.write("<td class='td-r subtitulo fontmed'>Tanques de aire:</td><td class='td-l fontmed'><select name='fTanques'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td></tr>")  

                  response.write("<tr><td class='td-r subtitulo fontmed'>Bater&iacute;a:</td><td class='td-l fontmed'><select name='fBateria'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td>")                     
                  response.write("<td class='td-r subtitulo fontmed'>Cableado el&eacute;ctrico:</td><td class='td-l fontmed'><select name='fCableado'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td></tr>") 
                  
                  closeTable

                          
                   response.write("</td></tr><tr><td colspan=2 class='td-c'><input type='submit' value='registrar'></form></td></tr>")

           closeTable
           rsUpdateEntry.close

           rsUpdateEntry2.close

    end if



    if request("control")="2" then
         strSQL="select case when a.Mode='in' then (select Kilometraje from check_unidad where ID_viaje=a.ID_viaje and Mode='out') else 0 end as flag from check_unidad a where a.ID_check="&request("ID")
         rsUpdateEntry2.Open strSQL, adoCon 

         vFlag=int(trim(rsUpdateEntry2("flag")))
         rsUpdateEntry2.close
  
         if vFlag>0 then
             strSQL="select case when "& trim(request("fkilometraje")) &"<"& trim(vFlag) &"then 1 else 0 end as flag"
             response.write strSQL
             rsUpdateEntry2.Open strSQL, adoCon 

             vFlag=int(trim(rsUpdateEntry2("flag")))
             rsUpdateEntry2.close

             if vFlag=1 then 'resulto que el kilometraje nuevo es menor al con que salio'
                   response.redirect("ShowContent.asp?contenido=88&action=1&msg=no puede registrar un kilometraje menor, del que salio")
             end if

         end if

 
         if trim(request("fkilometraje"))="" then
               response.redirect("ShowContent.asp?contenido=88&action=1&msg=no puede enviar el kilometraje vacio")
         else
               strSQL="update Check_unidad set "
               strSQL= strSQL &"kilometraje="& request("fkilometraje") &","
               strSQL= strSQL &"[IF]='"& request("fIF") &"',"
               strSQL= strSQL &"FF='"& request("fFF") &"',"
               strSQL= strSQL &"II='"& request("fII") &"',"
               strSQL= strSQL &"DI='"& request("fDI") &"',"
               strSQL= strSQL &"II2='"& request("fIT2") &"',"
               strSQL= strSQL &"IT='"& request("fIT") &"',"
               strSQL= strSQL &"DT='"& request("fDT") &"',"
               strSQL= strSQL &"IT2='"& request("fIT2") &"',"
               strSQL= strSQL &"DT2='"& request("fDT2") &"',"
               strSQL= strSQL &"cuerpo='"& request("fcuerpo") &"',"              
               strSQL= strSQL &"vidrios='"& request("fvidrios") &"',"
               strSQL= strSQL &"wippers='"& request("fwippers") &"',"
               strSQL= strSQL &"FrontLamp='"& request("fFrontLamp") &"',"
               strSQL= strSQL &"BackLamp='"& request("fBackLamp") &"',"
               strSQL= strSQL &"hazzard='"& request("fHazzard") &"',"
               strSQL= strSQL &"direccionales='"& request("fDireccionales") &"',"
               strSQL= strSQL &"cabina='"& request("fCabina") &"',"
               strSQL= strSQL &"AC='"& request("fAC") &"',"
               strSQL= strSQL &"bateria='"& request("fBateria") &"',"    
               strSQL= strSQL &"aceite='"& request("faceite") &"',"         
               strSQL= strSQL &"frenos='"& request("ffrenos") &"',"
               strSQL= strSQL &"direccion='"& request("fhidraulica") &"',"
               strSQL= strSQL &"agua='"& request("fagua") &"',"
               strSQL= strSQL &"Cinturones='"& request("fcinturon") &"',"
               strSQL= strSQL &"Antifreeze='"& request("fAntifreeze") &"',"
               strSQL= strSQL &"filtro='"& request("faire") &"',"
               strSQL= strSQL &"combustible='"& request("fcombustible") &"',"
               strSQL= strSQL &"TanqueAire='"& request("fTanques") &"',"          
               strSQL= strSQL &"Electrico='"& request("fCableado") &"',"  

               strSQL= strSQL &"id_usuario='"& session("session_id") &"'," 
               strSQL= strSQL &"LogDate=getdate()   " 
               strSQL= strSQL &"WHERE id_check="& request("ID")

               response.write strSQL
               rsUpdateEntry3.Open strSQL, adoCon 

               strSQL="select a.* from Peajes a inner join Viajes b on a.ciudad_origen=b.ciudad_origen and a.Entidad_Destino=b.Entidad_Destino and a.ciudad_destino=b.ciudad_destino inner join check_unidad c on c.Mode='in' and b.ID_viaje=c.ID_viaje where c.ID_check="&request("ID")
               rsUpdateEntry.Open strSQL, adoCon 
               flag=0
               vRuta=0

               if not rsUpdateEntry.EOF then
                   vRuta=int(trim(rsUpdateEntry("ID")))
                   vejes2=int(trim(rsUpdateEntry("2")))
                   vejes4=int(trim(rsUpdateEntry("4")))
                   if vEje2=0 or vejes4=0 then
                       flag=1
                   end if
               end if
               rsUpdateEntry.close
               
               if flag=1 then
                    response.redirect("ShowContent.asp?contenido=110&action=1&ID="&vRuta&"&msg=se actualizo revision de unidad, por favor actualice costos de peaje" )  
               else
                    response.redirect("ShowContent.asp?contenido=88&action=1&msg=se actualizo revision de unidad")
               end if

         end if



    end if

        
end sub



sub Check_unidad_edit  'contenido 96'  

    if request("control")="" then

        strSQL="select * from Check_unidad where mode='in' and id_check="& request("ID")
        'response.write strSQL
        rsUpdateEntry2.Open strSQL, adoCon 
 
        strSQL="select a.*,b.tipo_activo,b.*,c.activo,.dbo.fn_GetMonthName(b.Vigencia_Seguro,'spanish') as vigencia from viajes a inner join flotilla b on a.id_unidad=b.id_unidad inner join cat_activos c on b.tipo_activo=c.tipo_activo where a.id_viaje=" & rsUpdateEntry2("ID_viaje")
        'response.write strSQL
        rsUpdateEntry.Open strSQL, adoCon 

         

           DoAlert
           Titulo="Actualizando revision de unidad no. "&request("ID") & " / Viaje no. "& rsUpdateEntry2("ID_viaje")
           DoTitulo(TITULO)

           response.write("<form id='checkUnidad' method='post' action='ShowContent.asp'>")
           response.write("<input type='hidden' name='contenido' value='96'>")
           response.write("<input type='hidden' name='control' value='2'>")                  'va a ir a Check_unidad_edit'
           response.write("<input type='hidden' name='ID' value='"& request("ID") &"'>") 

           CreateTable(900)
           response.write("<tr><td with='40%' class='subtitulo strong'>INTERIOR / EXTERIOR</td><td with='60%' class='subtitulo strong'>DATOS GENERALES</td></tr>")
           
                 select case rsUpdateEntry("tipo_activo")  
                  case  "3"  response.write("<tr><td><img src='/images/camion.png' border=1 alt='' title='' width='300px'>")        'camion'                 
                  case  "2"  response.write("<tr><td><img src='/images/DeCarga.png' border=1 alt='' title='' width='300px'>")       'vehiculo de carga'
                  case  "1"  response.write("<tr><td><img src='/images/utilitario.png' border=1 alt='' title='' width='300px'>")    'utilitario'
                  end select

                 response.write("<td style='vertical-align:top;text-align:center'>")

                      CreateTable(540)
                      response.write("<tr><td class='td-r subtitulo fontmed' width='25%'>ID_UNIDAD:</td><td class='td-l fontmed strong' width='25%'>"& rsUpdateEntry("id_unidad") &"</td>")
                      response.write("<td class='td-r subtitulo fontmed strong' width='25%'>KILOMETRAJE:</td><td class='td-l fontmed' width='25%'>"& rsUpdateEntry2("kilometraje") &"</td>")

                      response.write("<tr><td class='td-r subtitulo fontmed'>Fabricante:</td><td class='td-l fontmed'>"& rsUpdateEntry("Fabricante") &"</td>")
                      response.write("<td class='td-r subtitulo fontmed'>Fecha captura:</td><td class='td-l fontmed'>"& date() &"</td></tr>")

                      response.write("<tr><td class='td-r subtitulo fontmed'>Modelo:</td><td class='td-l fontmed'>"& rsUpdateEntry("Modelo") &"</td>")
                      response.write("<td class='td-r subtitulo fontmed'>Placas:</td><td class='td-l fontmed'>"& rsUpdateEntry("Placas") &"</td></tr>")

                      response.write("<tr><td class='td-r subtitulo fontmed'>Seguro:</td><td class='td-l fontmed'>"& rsUpdateEntry("seguro") &"</td>")
                      response.write("<td class='td-r subtitulo fontmed'>Plataforma:</td><td class='td-l fontmed'>"& rsUpdateEntry("Plataforma") &"</td></tr>")

                      response.write("<tr><td class='td-r subtitulo fontmed'>Vigencia seguro:</td><td class='td-l fontmed'>"& rsUpdateEntry("vigencia") &"</td>")
                      response.write("<td class='td-r subtitulo fontmed'>Total llantas:</td><td class='td-l fontmed'>"& rsUpdateEntry("Total_Llantas") &"</td></tr>")

                      response.write("<tr><td class='td-r subtitulo fontmed'>Combustible:</td><td class='td-l fontmed'>"& rsUpdateEntry("Combustible") &"</td>")
                      response.write("<td class='td-r subtitulo fontmed'>Max carga (kg):</td><td class='td-l fontmed'>"& rsUpdateEntry("MaxCarga") &"</td></tr>")
                      closeTable

                 response.write("</td></tr>")
                 response.write("<tr><td style='vertical-align:top;text-align:center'>")


                  CreateTable(280)
                  response.write("<tr><td class='td-c subtitulo fontmed strong' colspan=2>NEUMATICOS</td></tr>")

                  response.write("<tr><td class='td-r subtitulo fontmed'>IF:</td><td class='td-l fontmed'><select name='fIF'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td></tr>")
                  response.write("<tr><td class='td-r subtitulo fontmed'>FF:</td><td class='td-l fontmed'><select name='fFF'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td></tr>")
                  response.write("<tr><td class='td-r subtitulo fontmed'>II:</td><td class='td-l fontmed'><select name='fII'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td></tr>")      
                   response.write("<tr><td class='td-r subtitulo fontmed'>DI:</td><td class='td-l fontmed'><select name='fDI'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td></tr>")
                  response.write("<tr><td class='td-r subtitulo fontmed'>II2:</td><td class='td-l fontmed'><select name='fII2'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td></tr>")
                  response.write("<tr><td class='td-r subtitulo fontmed'>IT:</td><td class='td-l fontmed'><select name='fIT'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td></tr>")            
                   response.write("<tr><td class='td-r subtitulo fontmed'>DT:</td><td class='td-l fontmed'><select name='fDT'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td></tr>")
                  response.write("<tr><td class='td-r subtitulo fontmed'>IT2:</td><td class='td-l fontmed'><select name='fIT2'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td></tr>")
                  response.write("<tr><td class='td-r subtitulo fontmed'>DT2:</td><td class='td-l fontmed'><select name='fDT2'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td></tr>")                     
                    
                  closeTable
                  response.write("</td><td style='vertical-align:top;text-align:center'>")
                  CreateTable(540)

                   response.write("<tr><td class='td-c subtitulo fontmed strong' colspan=2>VISIBLES</td><td class='td-c subtitulo fontmed strong' colspan=4>FUNCIONES</td></tr>")

                  response.write("<tr><td class='td-r subtitulo fontmed'>Cuerpo exterior:</td><td class='td-l fontmed'><select name='fcuerpo'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td>")                     
                  response.write("<td class='td-r subtitulo fontmed'>Aceite de motor:</td><td class='td-l fontmed'><select name='faceite'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td></tr>")

                  response.write("<tr><td class='td-r subtitulo fontmed'>Vidrios:</td><td class='td-l fontmed'><select name='fvidrios'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td>")                     
                  response.write("<td class='td-r subtitulo fontmed'>L&iacute;quido de frenos:</td><td class='td-l fontmed'><select name='ffrenos'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td></tr>")                           

                  response.write("<tr><td class='td-r subtitulo fontmed'>Wippers:</td><td class='td-l fontmed'><select name='fwippers'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td>")                     
                  response.write("<td class='td-r subtitulo fontmed'>Aceite direccion hidraulica:</td><td class='td-l fontmed'><select name='fhidraulica'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td></tr>")   
 
                  response.write("<tr><td class='td-r subtitulo fontmed'>Luces frontales:</td><td class='td-l fontmed'><select name='fFrontLamp'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td>")                     
                  response.write("<td class='td-r subtitulo fontmed'>Agua wippers:</td><td class='td-l fontmed'><select name='fagua'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td></tr>")                             

                  response.write("<tr><td class='td-r subtitulo fontmed'>Luces traseras:</td><td class='td-l fontmed'><select name='fBackLamp'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td>")                     
                  response.write("<td class='td-r subtitulo fontmed'>Cinturones:</td><td class='td-l fontmed'><select name='fcinturon'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td></tr>")                             

                  response.write("<tr><td class='td-r subtitulo fontmed'>Luces intermitentes:</td><td class='td-l fontmed'><select name='fHazzard'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td>")                     
                  response.write("<td class='td-r subtitulo fontmed'>Antifreeze:</td><td class='td-l fontmed'><select name='fAntifreeze'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td></tr>")    

                  response.write("<tr><td class='td-r subtitulo fontmed'>Direccionales:</td><td class='td-l fontmed'><select name='fDireccionales'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td>")                     
                  response.write("<td class='td-r subtitulo fontmed'>Filtro de aire:</td><td class='td-l fontmed'><select name='fAire'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td></tr>")   

                  response.write("<tr><td class='td-r subtitulo fontmed'>Luces interior cabina:</td><td class='td-l fontmed'><select name='fCabina'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td>")                     
                  response.write("<td class='td-r subtitulo fontmed'>Nivel Combustible:</td><td class='td-l fontmed'><select name='fcombustible'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td></tr>")   

                  response.write("<tr><td class='td-r subtitulo fontmed'>Aire acondicionado:</td><td class='td-l fontmed'><select name='fAC'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td>")                     
                  response.write("<td class='td-r subtitulo fontmed'>Tanques de aire:</td><td class='td-l fontmed'><select name='fTanques'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td></tr>")  

                  response.write("<tr><td class='td-r subtitulo fontmed'>Bater&iacute;a:</td><td class='td-l fontmed'><select name='fBateria'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td>")                     
                  response.write("<td class='td-r subtitulo fontmed'>Cableado el&eacute;ctrico:</td><td class='td-l fontmed'><select name='fCableado'>")
                           response.write("<option value='F'>Funcional</option><option value='A'>Requere atenci&oacute;n</option>")
                           response.write("</select></td></tr>") 
                  
                  closeTable

                          
                   response.write("</td></tr><tr><td colspan=2 class='td-c'><input type='submit' value='registrar'></form></td></tr>")

           closeTable
           rsUpdateEntry.close

           rsUpdateEntry2.close

    end if




    if request("control")="2" then
         
               strSQL="update Check_unidad set "              
               strSQL= strSQL &"[IF]='"& request("fIF") &"',"
               strSQL= strSQL &"FF='"& request("fFF") &"',"
               strSQL= strSQL &"II='"& request("fII") &"',"
               strSQL= strSQL &"DI='"& request("fDI") &"',"
               strSQL= strSQL &"II2='"& request("fIT2") &"',"
               strSQL= strSQL &"IT='"& request("fIT") &"',"
               strSQL= strSQL &"DT='"& request("fDT") &"',"
               strSQL= strSQL &"IT2='"& request("fIT2") &"',"
               strSQL= strSQL &"DT2='"& request("fDT2") &"',"
               strSQL= strSQL &"cuerpo='"& request("fcuerpo") &"',"              
               strSQL= strSQL &"vidrios='"& request("fvidrios") &"',"
               strSQL= strSQL &"wippers='"& request("fwippers") &"',"
               strSQL= strSQL &"FrontLamp='"& request("fFrontLamp") &"',"
               strSQL= strSQL &"BackLamp='"& request("fBackLamp") &"',"
               strSQL= strSQL &"hazzard='"& request("fHazzard") &"',"
               strSQL= strSQL &"direccionales='"& request("fDireccionales") &"',"
               strSQL= strSQL &"cabina='"& request("fCabina") &"',"
               strSQL= strSQL &"AC='"& request("fAC") &"',"
               strSQL= strSQL &"bateria='"& request("fBateria") &"',"    
               strSQL= strSQL &"aceite='"& request("faceite") &"',"         
               strSQL= strSQL &"frenos='"& request("ffrenos") &"',"
               strSQL= strSQL &"direccion='"& request("fhidraulica") &"',"
               strSQL= strSQL &"agua='"& request("fagua") &"',"
               strSQL= strSQL &"Cinturones='"& request("fcinturon") &"',"
               strSQL= strSQL &"Antifreeze='"& request("fAntifreeze") &"',"
               strSQL= strSQL &"filtro='"& request("faire") &"',"
               strSQL= strSQL &"combustible='"& request("fcombustible") &"',"
               strSQL= strSQL &"TanqueAire='"& request("fTanques") &"',"          
               strSQL= strSQL &"Electrico='"& request("fCableado") &"',"  

               strSQL= strSQL &"id_usuario='"& session("session_id") &"'," 
               strSQL= strSQL &"LogDate=getdate()   " 
               strSQL= strSQL &"WHERE id_check="& request("ID")

               response.write strSQL
               rsUpdateEntry3.Open strSQL, adoCon 

               response.redirect("ShowContent.asp?contenido=88&action=1&msg=se registro revision de unidad")

   
    end if

        
end sub





sub RegistroEmbarques   'contenido 90' 

 if request("action")="" then

   if request("ffecha")<>"" and request("fRS")<>""   then

    if request("fremision")<>"" then
          strSQL="select  * from envios where RazonSocial='"& request("fRS") &"' and DocNum=" & request("fremision") &" and tipo='remision'" 
          rsUpdateEntry2.Open strSQL, adoCon4

          strSQL="select  top 1 a.cardcode,a.cardname,b.WhsCode from ODLN a inner join DLN1 b on a.DocEntry=b.DocEntry where a.docnum="& request("fremision") 
          vfecha=date()
          vfecha=request("ffecha")

          select case request("fRS")
            case "DMX"  rsUpdateEntry.Open strSQL, adoCon2    'DMX'
            case "DFW"  rsUpdateEntry.Open strSQL, adoCon3
            case "MME"  rsUpdateEntry.Open strSQL, adoCon5
          end select 

          if request("fentidad")="" then
               response.redirect("ShowContent.asp?contenido=90&msg=debe seleccionar entidad y ciudad destino")
          end if

          strSQL="set dateformat ymd;insert into envios (ID,entidad_destino,mpio_destino,ciudad_destino,DocDate,ID_empleado,LogDate,ID_USUARIO,RazonSocial,DocNum,Tipo,WhsCode,id_Portador,Guia,Rastreo,subTotal,Cardcode,Cardname,Comentarios,ocurre,reenvio,CargoCliente) values ( (select max(ID) from envios)+1,'"& request("fentidad") &"','"& request("fmunicipio") &"','"& request("fciudad") &"','"

          strSQL= strSQL & year(vfecha) &"-" & month(vfecha) &"-" & day(vfecha) &"',"
          strSQL= strSQL & request("fempleado") &",getdate(),'"& session("session_id") &"','" & request("fRS") &"'," & request("fremision") &",'remision','" & rsUpdateEntry("WhsCode") &"',"& request("fCarrier") &",'"
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

          
          
          if rsUpdateEntry2.EOF then   'que no esta repetido el envio'
               response.write strSQL
               rsUpdateEntry7.Open strSQL, adoCon4  'se crea el envio'
               'Response.End
               
               '-- PRORRATEA VIATICOS nuevamente'
               strSQL="exec SP_Prorratea_viaticos"
               rsUpdateEntry3.Open strSQL, adoCon

               if not rsUpdateEntry.EOF then  'quiere decir que se encontro la remision'

                    strSQL="select top 1 Pedido from Notificaciones where tipo='remision' and RazonSocial='"&  request("fRS")&"' and DocNum="&  request("fremision") 

                    rsUpdateEntry4.Open strSQL, adoCon4

                    if not rsUpdateEntry4.EOF then

                       if trim(rsUpdateEntry4("Pedido"))<>"" and trim(rsUpdateEntry4("Pedido"))<>"0" then

                           strSQL="exec [dbo].[SP_AnalisisUtilidad] "& rsUpdateEntry4("Pedido") &",'"& request("fRS") &"','"& session("session_id") &"'"
                           rsUpdateEntry4.close
                           rsUpdateEntry4.Open strSQL, adoCon4

                       end if

                    end if

               end if
            

               response.redirect("ShowContent.asp?contenido=90&msg=el envio ha sido registrado")
              
          else
              if request("freenvio")="on" then
                  rsUpdateEntry5.Open strSQL, adoCon4  'se crea el envio'
                  response.write("<P class='alert td-c'>La informacion ha sido registrada</P>")
              else
                 response.write("<P class='alert td-c'>NO SE REGISTR&Oacute; SU ENV&Iacute;O, YA EST&Aacute; REGISTRADO UN ENV&Iacute;O PARA ESTA REMISI&Oacute;N </P>")
              end if
          end if
    else
          
             vfecha=date()
             vfecha=request("ffecha")

             if request("fentidad")="" then
               response.redirect("ShowContent.asp?contenido=90&msg=debe seleccionar entidad y ciudad destino")
             end if

              strSQL="set dateformat ymd;insert into envios (ID,entidad_destino,mpio_destino,ciudad_destino,DocDate,ID_empleado,LogDate,ID_USUARIO,RazonSocial,id_Portador,Guia,Rastreo,subTotal,Comentarios,reenvio,CargoCliente) values ( (select max(ID) from envios)+1, '"& request("fentidad") &"','"& request("fmunicipio") &"','"& request("fciudad") &"','"

              strSQL= strSQL & year(vfecha) &"-" & month(vfecha) &"-" & day(vfecha) &"',"
              strSQL= strSQL & request("fempleado") &",getdate(),'"& session("session_id") &"','" & request("fRS") &"',"& request("fCarrier") &",'"
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





    
      doAlert

      response.write("<div id='destino'>")
      titulo="Registro de uso de paqueter&iacute;a &oacute; mensajer&iacute;a"
      Dotitulo(titulo)
      CreateTable(600)
      response.write("<form id='envios' method='POST' action='ShowContent.asp'>")
      response.write("<input type='hidden' name='contenido' value='90'>")
      

      if request("fentidad")<>"" and request("fCiudadDestino")<>"" then
           vString=request("fCiudadDestino")
           'response.write( vString &"<br>")
           vPos=inStr(vString,"*")
           vString1=mid(vString,1,Vpos-1)
           vString2=mid(vString,Vpos+1,30)

           response.write("<input type='hidden' name='fentidad' value='"& request("fentidad") &"'>")
           response.write("<input type='hidden' name='fmunicipio' value='"& vString1 &"'>")
           response.write("<input type='hidden' name='fciudad' value='"& vString2 &"'>")

           response.write("<tr><td class='subtitulo td-r'>Ciudad destino:</td><td class='td-l'>"& vString2 &"</td></tr>")

      else
      %> 
     <tr><td class='subtitulo td-r'>Entidad Federativa Destino:</td><td class='td-l fontmed'>  

           <select id="entidad_destino" onchange="Javascript:PassValueChangeDiv('entidad_destino','destino','/modules/FormChangeStateEnvios.asp')">   <%  'elemento,divn,modulo  

           strSQL="select * from cat_entidades order by id_entidad"
           'response.write strSQL  
           rsUpdateEntry2.Open strSQL, adoCon 

          while not rsUpdateEntry2.EOF
              response.write("<option value='"& rsUpdateEntry2("id_entidad") &"'>"& rsUpdateEntry2("entidad")  &"</option>")             
              rsUpdateEntry2.MoveNext
          wend
          rsUpdateEntry2.Close
          response.write("</select></td></tr>")    

      end if

      strSQL="select id_usuario,ID,Nombres+' '+paterno+' '+Materno as nombre from Empleados where FechaEgreso is null order by nombre"
      'response.write strSQL  
      rsUpdateEntry7.Open strSQL, adoCon

      %>


     <tr><td class='fontbig subtitulo td-r' width="40%">Raz&oacute;n Social: </td><td class='fontbig td-l' width="60%">
         <select id="fRS" name="fRS"> 
            <option value="DMX">DMX</option> 
            <option value="MME">MME</option> 
            <option value="DFW">DFW</option> 
          </select></td></tr>
     <tr><td class='fontbig subtitulo td-r'># Remisi&oacute;n: </td><td class='fontbig td-l'><input type="number" id="fremision" name="fremision"></td></tr>
     <tr><td class='fontbig subtitulo td-r'>Colaborador que captura: </td><td class='fontbig td-l'><select name="fempleado" id="fempleado">
     <%
        while not rsUpdateEntry7.EOF
        	if UCASE(session("session_id"))=UCASE(rsUpdateEntry7("id_usuario")) then
                 response.write("<option value='"& rsUpdateEntry7("ID") &"' SELECTED>"& rsUpdateEntry7("nombre") &"</option>")
            else
                response.write("<option value='"& rsUpdateEntry7("ID") &"'>"& rsUpdateEntry7("nombre") &"</option>")
            end if
            
            rsUpdateEntry7.MoveNext
        wend
        rsUpdateEntry7.close
     %>
     </select></td></tr>
     
     <tr><td class='fontbig subtitulo td-r'># de gu&iacute;a: </td><td class='fontbig td-l'><input type="text" id="fguia" name="fguia"></td></tr>
     <tr><td class='fontbig subtitulo td-r'># de rastreo: </td><td class='fontbig td-l'><input type="text" id="frastreo" name="frastreo"></td></tr>

     <tr><td class='fontbig subtitulo td-r'>Subtotal sin impuestos: </td><td class='fontbig td-l'><input style="width: 200px" type="text" id="ftotal" name="ftotal" placeholder="no puede ir vac&iacute;o, no use comas"> </td></tr>
   
     
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
              <option value="003-Mty">003-Mty</option>
             </select>  </td></tr>


     <tr><td class='fontbig subtitulo td-r'>Re-env&iacute;o: </td><td class='fontbig td-l'><input type="checkbox" name="freenvio" id="freenvio"> (para un segundo env&iacute;o)</td></tr>    
     <tr><td colspan=2 class='td-c'>  <div id="box">&nbsp;</div> </td></tr>
     </form>
     </table>  
</div>
<P>&nbsp;</P>
<%
    titulo="&Uacute;ltimos registros"
    DoTitulo(titulo)
%>

    <form id="buscar" method="POST" action="ShowContent.asp">
    <input type="hidden" name="contenido" value="90">
    <H2 class="td-c fontmed">Buscar una gu&iacute;a o un rastero: <input type="text" name="vstring" value="<%=request("vstring")%>"> &nbsp;&nbsp;&nbsp;<input type="submit" value="buscar"></form> </H2>
    <H2 class="td-c fontmed">&nbsp;</H2>




 <table width="1260px" cellpadding="4" cellspacing="2" align="center" border=0>


<%
  Const adCmdText = &H0001
  Const adOpenStatic = 3
  nPage=0
  registros=1

  if request("vstring")<>"" and len(request("vstring"))>2 then
      strSQL="select  a.*,b.Portador,dbo.fn_GetMonthName(a.DocDate,'spanish') AS [DocDate2] from envios a inner join cat_portadores b on a.id_Portador=b.id_Portador left join INTRANET.dbo.cat_entidades c on a.entidad_destino collate Modern_Spanish_CI_AS = c.id_entidad collate Modern_Spanish_CI_AS where ( guia like '%"&request("vstring")&"%' OR rastreo like '%"&request("vstring")&"%' )  order by a.ID desc" 
  else
      
      strSQL="select  a.*,b.Portador,dbo.fn_GetMonthName(a.DocDate,'spanish') AS [DocDate2],c.entidad from envios a inner join cat_portadores b on a.id_Portador=b.id_Portador left join INTRANET.dbo.cat_entidades c on a.entidad_destino collate Modern_Spanish_CI_AS = c.id_entidad collate Modern_Spanish_CI_AS where DATEDIFF(DAY,a.DocDate,getdate())<40 order by a.ID desc" 
  end if
  'response.write  strSQL
  rsUpdateEntry.Open strSQL, adoCon4, adOpenStatic, adCmdText

  rsUpdateEntry.PageSize = 16 
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
                response.write("<a href='ShowContent.asp?contenido=90&vPage="&i &"'>") 
                response.write( i &"</A>&nbsp;&nbsp;&nbsp;&nbsp;") 
            end if
        next
        response.write("</td></tr>")    %>

 <tr>
      <td class="td-c subtitulo fontmed">ID</td>
      <td class="td-c subtitulo fontmed">Colaborador</td>
      <td class="td-c subtitulo fontmed" width='80px'>Fecha</td>
      <td class="td-c subtitulo fontmed">RS</td>
      <td class="td-c subtitulo fontmed">Dcto</td>
      <td class="td-c subtitulo fontmed">Tipo</td>
      <td class="td-c subtitulo fontmed">Almac&eacute;n <br>origen</td>
      <td class="td-c subtitulo fontmed">Portador</td>
      <td class="td-c subtitulo fontmed">Gu&iacute;a</td>
      <td class="td-c subtitulo fontmed">Rastreo</td>
      <td class="td-c subtitulo fontmed" width='80px'>Cargo <br> cliente</td>     
      <td class="td-c subtitulo fontmed">Ciudad</td>
      <td class="td-c subtitulo fontmed">Conso-<br>lidado</td>
      <td class="td-c subtitulo fontmed">Pedido</td>
      <td class="td-c subtitulo fontmed">&nbsp;&nbsp;&nbsp;&nbsp;</td>
      <td class="td-c subtitulo fontmed">Re env&iacute;o</td>

    </tr>
    <%        

  vfecha1=date()-5
  vfecha2=date()


  while not rsUpdateEntry.EOF AND registros<=16
     response.write("<tr>")

     response.write("<td class='td-r fontmed'><a href='ShowContent.asp?contenido=90&action=2&ID="& rsUpdateEntry("ID") &"'><img src='/img/b_edit.png' with='11px' alt='editar' title='editar' height='11px' border='0'>"& rsUpdateEntry("ID") &"</a></td>")   

     strSQL="select  Nombres+' '+paterno as Colaborador from Empleados where id=" &   rsUpdateEntry("ID_empleado")
     rsUpdateEntry4.Open strSQL, adoCon

     response.write("<td class='td-l fontmed'>"& mid(rsUpdateEntry4("Colaborador"),1,20) &"</td>")
     rsUpdateEntry4.close
     response.write("<td class='td-r fontmed'>"& rsUpdateEntry("DocDate2") &"</td>")
     response.write("<td class='td-c fontmed'>"& rsUpdateEntry("RazonSocial") &"</td>")
     response.write("<td class='td-r fontmed'>"& rsUpdateEntry("DocNum") &"</td>")
     response.write("<td class='td-c fontmed'>"& rsUpdateEntry("Tipo") &"</td>")
     response.write("<td class='td-c fontmed'>"& rsUpdateEntry("WhsCode") &"</td>")
     response.write("<td class='td-l fontmed'>"& rsUpdateEntry("Portador") &"</td>")
     response.write("<td class='td-r fontmed'>"& rsUpdateEntry("Guia") &"</td>")

     if len(rsUpdateEntry("Rastreo"))>4 then
     	   select case trim(rsUpdateEntry("Portador"))
     	   	   case "Paquete Express" response.write("<td class='td-r fontmed'><a href='https://www.paquetexpress.com.mx/rastreo/"& rsUpdateEntry("Rastreo")&"' target='rastreo'>"& rsUpdateEntry("Rastreo") &"</a></td>")
     	   	   case "Tres Guerras" response.write("<td class='td-r fontmed'>B-"& rsUpdateEntry("Rastreo") &"</td>")
     	   	   case "Castores" response.write("<td class='td-r fontmed'><a href='ttps://www.castores.com.mx/rastreo/"&rsUpdateEntry("Rastreo")&"' target='rastreo'>"& rsUpdateEntry("Rastreo") &"</a></td>")
     	   end select

     else
          response.write("<td class='td-r fontmed'>"& rsUpdateEntry("Rastreo") &"</td>")
     end if


     if rsUpdateEntry("CargoCliente") =1 then
         response.write("<td class='td-c fontmed'> si</td>")
     else
        response.write("<td class='td-c fontmed'>no</td>")
     end if

     
     response.write("<td class='td-l fontmed'>"& rsUpdateEntry("ciudad_destino") &"("& rsUpdateEntry("entidad") &")</td>")



     if rsUpdateEntry("Consolidado") =1 then
         response.write("<td class='td-c fontmed'> si</td>")
         response.write("<td class='td-l fontmed'> "& rsUpdateEntry("Pedido") &"</td>")
         response.write("<td class='td-l fontmed'>&nbsp;</td>")
         response.write("<tr><td class='td-r fontmed strong' colspan=3>[Comentarios:]</td><td class='td-l fontmed' colspan=6>"& rsUpdateEntry("Comentarios") &"</td><td class='td-r fontmed strong' colspan=2>[Socio negocio:]</td><td class='td-l fontmed' colspan=6>"& rsUpdateEntry("Cardname") &"</td></tr><tr><td colspan=17><HR></td></tr>")
     else
        response.write("<td class='td-c fontmed'>no</td>")
        response.write("<td class='td-l fontmed'> "& rsUpdateEntry("Pedido") &"</td>")
        vfecha2=cDate( day(rsUpdateEntry("DocDate"))&"/" & month(rsUpdateEntry("DocDate"))&"/" & year(rsUpdateEntry("DocDate")) )
        if vfecha2>vfecha1 then 
             response.write("<td class='td-l fontmed'><a href='ShowContent.asp?contenido=90&action=4&ID="& rsUpdateEntry("ID") &"'><img src='/img/b_drop.png' with='11px' alt='borrar' title='borrar' height='11px' border=0></a></td>")
        else
             response.write("<td class='td-l fontmed'>&nbsp;</td>")
        end if

        response.write("<td class='td-l fontmed'><a href='ShowContent.asp?contenido=90&action=6&ID="&rsUpdateEntry("ID") &"'><img src='/images/reEnvio.png' border='0' title='re envie email helpdesk' width='30px'></a> </td><tr><td class='td-r fontmed strong' colspan=3>[Comentarios:]</td><td class='td-l fontmed' colspan=6>"& rsUpdateEntry("Comentarios") &"</td><td class='td-r fontmed strong' colspan=2>[Socio negocio:]</td><td class='td-l fontmed' colspan=6>"& rsUpdateEntry("Cardname") &"</td></tr><tr><td colspan=17><HR></td></tr>")



     end if

     
      
     response.write("</tr>")
     rsUpdateEntry.MoveNext
     registros=registros+1

  wend
  response.write("</table><P>&nbsp;</P>") 

end if



if request("action")="2" then  'form edit'
   
    response.write("<form id='edit' method='POST' action='ShowContent.asp'>")

    response.write("<input type='hidden' name='contenido' value='90'>")
    response.write("<input type='hidden' name='ID' id='ID' value='"&  request("ID") &"'>")
    response.write("<input type='hidden' name='action' id='action' value='3'>")

    strSQL="select * from envios where ID=" & request("ID")
    rsUpdateEntry.Open strSQL, adoCon4  

    strSQL="select ID,Nombres+' '+paterno+' '+Materno as nombre from Empleados where FechaEgreso is null order by nombre"
    'response.write strSQL  
    rsUpdateEntry2.Open strSQL, adoCon      

    titulo="Editar registro de embarque no. "& rsUpdateEntry("ID") 
    DoTitulo(titulo)
     %>

     <table border=0 width="600px" cellpadding="2" cellpadding="2" align="center">
     <form id="envios" method="POST" action="envios.asp"> 
     <tr><td class='fontbig subtitulo td-r' width="40%">Raz&oacute;n Social: </td><td class='fontbig td-l' width="60%"><B><%=rsUpdateEntry("RazonSocial")%></B></td></tr>
     <tr><td class='fontbig subtitulo td-r'># Remisi&oacute;n: </td><td class='fontbig td-l'><B><%=rsUpdateEntry("DocNum")%></B></td></tr>
     <tr><td class='fontbig subtitulo td-r'>Colaborador que captura: </td><td class='fontbig td-l'><select name="fempleado" id="fempleado">
     <%
        while not rsUpdateEntry2.EOF
            if rsUpdateEntry("ID_empleado")=rsUpdateEntry2("ID") then
               response.write("<option value='"& rsUpdateEntry2("ID") &"' SELECTED>"& rsUpdateEntry2("nombre") &"</option>")
            else
               response.write("<option value='"& rsUpdateEntry2("ID") &"'>"& rsUpdateEntry2("nombre") &"</option>")
            end if
            rsUpdateEntry2.MoveNext
        wend
        rsUpdateEntry2.close
     %>
     </select></td></tr>
     
     <tr><td class='fontbig subtitulo td-r'># de gu&iacute;a: </td><td class='fontbig td-l'><input type="text" id="fguia" name="fguia" value="<%=rsUpdateEntry("Guia")%>"></td></tr>
     <tr><td class='fontbig subtitulo td-r'># de rastreo: </td><td class='fontbig td-l'><input type="text" id="frastreo" name="frastreo" value="<%=rsUpdateEntry("rastreo")%>"></td></tr>

     <tr><td class='fontbig subtitulo td-r'>Subtotal sin impuestos: </td><td class='fontbig td-l'><input type="text" id="ftotal" name="ftotal" value="<%=rsUpdateEntry("subTotal")%>"></td></tr>
   
     
     <tr><td class='fontbig subtitulo td-r'>Fecha del env&iacute;o: </td><td class='fontbig td-l'> D&iacute;a:  <select name=vDia> <%

              for i=1 to 31
                    if day(rsUpdateEntry("DocDate"))=i then
                       response.write("<option value="& i &" SELECTED>"& i &"</option>")
                   else
                       response.write("<option value="& i &" >"& i &"</option>")
                   end if
               next
               response.write("</select> Mes: <select name=vMes>")
               for i=1 to 12
                    if month(rsUpdateEntry("DocDate"))=i then
                       response.write("<option value="& i &" SELECTED>"& meses(i) &"</option>")
                   else
                       response.write("<option value="& i &" >"&  meses(i) &"</option>")
                   end if
               next
               response.write("</select> A&ntilde;o: <select name=vAnio>")
               for i=year(date())-1 to year(date())
                   if year(rsUpdateEntry("DocDate"))=i then
                       response.write("<option value="& i &" SELECTED>"& i &"</option>")
                   else
                       response.write("<option value="& i &" >"& i &"</option>")
                   end if
               next
               response.write("</select> </td></tr>  ")       %>

    
     <tr><td class='fontbig subtitulo td-r'>Portador: </td><td class='fontbig td-l'><select name="fcarrier" id="fcarrier" onchange="ValidaRemision()">
          <%
          strSQL="select * from cat_portadores order by ID_portador"
          rsUpdateEntry3.Open strSQL, adoCon4 

          while not rsUpdateEntry3.EOF
              if rsUpdateEntry("ID_portador")=rsUpdateEntry3("ID_portador") then
                   response.write("<option value="& rsUpdateEntry3("ID_portador") &" SELECTED>"& rsUpdateEntry3("Portador") &"</option>")
              else
                   response.write("<option value="& rsUpdateEntry3("ID_portador") &">"& rsUpdateEntry3("Portador") &"</option>")
              end if
              rsUpdateEntry3.MoveNext
          wend
          rsUpdateEntry3.close


%>
        </select></td></tr>

     <tr><td class='fontbig subtitulo td-r'>Cargo a Cliente: </td><td class='fontbig td-l'>
            <%
               if  rsUpdateEntry("CargoCliente")=1 then
                   response.write("<input type='checkbox'name='fCargoCliente' id='fCargoCliente' CHECKED>")
               else
                  response.write("<input type='checkbox'name='fCargoCliente' id='fCargoCliente'>")
               end if
            %>
       </td></tr>
     
     <tr><td class='fontbig subtitulo td-r'>Comentarios: </td><td class='fontbig td-l'><input type="text" id="fnotas" name="fnotas" class='anchox2' value="<%=rsUpdateEntry("Comentarios")%>"> </td></tr>
     <tr><td class='fontbig subtitulo td-r'>Re-env&iacute;o: </td><td class='fontbig td-l'>
     <%
               if  rsUpdateEntry("reenvio")=1 then
                   response.write("<input type='checkbox'name='freenvio' id='freenvio' CHECKED>")
               else
                  response.write("<input type='checkbox'name='freenvio' id='freenvio'>")
               end if
      %>
      (para un segundo env&iacute;o) </td></tr>
      <tr><td colspan=2 class='td-c'> &nbsp; </td></tr>
     <tr><td colspan=2 class='td-c'> <input type="submit" value="actualizar" >   </td></tr>
     </form>
     </table>

<%
end if


if request("action")="3" then  'edit  UP'

     strSQL="set dateformat ymd;update envios set ID_empleado="&request("fempleado") &",guia='"& request("fGuia") &"',rastreo='" & request("frastreo") &"',subTotal=" & request("fTotal") &","
     strSQL = strSQL & "DocDate='" & request("vAnio")&"-" & request("vMes")&"-" & request("vDia") &"',id_Portador=" & request("fCarrier") &",CargoCliente="
     if request("fCargoCliente")="on" then
          strSQL = strSQL & "1,"
     else
          strSQL = strSQL & "0,"
     end if
     strSQL = strSQL & "Comentarios='" & UCASE(request("fNotas")) &"',reenvio="
     if request("freenvio")="on" then
          strSQL = strSQL & "1 "
     else
          strSQL = strSQL & "0 "
     end if
     strSQL = strSQL & " where ID=" & request("ID")
     response.write strSQL
     rsUpdateEntry.Open strSQL, adoCon4

     response.redirect("ShowContent.asp?contenido=90&msg=la informacion ha sido actualizada")
end if




if request("action")="4" then  'Delete form'
      
    response.write("<form id='delete' method='POST' action='ShowContent.asp'>")

    response.write("<input type='hidden' name='contenido'  value='90'>")
    response.write("<input type='hidden' name='ID' id='ID' value='"&  request("ID") &"'>")
    response.write("<input type='hidden' name='action' id='action' value='5'>")

    strSQL="select a.*,b.Portador from envios a inner join cat_portadores b on a.id_Portador=b.id_Portador where a.ID=" & request("ID")
    rsUpdateEntry.Open strSQL, adoCon4  

    strSQL="select ID,Nombres+' '+paterno+' '+Materno as nombre from Empleados where id=" & rsUpdateEntry("ID_empleado")
    'response.write strSQL  
    rsUpdateEntry2.Open strSQL, adoCon      

  
    titulo="Borrar registro de embarque no. "& rsUpdateEntry("ID")
    DoTitulo(titulo)
    %>

     <table border=0 width="600px" cellpadding="2" cellpadding="2" align="center">
     <form id="envios" method="POST" action="envios.asp"> 
     <tr><td class='fontbig subtitulo td-r' width="40%">Raz&oacute;n Social: </td><td class='fontbig td-l' width="60%"><B><%=rsUpdateEntry("RazonSocial")%></B></td></tr>
     <tr><td class='fontbig subtitulo td-r'># Remisi&oacute;n: </td><td class='fontbig td-l'><B><%=rsUpdateEntry("DocNum")%></B></td></tr>
     <tr><td class='fontbig subtitulo td-r'>Colaborador que captura: </td><td class='fontbig td-l'><%=rsUpdateEntry2("nombre")%></td></tr>
     
     <tr><td class='fontbig subtitulo td-r'># de gu&iacute;a: </td><td class='fontbig td-l'><%=rsUpdateEntry("Guia")%></td></tr>
     <tr><td class='fontbig subtitulo td-r'># de rastreo: </td><td class='fontbig td-l'><%=rsUpdateEntry("rastreo")%></td></tr>
     <tr><td class='fontbig subtitulo td-r'>Subtotal sin impuestos: </td><td class='fontbig td-l'><%=rsUpdateEntry("subTotal")%></td></tr>
     <tr><td class='fontbig subtitulo td-r'>Fecha del env&iacute;o: </td><td class='fontbig td-l'><%=rsUpdateEntry("DocDate")%> </td></tr>
     <tr><td class='fontbig subtitulo td-r'>Portador: </td><td class='fontbig td-l'><%=rsUpdateEntry("Portador")%></td></tr>
     <tr><td class='fontbig subtitulo td-r'>Cargo a Cliente: </td><td class='fontbig td-l'>
            <%
               if  rsUpdateEntry("CargoCliente")=1 then
                   response.write("<input type='checkbox'name='fCargoCliente' id='fCargoCliente' CHECKED>")
               else
                  response.write("<input type='checkbox'name='fCargoCliente' id='fCargoCliente'>")
               end if
            %>
       </td></tr>
     
     <tr><td class='fontbig subtitulo td-r'>Comentarios: </td><td class='fontbig td-l'><%=rsUpdateEntry("Comentarios")%> </td></tr>
     <tr><td class='fontbig subtitulo td-r'>Re-env&iacute;o: </td><td class='fontbig td-l'>
     <%
               if  rsUpdateEntry("reenvio")=1 then
                   response.write("<input type='checkbox'name='freenvio' id='freenvio' CHECKED>")
               else
                  response.write("<input type='checkbox'name='freenvio' id='freenvio'>")
               end if
      %>
      (para un segundo env&iacute;o) </td></tr>
     <tr><td colspan=2 class='td-c'> &nbsp; </td></tr>
     <tr><td colspan=2 class='td-c'> <input type="submit" value="Borrar registro" >  &nbsp;&nbsp;&nbsp;&nbsp; <a href="/modules/ShowContent.asp?contenido=90">cancelar </a></td></tr>
     </form>
     </table>
   </div>
</div>
<%

end if


if request("action")="5" then  'Delete up'
      
     strSQL="delete envios where ID=" & request("ID")
     rsUpdateEntry.Open strSQL, adoCon4

     response.write strSQL
     response.redirect("ShowContent.asp?contenido=90&msg=el registro ha sido eliminado")
end if




if request("action")="6" then  're-envio de email' 
      
     strSQL="select a.*,b.portador from envios a inner join cat_portadores b on a.id_portador=b.id_portador where a.ID=" & request("ID")
     'response.write strSQL
     rsUpdateEntry.Open strSQL, adoCon4

     titulo="RE ENVIAR EMAIL HELPDESK DE EMBARQUE"     
     doTitulo(titulo)    

     response.write("<form id='reEnvio' method='POST' action='ShowContent.asp'>")

    response.write("<input type='hidden' name='contenido'  value='90'>")
    response.write("<input type='hidden' name='ID' id='ID' value='"&  rsUpdateEntry("ID") &"'>")
    response.write("<input type='hidden' name='action' id='action' value='7'>")

    CreateTable(470)    

     'response.write("<tr><td class='subtitulo td-r'>Entidad destino:</td><td class='td-l'>"& rsUpdateEntry("mpio_destino") &"</td></tr>")
     response.write("<tr><td class='subtitulo td-r'>Ciudad destino:</td><td class='td-l'>"& rsUpdateEntry("ciudad_destino") &"</td></tr>")
     response.write("<tr><td class='subtitulo td-r'>Razon Social:</td><td class='td-l'>"& rsUpdateEntry("RazonSocial") &"</td></tr>")
     response.write("<tr><td class='subtitulo td-r'>Socio de Negocio:</td><td class='td-l'>"& rsUpdateEntry("CardName") &"</td></tr>")
     response.write("<tr><td class='subtitulo td-r'>Remision:</td><td class='td-l'>"& rsUpdateEntry("Docnum") &"</td></tr>")
     response.write("<tr><td class='subtitulo td-r'>Portador:</td><td class='td-l'>"& rsUpdateEntry("Portador") &"</td></tr>")
     response.write("<tr><td class='subtitulo td-r'>Guia:</td><td class='td-l'>"& rsUpdateEntry("Guia") &"</td></tr>")
     response.write("<tr><td class='subtitulo td-r'>Rastreo:</td><td class='td-l'>"& rsUpdateEntry("rastreo") &"</td></tr>")
     response.write("<tr><td class='subtitulo td-r'>Fecha de embarque:</td><td class='td-l'>"& rsUpdateEntry("DocDate") &"</td></tr>")
     response.write("<tr><td class='subtitulo td-r'>Fecha del email:</td><td class='td-l'>"& rsUpdateEntry("sentdate") &"</td></tr>")
     response.write("<tr><td class='subtitulo td-r'>Re enviar email helpdesk:</td><td class='td-l'><select name='fseleccion'><option value='0'>no</option><option value='1'>si</option></select></td></tr>")
     response.write("<tr><td colspan=2>&nbsp;</td></tr><tr><td colspan=2 class='td-c'><input type='submit' value='continuar'></td></tr></form></table><P>&nbsp;</P>")
     
     
end if



    if request("action")="7" then  're-envio de email  UP' 
      if request("fseleccion")="1" then
         strSQL="update envios set sentdate=null where ID="& request("ID")
         response.write strSQL
         rsUpdateEntry4.Open strSQL, adoCon4

         strSQL="EXEC [dbo].[SP_envios]  "& request("ID")
         response.write strSQL
         rsUpdateEntry5.Open strSQL, adoCon4

         response.redirect("ShowContent.asp?contenido=90&msg=re envio programado")

      else
         response.redirect("ShowContent.asp?contenido=90&msg=re envio cancelado")
      end if
    end if



if request("action")="10" then  'solo informativo'
   

    strSQL="select * from envios where ID=" & request("ID")
    rsUpdateEntry.Open strSQL, adoCon4  

    strSQL="select ID,Nombres+' '+paterno+' '+Materno as nombre from Empleados where FechaEgreso is null order by nombre"
    'response.write strSQL  
    rsUpdateEntry2.Open strSQL, adoCon      

    titulo="Registro de embarque no. "& rsUpdateEntry("ID") 
    DoTitulo(titulo)
     %>

     <table border=0 width="600px" cellpadding="2" cellpadding="2" align="center">
     <form id="envios" method="POST" action="envios.asp"> 
     <tr><td class='fontbig subtitulo td-r' width="40%">Raz&oacute;n Social: </td><td class='fontbig td-l' width="60%"><B><%=rsUpdateEntry("RazonSocial")%></B></td></tr>
     <tr><td class='fontbig subtitulo td-r'># Remisi&oacute;n: </td><td class='fontbig td-l'><B><%=rsUpdateEntry("DocNum")%></B></td></tr>
     <tr><td class='fontbig subtitulo td-r'>Colaborador que captura: </td><td class='fontbig td-l'><select name="fempleado" id="fempleado">
     <%
        while not rsUpdateEntry2.EOF
            if rsUpdateEntry("ID_empleado")=rsUpdateEntry2("ID") then
               response.write("<option value='"& rsUpdateEntry2("ID") &"' SELECTED>"& rsUpdateEntry2("nombre") &"</option>")
            else
               response.write("<option value='"& rsUpdateEntry2("ID") &"'>"& rsUpdateEntry2("nombre") &"</option>")
            end if
            rsUpdateEntry2.MoveNext
        wend
        rsUpdateEntry2.close
     %>
     </select></td></tr>
     
     <tr><td class='fontbig subtitulo td-r'># de gu&iacute;a: </td><td class='fontbig td-l'><input type="text" id="fguia" name="fguia" value="<%=rsUpdateEntry("Guia")%>"></td></tr>
     <tr><td class='fontbig subtitulo td-r'># de rastreo: </td><td class='fontbig td-l'><input type="text" id="frastreo" name="frastreo" value="<%=rsUpdateEntry("rastreo")%>"></td></tr>

     <tr><td class='fontbig subtitulo td-r'>Subtotal sin impuestos: </td><td class='fontbig td-l'><input type="text" id="ftotal" name="ftotal" value="<%=rsUpdateEntry("subTotal")%>"></td></tr>
   
     
     <tr><td class='fontbig subtitulo td-r'>Fecha del env&iacute;o: </td><td class='fontbig td-l'> D&iacute;a:  <select name=vDia> <%

              for i=1 to 31
                    if day(rsUpdateEntry("DocDate"))=i then
                       response.write("<option value="& i &" SELECTED>"& i &"</option>")
                   else
                       response.write("<option value="& i &" >"& i &"</option>")
                   end if
               next
               response.write("</select> Mes: <select name=vMes>")
               for i=1 to 12
                    if month(rsUpdateEntry("DocDate"))=i then
                       response.write("<option value="& i &" SELECTED>"& meses(i) &"</option>")
                   else
                       response.write("<option value="& i &" >"&  meses(i) &"</option>")
                   end if
               next
               response.write("</select> A&ntilde;o: <select name=vAnio>")
               for i=year(date())-1 to year(date())
                   if year(rsUpdateEntry("DocDate"))=i then
                       response.write("<option value="& i &" SELECTED>"& i &"</option>")
                   else
                       response.write("<option value="& i &" >"& i &"</option>")
                   end if
               next
               response.write("</select> </td></tr>  ")       %>

    
     <tr><td class='fontbig subtitulo td-r'>Portador: </td><td class='fontbig td-l'><select name="fcarrier" id="fcarrier" onchange="ValidaRemision()">
          <%
          strSQL="select * from cat_portadores order by ID_portador"
          rsUpdateEntry3.Open strSQL, adoCon4 

          while not rsUpdateEntry3.EOF
              if rsUpdateEntry("ID_portador")=rsUpdateEntry3("ID_portador") then
                   response.write("<option value="& rsUpdateEntry3("ID_portador") &" SELECTED>"& rsUpdateEntry3("Portador") &"</option>")
              else
                   response.write("<option value="& rsUpdateEntry3("ID_portador") &">"& rsUpdateEntry3("Portador") &"</option>")
              end if
              rsUpdateEntry3.MoveNext
          wend
          rsUpdateEntry3.close


%>
        </select></td></tr>

     <tr><td class='fontbig subtitulo td-r'>Cargo a Cliente: </td><td class='fontbig td-l'>
            <%
               if  rsUpdateEntry("CargoCliente")=1 then
                   response.write("<input type='checkbox'name='fCargoCliente' id='fCargoCliente' CHECKED>")
               else
                  response.write("<input type='checkbox'name='fCargoCliente' id='fCargoCliente'>")
               end if
            %>
       </td></tr>
     
     <tr><td class='fontbig subtitulo td-r'>Comentarios: </td><td class='fontbig td-l'><input type="text" id="fnotas" name="fnotas" class='anchox2' value="<%=rsUpdateEntry("Comentarios")%>"> </td></tr>
     <tr><td class='fontbig subtitulo td-r'>Re-env&iacute;o: </td><td class='fontbig td-l'>
     <%
               if  rsUpdateEntry("reenvio")=1 then
                   response.write("<input type='checkbox'name='freenvio' id='freenvio' CHECKED>")
               else
                  response.write("<input type='checkbox'name='freenvio' id='freenvio'>")
               end if
      %>
      (para un segundo env&iacute;o) </td></tr>
      <tr><td colspan=2 class='td-c'> &nbsp; </td></tr>
     <tr><td colspan=2 class='td-c'>  </td></tr>
    
     </table>

<%
end if

end sub




sub EnvioConsolidado   'contenido 91'

   if request("action")="" then   'ShowRemisiones'

       doAlert

       response.write("<div id='destino'>")
       titulo="Agregando un registro de embarque consolidado"
       Dotitulo(titulo)

  %>


     <form id="envios" method="POST" action="ShowContent.asp"> 
     <input type="hidden" name="contenido" value="91">
     <input type="hidden" name="action" value="2">


     <table border=0 width="470px" cellpadding="2" cellpadding="2" align="center"><%

     if request("fentidad")<>"" and request("fCiudadDestino")<>"" then
           vString=request("fCiudadDestino")
           vPos=inStr(vString,"*")
           vString1=mid(vString,1,Vpos-1)
           vString2=mid(vString,Vpos+1,30)

           response.write("<input type='hidden' name='fentidad' value='"& request("fentidad") &"'>")
           response.write("<input type='hidden' name='fmunicipio' value='"& vString1 &"'>")
           response.write("<input type='hidden' name='fciudad' value='"& vString2 &"'>")

           response.write("<tr><td class='subtitulo td-r'>Ciudad destino:</td><td class='td-l'>"& vString2 &"</td></tr>")

      else
      %> 
     <tr><td class='fontbig subtitulo td-r'>Entidad Federativa Destino:</td><td class='td-l fontmed'>  

           <select id="entidad_destino" onchange="Javascript:PassValueChangeDiv('entidad_destino','destino','/modules/FormChangeStateEnviosC.asp')">   <%  'elemento,divn,modulo  

           strSQL="select * from cat_entidades order by id_entidad"
           response.write strSQL  
           rsUpdateEntry2.Open strSQL, adoCon 

          while not rsUpdateEntry2.EOF
              response.write("<option value='"& rsUpdateEntry2("id_entidad") &"'>"& rsUpdateEntry2("entidad")  &"</option>")             
              rsUpdateEntry2.MoveNext
          wend
          rsUpdateEntry2.Close
          response.write("</select></td></tr>")    

      end if
      %>
 
     <tr><td class="fontbig subtitulo td-r" width="65%">Raz&oacute;n Social: </td><td class="fontbig td-l" width="35%"><select id="fRS" name="fRS"> <option value="DMX">DMX</option> <option value="DFW" SELECTED>DFW</option> </select></td></tr>
     <tr><td class="fontbig subtitulo td-r">Cantidad total de remisiones a enviar: </td><td class="fontbig td-l"> <select name="fcantidad"><%
     for i=2 to 50 
           response.write("<option value="& i &">"& i &"</option>")
     next   %>
    </selec> </td></tr>
    <tr><td colspan=2>&nbsp;</td></tr>
    <tr><td colspan=2>&nbsp;</td></tr>
    <tr><td colspan=2 class="td-c"><input type="submit" value="continuar"></td></tr>
     </form>
     </table></div><P>&nbsp;</P>
 
<%

   end if



   if request("action")="2" then   'CheckRemisiones'

         titulo="Agregando un registro de embarque consolidado"
         Dotitulo(titulo)

        response.write("<form id='consolidado' method='POST' action='ShowContent.asp'>")
         response.write("<input type='hidden' name='contenido' value='91'>")
         response.write("<input type='hidden' name='action' value='3'>")
         
         response.write("<input type='hidden' name='fentidad' value='"& request("fentidad") &"'>")
         response.write("<input type='hidden' name='fmunicipio' value='"& request("fmunicipio")  &"'>")
         response.write("<input type='hidden' name='fciudad' value='"& request("fciudad")  &"'>")

         response.write("<table border=0 width='400px' cellpadding=2 cellpadding=2 align='center'>")
         response.write("<tr><td class='fontbig subtitulo td-r'>Ciudad destino:</td><td class='td-l'>"& request("fCiudad") &"</td></tr>")

         response.write("<input type='hidden' name='fRS' id='fRS' value="& request("fRS") &">")
         response.write("<input type='hidden' name='fcantidad' id='fcantidad' value="& request("fcantidad") &">")
         
         
         response.write("<tr><td class='fontbig subtitulo td-r' width='65%'>Raz&oacute;n Social: </td><td class='fontbig td-l' width='35%'>"& request("fRS") &"</td></tr>")

         for i=1 to request("fcantidad")
              response.write("<tr><td class='fontbig subtitulo td-r'>Remision "& i &": </td><td class='fontbig td-l'><input type='number' name='remision"&i&"'></td></tr>")
         next
         %>
          <tr><td colspan=2>&nbsp;</td></tr>
          <tr><td colspan=2>&nbsp;</td></tr>
          <tr><td colspan=2 class="td-c"><input type="submit" value="continuar"></td></tr>
           </form>
           </table>  <%
    
   end if




   if request("action")="3" then    'Detalle en Remisiones'

         vstring=""
         remisiones=""

     for i=1 to request("fcantidad") 
         vstring="remision"&i
         remisiones=remisiones & request(vstring)
         remisiones=remisiones&","
     next
      vpos=len(remisiones)
      remisiones=mid(remisiones,1,vpos-1)
      'response.write ("valor=" & remisiones &"<BR>")

      strSQL="select a.docnum,b.LineNum+1,convert(varchar,a.docnum)+'-'+convert(varchar,b.LineNum+1) as 'partida',b.ItemCode,c.ItemName,b.Quantity,b.Price,b.WhsCode "
      strSQL= strSQL & "from ODLN a inner join DLN1 b on a.DocEntry=b.DocEntry inner join OITM c on b.ItemCode=c.ItemCode "
      strSQL= strSQL & "where a.docnum in ("& remisiones &") order by a.DocNum,b.LineNum"

      strSQL2="select count(distinct(docnum)) as calculo from ODLN where docnum in ("& remisiones &")"

      'response.write strSQL

      if request("fRS")="DFW" then
           rsUpdateEntry.Open strSQL, adoCon3 
           rsUpdateEntry2.Open strSQL2, adoCon3 
      else
           rsUpdateEntry.Open strSQL, adoCon2 
           rsUpdateEntry2.Open strSQL2, adoCon2 
      end if


      if rsUpdateEntry.EOF then
           response.redirect("/modules/ShowContent.asp?contenido=91&msg=no existen remisiones con esa numeracion")
      end if


      if rsUpdateEntry2("calculo")=1 then
           response.redirect("/modules/ShowContent.asp?contenido=91&msg=para una sola remision utice registro de envio sencillo")
      end if

      rsUpdateEntry2.close
      titulo="Verifique cantidades del embarque consolidado ("&request("fRS") &")"
      DoTitulo(titulo)

      response.write("<table width='800px' cellpadding=3 cellpadding=2 align=center border=1>")

      response.write("<form id='envios' method='POST' action='ShowContent.asp'>")
      response.write("<input type='hidden' name='contenido' value='91'>")
      response.write("<input type='hidden' name='action' value='4'>")

      response.write("<input type='hidden' name='fentidad' value='"& request("fentidad") &"'>")
      response.write("<input type='hidden' name='fmunicipio' value='"& request("fmunicipio")  &"'>")
      response.write("<input type='hidden' name='fciudad' value='"& request("fciudad")  &"'>")
         
      response.write("<input type='hidden' name='fRS'  id='fRS' value="& request("fRS") &">")     
      response.write("<input type='hidden' name='fremisiones' id='fremisiones' value="& remisiones &">")
      %>


    <tr>
      <td class="td-c azul fontbig">Partida</td>
      <td class="td-c azul fontbig">C&oacute;digo</td>
      <td class="td-c azul fontbig">Nombre del c&oacute;digo</td>
      <td class="td-c azul fontbig">Cantidad</td>
      <td class="td-c azul fontbig">Precio</td>
      <td class="td-c azul fontbig">Almac&eacute;n</td>
    </tr>
<%
    i=1
    while not rsUpdateEntry.EOF
      response.write("<tr>")
      response.write("<td class='td-r fontmed'>"& rsUpdateEntry("partida") &"</td>")
      response.write("<td class='td-l fontmed'>"& rsUpdateEntry("ItemCode") &"</td>")
      response.write("<td class='td-l fontmed'>"& mid(rsUpdateEntry("ItemName"),1,24) &"</td>")
      response.write("<input type='hidden' name='partida"&i&"' value="& rsUpdateEntry("docnum") &">")
      response.write("<td class='td-r fontmed'><input type='number' name='cantidad"&i&"' value='"& rsUpdateEntry("Quantity")&"' style='width:60px;text-align:right;'></td>")
      response.write("<input type='hidden' name='precio"&i&"' value="& rsUpdateEntry("price") &">")
      response.write("<td class='td-r fontmed'>"& rsUpdateEntry("Price") &"</td>")
      response.write("<td class='td-r fontmed'>"& rsUpdateEntry("WhsCode") &"</td>")
      response.write("</tr>")
      rsUpdateEntry.MoveNext
      i=i+1
     wend
      rsUpdateEntry.close
      response.write("<input type='hidden' name='fcantidad' id='fcantidad' value="& i-1 &">")

    

     response.write("<tr><td class='fontbig subtitulo td-r' colspan=3>Ciudad destino:</td><td class='fontbig td-l' colspan=3>"&request("fciudad") &"</td></tr>")
      %>

     <tr><td class='fontbig subtitulo td-r'  colspan=3># de gu&iacute;a: </td><td class='fontbig td-l'  colspan="3"><input type="text" id="fguia" name="fguia"></td></tr>
     <tr><td class='fontbig subtitulo td-r'  colspan=3># de rastreo: </td><td class='fontbig td-l'  colspan="3"><input type="text" id="frastreo" name="frastreo"></td></tr>

     <tr><td class='fontbig subtitulo td-r'  colspan=3>Subtotal sin impuestos por todo el servicio: </td><td class='fontbig td-l'  colspan="3"><input type="text" id="ftotal" name="ftotal" placeholder="no puede ir vac&iacute;o"></td></tr>
   
     <tr><td class='fontbig subtitulo td-r'  colspan=3>Fecha del env&iacute;o: </td><td class='fontbig td-l'  colspan="3"><input name="ffecha" type="date"/> </td></tr>
     <tr><td class='fontbig subtitulo td-r'  colspan=3>Portador: </td><td class='fontbig td-l'  colspan="3"><select name="fcarrier" id="fcarrier" onchange="ValidaRemision()">
         <option value="0">seleccione</option> <%

          strSQL="select * from cat_portadores order by ID_portador"
          rsUpdateEntry3.Open strSQL, adoCon4 

          while not rsUpdateEntry3.EOF
              response.write("<option value="& rsUpdateEntry3("ID_portador") &">"& rsUpdateEntry3("Portador") &"</option>")
              rsUpdateEntry3.MoveNext
          wend
          rsUpdateEntry3.close  %>

          </select></td></tr>

      <tr><td class='fontbig subtitulo td-r' colspan=3>En OCURRE a otro Almac&eacute;n:</td><td class='fontbig td-l' colspan="3"><select name='focurre' class="anchox2">
              <option value="" SELECTED>si aplica seleccione</option>
              <option value="001-Mxl">001-Mxl</option>
              <option value="002-SJR">002-SJR</option>
              <option value="003-Mty">003-Mty</option>
             </select>  </td></tr>

     <tr><td class='fontbig subtitulo td-r'  colspan=3>Cargo a Cliente: </td><td class='fontbig td-l'  colspan=3><input type="checkbox" name="fCargoCliente" id="fCargoCliente"> </td></tr>
     <tr><td class='fontbig subtitulo td-r' colspan=3 >Comentarios: </td><td class='fontbig td-l'  colspan=3><input type="text" id="fnotas" name="fnotas" class='anchox2'></td></tr>
     <%
       strSQL="select * from Empleados where ID_USUARIO='"& session("session_id") &"'"
     'response.write strSQL  
      rsUpdateEntry.Open strSQL, adoCon

      response.write("<tr><td class='fontbig subtitulo td-r' colspan=3>Colaborador que captura: </td><td class='fontbig td-l' colspan=3><select name=fempleado>")    
        while not rsUpdateEntry.EOF
            response.write("<option value='"& rsUpdateEntry("ID") &"'>"& rsUpdateEntry("ID_USUARIO") &"</option>")
            rsUpdateEntry.MoveNext
        wend
        rsUpdateEntry.close
        response.write("</select></td></tr>")   
     %>
     <tr><td class='fontbig  td-c' colspan=6 >&nbsp;</td></tr>
     <tr><td class='fontbig  td-c' colspan=6 >&nbsp;</td></tr>
     <tr><td class='fontbig  td-c' colspan=6 ><input type="submit" value="registrar"></td></tr>
    </form></table>

<%
   
end if



if request("action")="4" then    'CONSOLIDADO UP'

if request("fCarrier")="0" then response.redirect("ShowContent.asp?contenido=91&msg=no selecciono un portador para su envio")

dim remisiones(99)
  vstring=trim(request("fremisiones"))

  for i=1 to 99
     vpos=inStr(vstring,",")
     if vpos=0 then exit for
         remisiones(i)=mid(vstring,1,vpos-1)
         vstring=mid(vstring,vpos+1,3000)
     next
  remisiones(i)=vstring
  
  strSQL="if exists (select * from INFORMATION_SCHEMA.TABLES where TABLE_NAME = 'enviosC' AND TABLE_SCHEMA = 'dbo')     drop table dbo.enviosC;"
  rsUpdateEntry7.Open strSQL, adoCon4

  strSQL="if not exists (select * from INFORMATION_SCHEMA.TABLES where TABLE_NAME = 'enviosC' AND TABLE_SCHEMA = 'dbo')   CREATE TABLE [dbo].[enviosC](    [DocNum] [int] NULL,    [Cantidad] [int] NULL,  [Precio] [decimal](18, 2) NULL,     [subTotal] [decimal](18, 2) NULL,   [Prorrateo] [decimal](18, 4) NULL ) ON [PRIMARY] "
  rsUpdateEntry7.Open strSQL, adoCon4 

  for n=1 to request("fcantidad")
      vstring1="partida"&n
      vstring2="cantidad"&n
      vstring3="precio"&n

      strSQL="insert into enviosC (DocNum,Cantidad,Precio) values ("& request(vstring1)  &"," & request(vstring2)  &"," & request(vstring3)  &")"
      rsUpdateEntry.Open strSQL, adoCon4 
  next

  strSQL="UPDATE enviosC set Subtotal=round(cantidad*Precio,2)"  
  rsUpdateEntry6.Open strSQL, adoCon4 

  strSQL="select sum(Subtotal) as calculo from enviosC"    
  rsUpdateEntry.Open strSQL, adoCon4 

  vfecha=date()

  for n=1 to i
       strSQL="select  top 1 a.cardcode,a.cardname,b.WhsCode from ODLN a inner join DLN1 b on a.DocEntry=b.DocEntry where a.docnum="& remisiones(n) 

       strSQL2="select sum(Subtotal) as calculo from enviosC where DocNum=" & remisiones(n)    
       'response.write(strSQL&"<br>")
       'response.write(strSQL2&"<br>")

       rsUpdateEntry3.Open strSQL2, adoCon4 

       if request("fRS")="DFW" then
              rsUpdateEntry2.Open strSQL, adoCon3 
       else
              rsUpdateEntry2.Open strSQL, adoCon2
       end if

       vfecha=request("ffecha")

       strSQL="set dateformat ymd;insert into envios (entidad_destino,mpio_Destino,ciudad_destino,ID,DocDate,ID_USUARIO,ID_empleado,LogDate,RazonSocial,DocNum,Tipo,WhsCode,id_Portador,Guia,Rastreo,subTotal,Ocurre,Cardcode,Cardname,Comentarios,SubtotalConsolidado,CargoCliente,consolidado) values ('"& request("fentidad") &"','"& request("fmunicipio") &"','"& request("fciudad") &"',(select max(ID) from envios)+1,'" & year(vfecha) &"-" & month(vfecha) &"-" & day(vfecha) &"','"& session("session_id") &"',"
          strSQL= strSQL & request("fempleado") &",getdate(),'" & request("fRS") &"'," & remisiones(n) &",'remision','" & rsUpdateEntry2("WhsCode") &"',"& request("fCarrier") &",'"
          strSQL= strSQL & request("fGuia") &"','" &  request("frastreo") &"',"           
          strSQL= strSQL & "round( CAST(" & request("fTotal") &" AS FLOAT)*CAST(" & rsUpdateEntry3("calculo")    &" AS FLOAT)/ CAST(" &  rsUpdateEntry("calculo")  &" AS FLOAT) , 2) "
          strSQL= strSQL  &",'" & request("focurre") &"','"  & rsUpdateEntry2("cardcode") &"','"& rsUpdateEntry2("cardname") &"','"& request("fNotas") &"'," & request("fTotal") &","
        
           if request("fCargoCliente")="on" then
             strSQL= strSQL & "1,1)"
          else
             strSQL= strSQL & "0,1)"
          end if
   
       response.write(strSQL&"<br>")

       rsUpdateEntry5.Open strSQL, adoCon4      

       rsUpdateEntry2.close
       rsUpdateEntry3.close
  next
  Sleep(1)
  response.redirect("ShowContent.asp?contenido=90&msg=se ha registrado un embarque consolidado")
 

end if

end sub


function Sleep(seconds)
            set oShell = CreateObject("Wscript.Shell")
            cmd = "%COMSPEC% /c timeout " & seconds & " /nobreak"
            oShell.Run cmd,0,1
End function


sub StockSinPeso
     DoAlert
     titulo="Codigos en DMX con stock y sin peso"
     DoTitulo(titulo)

     strSQL="SELECT y.ItemCode as Codigo,T0.ItemName as Descripcion,T3.ItmsGrpNam as Familia,T0.BWeight1 as Peso,cast(SUM(Y.InQty)-SUM(Y.OutQty) as int) as Stock  FROM PRODUCTIVA_DMX.dbo.OINM Y WITH (NOLOCK)     LEFT JOIN PRODUCTIVA_DMX.dbo.OITM T0 on Y.ItemCode=T0.ItemCode  LEFT JOIN PRODUCTIVA_DMX.dbo.OITB T3 WITH (NOLOCK) ON T0.ItmsGrpCod=T3.ItmsGrpCod where T0.BWeight1=0 group by y.ItemCode,T0.ItemName,T3.ItmsGrpNam,T0.BWeight1     having ( SUM(Y.InQty)-SUM(Y.OutQty) )>0"
     'response.write strSQL
     rsUpdateEntry.Open strSQL, adoCon2
     rsUpdateEntry2.Open strSQL, adoCon2

     CreateTable(800)
     CamposRS1

     while not rsUpdateEntry2.EOF
          RowIn
          response.write("<td class=''>"& rsUpdateEntry2("Codigo") &"</td>")
          response.write("<td class=''>"& rsUpdateEntry2("Descripcion") &"</td>")
          response.write("<td class=''>"& rsUpdateEntry2("Familia") &"</td>")
          response.write("<td class=''>"& rsUpdateEntry2("Peso") &"</td>")
          response.write("<td class=''>"& rsUpdateEntry2("Stock") &"</td>")
          RowOut
          rsUpdateEntry2.MoveNext
     wend
     rsUpdateEntry2.close
     closeTable

end sub






sub DetalleXMLEfectivale 'contenido 94

     strSQL="select * from XML where NombreXML='"& request("xml") &"'"
     'response.write strSQL
     rsUpdateEntry.Open strSQL, adoCon4     

     DoAlert
     titulo="Detalle xml de Estado de Cuenta Combustible ("& request("xml")  &")<br> " & rsUpdateEntry("emisor")
     rsUpdateEntry.close
     DoTitulo(titulo)

     strSQL="select ID,Identificador as Tarjeta,Fecha,RFC,ClaveEstacion as Estacion,cantidad,Unidad,NombreCombustible as Tipo,FolioOperacion as Folio,ValorUnitario as Precio,Importe as Importe_MXN,ID_unidad as ID from XML_detalle where NombreXML='"& request("xml") &"' order by ID_unidad,Fecha"
     'response.write strSQL
     rsUpdateEntry.Open strSQL, adoCon4
     rsUpdateEntry2.Open strSQL, adoCon4
     

     response.write("<table border=1 cellpadding=4 cellspacing=2 width='1160px' class='table2excel table2excel_with_colors' data-tableName='table1'>")

     Dim Campos(12)
     i=1
     For Each rsUpdateEntry in rsUpdateEntry.Fields
            Response.Write "<td class='subtitulo td-c'>" & rsUpdateEntry.Name & "</td>"
            campos(i)=rsUpdateEntry.Name
            i=i+1
     Next 
     Response.Write("<td class='subtitulo td-c anchox2'>Comprobacion</td>")   
     Response.Write("<td class='subtitulo td-c anchox2'>Chofer</td>") 
     Response.Write("<td class='subtitulo td-c'>Destino</td>")
     Response.Write("<td class='subtitulo td-c' width='75px'>Kms 1</td>")  
     Response.Write("<td class='subtitulo td-c' width='75px'>Kms 2</td>") 
     Response.Write("<td class='subtitulo td-c' width='75px'>kms</td></tr>") 
     vID_unidad=""     
     strSQL="select ("


     while not rsUpdateEntry2.EOF
           vID_unidad=rsUpdateEntry2("ID")
           rowin
             for i=1 to 12
               Response.write("<td class='td-r fonttiny'>"& rsUpdateEntry2(campos(i)) &"</td>")
             next  

             strSQL2="select a.*,b.Chofer,b.ciudad_origen,b.ciudad_destino,b.kms_viaje,c.Kilometraje as inicial,d.Kilometraje as Final from VoucherGas a left join Viajes b on a.ID_viatico=b.ID_viatico left join check_unidad c on c.Mode='out' and b.ID_viaje=c.ID_viaje left join check_unidad d on d.Mode='in' and b.ID_viaje=d.ID_viaje where a.Estacion="& rsUpdateEntry2("Estacion") &"  and a.DocDate='"& rsUpdateEntry2("Fecha") &"' and a.id_unidad='"&  rsUpdateEntry2("ID") &"'"
             'response.write(strSQL2&"<BR>")
             rsUpdateEntry4.Open strSQL2, adoCon

             if rsUpdateEntry4.EOF then
                   'Response.Write("<td class='td-c fonttiny alert' colspan=6>no se encontr&oacute;, revise</td>")
                   Response.Write("<td class='td-c fonttiny' colspan=6>"& strSQL2 &"</td>")
                   'Response.Write("<td class='td-c fonttiny'>-</td>")
                   'Response.Write("<td class='td-c fonttiny'>-</td>")
                   'Response.Write("<td class='td-c fonttiny'>-</td>")
                   'Response.Write("<td class='td-c fonttiny'>-</td>")
                   'Response.Write("<td class='td-c fonttiny'>-</td></tr>")
             else
                   Response.Write("<td class='td-c fonttiny'>"& rsUpdateEntry4("Motivo") &"</td>")   
                   if trim(rsUpdateEntry4("Chofer")) ="0" then            
                        Response.Write("<td class='td-c fonttiny'>N/A</td>") 
                   else
                        Response.Write("<td class='td-c fonttiny'>"& rsUpdateEntry4("Chofer") &"</td>") 
                   end if   
                   Response.Write("<td class='td-c fonttiny'>"& rsUpdateEntry4("ciudad_destino") &"</td>")     
                   Response.Write("<td class='td-c fonttiny'>"& rsUpdateEntry4("inicial") &"</td>")     
                   Response.Write("<td class='td-c fonttiny'>"& rsUpdateEntry4("Final") &"</td>")    
                   Response.Write("<td class='td-c fonttiny'>"& rsUpdateEntry4("kms_viaje") &"</td></tr>") 

             end if
             rsUpdateEntry4.close
                    
           rowout

           strSQL=strSQL&trim(rsUpdateEntry2("Importe_MXN"))&"+"
           rsUpdateEntry2.MoveNext
           if not rsUpdateEntry2.EOF then
               if rsUpdateEntry2("ID")<>vID_unidad then
                   separador
                    strSQL=mid(strSQL,1,len(strSQL)-1)
                    strSQL=strSQL&") as calculo"
                    rsUpdateEntry5.Open strSQL, adoCon4
                    Response.write("<td colspan=10 class='td-r fontmed strong'>Importe</td><td class='td-r strong'>$ "& rsUpdateEntry5("calculo") &"</td><td>&nbsp;</td>")
                    rsUpdateEntry5.close
                    strSQL="select ("
                   separador
               end if
           else
              separador
              strSQL=mid(strSQL,1,len(strSQL)-1)
              strSQL=strSQL&") as calculo"
              rsUpdateEntry5.Open strSQL, adoCon4
              Response.write("<td colspan=12 class='td-r fontmed strong'>Importe</td><td class='td-r strong'>$ "& rsUpdateEntry5("calculo") &"</td><td>&nbsp;</td>")
              rsUpdateEntry5.close
              separador
           end if
         wend
         rsUpdateEntry2.close

       %>
     
     </table><button class="exportToExcel">Exportar a excel</button> </center>  

           <script>
      $(function() {
        $(".exportToExcel").click(function(e){
          var table = $(this).prev('.table2excel');
          if(table && table.length){
            var preserveColors = (table.hasClass('table2excel_with_colors') ? true : false);
            $(table).table2excel({
              exclude: ".noExl",
              name: "Excel Document Name",
              filename: "myFileName" + new Date().toISOString().replace(/[\-\:\.]/g, "") + ".xls",
              fileext: ".xls",
              exclude_img: true,
              exclude_links: true,
              exclude_inputs: true,
              preserveColors: preserveColors
            });
          }
        });
        
      });
    </script>
    
    <P>&nbsp</P><P>&nbsp</P>
      <%


end sub


sub CargarXMLEfectivale  'contenido 93'

     DoAlert
     titulo="Ingresando informaci&oacute;n de combustible"
     DoTitulo(titulo)

     if request("action")="" then

         strSQL="select * from XML_VALIDA"
         'response.write strSQL
         rsUpdateEntry.Open strSQL, adoCon4
         rsUpdateEntry2.Open strSQL, adoCon4

         CreateTable(1400)
         CamposRS1
         ShowValoresRS2
         closeTable

         strSQL="select a.*,b.id_unidad from XML_VALIDA_detalle a left join INTRANET.dbo.flotilla b on a.Identificador collate Modern_Spanish_CI_AS=b.TarjetaCombustible collate Modern_Spanish_CI_AS"
         rsUpdateEntry3.Open strSQL, adoCon4
         rsUpdateEntry4.Open strSQL, adoCon4

         CreateTable(900)
         CamposRS3
         ShowValoresRS4
         closeTable

         response.write("<form id='combustibles' method='post' action='ShowContent.asp'>")
         response.write("<input type='hidden' name='contenido' value='93'>")
         response.write("<input type='hidden' name='action' value='2'>")  

         CreateTable(340)
         RowIn
         response.write("<td class='subtitulo td-c' width='80%'>Esta correcta la informacion<br> para ingresar en la base de datos?</td><td class='td-l'><select name='flag'><option value='0'>no</option><option value='1'>si</option></select></td>")
         RowOut
         RowIn
         response.write("<td colspan=2>&nbsp;</td>")
         RowOut
         RowIn
         response.write("<td colspan=2 class='td-c'><input type='submit' value='continuar'></form></td>")
         closeTable


      end if

     


     if request("action")="2" then

         select case request("flag")
            case "0"
                 response.redirect("ShowContent.asp?contenido=116&msg=proceso cancelado")
            case "1"
                 strSQL="insert into xml (Fecha,TipoComprobante,MetodoPago,FormaPago,Moneda,Folio,Subtotal,Impuestos,Total,RFCEmisor,Emisor,RFCReceptor,Receptor,UsoCFDi,LogDate,UUID,NombreXML,Concepto,cargado,CardCode) select Fecha,TipoComprobante,MetodoPago,FormaPago,Moneda,Folio,Subtotal,Total-Subtotal,Total,RFCEmisor,Emisor,RFCReceptor,Receptor,UsoCFDi,LogDate,UUID,NombreXML,Concepto,1,'A000015' from  XML_VALIDA"
                 response.write(strSQL &"<br>")
                 rsUpdateEntry.Open strSQL, adoCon4

                 strSQL="insert into XML_detalle (NombreXML,ID,Identificador,Fecha,RFC,ClaveEstacion,cantidad,TipoCombustible,Unidad,NombreCombustible,FolioOperacion,ValorUnitario,Importe,ID_unidad) select a.NombreXML,a.ID,a.Identificador,a.Fecha,a.RFC,a.ClaveEstacion,a.cantidad,a.TipoCombustible,a.Unidad,a.NombreCombustible,a.FolioOperacion,a.ValorUnitario,a.Importe,ID_unidad from XML_VALIDA_detalle a left join INTRANET.dbo.Flotilla b on a.Identificador collate Modern_Spanish_CI_AS=b.TarjetaCombustible collate Modern_Spanish_CI_AS"
                 response.write(strSQL &"<br>")
                 rsUpdateEntry2.Open strSQL, adoCon4
                 response.redirect("ShowContent.asp?contenido=116&msg=la informacion ha sido registrada en la base de datos")

         end select


     end if

end sub



sub VoucherGasOtros    

      vString=trim(request("fCiudadDestino"))
      vPos=inStr(vString,"*")
      vString1=mid(vString,1,Vpos-1)
      vString2=mid(vString,Vpos+1,30)

      doAlert
      titulo="Registro de vouchers de combustible"
      response.write("<P>&nbsp;</P>")
      DoTitulo(titulo)

      response.write("<form id='vouchers' action='ShowContent.asp' method='POST'>")
      response.write("<input type='hidden' name='contenido' value='95'>")
      response.write("<input type='hidden' name='action' value='2'>")     
      response.write("<input type='hidden' name='control' value='O'>")
      response.write("<input type='hidden' name='fentidad' value='"& request("fentidad")&"'>")     
      response.write("<input type='hidden' name='fciudad' value='"& vString2 &"'>")

    
      CreateTable(560)
      RowIn      
      response.write("<td width='40%' class='td-r subtitulo'>Naturaleza de la carga:</td><td width='60%' class='td-l'>otra causa </td></tr>")
      response.write("<td class='td-r subtitulo'>Ciudad Destino:</td><td width='40%' class='td-l'>"& vString2 &"</td></tr>")  
      response.write("<tr><td class='subtitulo td-r FontMed'>Chofer:</td><td class='td-l FontMed'><select name='fchofer'>")

      strSQL="select nombres+' '+paterno as colaborador,* from Empleados where FechaEgreso is null"
      'response.write strSQL  
      rsUpdateEntry.Open strSQL, adoCon

      while not rsUpdateEntry.EOF
           if UCASE(rsUpdateEntry("id_usuario"))=session("session_id")  then
                response.write("<option value='"& rsUpdateEntry("colaborador") &"' SELECTED>"&  rsUpdateEntry("colaborador") &"</option>")
           else
                response.write("<option value='"& rsUpdateEntry("colaborador") &"'>"&  rsUpdateEntry("colaborador") &"</option>")
           end if
           rsUpdateEntry.MoveNext
      wend
      rsUpdateEntry.close
      response.write("</select></td></tr>")

      'response.write("<tr><td class='subtitulo td-r FontMed'>Motivo:</td><td class='td-l FontMed'><input type='text' name='fmotivo' class='anchox2' placeholder='no debe ir vacio'></td></tr>")

      response.write("<tr><td class='subtitulo td-r FontMed'>Tarjeta combustible:</td><td class='td-l FontMed'><select name='ftarjeta'>")

            strSQL="select * from flotilla where ID_unidad!='000' and tipo_activo<4"
            rsUpdateEntry2.Open strSQL, adoCon 

               while not rsUpdateEntry2.EOF
                   response.write("<option value='"& rsUpdateEntry2("tarjetaCombustible")&"'>"& rsUpdateEntry2("tarjetaCombustible") &" (" & rsUpdateEntry2("id_unidad")&") </option>")
                   rsUpdateEntry2.MoveNext
               wend
            response.write("</select></td></tr>")
            rsUpdateEntry2.close

     response.write("<tr><td class='subtitulo td-r FontMed'>Fecha de la carga:</td><td class='td-l FontMed'><input type='date' name='ffecha'></td></tr>")        
           response.write("<tr><td class='subtitulo td-r FontMed'>Estacion:</td><td class='td-l FontMed'><input type='text' name='festacion'></td></tr>")  
           response.write("<tr><td class='subtitulo td-r FontMed'>Cantidad en litros:</td><td class='td-l FontMed'><input type='text' name='flitros' placeholder='no utilice comas'></td></tr>")  

         
           response.write("<tr><td colspan=2 class='td-c FontMed'><input type='submit' value='continuar'></form></td></tr>")
           response.write("<tr><td colspan=2 class='td-c FontMed'>&nbsp;</td></tr>")
           response.write("<tr><td colspan=2 class='td-c FontMed'><a href='ShowContent.asp?contenido=95'>regresar</a></td></tr>")
      
           CloseTable


end sub




sub voucherGas  'contenido 95'

    doAlert
    titulo="Registro de vouchers de combustible"
    response.write("<P>&nbsp;</P>")
    DoTitulo(titulo)

  if request("action")="" then
    response.write("<div id='causa'>")  
    CreateTable(300)
    RowIn
    response.write("<td width='60%' class='td-r subtitulo'>Motivo de la carga:</td><td width='40%' class='td-l'>") %>
    <select id="causal" onchange="Javascript:PassValueChangeDiv('causal','causa','/modules/FormCausaVouchers.asp')">   <% 'elemento,divn,modulo  
      response.write("<option value='0'>seleccione</option>")
      response.write("<option value='V'>por viatico</option>")
      response.write("<option value='O'>otra causa</option></select></td>")
    rowout        

    response.write("<tr><td colspan=2>&nbsp</td></tr>")
    closeTable
    response.write("</div>")

    ShowVouchers
  end if 


 if request("action")="2" then   'revisar integridad de los datos'

    Vviatico=0
    vMotivo="Combustible por viatico no. " & request("fviatico")

    if trim(request("ffecha"))="" then
         response.redirect("ShowContent.asp?contenido=95&msg=no puede enviar la fecha en blanco")
    end if

    if trim(request("festacion"))="" then
         response.redirect("ShowContent.asp?contenido=95&msg=no puede enviar la estacion de carga en blanco")
    end if

    if trim(request("flitros"))="" then
         response.redirect("ShowContent.asp?contenido=95&msg=no puede enviar la cantidad de litros en blanco")
    end if

    SELECT CASE request("control")
    case "V"
         strSQL="select * from VIATICOS_R1 where ID="&request("fviatico")
         rsUpdateEntry.Open strSQL, adoCon

         vUnidad=rsUpdateEntry("ID_unidad")
         vChofer=rsUpdateEntry("beneficiario")

         rsUpdateEntry.close

         if vUnidad="000" then
             strSQL="select * from Flotilla where TarjetaCombustible='"& request("ftarjeta")  &"'"
             rsUpdateEntry.Open strSQL, adoCon

             vUnidad=rsUpdateEntry("ID_unidad")
             rsUpdateEntry.close
         end if

         strSQL="set dateformat ymd;select * from VoucherGas where DocDate='" &request("ffecha") &"' and Estacion='" &request("festacion") &"' and TarjetaCombustible='" &request("ftarjeta") &"'"
         response.write(strSQL&"<br>")
         Vviatico=request("fviatico")    
         vTarjeta=request("ftarjeta")      

    case "O"
                  
         strSQL="select * from Flotilla where TarjetaCombustible='"& request("ftarjeta") &"'"
         rsUpdateEntry.Open strSQL, adoCon

         vMotivo="Diligencia en la ciudad de "&request("fciudad")
         vChofer=request("fchofer")
         vTarjeta=request("ftarjeta")
         vUnidad=rsUpdateEntry("ID_unidad")
         rsUpdateEntry.close

         strSQL="set dateformat ymd;select * from VoucherGas where DocDate='" &request("ffecha") &"' and Estacion='" &request("festacion") &"' and id_unidad='" & vUnidad &"'"
         response.write(strSQL&"<br>")

    end select
    response.write(strSQL&"<br>")
    rsUpdateEntry.Open strSQL, adoCon

    if not rsUpdateEntry.EOF then
        rsUpdateEntry.close        
        response.redirect("ShowContent.asp?contenido=95&msg=el registro ya existe y no puede duplicarlo")
    else
        rsUpdateEntry.close
    end if

    
    strSQL="insert into VoucherGas (ID,TarjetaCombustible,Estacion,ID_viatico,Motivo,DocDate,ID_unidad,Chofer,Logdate,id_usuario,litros) values ( (select max(ID) from VoucherGas)+1,'"& vTarjeta &"','"& request("festacion") &"','"& Vviatico &"','"& vMotivo &"','"&request("ffecha") &"','"& vUnidad &"','"& vChofer &"',getdate(),'"& session("session_id") &"',"& request("flitros") &")"

    response.write(strSQL&"<br>")
    rsUpdateEntry2.Open strSQL, adoCon

    response.redirect("ShowContent.asp?contenido=95&msg=la informacion ha sido registrada con exito")

 end if


if request("action")="3" then   'Form Delete'

    titulo="Borrar registro de vouchers"
    response.write("<P>&nbsp;</P>")
    DoTitulo(titulo)

  
    response.write("<form id='borrar' action='ShowContent.asp' method='POST'>")
    response.write("<input type='hidden' name='contenido' value='95'>")
    response.write("<input type='hidden' name='action' value='4'>")
    response.write("<input type='hidden' name='ID' value='"&request("ID")&"'>")

    CreateTable(400)
    RowIn
    strSQL="select *,.dbo.fn_GetMonthName(Docdate,'spanish') as Fecha from VoucherGas where ID="&request("ID")
    rsUpdateEntry.Open strSQL, adoCon

    response.write("<td width='50%' class='td-r subtitulo'>Tarjeta Combustible:</td><td width='50%' class='td-l'>"& rsUpdateEntry("TarjetaCombustible") &"</td></tr>")
    response.write("<td class='td-r subtitulo'>Fecha:</td><td class='td-l'>"& rsUpdateEntry("Fecha") &"</td></tr>")
    response.write("<td class='td-r subtitulo'>NumViatico:</td><td class='td-l'>"& rsUpdateEntry("id_viatico") &"</td></tr>")
    response.write("<td class='td-r subtitulo'>Motivo:</td><td class='td-l'>"& rsUpdateEntry("motivo") &"</td></tr>")
    response.write("<td class='td-r subtitulo'>Capturado en:</td><td class='td-l'>"& rsUpdateEntry("LogDate") &"</td></tr>")
    response.write("<td class='td-r subtitulo'>Confirma que desea borrarlo:</td><td class='td-l'><select name='decision'><option value=1>si</option></select></td></tr>") 
    response.write("<td class='td-r' colspan=2>&nbsp;</td></tr>") 
    response.write("<td class='td-c' colspan=2><input type='submit' value='borrar'></td></tr>") 
    rowout        
    response.write("</form></table>")
    closeTable

    rsUpdateEntry.close


end if


if request("action")="4" then   'Delete UP'

     strSQL="delete VoucherGas where ID="&request("ID")
     response.write strSQL

     rsUpdateEntry.Open strSQL, adoCon
     response.redirect("ShowContent.asp?contenido=95&msg=registro de voucher borrado")

end if



end sub


sub ShowVouchers

  titulo="Ultimos vouchers capturados"
  doTitulo(titulo)

  Dim strSQL

  strSQL="select a.ID,a.TarjetaCombustible as TarjtCombstbl,a.Estacion,case when convert(varchar,a.ID_viatico)='0' then 'N/A' else convert(varchar,a.ID_viatico) end as NumViatico,a.Motivo,.dbo.fn_GetMonthName(a.docdate,'spanish') as Fecha,a.id_unidad,substring(b.Descripcion,1,20) as Descripcion,Litros,a.LogDate as FechaCaptura from VoucherGas a left join Flotilla b on a.id_unidad=b.id_unidad WHERE a.ID>0 order by a.DocDate desc,a.Logdate desc"
  'Response.write strSQL

   Const adCmdText = &H0001
   Const adOpenStatic = 3
   nPage=0
   registros=1

  rsUpdateEntry.Open strSQL, adoCon
  rsUpdateEntry2.Open strSQL, adoCon, adOpenStatic, adCmdText

  rsUpdateEntry2.PageSize = 18
  nPageCount = rsUpdateEntry2.PageCount

  Dim Campos(10)
  Dim Contador
  contador=1

  CreateTable(1000)
  rowin
         For Each rsUpdateEntry in rsUpdateEntry.Fields
            Response.Write "<td class='subtitulo td-c'>" & rsUpdateEntry.Name & "</td>"
            campos(contador)=rsUpdateEntry.Name
            contador=contador+1
         Next
         Response.Write "<td class='subtitulo td-c'>&nbsp;</td></tr>"


  if not rsUpdateEntry2.EOF then
       if request("vPage")<>"" then
            nPage = int(trim(request("vPage")))
        else
            nPage=1
        end if
      rsUpdateEntry2.AbsolutePage=npage

      response.write("<tr><td colspan=3 class='td-r'><B><U>PAGINAS:</U></B></td><td colspan=6 class='td-l'>")
      for i=1 to nPageCount step 1
            if i<>npage then  
                response.write("<a href='/modules/ShowContent.asp?contenido="& request("contenido") &"&vPage="&i &"'>") 
                response.write( i &"</A>&nbsp;&nbsp;&nbsp;&nbsp;") 
            end if
      next
      response.write("</td></tr>")
  
  else

      NoRegistros

  end if
  

       

         while not rsUpdateEntry2.EOF  AND registros<=18
           rowin
             Response.write("<td class='td-r fontmed'>"& rsUpdateEntry2("ID") &"</td>")
             for contador=2 to 4
               Response.write("<td class='td-r fontmed'>"& rsUpdateEntry2(campos(contador)) &"</td>")
             next
             for contador=5 to 5
               Response.write("<td class='td-r fonttiny'>"& rsUpdateEntry2(campos(contador)) &"</td>")
             next
             for contador=6 to 7
               Response.write("<td class='td-r fontmed'>"& rsUpdateEntry2(campos(contador)) &"</td>")
             next
             for contador=8 to 9
               Response.write("<td class='td-r fonttiny'>"& rsUpdateEntry2(campos(contador)) &"</td>")
             next
             'Response.write("<td class='td-r fontmed'><a href='/modules/ShowContent.asp?contenido=94&xml="& rsUpdateEntry2("NombreXML") &"'>"& rsUpdateEntry2("NombreXML") &"</a></td>")
             'for contador=14 to 14
              Response.write("<td class='td-r fontmed'><a href='ShowContent.asp?contenido=95&action=3&ID="&rsUpdateEntry2("ID")&"'><img src='/img/b_drop.png' border='0' with='11px' height='11px' alt='borrar' title=''></td></tr>") 
             'next
           rowout
             rsUpdateEntry2.MoveNext
             registros=registros+1
         wend
         rsUpdateEntry2.close

 
  closeTable

end sub


sub notificaciones  'contenido 98'

     titulo="DOCUMENTOS DE MARKETING EN SAP"
     doTitulo(titulo)

     strSQL="exec SP_Notificaciones"
     'response.write strSQL
     rsUpdateEntry.Open strSQL, adoCon4
     response.write("<P class=alert style='text-align:center'>se accion&oacute; el env&iacute;o de notificaciones</P>")

     CreateTable(800)  %>
          
                 <tr><td width="12%" class="td-c submenu"><a href='helpdesk.asp?items=pedido'>Pedidos</a></td>
                     <!-- <td width="15%" class="td-c submenu"><a href='helpdesk.asp?items=OC&DocType=S'>OC-Servicios</a> </td> -->
                     <td width="12%" class="td-c submenu"><a href='ShowContent.asp?contenido=99'>Anticipos</a> </td>
                     <td width="14%" class="td-c submenu"><a href='helpdesk.asp?items=remision'>Remisiones</a> </td>
                     <td width="12%" class="td-c submenu"><a href='helpdesk.asp?items=factura'>Facturas</a> </td>
                     <td width="2%" class="td-c submenu">&nbsp; </td>

                     <td width="12%" class="td-c submenu"><a href='helpdesk.asp?items=OC&DocType=I'>OC-Materiales</a> </td>                                    
                     <td width="2%" class="td-c submenu">&nbsp; </td>

                     <td width="9%" class="td-c submenu"><a href='#'><a href='helpdesk.asp?items=entrada'>Entradas </a> </td>
                     <td width="9%" class="td-c submenu"><a href='#'>Traslados </a></td>
                     <td width="15%" class="submenu centrar" style="vertical-align: center"><a href="helpdesk.asp?action=1"><img src="/img/ejercer.png" width='18px' alt='enviar notificaciones' title='enviar notificaciones'>&nbsp; notificaciones</a></td>
     <%
     CloseTable

end sub




sub anticipos  'contenido 99'   anticipos parciales
 
   doAlert
   titulo="Control de Anticipos"
   response.write("<P>&nbsp;</P>")
   doTitulo(titulo)

   
   strSQL="select a.DocNum as Anticipo,a.DocDate,a.RazonSocial,a.CardName,a.SlpEmail,  CONVERT(VARCHAR,CONVERT(MONEY,a.subtotal),1)  as subtotal, CONVERT(VARCHAR,CONVERT(MONEY,a.total),1)  as Total,a.Pedido,a.CopiarCFDi,b.Ruta from Notificaciones a left join Repositorio b on a.RazonSocial=b.RazonSocial and a.cardCode=b.cardCode  where a.tipo='anticipo' and (DATEDIFF(MONTH,a.DocDate,getdate())<=2 OR a.CopiarCFDi=0 )"

   strSQL= strSQL &" order by DocDate Desc,DocNum desc "   
   'response.write strSQL
  

   rsUpdateEntry.Open strSQL, adoCon4
   rsUpdateEntry2.Open strSQL, adoCon4

   CreateTable(940)
   Dim Campos(10)
   i=1

   
   For Each rsUpdateEntry in rsUpdateEntry.Fields           
        campos(i)=rsUpdateEntry.Name
        i=i+1      
   Next

   rowin
   for i=1 to 9
   	    response.write("<th class='subtitulo td-c fontsmall'>"& campos(i) &"</th>")
   next
   rowout
   
   while not rsUpdateEntry2.EOF   
      if trim(rsUpdateEntry2("CopiarCFDi"))="0" then
            vcolor="#E3E1E1"
      else
            vcolor="#FFFFFF"
      end if   
      rowin  
      for i=1 to 7
           Response.write("<td class='td-r fonttiny' style='background-color:"&vcolor&"'>" & rsUpdateEntry2(campos(i)) &"</td>")
      next 
     
      if trim(rsUpdateEntry2("Pedido"))<>"" then
          Response.write("<td class='td-r fonttiny' style='background-color:"&vcolor&"'><a href='http://192.168.0.206/repositorio/"&rsUpdateEntry2("RazonSocial")&"/"& rsUpdateEntry2("Ruta") &"/" & rsUpdateEntry2("Pedido") &"' target='repositorio'>" & rsUpdateEntry2("Pedido") &"</a></td>")
      else 
          Response.write("<td class='td-r fonttiny' style='background-color:"&vcolor&"'> - </a></td>")
      end if      
      Response.write("<td class='td-r fonttiny' style='background-color:"&vcolor&"'>" & rsUpdateEntry2("CopiarCFDi") &"</a></td>")      
      RowOut      
      rsUpdateEntry2.MoveNext      
  wend
  
  closeTable

end sub




sub ViaticoChange   'contenido 100'

       titulo="Actualice remisiones de un viatico"
       DoAlert
       DoTitulo(titulo)

       strSQL="select a.Beneficiario,a.motivo,a.Notas,a.id_unidad,b.* from VIATICOS_R1 a inner join ViaticoRemision b on a.ID=b.ID where a.ID="&request("ID")
       'response.write strSQL
       rsUpdateEntry.Open strSQL, adoCon
       if not rsUpdateEntry.EOF then vUnidad=rsUpdateEntry("ID_unidad")

       CreateTable(660)

       
       response.write("<tr><td class='td-r subtitulo'># de viatico</td><td class='td-l'>"& rsUpdateEntry("ID") &"</td></tr>")
       response.write("<tr><td class='td-r subtitulo'>Beneficiario:</td><td class='td-l'>"& rsUpdateEntry("beneficiario") &"</td></tr>")
       response.write("<tr><td class='td-r subtitulo'>Motivo:</td><td class='td-l'>"& rsUpdateEntry("Motivo") &"</td></tr>")
       response.write("<tr><td class='td-r subtitulo'>Notas:</td><td class='td-l'>"& rsUpdateEntry("Notas") &"</td></tr>")
       response.write("<tr><td class='td-r subtitulo'>Raz&oacute;n social que suministra:</td><td class='td-l'>"& rsUpdateEntry("RS") &"</td></tr>")

       response.write("<Form id='form' method='POST' action='ShowContent.asp'>")
       response.write("<input type='hidden' name='contenido' value=100>")
       response.write("<input type='hidden' name='action' value=2>")
       response.write("<input type='hidden' name='ID' value="& rsUpdateEntry("ID") &">")
       response.write("<input type='hidden' name='RS' value="& rsUpdateEntry("RS") &">")


       if request("action")="" then

           i=1
           while not rsUpdateEntry.EOF
               response.write("<tr><td class='td-r subtitulo'>remision "&i&":</td><td class='td-l'><input type='number' name='remision"&i&"' value="&rsUpdateEntry("remision") &"></td></tr>")
               response.write("<input type='hidden' name='valor"&i&"' value='"&rsUpdateEntry("remision")&"'>")
               rsUpdateEntry.MoveNext
               i=i+1
           wend
           rsUpdateEntry.close
           i=i-1
           response.write("<input type='hidden' name='cantidad' value='"& i &"'>")
           response.write("<tr><td class='td-r subtitulo'>UNIDAD:</td><td class='td-l'><select name='fUnidad'>")
           strSQL="select * from flotilla where Activado=1 and tipo_activo!=6"
           rsUpdateEntry5.Open strSQL, adoCon
           while not rsUpdateEntry5.EOF
           	   if vUnidad=rsUpdateEntry5("ID_unidad") then
                   response.write("<option value='"& rsUpdateEntry5("ID_unidad") &"' SELECTED>"& rsUpdateEntry5("ID_unidad")&"-" & rsUpdateEntry5("Fabricante") &"</option>")
               else
                   response.write("<option value='"& rsUpdateEntry5("ID_unidad") &"'>"& rsUpdateEntry5("ID_unidad")&"-" & rsUpdateEntry5("Fabricante") &"</option>")
               end if
               rsUpdateEntry5.MoveNext
           wend
           rsUpdateEntry5.close
           response.write("</select></td></tr>")

           response.write("<tr><td class='td-c' colspan=2><input type='submit' value='actualizar'></form></td></tr>")
           closeTable

           response.write("<P class='fonttiny td-c'>en caso de ingresar una remision que no exista, el procedimiento quedara nulo y no provocara ningun cambio")
        end if


       if request("action")="2" then
          flag=1

          for i=1 to int(trim(request("cantidad"))) step 1
              vstring="remision"&i             
              strSQL="select * from ODLN where docnum="&request(vstring) &" and CANCELED='N' "
              'response.write(strSQL&"<BR>")
              if request("RS")="DMX" then
                  rsUpdateEntry2.Open strSQL, adoCon2  'DMX'
              else
                  rsUpdateEntry2.Open strSQL, adoCon3  'DFW'
              end if
              if rsUpdateEntry2.EOF then
                   flag=0
              end if
              rsUpdateEntry2.close
          next

          vMotivo=""

          if flag=0 then
              response.redirect("ShowContent.asp?contenido=100&ID="&request("ID")&"&msg=ingreso una remision que no existe o esta cancelada, no se produjo algun cambio" )
          else
              for i=1 to int(trim(request("cantidad")))  step 1
                vstring="remision"&i
                vstring2="valor"&i

                strSQL="update top(1) ViaticoRemision set remision="&request(vstring) &" where remision=" & request(vstring2) &" and ID="&request("ID")
                rsUpdateEntry5.Open strSQL, adoCon  
                vMotivo=vMotivo&" "&request(vstring)
                response.write(strSQL&"<BR>")

              next

              strSQL="update VIATICOS_R1 set motivo='"&vMotivo&"',ID_unidad='"& request("fUnidad") &"' where ID="&request("ID")
              rsUpdateEntry7.Open strSQL, adoCon 
              response.write strSQL

              response.redirect("ShowContent.asp?contenido=12&msg=se actualizaron remisiones para un viatico" )

          end if


       end if

 

end sub








'==============================================================================================================================================='





function Sleep(seconds)
            set oShell = CreateObject("Wscript.Shell")
            cmd = "%COMSPEC% /c timeout " & seconds & " /nobreak"
            oShell.Run cmd,0,1
End function



Sub EnterosDecimales
         vstring=""                  
         longitud=len(session("entero"))
         if longitud=5 then
               select case mid(session("entero"),1,2)
               case "10" vstring="diez mil "
               case "11" vstring="once mil "
               case "12" vstring="doce mil "
               case "13" vstring="trece mil "
               case "14" vstring="catorce mil "
               case "15" vstring="quince mil "  
               case "16" vstring="dieciseis mil " 
               case "17" vstring="diecisiete mil " 
               case "18" vstring="dieciocho mil " 
               case "19" vstring="diezcinueve mil "         
               case else
                   Select case mid(session("entero"),1,1)
                   case "2" vstring="veinte "
                   case "3" vstring="treinta "
                   case "4" vstring="cuarenta "
                   case "5" vstring="cincuenta "
                   case "6" vstring="sesenta "
                   case "7" vstring="setenta "
                   case "8" vstring="ochenta "
                   case "9" vstring="noventa "
                   end select
                   
               end select
               if mid(session("entero"),2,1)<>"0" and session("entero")>=20000  then vstring=vstring &"y " end if
          end if
          'response.write("entero=" & trim(session("entero")) &"<BR>")
          'response.write("decimales=" & trim(session("decimales")) &"<BR>")

          if longitud>3 and ( int(trim(session("entero")))<10000 or int(trim(session("entero"))) >=20000 ) then
                  Select case mid(session("entero"),longitud-3,1)
                   case "0" vstring=vstring &" mil "
                   case "1" vstring=vstring &"un mil "
                   case "2" vstring=vstring &"dos mil "
                   case "3" vstring=vstring &"tres mil "
                   case "4" vstring=vstring &"cuatro mil "
                   case "5" vstring=vstring &"cinco mil "
                   case "6" vstring=vstring &"seis mil "
                   case "7" vstring=vstring &"siete mil "
                   case "8" vstring=vstring &"ocho mil "
                   case "9" vstring=vstring &"nueve mil "
                   end select
            end if

           if longitud>2 then            
                   Select case mid(session("entero"),longitud-2,1)
                   case "1" vstring=vstring &"ciento "
                   case "2" vstring=vstring &"dos cientos "
                   case "3" vstring=vstring &"tres cientos "
                   case "4" vstring=vstring &"cuatro cientos "
                   case "5" vstring=vstring &"quinientos "
                   case "6" vstring=vstring &"seis cientos "
                   case "7" vstring=vstring &"siete cientos "
                   case "8" vstring=vstring &"ocho cientos "
                   case "9" vstring=vstring &"nueve cientos "
                   end select
                                    
            end if

            if longitud>1 then            
                   Select case mid(session("entero"),longitud-1,2)
                   case "10" vstring=vstring &"diez "
                   case "11" vstring=vstring &"once "
                   case "12" vstring=vstring &"doce "
                   case "13" vstring=vstring &"trece "
                   case "14" vstring=vstring &"catorce "
                   case "15" vstring=vstring &"quince "  
                   case "16" vstring=vstring &"dieciseis " 
                   case "17" vstring=vstring &"diecisiete " 
                   case "18" vstring=vstring &"dieciocho " 
                   case "19" vstring=vstring &"diezcinueve "         
                   case else

                         Select case mid(session("entero"),longitud-1,1)
                         case "2" vstring=vstring &"veinte "
                         case "3" vstring=vstring &"treinta "
                         case "4" vstring=vstring &"cuarenta "
                         case "5" vstring=vstring &"cincuenta "
                         case "6" vstring=vstring &"sesenta "
                         case "7" vstring=vstring &"setenta "
                         case "8" vstring=vstring &"ochenta "
                         case "9" vstring=vstring &"noventa "
                         end select   

                         Select case mid(session("entero"),longitud,1)
                         case "1" vstring=vstring &"uno "
                         case "2" vstring=vstring &"dos "
                         case "3" vstring=vstring &"tres "
                         case "4" vstring=vstring &"cuatro "
                         case "5" vstring=vstring &"cinco "
                         case "6" vstring=vstring &"seis "
                         case "7" vstring=vstring &"siete "
                         case "8" vstring=vstring &"ocho "
                         case "9" vstring=vstring &"nueve "
                         end select       


                    end select
               end if

         session("vstring")=vstring&"con "& session("decimales") &"/100 en MN"
End sub  


Sub ShowLogin 
       
       'response.write ( "modulo="&Session("modulo") &"<BR>")
       'response.write ( "session_modulo=" & Session("session_modulo") &"<BR>")

       Session("item") = request("item") 
       if request("modulo")<>"" then
           Session("modulo") = request("modulo") 
       end if
       response.write("<P>&nbsp;</P>" )

       CreateTable(360)
%>


        <form id="login" method="post" action="ShowContent.asp">
        <input type="hidden" id="contenido" name="contenido" value="10">                     
       
        <tr><td colspan=2 class=subtitulo>Inicio de sesi&oacute;n </td></tr>
         <tr><td colspan=2>&nbsp;</td></tr>
        <tr><td width="40%" class=td-r style="font-size: 13px;">usuario: </td><td width="60%" class=td-l><input type="text" name="user" id="user" class="anchox1"></td></tr>
        <tr><td  class=td-r style="font-size: 13px;">contrase&ntilde;a: </td><td  class=td-l><input type="password" name="pass" id="pass" class="anchox1"></td></tr>
        <tr><td class=td-c colspan="2">&nbsp; </td></tr>
        <tr><td class=td-c colspan=2><input type="submit" value="inicio" border=0 style="width:60px; text-align: center"> </td></tr>
        <tr><td class="td-c fonttiny" colspan=2>requiere sesi&oacute;n abierta</td></tr>
        </form>
        </table>
         
       <%              

end sub  





Sub CheckLogin 'contenido 10'

        if Session("modulo")="Controlviaticos" then
               Session("modulo")="viaticos"
               flag=1
        end if
        if Session("modulo")="ControlReembolsos" then
               Session("modulo")="reembolsos"
               flag=2
        end if  
       
        strSQL = "SELECT A.*,B.clave FROM permisos A INNER JOIN cat_usuario B ON A.id_usuario=B.id_usuario where A.id_usuario='" & UCase(request("user")) & "' AND A.modulo='"& Session("modulo") &"'"
        response.write strSQL
        rsUpdateEntry.Open strSQL, adoCon

  
         if not rsUpdateEntry.EOF then    'EXISTE EL USUARIO Y NO ESTA DADO DE BAJA
                 response.write("etap=1" &"<BR>")
                 if UCase(rsUpdateEntry("clave"))=UCase(Request("pass")) then
                          response.write("etap=2" &"<BR>")                          
                                                                               
                          Session("session_id") = UCase(Request("user"))    
                          Session("session_modulo") = Session("modulo")                                                
                          rsUpdateEntry.close  
                   
                          select case  flag
                          case 1    response.redirect("ShowContent.asp?contenido=12")  
                          case 2    response.redirect("ShowContent.asp?contenido=20")   
                          end select

                          response.write("sesion solicitada=" & Session("modulo")  )  
                          response.write("sesion logeada=" & Session("session_modulo")  )

                          select case  Session("modulo")
                          case "humres"         response.redirect("ShowContent.asp?contenido=9")
                          case "viaticos"       response.redirect("ShowContent.asp?contenido=13") 
                          case "reembolsos"     response.redirect("ShowContent.asp?contenido=16")     
                          case "H_Incidencias"  response.redirect("ShowContent.asp?contenido=32")                   
                          case "ControlEnvios"  response.redirect("ShowContent.asp?contenido=36")    
                          case "checada"        response.redirect("ShowContent.asp?contenido=41")
                          case "ModuloADD"      response.redirect("ShowContent.asp?contenido=42")  
                          case "ChecarIntranet"      response.redirect("ShowContent.asp?contenido=44")
                          case "SolicVacaciones"     response.redirect("ShowContent.asp?contenido=46")  
                          case "cobranza"     response.redirect("ShowContent.asp?contenido=52") 
                          case "Alarmas"     response.redirect("ShowContent.asp?contenido=54") 
                          case "PuntoEquilibrio"     response.redirect("ShowContent.asp?contenido=55")  
                          case "TipoCambio"     response.redirect("ShowContent.asp?contenido=57") 
                          case "ControlMantenimiento"  response.redirect("ShowContent.asp?contenido=61") 
                          case "TaskManager"  response.redirect("ShowContent.asp?contenido=62") 
                          case "CorteInventario"  response.redirect("ShowContent.asp?contenido=67") 
                          case "Inventario"   DoInventario
                          case "InformeEnvios"  response.redirect("ShowContent.asp?contenido=71") 
                          case "Contactos"  response.redirect("ShowContent.asp?contenido=72")    
                          case "DistribucionCuentas"  response.redirect("ShowContent.asp?contenido=74")  
                          case "SolicitudCompra" response.redirect("ShowContent.asp?contenido=77")  
                          case "AprobarSolicitudCompra" response.redirect("ShowContent.asp?contenido=78")
                          case "ShowSolicitudCompras" response.redirect("ShowContent.asp?contenido=79")   
                          case "ABCIngresos"    response.redirect("ShowContent.asp?contenido=80")
                          case "SolicitarCumple"  response.redirect("ShowContent.asp?contenido=58")
                          case "AnalisisUtilidadPedido"  response.redirect("ShowContent.asp?contenido=85")
                          case "Flotilla"  response.redirect("ShowContent.asp?contenido=87&action=1")
                          case "ControlViajes"  response.redirect("ShowContent.asp?contenido=88&action=1")
                          case "RegistroEmbarques"  response.redirect("ShowContent.asp?contenido=90")
                          case "EnvioConsolidado" response.redirect("ShowContent.asp?contenido=91")
                          case "voucherGas" response.redirect("ShowContent.asp?contenido=95")  
                          case "DIOT" response.redirect("ShowContent.asp?contenido=103")
                          case "PAGOS" response.redirect("ShowContent.asp?contenido=104&action=1")
                          case "EliminaFechaEmbarque"  response.redirect("ShowContent.asp?contenido=107")
                          case "AnalisisUtilidadHistorico"  response.redirect("ShowContent.asp?contenido=108")
                          case "BlogVentas"  response.redirect("ShowContent.asp?contenido=114")
                          case "feeds"  response.redirect("ShowContent.asp?contenido=118")
                          case "PotterRoemer"  response.redirect("ShowContent.asp?contenido=126")
                          case "ControlPagos"  response.redirect("ShowContent.asp?contenido=128")
                          case "SolicitarPagar"  response.redirect("ShowContent.asp?contenido=131")
                          end select
                                            

                 else
                     response.write("etap=3" &"<BR>")
                     rsUpdateEntry.close
                     response.redirect("ShowContent.asp?contenido=0&msg=credenciales incorrectas")
                                              
                 end if

         
         else 
            response.write("etap=4" &"<BR>")
            rsUpdateEntry.close
            response.redirect("ShowContent.asp?contenido=0&msg=usuario no tiene acceso al modulo: "& session("modulo") )
         end if   
        
End Sub


sub modulosHumRes
   response.write("<P style='font-size:4px'>&nbsp;</P>")   
   response.write("<H1>M&oacute;dulos disponibles:</H1>")      %>

    <table width="500px" border=0 cellpadding="1" cellspacing="1" align=center>

             <tr><td width="20%" class="td-c submenu"><a href='ShowContent.asp?contenido=37'>D&iacute;a inh&aacute;bil</a></td>   
                 <td width="20%" class="td-c submenu"><a href='ShowContent.asp?contenido=39'>Horarios</a></td> 
                 <td width="20%" class="td-c submenu"><a href='ShowContent.asp?contenido=40&accion=0'>Empleados</a> </td>                
                 <td width="20%" class="td-c submenu"><a href='ShowContent.asp?contenido=45'>Vacaciones</a> </td>                 
                 <td width="20%" class="td-c submenu"><a href='#'>&nbsp;</a> </td>
                            
</table>  
<%
end sub  




sub DoInventario

     strSQL="exec [dbo].[SP_InventarioComercial] " 
     response.write strSQL
     rsUpdateEntry.Open strSQL, adoCon4 'SBOTemp'    
     response.redirect("ShowContent.asp?contenido=69") 

end sub




session("campos")=1

sub CamposRS1
     session("campos")=1
     For Each rsUpdateEntry in rsUpdateEntry.Fields
         Campos( session("campos") )=rsUpdateEntry.Name         
         session("campos")=session("campos")+1
     Next
     session("campos")=session("campos")-1

              For i=1 to session("campos")
                Response.Write "<td class=subtitulo>" & campos(i) & "</td>"      
              Next
end sub


sub CamposRS3
     session("campos")=1
     For Each rsUpdateEntry3 in rsUpdateEntry3.Fields
         Campos( session("campos") )=rsUpdateEntry3.Name         
         session("campos")=session("campos")+1
     Next
     session("campos")=session("campos")-1

              For i=1 to session("campos")
                Response.Write "<td class=subtitulo>" & campos(i) & "</td>"      
              Next

end sub

sub CamposRS5
     session("campos")=1
     For Each rsUpdateEntry5 in rsUpdateEntry5.Fields
         Campos( session("campos") )=rsUpdateEntry5.Name         
         session("campos")=session("campos")+1
     Next
     session("campos")=session("campos")-1

              For i=1 to session("campos")
                Response.Write "<td class=subtitulo>" & campos(i) & "</td>"      
              Next
end sub



sub ShowValoresRS2
    n=1

    if rsUpdateEntry2.EOF then
         response.write("<tr><td colspan=20 class='td-c'>no existen registros con este criterio</td></tr>")
    else   
            While Not rsUpdateEntry2.EOF
            Response.Write "<tr>"
                      For i=1 to session("campos")
                            if n=1 then
                               Response.Write "<td class='td-r'>" & rsUpdateEntry2( campos(i) ) & "</td>"     
                            else
                               Response.Write "<td class='td-r' style='background-color:#E1DFDF;'>" & rsUpdateEntry2( campos(i) ) & "</td>"    
                            end if 
                      Next
                      n=n+1
                      if n=3 then 
                          n=1 
                      end if
            rsUpdateEntry2.MoveNext
            Response.Write "</tr>"
            Wend
    end if
    rsUpdateEntry2.close
end sub     



sub ShowValoresRS4
    While Not rsUpdateEntry4.EOF
    Response.Write "<tr>"
              For i=1 to session("campos")
                Response.Write "<td class='td-r'>" & rsUpdateEntry4(campos(i)) & "</td>"      
              Next
    rsUpdateEntry4.MoveNext
    Response.Write "</tr>"
    Wend
    rsUpdateEntry4.close
end sub   


sub ShowValoresRS6
    While Not rsUpdateEntry6.EOF
    Response.Write "<tr>"
              For i=1 to session("campos")
                Response.Write "<td class='td-l'>" & rsUpdateEntry6(campos(i)) & "</td>"      
              Next
    rsUpdateEntry6.MoveNext
    Response.Write "</tr>"
    Wend
    rsUpdateEntry6.close
end sub   


sub CambioModulo
	  if trim(Session("session_modulo"))<>"" and trim(Session("modulo"))<>""  and trim(Session("session_modulo"))<> trim(Session("modulo")) then
            'response.write("<P class='alert'>--"& Session("session_modulo") &"!"& Session("modulo") &"--</P>")
            strSQL = "SELECT A.*,B.clave FROM permisos A INNER JOIN cat_usuario B ON A.id_usuario=B.id_usuario where A.id_usuario='" & UCase(Session("session_id")) & "' AND A.modulo='"& Session("modulo") &"'"
            'response.write(strSQL &"<br>")
            rsUpdateEntry.Open strSQL, adoCon
            if rsUpdateEntry.EOF then
                    'response.write("no tiene permiso")
            else
                    'response.write("si tiene permiso")
                Session("session_modulo") = Session("modulo")
            end if
            rsUpdateEntry.close
	  end if	 
end sub

sub RowIn
	 response.write("<tr>")
end sub	 

sub RowOut
	 response.write("</tr>")
end sub	 

%>
