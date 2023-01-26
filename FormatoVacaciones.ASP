<!--#include virtual="/config/conn.asp"-->

<html>
<!DOCTYPE html>
<!--[if IE 9 ]> <html lang="mx-es" prefix="og: http://ogp.me/ns#" class="ie9 loading-site no-js bg-fill"> <![endif]-->
<!--[if IE 8 ]> <html lang="mx-es" prefix="og: http://ogp.me/ns#" class="ie8 loading-site no-js bg-fill"> <![endif]-->
<!--[if (gte IE 9)|!(IE)]><!--><html lang="mx-es" prefix="og: http://ogp.me/ns#" class="loading-site no-js bg-fill"> <!--<![endif]-->
<head>

<link rel="icon" href="/favicon.png" sizes="32x32" />

<meta name="viewport" content="width=device-width, initial-scale=1">
<meta http-equiv="content-type" content="text/html; charset=UTF-8" />
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">


<title>Intranet Deltaflow</title>
<link rel="stylesheet" href="/CSS/style.css">

<link href='//fonts.googleapis.com/css?family=Playfair+Display:400,700,900|Lora:400,700|Open+Sans:400,300,700|Oswald:700,300|Roboto:400,700,300|Signika:400,700,300' rel='stylesheet' type='text/css'>

<script type="text/javascript">
    WebFontConfig = {
      google: { families: [ "Lato:regular,default","Lato:regular,default","Lato:regular,default","Lato:regular,default", ] }
    };
    (function() {
      var wf = document.createElement('script');
      wf.src = 'https://ajax.googleapis.com/ajax/libs/webfont/1/webfont.js';
      wf.type = 'text/javascript';
      wf.async = 'true';
      var s = document.getElementsByTagName('script')[0];
      s.parentNode.insertBefore(wf, s);
    })(); </script>


</head>
<style type="text/css">
    @media print {
        .footer,
        #non-printable {
            display: none !important;
        }
        #printable {
            display: block;
        }
    }
    .strong{font-weight: 600!important}
</style>
<body style="margin: 0px; ">
<%

    response.write("<table cellpadding=6 cellpadding=3 align='center' border='1' style='border-color:black;width:930px'>")
    response.write("<tr><td width='33%' style='font-size:3px;height:3px'>&nbsp;</td><td width='33%' style='font-size:3px;height:3px'>&nbsp;</td><td width='33%' style='font-size:3px;height:3px'>&nbsp;</td></tr>")
    response.write("<tr><td colspan=3 class='td-c'>")
        DoTitulo("SOLICITUD DE VACACIONES")
    response.write("</tr><tr>")


    if request("ID") then
        flag=0
        strSQL="select max(identificador) as identificador from SolicitudVacaciones where Activo=1 and ID="&request("ID")  
        'response.write(strSQL&"<BR>") 
        rsUpdateEntry.Open strSQL, adoCon
         
        if not rsUpdateEntry.EOF then
              flag=1
              vIdentificador=trim(rsUpdateEntry("identificador")) 
        end if
        rsUpdateEntry.close

        if request("identificador")<>"" then vIdentificador=trim(request("identificador")) 

        strSQL="select a.nombres+' '+a.paterno+' '+a.materno as colaborador,.dbo.fn_GetMonthName(a.FechaIngreso,'spanish') as ingreso,a.*,b.sucursal,c.departamento,d.Puesto,round( DATEDIFF(YEAR,a.FechaIngreso,getdate()),1) as antiguo  from Empleados a inner join cat_sucursal b on a.clave_sucursal=b.clave_sucursal inner join cat_departamento c on a.Clave_depto=c.Clave_depto inner join cat_puesto d on a.clave_puesto=d.clave_puesto  where a.FechaEgreso is null and a.ID="&request("ID")    '121
        'response.write strSQL
        rsUpdateEntry.Open strSQL, adoCon
       

        if request("identificador")<>"" then
             strSQL="select a.*,cantidad=(select count(Fecha) from SolicitudVacaciones where ID="& request("ID") &" and cancelDate is null and identificador="&vIdentificador&" ) from SolicitudVacaciones a where a.ID="& request("ID") &" and identificador="&vIdentificador&" and cancelDate is null order by a.Fecha "
        else
            strSQL="select a.*,cantidad=(select count(Fecha) from SolicitudVacaciones where Activo=1 and ID="& request("ID") &" and identificador="&vIdentificador&" ) from SolicitudVacaciones a where Activo=1 and a.cancelDate is null and a.ID="& request("ID") &" and identificador="&vIdentificador&"  order by a.Fecha "

        end if
        'response.write strSQL

        Const adCmdText = &H0001
        Const adOpenStatic = 3

        rsUpdateEntry2.Open strSQL, adoCon, adOpenStatic, adCmdText

        select case rsUpdateEntry("id_empresa")
        case "DMX"  vLogo="logo-DMX.png"                
        case "SOL"  vLogo="logo-servPro.png"            
        end select
        response.write("<tr><td colspan=2 rowspan=2 class='td-c'><img src='/images/logo-DMX.png' border=0 width='400px'></td><td class='td-c strong'>FECHA DE SOLICITUD</td></tr>")
        response.write("<tr><td class='td-c'>")
           createtable(210)
           response.write("<tr><td class='td-c fonttiny' width='70px'>A&Ntilde;O</td><td class='td-c fonttiny' width='70px'>MES</td><td class='td-c fonttiny' width='70px'>DIA</td></tr>")           
           response.write("<tr><td class='td-c fonttiny strong' colspan=3>"& rsUpdateEntry2("LogDate") &"</td></tr>")
           CloseTable
        response.write("</td></tr>")
        response.write("<tr><td class='td-r fontbig'>Nombre del colaborador: </td><td class='td-l' colspan=2 class='fontbig strong'><b>"& rsUpdateEntry("colaborador") &"</b></td></tr>")
        response.write("<tr><td class='td-c fontbig'>Sucursal: <b>"& rsUpdateEntry("sucursal") &"</b></td><td class='td-c' class='fontbig'>Area: <b>"& rsUpdateEntry("departamento") &"</b></td><td class='td-c' class='fontbig'>Puesto: <b>"& rsUpdateEntry("Puesto") &"</b></td></tr>")
        response.write("<tr><td class='td-c fontbig'>Fecha de ingreso: <b>"& rsUpdateEntry("ingreso") &"</b></td><td class='td-l' class='fontbig'>Antig&utilde;edad: <b>" & rsUpdateEntry("antiguo") &" a&ntilde;os</b></td><td class='td-l' class='fontbig'>Per&iacute;do: <b>" & rsUpdateEntry2("Periodo") &"</b> </b></td></tr>")
         response.write("<tr><td class='td-l fontbig' colspan=3>&nbsp;</td></tr>")

        response.write("<tr><td class='td-l fontbig' colspan=3>Por medio de la presente, solicito me sea autorizado <B>"& rsUpdateEntry2("cantidad") &" d&iacute;a(s) a cuenta de mis vacaciones.</td></tr>")
        response.write("<tr><td class='td-l fontbig' colspan=3>Para ser gozadas en fecha inicial: &nbsp;&nbsp;&nbsp; <b>" & rsUpdateEntry2("Fecha") &"</b> &nbsp;&nbsp;&nbsp; a fecha final: &nbsp;&nbsp;&nbsp;<B>")
        rsUpdateEntry2.MoveLast
        response.write( rsUpdateEntry2("Fecha") &"</B> &nbsp;&nbsp;&nbsp;</td></tr>")        
        response.write("<tr><td class='td-l fontbig' colspan=3>Observaciones: <U>"& request("notas") &"</U></td></tr>")
        response.write("<tr><td class='td-l fontbig' colspan=3>&nbsp;</td></tr>")
        response.write("<tr><td class='td-l fontbig' colspan=3>&nbsp;</td></tr>")

        response.write("<tr><td class='td-c fontbig'>____________________________________________</td><td>&nbsp;</td>")
        response.write("<td class='td-c fontbig'>____________________________________________</td></tr>")

        response.write("<tr><td class='td-c fontmed'>FIRMA DEL COLABORADOR</td><td>&nbsp;</td>")
        response.write("<td class='td-c fontmed'>FIRMA DEL JEFE INMEDIATO</td></tr>")

        response.write("<tr><td class='td-c fontmed'>"& rsUpdateEntry("colaborador") &"</td><td>&nbsp;</td>")
        response.write("<td class='td-c fontmed'>NOMBRE: _________________________________</td></tr>")

        response.write("<tr><td class='td-l fontbig' colspan=3>&nbsp;</td></tr>")
        
        rsUpdateEntry.close

    end if 
    CloseTable
    response.write("<P>&nbsp;</P>")
    response.write("<table cellpadding=6 cellpadding=3 align='center' border='1' style='border-color:black;width:930px'>")
    response.write("<tr><td width='33%' style='font-size:3px;height:3px'>&nbsp;</td><td width='33%' style='font-size:3px;height:3px'>&nbsp;</td><td width='33%' style='font-size:3px;height:3px'>&nbsp;</td></tr>")
    response.write("<tr><td colspan=3 class='td-c'>")
        DoTitulo("USO EXCLUSIVO CAPITAL HUMANO")

    if request("ID")<>"" then
        strSQL="select * from PeriodosVacacionales where ID="& request("ID") &" and Periodo="& rsUpdateEntry2("Periodo") 
        rsUpdateEntry3.Open strSQL, adoCon

        response.write("</tr><tr><td class='td-c fontmed'>Antig&utilde;edad: <B>"& rsUpdateEntry3("AniosAntiguedad") &"</b></td>")       
        response.write("<td class='td-c fontmed'>D&iacute;as derecho del per&iacute;do "& rsUpdateEntry2("Periodo") &": <b>"& rsUpdateEntry3("DiasDerecho") &"</b></td><td class='td-c fontmed'>Dias Tomados: <B>"& rsUpdateEntry3("DiasTomados") &"</b></td></tr>")

        DiasDerecho=int(trim(rsUpdateEntry3("DiasDerecho")))
        DiasTomados=int(trim(rsUpdateEntry3("DiasTomados")))

        response.write("<tr><td class='td-l fontmed' colspan=3><B>DETALLE DIAS TOMADOS:</B><br> ")

        strSQL="SELECT Fechas=( SELECT .dbo.fn_GetMonthName(Fecha,'spanish')+'  |  ' FROM Incidencias  where ID="& request("ID") &" and INCIDENCIA=4 and Periodo="& rsUpdateEntry2("Periodo") &" order by Fecha for XML PATH ('')   )"
        rsUpdateEntry4.Open strSQL, adoCon

        response.write( rsUpdateEntry4("Fechas") &"</td></tr>")
        rsUpdateEntry4.close

        response.write("<tr><td class='td-l fontbig' colspan=2>&nbsp;</td>")
        response.write("<td class='td-c fontbig' colspan=2><B>Dias Restantes: "& DiasDerecho-DiasTomados &"</B></td></tr>")

        response.write("<tr><td class='td-l fontmed' colspan=2><B>DETALLE DIAS PROGRAMADOS:</B><br> ")
        strSQL="SELECT Fechas=( SELECT .dbo.fn_GetMonthName(Fecha,'spanish')+'  |  ' FROM SolicitudVacaciones  where ID="& request("ID") &" and Activo=1 and INCIDENCIA=4 and Periodo="& rsUpdateEntry2("Periodo") &" order by Fecha for XML PATH ('')   )"
        rsUpdateEntry4.Open strSQL, adoCon

        response.write( rsUpdateEntry4("Fechas") &"</td>")
        rsUpdateEntry4.close
        strSQL="SELECT (" & rsUpdateEntry3("DiasDerecho") &"-" & rsUpdateEntry3("DiasTomados") &"-" &" count(fecha) ) as calculo FROM SolicitudVacaciones  where ID="& request("ID") &" and Activo=1 and INCIDENCIA=4 and Periodo="& rsUpdateEntry2("Periodo") 
        'response.write strSQL
        rsUpdateEntry4.Open strSQL, adoCon

        response.write("<td class='td-c fontmed'>&nbsp;</td></tr>")
        'response.write("<td class='td-c fontmed'>D&iacute;as restantes despu&eacute;s de estas vacaciones: <b>" & rsUpdateEntry4("calculo") &"</td></tr>")        
        rsUpdateEntry4.close
        
    end if

    response.write("<tr><td class='td-l fontbig' colspan=3>&nbsp;</td></tr>")
    response.write("<tr><td class='td-l fontbig' colspan=3>&nbsp;</td></tr>")
    response.write("<tr><td class='td-c fontbig' colspan=3>________________________________________</td></tr>")
    response.write("<tr><td class='td-c fontbig' colspan=3>LIC. R. ISELA CERVANTES SANDOVAL</td></tr>")
    response.write("<tr><td class='td-c fontbig' colspan=3>GERENTE DE CAPITAL HUMANO</td></tr>")
    response.write("<tr><td class='td-l fontbig' colspan=3>&nbsp;</td></tr>")
    rsUpdateEntry2.close
    rsUpdateEntry3.close

    CloseTable

%>