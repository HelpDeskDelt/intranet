<!--#include virtual="/config/conn.asp"-->
<!--#include virtual="/header.asp"-->


<%
ruta="/docs/"&request("file")


titulo="mostrando archivo: "& request("file")
doTitulo(titulo)

response.write("<iframe id=marco src="& ruta &"></iframe>")   %>



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

