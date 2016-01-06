<html>
<head><title>CFP 1.16 - Ordem de Serviço</title></head>
<body bgcolor=#F0FFF0>
<form method=post action=acessorestrito.asp target=_self>
<table bgcolor=#C1CDC1 width=510 height=370 align=center cellspacing=0 cellpadding=0>
 <tr height=15>
  <td align=center>
   <%
    session.LCID = 1046
    response.expires = -1000 
    response.write " <i><font face=verdana, arial size=1 color=#ff0000>" & date() & " - Usuário: " & session("func") & "</font></i>" 
   %>
  </td>
 </tr>
 <tr bgcolor=#8B8682 height=15>
  <td align=center>ABRIR</td>
 </tr>
 <tr>
  <td align=center>
   <font face=arial size=2>




  <font face=Arial size=3>Acesso restrito!<br><br>


  </font>
 <a href=javascript:history.go(-1)>Voltar</a>


   </font>
  </td>
 </tr>
 <tr>
  <td valign=bottom>
   <!--#include file=botoes.asp-->
  </td>
 </tr>
</table>
</form>
</body>
</html>
