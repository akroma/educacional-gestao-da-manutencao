<%
set conexao = server.createobject("adodb.connection")
conexao.open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("oscadastroteste.mdb") '& ";uid=sa;pwd=;"
set bdcarg = conexao.execute("select codcarg from func where func = '" & session("func") & "'")
cargo= bdcarg("codcarg")
set bdcarg = nothing
conexao.close
set conexao = nothing
%>
<html>
<head><title>CFP 1.16 - Ordem de Serviço</title></head>
<body bgcolor=#F0FFF0>
<table bgcolor=#C1CDC1 width=510 height=370 align=center cellspacing=0 cellpadding=0>
 <tr height=15>
  <td align=center>
   <% response.write "<i><font face=verdana, arial size=1 color=#ff0000>" & date() & " - Usuário: " & session("func") & "</font></i>" %>
  </td>
 </tr>
 <tr bgcolor=#8B8682 height=15>
  <td><center>MANUTENÇÃO DO SISTEMA<center></td>
 <tr>
  <td>
   <table width=100% height=235>
    <tr>
     <td align=center>
      <font face=verdana, arial size=1>
   <%  if cargo = "1" then 'Administrador
        response.write "<a href=funcionariocadastro.asp>CADASTRO DE FUNCIONÁRIOS</a><br><br>"
        response.write "<a href=funcionarioalterar.asp>ALTERAR FUNCIONÁRIOS</a><br><br>"
        response.write "<a href=construcao.asp.asp>CADASTRO DE SITUAÇÕES</a><br><br>"
        response.write "<a href=construcao.asp>ALTERAR SITUAÇÕES</a><br><br>"
        response.write "<a href=tipooscadastro.asp>CADASTRO TIPO OS</a><br><br>"
        response.write "<a href=construcao.asp>ALTERAR TIPO OS</a><br><br>"
        response.write "<a href=construcao.asp>REPARAR BANCO DADOS</a><br><br>"
       else
        response.write "ACESSO PERMITIDO SOMENTE PARA ADMINISTRADOR"
       end if
%>
 </font>
     </td>
    </tr>
   </table>  
  </td>
 </tr>
 <tr>
  <td>
   <!--#include file=botoes.asp-->
  </td>
 </tr>
</table>
</body>
</html>
