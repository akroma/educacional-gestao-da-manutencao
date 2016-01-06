<html>
<head><title>CFP 1.16 - Ordem de Serviço</title></head>
<body bgcolor=#F0FFF0>
    <form method=post action=tipooscadastro.asp>
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
  <td align=center>CADASTRO DE TIPO DE OS</td>
 </tr>
 <tr>
  <td align=center>
   <font face=arial size=2>
<%   
   set conexao = server.createobject("adodb.connection")
   conexao.open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("oscadastroteste.mdb") & ";uid=sa;pwd=;"
   
  if request.form("ostipofr") = "" Then
  
 %> 
     <table width=100% border=0 cellpadding=2 cellspacing=0 align=center>
      <tr>       
       <td>Tipo:</font></td>
       <td><input type=text name=ostipofr size=50 maxlength=50></td>
      </tr>
     </table>
     <input type=hidden name=voltar value=nao>


<%
   else
    if request.form("ostipofr") <> "" then
     buscaostipo = "select tipo from ostipo where tipo = '" & request.form("ostipofr") & "' "	
     set bd = conexao.execute(buscaostipo)
     if not bd.eof then	'encontrou o registro procurado %>
      <center><br><br>
       <font face=Arial size=2><br>O tipo de OS já está cadastrado!<br>
        <a href=tipoosalterar>Alterar os dados deste tipo de OS</a>&nbsp;&nbsp;<br>
        <a href=javascript:history.go(-1)>Voltar</a>&nbsp;&nbsp;
       </font>
      </center>
<%   else
      conexao.execute("insert into ostipo (tipo) VALUES ('" & request.form("ostipofr") & "')")  %>
      <center><br><br>
       <font face=Arial size=2>Tipo de OS cadastrado com sucesso!<br><br> 
        <a href=tipooscadastro.asp><b>Cadastrar outro tipo</b></a>
       </font>
      </center>
<%   end if
     set bd = nothing
     
    else %>
     <center><br><br>
      <font face=Arial size=2>Atenção: Alguns dados necessário não foram preenchidos corretamente<br>
       <a href=javascript:history.go(-1)>Voltar</a>&nbsp;&nbsp;
      </font>
     </center> 
<% end if
  end if
 conexao.close
    set conexao = nothing
   


%></td>
 <tr>
  <td valign=bottom>
   <!--#include file=botoes.asp-->
  </td>
 </tr>
</table>
</form> 
</body>
</html>
