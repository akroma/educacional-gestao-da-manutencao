<html>
<head><title>CFP 1.16 - Ordem de Serviço</title></head>
<body bgcolor=#F0FFF0>
    <form method=post action=funcionariocadastro.asp>
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
  <td align=center>CADASTRO DE FUNCIONÁRIOS</td>
 </tr>
 <tr>
  <td align=center>
   <font face=arial size=2>
<%   
   set conexao = server.createobject("adodb.connection")
   conexao.open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("oscadastroteste.mdb") & ";uid=sa;pwd=;"
   
  if request.form("niffuncfr") = "" Then
  
 %> 

     <table width=100% border=1 cellpadding=2 cellspacing=0 align=center>
      <tr>
       <td width=20%>NIF:</td>  
       <td width=30%><input type=text name=niffuncfr size=15 maxlength=15></td>
      </tr>
      <tr>       
       <td>Nome:</font></td>
       <td><input type=text name=funcfr size=50 maxlength=50></td>
      </tr>
      <tr>       
       <td>Email:</font></td>
       <td><input type=text name=emailfr size=50 maxlength=50></td>
      </tr>
      <tr>
       <td>Ramal:</font></td>
       <td><input type=text name=ramalfr size=5 maxlength=15></td>
      </tr>
      <tr>
       <td>Núcleo:</td>
       <td>
        <select name=nucleofr>
<%   
    set bdnucleo = conexao.execute("select * from nucl")
    while not bdnucleo.eof
     response.write "<option value=" 
     response.write bdnucleo("codnucl") 
     response.write ">"
     response.write bdnucleo("nucl") 
     response.write "</option>"
     response.write "<br>"
     bdnucleo.movenext
    wend
    set bdnucleo = Nothing
%>   
    </select>
       </td>
      </tr>   
      <tr>
       <td>Tipo:</td>
       <td>
        <select name=tipofr>
         <option selected value=5>SOLICITANTE</option>  
<%
Set bdcargo = conexao.execute("select * from carg")
while not bdcargo.eof
  response.write "<option value=" 
  response.write bdcargo("codcarg") 
  response.write ">"
  response.write bdcargo("carg") 
  response.write "</option>"
  response.write "<br>"
  bdcargo.movenext
wend
set bdcargo = Nothing
%>

        </select>
       </td>
 
      </tr>
     </table>
     <input type=hidden name=voltar value=nao>


<%
   else
    if request.form("niffuncfr") <> "" then
     if request.form("funcfr") <> "" and request.form("ramalfr") <> "" and request.form("nucleofr") <> "" and request.form("tipofr") <> "" and request.form("emailfr") <> ""  then
     buscafunc = "select codfunc from func where codfunc = '" & request.form("niffuncfr") & "' "	
     set bd = conexao.execute(buscafunc)
     if not bd.eof then	'encontrou o registro procurado %>
      <center><br><br>
       <font face=Arial size=2><br>Já existe um funcionário cadastrado com este NIF!<br>
        <a href=javascript:history.go(-1)>Voltar</a>&nbsp;&nbsp;
       </font>
      </center>
<%   else
      comandosql = "insert into func (codfunc, func, codnucl, codcarg, ramal, email) VALUES ('"
      comandosql = comandosql & request.form("niffuncfr") & "','" & request.form("funcfr") & "'," & request.form("nucleofr") & "," & request.form("tipofr") & "," & request.form("ramalfr") & ",'" & request.form("emailfr") & "')"
      conexao.execute(comandosql)  %>
      <center><br><br>
       <font face=Arial size=2>Funcionário Cadastrado com sucesso!<br><br> 
        <a href=funcionariocadastro.asp><b>Cadastrar outro funcionário</b></a>
       </font>
      </center>
<%   end if
     else
      response.write "Os campos NIF, Nome e Email não podem ser em branco<br><a href=javascript:history.go(-1)>Voltar</a>"
    end if
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
