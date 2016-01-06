<html>
<head><title>CFP 1.16 - Ordem de serviço</title></head>
<body bgcolor=#F0FFF0>
<form method=post name=consultamicro action=funcionarioalterar.asp>
<table bgcolor=#C1CDC1 width=510 height=370 align=center cellspacing=0 cellpadding=0>
 <tr height=15>
  <td align=center>
<% Session.LCID = 1046 
   response.expires = -1000 
   response.write " <i><font face=verdana, arial size=1 color=#ff0000>" & date() & " - Usuário: " & session("FUNC") & "</font></i>" 
%></td>
 </tr>
 <tr bgcolor=#8B8682 height=15>
  <td align=center>ALTERAR FUNCIONÁRIO</td>
 </tr>
 <tr>
  <td align=center>
<% if request.form("opcaoform") = "" then %>  

     <input type=text name=procurafr>
     <select name=opcaoform>
      <option value=codfunc>NI DO FUNCIONÁRIO</option>
      <option value=func>NOME DO FUNCIONÁRIO</option>
     </select>

 
<% else
    set conexao = server.createobject("adodb.connection")
    conexao.open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("oscadastroteste.mdb") & ";uid=sa;pwd=;"
    if request.form("buscafunc") = "" then 
     set bdfunc = conexao.execute("select * from funccon where " & request.form("opcaoform") & " = '" & Request.Form("procurafr") & "'")
     if not bdfunc.eof then  'encontrou o registro procurado
      func = bdfunc("func") %>

      <input type=hidden name=buscafunc value=sim>
      <input type=hidden name=opcaoform value=sim>
      <table width=100% border=1 cellpadding=2 cellspacing=0 align=center>
       <tr>
        <td>NIF:</td>  
        <td><%=bdfunc("codfunc")%></td>
       </tr>
       <tr>       
        <td>Nome atual:</font></td>
        <td><%=bdfunc("func")%></td>
       </tr>
       <tr>       
        <td>Alterar nome:</font></td>
        <td><input type=text name=funcfr size=50 maxlength=50></td>
       </tr>
       <tr>       
        <td>Alterar email:</font></td>
        <td><input type=text name=emailfr size=50 maxlength=50 value=<%=bdfunc("email")%>></td>
       </tr>
       <tr>
        <td>Ramal:</font></td>
        <td><input type=text name=ramalfr value=<%=bdfunc("ramal")%> size=4 maxlength=10></td>
       </tr>
       <tr>
        <td>Núcleo:</td>
        <td>
         <select name=nucleofr>
          <option selected value=<%=bdfunc("codnucl")%>><%=bdfunc("nucl")%></option>
<%         set bdnucleo = conexao.execute("select * from nucl")
           while not bdnucleo.eof
            response.write "<option value=" & bdnucleo("codnucl") & ">" & bdnucleo("nucl") & "</option><br>"
           bdnucleo.movenext
           wend
           set bdnucleo = Nothing
%>       </select>
        </td>
       </tr>   
       <tr>
        <td>Tipo:</td>
        <td>
         <select name=tipofr>
          <option selected value=<%=bdfunc("codcarg")%>><%=bdfunc("carg")%></option>
<%         Set bdcargo = conexao.execute("select * from carg")
           while not bdcargo.eof
            response.write "<option value=" & bdcargo("codcarg") & ">" & bdcargo("carg") & "</option><br>"
           bdcargo.movenext
           wend
           set bdcargo = Nothing
%>       </select>
        </td>
       </tr>
       <tr>
        <td>Tipo:</td>
        <td>
         <select name=situafr>
          <option selected value=<%=bdfunc("codsitua")%>><%=bdfunc("situa")%></option>
          <option value=9>AFASTADO</option>
          <option value=8>DESLIGADO</option>
          <option value=10>LICENÇA</option>
          <option value=7>TRABALHANDO</option>
          <option value=11>TRANSFERIDO</option>
%>       </select>
        </td>
       </tr>
       <tr>
        <td>Resetar senha</td>
        <td>
         <input type=checkbox name=resetsenhafr>Marque aqui para resetar a senha<br>
        </td>
       </tr>
      </table>

     <input type=hidden name=voltar value=nao>

     <input type=hidden name=niffuncfr value=<%=bdfunc("codfunc")%>>

 <%   Set bdfunc = Nothing
     else 
%>    <br><br><br>
      <font face=Arial size=3 color=#ff0000><b>O funcionário não foi encontrado!</b><br></font><br>
      <font face=Arial size=2 color=#ff0000> 
       <a href=funcionarioalterar.asp>Alterar outro funcionário</a>
      </font>
<%   end if
    else
     if request.Form("emailfr") <> "" then       
     if request.Form("funcfr") <> "" then
      conexao.execute("update func set func = '" & request.form("funcfr") & "', email = '" & request.form("emailfr") & "', ramal = " & request.Form("ramalfr") & ", codnucl =" & request.Form("nucleofr") & ", codcarg = " & request.Form("tipofr") & ", codsitua = " & request.Form("situafr") & " where codfunc = '" & request.form("niffuncfr") & "';")  
     else
      conexao.execute("update func set email = '" & request.form("emailfr") & "', ramal = " & request.Form("ramalfr") & ", codnucl = " & request.Form("nucleofr") & ", codcarg = " & request.Form("tipofr") & ", codsitua = " & request.Form("situafr") & " where codfunc = '" & request.form("niffuncfr") & "';")   
     end if

     if request.Form("resetsenhafr") <> "" then
      conexao.execute("update func set sen = 'senaisp' where codfunc = '" & request.form("niffuncfr") & "';")         
     end if         
response.write "<font face=Arial size=2>Dados alterados com sucesso!<br><a href=funcionarioalterar.asp>Alterar outro funcionário</a></font>"
     else
       response.write "O campo email não pode ser em branco<br><a href=javascript:history.go(-1)>Voltar</a>" 

     end if
  end if
     conexao.close
     set conexao = nothing
    end if
%></td>
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