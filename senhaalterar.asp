<html>
<head><title>CFP 1.16 - Ordem de Serviço</title></head>
<body bgcolor=#F0FFF0>
<form method=post action=senhaalterar.asp target=_self>
<table bgcolor=#C1CDC1 width=510 height=370 align=center cellspacing=0 cellpadding=0>
 <tr height=15>
  <td align=center>
<%
   response.expires = -1000
   Session.LCID = 1046
   response.write " <i><font face=verdana, arial size=1 color=#ff0000>" & date() & " - Usuário: " & session("func") & "</font></i>" 
%>
  </td>
 </tr>
 <tr bgcolor=#8B8682 height=15>
  <td align=center>ALTERAR DADOS</td>
 </tr>
 <tr>
  <td align=center>
   <font face=arial size=2>
<%  set conexao = server.createobject("adodb.connection")
    conexao.open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("oscadastroteste.mdb") & ";uid=sa;pwd=;"
    if request.form("confsen") = "" then
     set bdnucleo = conexao.execute("SELECT FUNC.func, NUCL.nucl, FUNC.ramal, FUNC.email FROM FUNC INNER JOIN NUCL ON FUNC.codnucl = NUCL.codnucl where func = '" & session("func") & "'")
     nucleo = bdnucleo("nucl")
     ramal = bdnucleo("ramal")
     email = bdnucleo("email")
     set bdnucleo = nothing
%>    <table align=center border=1 bordercolor=#8B8682 cellpadding=2 cellspacing=0>
       <tr>
        <td width=35%>Solicitante:</td>
        <td><font face=Arial size=2><%=session("func")%></font></td>
       </tr>
       <tr>
        <td>Núcleo:</td>
        <td><font face=Arial size=2><%=nucleo %></font></td>
       </tr>
       <tr>
        <td>Ramal:</td>
        <td><input type=text name=ramalfr value=<%=ramal%> size=4 maxlength=4></td>
       </tr>
       <tr>
        <td>Email:</td>
        <td><input type=text name=emailfr value=<%=email%> size=50 maxlength=50></td>
       </tr>
       <tr>
        <td>Nova senha:</td>
        <td align=center><input type=password name=senfr size=30 maxlength=30 align=left></td>
       </tr>
       <tr>
        <td>Repita a Nova senha:</td>
        <td align=center><input type=password name=repsenfr size=30 maxlength=30 align=left></td>
       </tr>
      </table>
      <input type=hidden name=confsen value=sim>
      <input type=hidden name=voltar value=nao>
<%  else

    
     if request.form("emailfr") = "" then 

           response.write "O campo email não pode ser em branco<br><a href=javascript:history.go(-1)>Voltar</a>"     
     else
     if request.form("senfr") = "" or request.form("repsenfr") = "" then 
      response.write "A senha não pode ser em branco<br><a href=senhaalterar.asp><b>Voltar</b></a>"
     else
      if request.form("ramalfr") = "" then  
       ramal = 0
      else
       ramal = request.form("ramalfr")
      end if
      if isnumeric(ramal) = false then
       response.write "O campo ramal deve conter apenas números<br><a href=senhaalterar.asp><b>Voltar</b></a>"
      else
       if cstr(request.form("senfr")) <> cstr(request.form("repsenfr")) then 
        response.write "A senha não confere<br><a href=senhaalterar.asp><b>Voltar</b></a>"
       else
        set bdsolic = conexao.execute("select codfunc from func where func = '" & session("func") & "'")
        solicitante = bdsolic("codfunc")
        set bdsolic = nothing 
        conexao.execute("update func set sen ='" & request.form("senfr") & "', ramal = " & ramal & ", email ='" & request.form("emailfr") & "' where codfunc ='" & solicitante & "'" )
        response.write "Dados atualizados com sucesso!"
       end if
      end if
     end if
     end if
    end if 
    conexao.close
    set conexao = Nothing 
%> </font>
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
