<html>
<head><title>CFP 1.16 - Ordem de Servi�o</title></head>
<body bgcolor=#F0FFF0>
   <form method=post action=redirecionaros.asp target=_self>
<table bgcolor=#C1CDC1 width=510 height=370 align=center cellspacing=0 cellpadding=0>
 <tr height=15>
  <td align=center>
<% session.LCID = 1046
   response.expires = -1000
   response.write " <i><font face=verdana, arial size=1 color=#ff0000>" & date() & " - Usu�rio: " & session("func") & "</font></i>" 
   set conexao = server.createobject("adodb.connection")
   conexao.open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("oscadastroteste.mdb") & ";uid=sa;pwd=;"  
   if session("cargo") = "" or session("cargo") = 2 or session("cargo") = 5 or session("cargo") = 4 then 
    response.redirect("acessorestrito.asp")
   end if
%></td>
 </tr>
 <tr bgcolor=#8B8682 height=15>
  <td align=center>REDIRECIONAR</td>
 </tr>
 <tr>
  <td>

   <table align=center border=0 cellspacing=0 cellpadding=0>
    <tr>
     <td align=center>
      <font face=arial size=2>
<%     if request.form("diros") = "" then
        set bdosaberta = conexao.execute("select nros, codfunc from oscon where codsitua = 1 order by nros desc")
        if not bdosaberta.eof then    'encontrou o registro procurado
         response.write "<select name=nrosfr>"

         while not bdosaberta.eof
          response.write "<option value=" & bdosaberta("nros") & ">" 
          set bdsolic = conexao.execute("select func from func where codfunc = '" & bdosaberta("codfunc") & "'")
          solicitante = bdsolic("func")
          response.write bdosaberta("nros") & " - " & bdsolic("func") & "</option>" & "<br>"
          set bdsolic = nothing
          bdosaberta.movenext
         wend
         response.write "</select><input type=hidden name=diros value=sim>"
        else
         response.write "N�o existe OS para direcionar"
        end if
        set bdosaberta = nothing ' Tem que ser aqui sen�o apaga os campos
        conexao.close
        set conexao = nothing
       else
        if request.form("nrosfr2") = "" then
'         set conexao = server.createobject("adodb.connection")
'         conexao.open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("oscadastro.mdb") & ";uid=sa;pwd=;"
          set bddiros = conexao.execute("select nros, nipatr, local, codfunc, executor, servso, dtabe, dtapr from oscon where nros = " & request.form("nrosfr"))
 %>
          <table width=100% border=1 bordercolor=#8B8682 cellpadding=2 cellspacing=0>
           <tr>
            <td width=25%>Ordem Servi�o</td>
            <td><%=bddiros("nros")%></td>
           </tr>
<%         if bddiros("nipatr") <> 0 then %>
           <tr>
            <td width=33%><font face=Arial size=2>NI:</font></td>
             <td><font face=Arial size=2><%=bddiros("nipatr")%></font></td>
           </tr>
<%         end if %>
           <tr>
            <td width=33%>Local:</td>
            <td><%=bddiros("local")%></td>
           </tr>
           <tr>
            <td>Solicitante:</td>
            <td>
             <%
             set bdsolic = conexao.execute("select func, email from func where codfunc = '" & bddiros("codfunc") & "'")
             response.write bdsolic("func") %>
             <input type=hidden name=solicemailfr value=<%=bdsolic("email")%>>
            </td>
           </tr>
           <tr>
            <td>Executor:</td>
            <td>
             <%
             set bdsolic = nothing
             set bdexecatual = conexao.execute("select func from func where codfunc = '" & bddiros("executor") & "'")
             response.write bdexecatual("func")
             set bdexecatual = nothing %> 
            </td>
           </tr>
           <tr>
            <td valign=top>Dados Pedido:</td>
            <td><%=bddiros("servso")%></td>
           </tr>
           <tr>
            <td>Data abertura:</td>
            <td><%=bddiros("dtabe")%></td>
           </tr>
            <tr>
             <td>Data aprova��o:</td>
             <td><%=bddiros("dtapr")%></td>
            </tr>

           <tr>
            <td>Tipo manuten��o:</td>
            <td>
             <select name=codostipofr>
              <option selected value=""></option>
<%            set bdostipo = conexao.execute("select codostipo, tipo from ostipo order by tipo")
              while not bdostipo.eof
               response.write "<option value=" & bdostipo("codostipo") & ">"
               response.write bdostipo("tipo") & "</option><br>"
               bdostipo.movenext
              wend
              set bdostipo = nothing
%>           </select>
            </td>
           </tr>
           <tr>
            <td>Novo Executor:</td>
            <td>
             <select name=execfr>
              <option selected value="<%=bddiros("executor")%>><%=bddiros("executor")%></option>
<%            set bdexec = conexao.execute("select codfunc, func from func where codsitua = 7 and codnucl = 8 order by func")
              while not bdexec.eof
               response.write "<option value=" & bdexec("codfunc") & ">"
               response.write bdexec("func") & "</option><br>"
               bdexec.movenext
              wend
              set bdexec = nothing
%>           </select>
            </td>
           </tr>
          </table>
          <input type=hidden name=nrosfr2 value=<%=bddiros("nros")%>>
          <input type=hidden name=diros value=sim>
     <input type=hidden name=voltar value=nao>

<%       set bddiros = nothing
         conexao.close
        else 
         if request.form("codostipofr") = "" then
          response.write "Informe o tipo de manuten��o<br><br><a href=javascript:history.go(-1)>Voltar</a>"
          else
         if request.form("execfr") = "1" then
          response.write "Informe o nome do executor<br><br><a href=javascript:history.go(-1)>Voltar</a>"
         else
          set bddiros = conexao.execute("update os set codostipo = " & request.form("codostipofr") & ", executor = '" & request.form("execfr") & "', dtdir = now(), codsitua = 1 where nros = " & request.form("nrosfr2"))
set buscaemail = conexao.execute("select func, email from func where codfunc ='" & request.form("execfr") & "'")

assunto = "A OS " & request.form("nrosfr2") & " foi redirecionada para " & buscaemail("func")


set correio = Server.CreateObject("CDONTS.NewMail") 

correio.send "suporteinf116@sp.senai.br (Sistema de OS)", buscaemail("email"), assunto, "Este correio foi enviado automaticamente pelo programa de OS, por favor n�o responder.", 2

set correio = nothing

set correio = Server.CreateObject("CDONTS.NewMail") 

correio.send "suporteinf116@sp.senai.br (Sistema de OS)", request.form("solicemailfr"), assunto, "Este correio foi enviado automaticamente pelo programa de OS, por favor n�o responder.", 2


set correio = nothing
set buscaemail = nothing

         response.write "A OS " & request.form("nrosfr2") & " foi redirecionada com sucesso!<br><br><a href=redirecionaros.asp>Redirecionar outra OS</a>"
          set bddiros = nothing
          conexao.close
         end if
         end if
        end if 
       end if %>
      </font>
     </td>
    </tr>
   </table>
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
