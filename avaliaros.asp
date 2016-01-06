<html>
<head><title>CFP 1.16 - Ordem de Serviço</title></head>
<body bgcolor=#F0FFF0>
   <form method=post action=avaliaros.asp> 
<table bgcolor=#C1CDC1 width=510 height=370 align=center cellspacing=0 cellpadding=0>
 <tr height=15>
  <td align=center> 
<% session.LCID = 1046
   response.expires = -1000 
   response.write " <i><font face=verdana, arial size=1 color=#ff0000>" & date() & " - Usuário: " & session("func") & "</font></i>" 
%></td>
 </tr>
 <tr bgcolor=#8B8682 height=15>
  <td align=center>AVALIAR</td>
 </tr>
 <tr>
  <td>
    <table width=100% align=center border=0> 
     <tr>
      <td align=center valign=center>
       <font face=arial size=2>
<%      set conexao = server.createobject("adodb.connection")
        conexao.open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("oscadastroteste.mdb") & ";uid=sa;pwd=;"
        if request.form("fecharosfr1") = "" then

'         set bdcarg = conexao.execute("select codfunc, codcarg from func where func = '" & session("func") & "'")
'         cargo = bdcarg("codcarg")
'         codexec = bdcarg("codfunc")
'         set bdcarg = nothing

         set bdavalos = conexao.execute("select nros, local, codfunc, servso, obs, executor from os where codsitua = 14 and codfunc ='" & session("ni") & "' order by nros")

         if not bdavalos.eof then
          response.write "<select name=nrosfr>"
          while not bdavalos.eof
           set bdexec = conexao.execute("select func from func where codfunc = '" & bdavalos("executor") & "'")
           executor = bdexec("func")
           set bdexec = nothing
           response.write "<option value=" & bdavalos("nros") & ">" 
           response.write bdavalos("nros") & " - " & executor & "</option>" & "<br>"
           bdavalos.movenext
          wend
          set bdavalos = nothing
          response.write "</select><input type=hidden name=fecharosfr1 value=sim>"
         else
          response.write "Não existe OS para avaliar" 
          set bdavalos = nothing 
         end if


        else
         if request.form("fecharosform2") = "" then
          set bdavalos = conexao.execute("select * from oscon where nros =" & request.form("nrosfr"))
          set bdexec = conexao.execute("select func from func where codfunc = '" & bdavalos("executor") & "'")
          session("executor") = bdexec("func")
          set bdexec = nothing%>
          <table width=100% align=center border=1 bordercolor=#8B8682 cellpadding=1 cellspacing=0>
           <tr>
            <td width=30%>OS:</td>
            <td><font face=Arial size=2><%=bdavalos("nros")%></font></td>
           </tr>
           <tr>
            <td width=33%>NI:</td>
            <td>
             <font face=Arial size=2>
<%            if bdavalos("nipatr") <> 0 then 
               response.write bdavalos("nipatr")
              else
               response.write "Não possui"
              end if %> 
             </font>
            </td>
           </tr>
           <tr>
            <td width=33%>Local:</td>
             <td><font face=Arial size=2><%=bdavalos("local")%></font></td>
           </tr>
<%         if session("carg") = "EXECUTOR" then %>
            <tr>
             <td>Solicitante:</td>
             <td><font face=Arial size=2><%=bdavalos("func")%></font></td>
            </tr>
<% else %>

           <tr>
             <td>Executor:</td>
             <td><font face=Arial size=2><% =session("executor") %></font></td>
            </tr>



<%         end if %>
           <tr>
            <td valign=top>Dados do Pedido:</td>
            <td><font face=Arial size=2><%=bdavalos("servso")%></font></td>
           </tr>
         <tr>
          <td colspan=4><font face=Arial size=2><b><center>ANDAMENTO</center></td>
         </tr>


<%    
         set bdandam = conexao.execute("select * from andam where nros =" & bdavalos("nros") & " order by nrandam")
         if not bdandam.eof then
          while not bdandam.eof 
          response.write  "<tr><td><font face=Arial size=1>" & bdandam("dtandam") & "</font></td>"
           set bdexec = conexao.execute("select func from func where codfunc = '" & bdandam("codfunc") & "'")
           solic = bdexec("func")
          response.write "<td><font face=Arial size=1>" & bdexec("func") & "</font>"
           set bdexec = nothing    
          response.write "</td></tr>"
          response.write  "<tr><td colspan = 2><font face=Arial size=1>" & bdandam("andam") & "</font></td></tr>"          
           bdandam.movenext
          wend 
         else
          response.write "<tr><td colspan = 2><center>Sem andamento</center></td></tr>" 
         end if    %>
           <tr>
            <td colspan=2>Avaliação:</td>
           </tr>
           <tr> 
            <td colspan=2 align=center>
             <input type=radio name=avalfr value=1><font face=verdana, arial color=#ff0000>1</font>
             <input type=radio name=avalfr value=2><font face=verdana, arial color=#ff0000>2</font>
             <input type=radio name=avalfr value=3>3
             <input type=radio name=avalfr value=4>4
             <input type=radio name=avalfr value=5>5</td>
           </tr>
           <tr>
            <td colspan=2>Observação:</td>
           </tr>
           <tr> 
            <td colspan=2><textarea align=center name=obsfr rows=2 cols=60></textarea></td>
           </tr>
          </table>
          <input type=hidden name=nrosfr value=<%=bdavalos("nros")%>>
          <input type=hidden name=execfr value=<%=bdavalos("executor")%>>
          <input type=hidden name=fecharosfr1 value=sim>
          <input type=hidden name=fecharosform2 value=sim>
     <input type=hidden name=voltar value=nao>
<%      set bdavalos = nothing  
         else 
          if request.form("avalfr") = "" then
           response.write "Por favor informe a sua avaliação<br><a href=javascript:history.go(-1)>voltar</a>"
          else
            if request.form("avalfr") = 3 or request.form("avalfr") > 3  then
             set bdavalos = conexao.execute("update os set aval1 = " & request.form("avalfr") & ", codsitua = 2 where nros = " & request.form("nrosfr") & ";")

             if request.form("obsfr") <> "" then
              set bdandamos = conexao.execute("insert into andam (nros, andam, codfunc) values (" & request.form("nrosfr") & ",'" & request.form("obsfr") & "','" & session("ni") & "')")
             end if


             response.write "A OS " & request.form("nrosfr") & " foi avaliada satisfatóriamente e fechada com sucesso!<br><a href=avaliaros.asp>Avaliar outra OS</a>"

             ' ENVIA UM EMAIL PARA O EXECUTOR

             set buscaemail = conexao.execute("select email from func where func ='" & session("executor") & "'")

             assunto = "A OS " & request.form("nrosfr") & " foi avaliada satisfatóriamente por " & session("func")

             set correio = Server.CreateObject("CDONTS.NewMail") 

             correio.send "suporteinf116@sp.senai.br (Sistema de OS)", buscaemail("email"), assunto, "Este correio foi enviado automaticamente pelo programa de OS, por favor não responder.", 2

             set correio = nothing
             set buscaemail = nothing



            else

              if request.form("obsfr") = "" then

              response.write "O campo observação não pode ficar em branco quando a OS foi reprovada<br><a href=javascript:history.go(-1)>voltar</a>"
              else              

             set bdavalos = conexao.execute("select repr from os where nros = " & request.form("nrosfr") & ";") 

             if bdavalos("repr") = 1 then
              set bdavalos = conexao.execute("update os set codsitua = 3 where nros = " & request.form("nrosfr") & ";")
              set bdandamos = conexao.execute("insert into andam (nros, andam, codfunc) values (" & request.form("nrosfr") & ",'" & request.form("obsfr") & "','" & session("ni") & "')")
              response.write "A OS serviço " & request.form("nrosfr") & " foi avaliada insatistatóriamente duas vezes e será encaminhada para o responsável da manutenção!<br><a href=avaliaros.asp>Avaliar outra OS</a>"


              ' ENVIA UM EMAIL PARA O RESPONSAVEL DA MANUTENCAO

               set buscaemail = conexao.execute("select email from func where codsitua = 7 and codcarg = 1")

               assunto = "A OS " & request.form("nrosfr") & " foi reprovada 2 vezes pelo solicitante " & session("func") & " - Executor: " & session("executor")

               while not buscaemail.eof 
                set correio = Server.CreateObject("CDONTS.NewMail") 
                correio.send "suporteinf116@sp.senai.br (Sistema de OS)", buscaemail("email"), assunto, "Este correio foi enviado automaticamente pelo programa de OS, por favor não responder.", 2
               buscaemail.movenext
              wend
 
              set correio = nothing
              set buscaemail = nothing

             else

              set bdavalos = conexao.execute("update os set codsitua = 3 where nros = " & request.form("nrosfr") & ";")

              set bdavalos = conexao.execute("update os set repr = 1, codsitua = 1 where nros = " & request.form("nrosfr") & ";")

              set bdandamos = conexao.execute("insert into andam (nros, andam, codfunc) values (" & request.form("nrosfr") & ",'" & request.form("obsfr") & "','" & session("ni") & "')")

              response.write "A OS " & request.form("nrosfr") & " foi avaliada insatisfatóriamente e foi devolvida para o executor com sucesso!<br><a href=avaliaros.asp>Avaliar outra OS</a>"             

              assunto = "A OS " & request.form("nrosfr") & " foi avaliada insatisfatóriamente por " & session("func")

              set buscaemail = conexao.execute("select email from func where codfunc = '" & request.form("execfr") & "'")

              set correio = Server.CreateObject("CDONTS.NewMail") 

              correio.send "suporteinf116@sp.senai.br (Sistema de OS)", buscaemail("email"), assunto, "Este correio foi enviado automaticamente pelo programa de OS, por favor não responder.", 2

              set correio = nothing

              set buscaemail = nothing

            end if
            end if  
            end if 
           set bdavalos = nothing
          end if
         end if
        end if
        conexao.close
        set conexao = nothing
%>    </font>
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
