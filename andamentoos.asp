<html>
<head><title>CFP 1.16 - Ordem de Serviço</title></head>
<body bgcolor=#F0FFF0>
   <form method=post action=andamentoos.asp> 
<table bgcolor=#C1CDC1 width=510 height=370 align=center cellspacing=0 cellpadding=0>
 <tr height=15>
  <td align=center> 
<% session.LCID = 1046
   response.expires = -1000 
   response.write " <i><font face=verdana, arial size=1 color=#ff0000>" & date() & " - Usuário: " & session("func") & "</font></i>" 
%></td>
 </tr>
 <tr bgcolor=#8B8682 height=15>
  <td align=center>ANDAMENTO</td>
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
         
          
         if session("cargo") = 4 or session("cargo") = 5 then
          set bdandamos = conexao.execute("select nros, local, codfunc, servso, obs from oscon where codfunc ='" & session("ni") & "' and codsitua <> 2 or executor ='" & session("ni") & "' and codsitua <> 2 order by nros")
         else  
          set bdandamos = conexao.execute("select nros, local, codfunc, servso, obs from oscon where codsitua <> 2 order by nros")       
         end if
         if not bdandamos.eof then
          response.write "<select name=nrosfr>"
          while not bdandamos.eof
           set bdsolic = conexao.execute("select codfunc, func from func where codfunc = '" & bdandamos("codfunc") & "'")
           solic = bdsolic("func")
           set bdsolic = nothing
           response.write "<option value=" & bdandamos("nros") & ">" 
           response.write bdandamos("nros") & " - " & solic & "</option>" & "<br>"
           bdandamos.movenext
          wend
          set bdandamos = nothing
          response.write "</select><br><br><input type=hidden name=fecharosfr1 value=sim>"
         else
          response.write "Não existe OS aberta" 
          set bdandamos = nothing 
         end if


        else
         if request.form("fecharosform2") = "" then
          set bdandamos = conexao.execute("select * from oscon where nros =" & request.form("nrosfr"))%>
          <table width=100% align=center border=1 bordercolor=#8B8682 cellpadding=1 cellspacing=0>
           <tr>
            <td width=30%>OS:</td>
            <td><font face=Arial size=2><%=bdandamos("nros")%></font></td>
           </tr>
<%         if bdandamos("nipatr") <> 0 then %>
            <tr>
             <td width=33%>NI:</td>
             <td>
              <font face=Arial size=2>
               <%=bdandamos("nipatr")%>
              </font>
             </td>
            </tr>
<%         end if %> 
           <tr>
            <td width=33%>Local:</td>
             <td><font face=Arial size=2><%=bdandamos("local")%></font></td>
           </tr>
<%        set bdsitua = conexao.execute("select situa from situa where codsitua = " & bdandamos("codsitua")) %>
          <td><font face=Arial size=2><b>Situação</font></td>
          <td><font face=Arial size=2><%=bdsitua("situa")%></font></td></tr>
   <%    set bdsitua = Nothing %>

<%         if session("cargo") = "4" then %>
            <tr>
             <td>Solicitante:</td>
             <td><font face=Arial size=2><%=bdandamos("func")%></font></td>
                                  
<%         else
            if session("cargo") = "5" then 
            if bdandamos("executor") <> "" then
              
             set bdexec = conexao.execute("select func from func where codfunc = '" & bdandamos("executor") & "'")  
%>          
             <tr>
              <td>Executor:</td>
              <td><font face=Arial size=2><%=bdexec("func")%></font></td>
                        
<%           end if
            else
             set bdexec = conexao.execute("select func from func where codfunc = '" & bdandamos("executor") & "'") %>
              <tr>
               <td>Solicitante:</td>
               <td><font face=Arial size=2><%=bdandamos("func")%></font></td>
<%          if bdandamos("executor") <> "" then %>
              <tr>
               <td>Executor:</td>
               <td><font face=Arial size=2><%=bdexec("func")%></font></td>
<%          end if  
            end if
           end if %>

           <tr>
            <td colspan=4>Dados do Pedido:</td>
           </tr>  
           <tr>
            <td colspan=2><font face=Arial size=2><%=bdandamos("servso")%></font></td>
           </tr>
         <tr>
          <td colspan=2>Andamento</td>
         </tr>


<%       
         set bdandam = conexao.execute("select * from andam where nros =" & bdandamos("nros") & " order by nrandam")
         if not bdandam.eof then
          while not bdandam.eof 
           response.write  "<tr><td><font face=Arial size=2>" & bdandam("dtandam") & "</font></td>"
           set bdsolic = conexao.execute("select func from func where codfunc = '" & bdandam("codfunc") & "'")
           solic = bdsolic("func")
           response.write "<td><font face=Arial size=2>" & bdsolic("func") & "</font>"
           set bdsolic = nothing    
           response.write "</td></tr><tr><td colspan=2><font face=Arial size=2>" & bdandam("andam") & "</font></td></tr>"          
           bdandam.movenext
          wend 
         else
          response.write "<tr><td colspan=2><font face=Arial size=2>Sem registro de andamento</font></td></tr>" 
         end if    %>
           <tr>
            <td colspan=2>Registrar andamento:</td>
           </tr>
           <tr> 
            <td colspan=2><textarea align=center name=andamfr rows=2 cols=60></textarea></td>
           </tr>
          </table>
          <input type=hidden name=nrosfr value=<%=bdandamos("nros")%>>
          <input type=hidden name=execfr value=<%=bdandamos("executor")%>>
          <input type=hidden name=solicfr value=<%=bdandamos("codfunc")%>>
          <input type=hidden name=codsituafr value=<%=bdandamos("codsitua")%>>
          <input type=hidden name=fecharosfr1 value=sim>
          <input type=hidden name=fecharosform2 value=sim>
          <input type=hidden name=voltar value=nao>
<%      set bdandamos = nothing  
         else 
          if request.form("andamfr") = "" then
           response.write "Por favor registre um andamento<br><a href=javascript:history.go(-1)>voltar</a>"
          else
           set bdandamos = conexao.execute("insert into andam (nros, andam, codfunc) values (" & request.form("nrosfr") & ",'" & request.form("andamfr") & "','" & session("ni") & "')")
           response.write "Andamento registrado com sucesso!<br><a href=andamentoos.asp>Registrar outro andamento</a>"

'Enviar email para quem registrou o andamento

 assunto = "Foi registrado um andamento na OS " & request.form("nrosfr") & " por " & session("func")

if request.form("codsituafr") = 1 then
 if session("cargo") = 5 and request.form("execfr") <> "" then 'Busca email do executor
  set buscaemail = conexao.execute("select email from func where codfunc ='" & request.form("execfr") & "'")
 end if

 if session("cargo") = 4 then 'Busca email do solicitante
  set buscaemail = conexao.execute("select email from func where codfunc = '" & request.form("solicfr") & "'")
 end if

  set correio = Server.CreateObject("CDONTS.NewMail") 

  correio.send "suporteinf116@sp.senai.br (Sistema de OS)", buscaemail("email"), assunto, "Este correio foi enviado automaticamente pelo programa de OS, por favor não responder.", 2

else

if request.form("codsituafr") = 12 then ' Envia email para o coordenador


  set buscaemail = conexao.execute("select email from func where codsitua = 7 and codcarg = 2 and codnucl = " & session("nucleo"))

while not buscaemail.eof 

  set correio = Server.CreateObject("CDONTS.NewMail") 

  correio.send "suporteinf116@sp.senai.br (Sistema de OS)", buscaemail("email"), assunto, "Este correio foi enviado automaticamente pelo programa de OS, por favor não responder.", 2

  buscaemail.movenext
 wend

else ' Envia email para o administrador


  set buscaemail = conexao.execute("select email from func where codsitua = 7 and codcarg = 1")

while not buscaemail.eof 

  set correio = Server.CreateObject("CDONTS.NewMail") 

  correio.send "suporteinf116@sp.senai.br (Sistema de OS)", buscaemail("email"), assunto, "Este correio foi enviado automaticamente pelo programa de OS, por favor não responder.", 2

  buscaemail.movenext
 wend


end if

end if



set correio = nothing

set buscaemail = nothing






           set bdandamos = nothing
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
