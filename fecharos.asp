<html>
<head><title>CFP 1.16 - Ordem de Serviço</title></head>
<body bgcolor=#F0FFF0>
<form method=post action=fecharos.asp> 
<table bgcolor=#C1CDC1 width=510 height=370 align=center cellspacing=0 cellpadding=0>
 <tr height=15>
  <td align=center> 
<% session.LCID = 1046
   response.expires = -1000 
   response.write " <i><font face=verdana, arial size=1 color=#ff0000>" & date() & " - Usuário: " & session("func") & "</font></i>" 
%></td>
 </tr>
 <tr bgcolor=#8B8682 height=15>
  <td align=center>FECHAR</td>
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

         if session("cargo") = 1 or session("cargo") = 6 then
          set bdfecharos = conexao.execute("select nros, local, codfunc, func, servso, obs from oscon where codsitua = 1 order by nros")       
         else
          set bdfecharos = conexao.execute("select nros, local, codfunc, func, servso, obs from oscon where codsitua = 1 and executor ='" & session("ni") & "' order by nros") 
         end if

         if not bdfecharos.eof then
          response.write "<select name=nrosfr>"
          while not bdfecharos.eof
           response.write "<option value=" & bdfecharos("nros") & ">" & bdfecharos("nros") & " - " & bdfecharos("func") & "</option>" & "<br>"
           bdfecharos.movenext
          wend
          set bdfecharos = nothing
          response.write "</select><input type=hidden name=fecharosfr1 value=sim>"
         else
          response.write "Não existe OS aberta" 
          set bdfecharos = nothing 
         end if


        else
         if request.form("fecharosform2") = "" then
          set bdfecharos = conexao.execute("select * from oscon where nros =" & request.form("nrosfr"))%>
          <table width=100% align=center border=1 bordercolor=#8B8682 cellpadding=1 cellspacing=0>
           <tr>
            <td width=30%>OS:</td>
            <td><font face=Arial size=2><%=bdfecharos("nros")%></font></td>
           </tr>
<%         if bdfecharos("nipatr") <> 0 then %>
           <tr>
            <td width=33%>NI:</td>
             <td><font face=Arial size=2><%=bdfecharos("nipatr")%></font></td>
           </tr>
<%         end if %>
           <tr>
            <td width=33%>Local:</td>
             <td><font face=Arial size=2><%=bdfecharos("local")%></font></td>
           </tr>

<%        set bdsitua = conexao.execute("select situa from situa where codsitua = " & bdfecharos("codsitua")) %>
          <td><font face=Arial size=2><b>Situação</font></td>
          <td><font face=Arial size=2><%=bdsitua("situa")%></font></td></tr>
   <%    set bdsitua = Nothing %>

<%         if session("cargo") = "4" then %>
            <tr>
             <td>Solicitante:</td>
             <td><font face=Arial size=2><%=bdfecharos("func")%></font></td>
                                  
<%         else
            if session("cargo") = "5" then 
              
             set bdexec = conexao.execute("select func from func where codfunc = '" & bdfecharos("executor") & "'")  
%>          
             <tr>
              <td>Executor:</td>
              <td><font face=Arial size=2><%=bdexec("func")%></font></td>
                        
<%          else
             set bdexec = conexao.execute("select func from func where codfunc = '" & bdfecharos("executor") & "'") %>
              <tr>
               <td>Solicitante:</td>
               <td><font face=Arial size=2><%=bdfecharos("func")%></font></td>
              <tr>
               <td>Executor:</td>
               <td><font face=Arial size=2><%=bdexec("func")%></font></td>
<%            
            end if
           end if
%>




           <tr>
            <td colspan=2>Dados do Pedido:</td>
           </tr>
           <tr> 
            <td colspan=2><font face=Arial size=2><%=bdfecharos("servso")%></font></td>
           </tr>
           <tr>
            <td colspan=2>Andamento</td>
           </tr>
<%
          set bdandam = conexao.execute("select * from andam where nros =" & bdfecharos("nros") & " order by nrandam")
         if not bdandam.eof then
          while not bdandam.eof 
           response.write  "<tr><td><font face=Arial size=1>" & bdandam("dtandam") & "</font></td>"
           set bdsolic = conexao.execute("select func from func where codfunc = '" & bdandam("codfunc") & "'")
           solic = bdsolic("func")
           response.write "<td><font face=Arial size=1>" & bdsolic("func") & "</font>"
           set bdsolic = nothing    
           response.write "</td></tr><tr><td colspan=2><font face=Arial size=2>" & bdandam("andam") & "</font></td></tr>"          
           bdandam.movenext
          wend 
         else
          response.write "<tr><td colspan=2><font face=Arial size=2>Sem registro de andamento</font></td></tr>" 
         end if  
         set bdandamos = nothing  
%>

           <tr>
            <td>Tempo Gasto:</td>
            <td><input type=text name=horafr size=2 maxlength=3>H<input type=text name=minfr size=2 maxlength=2>M</td>
           </tr>
           <tr>
            <td colspan=2>Serviço realizado:</td>
           </tr>
           <tr> 
            <td colspan=2><textarea align=center name=servexefr rows=2 cols=60></textarea></td>
           </tr>
           <tr>
            <td colspan=2>Material Utilizado/Observações:</td>
           </tr>
           <tr> 
            <td colspan=2><textarea align=center name=matutilfr rows=2 cols=60></textarea></td>
           </tr>
          </table>
          <input type=hidden name=solicfr value=<%=bdfecharos("codfunc")%>>
          <input type=hidden name=nrosfr value=<%=bdfecharos("nros")%>>
          <input type=hidden name=fecharosfr1 value=sim>
          <input type=hidden name=fecharosform2 value=sim>
          <input type=hidden name=voltar value=nao>
<%        set bdfecharos = nothing  
         else 
          if request.form("horafr") = "" and request.form("minfr") = "" then 
           response.write "Campo tempo gasto não pode ser em branco<br><a href=javascript:history.go(-1)>Voltar</a>"            
          else 
           if request.form("horafr") = "" and request.form("minfr") <> "" then  
            horafr = 0
            minfr = request.form("minfr")
           end if 
           if request.form("minfr") = "" and request.form("horafr") <> "" then  
            horafr = request.form("horafr")
            minfr = 0
           end if
           if request.form("minfr") <> "" and request.form("horafr") <> "" then  
            horafr = request.form("horafr")
            minfr = request.form("minfr")
           end if
           if request.form("servexefr") = "" then
            response.write "Campo solução não pode ser em branco<br><a href=javascript:history.go(-1)>Voltar</a>"
           else
            if isNumeric(minfr) = false then
             response.write "O campo tempo gasto minuto deve ser um número<br><a href=javascript:history.go(-1)>Voltar</a>"   
            else
             if isNumeric(horafr) = false then 
              response.write "O campo tempo gasto hora deve ser um número<br><a href=javascript:history.go(-1)>Voltar</a>"   
             else
              if minfr > 59 then 
               response.write "Campo tempo gasto minutos deve ser um número menor que 60<br><a href=javascript:history.go(-1)>Voltar</a>"
              else
               tpgas = horafr + minfr / 100
               set bdandamos = conexao.execute("insert into andam (nros, andam, codfunc) values (" & request.form("nrosfr") & ",'" & request.form("servexefr") & "','" & session("ni") & "')")
               set bdfecharos = conexao.execute("update os set matutil = '" & request.form("matutilfr") & "', tpreal = '" & tpgas & "', dtsol = Now(), codsitua = 14 where nros = " & request.form("nrosfr"))
               response.write "A OS " & request.form("nrosfr") & " foi fechada com sucesso!<br><a href=fecharos.asp>Fechar outra OS</a>"
set buscaemail = conexao.execute("select email from func where codfunc ='" & request.form("solicfr") & "'")

assunto = "Por favor avaliar a OS " & request.form("nrosfr") & " fechada por " & session("func")

set correio = Server.CreateObject("CDONTS.NewMail") 

correio.send "suporteinf116@sp.senai.br (Sistema de OS)", buscaemail("email"), assunto, "Este correio foi enviado automaticamente pelo programa de OS, por favor não responder.", 2

set correio = nothing

set buscaemail = nothing

               set bdfecharos = nothing
              end if
             end if
            end if
           end if
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
