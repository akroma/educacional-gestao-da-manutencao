<html>
<head><title>CFP 1.16 - Ordem de Serviço</title></head>
<body bgcolor=#F0FFF0>
<form method=post action=consultaos.asp> 
<table bgcolor=#C1CDC1 width=510 height=370 align=center cellspacing=0 cellpadding=0>
 <tr height=15>
  <td align=center> 
<% session.LCID = 1046
   response.expires = -1000 
   response.write " <i><font face=verdana, arial size=1 color=#ff0000>" & date() & " - Usuário: " & session("func") & "</font></i>" 
%></td>
 </tr>
 <tr bgcolor=#8B8682 height=15>
  <td align=center>CONSULTAR OS</td>
 </tr>
 <tr>
  <td>
   <table width=100% align=center border=0> 
    <tr>
     <td align=center valign=center>
      <font face=arial size=2>
<%     set conexao = server.createobject("adodb.connection")
       conexao.open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("oscadastroteste.mdb") & ";uid=sa;pwd=;"
       if request.form("fecharosfr1") = "" then
        if session("cargo") = 1 or session("cargo") = 2 then
          set bdoscon = conexao.execute("select nros, local, codfunc, func, servso, obs from oscon order by nros desc")       
        else
          set bdoscon = conexao.execute("select nros, local, codfunc, func, servso, obs from oscon where codfunc ='" & session("ni") & "'or executor ='" & session("ni") & "' order by nros desc")
        end if
        if not bdoscon.eof then
         response.write "<select name=nrosfr>"
         while not bdoscon.eof
          response.write "<option value=" & bdoscon("nros") & ">" & bdoscon("nros") & " - " & bdoscon("func") & "</option>" & "<br>"
          bdoscon.movenext
         wend
         set bdoscon = nothing
         response.write "</select><input type=hidden name=fecharosfr1 value=sim>"
        else
         response.write "Não existe OS aberta" 
         set bdoscon = nothing 
        end if
       else
        if request.form("fecharosform2") = "" then
         set bdoscon = conexao.execute("select * from oscon where nros =" & request.form("nrosfr"))%>
         <table width=100% align=center border=1 bordercolor=#8B8682 cellpadding=1 cellspacing=0>
         <tr>
          <td width=10%><font face=Arial size=2><b>OS:</b></font></td>
          <td width=10%><font face=Arial size=2><%=bdoscon("nros")%></font></td>
          <td width=10%><font face=Arial size=2><b>Data</font></td>
          <td><font face=Arial size=2><font face=Arial size=2><%=bdoscon("dtabe")%></font></td>


          </tr>
<%         if bdoscon("nipatr") <> 0 then %>
           <tr>
            <td>NI:</td>
             <td><font face=Arial size=2><%=bdoscon("nipatr")%></font></td>
<%         else
            response.write "<td>NI:</td><td>0</td>"
           end if %>

            <td width=33%>Local:</td>
             <td><font face=Arial size=2><%=bdoscon("local")%></font></td>
           </tr>

<%      
        if session("cargo") = "4" then %>
            <tr>
             <td colspan=2>Solicitante:</td>
             <td colspan=2><font face=Arial size=2><%=bdoscon("func")%></font></td></tr>
                            
<%         else
            if session("cargo") = "5" then 
            if bdoscon("executor") <> "" then
              
             set bdexec = conexao.execute("select func from func where codfunc = '" & bdoscon("executor") & "'")  
%>          
     
              <td colspan=2>Executor:</td>
              <td colspan=2><font face=Arial size=2><%=bdexec("func")%></font></td>
                        
<%           end if   
       else
             set bdexec = conexao.execute("select func from func where codfunc = '" & bdoscon("executor") & "'") %>
              <tr>
               <td colspan=2>Solicitante:</td>
               <td colspan=2><font face=Arial size=2><%=bdoscon("func")%></font></td>
<%           if bdoscon("executor") <> "" then %>
              <tr>
               <td colspan=2>Executor:</td>
               <td colspan=2><font face=Arial size=2><%=bdexec("func")%></font></td>
<%           end if 
            end if
           end if
          set bdexec = Nothing %>
 
           <tr>
            <td colspan=4>Dados do Pedido:</td>
           </tr>
           <tr> 
            <td colspan=4><font face=Arial size=2><%=bdoscon("servso")%></font></td>
           </tr>


<%      
             
              
         observacoes = bdoscon("obs")
            
         if observacoes <> "" then
          response.write "<tr><td valign=top><font face=Arial size=2>Observações</font></td><td colspan=3 valign=top><font face=Arial size=1>" & observacoes & "</font></td></tr>"
         end if 
              
         materialutilizado = bdoscon("matutil")
             
         if materialutilizado <> "" then
          response.write "<tr><td valign=top height=30><font face=Arial size=2><b>Material Utilizado</font></td><td colspan=3 valign=top><font face=Arial size=1>" & materialutilizado & "</font></td></tr>"
         end if 
%>
           <tr>
            <td colspan=4>Andamento</td>
           </tr>
<%
          set bdandam = conexao.execute("select * from andam where nros =" & bdoscon("nros") & " order by nrandam")
         if not bdandam.eof then
          while not bdandam.eof 
           response.write  "<tr><td colspan=2><font face=Arial size=1>" & bdandam("dtandam") & "</font></td>"
           set bdsolic = conexao.execute("select func from func where codfunc = '" & bdandam("codfunc") & "'")
           solic = bdsolic("func")
           response.write "<td colspan=2><font face=Arial size=1>" & bdsolic("func") & "</font>"
           set bdsolic = nothing    
           response.write "</td></tr><tr><td colspan=6><font face=Arial size=2>" & bdandam("andam") & "</font></td></tr>"          
           bdandam.movenext
          wend 
         else
          response.write "<tr><td colspan=4><font face=Arial size=2>Sem registro de andamento</font></td></tr>" 
         end if  
         set bdandamos = nothing  

         set bdsitua = conexao.execute("select situa from situa where codsitua = " & bdoscon("codsitua")) %>
          <td colspan=1><font face=Arial size=2><b>Avaliação</font></td>
<% if bdoscon("aval1") <> "" then %>          
          <td colspan=1><font face=Arial size=2><%=bdoscon("aval1")%></font></td>
<%   else  %>

          <td colspan=1><font face=Arial size=2>N/A</font></td>
<% end if %>
          <td colspan=1><font face=Arial size=2><b>Situação</font></td>
          <td colspan=1><font face=Arial size=2><%=bdsitua("situa")%></font></td></tr>
   <%    set bdsitua = Nothing %>

          </table>
          <input type=hidden name=solicfr value=<%=bdoscon("codfunc")%>>
          <input type=hidden name=nrosfr value=<%=bdoscon("nros")%>>
          <input type=hidden name=fecharosfr1 value=sim>
          <input type=hidden name=fecharosform2 value=sim>
          <input type=hidden name=voltar value=nao>
<%        set bdoscon = nothing  
         else 
              
      
            
      
           if request.form("servexefr") = "" then
            response.write "Campo solução não pode ser em branco"
           else
            if isNumeric(minfr) = false then
             response.write "O campo tempo gasto minuto deve ser um número"   
            else
             if isNumeric(horafr) = false then 
              response.write "O campo tempo gasto hora deve ser um número"   
             else
              if minfr > 59 then 
               response.write "Campo tempo gasto minutos deve ser um número menor que 60"
              else


               set bdoscon = nothing
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
