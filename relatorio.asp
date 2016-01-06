<% 
if request.form("relaos") = "" then    %>
 <html>
  <head><title>CFP 1.16 - Ordem de Serviço</title></head>
  <body bgcolor=#F0FFF0>
   <table bgcolor=#C1CDC1 width=510 height=370 align=center cellspacing=0 cellpadding=0>
    <tr height=15>
     <td align=center>
<%    Session.LCID = 1046 
      response.expires = -1000 
      response.write " <i><font face=verdana, arial size=1 color=#ff0000>" & date() & " - Usuário: " & session("func") & "</font></i>" 
%>   </td>
    </tr>
    <tr bgcolor=#8B8682 height=15>
     <td align=center>RELATÓRIO</td>
    </tr>
    <tr>
    <td align=center>
     <form method=post action=relatorio.asp target=_new>
      <table align=center border=0>
        <tr> 
         <td align=center>
          <select name=relaopfr>
           <option selected value=1>OS ABERTA</option>
           <option value=2>OS FECHADA</option>
          </select>  
         </td>
         <td valign=center> 
          <input type=hidden name=relaos value=sim>
         </td>
        </tr>
       </table>
      </font>
     </td>
    </tr>
    <tr>
     <td valign=bottom>
      <!--#include file=botoes.asp-->
     </td>
    </tr>
   </form>
   </table>
  </body>
 </html>  <% 
else
 response.ContentType = "application/x-msexcel"
 response.write "<html><head><title>RELATÓRIO OS</title></head><body>"
 set conexao = server.createobject("adodb.connection") 
 conexao.open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("oscadastroteste.mdb") & ";uid=sa;pwd=;"

 set bdcarg = conexao.execute("select codfunc, codcarg, codnucl from func where func = '" & session("func") & "'")

 usuario = bdcarg("codfunc")
 cargo = bdcarg("codcarg")
 executor = bdcarg("codfunc")
 codnucl = bdcarg("codnucl")
 set bdcarg = nothing

 if cargo = 5 then 'Consulta de usuario
  set bdrela = conexao.execute("select * from relatorio where codfunc = '" & usuario & "' and codsitua = " & request.form("relaopfr") & " order by nros")
 else
  if cargo = 4 then 'Consulta de executor
   set bdrela = conexao.execute("select * from relatorio where executor = '" & usuario & "' and codsitua = " & request.form("relaopfr") & " order by nros")
  else
   if cargo = 2 then 'Consulta de coordenardor
    set bdrela = conexao.execute("select * from relatorio where codnucl = " & codnucl & " and codsitua = " & request.form("relaopfr") & " order by nros")
   else   'Consulta de administrador
    set bdrela = conexao.execute("select * from relatorio where codnucl = 8 and codsitua = " & request.form("relaopfr") & " order by nros")
   end if   
  end if 
 end if

 if not bdrela.eof then  'encontrou o registro procurado 
  
  if request.form("relaopfr") = 1 then
   tipo = "ABERTA"
  else
   tipo = "FECHADA"
  end if 
  response.write "<center><b>RELATÓRIO OS &nbsp;" & tipo & "</b></center><br><i><font face=verdana, arial size=2>" & date() & " - " & session("func") & "</font></i>" 
  response.write "<table width=1000 border=1 bordercolor=#8B8682 cellpadding=0 cellspacing=0><tr>"

  response.write "<td width=2% align=center><font face=Arial size=1>OS</font></td>"

  response.write "<td width=2% align=center><font face=Arial size=1>NI</font></td>"

  response.write "<td width=12% align=center><font face=Arial size=1>LOCAL</font></td>"

  if cargo = 4 or cargo = 1 or cargo = 2 then
   response.write "<td width=10% align=center><font face=Arial size=1>SOLICITANTE</font></td>"
   response.write "<td width=10% align=center><font face=Arial size=1>RAMAL</font></td>"
  end if

  if cargo = 5 or cargo = 1 or cargo = 2 then
   response.write "<td width=10% align=center><font face=Arial size=1>EXECUTOR</font></td>"
  end if  

  response.write "<td width=25% align=center><font face=Arial size=1>SERVIÇO SOLICITADO</font></td>"

  if tipo = "FECHADA" then 
   response.write "<td width=25% align=center><font face=Arial size=1>SERVIÇO EXECUTADO</font></td>"
  end if 

  if tipo = "FECHADA" then
   response.write "<td width=25% align=center><font face=Arial size=1>MAT.UTIL.</font></td>"
  end if

  if tipo = "FECHADA" then
   response.write "<td width=25% align=center><font face=Arial size=1>OBSERVAÇÕES</font></td>"
  end if 

  response.write "<td width=5% align=center><font face=Arial size=1>ABERTA</font></td>"

  if tipo = "FECHADA" then
   response.write "<td width=5% align=center><font face=Arial size=1>FECHADA</font></td></tr>"
  end if 

  while not bdrela.eof   
   response.write "<tr><td align=center valign=top><font face=Arial size=1>" & bdrela("nros") & "</font></td>"

   response.write "<td align=center valign=top><font face=Arial size=1>" & bdrela("nipatr") & "</font></td>"

   response.write "<td valign=top><font face=Arial size=1>" & bdrela("local") & "</font></td>"

   if cargo = 4 or cargo = 1 or cargo = 2 then
    set bdsolic = conexao.execute("select func, ramal from func where codfunc = '" & bdrela("codfunc") & "'") 
    response.write "<td valign=top><font face=Arial size=1>" & bdsolic("func") & "</font></td>"
    response.write "<td align=center valign=top><font face=Arial size=1>" & bdsolic("ramal")
    set bdsolic = Nothing

   end if

   if cargo = 5 or cargo = 1 or cargo = 2 then
    set bdnomeexecutor = conexao.execute("select func from func where codfunc = '" & bdrela("executor") & "'") 
    response.write "<td valign=top><font face=Arial size=1>" & bdnomeexecutor("func")
    set bdnomeexecutor = Nothing
    response.write "</font></td>"
   end if 

   response.write "<td valign=top><font face=Arial size=1>" & bdrela("servso") & "</font></td>"

   if tipo = "FECHADA" then
    response.write "<td valign=top><font face=Arial size=1>TESTE</font></td>"
   end if 

   if tipo = "FECHADA" then
    response.write "<td valign=top><font face=Arial size=1>" & bdrela("matutil") & "</font></td>"
   end if

   if tipo = "FECHADA" then
    response.write "<td valign=top><font face=Arial size=1>" & bdrela("obs") & "</font></td>"
   end if 

   response.write "<td align=center valign=top><font face=Arial size=1>" & mid(bdrela("dtabe"),1,10) & "</font></td>"

   if tipo = "FECHADA" then
    response.write "<td align=center valign=top><font face=Arial size=1>" & mid(bdrela("dtsol"),1,10) & "</td></font>"
   end if 

'   if bdrela("datasolucao") <> "" then 
'    response.write "<tr><td valign=top>Tempo gasto (hs)</td><td>" & bdrela("tpreal") & "</td></tr>"
'   end if 

   response.write "</tr>"
   bdrela.Movenext
  wend
  response.write "</table>"
 else 
  response.write "<center><font face=Arial size=2 color=#ff0000>Não existe ordem de serviço cadastrada para sua busca!</font></center>"
 end if 
 conexao.close
 set bdrela = nothing
 set conexao = nothing
 session("ossitua") = ""
 response.write "</body></html>"
end if
%>