<%
 if request.form("voltar") = "" then
  response.write "<center><a href=menu.asp><img src=menu.jpg border=0></a> <a href=javascript:history.go(-1)><img src=voltar.jpg border=0></a> <input type=image src=enviar.jpg> <a href=index.asp><img src=sair.jpg border=0></a></center>"
 else
  response.write "<center><a href=menu.asp><img src=menu.jpg border=0></a> </a><input type=image src=enviar.jpg> <a href=index.asp><img src=sair.jpg border=0></a></center>"
 end if
%>


 