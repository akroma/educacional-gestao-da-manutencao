<%
'crio o objeto correio'
'Set Mail = Server.CreateObject("SMTPsvg.Mailer")
set mail = server.createObject("Persits.MailSender") 
'configuro a mensagem 
'assinalo o servidor de sa�da para enviar o correio 
mail.host = "sn116anotes1.cfp116.sp.senai.br" 
'indico o endere�o de correio do remitente 
mail.from = "suporteinf116@sp.senai.br" 
'indico o endere�o do destinat�rio da mensagem 
mail.addAddress "edilsonfsp@ig.com.br" 
'indico o corpo da mensagem 
mail.body = "Se voc� recebeu funcionou" 
'o envio 
'certifico-me que n�o se apresentem erros na p�gina se se produzem 


On Error Resume Next 
mail.send 
if Err <> 0 then 
response.write "Erro, n�o pode completar a opera��o" 
else 
response.write "Obrigado por preencher o formul�rio. Foi enviado corretamente." end if 
%>