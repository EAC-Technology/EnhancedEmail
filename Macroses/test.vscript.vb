logger("begin")

Set sett = new smtpsettings
sett.server = mailbox.server
sett.port = mailbox.port
sett.user = mailbox.user
sett.password = mailbox.password
sett.sender = mailbox.sender
sett.ssl = mailbox.ssl

Set myEACMessage = new mailmessage
myEACMessage.ttl = 1
myEACMessage.subject = "test"
myEACMessage.sender = sett.user
myEACMessage.recipients = "<addr>"
myEACMessage.body = "test"

'result = server.mailer.send_via(myEACMessage, sett)

'logger(result)

logger("end")