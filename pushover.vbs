set xhr = Createobject("MSXML2.ServerXMLHTTP")
url = "https://api.pushover.net/1/messages.json"
data = "{*token*: *Your Token Here*, *user*: *Your User Key Here*, *title*: *Title Here*, *message*: *Message Here.*}"

xhr.open "POST", url, true
xhr.setRequestHeader "Content-type", "application/json"
xhr.send REPLACE(data,"*",Chr(34))
