				dim xhr
				dim url
				dim data
				dim starttime

				set xhr = Server.Createobject("MSXML2.ServerXMLHTTP")
				url = "https://api.pushover.net/1/messages.json"
				data = "{*token*: *enter token here*, *user*: *enter user*, *title*: *enter subject*, *message*: *enter message*}"

				xhr.open "POST", url, true
				xhr.setRequestHeader "Content-type", "application/json"
				xhr.send REPLACE(data,"*",Chr(34))

				startTime = Now
				Do Until DateDiff("s", startTime, now, 0, 0) > 1
				Loop
