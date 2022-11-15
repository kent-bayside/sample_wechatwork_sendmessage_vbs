' Sending messages with VBScript using WeChat Work API
' This is an exsample code to send a message with VBScript using WeChat Work API.

Dim http: Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

Dim corp_id: corp_id = "xxx"
Dim corp_secret: corp_secret = "xxx"
Dim url: url = "https://qyapi.weixin.qq.com/cgi-bin/gettoken?corpid=" + corp_id + "&corpsecret=" + corp_secret
Dim agentid: agentid = "999999"
Dim content: content = "Hello!"
Dim touser: touser = "test_user"

Dim res: res = ""
Dim token: token = ""

With http
    .Open "POST", url, False
    .SetRequestHeader "Content-Type", "application/json; charset=utf-8"
    .Send
    res = .ResponseText
    If InStr(res, "access_token") > 0 Then
      token = Mid(res, InStr(res, "access_token") + 15, 214)
    End If
End With

url = "https://qyapi.weixin.qq.com/cgi-bin/message/send?access_token=" + token

Dim data: data = "{ ""agentid"": " + agentid + ", ""msgtype"": ""text"", ""touser"": """ + touser + """, ""text"": {""content"": """ + content + """} }"

With http
    .Open "POST", url, False
    .SetRequestHeader "Content-Type", "application/json; charset=utf-8"
    .Send data
End With
