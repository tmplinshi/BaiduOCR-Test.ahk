#NoEnv
#SingleInstance Force
#KeyHistory 0
SetWorkingDir %A_ScriptDir%
SetBatchLines -1
ListLines Off

; 接入流程
; http://ai.baidu.com/docs#/Begin/top
apiKey := "这里替换为你的apiKey"
secretKey := "这里替换为你的secretKey"

; 获取Access Token
; http://ai.baidu.com/docs#/Auth/top
url := "https://aip.baidubce.com/oauth/2.0/token?grant_type=client_credentials&client_id=" apiKey "&client_secret=" secretKey
WinHttpRequest(url, body)
ret := Jxon_Load(body)
if !ret.access_token
	throw ret.error "`n" ret.error_description

; 通用文字识别
url := "https://aip.baidubce.com/rest/2.0/ocr/v1/general_basic?access_token=" ret.access_token
b64Data := Base64Enc_FromFile("D:\Desktop\test.png")
body := "image=" UriEncode(b64Data) "&language_type=CHN_ENG"
WinHttpRequest(url, body,, "Charset: UTF-8")
MsgBox % body
ExitApp