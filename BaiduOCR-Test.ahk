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
whr := ComObjCreate("WinHttp.WinHttpRequest.5.1")
whr.Open("GET", "https://aip.baidubce.com/oauth/2.0/token?grant_type=client_credentials&client_id=" apiKey "&client_secret=" secretKey)
whr.Send()
ret := Jxon_Load(whr.ResponseText)
if !ret.access_token
	throw ret.error "`n" ret.error_description

; 通用文字识别
whr.Open("POST", "https://aip.baidubce.com/rest/2.0/ocr/v1/general_basic?access_token=" ret.access_token)
whr.SetRequestHeader("Content-Type", "application/x-www-form-urlencoded")

b64Data := Base64Enc_FromFile("test.png")
body := "image=" UriEncode(b64Data) "&language_type=CHN_ENG"
whr.Send(body)

MsgBox % GetWinHttpText(whr) ; 结果要转换成UTF-8，不然会显示乱码
ExitApp

GetWinHttpText(objWinHttp, encoding := "UTF-8") {
	ado := ComObjCreate("adodb.stream")
	ado.Type     := 1 ; adTypeBinary = 1
	ado.Mode     := 3 ; adModeReadWrite = 3
	ado.Open()
	ado.Write(objWinHttp.ResponseBody)
	ado.Position := 0
	ado.Type     := 2 ; adTypeText = 2
	ado.Charset  := encoding
	return ado.ReadText(), ado.Close()
}