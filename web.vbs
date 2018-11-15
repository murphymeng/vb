
'抓取网页内容的函数'
Public Function getTimeFromAPI(url As String) As String


    Set HttpReq = CreateObject("MSXML2.ServerXMLHTTP")

    HttpReq.Open "get", url, False
    HttpReq.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    HttpReq.send
    
    getTimeFromAPI = HttpReq.responseText

End Function

'在一段文本内查找具有某种模式的文字，需要小宝贝了解下正则匹配是什么'
Public Function getText(pattern, content)
    Dim reg As Object
    Set reg = CreateObject("VBScript.Regexp")
    
    Dim str As String
    str = content
    
    reg.Global = True
    reg.pattern = pattern        '获取匹配结果'
    Dim matches As Object, match As Object
    Set allMatches = reg.Execute(str)
    
   If allMatches.Count <> 0 Then
      result = allMatches.Item(0).submatches.Item(0)
   End If
   getText = result
End Function

Sub main()
    response = getTimeFromAPI("https://mp.weixin.qq.com/s/byT6H5JXy85U53-776p4UQ")  '传入朱房路的某个文章地址获取网页内容'
    findString = getText("1、（高宇）国债方面(.*)收在3.32", response) '在上面的网页内容里根据某个正则匹配规则获取内容，小括号中间.*的部分就是要抓取的内容'

    Set WordApp = GetObject(class:="Word.Application")


    WordApp.Visible = True
    WordApp.Activate
    Set myDoc = WordApp.Documents.Open("C:\yiyi\test.docx")
    
    WordApp.ActiveDocument.Bookmarks("graph1").Select
    Set objSelection = WordApp.Selection
    objSelection.TypeText (findString)
    

End Sub


