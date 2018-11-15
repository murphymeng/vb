
'去除文本中所有html标签的函数，给小宝贝的
Public Function replaceHTMLTags(htmlText As String) As String
    Dim reg As Object
    Set reg = CreateObject("VBScript.Regexp")
    
    Dim str As String
    str = content
    
    reg.Global = True
    reg.IgnoreCase = True
    reg.pattern = "<(?:.|\n)*?>"
    htmlText = reg.Replace(htmlText, "")
 
    replaceHTMLTags = htmlText
End Function