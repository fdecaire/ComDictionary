<%
Dim MyComObject
Dim MyText

Set MyComObject = Server.CreateObject("ComDictionary.MyDictionary")

MyComObject("testvar") = "test data"

MyComObject.Clear

MyText = MyComObject("testvar")

%>
<html>
<head></head>
<body>
	<%=MyText %>
</body>
</html>
