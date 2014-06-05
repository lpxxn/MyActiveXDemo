<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="WebForm1.aspx.cs" Inherits="WebApplication1.WebForm1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <script type="text/javascript">
        function ClearClip() {
            var activeObj = document.getElementById("mytt");
            if (activeObj) {
                activeObj.ClearClip();               
            }
        }
        function ConvertAndUploadHtml() {
            var activeObj = document.getElementById("mytt");
            if (activeObj) {
                activeObj.initService(this.window, 'callbackFun', '1',"http://192.168.112.205:9090/re/file/fileupload.koala");
                activeObj.ConvertClip();
            }
        }
        function ConvertAndUploadWord() {
            var activeObj = document.getElementById("mytt");
            if (activeObj) {
                activeObj.initService(this.window, 'callbackFun', '2',"http://192.168.112.205:9090/re/file/fileupload.koala");
                activeObj.ConvertClip();
            }
        }
        function callbackFun(state, msg) {
            alert(state);
            var activeObj = document.getElementById("mytt");
            if (activeObj) {
                var content = activeObj.GetReultHTML();
                document.getElementById('output').innerHTML = content;
                alert(content);
            }
        } 
</script>

</head>
<body>
    <form id="form1" runat="server">
    <div>
    <object id="mytt" classid="clsid:6169E98E-DA08-4E87-81B6-EE3A5034C0E2"        
        codebase="/Ac/WordToHTML.cab"></object>
    </div>

    </br>
    <input type="button" onClick="javascript:ClearClip();" value="清除剪切板"></input>
    </br>
    <input type="button" onClick="javascript:ConvertAndUploadHtml();" value="转换上传Html"></input>
    </br>
    <input type="button" onClick="javascript:ConvertAndUploadWord();" value="转换上传Word"></input>
    </form>
</body>
</html>
