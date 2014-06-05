using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Reflection;
using System.IO;
using Microsoft.Office.Interop.Word;
using System.Net;
using mshtml;
using System.Threading;
namespace PERA.WORDHELPER
{

    internal enum ConvertType
    {
        HTML = 1,
        WORD
    }

    [Guid("6169E98E-DA08-4E87-81B6-EE3A5034C0E2"), ProgId("MyActiveXDemo.UserControl1"), ComVisible(true)]
    public partial class UserControl1 : UserControl, IObjectSafety
    {

        #region IObjectSafety 成员 格式固定

        private const string _IID_IDispatch = "{00020400-0000-0000-C000-000000000046}";
        private const string _IID_IDispatchEx = "{a6ef9860-c720-11d0-9337-00a0c90dcaa9}";
        private const string _IID_IPersistStorage = "{0000010A-0000-0000-C000-000000000046}";
        private const string _IID_IPersistStream = "{00000109-0000-0000-C000-000000000046}";
        private const string _IID_IPersistPropertyBag = "{37D84F60-42CB-11CE-8135-00AA004BB851}";

        private const int INTERFACESAFE_FOR_UNTRUSTED_CALLER = 0x00000001;
        private const int INTERFACESAFE_FOR_UNTRUSTED_DATA = 0x00000002;
        private const int S_OK = 0;
        private const int E_FAIL = unchecked((int)0x80004005);
        private const int E_NOINTERFACE = unchecked((int)0x80004002);

        private bool _fSafeForScripting = true;
        private bool _fSafeForInitializing = true;

        public int GetInterfaceSafetyOptions(ref Guid riid, ref int pdwSupportedOptions, ref int pdwEnabledOptions)
        {
            int Rslt = E_FAIL;

            string strGUID = riid.ToString("B");
            pdwSupportedOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER | INTERFACESAFE_FOR_UNTRUSTED_DATA;
            switch (strGUID)
            {
                case _IID_IDispatch:
                case _IID_IDispatchEx:
                    Rslt = S_OK;
                    pdwEnabledOptions = 0;
                    if (_fSafeForScripting == true)
                        pdwEnabledOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER;
                    break;
                case _IID_IPersistStorage:
                case _IID_IPersistStream:
                case _IID_IPersistPropertyBag:
                    Rslt = S_OK;
                    pdwEnabledOptions = 0;
                    if (_fSafeForInitializing == true)
                        pdwEnabledOptions = INTERFACESAFE_FOR_UNTRUSTED_DATA;
                    break;
                default:
                    Rslt = E_NOINTERFACE;
                    break;
            }

            return Rslt;
        }

        public int SetInterfaceSafetyOptions(ref Guid riid, int dwOptionSetMask, int dwEnabledOptions)
        {
            int Rslt = E_FAIL;
            string strGUID = riid.ToString("B");
            switch (strGUID)
            {
                case _IID_IDispatch:
                case _IID_IDispatchEx:
                    if (((dwEnabledOptions & dwOptionSetMask) == INTERFACESAFE_FOR_UNTRUSTED_CALLER) && (_fSafeForScripting == true))
                        Rslt = S_OK;
                    break;
                case _IID_IPersistStorage:
                case _IID_IPersistStream:
                case _IID_IPersistPropertyBag:
                    if (((dwEnabledOptions & dwOptionSetMask) == INTERFACESAFE_FOR_UNTRUSTED_DATA) && (_fSafeForInitializing == true))
                        Rslt = S_OK;
                    break;
                default:
                    Rslt = E_NOINTERFACE;
                    break;
            }

            return Rslt;
        }

        #endregion

        public UserControl1()
        {
            InitializeComponent();                               
        } 

        #region 字段和属性

        readonly string m_SAVEFOLDER = Environment.ExpandEnvironmentVariables("%temp%")+@"\ConvertHtml";
        private Microsoft.Office.Interop.Word.Application m_app;
        private Microsoft.Office.Interop.Word.Document m_oDoc;
        private ConvertType myConvertType;
        #region Office格式 
        private readonly string[] m_docType = { "Object Descriptor", "Rich Text Format", "HTML Format", "System.String","UnicodeText", "Text", "EnhancedMetafile", "MetaFilePict", "Embed Source","Link Source", "Link Source Descriptor", "ObjectLink","Hyperlink"};
        private readonly string[] m_docType2 = { "Object Descriptor", "Rich Text Format", "HTML Format", "System.String","UnicodeText", "Text", "EnhancedMetafile", "MetaFilePict","Embed Source", "Link Source", "Link Source Descriptor", "ObjectLink"};
        private readonly string[] m_docType3 = { "Object Descriptor","Rich Text Format","HTML Format","EnhancedMetafile","MetaFilePict","PNG","GIF","JFIF","PNG+Office Art","GIF+Office Art","JFIF+Office Art","Office Drawing Shape Format","ActiveClipboard","DeviceIndependentBitmap","System.Drawing.Bitmap","Bitmap","Embed Source","Link Source","Link Source Descriptor","ObjectLink","Hyperlink"};
        private readonly string[] m_docType4 = { "Art::GVML ClipFormat", "System.Drawing.Bitmap", "Bitmap", "PNG", "JFIF", "GIF", "EnhancedMetafile", "MetaFilePict", "Object Descriptor"};
        private readonly string[] m_docType5 = { "Rich Text Format", "HTML Format", "EnhancedMetafile", "MetaFilePict", "PNG", "GIF", "JFIF", "ActiveClipboard", "DeviceIndependentBitmap", "System.Drawing.Bitmap", "Bitmap", "Embed Source", "Link Source", "Link Source Descriptor", "ObjectLink", "Hyperlink" };
        private readonly string[] m_docType6 = { "MathML", "MathML Presentation", "Object Descriptor", "Rich Text Format", "HTML Format", "System.String", "UnicodeText", "Text", "EnhancedMetafile", "MetaFilePict", "Embed Source", "Link Source", "Link Source Descriptor", "ObjectLink", "Hyperlink" };
        private readonly string[] m_docType7 = { "Office Drawing Shape Format", "Embedded Object", "MetaFilePict", "EnhancedMetafile", "Object Descriptor", "Link Source", "PNG+Office Art", "JFIF+Office Art", "GIF+Office Art", "PNG", "JFIF", "GIF", "ActiveClipboard", "HTML Format" };
        //这个_docType9 和上边的7在xp系统和win7系统不一样，只有一个字母大小写不一样
        private readonly string[] m_docType9 = { "Office Drawing Shape Format", "Embedded Object", "MetaFilePict", "EnhancedMetafile", "Object Descriptor", "Link Source", "PNG+Office Art", "JFIF+Office Art", "GIF+Office Art", "PNG", "JFIF", "GIF", "ActiveClipBoard", "HTML Format" };
        private readonly string[] m_docType8 = { "Object Descriptor", "Rich Text Format", "HTML Format", "System.String", "UnicodeText", "Text", "EnhancedMetafile", "MetaFilePict", "Embed Source", "Link Source", "Link Source Descriptor", "ObjectLink", "HyperlinkWordBkmk", "Hyperlink" };
        #endregion
        private HTMLWindow2Class m_htmlWindow;
        private string m_jsCallBackFun;
        private string m_strURL;
        private string m_returnHTML=string.Empty;
        private string m_isDebugState = "0";//如果是1就打开debug状态
        #endregion

        #region 公开的方法
        /// <summary>
        /// 初始化
        /// </summary>
        /// <param name="htmlWindow">调用页面</param>
        /// <param name="jsCallBackFun">回调方法</param>
        /// <param name="flag">1为html 2为WORD</param>
        /// <param name="url">上传服务器的地址</param>
        /// <param name="debug">可选参数为1时才弹出bug窗体</param>       
        public void initService(IHTMLWindow2 htmlWindow, String jsCallBackFun, string flag,string url,string debug="0")
        {
            myConvertType = (ConvertType)Enum.Parse(typeof(ConvertType), flag);
            this.m_htmlWindow = (HTMLWindow2Class)Marshal.CreateWrapperOfType(htmlWindow, typeof(HTMLWindow2Class));
            this.m_jsCallBackFun = jsCallBackFun;
            m_isDebugState = debug;
            m_strURL = url;            
        }

        /// <summary>
        /// 把剪切板里的内容转换
        /// </summary>
        public void ConvertClip()
        {
            if (CheckValue())
            {
                DoAction(() => ConvertToHtml(myConvertType));
            }
        }

        /// <summary>
        /// 把剪切板里的内容转换
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="para"></param>
        public void doConvert(String filePath, String para)
        {


            if (string.IsNullOrEmpty(filePath.Trim()))
            {
                if (!CheckValue())
                    return;
            }
            else
            {
                if (!File.Exists(filePath.Trim()))
                {
                    //ShowDebugMessage("Word的存放路径不正确！请检查" );
                    MessageBox.Show("Word的存放路径" + filePath + "不正确,请检查!");
                    return;
                }
                //判断是否是打开状态
                try
                {
                    Stream s = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.None);
                    s.Close();
                }
                catch (Exception)
                {
                    MessageBox.Show("要上传的Word处于打开状态，请关闭后再上传！");
                    return;
                }
            }
            DoAction(() => ConvertToHtml(myConvertType, filePath, para));

        }

        public void ClearClip()
        {
            Clipboard.Clear();
            
            MessageBox.Show("清理剪切板成功!");
        }

        /// <summary>
        /// 得到返回的HTML
        /// </summary>
        /// <returns></returns>
        public string GetReultHTML()
        {
            return m_returnHTML;
        }
        
        #endregion

        /// <summary>
        /// 检查数据是否是word里复制的
        /// </summary>
        /// <returns></returns>
        private bool CheckValue()
        {
            
            if (string.IsNullOrEmpty(m_strURL))
            {
                javaScriptEventFire("0", "上传服务器的Url错误");
                //ShowDebugMessage("上传服务器的Url错误");                
                return false;
            }
            
            return true;

            //TODO 为了能粘贴excel还有point等内容这个判断先去掉，
            //IDataObject data = Clipboard.GetDataObject();
            
            //string[] m_strFormats = data.GetFormats();
            
            //if (compareArray(m_strFormats, m_docType) || compareArray(m_strFormats, m_docType2) || compareArray(m_strFormats, m_docType3) ||
            //    compareArray(m_strFormats, m_docType4) || compareArray(m_strFormats, m_docType5) || compareArray(m_strFormats, m_docType6) ||
            //    compareArray(m_strFormats, m_docType7) || compareArray(m_strFormats, m_docType8) || compareArray(m_strFormats, m_docType9))
            //{
            //}
            //else
            //{
            //    javaScriptEventFire("0", "请复制word里的内容");
            //    string _sF = ""; //bug状态下查找在word里复制的内容是什么格式 有不支持的模式时可以打开查看格式
            //    m_strFormats.ToList().ForEach(x => _sF += "\"" + x + "\",");
            //    //MessageBox.Show(_sF);                
            //    ShowDebugMessage("内容的格式为"+_sF);
            //    return false;
            //}
            //return true;
        }

        private void DoAction(Action action)
        {
            Thread thread = new Thread(() =>
            {
                action();
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.IsBackground = true;
            thread.Start();
        }

        /// <summary>
        /// 回调HTML页面方法
        /// </summary>
        /// <param name="state">1是成功0是失败</param>
        /// <param name="meg">失败时传的信息</param>
        private void javaScriptEventFire(string state, string meg = "")
        {
            try
            {
                string script = string.Format("{0}('{1}','{2}');", m_jsCallBackFun, state, meg);
                m_htmlWindow.execScript(script, "JavaScript");
            }
            catch (Exception ex)
            {                
            }
        }

        #region  检查复制的内容是不是从word里复制的
        /// <summary>
        /// 检查复制的内容是不是从word里复制的
        /// </summary>
        /// <param name="obj1"></param>
        /// <param name="obj2"></param>
        /// <returns></returns>
        private bool compareArray(string[] obj1, string[] obj2)
        {
            if (obj1.Length == obj2.Length)
            {
                var m_count = obj1.Except(obj2).ToArray();
                if (m_count.Length == 0)
                {
                    return true;
                }
            }
            return false;
        }
        int _index = 0;

        #region
        private void ConvertToHtml(ConvertType convertType, String filePath = "", String para = "")
        {
            Monitor.Enter(this);
            bool _haveErreo = false;//是否有错误
            object FileName = m_SAVEFOLDER + @"\myfile" + _index++ + ".htm";
            bool isExists = Directory.Exists(m_SAVEFOLDER);
            if (!isExists)
                Directory.CreateDirectory(m_SAVEFOLDER);
            try
            {
                #region 用word转换成html
                //filePath如果为空就用剪切板里的内容
                if (string.IsNullOrEmpty(filePath.Trim()))
                {

                    m_app = new Microsoft.Office.Interop.Word.Application();
                    m_app.Visible = false;
                    object oMissing = Missing.Value;
                    m_oDoc = m_app.Documents.Add(ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                    m_oDoc.Content.Paste();

                   

                    object FileFormat = WdSaveFormat.wdFormatFilteredHTML;
                    if (convertType == ConvertType.WORD)
                    {
                        FileFormat = WdSaveFormat.wdFormatDocument;
                        FileName = FileName.ToString().Replace(".htm", ".doc");
                    }
                    m_oDoc.SaveAs(ref FileName, ref FileFormat);
                }
                #endregion
            }
            catch (Exception ex)
            {
                ShowDebugMessage("保存Word失败！错误方法名为ConvertToHtml：详细信息：" + ex.Message);
                _haveErreo = true;
            }
            finally
            {
                if (string.IsNullOrEmpty(filePath.Trim()))
                {
                    m_oDoc.Close(null, null, null);
                    m_app.Quit();
                    m_oDoc = null;
                    m_app = null;
                    GC.Collect();
                }

                #region 上传
                if (!_haveErreo)
                {
                    if(convertType == ConvertType.WORD)
                    {
                        m_returnHTML = JAVAFieldsMethod(m_SAVEFOLDER, string.Empty, convertType, filePath, para);
                    }
                    else if (convertType == ConvertType.HTML)
                    {
                        //上传文件                        
                        m_returnHTML = JAVAFieldsMethod(m_SAVEFOLDER + @"\myfile" + --_index + ".files", FileName.ToString(), convertType);
                    }
                    if (!string.IsNullOrEmpty(m_returnHTML.Trim()))
                    {
                        javaScriptEventFire("1");
                    }
                    else
                    {
                        javaScriptEventFire("0", "上传失败");    
                    }
                    DeleteFolderrsFiles(m_SAVEFOLDER);//删除文件
                }
                else
                {
                    javaScriptEventFire("0", "转换失败");
                }
                #endregion
                Monitor.Exit(this);
            }            
        }
        
        #endregion

        
        #endregion


        #region 把图片和html上传到服务器上
        /// <summary>
        /// 把图片和html整理成集合上传
        /// </summary>
        /// <param name="imageOrWordPath">图片地址</param>
        /// <param name="htmlPath">html地址</param>
        /// <returns></returns>
        private string JAVAFieldsMethod(string imageOrWordPath, string htmlPath, ConvertType convertType,String filePath = "", String para = "")
        {
            string _returnValue = string.Empty;
            if (Directory.Exists(imageOrWordPath))
            {
                try
                {
                    int _fileIndex = 0;
                    Dictionary<string, string> _dic = new Dictionary<string, string>();//图片或doc
                    Dictionary<string, string> _Html = new Dictionary<string, string>();//html
                    DirectoryInfo _di = new DirectoryInfo(imageOrWordPath);
                    switch (convertType)
                    {
                        case ConvertType.HTML:
                            if (File.Exists(htmlPath))
                            {
                                FileInfo[] _fi = _di.GetFiles("*.png");
                                foreach (FileInfo tmpfi in _fi)
                                {
                                    string _fileKey = "fileUpload" + _fileIndex++;
                                    _dic.Add(_fileKey, tmpfi.FullName);
                                }
                                _fi = _di.GetFiles("*.jpg");
                                foreach (FileInfo tmpfi in _fi)
                                {
                                    string _fileKey = "fileUpload" + _fileIndex++;
                                    _dic.Add(_fileKey, tmpfi.FullName);
                                }
                                FileInfo _fileInfo = new FileInfo(htmlPath);
                                FileStream _fileStream = new FileStream(_fileInfo.FullName, FileMode.Open, FileAccess.Read);
                                StreamReader _reader = new StreamReader(_fileStream, Encoding.Default);
                                string _strRetrun = _reader.ReadToEnd();

                                _Html.Add("WordHtml", _strRetrun);
                                _reader.Close();
                                _fileStream.Close();
                            }
                            break;
                        case ConvertType.WORD:
                            //如果是空的时候用剪切板里转换的wod
                            if (string.IsNullOrEmpty(filePath.Trim()))
                            {
                                FileInfo[] _fileInfoDoc = _di.GetFiles("*.doc");
                                foreach (FileInfo tmpfi in _fileInfoDoc)
                                {
                                    string _fileKey = "fileUpload" + _fileIndex++;
                                    _dic.Add(_fileKey, tmpfi.FullName);
                                }
                                _fileInfoDoc = _di.GetFiles("*.docx");
                                foreach (FileInfo tmpfi in _fileInfoDoc)
                                {
                                    string _fileKey = "fileUpload" + _fileIndex++;
                                    _dic.Add(_fileKey, tmpfi.FullName);
                                }
                            }
                            else
                            {
                                _dic.Add("fileUpload", filePath);
                                
                            }
                            
                            _Html.Add("EditorParams", para);                            
                            break;
                    }
                    _returnValue = JavaUploadFile(m_strURL, _Html, _dic);
                }
                catch (Exception ex)
                {
                    switch (convertType)
                    {
                        case ConvertType.HTML:
                            ShowDebugMessage("把图片和html上传到服务器上发生错误：方法名JAVAFieldsMethod" + ex.Message);
                            break;
                        case ConvertType.WORD:
                            ShowDebugMessage("把Word上传到服务器上发生错误：方法名JAVAFieldsMethod" + ex.Message);
                            break;
                    }
                    return _returnValue;
                }
                finally
                {

                }
            }
            return _returnValue;
        }

        /// <summary>
        /// 把图片和HTML上传到服务器上去
        /// </summary>
        /// <param name="url">服务器URL</param>
        /// <param name="nvc">图片集合</param>
        /// <param name="fileMap">HTML集合</param>
        /// <returns></returns>
        public string JavaUploadFile(string url, Dictionary<string,string> nvc,Dictionary<string,string> fileMap)
        {

            string BOUNDARY = "---------------------------123821742118716";//
            byte[] boundarybytes = Encoding.UTF8.GetBytes("\r\n--" + BOUNDARY + "\r\n");

            HttpWebRequest wr = (HttpWebRequest)WebRequest.Create(url);
            wr.ContentType = "multipart/form-data; boundary=" + BOUNDARY;
            wr.Method = "POST";
            wr.UserAgent = "Mozilla/5.0 (Windows; U; Windows NT 6.1; zh-CN; rv:1.9.2.6)";
            wr.KeepAlive = true;
            wr.Credentials = CredentialCache.DefaultCredentials;            
            Stream rs = wr.GetRequestStream();
            if (nvc != null && nvc.Count > 0)
            {
                StringBuilder strBuf = new StringBuilder();
                foreach (string key in nvc.Keys)
                {
                    string _inputValue = nvc[key];
                    strBuf.Append("\r\n").Append("--").Append(BOUNDARY)
                            .Append("\r\n");
                    strBuf.Append("Content-Disposition: form-data; name=\""
                            + key + "\"\r\n\r\n");
                    strBuf.Append(_inputValue);
                }
                byte[] formitembytes = Encoding.UTF8.GetBytes(strBuf.ToString());
                rs.Write(formitembytes, 0, formitembytes.Length);
            }          

            if (fileMap != null && fileMap.Count > 0)
            {
                foreach (string key in fileMap.Keys)
                {
                    FileInfo _fileInfo = new FileInfo(fileMap[key]);
                    string _fileName = _fileInfo.Name;
                    string _contType = "application/octet-stream";
                    if (_fileName.EndsWith(".png"))
                    {
                        _contType = "image/png";
                    }
                    else if (_fileName.EndsWith(".jpg"))
                    {
                        _contType = "image/jpg";
                    }
                    else if (_fileName.EndsWith(".doc") || _fileName.EndsWith(".docx"))
                    {
                        _contType = "application/msword";
                    }
                    StringBuilder strBuf = new StringBuilder();
                    strBuf.Append("\r\n").Append("--").Append(BOUNDARY)
                            .Append("\r\n");
                    strBuf.Append("Content-Disposition: form-data; name=\""
                            + key + "\"; filename=\"" + _fileName
                            + "\"\r\n");
                    strBuf.Append("Content-Type:" + _contType + "\r\n\r\n");
                    byte[] formitembytes = Encoding.UTF8.GetBytes(strBuf.ToString());
                    rs.Write(formitembytes, 0, formitembytes.Length);

                    FileStream fileStream = new FileStream(_fileInfo.FullName, FileMode.Open, FileAccess.Read);
                    byte[] buffer = new byte[4096];
                    int bytesRead = 0;
                    while ((bytesRead = fileStream.Read(buffer, 0, buffer.Length)) != 0)
                    {
                        rs.Write(buffer, 0, bytesRead);
                    }
                    fileStream.Close();
                }
            }

            byte[] trailer = Encoding.UTF8.GetBytes("\r\n--" + BOUNDARY + "--\r\n");
            rs.Write(trailer, 0, trailer.Length);
            rs.Close();

            WebResponse wresp = null;
            string _strRetrun = string.Empty;
            try
            {
                wresp = wr.GetResponse();
                Stream _streamR = wresp.GetResponseStream();
                StreamReader _reader = new StreamReader(_streamR, Encoding.UTF8);
                _strRetrun = _reader.ReadToEnd();              
            }
            catch (Exception ex)
            {                
                if (wresp != null)
                {
                    wresp.Close();
                    wresp = null;
                }
                ShowDebugMessage("转换错误 方法名：JavaUploadFile 详细息：" + ex.Message);
                return string.Empty;
            }
            finally
            {
                wr = null;                    
            }
            return _strRetrun;
        }

        #region 删除文件夹下的所有文件
        public void DeleteFolderrsFiles(string path)
        {
            try
            {
                path = path.Trim();
                if (Directory.Exists(path))
                {
                    string[] _Foloders = Directory.GetDirectories(path);
                    string[] _Files = Directory.GetFiles(path);
                    foreach (var _fiel in _Files)
                    {
                        File.Delete(_fiel);
                    }
                    foreach (var _folor in _Foloders)
                    {
                        Directory.Delete(_folor, true);
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }
        #endregion
        #endregion

        #region 显示信息
        public void ShowDebugMessage(string error)
        {
            if (m_isDebugState == "1")
            {
                MessageBox.Show(error);
            }
        }
        #endregion
    }

    
}
