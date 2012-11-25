using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using System.Globalization;
using System.Xml;
using System.Xml.Xsl;

namespace Askme
{
    public static class ErrorCode
    {
        public const string NOT_INI_SYSPROP = "Не инициализированны переменные настройки программы. \nДальнейшее выполнение не возможно, \nпрограмма завершает свою работу.";
        public const string UNKNOWN_ERROR = "Возникла ошибка, которая не может быть обработанна программой. \nДальнейшее выполнение не возможно, \nпрограмма завершает свою работу.";

    }

    public static class Common
    {
        public static SysProp SysProp = (SysProp)AskSerialize.Deserialize(Path.GetDirectoryName(Application.ExecutablePath) + "\\sysprop.dmp");
        
        // Путь к файлу помощи
        public static string HELP_PATH = Path.ChangeExtension(Application.ExecutablePath, ".chm");
        
        // Пути к файлам для преобразования Excel
        public static string SHARED_STRING = Path.GetDirectoryName(Application.ExecutablePath) + "\\convert\\sharedStrings.xml";
        public static string XML_CODE_BASE = Path.GetDirectoryName(Application.ExecutablePath) + "\\convert\\xmlCodeBase.xml";
        public static string XSLT_CHK_COUNT = Path.GetDirectoryName(Application.ExecutablePath) + "\\convert\\ChkCountChanel.xslt";
        public static string XSLT_TO_ROW = Path.GetDirectoryName(Application.ExecutablePath) + "\\convert\\xslToRow.xslt";
        public static string XML_ROW = Path.GetDirectoryName(Application.ExecutablePath) + "\\convert\\xmlRow.xml";
        public static string EXCEL_CODE_BASE = Path.GetDirectoryName(Application.ExecutablePath) + "\\office\\Codes_base.xlsx";
        public static string XSLT_CODE_BASE = Path.GetDirectoryName(Application.ExecutablePath) + "\\convert\\extractXlsxData.xslt";

        // Пути к файлам итогового преобразования
        public static string XSLT_RESULT = Path.GetDirectoryName(Application.ExecutablePath) + "\\convert\\xslt_result.xslt";
        public static string XML_ZIP_RESULT(TypeGetReport typeGet)
        {
            return Path.ChangeExtension(XML_RESULT(typeGet), ".zip");
        }
        public static string XML_RESULT(TypeGetReport typeGet)
        {
            if (typeGet == TypeGetReport.handle)
                return Path.GetDirectoryName(Application.ExecutablePath) + "\\resultFile\\" +
                    String.Format("{0}_{1}_{2}_{3}_PRUSSES1.xml", Common.SysProp.classValue, Common.SysProp.innValue,
                     Common.SysProp.officeDateReport.ToString("yyyyMMdd"), Common.SysProp.numberValue);
            else 
                return Path.GetDirectoryName(Application.ExecutablePath) + "\\resultFile\\" +
                    String.Format("{0}_{1}_{2}_{3}_PRUSSES1.xml", Common.SysProp.classValue, Common.SysProp.innValue,
                    DateTime.Now.AddDays(-1).ToString("yyyyMMdd"), Common.SysProp.numberValue);
        }

        // Путь к xml-файлу sql-запроса
        public static string ORA_FILE = Path.GetDirectoryName(Application.ExecutablePath) + "\\convert\\ora_result.xml";
        public static string ORA_XML = Path.GetDirectoryName(Application.ExecutablePath) + "\\convert\\ora_xml_result.xml";
        public static string ORA_XSLT = Path.GetDirectoryName(Application.ExecutablePath) + "\\convert\\ora_xslt.xslt";

        // Пути к файлам настроек
        public static string TNS_FILE = Path.GetDirectoryName(Application.ExecutablePath) + "\\tnsnames.ora";

        public static void SetTnsFile(string serverName, string baseName, string port, string connectionName)
        {
            try
            {
                string contentFile = String.Format(
                    "{0} = (DESCRIPTION = (ADDRESS = (PROTOCOL = TCP)(HOST = {1})(PORT = {2})) " +
                    "(CONNECT_DATA = (SERVER = DEDICATED) (SERVICE_NAME = {3})))", 
                    connectionName.Trim(), serverName.Trim(), port.Trim(), baseName.Trim());

                using (FileStream fileStream = new FileStream(
                        TNS_FILE,
                        FileMode.Create,
                        FileAccess.Write,
                        FileShare.None))
                {
                    using (StreamWriter streamWriter = new StreamWriter(fileStream))
                    {
                        StringBuilder strBuild = new StringBuilder(contentFile);
                        streamWriter.WriteLine(strBuild.ToString());
                    }
                }
            }
            catch (Exception err)
            {
                Log.ToLog(err.Message);
            }

        }
        public static void ToFile(string stringValue, string fileName)
        {
            try
            {
                using (FileStream fileStream = new FileStream(
                        fileName,
                        FileMode.Create ,
                        FileAccess.Write,
                        FileShare.None))
                {
                    using (StreamWriter streamWriter = new StreamWriter(fileStream))
                    {
                        streamWriter.WriteLine(stringValue);
                    }
                }
            }
            catch
            {
                return;
            }
        }
        public static double Round(double value, int digits)
        {
            double scale = Math.Pow(10.0, digits);
            double round = Math.Floor(Math.Abs(value) * scale + 0.5);
            return (Math.Sign(value) * round / scale);
        }
        public static XmlDocument ConvertXML(string xmlFile, string xsltFile)
        {
            XmlDocument doc = new XmlDocument();
            XslCompiledTransform xslt = new XslCompiledTransform();
            xslt.Load(xsltFile);

            using (XmlWriter writer = doc.CreateNavigator().AppendChild()) 
            {
                xslt.Transform(xmlFile, null, writer);
            }
            return doc;
        }
    }
    public static class Log
    {
        public static void ToLog(string data)
        {
            try
            {
                using (FileStream fileStream = new FileStream(
                        Path.ChangeExtension(Application.ExecutablePath, ".log"),
                        FileMode.Append,
                        FileAccess.Write,
                        FileShare.None))
                {
                    using (StreamWriter streamWriter = new StreamWriter(fileStream))
                    {
                        StringBuilder strBuild = new StringBuilder("");
                        if (String.IsNullOrEmpty(data))
                            strBuild.Append(" ");
                        else
                            strBuild.Append(DateTime.Now.ToString() + "\t");
                        streamWriter.WriteLine(strBuild.ToString() + data);
                    }
                }
            }
            catch
            {
                return;
            }
        }
        public static string PROP_FILE = Path.GetDirectoryName(Application.ExecutablePath) + "\\sysprop.dmp";
    }
    public sealed class AskSerialize
    {
        static private Stream NewFileStream(string fileName)
        {
            return File.Create(fileName);
        }
        static private Stream GetFileStream(string fileName)
        {
            if (File.Exists(fileName))
                return File.OpenRead(fileName);
            else
                return null;
        }
        static public void Serialize(string path, object value)
        {
            if (value == null)
                return;
            BinaryFormatter f = new BinaryFormatter();
            using (Stream nf = NewFileStream(path))
            {
                f.Serialize(nf, value);
            }
        }
        static public object Deserialize(string path)
        {
            BinaryFormatter f = new BinaryFormatter();
            using (Stream strm = GetFileStream(path))
            {
                if ((strm != null) && (strm.Length != 0))
                {
                    return f.Deserialize(strm);
                }
                else
                    return null;
            }
        }

        private AskSerialize() { }
    }

    public static class AwareMessageBox
    {
        public static DialogResult Show(IWin32Window owner, string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon, MessageBoxDefaultButton defaultButton, MessageBoxOptions options)
        {
            if (IsRightToLeft(owner))
            {
                options |= MessageBoxOptions.RtlReading |
                MessageBoxOptions.RightAlign;
            }
            return MessageBox.Show(owner, text, caption,
            buttons, icon, defaultButton, options);
        }
        private static bool IsRightToLeft(IWin32Window owner)
        {
            Control control = owner as Control;

            if (control != null)
            {
                return control.RightToLeft == RightToLeft.Yes;
            }

            // If no parent control is available, ask the CurrentUICulture
            // if we are running under right-to-left.
            return CultureInfo.CurrentUICulture.TextInfo.IsRightToLeft;
        }
    }

}
