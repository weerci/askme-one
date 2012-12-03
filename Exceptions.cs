using System;
using System.IO;
using System.Collections;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Text;
using System.Globalization;
using System.Runtime.Serialization;
using Askme;

namespace WFExceptions
{
    /// <summary>
    /// Типы сообщений об ошибке: 
    /// <list type="bullet">
    /// <item><description>Message - сообещние о предвиденной ошибке в программе</description></item>
    /// <item><description>Assert - сообщение о невыполненных краевых условиях </description></item>
    /// <item><description>Critical - неперехваченная ошибка</description></item>
    /// </list>    
    /// </summary>
    public enum ErrType { None, Message, Assert, Critical };
    /// <summary>
    /// Класс реализует обработку исключительных ситуаций
    /// </summary>
    public class WFException : Exception
    {
        // Constructors
        public WFException(): base() {}
        public WFException(ErrType errType, string message) : base(message)
        {
            _message = message;
            _errType = errType;
        }
        public WFException(ErrType errType, string message, Exception error) : base(message, error)
        {
            _exception = error;
            _message = message;
            _errType = errType;
        }
        protected WFException(SerializationInfo info, StreamingContext context) : base(info, context) { }

        // Properties
        public string Msg
        {
            get
            {
                return _message;
            }
        }
        public Exception Exp
        {
            get
            {
                return _exception;
            }
        }
        public ErrType ErrorType
        {
            get
            {
                return _errType;
            }
        }

        // Private methods
        private static void ToLog(ErrType logType, string data)
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
                        StringBuilder _strBuild = new StringBuilder("");
                        switch (logType)
                        {
                            case ErrType.Message:
                                _strBuild.Append("\r***MESSAGE***\r" + DateTime.Now.ToString() + "\t");
                                break;
                            case ErrType.Critical:
                                _strBuild.Append("\r***CRITICAL***\r" + DateTime.Now.ToString() + "\t");
                                break;
                            case ErrType.Assert:
                                _strBuild.Append("\r***ASSERT***\r" + DateTime.Now.ToString() + "\t");
                                break;
                            case ErrType.None:
                            default:
                                break;
                        }
                        streamWriter.WriteLine(_strBuild.ToString() + data);
                    }
                }
            }
            catch
            {
                return;
            }
        }
        private static void ShowError(string data)
        {
            AwareMessageBox.Show(null, data, Application.ProductName,
                MessageBoxButtons.OK, MessageBoxIcon.Error, 
                MessageBoxDefaultButton.Button1, (MessageBoxOptions)0);
        }
        private static void LevelExceptionToLog(Exception e, int exceptionLevel, StringBuilder logString)
        {
            exceptionLevel++;
            object[] arrayErr = { exceptionLevel, e.GetType().Name.ToString(), e.HelpLink, e.Message, e.Source, e.StackTrace, e.TargetSite };
            logString.AppendLine(String.Format("***Exception Level: {0}", exceptionLevel));
            logString.AppendLine(String.Format("***ExceptionType: {0}",  e.GetType().Name.ToString()));
            logString.AppendLine(String.Format("***HelpLine: {0}", e.HelpLink));
            logString.AppendLine(String.Format("***Message: {0}", e.Message));
            logString.AppendLine(String.Format("***Source: {0}", e.Source));
            logString.AppendLine(String.Format("***StackTrace: {0}", e.StackTrace));
            logString.AppendLine(String.Format("***TargetSite: {0}", e.TargetSite));
            logString.Append(String.Format("***Data: {0}", e.Data));
            foreach (DictionaryEntry de in e.Data)
            {
                logString.AppendLine(String.Format(CultureInfo.InvariantCulture, "  {0} : {1}", de.Key, de.Value));
            }
            ToLog(ErrType.None, logString.ToString());

            Exception ie = e.InnerException;
            while (ie != null)
            {
                LevelExceptionToLog(ie, exceptionLevel++, logString);
                if (exceptionLevel > 1)
                    ie = ie.InnerException;
                else
                    ie = null;
            }
            exceptionLevel--;
        }
        private static void EmptyErrorStackToLog()
        {
            string currentStackTrace = System.Environment.StackTrace;
            string firstStackTraceCall = "System.Environment.get_StackTrace( )";

            int posOfStackTraceCall = currentStackTrace.IndexOf(firstStackTraceCall, StringComparison.Ordinal);
            string finalStackTrace = currentStackTrace.Substring(posOfStackTraceCall + firstStackTraceCall.Length);
            MatchCollection methodCallMatches = Regex.Matches(finalStackTrace, @"\sat\s.*(\sin\s.*\:line\s\d*)?");
            StringBuilder log = new StringBuilder("-------------");
            foreach (Match m in methodCallMatches)
            {
                log.Append(m.Value + System.Environment.NewLine + "-------------");
            }
        }
        public static void HandleError(UnhandledExceptionEventArgs e)
        {
            EmptyErrorStackToLog();
        }
        public static void HandleError(Exception error)
        {
            bool IsUriException = error.GetType() == typeof(WFException);
            StringBuilder stringBuilder = new StringBuilder();

            if (IsUriException)
            {
                WFException wfErr = (WFException)error;
                switch (wfErr.ErrorType)
                {
                    case ErrType.Message:
                        ShowError(wfErr.Msg);
                        if (wfErr.Exp != null)
                        {
                            ToLog(ErrType.Message, wfErr.Msg);
                            LevelExceptionToLog(wfErr, 0, stringBuilder);
                            ToLog(ErrType.Message, stringBuilder.ToString());
                        }
                        break;
                    case ErrType.Critical:
                        ToLog(ErrType.Critical, wfErr.Msg);
                        stringBuilder.AppendLine("***CRITICAL");
                        LevelExceptionToLog(wfErr, 0, stringBuilder);
                        ToLog(ErrType.Critical, stringBuilder.ToString());
                        break;
                    case ErrType.Assert:
                        ToLog(ErrType.Assert, wfErr.Msg);
                        stringBuilder.AppendLine("***ASSERT");
                        LevelExceptionToLog(wfErr, 0, stringBuilder);
                        ToLog(ErrType.Assert, stringBuilder.ToString());
                        break;
                    case ErrType.None:
                    default:
                        break;
                }
            }
            else
            {
                ToLog(ErrType.Critical, error.Message);
                LevelExceptionToLog(error, 0, stringBuilder);
//                ToLog(ErrType.Critical, stringBuilder.ToString());
            }
        }

        public WFException(string message) : base(message) { }
        public WFException(string message, Exception error) : base(message, error) { }

        // Fields
        private Exception _exception;
        private ErrType _errType;
        private string _message;
    }


}
