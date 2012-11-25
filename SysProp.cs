using System;
using System.Collections.Generic;
using System.Text;

namespace Askme
{
    [Serializable]
    public class SysProp
    {
        // Настройки таймера
        public string FromTime;
        public string ToTime;
        public decimal Interval;
        public bool ImmediateStart;

        // Настройки e-mail
        public string MessageFrom;
        public string MessageTo;
        public string MessageSubject;
        public string MessageBody;
        public string SmtpServer;
        public int Port;
        public string login;
        public string password;

        // Настройки соединения с базой
        public string oracleLogin;
        public string oraclePassword;
        public string oracleAliase;

        // Настройки пути файлов офиса
        public string officeCodeBase;
        public DateTime officeDateReport;

        // Настройка переменных запроса
        public string classValue;
        public string versionValue;
        public string numberValue;
        public string dayLigthSavigTimeValue;
        public string nameValue;
        public string innValue;
        public string name2Value;
        public string inn3Value;
        public string statusValue;
        public string extStatusValue;
        public string param1Value;
        public string nameSendFile;
        public bool mailedOnHandle;
        public bool archiveMail;

        // Текст запроса
        public string sqlText;

        public string LastSendedDate = DateTime.Now.AddDays(-1).ToShortDateString();
    }
}
