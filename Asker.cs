using System.Linq;
using System;
using System.Text;
using System.Xml;
using System.Windows.Forms;
using System.IO;
using System.IO.Packaging;
using System.Collections.Generic;
using System.Xml.Xsl;
using System.Data;
using WFExceptions;

namespace Askme
{
    /// <summary>
    /// Основоной класс - формирует запрос к базе, проводит слияние дата сетов, отправляет данные по e-mail
    /// </summary>
    class Asker
    {
        public Asker(TypeGetReport typeGet, bool mailed, bool archiveMail)
        {
            this.typeGet = typeGet;
            this.mailed = mailed;
            this.archiveMail = archiveMail;
        }

        // Public method
        public string PrepareSql(ref bool result)
        {
            // Подготовка запроса для Oracle
            try
            {
                Log.ToLog("Начало формирования кода SQL-запроса");
                OfficeFiles officeFiles = new OfficeFiles(Common.EXCEL_CODE_BASE);
                officeFiles.CreateCodeBaseXML();
                this.listPoint = officeFiles.GetObjectIds();
                if (this.typeGet == TypeGetReport.handle)
                    return String.Format(
                        Common.SysProp.sqlText,
                        listPoint.Aggregate((curr, next)=>curr +", "+ next),
                        Convert.ToDateTime(Common.SysProp.officeDateReport).ToString("dd-MM-yyyy"));
                else 
                    return String.Format(
                        Common.SysProp.sqlText,
                        listPoint.Aggregate((curr, next) => curr + ", " + next),
                        DateTime.Now.AddDays(-1).ToString("dd-MM-yyyy"));
            }
            catch (Exception err)
            {
                WFException.HandleError(err);
                result = false;
                return "";
            }
            finally
            {
                Log.ToLog("Завершение формирования кода SQL-запроса");
            }
        }
        public void SendMail(ref bool result)
        {

            if (this.typeGet == TypeGetReport.handle && !mailed)
            {
                result = true;
                return;
            }

            try
            {
                Log.ToLog("Начало отправки сообщения на e-mail");

                Mailer mailer = new Mailer();
                mailer.LoadData();
                mailer.SendMail(this.typeGet, archiveMail);
            }
            catch (Exception err)
            {
                WFException.HandleError(err);
                result = false;
            }
            finally
            {
                Log.ToLog("Завершение отправки сообщения на e-mail");
            }
        }
        public void ExecSQL(string sqlQuery, ref bool result)
        {
            try
            {
                Log.ToLog("Начало выполения запроса к базе");

                OracleQuery oracleQuery = new OracleQuery(sqlQuery);
                oracleQuery.Query();

                double diff = 0;
                double currPureValue = 0;
                double currCheckValue = 0;
                double roundCheckValue = 0;

                int gr = 0;
                StringBuilder sb = new StringBuilder("<?xml version=\"1.0\" encoding=\"utf-8\"?>\n");
                sb.Append("<DocumentElement>\n");
                for (int i = 0; i < oracleQuery.DT.Rows.Count; i++)
                {
                    if (gr != Convert.ToInt32(oracleQuery.DT.Rows[i]["N_GR_TY"]))
                    {
                        diff = 0;
                        gr = Convert.ToInt32(oracleQuery.DT.Rows[i]["N_GR_TY"]);
                    }
                    currPureValue = Convert.ToDouble(oracleQuery.DT.Rows[i]["VAL"]);
                    currCheckValue = currPureValue + diff;
                    roundCheckValue = Common.Round(currCheckValue, 0);
                    diff = currCheckValue - roundCheckValue;
                    sb.Append(
                        String.Format("<oracleResult N_OB=\"{0}\" N_FID=\"{1}\" N_GR_TY=\"{2}\" DD_MM_YYYY=\"{3}\" N_INTER_RAS=\"{4}\" VAL=\"{5}\" MIN_0=\"{6}\" MIN_1=\"{7}\" />\n",
                        oracleQuery.DT.Rows[i]["N_OB"].ToString(),
                        oracleQuery.DT.Rows[i]["N_FID"].ToString(),
                        oracleQuery.DT.Rows[i]["N_GR_TY"].ToString(),
                        oracleQuery.DT.Rows[i]["DD_MM_YYYY"].ToString(),
                        oracleQuery.DT.Rows[i]["N_INTER_RAS"].ToString(),
                        roundCheckValue,
                        oracleQuery.DT.Rows[i]["MIN_0"].ToString(),
                        oracleQuery.DT.Rows[i]["MIN_1"].ToString())
                        );
                }
                sb.Append("</DocumentElement>");

                Common.ToFile(sb.ToString(), Common.ORA_XML);

                result = true;
            }
            catch (Exception err)
            {
                result = false;
                WFException.HandleError(err);
            }
            finally
            {
                Log.ToLog("Завершение выполения запроса к базе");
            }
            
        }
        public void PrepareXML(ref bool result)
        {
            try
            {
                Log.ToLog("Начало преобразования полученных данных в итоговый xml");

                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(Common.ORA_XML);

                XsltSettings settings = new XsltSettings(true, true);
                XslCompiledTransform xslTransSheet = new XslCompiledTransform();
                xslTransSheet.Load(Common.XSLT_RESULT, settings, new XmlUrlResolver());

                XsltArgumentList xsltArgList = new XsltArgumentList();
                xsltArgList.AddParam("class", "", Common.SysProp.classValue);
                xsltArgList.AddParam("version", "", Common.SysProp.versionValue);
                xsltArgList.AddParam("number", "", Common.SysProp.numberValue);
                xsltArgList.AddParam("timestamp", "", DateTime.Now.ToString("yyyyMMddHHmmss"));
                xsltArgList.AddParam("daylightsavingtime", "", Common.SysProp.dayLigthSavigTimeValue);
                xsltArgList.AddParam("day", "", Convert.ToDateTime(Common.SysProp.officeDateReport).ToString("yyyyMMdd"));
                xsltArgList.AddParam("name", "", Common.SysProp.nameValue);
                xsltArgList.AddParam("inn", "", Common.SysProp.innValue);
                xsltArgList.AddParam("name2", "", Common.SysProp.name2Value);
                xsltArgList.AddParam("inn3", "", Common.SysProp.inn3Value);

                using (FileStream fs = new FileStream(Common.XML_RESULT(this.typeGet), FileMode.Create))
                {
                    xslTransSheet.Transform(xmlDoc.CreateNavigator(), xsltArgList, fs);
                }

                // Проверка количества сформированных фидеров
                Log.ToLog("---------Начало проверки данных на количество получасовок (кратность 48)");
                CheckFiderCount();
                Log.ToLog("---------Окончание проверки данных на количество получасовок (кратность 48)");

                result = true;
            }
            catch (Exception err)
            {
                result = false;
                WFException.HandleError(err);
            }
            finally
            {
                Log.ToLog("Завершение преобразования полученных данных в итоговый xml");
            }
        }

        // Properties
        public TypeGetReport TypeGet { get { return this.typeGet; } }
        public bool Mailed { get { return this.mailed; } }

        // Private mathods
        private void CheckFiderCount()
        {
            XmlDocument doc = Common.ConvertXML(Common.XML_RESULT(typeGet), Common.XSLT_CHK_COUNT);
            if (doc.SelectNodes("//cnl").Count > 0)
            {
                Log.ToLog(doc. ToString());
                throw new Exception(CommonString.ErrorFiderCount);
            }

            //XmlDocument result = new XmlDocument();
            //result.Load(Common.ORA_XML);

            //foreach (string item in this.listPoint)
            //    if (result.SelectNodes("//oracleResult[@N_OB=" + item + "]").Count % 48 > 0)
            //         throw new Exception(String.Format(CommonString.ErrorFiderCount, item));
        
        }

        //Fields
        private TypeGetReport typeGet;
        private bool mailed;
        private bool archiveMail;
        private List<string> listPoint;

    }

    public enum TypeGetReport {handle, auto};
}
