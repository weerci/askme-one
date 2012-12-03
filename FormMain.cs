﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Devart.Data.Oracle;
using System.Globalization;
using System.Data.Odbc;
using System.Threading;
using WFExceptions;

namespace Askme
{
    public partial class FormMain : Form
    {
        // Constructors
        public FormMain()
        {
            InitializeComponent();
        }

        // Events
        private void FormMain_Load(object sender, EventArgs e)
        {
            // Форма скрывается
            this.Visible = false;
            ShowInTaskbar = false;

            if (Common.SysProp == null)
                throw new WFException(ErrType.Critical, ErrorCode.NOT_INI_SYSPROP);

            Common.SysProp = (SysProp)AskSerialize.Deserialize(Log.PROP_FILE);
            LoadFromSaveFile();

            // Запуск сервера выполнения запросов
            if (Common.SysProp.ImmediateStart)
                StartTimer(); 
            

            IsChanged = false;
            EnabledControl();

            Log.ToLog(CommonString.ApplicationStart);
        }
        private void toolStripMenuItemProperties_Click(object sender, EventArgs e)
        {
            this.Show();
            this.ShowInTaskbar = true;
        }
        private void toolStripMenuItemClose_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        private void FormMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                this.Visible = false;
                this.ShowInTaskbar = false;
            }
        }
        private void timerMain_Tick(object sender, EventArgs e)
        {
            if (this.inRun)
                return;

            bwSQL.RunWorkerAsync(TypeGetReport.auto);
        }
        private void buttonSave_Click(object sender, EventArgs e)
        {
            try
            {
                this.Cursor = Cursors.WaitCursor;

                SysProp sysProp = new SysProp();

                #region Сохранение настроек
                // Сохранение значений настроек таймера
                sysProp.FromTime = mtbFrom.Text;
                sysProp.ToTime = mtbTo.Text;
                sysProp.Interval = nudInterval.Value;
                sysProp.ImmediateStart = cbImmediateStart.Checked;

                // Сохранение настроек e-mail
                sysProp.MessageFrom = tbFrom.Text;
                sysProp.MessageTo = tbTo.Text;
                sysProp.MessageSubject = tbSubject.Text;
                sysProp.MessageBody = tbBody.Text;
                sysProp.SmtpServer = tbSmtp.Text;
                sysProp.Port = Convert.ToInt32(tbPort.Text);
                sysProp.login = tbLogin.Text;
                sysProp.password = tbPassword.Text;

                // Сохранение настроек соединения с базой данных
                sysProp.oracleLogin = tbOracleLogin.Text;
                sysProp.oraclePassword = tbOraclelPassword.Text;
                sysProp.oracleAliase = tbOracleAlias.Text;

                // Сохранение настроек файлов офиса
                sysProp.officeDateReport = dtpDataReport.Value;

                // Сохранение переменных
                sysProp.classValue = tbClass.Text;
                sysProp.versionValue = tbVersion.Text;
                sysProp.numberValue = tbNumber.Text;
                sysProp.dayLigthSavigTimeValue = tbDayLight.Text;
                sysProp.nameValue = tbName.Text;
                sysProp.innValue = tbInn.Text;
                sysProp.name2Value = tbName2.Text;
                sysProp.inn3Value = tbInn3.Text;
                sysProp.statusValue = tbStatus.Text;
                sysProp.extStatusValue = tbExtendStatus.Text;
                sysProp.param1Value = tbParam1.Text;
                sysProp.mailedOnHandle = cbMailed.Checked;
                sysProp.archiveMail = cbArchiveMail.Checked;

                // Сохранение текста запроса
                sysProp.sqlText = richTextBoxSQL.Text;

                #endregion

                AskSerialize.Serialize(Log.PROP_FILE, sysProp);
                Common.SysProp = (SysProp)AskSerialize.Deserialize(Log.PROP_FILE);

                IsChanged = false;
                EnabledControl();
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }
        private void ValueChanged(object sender, EventArgs e)
        {
            IsChanged = true;
            EnabledControl();
        }
        private void btnRun_Click(object sender, EventArgs e)
        {
            StartTimer();
        }
        private void btnConnectTest_Click(object sender, EventArgs e)
        {
            try
            {
                this.Cursor = Cursors.WaitCursor;

                using (OracleConnection connection = new OracleConnection(
                    String.Format("User Id={0}; Password={1}; Data Source={2}",
                    Common.SysProp.oracleLogin, Common.SysProp.oraclePassword, Common.SysProp.oracleAliase)))
                {
                    connection.Open();
                    connection.Close();
                }
                MessageBox.Show(CommonString.TestConnectionOk);
            }
            catch (Exception err)
            {
                MessageBox.Show(String.Format(CommonString.TestConnectionFailed, err.Message));
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }
        private void btnLogFile_Click(object sender, EventArgs e)
        {
            try
            {
                this.Cursor = Cursors.WaitCursor;
                System.Diagnostics.Process.Start("notepad.exe", Path.ChangeExtension(Application.ExecutablePath, ".log"));
            }
            catch (Exception err)
            {
                Log.ToLog(err.Message);
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }
        private void notifyIconMain_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            toolStripMenuItemProperties_Click(null, null);
        }
        private void btnCancel_Click(object sender, EventArgs e)
        {
            try
            {
                this.Cursor = Cursors.WaitCursor;
                Common.SysProp = (SysProp)AskSerialize.Deserialize(Log.PROP_FILE);
                LoadFromSaveFile();
                IsChanged = false;
                EnabledControl();
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }
        private void StartTimer()
        {
            if (timerMain.Enabled)
            {
                timerMain.Stop();
                btnRun.Text = CommonString.buttonStart;
                lbStartServer.Text = CommonString.labelStop;
            }
            else 
            {
                timerMain.Stop();
                btnRun.Text = CommonString.buttonStop;
                lbStartServer.Text = CommonString.labelStart;
                timerMain.Interval = Convert.ToInt32(Common.SysProp.Interval) * 60000;
                timerMain.Start();
            }
        }
        private void btnGetReport_Click(object sender, EventArgs e)
        {
            if (this.IsChanged)
            {
                if (AwareMessageBox.Show(
                            this,
                            CommonString.ConfirmUnsafeRun,
                            Application.ProductName,
                            MessageBoxButtons.OKCancel,
                            MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1,
                            (MessageBoxOptions)0) == DialogResult.OK)
                {
                    buttonSave_Click(null, null);
                    btnGetReport.Enabled = false;
                    bwSQL.RunWorkerAsync(TypeGetReport.handle);
                }
            }
            else
            {
                buttonSave_Click(null, null);
                btnGetReport.Enabled = false;
                bwSQL.RunWorkerAsync(TypeGetReport.handle);
            }
        }
        private void btnViewReport_Click(object sender, EventArgs e)
        {
            try
            {
                this.Cursor = Cursors.WaitCursor;
                System.Diagnostics.Process.Start("notepad.exe", Common.XML_RESULT(TypeGetReport.handle));
            }
            catch (Exception err)
            {
                Log.ToLog(err.Message);
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }
        private void tsmiFile_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem item = sender as ToolStripMenuItem;
            if (item != null)
            {
                switch (item.Name)
                {
                    case "tsmiConst":
                        tabControlProperties.SelectedIndex = 0;
                        break;
                    case "tsmiQuery":
                        tabControlProperties.SelectedIndex = 1;
                        break;
                    case "tsmiConnection":
                        tabControlProperties.SelectedIndex = 2;
                        break;
                    case "tsmiProperty":
                        tabControlProperties.SelectedIndex = 3;
                        break;
                    case "tsmiExit":
                        toolStripMenuItemClose_Click(null, null);
                        break;
                    default:
                        break;
                }
            }
        }
        private void tsmiOpenHelp_Click(object sender, EventArgs e)
        {
            tsmHelp_Click(null, null);
        }
        private void tsmiAbout_Click(object sender, EventArgs e)
        {
            new AboutBox().ShowDialog();
        }
        private void bwSQL_DoWork(object sender, DoWorkEventArgs e)
        {
            bool result = true;
            this.inRun = true;
            Asker asker = new Asker((TypeGetReport)e.Argument, cbMailed.Checked, cbArchiveMail.Checked);

            // Проверка условий выполнения запроса
            if (asker.TypeGet == TypeGetReport.handle ||
                (new CheckSendCondition().CheckCondition() && asker.TypeGet == TypeGetReport.auto))
            {
                try
                {
                    Log.ToLog("");
                    Log.ToLog("****************Формирование нового отчета**************************");
                    
                    // Подготовка запроса
                    bwSQL.ReportProgress(1);
                    string sqlText = asker.PrepareSql(ref result);
                    if (!result)
                        throw new Exception(CommonString.ErrorOnPrepareSQL);
                    Log.ToLog(sqlText);
                    
                    // Выполение запроса
                    bwSQL.ReportProgress(2);
                    asker.ExecSQL(sqlText, ref result);
                    if (!result)
                        throw new Exception(CommonString.ErrorOnExecSQL);

                    // Конвертация в XML
                    bwSQL.ReportProgress(3);
                    asker.PrepareXML(ref result);
                    if (!result)
                        throw new Exception(CommonString.ErrorOnConvertXML);

                    // Отправка сообщения
                    bwSQL.ReportProgress(4);
                    asker.SendMail(ref result);
                    if (!result)
                        throw new Exception(CommonString.ErrorOnMail);

                    Common.SysProp.LastSendedDate = DateTime.Now.ToShortDateString();
                    Log.ToLog("****************Отчет успешно сформирован***************************");
                }
                catch (Exception err)
                {
                    this.inRun = false;
                    Log.ToLog(err.Message);
                    Log.ToLog("**********При формировании отчета возникли ошибки*******************");
                }
            }
        }
        private void bwSQL_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            this.inRun = false;
            lbState.Text = CommonString.stateWait;
            btnGetReport.Enabled = true;
        }
        private void bwSQL_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            switch (e.ProgressPercentage)
	        {
                case 1:
                    lbState.Text = CommonString.statePrepareSQL;
                    break;
                case 2:
                    lbState.Text = CommonString.stateGetSQL;
                    break;
                case 3:
                    lbState.Text = CommonString.stateConvertXML;
                    break;
                case 4:
                    lbState.Text = CommonString.stateSendMail;
                    break;
                default:
                    lbState.Text = CommonString.stateWait;
                    break;
	        }
        }
        private void btnCodeBase_Click(object sender, EventArgs e)
        {
            try
            {
                this.Cursor = Cursors.WaitCursor;
                System.Diagnostics.Process.Start("excel", String.Format("\"{0}\"", Common.EXCEL_CODE_BASE));
            }
            catch (Exception err)
            {
                Log.ToLog(err.Message);
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }
        private void tsmHelp_Click(object sender, EventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(Common.HELP_PATH);
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message);
            }
        }

        // Private methods
        private void EnabledControl()
        {
            btnSave.Enabled = IsChanged;
            btnCancel.Enabled = IsChanged;
        }
        private void LoadFromSaveFile()
        {
            // Load page properties
            mtbFrom.Text = Common.SysProp.FromTime;
            mtbTo.Text = Common.SysProp.ToTime;
            nudInterval.Value = Common.SysProp.Interval;
            cbImmediateStart.Checked = Common.SysProp.ImmediateStart;

            tbFrom.Text = Common.SysProp.MessageFrom;
            tbTo.Text = Common.SysProp.MessageTo;
            tbSubject.Text = Common.SysProp.MessageSubject;
            tbBody.Text = Common.SysProp.MessageBody;
            tbSmtp.Text = Common.SysProp.SmtpServer;
            tbPort.Text = Common.SysProp.Port.ToString();
            tbLogin.Text = Common.SysProp.login;
            tbPassword.Text = Common.SysProp.password;

            // Load oracle connection
            tbOracleLogin.Text = Common.SysProp.oracleLogin;
            tbOraclelPassword.Text = Common.SysProp.oraclePassword;
            tbOracleAlias.Text = Common.SysProp.oracleAliase;

            // Load constants
            tbClass.Text = Common.SysProp.classValue;
            tbVersion.Text = Common.SysProp.versionValue;
            tbNumber.Text = Common.SysProp.numberValue;
            tbDayLight.Text = Common.SysProp.dayLigthSavigTimeValue;
            tbName.Text = Common.SysProp.nameValue;
            tbInn.Text = Common.SysProp.innValue;
            tbName2.Text = Common.SysProp.name2Value;
            tbInn3.Text = Common.SysProp.inn3Value;
            tbStatus.Text = Common.SysProp.statusValue;
            tbExtendStatus.Text = Common.SysProp.extStatusValue;
            tbParam1.Text = Common.SysProp.param1Value;
            dtpDataReport.Value = Common.SysProp.officeDateReport;
            cbMailed.Checked = Common.SysProp.mailedOnHandle;
            cbArchiveMail.Checked = Common.SysProp.archiveMail;

            // Load sql
            richTextBoxSQL.Text = Common.SysProp.sqlText;

        }

        // Fields
        private bool IsChanged;
        private volatile bool inRun;

   }
}
