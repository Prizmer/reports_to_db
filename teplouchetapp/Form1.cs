using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.IO;
using System.IO.Ports;
using ExcelLibrary.SpreadSheet;
using System.Configuration;
using System.Threading;
using System.Diagnostics;
//using System.Configuration.Assemblies;

using Npgsql;


namespace teplouchetapp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            _Form1 = this;

            this.Text = FORM_TEXT_DEFAULT;

            InProgress = false;
            DemoMode = false;
            InputDataReady = false;

        }

        Form1 _Form1;

        //при опросе или тесте связи
        bool bInProcess = false;
        public bool InProgress
        {
            get { return bInProcess; }
            set
            {
                bInProcess = value;

                if (bInProcess)
                {
                    toolStripProgressBar1.Value = 0;

                    buttonPoll.Enabled = false;
                    buttonImport.Enabled = false;           
                    buttonExport.Enabled = false;
                    buttonStop.Enabled = true;


                    this.Text += FORM_TEXT_INPROCESS;
                }
                else
                {
                    buttonPoll.Enabled = true;
                    buttonImport.Enabled = true;
                    buttonExport.Enabled = true;
                    buttonStop.Enabled = false;
                    dgv1.Enabled = true;

                    this.Text = this.Text.Replace(FORM_TEXT_INPROCESS, String.Empty);
                }
            }
        }

        //Демонстрационный режим - отключает сервисные сообщения
        bool bDemoMode = false;
        public bool DemoMode
        {
            get { return bDemoMode; }
            set
            {
                bDemoMode = value;

                if (bDemoMode)
                {
                    this.Text = this.Text.Replace(FORM_TEXT_DEMO_OFF, String.Empty);
                }
                else
                {
                    this.Text += FORM_TEXT_DEMO_OFF;
                }
            }

        }

        bool bInputDataReady = false;
        public bool InputDataReady
        {
            get { return bInputDataReady; }
            set
            {
                bInputDataReady = value;

                if (!bInputDataReady)
                {
                    toolStripProgressBar1.Value = 0;
                    buttonPoll.Enabled = false;
                    buttonImport.Enabled = true;
                    buttonExport.Enabled = false;

                    buttonStop.Enabled = false;

                }
                else
                {
                    buttonPoll.Enabled = true;
                    buttonImport.Enabled = true;
                    buttonExport.Enabled = true;
                    buttonStop.Enabled = false;
                }
            }
        }

        #region Строковые постоянные 

            const string METER_IS_ONLINE = "ОК";
            const string METER_IS_OFFLINE = "Нет связи";
            const string METER_WAIT = "Ждите";
            const string REPEAT_REQUEST = "Повтор";
            const string METER_IS_BUSY = "Занят";

            const string FORM_TEXT_DEFAULT = "ПИ - программа записи отчетов в БД";
            const string FORM_TEXT_DEMO_OFF = "";
            const string FORM_TEXT_DEV_ON = " - режим разработчика";

            const string FORM_TEXT_INPROCESS = " - чтение данных";

        #endregion


        const string DOUBLE_STRING_FORMATER = "0.#######";
        private string m_connection_string; //= "Server=localhost;Port=5432;User Id=postgres;Password=;Database=prizmer;Pooling=true;MinPoolSize=1;MaxPoolSize=1000;";
        private Npgsql.NpgsqlConnection m_pg_con = null;

        public ConnectionState Open(String ConnectionString)
        {
            m_connection_string = "";
            m_pg_con = new Npgsql.NpgsqlConnection(ConnectionString);

            try
            {
                m_pg_con.Open();
            }
            catch (Exception ex)
            {
                return ConnectionState.Broken;
            }

            return m_pg_con.State;
        }

        public void Close()
        {
            if (m_pg_con != null)
            {
                m_pg_con.Close();
                m_pg_con = null;
            }
        }
        private bool GetMeterAddressBySerial(string sn, out string a)
        {
            a = "";

            string query = "select address from meters where (factory_number_manual='"+ sn + "' OR factory_number_readed='"+ sn +"')";
            NpgsqlCommand command = new NpgsqlCommand(query, m_pg_con);
            NpgsqlDataReader dr = null;
            List<object> results = new List<object>();

            try
            {
                dr = command.ExecuteReader();

                if (dr.HasRows)
                {
                    while (dr.Read())
                    {
                        results.Add((Object)dr["address"]);
                        WriteToLog("address: " + dr["address"]);
                    }
                }
            }
            catch (Exception e)
            {

            }
            finally
            {
                if (dr != null)
                {
                    if (!dr.IsClosed)
                    {
                        dr.Close();
                    }
                }
            }

            if (results.Count == 0)
                return false;

            a = results[0].ToString();
            return true;
        }
        private bool GetTakenParamIdByGuidAndMetersAddress(string paramGuid, int mAddress, out uint takenPId)
        {
            takenPId = 0;

            string query = "SELECT taken_params.id FROM public.taken_params, public.meters, public.params, public.types_params " +
                "WHERE taken_params.guid_meters = meters.guid AND taken_params.guid_params = params.guid AND params.guid_types_params = types_params.guid AND types_params.type = 4 AND " +
                "params.guid_names_params = '" + paramGuid + "' AND meters.address=" + mAddress;

            NpgsqlCommand command = new NpgsqlCommand(query, m_pg_con);
            NpgsqlDataReader dr = null;
            List<object> results = new List<object>();

            try
            {
                dr = command.ExecuteReader();

                if (dr.HasRows)
                {
                    while (dr.Read())
                    {
                        results.Add((Object)dr["id"]);
                    }
                }
            }
            catch (Exception e)
            {

            }
            finally
            {
                if (dr != null)
                {
                    if (!dr.IsClosed)
                    {
                        dr.Close();
                    }
                }
            }

            if (results.Count == 0)
                return false;

            takenPId = uint.Parse(results[0].ToString());
            return true;


        }
        public struct Value
        {
            public DateTime dt;
            public float value;
            public Boolean status;
            public UInt32 id_taken_params;
        }
        private int AddVariousValue(Value value)
        {

            string query = "INSERT INTO various_values (date, time, value, status, id_taken_params) " +
                "VALUES (" +
                "'" + value.dt.ToShortDateString() + "', " +
                "'" + value.dt.ToShortTimeString() + "', " +
                value.value.ToString(DOUBLE_STRING_FORMATER).Replace(',', '.') + ", " +
                value.status.ToString() + ", " +
                value.id_taken_params.ToString() +
                ")";


            NpgsqlCommand command = new NpgsqlCommand(query, m_pg_con);

            try
            {
                return command.ExecuteNonQuery();

            } 
            catch
            {
                return -1;
            }
        }

        const string AP_PNAME_GUID = "9ad9b931-fe2b-463d-b47f-f0a471279313";
        const string AM_PNAME_GUID = "087475af-2791-48fe-87fb-f5e883c97528";
        const string RP_PNAME_GUID = "475ac5ee-3ddd-4311-a0fb-d4bf531cbafd";
        const string RM_PNAME_GUID = "16462cac-a9cd-48df-b914-bff0cb9c5d94";



        //изначально ни один процесс не выполняется, все остановлены
        volatile bool doStopProcess = false;
        bool bPollOnlyOffline = false;

        //default settings for input *.xls file
        int flatNumberColumnIndex = 0;
        int factoryNumberColumnIndex = 1;
        int firstRowIndex = 1;


        private bool setXlsParser()
        {
            try
            {
                flatNumberColumnIndex = int.Parse(ConfigurationSettings.AppSettings["flatColumn"]) - 1;
                factoryNumberColumnIndex = int.Parse(ConfigurationSettings.AppSettings["factoryColumn"]) - 1;
                firstRowIndex = int.Parse(ConfigurationSettings.AppSettings["firstRow"]) - 1;
                m_connection_string = ConfigurationSettings.AppSettings["constr"];
                return true;
            }
            catch (Exception ex)
            {
                WriteToStatus("Ошибка разбора блока \"Настройка парсера\" в файле конфигурации: " + ex.Message);
                return false;
            }

        }

        private void WriteToStatus(string str)
        {
            MessageBox.Show(str, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void Form1_Load(object sender, EventArgs e)
        {           
            //setting up dialogs
            ofd1.Filter = "Excel files (*.xls) | *.xls";
            sfd1.Filter = ofd1.Filter;
            ofd1.FileName = "FactoryNumbersTable";
            sfd1.FileName = ofd1.FileName;

            if (!setXlsParser()) return;

            meterPinged += new EventHandler(Form1_meterPinged);
            pollingEnd += new EventHandler(Form1_pollingEnd);

        }


        DataTable dt = new DataTable("meters");
        public string worksheetName = "Лист1";

        //список, хранящий номера параметров в перечислении Params драйвера
        //целесообразно его сделать здесь, так как кол-во считываемых значений зависит от кол-ва колонок
        List<int> paramCodes = null;
        private void createMainTable(ref DataTable dt)
        {
            paramCodes = new List<int>();

            //creating columns for internal data table
            DataColumn column = dt.Columns.Add();
            column.DataType = typeof(string);
            column.Caption = "№";
            column.ColumnName = "сolNum";
     
            column = dt.Columns.Add();
            column.DataType = typeof(string);
            column.Caption = "A+";
            column.ColumnName = "colAp";

            column = dt.Columns.Add();
            column.DataType = typeof(string);
            column.Caption = "A-";
            column.ColumnName = "colAm";

            column = dt.Columns.Add();
            column.DataType = typeof(string);
            column.Caption = "R+";
            column.ColumnName = "ColRp";

            column = dt.Columns.Add();
            column.DataType = typeof(string);
            column.Caption = "R-  ";
            column.ColumnName = "colRm";

            column = dt.Columns.Add();
            column.DataType = typeof(string);
            column.Caption = "Время";
            column.ColumnName = "colTime";


            column = dt.Columns.Add();
            column.DataType = typeof(string);
            column.Caption = "Дата";
            column.ColumnName = "colDate";

            //paramCodes.Add(2);

            DataRow captionRow = dt.NewRow();
            for (int i = 0; i < dt.Columns.Count; i++)
                captionRow[i] = dt.Columns[i].Caption;
            dt.Rows.Add(captionRow);

        }

        string meterNumberFromCell = "";
        string meterAddressFromCell = "";
        private void loadXlsFile(float coeff = 1f)
        {
            doStopProcess = false;
            buttonStop.Enabled = true;

            dt = new DataTable();
            createMainTable(ref dt);
                       
            string fileName = ofd1.FileName;
            FileInfo fi = new FileInfo(fileName);
        
            string meterNumberFromFilename = fi.Name.Split('_')[0].Split('.')[0];

            Workbook book = Workbook.Load(fileName);
           
            int rowsInFile = 0;
            for (int i = 0; i < book.Worksheets.Count; i++)
                rowsInFile += book.Worksheets[i].Cells.LastRowIndex - firstRowIndex;

            //setting up progress bar
            toolStripProgressBar1.Minimum = 0;
            toolStripProgressBar1.Maximum = rowsInFile;
            toolStripProgressBar1.Step = 1;

            int worksheetsCnt = 1;

            //filling internal data table with *.xls file data according to *.config file
            for (int i = 0; i < worksheetsCnt; i++)
            {
                Worksheet sheet = book.Worksheets[i];
                meterNumberFromCell = sheet.Cells[3, 1].Value.ToString();
                meterAddressFromCell = sheet.Cells[4, 1].Value.ToString();

                if (meterNumberFromCell  != meterNumberFromFilename)
                {
                    MessageBox.Show(String.Format("Несовпадение номеров. Имя файла: {0}; Ячейка: {1};", meterNumberFromFilename, meterNumberFromCell), "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                double[] sumHalfs = new double[4];

                for (int rowIndex = firstRowIndex; rowIndex <= sheet.Cells.LastRowIndex; rowIndex++)
                {
                    if (doStopProcess)
                    {
                        buttonStop.Enabled = false;
                        return;
                    }

                    Row row_l = sheet.Cells.GetRow(rowIndex);
                    DataRow dataRow = dt.NewRow();

                    object oNumber = row_l.GetCell(0).Value;
                    int iNumber = 0;

                    if (oNumber != null && int.TryParse(oNumber.ToString(), out iNumber))
                    {
                        dataRow[0] = iNumber;

                        float tmpVal = 0;

                        dataRow[1] = row_l.GetCell(1).Value;
                        string tmpValStrrrr = dataRow[1].ToString();
                        float.TryParse(dataRow[1].ToString().Replace(',','.'), System.Globalization.NumberStyles.Any,
                            System.Globalization.CultureInfo.InvariantCulture, out tmpVal);
                        sumHalfs[0] += tmpVal * coeff;

                        dataRow[2] = row_l.GetCell(2).Value;
                        float.TryParse(dataRow[2].ToString().Replace(',', '.'), System.Globalization.NumberStyles.Any,
    System.Globalization.CultureInfo.InvariantCulture, out tmpVal);
                        sumHalfs[1] += tmpVal * coeff;

                        dataRow[3] = row_l.GetCell(3).Value;
                        float.TryParse(dataRow[3].ToString().Replace(',', '.'), System.Globalization.NumberStyles.Any,
System.Globalization.CultureInfo.InvariantCulture, out tmpVal);
                        sumHalfs[2] += tmpVal * coeff;


                        dataRow[4] = row_l.GetCell(4).Value;
                        float.TryParse(dataRow[4].ToString().Replace(',', '.'), System.Globalization.NumberStyles.Any,
System.Globalization.CultureInfo.InvariantCulture, out tmpVal);
                        sumHalfs[3] += tmpVal * coeff;


                        string tmps = row_l.GetCell(5).Value.ToString();


                       

                        dataRow[5] = row_l.GetCell(5).Value;
                        dataRow[6] = row_l.GetCell(6).Value;

                        incrProgressBar();
                    }
                    else
                    {
                        incrProgressBar();
                        continue;
                    }



                    dt.Rows.Add(dataRow);
                }

                richTextBox1.Clear();

                richTextBox1.Text += "Сумма A+: " + sumHalfs[0].ToString(DOUBLE_STRING_FORMATER) + ";\n";
                richTextBox1.Text += "Сумма A-: " + sumHalfs[1].ToString(DOUBLE_STRING_FORMATER) + ";\n";
                richTextBox1.Text += "Сумма R+: " + sumHalfs[2].ToString(DOUBLE_STRING_FORMATER) + ";\n";
                richTextBox1.Text += "Сумма R-: " + sumHalfs[3].ToString(DOUBLE_STRING_FORMATER) + ";\n\n";
            }


            dgv1.DataSource = dt;

            toolStripProgressBar1.Value = 0;
            toolStripProgressBar1.Maximum = dt.Rows.Count - 1;
            toolStripStatusLabel1.Text = String.Format("({0}/{1})", toolStripProgressBar1.Value, toolStripProgressBar1.Maximum);

            InputDataReady = true;
        }
        private void buttonImport_Click(object sender, EventArgs e)
        {
            if (ofd1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                loadXlsFile();
        }

        private void buttonExport_Click(object sender, EventArgs e)
        {
            if (sfd1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                //create new xls file
                string file = sfd1.FileName;
                Workbook workbook = new Workbook();
                Worksheet worksheet = new Worksheet(worksheetName);

                //office 2010 will not open file if there is less than 100 cells
                for (int i = 0; i < 100; i++)
                    worksheet.Cells[i, 0] = new Cell("");

                //copying data from data table
                for (int rowIndex = 0; rowIndex < dt.Rows.Count; rowIndex++)
                {
                    for (int colIndex = 0; colIndex < dt.Columns.Count; colIndex++)
                    {
                        worksheet.Cells[rowIndex, colIndex] = new Cell(dt.Rows[rowIndex][colIndex].ToString());
                    }
                }

                workbook.Worksheets.Add(worksheet);
                workbook.Save(file);
            }
        }

        private void incrProgressBar()
        {
            if (toolStripProgressBar1.Value < toolStripProgressBar1.Maximum)
            {
                toolStripProgressBar1.Value += 1;
                toolStripStatusLabel1.Text = String.Format("({0}/{1})", toolStripProgressBar1.Value, toolStripProgressBar1.Maximum);
            }
        }

        //Возникает по окончании Теста связи или Опроса ОДНОГО счетчика из списка
        public event EventHandler meterPinged;
        void Form1_meterPinged(object sender, EventArgs e)
        {
            incrProgressBar();
        }

        //Возникает по окончании Теста связи или Опроса ВСЕХ счетчиков списка
        public event EventHandler pollingEnd;
        void Form1_pollingEnd(object sender, EventArgs e)
        {
            InProgress = false;
            doStopProcess = false;

        }

        Thread pingThr = null;
        //Обработчик кнопки "Тест связи"
        private void buttonPing_Click(object sender, EventArgs e)
        {
            InProgress = true;
            doStopProcess = false;

            DeleteLogFiles();

            pingThr = new Thread(pingMeters);
            pingThr.Start((object)dt);
        }

        int attempts = 3;
        private void pingMeters(Object metersDt)
        {
            DataTable dt = (DataTable)metersDt;
            int columnIndexFactory = 1;
            int columnIndexResult = 2;

            List<string> factoryNumbers = new List<string>();
            for (int i = 1; i < dt.Rows.Count; i++)
            {
                int tmpNumb = 0;
                object oColFactory = dt.Rows[i][columnIndexFactory];
                object oColResult = dt.Rows[i][columnIndexResult];

           
            }
        END:

            Invoke(pollingEnd);
        }

        //Обработчик кнопки "Опрос"
        struct PollMetersArguments
        {
            public DataTable dt;
            public List<int> incorrectRows;
        }

        private void buttonPoll_Click(object sender, EventArgs e)
        {

            //проверить сетевой адрес прибора

            InProgress = true;
            doStopProcess = false;

            DeleteLogFiles();

            PollMetersArguments pma = new PollMetersArguments();
            pma.dt = dt;
            pma.incorrectRows = null;

            pingThr = new Thread(pollMeters);
            pingThr.Start((object)pma);
        }

        private void DeleteLogFiles()
        {
            string curDir = AppDomain.CurrentDomain.BaseDirectory;
            try
            {
                FileInfo fi = new FileInfo(curDir + "teplouchetlog.pi");
                if (fi.Exists)
                    fi.Delete();

                fi = new FileInfo(curDir + "metersinfo.pi");
                if (fi.Exists)
                    fi.Delete();

                fi = new FileInfo(curDir + "datainfo.pi");
                if (fi.Exists)
                    fi.Delete();
            }
            catch (Exception ex)
            {
                //
            }
        }



        private void pollMeters(Object pollMetersArgs)
        {
            PollMetersArguments pmaInp = (PollMetersArguments)pollMetersArgs;

            m_connection_string = ConfigurationSettings.AppSettings["constr"];
            Open(m_connection_string);

            string meterAddress = "";
            GetMeterAddressBySerial(this.meterNumberFromCell, out meterAddress);

            if (meterAddress != meterAddressFromCell)
            {
                MessageBox.Show("Адреса счетчиков в БД и таблице не совпадают..." + meterAddress + "//" + meterAddressFromCell);
                  goto END;
            }

            //richTextBox1.Text += "Адрес счетчика, полученный по s/n: " + meterAddress + "\n\n";


            for (int m = 0; m < dt.Rows.Count; m++)
            {
                //richTextBox1.Text += "Получасовка " + m + ":\n";

                //интерпретируем значения
                Value[] valArr = new Value[4];

                bool success = true;

                for (int v = 0; v < valArr.Length; v++)
                {


                    valArr[v] = new Value();
                    valArr[v].status = false;

                    string fmt = "dd.MM.yyyy HH:mm";
                    string tmpDateString = dt.Rows[m][6].ToString() + " " + dt.Rows[m][5].ToString();
                    DateTime tmpDateTime = new DateTime();
                    bool dtParseRes = DateTime.TryParseExact(tmpDateString, fmt, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out tmpDateTime);

                    if (dtParseRes)
                    {
                        valArr[v].dt = tmpDateTime;
                    }
                    else
                    {
                        success = false;
                        WriteToLog("Дата не разобрана...m/v: " + m + "//" + v);
                        WriteToLog("\nFormat: " + fmt);
                        WriteToLog("\nDtStr: " + tmpDateString);
                        break;
                    }


                    string selectedParamNameGuid = "";

                    switch (v)
                    {
                        case 0:
                            {
                                selectedParamNameGuid = AP_PNAME_GUID;
                                break;
                            }
                        case 1:
                            {
                                selectedParamNameGuid = AM_PNAME_GUID;
                                break;
                            }
                        case 2:
                            {
                                selectedParamNameGuid = RP_PNAME_GUID;
                                break;
                            }
                        case 3:
                            {
                                selectedParamNameGuid = RM_PNAME_GUID;
                                break;
                            }
                    }

                    bool tpres = GetTakenParamIdByGuidAndMetersAddress(selectedParamNameGuid, int.Parse(meterAddress), out valArr[v].id_taken_params);
                    if (!tpres)
                    {
                        WriteToLog("Такен парамес не получен...m/v: " + m + "//" + v);
                       // success = false;
                        //break;
                    }



                    int ApRowIndex = 1;
                    string tmpValString = dt.Rows[m][ApRowIndex + v].ToString().Replace(',', '.');
                    bool vres = float.TryParse(tmpValString, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out valArr[v].value);

                    if (!vres)
                    {
                        WriteToLog("Ошибка преобразования...m/v: " + m + "//" + v);
                        success = false;
                        break;
                    }
                }

                WriteToLog(String.Format("Поулчасовка {0}, 0: id_taken_p: {1}; val: {2}; dt: {3};\n", m, valArr[0].id_taken_params, valArr[0].value, valArr[0].dt));

                if (success)
                { 
                    //запишем в БД
                    if (cbColB.Checked)
                        AddVariousValue(valArr[0]);

                    if (cbColC.Checked)
                        AddVariousValue(valArr[1]);

                    if (cbColD.Checked)
                        AddVariousValue(valArr[2]);

                    if (cbColD.Checked)
                        AddVariousValue(valArr[3]);

                    if (doStopProcess)
                        break;
                }

                Invoke(meterPinged);
            }


        END:
            Close();
            Invoke(pollingEnd);

        }



        //Обработчик клавиши "Стоп"
        private void buttonStop_Click(object sender, EventArgs e)
        {
            doStopProcess = true;

            buttonStop.Enabled = false;
            dgv1.Enabled = false;
        }


        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (InProgress)
            {
                MessageBox.Show("Остановите опрос перед закрытием программы","Напоминание");
                e.Cancel = true;
                return;
            }


        }

        /// <summary>
        /// Запись в ЛОГ-файл
        /// </summary>
        /// <param name="str"></param>
        public void WriteToLog(string str, bool doWrite = true)
        {
            if (doWrite)
            {
                StreamWriter sw = null;
                FileStream fs = null;
                try
                {
                    string curDir = AppDomain.CurrentDomain.BaseDirectory;
                    fs = new FileStream(curDir + "metersinfo.pi", FileMode.Append, FileAccess.Write, FileShare.ReadWrite);
                    sw = new StreamWriter(fs, Encoding.Default);
                    sw.WriteLine(DateTime.Now.ToString() + ": " + str);
                    sw.Close();
                    fs.Close();
                }
                catch
                {
                }
                finally
                {
                    if (sw != null)
                    {
                        sw.Close();
                        sw = null;
                    }
                    if (fs != null)
                    {
                        fs.Close();
                        fs = null;
                    }
                }
            }
        }
        public void WriteToSeparateLog(string str, bool doWrite = true)
        {
            if (doWrite)
            {
                StreamWriter sw = null;
                FileStream fs = null;
                try
                {
                    string curDir = AppDomain.CurrentDomain.BaseDirectory;
                    fs = new FileStream(curDir + "datainfo.pi", FileMode.Append, FileAccess.Write, FileShare.ReadWrite);
                    sw = new StreamWriter(fs, Encoding.Default);
                    sw.WriteLine(DateTime.Now.ToString() + ": " + str);
                    sw.Close();
                    fs.Close();
                }
                catch
                {
                }
                finally
                {
                    if (sw != null)
                    {
                        sw.Close();
                        sw = null;
                    }
                    if (fs != null)
                    {
                        fs.Close();
                        fs = null;
                    }
                }
            }
        }



        private void pictureBoxLogo_Click(object sender, EventArgs e)
        {
            Process.Start("http://prizmer.ru/");
        }


        string tmpFactoryNumberString = "";

        private void button1_Click(object sender, EventArgs e)
        {
            loadXlsFile(float.Parse(numericUpDown1.Value.ToString()));
        }
    }
}
