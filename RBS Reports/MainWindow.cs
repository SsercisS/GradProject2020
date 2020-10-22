using System;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using System.IO;
using ClosedXML.Excel;

namespace RBS_Reports
{
   

    public partial class MainWindow : Form
    {
        public MainWindow()
        {
            InitializeComponent();
        
        }

        private void tabControl1_Selecting(object sender, TabControlCancelEventArgs e)
        {
                     
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker2.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.CustomFormat = "dd.MM.yyyy";
            dateTimePicker2.CustomFormat = "dd.MM.yyyy";
            openFileDialog1.Filter = "Реестр заявок|*.xlsx|Этапы заявок|*.xlsx"; //ставим фильтр для OpenFileDialog
            saveFileDialog1.Filter = "Документ Excel|*.xlsx";
            checkedListBox1.Items.AddRange(Reestr);
            checkedListBox2.Items.AddRange(Etap);

        }

        private int i = 2;
        private int y = 18;
        private int nameIndex = 1;
        private string[] Etap = {
            "00. Не рассмотрено",
            "01. Ожидание ответа",
            "02. Подготовка ИС/ИДОЗ/ЗапросТКП",
            "03. Согласование ИДОЗ/ЗапросТКП",
            "03.1 Отработка замечаний",
            "04. Сбор ТКП",
            "05. Формирование АС/оценка тендера",
            "06. Согласование итогов (ИС/АС)",
            "06.1 Отработка замечаний",
            "07. Передано в ЗК/ЦЗК/СЗ"
        };

        private string[] Reestr = {
            "Окончание срока план",
            "Фактическое окончание",
            "Просроченные заявки",
            "Дней просрочки",
            "Категория штрафов",
            
        };

        #region Код динамичного создания элементов

        DoubleBufferedDataGridView dataGridView = new DoubleBufferedDataGridView();
        private bool online = false;
        DataTable tableRem= new DataTable();


        /// <summary>
        /// Процедура создания новой вкладки для выбора Excel- документа, title обязательно принимает в себя значение "Добавить документ"
        /// </summary>
        /// <param name="title"></param>
        private void NewPage(string title)
        {
            //if ()

            y = 18; //скидываем позицию для чек боксов

            TabPage tabPage = new TabPage(title);
            TabPage tab = new TabPage(i.ToString());
            
            tabPage.ContextMenuStrip = null;

            dataGridView = new DoubleBufferedDataGridView();

            dataGridView.Name = tabControl1.SelectedIndex.ToString();
            dataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            tabControl1.TabPages.Add(tabPage); //новая вкладка с имененм файла
            tabControl2.TabPages.Add(tab); //Новая вкладка с порядковым номером на именах показателей

            
            tabControl1.TabPages[tabControl1.TabPages.Count - 2].BackColor = Color.White;
           

            dataGridView.Dock = DockStyle.Fill;
            tabControl1.TabPages[tabControl1.TabPages.Count - 2].Controls.Add(dataGridView); //добваляю DataGridView на TabPage

            if (!online) { OpenTable(dataGridView); }


            if (tabControl1.TabPages.Count - 2 > -1)
            {
                tabControl1.TabPages[tabControl1.TabPages.Count - 2].Text = Path.GetFileNameWithoutExtension(openFileDialog1.FileName);
                tabControl1.SelectedIndex = tabControl1.TabPages.Count - 2;
                i++;
            }
            




            switch (openFileDialog1.FilterIndex)
            {
                case (1):
                    try
                    {
                        
                        checkedList.Items.AddRange(Reestr);
                        dataGridView.Columns[0].Frozen = true;
                        dataGridView.Columns[1].Frozen = true;
                    }
                    catch (System.ArgumentOutOfRangeException ex)
                    {
                        MessageBox.Show(
                                                "Произошла ошибка при попытке отображения файла. Пожалуйста, удалите вкладку и повторите попытку\n" +
                                                "Если ошибка потворится вновь, обратитесь к руководству пользователя\n" +
                                                "Если вкладка не удалилась автоматически, то проделайте это вручную",
                                                ex.ToString(),
                                                MessageBoxButtons.OK,
                                                MessageBoxIcon.Error);
                        tabControl1.TabPages[tabControl1.SelectedIndex].Dispose();
                        tabControl2.TabPages[tabControl2.SelectedIndex].Dispose();
                        res = false;
                    }
                    
                    break;
                case (2):
                    try
                    {
                        checkedList.Items.AddRange(Etap);
                        dataGridView.Columns[0].Frozen = true;
                        dataGridView.Columns[1].Frozen = true;
                    }
                    catch (System.ArgumentOutOfRangeException ex)
                    {
                        MessageBox.Show(
                                                "Произошла ошибка при попытке отображения файла. Пожалуйста, удалите вкладку и повторите попытку\n" +
                                                "Если ошибка потворится вновь, обратитесь к руководству пользователя\n" +
                                                "Если вкладка не удалилась автоматически, то проделайте это вручную",
                                                ex.ToString(),
                                                MessageBoxButtons.OK,
                                                MessageBoxIcon.Error);
                        tabControl1.TabPages[tabControl1.SelectedIndex].Dispose();
                        tabControl2.TabPages[tabControl2.SelectedIndex].Dispose();
                        res = false;
                    }
                    
                    break;
            }
        }
        /// <summary>
        /// Процедура создания новой вкладки для выгрузки из БД, title обязательно принимает в себя значение "Добавить документ"
        /// </summary>
        /// <param name="title"></param>
        /// <param name="txt"></param>
        private void NewPage(string title, string txt)
        {

            y = 18; //скидываем позицию для чек боксов
            


            TabPage tabPage = new TabPage(title);
            TabPage tab = new TabPage(i.ToString());

            tableRem = new DataTable();

            tabPage.ContextMenuStrip = null;

            dataGridView = new DoubleBufferedDataGridView();
            //CheckBox checkBox = new CheckBox();
            dataGridView.Name = tabControl1.SelectedIndex.ToString();
            dataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            DoubleBufferedDataGridView d = new DoubleBufferedDataGridView();

            tabControl1.TabPages.Add(tabPage); //новая вкладка с имененм файла
            tabControl2.TabPages.Add(tab); //Новая вкладка с порядковым номером на именах показателей

            tabControl1.TabPages[tabControl1.TabPages.Count - 2].BackColor = Color.Transparent;
            //openFileDialog1.ShowDialog();

            dataGridView.Dock = DockStyle.Fill;
            tabControl1.TabPages[tabControl1.TabPages.Count - 2].Controls.Add(dataGridView); //добваляю DataGridView на TabPage

            Setting s = new Setting();
            s.LoadRemarks(dateTimePicker1.Value.ToShortDateString(), dateTimePicker2.Value.ToShortDateString(), tableRem);

            dataGridView.DataSource = tableRem;

            for (int i=0; i<dataGridView.Columns.Count; i++)
            {
                NewCheckBox(dataGridView.Columns[i].Name, y, nameIndex);
                y += 23;
                nameIndex++;
            }
            tabControl1.TabPages[tabControl1.TabPages.Count - 2].Text = "Замечания " + i + "за " + DateTime.Today.ToShortDateString();
            tabControl1.SelectedIndex = tabControl1.TabPages.Count - 2;
            tabControl2.SelectedIndex = tabControl2.TabPages.Count - 2;
            i++;


        }

        CheckBox checkBox = new CheckBox();
        /// <summary>
        /// Процедура создания CheckBox, который принимает в себя событие EventHandler
        /// </summary>
        /// <param name="text"></param>
        /// <param name="y"></param>
        /// <param name="nameIndex"></param>
        private void NewCheckBox(string text, int y, int nameIndex)
        {
            checkBox = new CheckBox();
            checkBox.Name = "checkBox" + nameIndex.ToString(); //задаем имя CheckBox
            checkBox.Appearance = Appearance.Normal; //задаем тип CheckBox
            checkBox.Location = new Point(3, y); //Задаем локацию CheckBox
            checkBox.Text = text;
            checkBox.Checked = true;
            checkBox.Click += new System.EventHandler(checkBox_CheckedChanged); //Присваиваем CheckBox событие
            checkBox.Width = tabControl2.TabPages[tabControl2.TabPages.Count - 2].Width - 10;
            tabControl2.TabPages[tabControl2.TabPages.Count - 2].Controls.Add(checkBox); //Добавляем CheckBox на форму
                
        }

        ContextMenuStrip contextMenuStrip = new ContextMenuStrip();

       

        CheckedListBox checkedList = new CheckedListBox();
        /// <summary>
        /// Процедура создания ChechekedListBox, который принимает в себя событие ItemCheckEventHundler
        /// </summary>
        /// <param name="nameIndex"></param>
        private void NewCheckedListBox(int nameIndex)
        {
            checkedList = new CheckedListBox();
            checkedList.Name =  nameIndex.ToString();
            checkedList.Dock = System.Windows.Forms.DockStyle.Bottom;
            checkedList.FormattingEnabled = true;
            checkedList.Location = new System.Drawing.Point(3, 295);
            checkedList.Size = new System.Drawing.Size(227, 154);
            checkedList.TabIndex = 0;
            tabControl2.TabPages[tabControl2.TabPages.Count - 2].Controls.Add(checkedList);
            checkedList.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.checkedList_ItemCheck);
        }

        #endregion

        #region код обработки и открытия данных
       
        /// <summary>
        /// Процедура события, возникающего при изменении статуса элемента CheckedListBox, служащего для добавления рассчитываемых столбиков в рабочую область
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void checkedList_ItemCheck(object sender, ItemCheckEventArgs  e)
        {
            var dg = tabControl1.TabPages[tabControl1.SelectedIndex].Controls
                .OfType<DoubleBufferedDataGridView>()
                .FirstOrDefault(x => x.Name == tabControl1.SelectedIndex.ToString());

            int check = 0; ;

            for (int n=0; n<dg.Columns.Count; n++)
            {
                if (dg.Columns[n].HeaderText == (sender as CheckedListBox).Items[e.Index].ToString())
                {
                    check++;
                }
                
            }

            dg.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
            
            //checkedList.

            if (dg != null)
            {
                if (e.NewValue== CheckState.Checked)
                {
                    if (check < 1)
                    {
                        dg.Columns.Add((sender as CheckedListBox).Items[e.Index].ToString(), (sender as CheckedListBox).Items[e.Index].ToString());
                        check = 0;


                        switch (openFileDialog1.FilterIndex)
                        {
                            case 1://реестр заявок 
                                checkedListBox1.Enabled = true;
                                checkedListBox2.Enabled = false;
                                switch ((sender as CheckedListBox).Items[e.Index].ToString())
                                {
                                    case "Окончание срока план":
                                        try
                                        {
                                            for (int n = 0; n < dg.Rows.Count - 1; n++)
                                            {
                                                dg[dg.Columns[(sender as CheckedListBox).Items[e.Index].ToString()].Index, n].Value =
                                                Raschety.OkonchanyePlan(dg[dg.Columns["Тип заявки"].Index, n].Value.ToString()
                                                , dg[dg.Columns["Дата распределения"].Index, n].Value.ToString());
                                            }
                                            checkedListBox1.SetItemCheckState(e.Index, CheckState.Checked);
                                            checkedListBox1.Enabled = true;
                                            checkedListBox2.Enabled = false;
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show(
                                                "Произошла ошибка в расчетах. Для начала,выберите предыдущие столбцы",
                                                ex.ToString(),
                                                MessageBoxButtons.OK,
                                                MessageBoxIcon.Error);
                                            dg.Columns.Remove("Окончание срока план");
                                            e.NewValue = CheckState.Unchecked;
                                        }
                                        
                                        break;
                                    case "Фактическое окончание":
                                        try
                                        {
                                            for (int n = 0; n < dg.Rows.Count - 1; n++)
                                            {
                                                dg[dg.Columns[(sender as CheckedListBox).Items[e.Index].ToString()].Index, n].Value =
                                                Raschety.OkonchanyeFakt(dg[dg.Columns["Окончание срока план"].Index, n].Value.ToString()
                                                , dg[dg.Columns["Количество дней продления"].Index, n].Value.ToString());
                                            }
                                            checkedListBox1.SetItemCheckState(e.Index, CheckState.Checked);
                                            checkedListBox1.Enabled = true;
                                            checkedListBox2.Enabled = false;
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show(
                                                "Произошла ошибка в расчетах. Для начала,выберите предыдущие столбцы",
                                                ex.ToString(),
                                                MessageBoxButtons.OK,
                                                MessageBoxIcon.Error);
                                            dg.Columns.Remove("Фактическое окончание");
                                            e.NewValue = CheckState.Unchecked;
                                            checkedListBox1.Enabled = true;
                                            checkedListBox2.Enabled = false;
                                        }
                                        
                                        break;
                                    case "Просроченные заявки":
                                        try
                                        {
                                            for (int n = 0; n < dg.Rows.Count - 1; n++)
                                            {

                                                if (dg[dg.Columns["дата передачи на контрактование"].Index, n].Value.ToString().Length == 0)
                                                {
                                                    dg[dg.Columns["дата передачи на контрактование"].Index, n].Value = "";
                                                }

                                                dg[dg.Columns[(sender as CheckedListBox).Items[e.Index].ToString()].Index, n].Value =
                                                    Raschety.Prosrochka(dg[dg.Columns["дата передачи на контрактование"].Index, n].Value.ToString()
                                                    , dg[dg.Columns["Фактическое окончание"].Index, n].Value.ToString());
                                            }
                                            checkedListBox1.SetItemCheckState(e.Index, CheckState.Checked);
                                            checkedListBox1.Enabled = true;
                                            checkedListBox2.Enabled = false;
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show(
                                                "Произошла ошибка в расчетах. Для начала,выберите предыдущие столбцы",
                                                ex.ToString(),
                                                MessageBoxButtons.OK,
                                                MessageBoxIcon.Error);
                                            dg.Columns.Remove("Просроченные заявки");
                                            e.NewValue = CheckState.Unchecked;
                                        }
                                        
                                        break;
                                    case "Дней просрочки":
                                        try
                                        {
                                            for (int n = 0; n < dg.Rows.Count - 1; n++)
                                            {
                                                if (dg[dg.Columns["дата передачи на контрактование"].Index, n].Value.ToString().Length == 0)
                                                {
                                                 dg[dg.Columns["дата передачи на контрактование"].Index, n].Value = "";
                                                }

                                                  dg[dg.Columns[(sender as CheckedListBox).Items[e.Index].ToString()].Index, n].Value =
                                                Raschety.DneyProsrochky(dg[dg.Columns["дата передачи на контрактование"].Index, n].Value.ToString()
                                                , dg[dg.Columns["Фактическое окончание"].Index, n].Value.ToString());
                                            }
                                            checkedListBox1.SetItemCheckState(e.Index, CheckState.Checked);
                                            checkedListBox1.Enabled = true;
                                            checkedListBox2.Enabled = false;
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show(
                                                "Произошла ошибка в расчетах. Для начала,выберите предыдущие столбцы",
                                                ex.ToString(),
                                                MessageBoxButtons.OK,
                                                MessageBoxIcon.Error);
                                            dg.Columns.Remove("Дней просрочки");
                                            e.NewValue = CheckState.Unchecked;
                                        }
                                        
                                        break;
                                    case "Категория штрафов":
                                        try
                                        {
                                            for (int n = 0; n < dg.Rows.Count - 1; n++)
                                            {
                                                dg[dg.Columns[(sender as CheckedListBox).Items[e.Index].ToString()].Index, n].Value =
                                                    Raschety.Shtrafy(dg[dg.Columns["Дней просрочки"].Index, n].Value.ToString());
                                            }
                                            checkedListBox1.SetItemCheckState(e.Index, CheckState.Checked);
                                            checkedListBox1.Enabled = true;
                                            checkedListBox2.Enabled = false;
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show(
                                                "Произошла ошибка в расчетах. Для начала,выберите предыдущие столбцы",
                                                ex.ToString(),
                                                MessageBoxButtons.OK,
                                                MessageBoxIcon.Error);
                                            dg.Columns.Remove("Категория штрафов");
                                            e.NewValue = CheckState.Unchecked;
                                        }
                                        
                                        break;
                                }
                                break;

                            case 2: //этапы заявок
                                checkedListBox1.Enabled = false;
                                checkedListBox2.Enabled = true;
                                switch ((sender as CheckedListBox).Items[e.Index].ToString())
                                {
                                    case "00. Не рассмотрено":
                                        try
                                        {
                                            for (int n = 0; n < dg.Rows.Count - 1; n++)
                                            {
                                                dg[dg.Columns[(sender as CheckedListBox).Items[e.Index].ToString()].Index, n].Value =
                                                    Raschety.NeRassmotreno(dg[dg.Columns["Факт приема в работу"].Index, n].Value.ToString());

                                            }
                                            checkedListBox2.SetItemCheckState(e.Index, CheckState.Checked);
                                            checkedListBox1.Enabled = false;
                                            checkedListBox2.Enabled = true;
                                        }
                                        catch (Exception ex)
                                        {

                                            MessageBox.Show(
                                                 "Произошла ошибка в расчетах",
                                                 ex.ToString(),
                                                 MessageBoxButtons.OK,
                                                 MessageBoxIcon.Error);
                                            
                                        }

                                        break;
                                    case "02. Подготовка ИС/ИДОЗ/ЗапросТКП":
                                        try
                                        {
                                            for (int n = 0; n < dg.Rows.Count - 1; n++)
                                            {
                                                if (dg[dg.Columns["01. Ожидание ответа"].Index, n].Value == null)
                                                {
                                                    dg[dg.Columns[(sender as CheckedListBox).Items[e.Index].ToString()].Index, n].Value =
                                                   Raschety.Podgotovka(
                                                       dg[dg.Columns["Тип заявки"].Index, n].Value.ToString()
                                                      , " "
                                                      , dg[dg.Columns["00. Не рассмотрено"].Index, n].Value.ToString());
                                                }
                                                else
                                                {
                                                    dg[dg.Columns[(sender as CheckedListBox).Items[e.Index].ToString()].Index, n].Value =
                                                   Raschety.Podgotovka(
                                                       dg[dg.Columns["Тип заявки"].Index, n].Value.ToString()
                                                      , dg[dg.Columns["01. Ожидание ответа"].Index, n].Value.ToString()
                                                      , dg[dg.Columns["00. Не рассмотрено"].Index, n].Value.ToString());
                                                }
                                                checkedListBox2.SetItemCheckState(e.Index, CheckState.Checked);
                                                checkedListBox1.Enabled = false;
                                                checkedListBox2.Enabled = true;

                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show(
                                                "Произошла ошибка в расчетах. Для начала,выберите предыдущие столбцы",
                                                ex.ToString(),
                                                MessageBoxButtons.OK,
                                                MessageBoxIcon.Error);
                                            dg.Columns.Remove("02. Подготовка ИС/ИДОЗ/ЗапросТКП");
                                            e.NewValue = CheckState.Unchecked;

                                        }

                                        break;
                                    case "03. Согласование ИДОЗ/ЗапросТКП":
                                        try
                                        {
                                            for (int n = 0; n < dg.Rows.Count - 1; n++)
                                            {
                                                dg[dg.Columns[(sender as CheckedListBox).Items[e.Index].ToString()].Index, n].Value =
                                                Raschety.Soglasovanie(dg[dg.Columns["Тип заявки"].Index, n].Value.ToString()
                                                                     , dg[dg.Columns["02. Подготовка ИС/ИДОЗ/ЗапросТКП"].Index, n].Value.ToString());

                                            }
                                            checkedListBox2.SetItemCheckState(e.Index, CheckState.Checked);
                                            checkedListBox1.Enabled = false;
                                            checkedListBox2.Enabled = true;
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show(
                                                "Произошла ошибка в расчетах. Для начала,выберите предыдущие столбцы",
                                                ex.ToString(),
                                                MessageBoxButtons.OK,
                                                MessageBoxIcon.Error);
                                            dg.Columns.Remove("03. Согласование ИДОЗ/ЗапросТКП");
                                            e.NewValue = CheckState.Unchecked;
                                        }

                                        break;
                                    case "04. Сбор ТКП":
                                        try
                                        {
                                            for (int n = 0; n < dg.Rows.Count - 1; n++)
                                            {
                                                if (dg[dg.Columns["03.1 Отработка замечаний"].Index, n].Value == null)
                                                {
                                                    dg[dg.Columns["03.1 Отработка замечаний"].Index, n].Value = " ";
                                                }

                                                if (dg[dg.Columns["03. Согласование ИДОЗ/ЗапросТКП"].Index, n].Value == null)
                                                {
                                                    dg[dg.Columns["03. Согласование ИДОЗ/ЗапросТКП"].Index, n].Value = "";
                                                }

                                                dg[dg.Columns[(sender as CheckedListBox).Items[e.Index].ToString()].Index, n].Value =
                                            Raschety.Sbor(dg[dg.Columns["Тип заявки"].Index, n].Value.ToString()
                                                        , dg[dg.Columns["02. Подготовка ИС/ИДОЗ/ЗапросТКП"].Index, n].Value.ToString()
                                                        , dg[dg.Columns["03.1 Отработка замечаний"].Index, n].Value.ToString()
                                                        , dg[dg.Columns["03. Согласование ИДОЗ/ЗапросТКП"].Index, n].Value.ToString());
                                                                                               

                                            }
                                            checkedListBox2.SetItemCheckState(e.Index, CheckState.Checked);
                                            checkedListBox1.Enabled = false;
                                            checkedListBox2.Enabled = true;
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show(
                                                "Произошла ошибка в расчетах. Для начала,выберите предыдущие столбцы",
                                                ex.ToString(),
                                                MessageBoxButtons.OK,
                                                MessageBoxIcon.Error);
                                            dg.Columns.Remove("04. Сбор ТКП");
                                            e.NewValue = CheckState.Unchecked;
                                        }


                                        break;
                                    case "05. Формирование АС/оценка тендера":
                                        try
                                        {
                                            for (int n = 0; n < dg.Rows.Count - 1; n++)
                                            {
                                                dg[dg.Columns[(sender as CheckedListBox).Items[e.Index].ToString()].Index, n].Value =
                                                Raschety.Formirovanye(dg[dg.Columns["Тип заявки"].Index, n].Value.ToString()
                                                                    , dg[dg.Columns["04. Сбор ТКП"].Index, n].Value.ToString());

                                            }
                                            checkedListBox2.SetItemCheckState(e.Index, CheckState.Checked);
                                            checkedListBox1.Enabled = false;
                                            checkedListBox2.Enabled = true;
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show(
                                               "Произошла ошибка в расчетах. Для начала,выберите предыдущие столбцы",
                                                ex.ToString(),
                                                MessageBoxButtons.OK,
                                                MessageBoxIcon.Error);
                                            dg.Columns.Remove("05. Формирование АС/оценка тендера");
                                            e.NewValue = CheckState.Unchecked;
                                        }
                                        
                                        break;
                                    case "06. Согласование итогов (ИС/АС)":
                                        try
                                        {
                                            for (int n = 0; n < dg.Rows.Count - 1; n++)
                                            {
                                                dg[dg.Columns[(sender as CheckedListBox).Items[e.Index].ToString()].Index, n].Value =
                                                Raschety.SoglasovanieItogy(dg[dg.Columns["Тип заявки"].Index, n].Value.ToString()
                                                                    , dg[dg.Columns["05. Формирование АС/оценка тендера"].Index, n].Value.ToString());

                                            }
                                            checkedListBox2.SetItemCheckState(e.Index, CheckState.Checked);
                                            checkedListBox1.Enabled = false;
                                            checkedListBox2.Enabled = true;
                                        }
                                        catch(Exception ex)
                                        {
                                            MessageBox.Show(
                                               "Произошла ошибка в расчетах. Для начала,выберите предыдущие столбцы",
                                                ex.ToString(),
                                                MessageBoxButtons.OK,
                                                MessageBoxIcon.Error);
                                            dg.Columns.Remove("06. Согласование итогов (ИС/АС)");
                                            e.NewValue = CheckState.Unchecked;
                                        }
                                        
                                        break;

                                    case "07. Передано в ЗК/ЦЗК/СЗ":
                                        try
                                        {
                                            for (int n = 0; n < dg.Rows.Count - 1; n++)
                                            {
                                                if (dg[dg.Columns["06.1 Отработка замечаний"].Index, n].Value == null)
                                                {
                                                    dg[dg.Columns["06.1 Отработка замечаний"].Index, n].Value = " ";
                                                }

                                                dg[dg.Columns[(sender as CheckedListBox).Items[e.Index].ToString()].Index, n].Value =
                                                    Raschety.Peredano(dg[dg.Columns["Тип заявки"].Index, n].Value.ToString()
                                                    , dg[dg.Columns["06. Согласование итогов (ИС/АС)"].Index, n].Value.ToString()
                                                    , dg[dg.Columns["05. Формирование АС/оценка тендера"].Index, n].Value.ToString()
                                                    , dg[dg.Columns["06.1 Отработка замечаний"].Index, n].Value.ToString());

                                            }
                                            checkedListBox2.SetItemCheckState(e.Index, CheckState.Checked);
                                            checkedListBox1.Enabled = false;
                                            checkedListBox2.Enabled = true;
                                        }
                                        catch(Exception ex)
                                        {
                                            MessageBox.Show(
                                                "Произошла ошибка в расчетах. Для начала,выберите предыдущие столбцы",
                                                ex.ToString(),
                                                MessageBoxButtons.OK,
                                                MessageBoxIcon.Error);
                                            dg.Columns.Remove("07. Передано в ЗК/ЦЗК/СЗ");
                                            e.NewValue = CheckState.Unchecked;
                                        }
                                        

                                        break;
                                }
                                break;
                        }

                    }
                    else
                    {
                        if (e.NewValue == CheckState.Unchecked)
                        {
                            dg.Columns[(sender as CheckedListBox).Items[e.Index].ToString()].Visible = false;
                            
                            check = 0;
                        }
                        else
                        {
                            dg.Columns[(sender as CheckedListBox).Items[e.Index].ToString()].Visible = true;
                           
                            check = 0;
                        }
                    }
                }
                else
                {
                    if (e.NewValue == CheckState.Unchecked)
                    {
                        dg.Columns[(sender as CheckedListBox).Items[e.Index].ToString()].Visible = false;
                        
                        check = 0;
                    }
                    else
                    {
                        dg.Columns[(sender as CheckedListBox).Items[e.Index].ToString()].Visible = true;
                       
                        check = 0;
                    }
                }
                
                       
            }
                
        }
        
        private static int nameChd=0;

        private static bool res = false;
        /// <summary>
        /// Процедура открытия Excel-документа в рабочую область приложения
        /// </summary>
        /// <param name="dataGrid"></param>
        private void OpenTable(DoubleBufferedDataGridView dataGrid)
        {
            try
            {
                //openFileDialog1.ShowDialog();
                if (openFileDialog1.ShowDialog() == DialogResult.Cancel) { res = false; return; }
                    
                else
                { 
                    string filePath = openFileDialog1.FileName;
                    NewCheckedListBox(nameChd);
                    nameChd++;
                    //Открываем Эксель используя ClosedXML
                    using (XLWorkbook workBook = new XLWorkbook(filePath))
                    {
                        //Считываем первый лист документа
                        IXLWorksheet workSheet = workBook.Worksheet(1);

                        //Создаем новую таблицу DataTable.
                        DataTable dt = new DataTable();

                        //Считываем данный
                        bool firstRow = true;
                        foreach (IXLRow row in workSheet.Rows())
                        {
                        //Первая строка становится названием колонок
                            if (firstRow)
                            {
                                foreach (IXLCell cell in row.Cells())
                                {
                                    dt.Columns.Add(cell.Value.ToString());
                                    NewCheckBox(cell.Value.ToString(), y, nameIndex);
                                    y += 23;
                                    nameIndex++;
                                }
                                firstRow = false;
                            }
                            else
                            {
                                //Добавляем строки в DataTable
                                dt.Rows.Add();
                                int i = 0;
                                foreach (IXLCell cell in row.Cells())
                                {
                                    dt.Rows[dt.Rows.Count - 1][i] = cell.Value.ToString();
                                    i++;
                                }
                            }
                            //Устанавливаем в качестве источника данных DataTable dt
                            dataGrid.DataSource = dt;
                        }

                    }
                }
            }
            catch(ClosedXML.Excel.CalcEngine.ExpressionParseException ex)
            {
                MessageBox.Show(
                                                "Ошибка при прочтении файла",
                                                ex.ToString(),
                                                MessageBoxButtons.OK,
                                                MessageBoxIcon.Error);
                tabControl1.TabPages[tabControl1.SelectedIndex].Dispose();
                tabControl2.TabPages[tabControl2.SelectedIndex].Dispose();
                res = false;
            }
            catch(Exception ex)
            {
                MessageBox.Show(
                                                "Ошибка при загрузке файла",
                                                ex.ToString(),
                                                MessageBoxButtons.OK,
                                                MessageBoxIcon.Error);
                tabControl1.TabPages[tabControl1.SelectedIndex].Dispose();
                tabControl2.TabPages[tabControl2.SelectedIndex].Dispose();
                res = false;
            }
            //int y = 22;
           
        }
        #endregion

        /// <summary>
        /// Процедура события, возникающего при изменении состояния CheckBox, служит для отображения вкладок в рабочей области (рассчитываемые столбцы не учитываются)
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void checkBox_CheckedChanged(object sender, EventArgs e)
        {
            var dg = tabControl1.TabPages[tabControl1.SelectedIndex].Controls
                .OfType<DoubleBufferedDataGridView>()
                .FirstOrDefault(x => x.Name == tabControl1.SelectedIndex.ToString());

            

            if(dg!= null)
            {
                if ((sender as CheckBox).Checked) { dg.Columns[(sender as CheckBox).Text].Visible = true; }
                else { dg.Columns[(sender as CheckBox).Text].Visible = false; }
            }
          
        }

        private void tabControl2_Selecting(object sender, TabControlCancelEventArgs e)
        {
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SettingBD.firstOp = false;
            SettingBD bd = new SettingBD();
            
            bd.Show();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            saveFileDialog1.ShowDialog();
            string Path = saveFileDialog1.FileName;

            using (var workbook = new XLWorkbook())
            {
                var dg = tabControl1.TabPages[tabControl1.SelectedIndex].Controls
               .OfType<DoubleBufferedDataGridView>()
               .FirstOrDefault(x => x.Name == tabControl1.SelectedIndex.ToString());

                DataTable table = new DataTable("shnaga");

                for (int iCol = 0; iCol < dg.Columns.Count; iCol++)
                {
                    table.Columns.Add(dg.Columns[iCol].Name);
                }

                foreach (DataGridViewRow row in dg.Rows)
                {

                    DataRow datarw = table.NewRow();

                    for (int iCol = 0; iCol < dg.Columns.Count; iCol++)
                    {
                        datarw[iCol] = row.Cells[iCol].Value;
                    }

                    table.Rows.Add(datarw);
                }

                
                workbook.Worksheets.Add(table, tabControl1.TabPages[tabControl1.SelectedIndex].Text);

                workbook.SaveAs(Path);

            }
        }

        private void panel3_Paint(object sender, PaintEventArgs e)
        {
            
        }

        private void tabControl1_MouseDown(object sender, MouseEventArgs e)
        {
            online = false;

            switch (e.Button)
            {
                
                case (MouseButtons.Left)://Возникает при клике на левую кнопку мыши
                    if (tabControl1.TabPages.Count == 1)
                    {
                        if (tabControl1.TabPages[tabControl1.TabPages.Count - 1].Text == "Добавить документ")
                        {
                            comboBox2.Items.Clear();

                            NewPage("Добавить документ");
                            var dg = tabControl1.TabPages[tabControl1.SelectedIndex].Controls
                                .OfType<DoubleBufferedDataGridView>()
                                .FirstOrDefault(x => x.Name == tabControl1.SelectedIndex.ToString());
                            for (int i = 0; i < dg.Columns.Count; i++)
                            {
                                comboBox2.Items.Add(dg.Columns[i].HeaderText);
                            }
                        }
                    }
                    
                    break;                
            }

            
        }

        private void tabControl1_Selecting(object sender, TabControlEventArgs e)
        {
            if (res)
            {
                y = 18;
                online = false;
                if (tabControl1.TabPages[e.TabPageIndex].Text == "Добавить документ")
                {
                    NewPage("Добавить документ");
                    comboBox2.Items.Clear();
                    var dg = tabControl1.TabPages[tabControl1.SelectedIndex].Controls
                        .OfType<DoubleBufferedDataGridView>()
                        .FirstOrDefault(x => x.Name == tabControl1.SelectedIndex.ToString());
                    for (int i = 0; i < dg.Columns.Count; i++)
                    {
                        comboBox2.Items.Add(dg.Columns[i].HeaderText);
                    }
                }
                else
                {
                    comboBox2.Items.Clear();
                    res = true;
                    var dg = tabControl1.TabPages[tabControl1.SelectedIndex].Controls
                            .OfType<DoubleBufferedDataGridView>()
                            .FirstOrDefault(x => x.Name == tabControl1.SelectedIndex.ToString());
                    if (dg != null)
                    {
                        for (int i = 0; i < dg.Columns.Count; i++)
                        {
                            comboBox2.Items.Add(dg.Columns[i].HeaderText);
                        }
                    }
                }

                tabControl2.SelectedIndex = tabControl1.SelectedIndex;
            }
            else
            {
                comboBox2.Items.Clear();
                res = true;
                try
                {
                    var dg = tabControl1.TabPages[tabControl1.SelectedIndex].Controls
                        .OfType<DoubleBufferedDataGridView>()
                        .FirstOrDefault(x => x.Name == tabControl1.SelectedIndex.ToString());

                    if (dg != null)
                    {
                        for (int i = 0; i < dg.Columns.Count; i++)
                        {
                            comboBox2.Items.Add(dg.Columns[i].HeaderText);
                        }
                    }
                }
                catch
                {
                }
    
            }
           
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Setting s = new Setting();
                                 
            //MessageBox.Show(dateTimePicker1.Value.ToShortDateString());
           
            //MessageBox.Show(Setting.cnt.ConnectionString);
            //Setting.cnt.Open();
            NewPage("Добавить документ", "");



        }

        /// <summary>
        /// Наследуемый класс для реализации двойной буфферизации элемента отображения данных
        /// </summary>
        class DoubleBufferedDataGridView : DataGridView
        {
            protected override bool DoubleBuffered { get => true; }
        }

        private void checkedListBox2_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            var dg = tabControl1.TabPages[tabControl1.SelectedIndex].Controls
                .OfType<DoubleBufferedDataGridView>()
                .FirstOrDefault(x => x.Name == tabControl1.SelectedIndex.ToString());
            if (e.NewValue == CheckState.Checked)
            {
                switch (checkedListBox2.Items[e.Index].ToString())
                {
                    case "00. Не рассмотрено":
                        try
                        {
                            for (int n = 0; n < dg.Rows.Count - 1; n++)
                            {
                                if (dg[dg.Columns["Статус заявки"].Index, n].Value.ToString() == dg.Columns["00. Не рассмотрено"].HeaderText)
                                {
                                    if (Convert.ToDateTime(dg[dg.Columns["00. Не рассмотрено"].Index, n].Value.ToString()) < DateTime.Now.AddDays(1))
                                    {
                                        dg[dg.Columns["00. Не рассмотрено"].Index, n].Style.BackColor = Color.OrangeRed;
                                    }
                                    else
                                    {
                                        dg[dg.Columns["00. Не рассмотрено"].Index, n].Style.BackColor = Color.AliceBlue;
                                    }
                                }
                                else
                                {
                                    dg[dg.Columns["00. Не рассмотрено"].Index, n].Style.BackColor = Color.AliceBlue;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(
                                 "",
                                 ex.ToString(),
                                 MessageBoxButtons.OK,
                                 MessageBoxIcon.Error);
                        }

                        break;
                    case "02. Подготовка ИС/ИДОЗ/ЗапросТКП":
                        try
                        {
                            for (int n = 0; n < dg.Rows.Count - 1; n++)
                            {
                                if (dg[dg.Columns["Статус заявки"].Index, n].Value.ToString() == dg.Columns["02. Подготовка ИС/ИДОЗ/ЗапросТКП"].HeaderText)
                                {
                                    if (Convert.ToDateTime(dg[dg.Columns["02. Подготовка ИС/ИДОЗ/ЗапросТКП"].Index, n].Value.ToString()) < DateTime.Now.AddDays(1))
                                    {
                                        dg[dg.Columns["02. Подготовка ИС/ИДОЗ/ЗапросТКП"].Index, n].Style.BackColor = Color.OrangeRed;
                                    }
                                    else
                                    {
                                        dg[dg.Columns["02. Подготовка ИС/ИДОЗ/ЗапросТКП"].Index, n].Style.BackColor = Color.AliceBlue;
                                    }
                                }
                                else
                                {
                                    dg[dg.Columns["02. Подготовка ИС/ИДОЗ/ЗапросТКП"].Index, n].Style.BackColor = Color.AliceBlue;
                                }

                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(
                                "Произошла ошибка в расчетах. Для начала,выберите предыдущие столбцы",
                                ex.ToString(),
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);

                        }
                        break;
                    case "03. Согласование ИДОЗ/ЗапросТКП":
                        try
                        {
                            for (int n = 0; n < dg.Rows.Count - 1; n++)
                            {
                                if (dg[dg.Columns["Статус заявки"].Index, n].Value.ToString() == dg.Columns["03. Согласование ИДОЗ/ЗапросТКП"].HeaderText)
                                {
                                    if (Convert.ToDateTime(dg[dg.Columns["03. Согласование ИДОЗ/ЗапросТКП"].Index, n].Value.ToString()) < DateTime.Now.AddDays(1))
                                    {
                                        dg[dg.Columns["03. Согласование ИДОЗ/ЗапросТКП"].Index, n].Style.BackColor = Color.OrangeRed;
                                    }
                                    else
                                    {
                                        dg[dg.Columns["03. Согласование ИДОЗ/ЗапросТКП"].Index, n].Style.BackColor = Color.AliceBlue;
                                    }
                                }
                                else
                                {
                                    dg[dg.Columns["03. Согласование ИДОЗ/ЗапросТКП"].Index, n].Style.BackColor = Color.AliceBlue;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(
                                "Произошла ошибка в расчетах. Для начала,выберите предыдущие столбцы",
                                ex.ToString(),
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
                        }

                        break;
                    case "04. Сбор ТКП":
                        try
                        {
                            for (int n = 0; n < dg.Rows.Count - 1; n++)
                            {

                                if (dg[dg.Columns["Статус заявки"].Index, n].Value.ToString() == dg.Columns["04. Сбор ТКП"].HeaderText)
                                {
                                    if (Convert.ToDateTime(dg[dg.Columns["04. Сбор ТКП"].Index, n].Value.ToString()) < DateTime.Now.AddDays(1))
                                    {
                                        dg[dg.Columns["04. Сбор ТКП"].Index, n].Style.BackColor = Color.OrangeRed;
                                    }
                                    else
                                    {
                                        dg[dg.Columns["04. Сбор ТКП"].Index, n].Style.BackColor = Color.AliceBlue;
                                    }
                                }
                                else
                                {
                                    dg[dg.Columns["04. Сбор ТКП"].Index, n].Style.BackColor = Color.AliceBlue;
                                }

                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(
                                "",
                                ex.ToString(),
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);

                        }
                        break;
                    case "05. Формирование АС/оценка тендера":
                        try
                        {
                            for (int n = 0; n < dg.Rows.Count - 1; n++)
                            {

                                if (dg[dg.Columns["Статус заявки"].Index, n].Value.ToString() == dg.Columns["05. Формирование АС/оценка тендера"].HeaderText)
                                {
                                    if (Convert.ToDateTime(dg[dg.Columns["05. Формирование АС/оценка тендера"].Index, n].Value.ToString()) < DateTime.Now.AddDays(1))
                                    {
                                        dg[dg.Columns["05. Формирование АС/оценка тендера"].Index, n].Style.BackColor = Color.OrangeRed;
                                    }
                                    else
                                    {
                                        dg[dg.Columns["05. Формирование АС/оценка тендера"].Index, n].Style.BackColor = Color.AliceBlue;
                                    }
                                }
                                else
                                {
                                    dg[dg.Columns["05. Формирование АС/оценка тендера"].Index, n].Style.BackColor = Color.AliceBlue;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(
                               "",
                                ex.ToString(),
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
                        }

                        break;
                    case "06. Согласование итогов (ИС/АС)":
                        try
                        {
                            for (int n = 0; n < dg.Rows.Count - 1; n++)
                            {

                                if (dg[dg.Columns["Статус заявки"].Index, n].Value.ToString() == dg.Columns["06. Согласование итогов (ИС/АС)"].HeaderText)
                                {
                                    if (Convert.ToDateTime(dg[dg.Columns["06. Согласование итогов (ИС/АС)"].Index, n].Value.ToString()) < DateTime.Now.AddDays(1))
                                    {
                                        dg[dg.Columns["06. Согласование итогов (ИС/АС)"].Index, n].Style.BackColor = Color.OrangeRed;
                                    }
                                    else
                                    {
                                        dg[dg.Columns["06. Согласование итогов (ИС/АС)"].Index, n].Style.BackColor = Color.AliceBlue;
                                    }
                                }
                                else
                                {
                                    dg[dg.Columns["06. Согласование итогов (ИС/АС)"].Index, n].Style.BackColor = Color.AliceBlue;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(
                               "",
                                ex.ToString(),
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);

                        }

                        break;

                    case "07. Передано в ЗК/ЦЗК/СЗ":
                        try
                        {
                            for (int n = 0; n < dg.Rows.Count - 1; n++)
                            {

                                if (dg[dg.Columns["07. Передано в ЗК/ЦЗК/СЗ"].Index, n].Value.ToString() != " ")
                                {
                                    if (dg[dg.Columns["Статус заявки"].Index, n].Value.ToString() == dg.Columns["07. Передано в ЗК/ЦЗК/СЗ"].HeaderText)
                                    {
                                        if (Convert.ToDateTime(dg[dg.Columns["07. Передано в ЗК/ЦЗК/СЗ"].Index, n].Value.ToString()) < DateTime.Now.AddDays(1))
                                        {
                                            dg[dg.Columns["07. Передано в ЗК/ЦЗК/СЗ"].Index, n].Style.BackColor = Color.OrangeRed;
                                        }
                                        else
                                        {
                                            dg[dg.Columns["07. Передано в ЗК/ЦЗК/СЗ"].Index, n].Style.BackColor = Color.AliceBlue;
                                        }
                                    }
                                    else
                                    {
                                        dg[dg.Columns["07. Передано в ЗК/ЦЗК/СЗ"].Index, n].Style.BackColor = Color.AliceBlue;
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(
                                "",
                                ex.ToString(),
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);

                        }
                        break;
                }

            }
            else
            {
                switch (checkedListBox2.Items[e.Index].ToString())
                {
                    case "00. Не рассмотрено":
                        try
                        {
                            for (int n = 0; n < dg.Rows.Count - 1; n++)
                            {
                                dg[dg.Columns["00. Не рассмотрено"].Index, n].Style.BackColor = Color.White;
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(
                                 "Для начала добватье столбик в рабочую область, это можно сделать на панели отображения столбцов",
                                 ex.ToString(),
                                 MessageBoxButtons.OK,
                                 MessageBoxIcon.Error);

                        }

                        break;
                    case "02. Подготовка ИС/ИДОЗ/ЗапросТКП":
                        try
                        {
                            for (int n = 0; n < dg.Rows.Count - 1; n++)
                            {
                                dg[dg.Columns["02. Подготовка ИС/ИДОЗ/ЗапросТКП"].Index, n].Style.BackColor = Color.White;
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(
                                "Для начала добватье столбик в рабочую область, это можно сделать на панели отображения столбцов",
                                ex.ToString(),
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);

                        }

                        break;
                    case "03. Согласование ИДОЗ/ЗапросТКП":
                        try
                        {
                            for (int n = 0; n < dg.Rows.Count - 1; n++)
                            {
                                dg[dg.Columns["03. Согласование ИДОЗ/ЗапросТКП"].Index, n].Style.BackColor = Color.White;
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(
                                "Для начала добватье столбик в рабочую область, это можно сделать на панели отображения столбцов",
                                ex.ToString(),
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
                        }

                        break;
                    case "04. Сбор ТКП":
                        try
                        {
                            for (int n = 0; n < dg.Rows.Count - 1; n++)
                            {
                                dg[dg.Columns["04. Сбор ТКП"].Index, n].Style.BackColor = Color.White;
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(
                                "Для начала добватье столбик в рабочую область, это можно сделать на панели отображения столбцов",
                                ex.ToString(),
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
                        }


                        break;
                    case "05. Формирование АС/оценка тендера":
                        try
                        {
                            for (int n = 0; n < dg.Rows.Count - 1; n++)
                            {
                                dg[dg.Columns["05. Формирование АС/оценка тендера"].Index, n].Style.BackColor = Color.White;
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(
                               "Для начала добватье столбик в рабочую область, это можно сделать на панели отображения столбцов",
                                ex.ToString(),
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
                        }

                        break;
                    case "06. Согласование итогов (ИС/АС)":
                        try
                        {
                            for (int n = 0; n < dg.Rows.Count - 1; n++)
                            {
                                dg[dg.Columns["06. Согласование итогов (ИС/АС)"].Index, n].Style.BackColor = Color.White;
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(
                               "Для начала добватье столбик в рабочую область, это можно сделать на панели отображения столбцов",
                                ex.ToString(),
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
                        }

                        break;

                    case "07. Передано в ЗК/ЦЗК/СЗ":
                        try
                        {
                            for (int n = 0; n < dg.Rows.Count - 1; n++)
                            {
                                dg[dg.Columns["07. Передано в ЗК/ЦЗК/СЗ"].Index, n].Style.BackColor = Color.White;
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(
                                "Для начала добватье столбик в рабочую область, это можно сделать на панели отображения столбцов",
                                ex.ToString(),
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
                        }
                        break;
                }
            }
            
        }

        private void checkedListBox1_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            var dg = tabControl1.TabPages[tabControl1.SelectedIndex].Controls
                    .OfType<DoubleBufferedDataGridView>()
                    .FirstOrDefault(x => x.Name == tabControl1.SelectedIndex.ToString());

            if (e.NewValue == CheckState.Checked)
            {
                switch (checkedListBox1.Items[e.Index].ToString())
                {
                    case "Окончание срока план":
                        try
                        {
                            for (int n = 0; n < dg.Rows.Count - 1; n++)
                            {
                                dg[dg.Columns["Окончание срока план"].Index, n].Style.BackColor = Color.AliceBlue;
                            }
                        }
                        catch(System.NullReferenceException ex)
                        {
                            MessageBox.Show(
                               "Для начала добватье столбик в рабочую область, это можно сделать на панели отображения столбцов",
                                ex.ToString(),
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
                        }
                       
                        break;
                    case "Фактическое окончание":
                        try
                        {
                            for (int n = 0; n < dg.Rows.Count - 1; n++)
                            {
                                dg[dg.Columns["Фактическое окончание"].Index, n].Style.BackColor = Color.AliceBlue;
                            }
                        }
                        catch (System.NullReferenceException ex)
                        {
                            MessageBox.Show(
                               "Для начала добватье столбик в рабочую область, это можно сделать на панели отображения столбцов",
                                ex.ToString(),
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
                        }
                        
                        break;
                    case "Просроченные заявки":
                        try
                        {
                            for (int n = 0; n < dg.Rows.Count - 1; n++)
                            {

                                if (dg[dg.Columns["Просроченные заявки"].Index, n].Value.ToString() != "-")
                                {
                                    dg[dg.Columns["Просроченные заявки"].Index, n].Style.BackColor = Color.OrangeRed;
                                }
                                else
                                {
                                    dg[dg.Columns["Просроченные заявки"].Index, n].Style.BackColor = Color.AliceBlue;
                                }

                            }
                            
                        }
                        catch (System.NullReferenceException ex)
                        {
                            MessageBox.Show(
                               "Для начала добватье столбик в рабочую область, это можно сделать на панели отображения столбцов",
                                ex.ToString(),
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
                        }
                        
                        break;
                    case "Дней просрочки":
                        try
                        {
                            for (int n = 0; n < dg.Rows.Count - 1; n++)
                            {

                            }
                        }
                        catch (System.NullReferenceException ex)
                        {
                            MessageBox.Show(
                               "Для начала добватье столбик в рабочую область, это можно сделать на панели отображения столбцов",
                                ex.ToString(),
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
                        }
                        
                        break;
                    case "Категория штрафов":
                        try
                        {
                            for (int n = 0; n < dg.Rows.Count - 1; n++)
                            {
                                if (dg[dg.Columns["Категория штрафов"].Index, n].Value.ToString() != "0")
                                {
                                    dg[dg.Columns["Категория штрафов"].Index, n].Style.BackColor = Color.OrangeRed;
                                }
                                else
                                {
                                    dg[dg.Columns["Категория штрафов"].Index, n].Style.BackColor = Color.AliceBlue;
                                }
                            }
                        }
                        catch (System.NullReferenceException ex)
                        {
                            MessageBox.Show(
                               "Для начала добватье столбик в рабочую область, это можно сделать на панели отображения столбцов",
                                ex.ToString(),
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
                        }
                        
                        break;
                }
                
            }
            else
            {
                switch (checkedListBox1.Items[e.Index].ToString())
                {
                    case "Окончание срока план":
                        for (int n = 0; n < dg.Rows.Count - 1; n++)
                        {
                            dg[dg.Columns["Окончание срока план"].Index, n].Style.BackColor = Color.White;
                        }
                        break;
                    case "Фактическое окончание":
                        for (int n = 0; n < dg.Rows.Count - 1; n++)
                        {
                            dg[dg.Columns["Фактическое окончание"].Index, n].Style.BackColor = Color.White;
                        }
                        break;
                    case "Просроченные заявки":
                        for (int n = 0; n < dg.Rows.Count - 1; n++)
                        {
                            dg[dg.Columns["Просроченные заявки"].Index, n].Style.BackColor = Color.White;
                        }
                        break;
                    case "Дней просрочки":
                        for (int n = 0; n < dg.Rows.Count - 1; n++)
                        {
                            dg[dg.Columns["Дней просрочки"].Index, n].Style.BackColor = Color.White;
                        }
                        break;
                    case "Категория штрафов":
                        for (int n = 0; n < dg.Rows.Count - 1; n++)
                        {
                            dg[dg.Columns["Категория штрафов"].Index, n].Style.BackColor = Color.White;
                        }
                        break;
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (tabControl1.TabPages[tabControl1.SelectedIndex].Text != "Добавить документ")
            {
                tabControl1.TabPages[tabControl1.SelectedIndex].Dispose();
                tabControl2.TabPages[tabControl2.SelectedIndex].Dispose();
            }
            
        }

        private void button5_Click(object sender, EventArgs e)
        {
            var dg = tabControl1.TabPages[tabControl1.SelectedIndex].Controls
                    .OfType<DoubleBufferedDataGridView>()
                    .FirstOrDefault(x => x.Name == tabControl1.SelectedIndex.ToString());

            try
            {
                         
            

                (dg.DataSource as DataTable).DefaultView.RowFilter =
                String.Format("" + comboBox2.Text + " like '%{0}%'", textBox1.Text);
            }
            catch(System.Data.SyntaxErrorException ex)
            {
                MessageBox.Show(
                              "В текущей версии программы реализован поиск только для столбцов, в названии которых нет пробелов.",
                               ex.ToString(),
                               MessageBoxButtons.OK,
                               MessageBoxIcon.Error);
            }
            
        }

        private void button6_Click(object sender, EventArgs e)
        {
            var dg = tabControl1.TabPages[tabControl1.SelectedIndex].Controls
                   .OfType<DoubleBufferedDataGridView>()
                   .FirstOrDefault(x => x.Name == tabControl1.SelectedIndex.ToString());



            (dg.DataSource as DataTable).DefaultView.RowFilter = null;
        }

        private void tabControl2_Selected(object sender, TabControlEventArgs e)
        {
            
            tabControl1.SelectedIndex = tabControl2.SelectedIndex;
        }

        private void MainWindow_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }
    }
}
