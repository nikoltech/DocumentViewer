using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Text.RegularExpressions;
using AddTable = System.Int32;
using System.Xml;
using System.IO;
using System.Text;

namespace XMLGenerator
{
    public partial class MainForm : Form
    {
        XDocument document;
        private ContextMenu contextMenu;
        private IEnumerable<XElement> students;
        private XElement Student;
        public MainForm()
        {

            InitializeComponent();
            /*
            textBox19.Value = DateTime.Today;
            textBox20.Value = DateTime.Today;
            textBox22.Value = DateTime.Today;
            textBox23.Value = DateTime.Today;
            textBox21.Value = DateTime.Today;
            textBox24.Value = DateTime.Today;*/

            // Выпадающие списки
            SexBox.SelectedIndex = 0;
            comboBox4.SelectedIndex = 0;
            comboBox1.SelectedIndex = 0;
            comboBox37.SelectedIndex = 0;


            // подгоняем данные из текстового документа в comboBox37
            try {
                comboBox37.Items.Clear();

                var files = from file in Directory.EnumerateFiles(Directory.GetCurrentDirectory(), "*.txt")
                            from line in File.ReadAllLines(file, Encoding.Default)
                            where line.Contains("Institutes:")

                            select new
                            {
                                File = file,
                                Line = line
                            };

                        foreach (var line in files)
                        {
                            //line.Line.Substring(0, line.Line.Length - 2);
                            comboBox37.Items.Add(line.Line.Substring(11));
                        }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Problem with file", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            // подгоняем данные из текстового документа в comboBoxID
            try
            {
                comboBoxID.Items.Clear();

                var files = from file in Directory.EnumerateFiles(Directory.GetCurrentDirectory(), "*.txt")
                            from line in File.ReadAllLines(file, Encoding.Default)
                            where line.Contains("Identity:")

                            select new
                            {
                                File = file,
                                Line = line
                            };

                foreach (var line in files)
                {
                    //line.Line.Substring(0, line.Line.Length - 2);
                    comboBoxID.Items.Add(line.Line.Substring(9));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Problem with file", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


            // подгоняем данные из текстового документа в comboBox1
            try
            {
                comboBox1.Items.Clear();

                var files = from file in Directory.EnumerateFiles(Directory.GetCurrentDirectory(), "*.txt")
                            from line in File.ReadAllLines(file, Encoding.Default)
                            where line.Contains("State:")

                            select new
                            {
                                File = file,
                                Line = line
                            };

                foreach (var line in files)
                {
                    //line.Line.Substring(0, line.Line.Length - 2);
                    comboBox1.Items.Add(line.Line.Substring(6));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Problem with file", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            // подгоняем данные из текстового документа в comboBoxID
            try
            {
                DataGridViewComboBoxColumn a = dataGridView1.Columns["national_mark"] as DataGridViewComboBoxColumn;
                a.Items.Clear();

                var files = from file in Directory.EnumerateFiles(Directory.GetCurrentDirectory(), "*.txt")
                            from line in File.ReadAllLines(file, Encoding.Default)
                            where line.Contains("NationalMark:")

                            select new
                            {
                                File = file,
                                Line = line
                            };

                foreach (var line in files)
                {
                    //line.Line.Substring(0, line.Line.Length - 2);
                    a.Items.Add(line.Line.Substring(13));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Problem with file", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            try
            {
                DataGridViewComboBoxColumn b = dataGridView1.Columns["ects_mark"] as DataGridViewComboBoxColumn;
                b.Items.Clear();

                var files = from file in Directory.EnumerateFiles(Directory.GetCurrentDirectory(), "*.txt")
                            from line in File.ReadAllLines(file, Encoding.Default)
                            where line.Contains("ECTS:")

                            select new
                            {
                                File = file,
                                Line = line
                            };

                foreach (var line in files)
                {
                    //line.Line.Substring(0, line.Line.Length - 2);
                    b.Items.Add(line.Line.Substring(5));
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Problem with file", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


            try
            {
                DataGridViewComboBoxColumn c = dataGridView1.Columns["type"] as DataGridViewComboBoxColumn;
                c.Items.Clear();

                var files = from file in Directory.EnumerateFiles(Directory.GetCurrentDirectory(), "*.txt")
                            from line in File.ReadAllLines(file, Encoding.Default)
                            where line.Contains("Types:")

                            select new
                            {
                                File = file,
                                Line = line
                            };

                foreach (var line in files)
                {
                    //line.Line.Substring(0, line.Line.Length - 2);
                    c.Items.Add(line.Line.Substring(6));
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Problem with file", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            //total
            dataGridView2.Rows.Add(1);

            //Контекстное меню
            contextMenu = new ContextMenu();
            MenuItem cut = new MenuItem("&Вырезать");
            MenuItem copy = new MenuItem("&Копировать");
            MenuItem paste = new MenuItem("&Вставить");
            MenuItem delete = new MenuItem("&Удалить выделенное");
            MenuItem clear = new MenuItem("&Очистить поле");
            cut.Click += new System.EventHandler(this.cut_Click);
            copy.Click += new System.EventHandler(this.copy_Click);
            paste.Click += new System.EventHandler(this.paste_Click);
            delete.Click += new System.EventHandler(this.delete_Click);
            clear.Click += new System.EventHandler(this.clear_Click);
            contextMenu.MenuItems.Add(cut);
            contextMenu.MenuItems.Add("-");
            contextMenu.MenuItems.Add(copy);
            contextMenu.MenuItems.Add(paste);
            contextMenu.MenuItems.Add("-");
            contextMenu.MenuItems.Add(delete);
            contextMenu.MenuItems.Add("-");
            contextMenu.MenuItems.Add("-");
            contextMenu.MenuItems.Add(clear);



            //Соединение контекстного меню с элементами
            foreach (Control tab in this.tabControl1.Controls)
            {
                foreach (var i in tab.Controls)
                {
                    var control = i as RichTextBox;
                    if (control != null)
                    {
                        control.ContextMenu = contextMenu;
                        continue;
                    }
                }

            }

            //Выделяем поля
            //sex
            this.SexBox.BackColor = System.Drawing.Color.Azure;

            //dates
            this.textBox20.BackColor = System.Drawing.Color.Azure;
            this.textBox21.BackColor = System.Drawing.Color.Azure;


            //Область выбора подгружаемых данных
            listView2.View = View.Details;
            listView2.GridLines = true;
            listView2.MultiSelect = false;
            listView2.Columns.Add("ФИО");
            listView2.Columns[0].Width = listView2.Width - 4 - 17;

            

        }
        
        // Загрузить документ в таблицу
        private void LoadButton_Click(object sender, EventArgs e)
        {

            OpenFileDialog loadFile = new OpenFileDialog();
            loadFile.DefaultExt = "*.xml";
            loadFile.Filter = "ONPU DIPLOMA FILES |*.xml";

            if (loadFile.ShowDialog() == DialogResult.OK && loadFile.FileName.Length > 0)
            {
                try
                {
                    document = XDocument.Load(loadFile.FileName);
                }
                catch (Exception exeption)
                {
                    MessageBox.Show(exeption.Message, "Загрузка XML доккумента", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                try
                {
                    // подгрузка данных в поля / lvl j
                    //LoadToFields(document);
                    this.listView2_updateItems();
                    /*XElement student = students.FirstOrDefault();
                    loadStudentToFieldsSingle(student);*/
                    this.tabControl1.SelectedTab = this.tabControl1.TabPages[0];
                    updateBtnDisable();
                    ComboBoxIDSetDefaultValue();
                }
                catch(Exception ex)
                {
                    throw ex;
                    MessageBox.Show("Обнаружено несовпадение в форматах данных. \nПроцесс загрузки данных приостановлен.");
                }

            }
        }

        private void LoadToFields(XDocument document)
        {
            XElement updateDoc = document.Element("documents").Element("document");

            //university
            XElement university = document.Descendants("university").FirstOrDefault();
            this.richTextBox7.Text = university.Attribute("uk").Value;
            this.richTextBox6.Text = university.Attribute("en").Value;
            this.textBox4.Text = university.Attribute("ownershipuk").Value;
            this.textBox6.Text = university.Attribute("ownershipen").Value;
            //faculty
            XElement faculty = document.Descendants("faculty").FirstOrDefault();
            this.comboBox37.Text = faculty.Attribute("uk").Value;
            //scale
            XElement scale = document.Descendants("scale").FirstOrDefault();
            if (scale != null)
                this.ScaleBox.Text = scale.Value == "4" ? "4-бальна" : (scale.Value == "7" ? "7-бальна" : (scale.Value == "5" ? "5-бальна(радянська)" : ""));
            else
                updateDoc.Add(new XElement("scale", ""));
            //educationformname
            XElement educationformname = document.Descendants("educationformname").FirstOrDefault();
            this.comboBox4.Text = educationformname.Attribute("uk").Value;
            //studylanguage
            XElement studylanguage = document.Descendants("studylanguage").FirstOrDefault();
            this.richTextBox9.Text = studylanguage.Attribute("uk").Value;
            this.richTextBox8.Text = studylanguage.Attribute("en").Value;
            //qualification
            XElement qualification = document.Descendants("qualification").FirstOrDefault();
            this.comboBox1.Text = qualification.Attribute("degreeuk").Value;
            this.richTextBox1.Text = qualification.Attribute("specialityuk").Value;
            this.richTextBox2.Text = qualification.Attribute("specialityen").Value;
            this.richTextBox4.Text = qualification.Attribute("profqualificationuk").Value;
            this.richTextBox3.Text = qualification.Attribute("profqualificationen").Value;
            //fieldofstudy
            XElement fieldofstudy = document.Descendants("fieldofstudy").FirstOrDefault();
            this.richTextBox5.Text = fieldofstudy.Attribute("uk").Value;
            this.richTextBox38.Text = fieldofstudy.Attribute("en").Value;
            //levelofqualification
            XElement levelofqualification = document.Descendants("levelofqualification").FirstOrDefault();
            this.richTextBox11.Text = levelofqualification.Attribute("uk").Value;
            this.richTextBox10.Text = levelofqualification.Attribute("en").Value;
            //studyduration
            XElement studyduration = document.Descendants("studyduration").FirstOrDefault();
            this.textBox2.Text = studyduration.Attribute("years").Value;
            this.textBox3.Text = studyduration.Attribute("months").Value;
            //accessrequirements
            XElement accessrequirements = document.Descendants("accessrequirements").FirstOrDefault();
            this.richTextBox13.Text = accessrequirements.Attribute("mainuk").Value;
            this.richTextBox12.Text = accessrequirements.Attribute("mainen").Value;
            //accessrequirements additionaluk ??    Для каждого студента свой документ о предыдущем образовании!  Уточнить 
           /* this.richTextBox21.Text = accessrequirements.Attribute("additionaluk").Value;
            //accessrequirements additionalen ??
            this.richTextBox20.Text = accessrequirements.Attribute("additionalen").Value;*/
            
            //accessfurtherstudy
            XElement accessfurtherstudy = document.Descendants("accessfurtherstudy").FirstOrDefault();
            this.richTextBox15.Text = accessfurtherstudy.Attribute("uk").Value;
            this.richTextBox14.Text = accessfurtherstudy.Attribute("en").Value;
            //professionalstatus
            XElement professionalstatus = document.Descendants("professionalstatus").FirstOrDefault();
            this.richTextBox17.Text = professionalstatus.Attribute("uk").Value;
            this.richTextBox16.Text = professionalstatus.Attribute("en").Value;
            //sign
            XElement sign = document.Descendants("sign").FirstOrDefault();
            this.textBox8.Text = sign.Attribute("positionuk").Value;
            this.textBox7.Text = sign.Attribute("positionen").Value;
            this.textBox10.Text = sign.Attribute("signernameuk").Value;
            this.textBox9.Text = sign.Attribute("signernameen").Value;
            //programmerequirements
            //learnermustsutisfy
            //sutisfy
            XElement sutisfy = document.Descendants("sutisfy").FirstOrDefault();
            if(sutisfy == null)
            {
                document.Descendants("learnermustsutisfy").FirstOrDefault().Add(
                    new XElement("sutisfy",
                        new XAttribute("uk", ""),
                        new XAttribute("en", "")) );
                sutisfy = document.Descendants("sutisfy").FirstOrDefault();
            }
            this.richTextBox29.Text = sutisfy.Attribute("uk").Value;
            this.richTextBox28.Text = sutisfy.Attribute("en").Value;
            //knowledgeundestanding
            //knowledge
            XElement knowledge = document.Descendants("knowledge").FirstOrDefault();
            this.richTextBox31.Text = knowledge.Attribute("uk").Value;
            this.richTextBox30.Text = knowledge.Attribute("en").Value;
            //applyingknowledgeunderstanding
            //understanding
            XElement understanding = document.Descendants("understanding").FirstOrDefault();
            this.richTextBox33.Text = understanding.Attribute("uk").Value;
            this.richTextBox32.Text = understanding.Attribute("en").Value;
            //makingjudgments
            //judgments
            XElement judgments = document.Descendants("judgments").FirstOrDefault();
            this.richTextBox35.Text = judgments.Attribute("uk").Value;
            this.richTextBox34.Text = judgments.Attribute("en").Value;
            //end programmerequirements


            // Данные о студенте. При патчах обратить внимание!!!!! SingleOrDefault | FirstOrDefault
            //student
            //this.students = document.Descendants("student");
            this.listView2_updateItems();

            
        }

        // Создать XML-файл
        private void CreateButton_Click(object sender, EventArgs e)
        {
            if(this.updateStud.Enabled)
            {
                string message = "Ви не підтвердили зміни. Бажаєте продовжити запис до файлу?";
                string caption = "Редагування даних";
                MessageBoxButtons buttons = MessageBoxButtons.YesNo;
                DialogResult result;

                // Displays the MessageBox.

                result = MessageBox.Show(message, caption, buttons);

                if (result == System.Windows.Forms.DialogResult.No)
                {
                    return;
                }
            }

            SaveFileDialog sfile = new SaveFileDialog();
            sfile.DefaultExt = "*.xml";
            sfile.Filter = "ONPU DIPLOMA FILES |*.xml";
            if (sfile.ShowDialog() == DialogResult.OK && sfile.FileName.Length > 0)
            {
                if (document == null) this.createXDocFromFields();
                document.Save(sfile.FileName);
                this.loadedFIO.Text = "";

                ClearAllFields();
                updateBtnDisable();
            }

        }

        //Сбор данных с полей в document /додатки
        private void createXDocFromFields()
        {
            string scale;
            switch(this.ScaleBox.Text)
            {
                case "4-бальна": scale = "4"; break;
                    
                case "7-бальна": scale = "7"; break;
                    
                case "5-бальна(радянська)": scale = "5"; break;
                 
                default: scale = ""; break;
                    
            }

            // Описание XML-доккумента, структуру доккумента реализовывать ТОЛЬКО ЗДЕСЬ!
            document = new XDocument(
            new XElement("documents",
               new XElement("document",
                  new XElement("university", // START University
                    new XAttribute("uk", this.richTextBox7.Text),
                    new XAttribute("en", this.richTextBox6.Text),
                    new XAttribute("ownershipuk", this.textBox4.Text),
                    new XAttribute("ownershipen", this.textBox6.Text)
                            ),// END University
                  new XElement("faculty",// START faculty
                    new XAttribute("uk", this.comboBox37.Text),
                    new XAttribute("en", this.instituteEN(this.comboBox37.Text))
                    ),// END faculty
                  new XElement("scale", scale), // END scale
                  new XElement("educationformname", // START educationformname(VipSpisok?)
                    new XAttribute("uk", this.comboBox4.Text),
                    new XAttribute("en", this.comboBox4.Text == "Денна" ? "Full-time" : "Correspondence")
                            ),// END educationformname
                  new XElement("studylanguage",// START studylanguage(VipSpisok?)
                    new XAttribute("uk", this.richTextBox9.Text),
                    new XAttribute("en", this.richTextBox8.Text)
                            ),// END studylanguage
                  new XElement("qualification",// START qualification
                    new XAttribute("degreeuk", this.comboBox1.Text),
                    new XAttribute("degreeen", this.comboBox1.Text == "Бакалавр" ? "Bachelor" : (this.comboBox1.Text == "Спеціаліст" ? "Specialist" : "Master")),
                    new XAttribute("specialityuk", this.richTextBox1.Text),
                    new XAttribute("specialityen", this.richTextBox2.Text),
                    new XAttribute("profqualificationuk", this.richTextBox4.Text),
                    new XAttribute("profqualificationen", this.richTextBox3.Text)
                            ),// END qualification
                  new XElement("fieldofstudy",// START fieldofstudy
                    new XAttribute("uk", this.richTextBox5.Text),
                    new XAttribute("en", this.richTextBox38.Text)
                            ),// END fieldofstudy
                  new XElement("levelofqualification",// START levelofqualification(BigText?)
                    new XAttribute("uk", this.richTextBox11.Text),
                    new XAttribute("en", this.richTextBox10.Text)
                        ),// END levelofqualification
                  new XElement("studyduration",// START studyduration
                    new XAttribute("years", this.textBox2.Text),// Shislo
                    new XAttribute("months", this.textBox3.Text),// Shislo
                    new XAttribute("formnameuk", this.comboBox4.Text == "Денна" ? "Денна навчання" : "Заочна навчання"),
                    new XAttribute("formnameen", this.comboBox4.Text == "Денна" ? "Full-time form of studies" : "Correspondence form of studies")
                        ),// END studyduration
                  new XElement("accessrequirements",// START accessrequirements(BigText?)
                    new XAttribute("mainuk", this.richTextBox13.Text),
                    new XAttribute("mainen", this.richTextBox12.Text),
                    new XAttribute("additionaluk", ""),
                    new XAttribute("additionalen", "")
                        ),// END accessrequirements
                  new XElement("accessfurtherstudy",// START accessfurtherstudy(BigText?)
                    new XAttribute("uk", this.richTextBox15.Text),
                    new XAttribute("en", this.richTextBox14.Text)
                        ),// END accessfurtherstudy
                  new XElement("professionalstatus",// START professionalstatus(BigText?)
                    new XAttribute("uk", this.richTextBox17.Text),
                    new XAttribute("en", this.richTextBox16.Text)
                        ),// END professionalstatus
                  new XElement("sign",// START sign
                    new XAttribute("positionuk", this.textBox8.Text),
                    new XAttribute("positionen", this.textBox7.Text),
                    new XAttribute("signernameuk", this.textBox10.Text),
                    new XAttribute("signernameen", this.textBox9.Text)
                        ),// END sign
                  new XElement("programmerequirements", // START programmerequirements
                    new XElement("learnermustsutisfy",// START learnermustsutisfy
                        new XElement("sutisfy",// START sutisfy
                             new XAttribute("uk", this.richTextBox29.Text),
                             new XAttribute("en", this.richTextBox28.Text)
                             )// END sutisfy
                        ),// END learnermustsutisfy
                    new XElement("knowledgeundestanding",// START knowledgeundestanding
                        new XElement("knowledge",// START knowledge(BigText?)
                             new XAttribute("uk", this.richTextBox31.Text),
                             new XAttribute("en", this.richTextBox30.Text)
                            )// END knowledge
                        ),// END knowledgeundestanding
                    new XElement("applyingknowledgeunderstanding",// START applyingknowledgeunderstanding
                        new XElement("understanding",// START understanding(BigText?)
                             new XAttribute("uk", this.richTextBox33.Text),
                             new XAttribute("en", this.richTextBox32.Text)
                            )// END knowledge
                        ),// END applyingknowledgeunderstanding
                    new XElement("makingjudgments",// START makingjudgments
                        new XElement("judgments",// START judgments(BigText?)
                             new XAttribute("uk", this.richTextBox35.Text),
                             new XAttribute("en", this.richTextBox34.Text)
                            )// END judgments
                        )// END makingjudgments
                    )
              )
            )
        );

            
        }
        //Обновить данные с полей в document
        private bool updateXDoc()
        {
            if (document == null) return false;

            string scale;
            switch (this.ScaleBox.Text)
            {
                case "4-бальна": scale = "4"; break;

                case "7-бальна": scale = "7"; break;

                case "5-бальна(радянська)": scale = "5"; break;

                default: scale = ""; break;

            }

            XElement updateDoc = document.Element("documents").Element("document");

            //university
            updateDoc.Element("university").Attribute("uk").SetValue(this.richTextBox7.Text);
            updateDoc.Element("university").Attribute("en").SetValue(this.richTextBox6.Text);
            updateDoc.Element("university").Attribute("ownershipuk").SetValue(this.textBox4.Text);
            updateDoc.Element("university").Attribute("ownershipen").SetValue(this.textBox6.Text);
            //faculty
            updateDoc.Element("faculty").Attribute("uk").SetValue(this.comboBox37.Text);
            updateDoc.Element("faculty").Attribute("en").SetValue(this.instituteEN(this.comboBox37.Text));
            //scale
            updateDoc.Element("scale").SetValue(scale);
            //educationformname
            updateDoc.Element("educationformname").Attribute("uk").SetValue(this.comboBox4.Text);
            updateDoc.Element("educationformname").Attribute("en").SetValue(this.comboBox4.Text == "Денна" ? "Full-time" : "Correspondence");
            //studylanguage
            updateDoc.Element("studylanguage").Attribute("uk").SetValue(this.richTextBox9.Text);
            updateDoc.Element("studylanguage").Attribute("en").SetValue(this.richTextBox8.Text);
            //qualification
            updateDoc.Element("qualification").Attribute("degreeuk").SetValue(this.comboBox1.Text);
            updateDoc.Element("qualification").Attribute("degreeen").SetValue(this.comboBox1.Text == "Бакалавр" ? "Bachelor" : (this.comboBox1.Text == "Спеціаліст" ? "Specialist" : "Master"));
            updateDoc.Element("qualification").Attribute("specialityuk").SetValue(this.richTextBox1.Text);
            updateDoc.Element("qualification").Attribute("specialityen").SetValue(this.richTextBox2.Text);
            updateDoc.Element("qualification").Attribute("profqualificationuk").SetValue(this.richTextBox4.Text);
            updateDoc.Element("qualification").Attribute("profqualificationen").SetValue(this.richTextBox3.Text);
            //fieldofstudy
            updateDoc.Element("fieldofstudy").Attribute("uk").SetValue(this.richTextBox5.Text);
            updateDoc.Element("fieldofstudy").Attribute("en").SetValue(this.richTextBox38.Text);
            //levelofqualification
            updateDoc.Element("levelofqualification").Attribute("uk").SetValue(this.richTextBox11.Text);
            updateDoc.Element("levelofqualification").Attribute("en").SetValue(this.richTextBox10.Text);
            //studyduration
            updateDoc.Element("studyduration").Attribute("years").SetValue(this.textBox2.Text);
            updateDoc.Element("studyduration").Attribute("months").SetValue(this.textBox3.Text);
            updateDoc.Element("studyduration").Attribute("formnameuk").SetValue(this.comboBox4.Text == "Денна" ? "Денна навчання" : "Заочна навчання");
            updateDoc.Element("studyduration").Attribute("formnameen").SetValue(this.comboBox4.Text == "Денна" ? "Full-time form of studies" : "Correspondence form of studies");
            //accessrequirements
            updateDoc.Element("accessrequirements").Attribute("mainuk").SetValue(this.richTextBox13.Text);
            updateDoc.Element("accessrequirements").Attribute("mainen").SetValue(this.richTextBox12.Text);
            updateDoc.Element("accessrequirements").Attribute("additionaluk").SetValue("");
            updateDoc.Element("accessrequirements").Attribute("additionalen").SetValue("");
            //accessfurtherstudy
            updateDoc.Element("accessfurtherstudy").Attribute("uk").SetValue(this.richTextBox15.Text);
            updateDoc.Element("accessfurtherstudy").Attribute("en").SetValue(this.richTextBox14.Text);
            //professionalstatus
            updateDoc.Element("professionalstatus").Attribute("uk").SetValue(this.richTextBox17.Text);
            updateDoc.Element("professionalstatus").Attribute("en").SetValue(this.richTextBox16.Text);
            //sign
            updateDoc.Element("sign").Attribute("positionuk").SetValue(this.textBox8.Text);
            updateDoc.Element("sign").Attribute("positionen").SetValue(this.textBox7.Text);
            updateDoc.Element("sign").Attribute("signernameuk").SetValue(this.textBox10.Text);
            updateDoc.Element("sign").Attribute("signernameen").SetValue(this.textBox9.Text);

            XElement programmerequirementsUp = updateDoc.Element("programmerequirements");
            //learnermustsutisfy
            programmerequirementsUp.Element("learnermustsutisfy").Element("sutisfy").Attribute("uk").SetValue(this.richTextBox29.Text);
            programmerequirementsUp.Element("learnermustsutisfy").Element("sutisfy").Attribute("en").SetValue(this.richTextBox28.Text);
            //knowledgeundestanding
            programmerequirementsUp.Element("knowledgeundestanding").Element("knowledge").Attribute("uk").SetValue(this.richTextBox31.Text);
            programmerequirementsUp.Element("knowledgeundestanding").Element("knowledge").Attribute("en").SetValue(this.richTextBox30.Text);
            //applyingknowledgeunderstanding
            programmerequirementsUp.Element("applyingknowledgeunderstanding").Element("understanding").Attribute("uk").SetValue(this.richTextBox33.Text);
            programmerequirementsUp.Element("applyingknowledgeunderstanding").Element("understanding").Attribute("en").SetValue(this.richTextBox32.Text);

            programmerequirementsUp.Element("makingjudgments").Element("judgments").Attribute("uk").SetValue(this.richTextBox35.Text);
            programmerequirementsUp.Element("makingjudgments").Element("judgments").Attribute("en").SetValue(this.richTextBox34.Text);

            return true;
        }
        private string instituteEN(string input)
        {
            string FoH = "";
            switch (input)
            {
                case "Гуманітарний факультет": FoH = "Faculty of Humanities";
                    break;
                case "Інститут бізнесу економіки та інформаційних технологій": FoH = "Institute of Business Economics and Information Technology";
                    break;
                case "Інститут електромеханіки та енергоменеджменту": FoH = "Institute of Electromechanics and Power";
                    break;
                case "Інститут енергетики і комп'ютерно-інтегрованих систем управління": FoH = "Institute of Power Engineering and Computer Integrated Management Systems";
                    break;
                case "Інститут інформаційної безпеки, радіоелектроніки та телекомунікацій": FoH = "Institute of Information Security, electronics and telecommunications";
                    break;
                case "Інститут комп'ютерних систем": FoH = "Institute of Computer Systems";
                    break;
                case "Інститут машинобудування": FoH = "Institute of engineering";
                    break;
                case "Інститут промислових технологій дизайну та менеджменту": FoH = "Industrial Technologies Design and Management";
                    break;
                case "Хіміко-технологічний факультет": FoH = "Chemical Engineering Department";
                    break;
                case "Інститут медичної інженерії": FoH = "Institute of Medical Engineering";
                    break;
                case "Інститут дистанційної і заочної освіти": FoH = "Institute of Distance and distance education";
                    break;
                case "Українсько-німецький інститут": FoH = "Ukrainian-German Institute";
                    break;
                case "Українсько-польський інститут": FoH = "Ukrainian-Polish Institute";
                    break;
                case "Українсько-іспанський інститут": FoH = "Ukrainian-Spanish Institute";
                    break;
                default: break;
            }
            return FoH;
        }

        //Видалити поле
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                AddTable rowToDelete = this.dataGridView1.Rows.GetFirstRow(DataGridViewElementStates.Selected);
                if (rowToDelete > -1)
                {
                    this.dataGridView1.Rows.RemoveAt(rowToDelete);
                }
            }
            catch (InvalidOperationException)
            {
                MessageBox.Show("Для удаления строки неободимо убедится что она существует и была выделена");
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.dataGridView1.Rows.Add(1);
        }

        //Clear all fields
        private void button5_Click(object sender, EventArgs e)
        {
            ClearAllFields();
        }

        private void ClearAllFields()
        {
            foreach (Control tab in this.tabControl1.Controls)
            {
                foreach (var i in tab.Controls)
                {
                    var control = i as RichTextBox;
                    if (control != null)
                    {
                        control.Text = "";
                    }
                    var control2 = i as TextBox;
                    if (control2 != null)
                    {
                        control2.Text = "";
                    }
                }

            }

            this.textBox10.Text = "Оборський Геннадій Олександрович";
            this.textBox2.Text = "0";
            this.textBox3.Text = "0";
            this.textBox4.Text = "Державна";
            this.textBox6.Text = "State";
            this.textBox7.Text = "Rector";
            this.textBox8.Text = "Ректор";
            this.textBox9.Text = "Oborskyi Hennadii Oleksandrovych";

            this.richTextBox6.Text = "Odessa National Polytechnic University";
            this.richTextBox7.Text = "Одеський національний політехнічний університет";
            this.richTextBox8.Text = "ukrainian";
            this.richTextBox9.Text = "українська";
            this.richTextBox4.Text = "-";
            this.richTextBox3.Text = "-";

            this.comboBox1.SelectedIndex = 0;
            this.comboBox37.SelectedIndex = 0;
            this.comboBox4.SelectedIndex = 0;
            this.SexBox.SelectedIndex = 0;
            this.ComboBoxIDSetDefaultValue();

            this.checkBox1.Checked = false;
            this.checkBox2.Checked = false;
            this.checkBox3.Checked = false;
            this.checkBox4.Checked = false;

            this.dataGridView1.Rows.Clear();
            this.dataGridView2.Rows.Clear();
            this.dataGridView2.Rows.Add(1);

            switch (this.comboBox1.Text)
            {
                case "Бакалавр":
                    this.richTextBox29.Text = "30 кредитів ECTS у семестр за акредитованою бакалаврською програмою, атестаційний екзамен та дипломна робота (проект)";
                    this.richTextBox28.Text = "30 ECTS credits per every semester of Bachelor`s degree, training adopted program, plus the graduation thesis (Diploma) and state qualification examination";
                    this.label12.Text = "\"Доповнення\"->\"Вимоги освітньої програми..\" сформовані для Бакалавра";
                    break;
                case "Спеціаліст":
                    this.richTextBox29.Text = "30 кредитів ECTS у семестр за акредитованою спеціалістською програмою, дипломна робота (проект)";
                    this.richTextBox28.Text = "30 ECTS credits per every semester of Specialist`s degree, training adopted program, plus the graduation thesis (Diploma)";
                    this.label12.Text = "\"Доповнення\"->\"Вимоги освітньої програми..\" сформовані для Спеціаліста";
                    break;
                case "Магістр":
                    this.richTextBox29.Text = "30 кредитів ECTS у семестр за акредитованою магістерською програмою, дипломна робота (проект)";
                    this.richTextBox28.Text = "30 ECTS credits per every semester of Master`s degree, training adopted program, plus the graduation thesis (Diploma)";
                    this.label12.Text = "\"Доповнення\"->\"Вимоги освітньої програми..\" сформовані для Магістра";
                    break;
                default:
                    this.richTextBox29.Text = "";
                    this.richTextBox28.Text = "";
                    this.label12.Text = " ";
                    break;
            }

            this.loadedFIO.Text = "";
            

        }

        private void dataGridView1_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            if (dataGridView1.Rows[e.RowIndex].IsNewRow) { return; }
            if (e.ColumnIndex == 2) //Кредити / Цифры и точка
            {
                string str = e.FormattedValue.ToString();
                if (Regex.IsMatch(str, @"^([0-9]*\.[0-9]+)$"))
                {
                    this.dataGridView1.Rows[e.RowIndex].ErrorText = "";
                    str.Replace(',', '.'); // doesnt work
                }
                else
                {
                    this.dataGridView1.Rows[e.RowIndex].ErrorText = "Неверный формат записи. Для поля \"Кредити\" доступны дробные записи с точкой";
                    //e.Cancel = true;
                }

            }
            if (e.ColumnIndex == 3 || e.ColumnIndex == 4) //Години або Оцінка / Только цифры
            {
                string str = e.FormattedValue.ToString();
                if (Regex.IsMatch(str, @"^([0-9])"))
                {
                    this.dataGridView1.Rows[e.RowIndex].ErrorText = "";

                    if (e.ColumnIndex == 4)     //Оцінка - генерация national_mark и ects_mark
                    {
                        this.dataGridView1.Rows[e.RowIndex].Cells["national_mark"].Value = ConvertMarkToNationalMark(str);
                        this.dataGridView1.Rows[e.RowIndex].Cells["ects_mark"].Value = ConvertMarkToEctsMark(str);
                    }
                }
                else
                {
                    this.dataGridView1.Rows[e.RowIndex].ErrorText = "Неверный формат записи. Для поля \"Години\" и \"Оцінка\"Доступны только цифры";
                    //e.Cancel = true;
                }
            }
            

            

        }

       /*
        * national_mark
        * 
            Відмінно/Excellent
            Добре/Good
            Дуже добре/Very good
            Задовільно/Satisfactory
            Достатньо/Enough
            Зараховано/Passed
            Незадовільно/Fail
            Не зараховано/Fail
        * 
        * 
        * ects_mark
        * 
            A
            B
            C
            D
            E
            F
        * 
        */

        //Преобразуем Оценку в Оценку по национальной шкале
        private string ConvertMarkToNationalMark(string mark)
        {
            string nationalMark = "";
            AddTable markForConvert;
            try
            {
                markForConvert = AddTable.Parse(mark);
            }
            catch (Exception)
            {
                MessageBox.Show("Ошибка при конвертации Оценки в Национальную шкалу \n Проверьте введенные данные и попробуйте снова");
                return "F";
            }

            if (markForConvert >= 90)
            {
                nationalMark = "Відмінно/Excellent";
            }
            else
            if (markForConvert >= 82 && markForConvert <= 89)
            {
                nationalMark = "Добре/Good";
            }
            else
            if (markForConvert >= 75 && markForConvert <= 81)
            {
                nationalMark = "Добре/Good";
            }
            else
            if (markForConvert >= 67 && markForConvert <= 74)
            {
                nationalMark = "Задовільно/Satisfactory";
            }
            else
            if (markForConvert >= 60 && markForConvert <= 66)
            {
                nationalMark = "Задовільно/Satisfactory";
            }
            else
            if (markForConvert <= 59)
            {
                nationalMark = "Незадовільно/Fail";
            }

            return nationalMark;
        }
        //Преобразуем Оценку в Оценку по ЕКТС
        private string ConvertMarkToEctsMark(string mark)
        {
            string EctsMark = "";
            AddTable markForConvert;
            try
            {
                markForConvert = AddTable.Parse(mark);
            }
            catch(Exception)
            {
                MessageBox.Show("Ошибка при конвертации Оценки в ЕКТС \n Проверьте введенные данные и попробуйте снова");
                return "F";
            }

            if (markForConvert >= 90)
            {
                EctsMark = "A";
            }
            else
            if (markForConvert >= 82 && markForConvert <= 89)
            {
                EctsMark = "B";
            }
            else
            if (markForConvert >= 75 && markForConvert <= 81)
            {
                EctsMark = "C";
            }
            else
            if (markForConvert >= 67 && markForConvert <= 74)
            {
                EctsMark = "D";
            }
            else
            if (markForConvert >= 60 && markForConvert <= 66)
            {
                EctsMark = "E";
            }
            else
            if (markForConvert <= 59)
            {
                EctsMark = "F";
            }

            return EctsMark;
        }

        //Одинарный клик по выпадающему списку в таблице /#comboClick
        private void datagridview1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {/*
            var datagridview = sender as DataGridView;
            bool validClick = (e.RowIndex != -1 && e.ColumnIndex != -1 && e.RowIndex != 0); //Make sure the clicked row/column is valid.

            // Check to make sure the cell clicked is the cell containing the combobox 
            if (datagridview.Columns[e.ColumnIndex] is DataGridViewComboBoxColumn && validClick)
            {
                datagridview.BeginEdit(true);
                ((ComboBox)datagridview.EditingControl).DroppedDown = true;
            }*/
        }
        //Розрахувати
        private void button6_Click(object sender, EventArgs e)
        {
            double sumCredits = 0;
            double sumHours = 0;
            double sumMarkSubj = 0;
            int markCount = 0;
            int attestCount = 0;
            double sumMarkAttest = 0;
            try
            {
                foreach (DataGridViewRow row in this.dataGridView1.Rows)
                {
                    if (!row.IsNewRow)
                    {
                        sumCredits += Convert.ToDouble(row.Cells["credits"].Value.ToString().Replace('.', ','));
                        sumHours += Convert.ToDouble(row.Cells["hours"].Value.ToString().Replace('.', ','));
                        if (row.Cells["type"].Value.ToString().Equals("Аттестация"))
                        {
                            sumMarkAttest += Convert.ToDouble(row.Cells["mark"].Value.ToString().Replace('.', ','));
                            attestCount++;
                        }
                        else
                        {
                            sumMarkSubj += Convert.ToDouble(row.Cells["mark"].Value.ToString().Replace('.', ','));
                            markCount++;
                        }

                    }

                }
            }
            catch (Exception)
            {
                MessageBox.Show("Все поля таблицы должны быть введены");
                return;
            }
            if (this.dataGridView1.RowCount == 1) return; //если строк с данными нету - выход
            int isatt = 0;
            if (markCount != 0) isatt++;
            if (attestCount != 0) isatt++;
            markCount = markCount == 0 ? 1 : markCount;
            attestCount = attestCount == 0 ? 1 : attestCount;
            this.dataGridView2.Rows[0].Cells["totalCredits"].Value = sumCredits.ToString(".#").Replace(',', '.');
            this.dataGridView2.Rows[0].Cells["totalHours"].Value = sumHours.ToString(".#").Replace(',', '.');
            this.dataGridView2.Rows[0].Cells["middleMark"].Value = ((sumMarkSubj / markCount + sumMarkAttest / attestCount) / isatt).ToString(".#").Replace(',', '.');

        }


        //Контекстное меню
        private void cut_Click(object sender, System.EventArgs e)
        {
            RichTextBox richTextBox = ((sender as MenuItem).Parent as ContextMenu).SourceControl as RichTextBox;
            Clipboard.SetText(richTextBox.SelectedText);
            richTextBox.SelectedText = "";
        }
        private void copy_Click(object sender, System.EventArgs e)
        {
            RichTextBox richTextBox = ((sender as MenuItem).Parent as ContextMenu).SourceControl as RichTextBox;
            Clipboard.SetText(richTextBox.SelectedText);
        }
        private void paste_Click(object sender, System.EventArgs e)
        {
            RichTextBox richTextBox = ((sender as MenuItem).Parent as ContextMenu).SourceControl as RichTextBox;
            richTextBox.SelectedText = Clipboard.GetText();
        }
        private void delete_Click(object sender, System.EventArgs e)
        {
            RichTextBox richTextBox = ((sender as MenuItem).Parent as ContextMenu).SourceControl as RichTextBox;
            richTextBox.SelectedText = "";
        }
        private void clear_Click(object sender, System.EventArgs e)
        {
            RichTextBox richTextBox = ((sender as MenuItem).Parent as ContextMenu).SourceControl as RichTextBox;
            richTextBox.Text = "";
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch ((sender as ComboBox).Text)
            {
                case "Бакалавр":
                    this.richTextBox29.Text = "30 кредитів ECTS у семестр за акредитованою бакалаврською програмою, атестаційний екзамен та дипломна робота (проект)";
                    this.richTextBox28.Text = "30 ECTS credits per every semester of Bachelor`s degree, training adopted program, plus the graduation thesis (Diploma) and state qualification examination";
                    this.label12.Text = "\"Доповнення\"->\"Вимоги освітньої програми..\" сформовані для Бакалавра";
                    break;
                case "Спеціаліст":
                    this.richTextBox29.Text = "30 кредитів ECTS у семестр за акредитованою спеціалістською програмою, дипломна робота (проект)";
                    this.richTextBox28.Text = "30 ECTS credits per every semester of Specialist`s degree, training adopted program, plus the graduation thesis (Diploma)";
                    this.label12.Text = "\"Доповнення\"->\"Вимоги освітньої програми..\" сформовані для Спеціаліста";
                    break;
                case "Магістр":
                    this.richTextBox29.Text = "30 кредитів ECTS у семестр за акредитованою магістерською програмою, дипломна робота (проект)";
                    this.richTextBox28.Text = "30 ECTS credits per every semester of Master`s degree, training adopted program, plus the graduation thesis (Diploma)";
                    this.label12.Text = "\"Доповнення\"->\"Вимоги освітньої програми..\" сформовані для Магістра";
                    break;
                default:
                    this.richTextBox29.Text = "";
                    this.richTextBox28.Text = "";
                    this.label12.Text = " ";
                    break;

            }
        }

        //Попередній перегляд
        /*private void button7_Click(object sender, EventArgs e)
        {
            treeView1.Nodes.Clear();
            if (document == null) createXDocFromFields();
            TreeNode root = new TreeNode(document.Root.Name.ToString());
            treeView1.Nodes.Add(root);

            ReadNode(document.Root, root);
            treeView1.ExpandAll();
        }*/

        private void ReadNode(XElement xElement, TreeNode treeNode)
        {
            foreach (XElement element in xElement.Elements())
            {
                TreeNode node = new TreeNode(element.Name.ToString());
                treeNode.Nodes.Add(node);

                if (element.HasAttributes)
                {
                    TreeNode attributesNode = new TreeNode("Attributes");
                    ReadAttributes(element, attributesNode);
                    node.Nodes.Add(attributesNode);
                }

                ReadNode(element, node);
            }
        }

        private void ReadAttributes(XElement element, TreeNode treeNode)
        {
            foreach (XAttribute attribute in element.Attributes())
            {
                TreeNode node = new TreeNode(attribute.ToString());
                treeNode.Nodes.Add(node);
            }
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            this.richTextBox27.Text += this.dateTimePicker4.Text + "-" + this.dateTimePicker2.Text;
            this.richTextBox26.Text += this.dateTimePicker4.Text + "-" + this.dateTimePicker2.Text;
        }

        //Зберегти для диплома
        private void create_for_dyplome_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfile = new SaveFileDialog();
            sfile.DefaultExt = "*.xml";
            sfile.Filter = "ONPU DIPLOMA FILES |*.xml";
            if (sfile.ShowDialog() == DialogResult.OK && sfile.FileName.Length > 0)
            {
                if (document == null) this.createXDocFromFields(); // переделать на создание для формата диплома
                //XDocument d = document.Document; //для совместимости?
                readFieldsToDyplomeXDoc();
                document.Save(sfile.FileName);
                //document = d.Document; //для совместимости?
                this.loadedFIO.Text = "";
                ComboBoxIDSetDefaultValue();
                ClearAllFields();
            }
        }
        private void readFieldsToDyplomeXDoc()
        {
            /*this.textBox20.CustomFormat = "dd/MM/yyyy";
            this.textBox23.CustomFormat = "dd/MM/yyyy";
            this.textBox24.CustomFormat = "dd/MM/yyyy";
            this.textBox21.CustomFormat = "dd/MM/yyyy";*/

            document = new XDocument(
                new XElement("Documents",
                    new XElement("Document",
                        //new XElement("repeat", this.checkBox3.Checked ? this.checkBox4.Checked : false),

                        // Constants
                        new XElement("DocumentSourceId", "0"),
                        new XElement("DocumentSourceName", "XMLGenerator"),
                        new XElement("DocumentCode", "0"),
                        new XElement("PersonEducationId", "0"),
                        new XElement("GraduateDate", this.textBox21.Text),
                        new XElement("UniversityId", "203"),
                        new XElement("DocumentTypeId", this.comboBox1.SelectedIndex == 0 ? "6" : this.comboBox1.SelectedIndex == 1 ? "7" : "8"),
                        new XElement("DocumentTypeName", this.comboBox1.SelectedIndex == 0 ? "Диплом бакалавра державного зразка" : this.comboBox1.SelectedIndex == 1 ? "Диплом спеціаліста державного зразка" : "Диплом магістра державного зразка"),
                        new XElement("DocumentSeries", this.seriaBox.Text),
                        new XElement("DocumentNumber", this.numberBox.Text),
                        new XElement("PersonCode", "0"),
                        new XElement("LastName", this.textBox1.Text),
                        new XElement("FirstName", this.textBox11.Text),
                        new XElement("MiddleName", this.textBox12.Text),
                        new XElement("LastNameEn", this.textBox13.Text),
                        new XElement("FirstNameEn", this.textBox14.Text),
                        new XElement("MiddleNameEn", this.textBox15.Text),
                        new XElement("Birthday", this.textBox24.Text),
                        new XElement("SexId", this.SexBox.Text == "Чоловiча" ? "1" : "2"),
                        new XElement("SexName", this.SexBox.Text),
                        new XElement("AwardTypeId", this.checkBox2.Checked ? "1" : "0"),
                        new XElement("AwardTypeName", ""),
                        new XElement("PaymentTypeId", "0"),
                        new XElement("PaymentTypeName", ""),
                        new XElement("IsBenefits", ""),
                        new XElement("EducationFormId", this.comboBox4.Text == "Денна" ? "1" : "2"),
                        new XElement("EducationFormName", this.comboBox4.Text),
                        new XElement("PersonDocumentTypeId", "0"),
                        new XElement("PersonDocumentTypeName", ""),
                        new XElement("PersonDocumentNumber", "0"),
                        new XElement("INN", "0"),
                        new XElement("IsDuplicate", this.checkBox3.Checked ? "True" : "False"),
                        new XElement("CreateDate", this.textBox20.Text),
                        new XElement("IssueDate", this.textBox23.Text),
                        new XElement("UniversityPrintName", this.richTextBox7.Text),
                        new XElement("UniversityPrintNameEn", this.richTextBox6.Text),
                        new XElement("UniversityName", this.richTextBox7.Text),
                        new XElement("UniversityNameEn", this.richTextBox6.Text),
                        new XElement("BossPost", this.textBox8.Text),
                        new XElement("BossPostEn", this.textBox7.Text),
                        new XElement("BossFIO", this.textBox10.Text),
                        new XElement("BossFIOEn", this.textBox9.Text),
                        new XElement("SpecialityName", this.richTextBox1.Text),
                        new XElement("SpecialityNameEn", this.richTextBox2.Text),
                        new XElement("SpecializationName", this.richTextBox4.Text),
                        new XElement("SpecializationNameEn", this.richTextBox3.Text),
                        new XElement("FacultyName", this.comboBox37.Text)
                        ))
                );
          /*  this.textBox20.Text = "dd.MM.yyyy";Time create see at this
            this.textBox23.Text = "dd.MM.yyyy";
            this.textBox24.Text = "dd.MM.yyyy";
            this.textBox21.Text = "dd.MM.yyyy";*/
        }


        private void load_for_dyplome_Click(object sender, EventArgs e)
        {
            OpenFileDialog loadFile = new OpenFileDialog();
            loadFile.DefaultExt = "*.xml";
            loadFile.Filter = "ONPU DIPLOMA FILES |*.xml";

            if (loadFile.ShowDialog() == DialogResult.OK && loadFile.FileName.Length > 0)
            {
                try
                {
                    document = XDocument.Load(loadFile.FileName);
                }
                catch (Exception exeption)
                {
                    MessageBox.Show(exeption.Message, "Загрузка XML доккумента", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                try
                {
                    // подгрузка данных в поля / lvl j
                    XElement documentdata = document.Descendants("Document").FirstOrDefault();
                    if (documentdata == null) return;
                   this.textBox23.Text = documentdata.Element("IssueDate").Value;
                    this.textBox24.Text = documentdata.Element("Birthday").Value;
                    this.textBox20.Text = documentdata.Element("CreateDate").Value;
                    this.textBox21.Text = documentdata.Element("GraduateDate").Value;
                    this.checkBox3.Checked = (documentdata.Element("IsDuplicate") != null) ? (documentdata.Element("IsDuplicate").Value == "True") : false;
                    this.comboBox37.Text = documentdata.Element("FacultyName").Value;

                    //seria number
                    this.seriaBox.Text = documentdata.Element("DocumentSeries").Value;
                    this.numberBox.Text = documentdata.Element("DocumentNumber").Value;
                    //honour
                    this.checkBox2.Checked = documentdata.Element("AwardTypeId").Value == "1";
                    //fio ua
                    this.textBox1.Text = documentdata.Element("LastName").Value;
                    this.textBox11.Text = documentdata.Element("FirstName").Value;
                    this.textBox12.Text = documentdata.Element("MiddleName").Value;
                    //fio en
                    this.textBox13.Text = documentdata.Element("LastNameEn").Value;
                     this.textBox14.Text = documentdata.Element("FirstNameEn").Value;
                    this.textBox15.Text = documentdata.Element("MiddleNameEn").Value;
                    //sex
                    this.SexBox.SelectedIndex = documentdata.Element("SexId").Value == "1" ? 0 : 1;
                    //university
                   
                    //education spec/qualification
                   // this.richTextBox4.Text = documentdata.Element("SpecializationName").Value;
                   // this.richTextBox3.Text = documentdata.Element("SpecializationNameEn").Value;
                    this.richTextBox1.Text = documentdata.Element("SpecialityName").Value;
                    this.richTextBox2.Text = documentdata.Element("SpecialityNameEn").Value;
                    this.richTextBox7.Text = documentdata.Element("UniversityPrintName").Value;
                    this.richTextBox6.Text = documentdata.Element("UniversityPrintNameEn").Value;
                    //boss
                    this.textBox8.Text = documentdata.Element("BossPost").Value;
                    this.textBox7.Text = documentdata.Element("BossPostEn").Value;
                    this.textBox10.Text = documentdata.Element("BossFIO").Value;
                    this.textBox9.Text = documentdata.Element("BossFIOEn").Value;
                     this.comboBox4.SelectedIndex = documentdata.Element("EducationFormId").Value == "Денна" ? 0 : 1;
                    //dates
                    /*this.textBox20.CustomFormat = "dd/MM/yyyy";
                    this.textBox24.CustomFormat = "dd/MM/yyyy";
                    this.textBox23.CustomFormat = "dd/MM/yyyy";
                    this.textBox21.CustomFormat = "dd/MM/yyyy";*/

                   
                   
                  /*  if (documentdata.Element("IssueDate") != null)
                    {
                        
                    }
                    */
                    /*this.textBox20.CustomFormat = "dd.MM.yyyy";
                    this.textBox23.CustomFormat = "dd.MM.yyyy";
                    this.textBox24.CustomFormat = "dd.MM.yyyy";
                    this.textBox21.CustomFormat = "dd.MM.yyyy";*/

                    this.loadedFIO.Text = "";
                    this.ComboBoxIDSetDefaultValue();
                }
                catch(Exception)
                {
                    MessageBox.Show("Обнаружено несовпадение в форматах данных. \nПроцесс загрузки данных приостановлен.");
                }

            }
        }

       /* private void button8_Click(object sender, EventArgs e)
        {
            treeView1.Nodes.Clear();
            readFieldsToDyplomeXDoc();
            TreeNode root = new TreeNode(document.Root.Name.ToString());
            treeView1.Nodes.Add(root);

            ReadNode(document.Root, root);
            treeView1.ExpandAll();
        }*/

        //Загружаем данные студента в поля. Формат додатков
        public void loadStudentToFieldsSingle(XElement student)
        {
            //подгоняем под формат xml
            /*this.textBox20.CustomFormat = "dd/MM/yyyy";
            this.textBox21.CustomFormat = "dd/MM/yyyy";
            this.textBox22.CustomFormat = "dd/MM/yyyy";
            this.textBox23.CustomFormat = "dd/MM/yyyy";*/

            this.loadedFIO.Text = student.Attribute("lastnameuk").Value + "  " + student.Attribute("firstnameuk").Value + " " + student.Attribute("middlenameuk").Value;
            this.textBox1.Text = student.Attribute("lastnameuk").Value;
            this.textBox11.Text = student.Attribute("firstnameuk").Value;
            this.textBox12.Text = student.Attribute("middlenameuk").Value;
            this.textBox13.Text = student.Attribute("lastnameen").Value;
            this.textBox14.Text = student.Attribute("firstnameen").Value;
            this.textBox15.Text = student.Attribute("middlenameen").Value;
            this.SexBox.Text = (student.Attribute("sex").Value == "1") ? "Чоловiча" : "Жiноча";
            this.textBox24.Text = student.Attribute("birthday").Value;
            this.checkBox1.Checked = (student.Attribute("foreigner").Value == "1");
            //diplom
            XElement diplom = student.Descendants("diplom").FirstOrDefault();
            this.seriaBox.Text = diplom.Attribute("seria").Value;
            this.numberBox.Text = diplom.Attribute("number").Value;
            this.checkBox2.Checked = (diplom.Attribute("honour").Value == "1");
            this.checkBox3.Checked = (diplom.Attribute("IsDuplicate") != null) ? (diplom.Attribute("IsDuplicate").Value == "True") : false;
            this.checkBox4.Checked = (diplom.Attribute("repeatDodatok") != null) ? (diplom.Attribute("repeatDodatok").Value == "1") : false;
            this.textBox21.Text = diplom.Attribute("givendate").Value;
            if ( diplom.Attribute("givendodatokdate") != null)
                this.textBox22.Text = diplom.Attribute("givendodatokdate").Value;
            if ( diplom.Attribute("issue_date") != null )
                this.textBox23.Text = diplom.Attribute("issue_date").Value;
            this.textBox5.Text = diplom.Attribute("numbdodatok") != null ? diplom.Attribute("numbdodatok").Value : "";
            //prev_institution
            XElement prev_institution = student.Descendants("prev_institution").FirstOrDefault();
            this.richTextBox27.Text = (prev_institution != null && prev_institution.Attribute("uk") != null) ? prev_institution.Attribute("uk").Value : "";
            this.richTextBox26.Text = (prev_institution != null && prev_institution.Attribute("en") != null) ? prev_institution.Attribute("en").Value : "";

            //prev_document / Обнаружено! На практике prev_document бывает пустым
            XElement prev_document = student.Descendants("prev_document").FirstOrDefault();
            if (prev_document.Attribute("ID") == null)
            {
                prev_document.Add(new XAttribute("ID", ""));
            }
            this.textBox16.Text = prev_document.Attribute("ID").Value;
            ComboboxIDSet(prev_document.Attribute("ID").Value);
            

            if (prev_document.Attribute("seria") == null)
            {
                prev_document.Add(new XAttribute("seria", ""));
            }
            this.textBox17.Text = prev_document.Attribute("seria").Value;

            if (prev_document.Attribute("number") == null)
            {
                prev_document.Add(new XAttribute("number", ""));
            }
            this.textBox18.Text = prev_document.Attribute("number").Value;

            if (prev_document.Attribute("prevqualificationuk") == null)
            {
                prev_document.Add(new XAttribute("prevqualificationuk", ""));
            }
            this.richTextBox23.Text = prev_document.Attribute("prevqualificationuk").Value;

            if (prev_document.Attribute("prevqualificationen") == null)
            {
                prev_document.Add(new XAttribute("prevqualificationen", ""));
            }
            this.richTextBox22.Text = prev_document.Attribute("prevqualificationen").Value;
            
            //Изменения нострификации. В одном поле аттрибута
            if (prev_document.Attribute("nostr") == null)
            {
                prev_document.Add(new XAttribute("nostr", ""));
            }
            this.richTextBox37.Text = prev_document.Attribute("nostr").Value;
            /*
            if (prev_document.Attribute("nostruk") == null)
            {
                prev_document.Add(new XAttribute("nostruk", ""));
            }
            this.richTextBox37.Text = prev_document.Attribute("nostruk").Value;

            if (prev_document.Attribute("nostren") == null)
            {
                prev_document.Add(new XAttribute("nostren", ""));
            }
            this.richTextBox36.Text = prev_document.Attribute("nostren").Value;
             */

            //prev_speciality
            XElement prev_speciality = student.Descendants("prev_speciality").FirstOrDefault();
            this.richTextBox25.Text = prev_speciality.Attribute("uk").Value;
            this.richTextBox24.Text = prev_speciality.Attribute("en").Value;
            //issue // Для каждого студента свой документ о предыдущем образовании!  Уточнить 
            XElement issue = student.Descendants("issue").FirstOrDefault();
            if (issue != null)
            {
                this.richTextBox21.Text = issue.Attribute("uk").Value;
                this.richTextBox20.Text = issue.Attribute("en").Value;
            }
            
            //certifiactioninfo
            XElement certifiactioninfo = student.Descendants("certifiactioninfo").FirstOrDefault();
            this.richTextBox19.Text = certifiactioninfo.Attribute("uk").Value;
            this.richTextBox18.Text = certifiactioninfo.Attribute("en").Value;
            //educationdates
            XElement educationdates = student.Descendants("educationdates").FirstOrDefault();
            this.textBox19.Text = educationdates.Attribute("receipt_date").Value;
            this.textBox20.Text = educationdates.Attribute("graduated").Value;
            // END подгрузка данных в поля

            //возвращаем прежний удобный формат
           /* this.textBox20.CustomFormat = "dd.MM.yyyy";
            this.textBox21.CustomFormat = "dd.MM.yyyy";
            this.textBox22.CustomFormat = "dd.MM.yyyy";
            this.textBox23.CustomFormat = "dd.MM.yyyy";*/

            // подгрузка в таблицу
            this.dataGridView1.Rows.Clear();
            foreach (var discipline in student.Descendants("discipline"))
            {
                string nameuk = discipline.Attribute("nameuk").Value;
                string nameen = discipline.Attribute("nameen").Value;
                string credits = discipline.Attribute("credits").Value;
                string hours = discipline.Attribute("hours").Value;
                string mark = discipline.Attribute("mark").Value;
                string ects_mark = discipline.Attribute("ects_mark").Value;
                string national_mark = National_markValidation( discipline.Attribute("national_mark").Value );
                string type = discipline.Attribute("type").Value;
                if (Regex.IsMatch(type, @"^([0-9])"))
                {
                    type = this.ConvertSubjectIndexToName(type);
                }
                else
                {
                    type = this.ConvertSubjectIndexToName(this.ConvertSubjectNameToIndex(type)); // чтобы небыло исключений из-за несовпадения с данными столбца
                }

                this.dataGridView1.Rows.Add(nameuk, nameen, credits, hours, mark, national_mark, ects_mark, type);

            }
            //total
            XElement total = student.Descendants("total").FirstOrDefault();
            try
            {
                if (total != null)
                {
                    this.dataGridView2.Rows[0].Cells["totalCredits"].Value = total.Attribute("credits").Value;
                    this.dataGridView2.Rows[0].Cells["totalHours"].Value = total.Attribute("hours").Value;
                    this.dataGridView2.Rows[0].Cells["middleMark"].Value = total.Attribute("mark").Value;
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Ошибка при чтении общих данных оценок!\n Программа продолжит работу");
            }

        }

        private void addStudentToMainObj()
        {
            //подгоняем под формат xml
            /*this.textBox20.CustomFormat = "dd/MM/yyyy";
            this.textBox21.CustomFormat = "dd/MM/yyyy";
            this.textBox22.CustomFormat = "dd/MM/yyyy";
            this.textBox23.CustomFormat = "dd/MM/yyyy";*/

            // Данные о студенте
            XElement newStudent = new XElement("student",// START student
                                        new XAttribute("lastnameuk", textBox1.Text),
                                        new XAttribute("firstnameuk", textBox11.Text),
                                        new XAttribute("middlenameuk", textBox12.Text),
                                        new XAttribute("lastnameen", textBox13.Text),
                                        new XAttribute("firstnameen", textBox14.Text),
                                        new XAttribute("middlenameen", textBox15.Text),
                                        new XAttribute("sex", SexBox.Text == "Чоловiча" ? "1" : "0"),
                                        new XAttribute("birthday", textBox24.Text),
                                        new XAttribute("foreigner", this.checkBox1.Checked ? "1" : "0"),
                                      // END student
                                      new XElement("diplom",// START diplom
                                        new XAttribute("seria", seriaBox.Text),
                                        new XAttribute("number", numberBox.Text),
                                        new XAttribute("honour", this.checkBox2.Checked ? "1" : "0"),
                                        new XAttribute("IsDuplicate", this.checkBox3.Checked ? "True" : "False"),
                                        new XAttribute("repeatDodatok", this.checkBox4.Checked ? "1" : "0"),
                                        new XAttribute("givendate", textBox21.Text),
                                        new XAttribute("givendodatokdate", textBox22.Text),
                                        new XAttribute("type", this.checkBox2.Checked ? "1" : "0"),
                                        new XAttribute("numbdodatok", textBox5.Text)
                                      ),// END diplom
                                      new XElement("prev_institution",// START prev_institution
                                        new XAttribute("uk", richTextBox27.Text),
                                        new XAttribute("en", richTextBox26.Text)
                                      ),// END prev_institution
                                      new XElement("prev_document",// START prev_document
                                        //new XAttribute("ID", textBox16.Text),
                                        new XAttribute("ID", ComboBoxIDGet()),
                                        new XAttribute("seria", textBox17.Text),
                                        new XAttribute("number", textBox18.Text),
                                        new XAttribute("prevqualificationuk", richTextBox23.Text),
                                        new XAttribute("prevqualificationen", richTextBox22.Text),
                                        new XAttribute("nostr", richTextBox37.Text),
                                        //new XAttribute("nostren", richTextBox36.Text),
                                        new XAttribute("issue_date", textBox23.Text)//????????????????????
                                      ),// END prev_document
                                      new XElement("prev_speciality",// START prev_speciality
                                        new XAttribute("uk", richTextBox25.Text),
                                        new XAttribute("en", richTextBox24.Text)
                                        ),// END prev_speciality
                                      new XElement("issue",// START issue /// Для каждого студента свой документ о предыдущем образовании!  Уточнить 
                                        new XAttribute("uk", richTextBox21.Text),
                                        new XAttribute("en", richTextBox20.Text)
                                        ),// END issue
                                      new XElement("certifiactioninfo",// START certifiactioninfo(BigText?)
                                        new XAttribute("uk", richTextBox19.Text),
                                        new XAttribute("en", richTextBox18.Text)
                                        ),// END certifiactioninfo
                                      new XElement("educationdates",// START educationdates
                                        new XAttribute("receipt_date", textBox19.Text), // Поле (date)
                                        new XAttribute("graduated", textBox20.Text)       // Поле (date)
                                        ),// END educationdates
                                      new XElement("duration",// START duration
                                        new XAttribute("uk", textBox19.Text + "-" + textBox20.Text), // Поле (date)
                                        new XAttribute("en", textBox19.Text + "-" + textBox20.Text)       // Поле (date)
                                        ),// END duration


                                      new XElement("marks")// END marks
                                    );
            //возвращаем прежний удобный формат
            /*this.textBox20.CustomFormat = "dd.MM.yyyy";
            this.textBox21.CustomFormat = "dd.MM.yyyy";
            this.textBox22.CustomFormat = "dd.MM.yyyy";
            this.textBox23.CustomFormat = "dd.MM.yyyy";*/

            XElement documentContainer = this.document.Element("documents")
                                                    .Element("document");
            //если есть другие то предметы просто копируем
            IEnumerable<XElement> studentsList = document.Descendants("student");
            if (studentsList != null && studentsList.Count() > 0)
            {
                newStudent.Element("marks").Remove();
                newStudent.Add(studentsList.First().Element("marks"));///test
                documentContainer.Add(newStudent);
                return;
            }

            
            documentContainer.Add(newStudent);
            XElement marks = newStudent.Element("marks");

            foreach (DataGridViewRow row in this.dataGridView1.Rows)
            {
                try
                {
                    if (!row.IsNewRow)
                        marks.Add(new XElement("discipline",
                                    new XAttribute("nameuk", row.Cells["nameuk"].Value.ToString()),
                                    new XAttribute("nameen", row.Cells["nameen"].Value.ToString()),
                                    new XAttribute("credits", row.Cells["credits"].Value.ToString()),
                                    new XAttribute("hours", row.Cells["hours"].Value.ToString()),
                                    new XAttribute("mark", row.Cells["mark"].Value.ToString()),
                                    new XAttribute("national_mark", row.Cells["national_mark"].Value.ToString()),
                                    new XAttribute("ects_mark", row.Cells["ects_mark"].Value.ToString()),
                                    new XAttribute("type", this.ConvertSubjectNameToIndex(row.Cells["type"].Value.ToString()))
                        ));
                }
                catch (Exception)
                {
                    MessageBox.Show("Все поля таблицы должны быть введены");
                    return;
                }
            }

            double sumCredits = 0;
            double sumHours = 0;
            double sumMarkSubj = 0;
            int markCount = 0;
            int attestCount = 0;
            double sumMarkAttest = 0;
            foreach (DataGridViewRow row in this.dataGridView1.Rows)
            {
                if (!row.IsNewRow)
                {
                    sumCredits += Convert.ToDouble(row.Cells["credits"].Value.ToString().Replace('.', ','));
                    sumHours += Convert.ToDouble(row.Cells["hours"].Value.ToString().Replace('.', ','));
                    if (row.Cells["type"].Value.ToString().Equals("Аттестация"))
                    {
                        sumMarkAttest += Convert.ToDouble(row.Cells["mark"].Value.ToString().Replace('.', ','));
                        attestCount++;
                    }
                    else
                    {
                        sumMarkSubj += Convert.ToDouble(row.Cells["mark"].Value.ToString().Replace('.', ','));
                        markCount++;
                    }

                }

            }

            try
            {
                int isatt = 0;
                if (markCount != 0) isatt++;
                if (attestCount != 0) isatt++;
                markCount = markCount == 0 ? 1 : markCount;
                attestCount = attestCount == 0 ? 1 : attestCount;
                if (isatt > 0)
                    marks.Add(new XElement("total",
                                new XAttribute("credits", sumCredits.ToString(".#").Replace(',', '.')),
                                new XAttribute("hours", sumHours.ToString(".#").Replace(',', '.')),
                                new XAttribute("mark", ((sumMarkSubj / markCount + sumMarkAttest / attestCount) / isatt).ToString(".#").Replace(',', '.'))
                        ));
            }
            catch (Exception)
            {
                MessageBox.Show("Все поля таблицы итогов должны быть заполнены");
                return;
            }

        }

        private void listView2_DoubleClick(object sender, EventArgs e)
        {
            this.listView2LoadSelectedItem();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            var items = listView2.SelectedItems;
            if (items.Count == 0) return;
            this.listView2LoadSelectedItem();
            
        }

        private void listView2LoadSelectedItem()
        {
            ComboBoxIDSetDefaultValue();
            var items = listView2.SelectedItems;
            if (items.Count == 0) return;
            ListViewItem clickedItem = null;
            foreach (var item in items)
            {
                clickedItem = item as ListViewItem;
                break;
            }
            int i = 0;
            foreach (var student in students)
            {
                if (i == clickedItem.Index)
                {
                    LoadToFields(document); //при загрузке для работы с каждым студентом подгружаем и все общие данные
                    loadStudentToFieldsSingle(student as XElement);
                    this.Student = student;
                    break;
                }

                i++;
            }
            updateBtnEnable();
        }



        //обновить список студентов
        private void listView2_updateItems()
        {
            this.students = document.Descendants("student");
            listView2.Items.Clear();
            foreach(var elem in students)
            {
                listView2.Items.Add(
                    elem.Attribute("lastnameuk").Value + "  " + 
                    elem.Attribute("firstnameuk").Value + " " + 
                    elem.Attribute("middlenameuk").Value);
            }
        }
        //добавить студента
        private void button9_Click(object sender, EventArgs e)
        {
            if (document == null) this.createXDocFromFields();
            addStudentToMainObj();
            listView2_updateItems();
            ClearAllFields();
        }
        //удалить студента
        private void button10_Click(object sender, EventArgs e)
        {
            var items = listView2.SelectedItems;
            if (items.Count == 0)
            {
                MessageBox.Show("Вы не выбрали запись которую нужно удалить!");
                return;
            }

            ListViewItem clickedItem = null;
            foreach (var item in items)
            {
                clickedItem = item as ListViewItem;
                break;
            }
            int i = 0;
            foreach (var student in students)
            {
                if (i == clickedItem.Index)
                {
                    students.ElementAt(i).Remove();
                    this.listView2_updateItems();
                    break;
                }

                i++;
            }


        }

        //Обновить данные студента в XML
        private void updateStud_Click(object sender, EventArgs e)
        {
            //реализация
            updateXDoc();
            updateStudent();
            listView2_updateItems();
            ClearAllFields();
            updateBtnDisable();
            this.tabControl1.SelectedTab = this.tabControl1.TabPages[0];
        }

        private void updateStudent()
        {
            if (this.Student == null) return;

            this.Student.Attribute("lastnameuk").SetValue(textBox1.Text);
            this.Student.Attribute("firstnameuk").SetValue(textBox11.Text);
            this.Student.Attribute("middlenameuk").SetValue(textBox12.Text);
            this.Student.Attribute("lastnameen").SetValue(textBox13.Text);
            this.Student.Attribute("firstnameen").SetValue(textBox14.Text);
            this.Student.Attribute("middlenameen").SetValue(textBox15.Text);
            this.Student.Attribute("sex").SetValue(SexBox.Text == "Чоловiча" ? "1" : "0");
            this.Student.Attribute("birthday").SetValue(textBox24.Text);
            this.Student.Attribute("foreigner").SetValue(this.checkBox1.Checked ? "1" : "0");
            
            this.Student.Element("diplom").Attribute("seria").SetValue(seriaBox.Text); 
            this.Student.Element("diplom").Attribute("number").SetValue(numberBox.Text);
            this.Student.Element("diplom").Attribute("honour").SetValue(this.checkBox2.Checked ? "1" : "0");
            //IsDuplicate
            if (this.Student.Element("diplom").Attribute("IsDuplicate") == null)
                this.Student.Element("diplom").Add(new XAttribute("IsDuplicate", "False")); 
            this.Student.Element("diplom").Attribute("IsDuplicate").SetValue(this.checkBox3.Checked ? "True" : "False");
            //repeatDodatok
            if (this.Student.Element("diplom").Attribute("repeatDodatok") == null)
                this.Student.Element("diplom").Add(new XAttribute("repeatDodatok", "0")); 
            this.Student.Element("diplom").Attribute("repeatDodatok").SetValue(this.checkBox4.Checked ? "1" : "0");

            this.Student.Element("diplom").Attribute("givendate").SetValue(textBox21.Text);
            //givendodatokdate
            if (this.Student.Element("diplom").Attribute("givendodatokdate") == null)
                this.Student.Element("diplom").Add(new XAttribute("givendodatokdate", "0"));
            this.Student.Element("diplom").Attribute("givendodatokdate").SetValue(textBox22.Text);
            //issue_date
            if (this.Student.Element("diplom").Attribute("issue_date") == null)
                this.Student.Element("diplom").Add(new XAttribute("issue_date", "0"));
            this.Student.Element("diplom").Attribute("issue_date").SetValue(textBox23.Text);

            //this.Student.Element("diplom").Attribute("type").SetValue(this.checkBox2.Checked ? "1" : "0");
            //numbdodatok
            if(this.Student.Element("diplom").Attribute("numbdodatok") == null)
                this.Student.Element("diplom").Add(new XAttribute("numbdodatok", ""));
            this.Student.Element("diplom").Attribute("numbdodatok").SetValue(textBox5.Text);

            if (this.Student.Element("prev_institution") != null)
            {
                if (this.Student.Element("prev_institution").Attribute("uk") != null)
                    this.Student.Element("prev_institution").Attribute("uk").SetValue(richTextBox27.Text);
                if (this.Student.Element("prev_institution").Attribute("en") != null)
                    this.Student.Element("prev_institution").Attribute("en").SetValue(richTextBox26.Text);
            }
            else
            {
                this.Student.Add(new XElement("prev_institution",
                    new XAttribute("uk", richTextBox27.Text),
                    new XAttribute("en", richTextBox26.Text)
                    ));
            }


            //this.Student.Element("prev_document").Attribute("ID").SetValue(textBox16.Text);
            this.Student.Element("prev_document").Attribute("ID").SetValue(ComboBoxIDGet());
            this.Student.Element("prev_document").Attribute("seria").SetValue(textBox17.Text);
            this.Student.Element("prev_document").Attribute("number").SetValue(textBox18.Text);
            this.Student.Element("prev_document").Attribute("prevqualificationuk").SetValue(richTextBox23.Text);
            this.Student.Element("prev_document").Attribute("prevqualificationen").SetValue(richTextBox22.Text);
            this.Student.Element("prev_document").Attribute("nostr").SetValue(richTextBox37.Text);
            //this.Student.Element("prev_document").Attribute("nostren").SetValue(richTextBox36.Text);

            this.Student.Element("prev_speciality").Attribute("uk").SetValue(richTextBox25.Text);
            this.Student.Element("prev_speciality").Attribute("en").SetValue(richTextBox24.Text);

            if (this.Student.Element("issue") == null)
                this.Student.Add(new XElement("issue",
                    new XAttribute("uk", ""),
                    new XAttribute("en", "")) );
            this.Student.Element("issue").Attribute("uk").SetValue(richTextBox21.Text);
            this.Student.Element("issue").Attribute("en").SetValue(richTextBox20.Text);


            this.Student.Element("certifiactioninfo").Attribute("uk").SetValue(richTextBox19.Text);
            this.Student.Element("certifiactioninfo").Attribute("en").SetValue(richTextBox18.Text);

            this.Student.Element("educationdates").Attribute("receipt_date").SetValue(textBox19.Text);
            this.Student.Element("educationdates").Attribute("graduated").SetValue(textBox20.Text);


            this.Student.Element("duration").Attribute("uk").SetValue(textBox19.Text + "-" + textBox20.Text);
            this.Student.Element("duration").Attribute("en").SetValue(textBox19.Text + "-" + textBox20.Text);
            
            //marks
            XElement marksFromTable = new XElement("marks");
            
            foreach (DataGridViewRow row in this.dataGridView1.Rows)
            {
                try
                {
                    if (!row.IsNewRow)
                        marksFromTable.Add(new XElement("discipline",
                                    new XAttribute("nameuk", row.Cells["nameuk"].Value.ToString()),
                                    new XAttribute("nameen", row.Cells["nameen"].Value.ToString()),
                                    new XAttribute("credits", row.Cells["credits"].Value.ToString()),
                                    new XAttribute("hours", row.Cells["hours"].Value.ToString()),
                                    new XAttribute("mark", row.Cells["mark"].Value.ToString()),
                                    new XAttribute("national_mark", row.Cells["national_mark"].Value.ToString()),
                                    new XAttribute("ects_mark", row.Cells["ects_mark"].Value.ToString()),
                                    new XAttribute("type", this.ConvertSubjectNameToIndex(row.Cells["type"].Value.ToString()))

                        ));
                }
                catch (Exception)
                {
                    MessageBox.Show("Все поля таблицы должны быть введены");
                    return;
                }
            }

            double sumCredits = 0;
            double sumHours = 0;
            double sumMarkSubj = 0;
            int markCount = 0;
            int attestCount = 0;
            double sumMarkAttest = 0;
            foreach (DataGridViewRow row in this.dataGridView1.Rows)
            {
                if (!row.IsNewRow)
                {
                    sumCredits += Convert.ToDouble(row.Cells["credits"].Value.ToString().Replace('.', ','));
                    sumHours += Convert.ToDouble(row.Cells["hours"].Value.ToString().Replace('.', ','));
                    if (row.Cells["type"].Value.ToString().Equals("Аттестация"))
                    {
                        sumMarkAttest += Convert.ToDouble(row.Cells["mark"].Value.ToString().Replace('.', ','));
                        attestCount++;
                    }
                    else
                    {
                        sumMarkSubj += Convert.ToDouble(row.Cells["mark"].Value.ToString().Replace('.', ','));
                        markCount++;
                    }

                }

            }

            try
            {
                int isatt = 0;
                if (markCount != 0) isatt++;
                if (attestCount != 0) isatt++;
                markCount = markCount == 0 ? 1 : markCount;
                attestCount = attestCount == 0 ? 1 : attestCount;
                if (isatt > 0)
                    marksFromTable.Add(new XElement("total",
                                new XAttribute("credits", sumCredits.ToString(".#").Replace(',', '.')),
                                new XAttribute("hours", sumHours.ToString(".#").Replace(',', '.')),
                                new XAttribute("mark", ((sumMarkSubj / markCount + sumMarkAttest / attestCount) / isatt).ToString(".#").Replace(',', '.'))
                        ));
            }
            catch (Exception)
            {
                MessageBox.Show("Все поля таблицы итогов должны быть заполнены");
                return;
            }
            
            this.Student.Element("marks").Remove();
            this.Student.Add(marksFromTable);

            this.Student = null;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            

        }

        //конвертация названия предмета в индекс (в исходных xml данные идут числовим индексом)
        private string ConvertSubjectNameToIndex(string subjectName = "")
        {
            if (subjectName == "") return "1";
            string indexStr = "1";
            switch (subjectName)
            {
                case "Предмет": indexStr = "1"; break;
                case "Практика": indexStr = "2"; break;
                case "Аттестация": indexStr = "3"; break;
                case "Курсовая": indexStr = "4"; break;
                default: indexStr = "1"; break;
            }
            return indexStr;

        }

        //конвертация индекса предмета в название (в исходных xml данные идут числовим индексом)
        private string ConvertSubjectIndexToName(string subjectIndex = "")
        {
            if (subjectIndex == "") return "Предмет";
            string nameStr = "Предмет";
            switch (subjectIndex)
            {
                case "1": nameStr = "Предмет"; break;
                case "2": nameStr = "Практика"; break;
                case "3": nameStr = "Аттестация"; break;
                case "4": nameStr = "Курсовая"; break;
                default: nameStr = "Предмет"; break;
            }
            return nameStr;
        }

        //проверка и корректировка совместимости получаемых данных national_mark из xml с допустимыми значениями в выпадающем списке national_mark из таблицы Дисциплин
        private string National_markValidation(string value)
        {
            //Выбираем колекцию с данными национальных оценок, зашитых в свойствах элемента списка в таблице
            DataGridViewComboBoxColumn national_markColumn = this.dataGridView1.Columns["national_mark"] as DataGridViewComboBoxColumn;
            var columnItems = national_markColumn.Items;

            //перебираем и сравниваем на совместимость значения
            string result = "";
            bool isEqual = false;
            foreach (var item in columnItems)
            {
                isEqual = value.Equals(item.ToString());
                if(isEqual)
                {
                    result = value;
                    break;
                }
                else //Если совпадение не обнаружится у нас останется последний эл-т из колекции значений списка
                {
                    result = item.ToString();
                }
            }
            return result;
        }

        //Подтвердить изменение. Включить кнопку
        private bool updateBtnEnable()
        {
            this.updateStud.Enabled = true;
            this.updateStud.BackColor = System.Drawing.Color.LightYellow;
            return this.updateStud.Enabled;
        }
        //Подтвердить изменение. Отключить кнопку
        private bool updateBtnDisable()
        {
            this.updateStud.Enabled = false;
            this.updateStud.BackColor = System.Drawing.Color.WhiteSmoke;
            return this.updateStud.Enabled;
        }
        //Установить значение в ComboboxID - Идентификатор
        private void ComboboxIDSet(string a)
        {
            switch(a)
            {
                case  "2":
                    this.comboBoxID.SelectedIndex = 1;
                    break;
                case  "5":
                case "10":
                    this.comboBoxID.SelectedIndex = 2;
                    break;
                case  "6":
                case "11":
                    this.comboBoxID.SelectedIndex = 3;
                    break;
                case "12":
                    this.comboBoxID.SelectedIndex = 4;
                    break;
                case "13":
                    this.comboBoxID.SelectedIndex = 5;
                    break;
                default  :
                    this.comboBoxID.SelectedIndex = 0;
                    break;
            }

        }
        //Получить значение из ComboboxID - Идентификатор
        private string ComboBoxIDGet()
        {
            string ID = "";
            switch(this.comboBoxID.SelectedIndex)
            {
                case 1:
                    ID = "2";
                    break;
                case 2:
                    ID = "10";
                    break;
                case 3:
                    ID = "11";
                    break;
                case 4:
                    ID = "12";
                    break;
                case 5:
                    ID = "13";
                    break;
                default:
                    ID = "";
                    break;
            }

            return ID;
        }
        //Обнулить значение ComboboxID. Первый элемент = ""
        private void ComboBoxIDSetDefaultValue()
        {
            this.comboBoxID.SelectedIndex = 0;
        }

        private void comboBox37_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

    }
}
