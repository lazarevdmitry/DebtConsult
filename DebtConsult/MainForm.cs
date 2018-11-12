/*
 * Создано в SharpDevelop.
 * Пользователь: Дмитрий
 * Дата: 23.01.2015
 * Время: 22:03
 * 
 * Для изменения этого шаблона используйте меню "Инструменты | Параметры | Кодирование | Стандартные заголовки".
 */
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Word=Microsoft.Office.Interop.Word;


namespace DebtConsult
{
	/// <summary>
	/// Description of MainForm.
	/// </summary>
	public partial class MainForm : Form
	{
		// Переменные формы
		public String tIFamilia, tIName, tIOtchestvo;
		public String tRFamilia, tRName, tROtchestvo;
		public String tDFamilia, tDName, tDOtchestvo;
		public String tPol, tDataRozhd;
		public String tSeria, tNomer, tKemVydan, tDataVydachi;
		public String tAdresReg, tAdresProzh;
		public String tTelefon;
		public String tOfertaNomer, tOfertaData, tOfertaSumma, tOfertaRassrochka, tOfertaVznos;		
		//
		Word._Application app;
		Word._Document doc1;
		Word._Document doc2;
		Word._Document doc3;
		Word._Document doc4;
		Word._Document doc5;
//		Word._Document doc6;
//		Word._Document doc7;
//		Word._Document doc8;
//		Word._Document doc9;
//		Word._Document doc10;
		String dir;
		// 
		public const String NOVALUE = "Не введено значение";
		
		public String[] months = {"января","февраля","марта","апреля","мая","июня",
			"июля","августа","сентября","октября","ноября","декабря"};
		// Строки, которые будут хранить названия шаблонов документов
		public const String strDoc1="yurspravka.dotx";
		public const String strDoc2="dogovorip.dotx";
		public const String strDoc3="schetip.dotx";
		public const String strDoc4="dogovorvechsel.dotx";
		public const String strDoc5="garant.dotx";
		public const String strDoc6="";
		public const String strDoc7="";
		public const String strDoc8="";
		public const String strDoc9="";
		public const String strDoc10="";
		public String tmpDir=System.IO.Directory.GetCurrentDirectory()+"\\";
		//public const String tmpDir="C:\\Debt\\";
		
		public MainForm()
		{
			//
			// The InitializeComponent() call is required for Windows Forms designer support.
			//
			InitializeComponent();
		
			//
			// TODO: Add constructor code after the InitializeComponent() call.
			//
			//test();
			this.Text+=" ["+tmpDir+"]";
		}
		

		// 		Основной алгоритм действий программы
		void mainapp(){
			// После того, как была нажата кнопка НАЧАТЬ
			
			// Проверяем наличие папки для вывода
			if (!checkDirectoryExists()) return;
			// Создаем объект приложения
			makeApplication();
			// Устанавливаем базовую папку
			
//			Все документы:
//				- юридическая справка должника
//				- договор агента ИП
//				- счет на ИП
//				- договор покупки векселя
			// Проверяем состояние флажков. Если флажок отмечен, то
			// запускаем метод работы с отмеченным документом
			if (checkBox1.Checked) makeDoc1();
			if (checkBox2.Checked) makeDoc2();
			if (checkBox3.Checked) makeDoc3();
			if (checkBox4.Checked) makeDoc4();
			if (checkBox7.Checked) makeDoc5();
//			if (checkBox6.Checked) makeDoc6();
//			if (checkBox7.Checked) makeDoc7();
//			if (checkBox8.Checked) makeDoc8();
//			if (checkBox9.Checked) makeDoc9();
//			if (checkBox10.Checked) makeDoc10();
			
			// Закрываем объект приложения
			deleteApplication();
			// Предалагаем Пользователю открыть в Проводнике папку сохранения
			//doYouWantToViewInExplorer();
		}		 
		bool checkDirectoryExists(){
		// Проверяем, введен ли путь к папке сохранения
		// Преамбула:
		// 1) путь введен - не введен
		// 2) путь существует - не существует
		// 3) путь можно создать - нельзя создать
		//
			dir= textBox19.Text;
			// Путь введен или не введен
			if (dir==null||dir.Equals("")){
				MessageBox.Show("Не введен путь для сохранения файлов", "Ошибка");
				return false;
			}
			// Путь существует или не существует
			if (System.IO.Directory.Exists(dir)){
				// Ничего не делаем			
			} else {
			// Создаем папку сохранения файлов 
				try {
					System.IO.Directory.CreateDirectory(dir);
				} catch (Exception dnfe){
				// Если путь папки недопустим
				MessageBox.Show("Невозможно создать папку для сохранения"+dnfe.Message, "Ошибка");
				return false;
				}			
			}
			return true;
		}
		void makeApplication(){
			// Создаем объект приложения
			app = new Word.ApplicationClass();
			app.Visible=false;
		}
		void deleteApplication(){
			app.Quit();
		}
		void makeDoc1(){
			// Создаем документ "Юридическая справка должника"
			// Необходимые поля: ФИО в дательном падеже, текущая дата			
			int i=0;
			// Проверяем наличие введенных в текстовые поля данные 
			if (textBox9.Text.Equals("")) {noValue(textBox9);i++;}
			if (textBox8.Text.Equals("")) {noValue(textBox8);i++;}
			if (textBox7.Text.Equals("")) {noValue(textBox7);i++;}
			if (i!=0){message(NOVALUE);return;}
			try {
			// Если в наличии все необходимые данные, то
			// создаем объект документа
			doc1= new Word.DocumentClass();
			// Добавляем документ в объект приложения
			// template - шаблон "yurspravka.dotx"
			// newTemplate - должен сводиться к false
			// docType - тип документа (нужен *.docx)
			// visible = false
			String temp=tmpDir+strDoc1;
			object template = @temp;
			object newTemplate = Type.Missing;
			object docType = Word.WdNewDocumentType.wdNewBlankDocument;
			object visible = true;			
			doc1=app.Documents.Add(ref template, ref newTemplate, ref docType, ref visible);						
			// Вводим данные для закладок (bookmarks)			
			Word.Range rn=doc1.Bookmarks["yurspravka1"].Range;
			rn.Text=textBox9.Text+" "+textBox8.Text+" "+textBox7.Text;					
			DateTime currDate = DateTime.Now;
			doc1.Bookmarks["yurspravka2"].Range.Text=currDate.Day.ToString();
			doc1.Bookmarks["yurspravka3"].Range.Text=months[currDate.Month-1];			
			doc1.Bookmarks["yurspravka4"].Range.Text=currDate.Year.ToString();
			  String fileSave = dir+"\\Юридическая справка должника "+textBox9.Text+" "+textBox8.Text+" "+textBox7.Text;
			  object fn = @fileSave;
			  object ff = Word.WdSaveFormat.wdFormatDocument;
			  object lc = false;
			  object psswd = "";
			  object f3 = false;
			  object f4 = "";
			  object f5 = false;
			  object f6 = false;
			  object f7 = false;
			  object f8 = false;
			  object f9 = Type.Missing;
			  object f10 = Type.Missing;
			  object f11 = Type.Missing;
			  object f12 = Type.Missing;
			  object f13 = Type.Missing;
			  object f14 = Type.Missing;			  
			doc1.SaveAs(ref fn, ref ff,ref lc, ref psswd, ref f3, ref f4, ref f5, ref f6, ref f7, ref f8,
			              ref f9, ref f10,ref f11, ref f12, ref f13, ref f14);			
			  
			doc1.Close();
			temp=null;
			} catch (System.IO.IOException ioe){
				message(ioe.Message);
				return;
			}
		}
		
		void makeDoc2(){
			// Создаем документ "Договор агента ИП"
			// Необходимые поля: № договора, дата подписания, ФИО, сумма цифрами, сумма прописью,
			// снова дата подписания
			int i=0;
			// Проверяем наличие введенных в текстовые поля данные 
			if (textBox16.Text.Equals("")) {noValue(textBox16);i++;}	// № договора
			if (textBox1.Text.Equals("")) {noValue(textBox1);i++;}	// Фамилия
			if (textBox2.Text.Equals("")) {noValue(textBox2);i++;}	// Имя
			if (textBox3.Text.Equals("")) {noValue(textBox3);i++;}	// Отчество
			if (textBox18.Text.Equals("")) {noValue(textBox18);i++;}	// Сумма цифрами
			if (i!=0){message(NOVALUE);return;}
			// Если в наличии все необходимые данные, то
			// создаем объект документа
			try {
			doc2= new Word.DocumentClass();		
			String temp=tmpDir+strDoc2;
			object template = @temp;
			object newTemplate = Type.Missing;
			object docType = Word.WdNewDocumentType.wdNewBlankDocument;
			object visible = true;
			doc2=app.Documents.Add(ref template, ref newTemplate, ref docType, ref visible);			
			// Вводим данные для закладок (bookmarks)						
			doc2.Bookmarks["dogovorip1"].Range.Text=textBox16.Text;	// № договора
			doc2.Bookmarks["dogovorip2"].Range.Text=dateTimePicker3.Value.Day.ToString();	// день подписания
			doc2.Bookmarks["dogovorip3"].Range.Text=months[dateTimePicker3.Value.Month-1];	// месяц подписания
			//doc2.Bookmarks["dogovorip4"].Range.Text=dateTimePicker3.Value.Year.ToString();	// год подписания
			doc2.Bookmarks["dogovorip4"].Range.Text=textBox1.Text+" "+textBox2.Text+" "+textBox3.Text;	// ФИО клиента
			doc2.Bookmarks["dogovorip5"].Range.Text=placeWhiteSpace(textBox18.Text);	// Сумма цифрами
			doc2.Bookmarks["dogovorip6"].Range.Text=cipherToChars(textBox18.Text);	// Сумма прописью
			doc2.Bookmarks["dogovorip7"].Range.Text=dateTimePicker3.Value.Day.ToString();	// день подписания
			doc2.Bookmarks["dogovorip8"].Range.Text=months[dateTimePicker3.Value.Month-1];	// месяц подписания
			//doc2.Bookmarks[""].Range.Text=;
			  String fileSave = dir+"\\Договор с ИП Придача ОА "+textBox1.Text+" "+textBox2.Text+" "+textBox3.Text;
			  object fn = @fileSave;
			  object ff = Word.WdSaveFormat.wdFormatDocument;
			  object lc = false;
			  object psswd = "";
			  object f3 = false;
			  object f4 = "";
			  object f5 = false;
			  object f6 = false;
			  object f7 = false;
			  object f8 = false;
			  object f9 = Type.Missing;
			  object f10 = Type.Missing;
			  object f11 = Type.Missing;
			  object f12 = Type.Missing;
			  object f13 = Type.Missing;
			  object f14 = Type.Missing;			  
			doc2.SaveAs(ref fn, ref ff,ref lc, ref psswd, ref f3, ref f4, ref f5, ref f6, ref f7, ref f8,
			              ref f9, ref f10,ref f11, ref f12, ref f13, ref f14);			
			doc2.Close();
			temp=null;
			} catch (System.IO.IOException ioe){
				message(ioe.Message);
			}
		}
		
		void makeDoc3(){
			// Создаем документ "Счет ИП"
			// Необходимые поля: № договора, дата в формате "дд.мм.гггг г.", ФИО клиента, адрес клиента, сумма цифрами
			int i=0;
			// Проверяем наличие введенных в текстовые поля данные 
			if (textBox16.Text.Equals("")) {noValue(textBox16);i++;}	// № договора
			if (textBox1.Text.Equals("")) {noValue(textBox1);i++;}	// Фамилия
			if (textBox2.Text.Equals("")) {noValue(textBox2);i++;}	// Имя
			if (textBox3.Text.Equals("")) {noValue(textBox3);i++;}	// Отчество
			if (textBox13.Text.Equals("")) {noValue(textBox13);i++;}	// Адрес клиента
			if (textBox18.Text.Equals("")) {noValue(textBox18);i++;}	// Сумма цифрами
			if (i!=0){message(NOVALUE);return;}
			// Если в наличии все необходимые данные, то
			// создаем объект документа
			try {
			doc3= new Word.DocumentClass();		
			String temp=tmpDir+strDoc3;
			object template = @temp;
			object newTemplate = Type.Missing;
			object docType = Word.WdNewDocumentType.wdNewBlankDocument;
			object visible = true;
			doc3=app.Documents.Add(ref template, ref newTemplate, ref docType, ref visible);			
			// Вводим данные для закладок (bookmarks)						
			doc3.Bookmarks["schetip1"].Range.Text=textBox1.Text+" "+textBox2.Text+" "+textBox3.Text;	// ФИО клиента
			doc3.Bookmarks["schetip2"].Range.Text=textBox13.Text;	// Адрес регистрации клиента
			doc3.Bookmarks["schetip3"].Range.Text=textBox16.Text;	// номер договора
			doc3.Bookmarks["schetip4"].Range.Text=dateTimePicker3.Value.Date.ToLongDateString();	// дата подписания договора в формате "дд.мм.гггг г."
			doc3.Bookmarks["schetip5"].Range.Text=textBox18.Text+"-00";	// сумма цифрами
			doc3.Bookmarks["schetip6"].Range.Text=textBox18.Text+"-00";	// сумма цифрами
			doc3.Bookmarks["schetip7"].Range.Text=textBox18.Text+"-00";	// сумма цифрами
			doc3.Bookmarks["schetip8"].Range.Text=textBox18.Text+"-00";	// сумма цифрами
			doc3.Bookmarks["schetip9"].Range.Text=textBox18.Text+"-00";	// сумма цифрами			
			
			 String fileSave = dir+"\\Счет на ИП Придача ОА "+textBox1.Text+" "+textBox2.Text+" "+textBox3.Text;
			  object fn = @fileSave;
			  object ff = Word.WdSaveFormat.wdFormatDocument;
			  object lc = false;
			  object psswd = "";
			  object f3 = false;
			  object f4 = "";
			  object f5 = false;
			  object f6 = false;
			  object f7 = false;
			  object f8 = false;
			  object f9 = Type.Missing;
			  object f10 = Type.Missing;
			  object f11 = Type.Missing;
			  object f12 = Type.Missing;
			  object f13 = Type.Missing;
			  object f14 = Type.Missing;			  
			doc3.SaveAs(ref fn, ref ff,ref lc, ref psswd, ref f3, ref f4, ref f5, ref f6, ref f7, ref f8,
			              ref f9, ref f10,ref f11, ref f12, ref f13, ref f14);			
			doc3.Close();
			temp=null;
			} catch (System.IO.IOException ioe){
				message(ioe.Message);
			}
		}
		
		void makeDoc4(){
			// Создаем документ "Договор покупки векселя"
			// Необходимые поля: номер договора, день, месяц, год, 
			// номер договора, дата подписания договора в формате "дд.мм.гггг г.",
			// ФИО в родительном падеже, адрес регистрации клиента, номер телефона,
			// день, месяц, год, 
			int i=0;
			// Проверяем наличие введенных в текстовые поля данные 
			if (textBox16.Text.Equals("")) {noValue(textBox16);i++;}	// № договора
			if (textBox6.Text.Equals("")) {noValue(textBox6);i++;}	// Фамилия в родительном падеже
			if (textBox5.Text.Equals("")) {noValue(textBox5);i++;}	// Имя в родительном падеже
			if (textBox4.Text.Equals("")) {noValue(textBox4);i++;}	// Отчество в родительном падеже
			if (textBox13.Text.Equals("")) {noValue(textBox13);i++;}	// Адрес регистрации клиента
			if (textBox15.Text.Equals("")) {noValue(textBox15);i++;}	// номер телефона
			if (i!=0){message(NOVALUE);return;}
			// Если в наличии все необходимые данные, то
			// создаем объект документа
			try {
			doc4= new Word.DocumentClass();		
			String temp=tmpDir+strDoc4;
			object template = @temp;
			object newTemplate = Type.Missing;
			object docType = Word.WdNewDocumentType.wdNewBlankDocument;
			object visible = true;
			doc4=app.Documents.Add(ref template, ref newTemplate, ref docType, ref visible);			
			// Вводим данные для закладок (bookmarks)						
			doc4.Bookmarks["dogovorvechsel1"].Range.Text=textBox16.Text;	// номер договора
			doc4.Bookmarks["dogovorvechsel2"].Range.Text=dateTimePicker3.Value.Day.ToString();	// день
			doc4.Bookmarks["dogovorvechsel3"].Range.Text=months[dateTimePicker3.Value.Month-1];	// месяц
			doc4.Bookmarks["dogovorvechsel4"].Range.Text=dateTimePicker3.Value.Year.ToString();	// год
			doc4.Bookmarks["dogovorvechsel5"].Range.Text=textBox16.Text;	// номер договора
			doc4.Bookmarks["dogovorvechsel6"].Range.Text=dateTimePicker3.Value.ToShortDateString();	// дата подписания договора
			doc4.Bookmarks["dogovorvechsel7"].Range.Text=textBox6.Text+" "+textBox5.Text+" "+textBox4.Text;	// ФИО в родительном падеже
			doc4.Bookmarks["dogovorvechsel8"].Range.Text=textBox13.Text;	// адрес регистрации клиента
			doc4.Bookmarks["dogovorvechsel9"].Range.Text=textBox15.Text;	// номер телефона
			doc4.Bookmarks["dogovorvechsel10"].Range.Text=dateTimePicker3.Value.Day.ToString();	// день
			doc4.Bookmarks["dogovorvechsel11"].Range.Text=months[dateTimePicker3.Value.Month-1];	// месяц
			doc4.Bookmarks["dogovorvechsel12"].Range.Text=dateTimePicker3.Value.Year.ToString();	// год
			
			 String fileSave = dir+"\\Договор покупки векселя "+textBox1.Text+" "+textBox2.Text+" "+textBox3.Text;
			  object fn = @fileSave;
			  object ff = Word.WdSaveFormat.wdFormatDocument;
			  object lc = false;
			  object psswd = "";
			  object f3 = false;
			  object f4 = "";
			  object f5 = false;
			  object f6 = false;
			  object f7 = false;
			  object f8 = false;
			  object f9 = Type.Missing;
			  object f10 = Type.Missing;
			  object f11 = Type.Missing;
			  object f12 = Type.Missing;
			  object f13 = Type.Missing;
			  object f14 = Type.Missing;			  
			doc4.SaveAs(ref fn, ref ff,ref lc, ref psswd, ref f3, ref f4, ref f5, ref f6, ref f7, ref f8,
			              ref f9, ref f10,ref f11, ref f12, ref f13, ref f14);			
			doc4.Close();
			temp=null;
			} catch (System.IO.IOException ioe){
				message(ioe.Message);
			}
		}
		
		void makeDoc5(){
			// Создаем документ "Гарантийное письмо"
			// Необходимые поля: день, месяц, год, ФИО в им. падеже, № договора
			// день заключения договора, месяц заключения договора, год заключения договора
			// сумма векселя числом
			int i=0;
			// Проверяем наличие введенных в текстовые поля данные 
			if (textBox1.Text.Equals("")) {noValue(textBox1);i++;}	//	Фамилия
			if (textBox2.Text.Equals("")) {noValue(textBox2);i++;}	// Имя
			if (textBox3.Text.Equals("")) {noValue(textBox3);i++;}	// Отчество
			if (textBox16.Text.Equals("")) {noValue(textBox16);i++;}	// номер договора
			if (textBox17.Text.Equals("")) {noValue(textBox17);i++;}	// сумма векселя числом

			if (i!=0){message(NOVALUE);return;}
			// Если в наличии все необходимые данные, то
			// создаем объект документа
			try {
			doc5= new Word.DocumentClass();
			String temp=tmpDir+strDoc5;
			object template = @temp;
			object newTemplate = Type.Missing;
			object docType = Word.WdNewDocumentType.wdNewBlankDocument;
			object visible = true;
			doc5=app.Documents.Add(ref template, ref newTemplate, ref docType, ref visible);
			// Вводим данные для закладок (bookmarks)						
			DateTime dt = DateTime.Now;
			doc5.Bookmarks["garant1"].Range.Text=dt.Day.ToString();	//
			doc5.Bookmarks["garant2"].Range.Text=months[dt.Month-1];	// 
			doc5.Bookmarks["garant3"].Range.Text=dt.Year.ToString();	// 
			doc5.Bookmarks["garant4"].Range.Text=textBox1.Text+" "+textBox2.Text+" "+textBox3.Text;	// 
			doc5.Bookmarks["garant5"].Range.Text=textBox16.Text;	// 
			doc5.Bookmarks["garant6"].Range.Text=dateTimePicker3.Value.Day.ToString();	// 
			doc5.Bookmarks["garant7"].Range.Text=months[dateTimePicker3.Value.Month-1];	// 
			doc5.Bookmarks["garant8"].Range.Text=dateTimePicker3.Value.Year.ToString();	// 
			doc5.Bookmarks["garant9"].Range.Text=placeWhiteSpace(textBox17.Text);	//

			
			 String fileSave = dir+"\\Гарантийное письмо "+textBox1.Text+" "+textBox2.Text+" "+textBox3.Text;
			  object fn = @fileSave;
			  object ff = Word.WdSaveFormat.wdFormatDocument;
			  object lc = false;
			  object psswd = "";
			  object f3 = false;
			  object f4 = "";
			  object f5 = false;
			  object f6 = false;
			  object f7 = false;
			  object f8 = false;
			  object f9 = Type.Missing;
			  object f10 = Type.Missing;
			  object f11 = Type.Missing;
			  object f12 = Type.Missing;
			  object f13 = Type.Missing;
			  object f14 = Type.Missing;			  
			doc5.SaveAs(ref fn, ref ff,ref lc, ref psswd, ref f3, ref f4, ref f5, ref f6, ref f7, ref f8,
			              ref f9, ref f10,ref f11, ref f12, ref f13, ref f14);			
			doc5.Close();
			temp=null;
			} catch (System.IO.IOException ioe){
				message(ioe.Message);
			}		
		}
		
		//
		//	ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ
		//
		void message(String s){
			// Вспомогательный метод для вывода окна сообщения
			MessageBox.Show(s,"");
		}
		void noValue(TextBox tb,String s){
			// Метод подсвечивает поле ввода, в которое не было введено значение
			tb.BackColor=Color.Red;
			message(s);
		}
		void noValue(TextBox tb){
			// Метод подсвечивает поле ввода, в которое не было введено значение
			tb.BackColor=Color.Red;			
		}
		void hasValue(TextBox tb){
			// Метод убирает подсветку поля ввода
			tb.BackColor=Color.White;
		}
		void Button1Click(object sender, EventArgs e)
		{
			mainapp();
		}
		void test(){
			textBox9.Text="Иванову";
			textBox8.Text="Ивану";
			textBox7.Text="Ивановичу";
			textBox19.Text="C:\\Debt\\";
			checkBox1.Checked=true;		
		}
		String placeWhiteSpace(String s){
		// Если s можно преобразовать в целое число, 
		// то вставляет пробел перед тремя младшими разрядами числа
			int a;
			if (Int32.TryParse(s, out a)) {
				int g=a/1000;
				int l=a%1000;
				if (l==0){
				return g.ToString()+" "+l.ToString()+l.ToString()+l.ToString();
				} else {
				return g.ToString()+" "+l.ToString();
				}
			} else {
				message("Сумма не является целым числом");
				return "";
			}
		}
		String cipherToChars(String s){
			// Принимает на вход целое число (сумму по договору) и переводит его в прописную форму
			string[] c5={"десять ", "двадцать ", "тридцать ", "сорок ", "пятьдесят ", "шестьдесят ", "семьдесят ", "восемьдесят ", "девяносто "};
			string[] c51={"", "одиннадцать ", "двенадцать ", "тринадцать ", "четырнадцать ", "пятнадцать ", "шестнадцать ", "семнадцать ", "восемнадцать ", "девятнадцать "};
			string[] c4={"", "одна ", "две ", "три ", "четыре ", "пять ", "шесть ", "семь ", "восемь ", "девять "};
			string[] c3={"", "сто ", "двести ", "триста ", "четыреста ", "пятьсот ", "шестьсот ", "семьсот ", "восемьсот ", "девятьсот "};
			string[] c2=c5;
			string[] c1={"", "один ", "два ", "три ", "четыре ", "пять ", "шесть ", "семь ", "восемь ", "девять "};
			string[] th={"тысяч ", "тысяча ", "тысячи ", "тысячи ", "тысячи ", "тысяч ", "тысяч ", "тысяч ", "тысяч ", "тысяч "};
			string[] rubles={"рублей ", "рубль ", "рубля ", "рубля ", "рублей ", "рублей ", "рублей ", "рублей ", "рублей "};
			int summa;
			Int32.TryParse(s, out summa);
			String output="";
			int g=summa/1000;
			// Пройдемся по двум старшим разрядам
			if (g<10) {
				output+=c4[g]+th[g];
			}
			if (g>10&&g<20){
				output+=c51[g-10]+th[0];
			}
			if (g>19){
				output+=c5[g/10-1]+c4[g%10]+th[g%10];
			}
			if (g==10){
				output+=c5[0]+th[0];
			}
			// Теперь пройдемся по трем младшим разрядам
			int l=summa%1000; 
			int l3=(l-l%100)/100;
			int l1=l%10;
			int l2=(l-l3*100-l1)/10;
			output+=c3[l3];
			if (l2==0) {
				output+=c1[l1];
			}
			if (l2==1&&l3!=0){
				output+=c51[l1];
			}
			if (l2==1&&l3==0){
				output+=c5[0];
			}
			if (l2>1){
				output+=c5[l2-1]+c1[l1];
			}
			
			return output;
			
		}
		void TextBox1Leave(object sender, EventArgs e)
		{
			// Реализуем метод дозаполнения фамилии в родительном и дательном падежах
			String tb = textBox1.Text;
			if (tb.EndsWith("ов")) {
				textBox6.Text=tb+"а";
				textBox9.Text=tb+"у";
				return;
			}
			if (tb.EndsWith("ев")) {
				textBox6.Text=tb+"а";
				textBox9.Text=tb+"у";
				return;
			}
			if (tb.EndsWith("ёв")) {
				textBox6.Text=tb+"а";
				textBox9.Text=tb+"у";
				return;
			}
			if (tb.EndsWith("кий")) {
				//textBox6.Text;
				//textBox9.Text;
				return;
			}
			if (tb.EndsWith("кой")) {
				//textBox6.Text;
				//textBox9.Text;
				return;
			}
		}
		void TextBox3Leave(object sender, EventArgs e)
		{
			// Реализуем дозаполнение отчества в родительном и дательном падежах
			String tb1=textBox3.Text;
			if (tb1.EndsWith("ич")){
				textBox4.Text=tb1+"а";
				textBox7.Text=tb1+"у";
				return;
			}
			if (tb1.EndsWith("вна")){
				textBox4.Text=tb1.Insert(tb1.Length,"ы");
				textBox7.Text=tb1.Insert(tb1.Length,"е");
				return;
			}	
		}
		void CheckBox5CheckedChanged(object sender, EventArgs e)
		{
			if (checkBox5.Checked){
				checkBox1.Checked=true;
				checkBox2.Checked=true;
				checkBox3.Checked=true;
				checkBox4.Checked=true;
				checkBox6.Checked=true;
				checkBox7.Checked=true;
			} else {
				checkBox1.Checked=false;
				checkBox2.Checked=false;
				checkBox3.Checked=false;
				checkBox4.Checked=false;
				checkBox6.Checked=false;
				checkBox7.Checked=false;
			}
		}
		void TextBox1Enter(object sender, EventArgs e)
		{
			hasValue(textBox1);
		}
		void TextBox2Enter(object sender, EventArgs e)
		{
			hasValue(textBox2);
		}
		void TextBox3Enter(object sender, EventArgs e)
		{
			hasValue(textBox3);
		}
		void TextBox4Enter(object sender, EventArgs e)
		{
			hasValue(textBox4);
		}
		void TextBox5Enter(object sender, EventArgs e)
		{
			hasValue(textBox5);
		}
		void TextBox6Enter(object sender, EventArgs e)
		{
			hasValue(textBox6);
		}
		void TextBox7Enter(object sender, EventArgs e)
		{
			hasValue(textBox7);
		}
		void TextBox8Enter(object sender, EventArgs e)
		{
			hasValue(textBox8);
		}
		void TextBox9Enter(object sender, EventArgs e)
		{
			hasValue(textBox9);
		}
		void TextBox10Enter(object sender, EventArgs e)
		{
			hasValue(textBox10);
		}
		void TextBox11Enter(object sender, EventArgs e)
		{
			hasValue(textBox11);
		}
		void TextBox12Enter(object sender, EventArgs e)
		{
			hasValue(textBox12);
		}
		void TextBox13Enter(object sender, EventArgs e)
		{
			hasValue(textBox13);
		}
		void TextBox14Enter(object sender, EventArgs e)
		{
			hasValue(textBox14);
		}
		void TextBox15Enter(object sender, EventArgs e)
		{
			hasValue(textBox15);
		}
		void TextBox16Enter(object sender, EventArgs e)
		{
			hasValue(textBox16);
		}
		void TextBox17Enter(object sender, EventArgs e)
		{
			hasValue(textBox17);
		}
		void TextBox18Enter(object sender, EventArgs e)
		{
			hasValue(textBox18);
		}
		void Button5Click(object sender, EventArgs e)
		{
			// Кнопка ОБЗОР выводит на экран окно выбора папки для сохранения
			FolderBrowserDialog	chooseDirectory = new FolderBrowserDialog();
			DialogResult folderDialogResult=chooseDirectory.ShowDialog();
			// Если пользователь выбрал папку и нажал на кнопку ОК диалога
			if (folderDialogResult.Equals(DialogResult.OK)){
				String ch = chooseDirectory.SelectedPath;
				textBox19.Text=ch;
			}			
			chooseDirectory.Dispose();	
		}
		void Button4Click(object sender, EventArgs e)
		{
			// При нажатии на кнопку ОЧИСТИТЬ удаляются значения всех текстовых полей окна
			textBox1.Text="";	textBox2.Text="";	textBox3.Text="";	textBox4.Text="";
			textBox5.Text="";	textBox6.Text="";	textBox7.Text="";	textBox8.Text="";
			textBox9.Text="";	textBox10.Text="";	textBox11.Text="";	textBox12.Text="";
			textBox13.Text="";	textBox14.Text="";	textBox15.Text="";	textBox16.Text="";
			textBox17.Text="";	textBox18.Text="";	//textBox19.Text="";	
			// А также снимаются все флажки
			checkBox1.Checked=false;	checkBox2.Checked=false;	checkBox3.Checked=false;
			checkBox4.Checked=false;	checkBox5.Checked=false;	checkBox6.Checked=false;
			checkBox7.Checked=false;	
		}
	}
	
	
	
}