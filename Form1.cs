using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using CsQuery;
using System.Net;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using iText.Kernel.Pdf;

namespace Parser
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        bool isJob = false;
        int chet = 0;
        int allCompanies = 39;


        public string downloadPage(string url)
        {
            {
                string htmlCode = "";
                using (WebClient client = new WebClient())
                {
                    client.Encoding = System.Text.Encoding.Default;
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
                    var htmlData = client.DownloadData(url);
                    htmlCode = Encoding.Default.GetString(htmlData);
                }
                return htmlCode;
            }
        }
        public string downloadPageUTF8(string url)
        {
            string htmlCode = "";
            using (WebClient client = new WebClient())
            {
                client.Encoding = System.Text.Encoding.UTF8;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
                var htmlData = client.DownloadData(url);
                htmlCode = Encoding.UTF8.GetString(htmlData);
            }
            return htmlCode;
        }

        private string getHref(string url, string htmlCode, string tag)
        {
            string href = "";
            CQ dom = CQ.CreateDocument(htmlCode);
            foreach (IDomObject obj in dom.Find(tag))
            {
                href = url + obj.GetAttribute("href");
            }
            return href;
        }

        private void managerThread(){
            
            
            if (!isJob) { 
            
            }
        
        }

        private void displayChange(string name) {
            chet++;
            int percentage = (int)(100*chet/allCompanies);
            this.Invoke(new MethodInvoker(delegate
            {
                label2.Text = name;
            }));
            backgroundWorker1.ReportProgress(percentage);
        }


        void getAllInfo()
        {
            chet = 0;
            string url = "";
            string htmlCode = "";
            CQ dom;

            //АО НПФ «Атомгарант»
            displayChange("АО НПФ «Атомгарант»");
          
                //телефон фонда
                string telnpfatom = "";

                url = "https://www.npf-atom.ru/information_disclosure/?special_version=Y";
                htmlCode = downloadPage(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".address:eq(2)"))
                {
                    string tag = obj.TextContent;
                    telnpfatom = tag.Split(':')[1];
                }
                //адрес атомгарант
                string addressnpfatommsk = "";
                foreach (IDomObject obj in dom.Find(".address:eq(0)"))
                {
                    string tag = obj.TextContent;
                    addressnpfatommsk = tag.Split(':')[1];
                }
                


                //Раздел с информацией о раскрытии на главной странице
                string razdelraskrnpfatom = "";
                url = "https://www.npf-atom.ru/";
                htmlCode = downloadPage(url);
                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".dropdown a:eq(6)"))
                {

                    razdelraskrnpfatom = "https://www.npf-atom.ru/" + obj.GetAttribute("href");
                }
                //Органы и управления контроля
                string orguprnpfatom = "";
                foreach (IDomObject obj in dom.Find(".dropdown a:eq(9)"))
                {
                    orguprnpfatom = "https://www.npf-atom.ru/" + obj.GetAttribute("href");
                }
                //Финансовая отчётность
                string finotchetnpfatom = "";
                foreach (IDomObject obj in dom.Find(".dropdown a:eq(10)"))
                {
                    finotchetnpfatom = "https://www.npf-atom.ru/" + obj.GetAttribute("href");
                }
                //Спецдепозитарий
                string specdepoznpfatom = "";
                foreach (IDomObject obj in dom.Find(".dropdown a:eq(15)"))
                {
                    specdepoznpfatom = "https://www.npf-atom.ru/" + obj.GetAttribute("href");
                }
                //о прекращенных договорах с УК
                string prkrdogovUKnpfatom = "";
                foreach (IDomObject obj in dom.Find(".dropdown a:eq(14)"))
                {
                    prkrdogovUKnpfatom = "https://www.npf-atom.ru/" + obj.GetAttribute("href");
                }
                //Структура и состав акционеров
                string sostavnpfatom = "";
                foreach (IDomObject obj in dom.Find(".dropdown a:eq(8)"))
                {
                    sostavnpfatom = "https://www.npf-atom.ru/" + obj.GetAttribute("href");
                }
                //Информация о конечных владельцах фонда
                string infkonechvladnpfatom = "";
                foreach (IDomObject obj in dom.Find(".dropdown a:eq(8)"))
                {
                    infkonechvladnpfatom = "https://www.npf-atom.ru/" + obj.GetAttribute("href");
                }


                //Устав фонда
                string ustavnpfatom = "";
                url = "https://www.npf-atom.ru/information_disclosure/dokumenty-dlya-raskrytiya/";
                htmlCode = downloadPage(url);
                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find("table a:eq(0)"))
                {
                    ustavnpfatom = "https://www.npf-atom.ru/" + obj.GetAttribute("href");
                }
                //Пенсионные правила
                string penspravilanpfatom = "";
                foreach (IDomObject obj in dom.Find("table a:eq(2)"))
                {
                    penspravilanpfatom = "https://www.npf-atom.ru/" + obj.GetAttribute("href");
                }
                //Лицензия, выдаваемая БР
                string licenznpfatom = "";
                foreach (IDomObject obj in dom.Find("table a:eq(1)"))
                {
                    licenznpfatom = "https://www.npf-atom.ru/" + obj.GetAttribute("href");
                }
                //Регистрация изменений и дополнений в пенсионные правила
                string izmendopvpenspravnpfatom = "";
                foreach (IDomObject obj in dom.Find("table a:eq(4)"))
                {
                    izmendopvpenspravnpfatom = "https://www.npf-atom.ru/" + obj.GetAttribute("href");
                }
            //Адреса обособленных подразделений
            string adrespodrazlelnpfatom = "";
            foreach (IDomObject obj in dom.Find("table a:eq(16)"))
            {
                adrespodrazlelnpfatom = "https://www.npf-atom.ru/" + obj.GetAttribute("href");
            }
            string addressnpfatom = "";
            addressnpfatom = "Место нахождения фонда:" + " " + addressnpfatommsk + " " + "Место нахождения обособленных подразделений фонда:" + " " + adrespodrazlelnpfatom;

            //Размер дохода от размещения ПР(страховой резерв) Атомагрант
            string stahovrezervnpfatom = "";
                url = "https://www.npf-atom.ru/information_disclosure/pokazateli-deyatelnosti/";
                htmlCode = downloadPage(url);
                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".col-md-9 a:eq(3)"))
                {
                    stahovrezervnpfatom = "https://www.npf-atom.ru/" + obj.GetAttribute("href");
                }
                //Количество вкадчиков и участников фонда атомгрант
                string kolvkluchas = "";
                foreach (IDomObject obj in dom.Find(".col-md-9 a:eq(3)"))
                {
                    kolvkluchas = "https://www.npf-atom.ru/" + obj.GetAttribute("href");
                }
                //Размер ПР
                string razmerprnpfatom = "";
                foreach (IDomObject obj in dom.Find(".col-md-9 a:eq(3)"))
                {
                    razmerprnpfatom = "https://www.npf-atom.ru/" + obj.GetAttribute("href");
                }

                //Результаты размещения ПР
                string rezrazmPRnpfatom = "";
                url = "https://www.npf-atom.ru/information_disclosure/rezultaty-investirovaniya/";
                htmlCode = downloadPage(url);
                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".col-md-9 a:eq(0)"))
                {
                    rezrazmPRnpfatom = "https://www.npf-atom.ru/" + obj.GetAttribute("href");
                }
                //Средневзвешенный процент
                string srvzvprocnpfatom = "";
                foreach (IDomObject obj in dom.Find(".col-md-9 a:eq(1)"))
                {
                    srvzvprocnpfatom = "https://www.npf-atom.ru/" + obj.GetAttribute("href");
                }
                //Структура инвестиционного портфеля ПР
                string structinvestportnpfatom = "";
                foreach (IDomObject obj in dom.Find(".col-md-9 a:eq(2)"))
                {
                    structinvestportnpfatom = "https://www.npf-atom.ru/" + obj.GetAttribute("href");
                }
                //Информация о процессе размещении ПР
                string procesrazmPRnpfatom = "";
                foreach (IDomObject obj in dom.Find(".col-md-9 a:eq(3)"))
                {
                    procesrazmPRnpfatom = "https://www.npf-atom.ru/" + obj.GetAttribute("href");
                }
                //Информация о событиях, существенно влияющих на стоимость активов
                string sobvlstoimactnpfatom = "";
                foreach (IDomObject obj in dom.Find(".col-md-9 a:eq(4)"))
                {
                    sobvlstoimactnpfatom = "https://www.npf-atom.ru/" + obj.GetAttribute("href");
                }
                
                //О составе средств ПР
                string sostavsredstvPRnpfatom = "";
                foreach (IDomObject obj in dom.Find(".col-md-9 a:eq(6)"))
                {
                    sostavsredstvPRnpfatom = "https://www.npf-atom.ru/" + obj.GetAttribute("href");
                }
          
                //Решения БР о запрете проведения всех или части операций
                string reshozapretenpfatom = "";
                url = "https://www.npf-atom.ru//information_disclosure/rezultaty-investirovaniya/resheniya-cbrf/";
                htmlCode = downloadPage(url);
                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".inner ul")) 
                {
                    IDomObject tmp = obj.ParentNode;
                    tmp.RemoveChild(obj);
                    reshozapretenpfatom = tmp.InnerText;
                }

            
            //НПФ СБЕРБАНК
            displayChange("НПФ СБЕРБАНК");
          
                //Раздел с раскрытием информации на главном экране
                string razdelraskrnpfsberbanka = "";
                url = "https://npfsberbanka.ru/";
                htmlCode = downloadPage(url);
                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".columns a:eq(20)"))
                {
                    razdelraskrnpfsberbanka = "https://npfsberbanka.ru/" + obj.GetAttribute("href");
                }
                //номер телефона
                string telnpfsberbanka = "";
                url = "https://npfsberbanka.ru/about/information-to-be-disclosed/";
                htmlCode = downloadPage(url);
                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".columns a:eq(0)"))
                {
                    string tag = obj.TextContent;
                    telnpfsberbanka = tag;
                }
                //адрес 
                string addressnfsberbankamsk = "";
            htmlCode = downloadPageUTF8(url);

            dom = CQ.CreateDocument(htmlCode);
            foreach (IDomObject obj in dom.Find(".address__coordinates:eq(0)"))
                {
                    string tag = obj.TextContent;
                    addressnfsberbankamsk = tag;
                }
            string adressnpfsbrbanksarov = "";
            foreach (IDomObject obj in dom.Find(".address__coordinates:eq(1)"))
            {
                adressnpfsbrbanksarov = obj.TextContent;
            }
            string addressnfsberbanka = addressnfsberbankamsk + " " + adressnpfsbrbanksarov;

            //Руководство
            string rukovodnpfsber = "";
                foreach (IDomObject obj in dom.Find(".container a:eq(0)"))
                {
                    rukovodnpfsber = "https://npfsberbanka.ru/" + obj.GetAttribute("href");
                }
                //Законодательные ограничения(Заключение и прекращ.договора со спец.депозитарием)
                string rastordogovSDnpfsber = "";
                foreach (IDomObject obj in dom.Find(".container a:eq(1)"))
                {
                    rastordogovSDnpfsber = "https://npfsberbanka.ru/" + obj.GetAttribute("href");
                }
                //Инвестиционная деятельность(Сведения о расторжении договоров с управляющими компаниями; информация о составе и структуре инвест.портфеля по ОПС) 
                string rastordogovUKnpfsber = "";
                foreach (IDomObject obj in dom.Find(".container a:eq(2)"))
                {
                    rastordogovUKnpfsber = "https://npfsberbanka.ru/" + obj.GetAttribute("href");
                }
                //Отчетность (бух отчет, аудиторское и актуарное заключение, отчет МСФО)
                string otchetnostnpfsber = "";
                foreach (IDomObject obj in dom.Find(".container a:eq(3)"))
                {
                    otchetnostnpfsber = "https://npfsberbanka.ru/" + obj.GetAttribute("href");
                }
                //Результаты деятельности (кол-во вкладчиков, уч-ков, застр.лиц; рез-т размещения ПР, рез-т инвестирования ПН; размер ПР и ПН; доходность)
                string rezuldeytnpfsberbanka = "";
                foreach (IDomObject obj in dom.Find(".container a:eq(4)"))
                {
                    rezuldeytnpfsberbanka = "https://npfsberbanka.ru/" + obj.GetAttribute("href");
                }



                //устав(ЕГРЮЛ) -
                foreach (IDomObject obj in dom.Find(".columns a:eq(5)"))
                {
                    Console.WriteLine("https://npfsberbanka.ru/" + obj.GetAttribute("href"));
                }
                //устав фонда
                string ustavnpfsberbanka = "";
                foreach (IDomObject obj in dom.Find(".columns a:eq(6)"))
                {
                    ustavnpfsberbanka = "https://npfsberbanka.ru/" + obj.GetAttribute("href");
                }
                //Правила фонда
                //Страховые правила
                string strahovpravilanpfsberbanka = "";
                foreach (IDomObject obj in dom.Find(".columns a:eq(7)"))
                {
                    strahovpravilanpfsberbanka = "https://npfsberbanka.ru/" + obj.GetAttribute("href");
                }
                //Уведомление ЦБ
                string izmenpravilnpfsberbanka = "";
                foreach (IDomObject obj in dom.Find(".columns a:eq(8)"))
                {
                    izmenpravilnpfsberbanka = "https://npfsberbanka.ru/" + obj.GetAttribute("href");
                }
                //Пенсионные правила
                string pensionpravilnpfsberbanka = "";
                foreach (IDomObject obj in dom.Find(".columns a:eq(9)"))
                {
                    pensionpravilnpfsberbanka = "https://npfsberbanka.ru/" + obj.GetAttribute("href");
                }

                //Сведения об органах управления, членах совета директоров (наблюдательного совета) фонда, должностных лицах и работниках фонда
                //Сведения о членах Сов.директоров
                foreach (IDomObject obj in dom.Find(".columns a:eq(10)"))
                {
                    Console.WriteLine("https://npfsberbanka.ru/" + obj.GetAttribute("href"));
                }
                //Генеральный директор
                foreach (IDomObject obj in dom.Find(".columns a:eq(11)"))
                {
                    Console.WriteLine("https://npfsberbanka.ru/" + obj.GetAttribute("href"));
                }
                //Заместитель ген.директора
                foreach (IDomObject obj in dom.Find(".columns a:eq(12)"))
                {
                    Console.WriteLine("https://npfsberbanka.ru/" + obj.GetAttribute("href"));
                }
                //Главный бухгалтер
                foreach (IDomObject obj in dom.Find(".columns a:eq(13)"))
                {
                    Console.WriteLine("https://npfsberbanka.ru/" + obj.GetAttribute("href"));
                }
                //Директор по развитию
                foreach (IDomObject obj in dom.Find(".columns a:eq(14)"))
                {
                    Console.WriteLine("https://npfsberbanka.ru/" + obj.GetAttribute("href"));
                }
                //Операционный директор
                foreach (IDomObject obj in dom.Find(".columns a:eq(15)"))
                {
                    Console.WriteLine("https://npfsberbanka.ru/" + obj.GetAttribute("href"));
                }
                //Контроллер
                foreach (IDomObject obj in dom.Find(".columns a:eq(16)"))
                {
                    Console.WriteLine("https://npfsberbanka.ru/" + obj.GetAttribute("href"));
                }


                //Правоустанавливающие документы АО НПФ Сбербанка

                //лицензия
                string licenznpfsberbanka = "";
                foreach (IDomObject obj in dom.Find(".columns a:eq(24)"))
                {
                    licenznpfsberbanka = "https://npfsberbanka.ru/" + obj.GetAttribute("href");
                }


                //Свид-тво о внесении НПФ в реестр НПФ - участников  системы гарантирования прав застрахованных лиц в сис-ме ОПС ( в лицензию)
                foreach (IDomObject obj in dom.Find(".columns a:eq(34)"))
                {
                    Console.WriteLine("https://npfsberbanka.ru/" + obj.GetAttribute("href"));
                }

                //Структура и состав акционеров НПФ Сербанка
                string struktsostavnpfsberbanka = "";
                foreach (IDomObject obj in dom.Find(".footer__content a:eq(1)"))
                {
                    struktsostavnpfsberbanka = "https://npfsberbanka.ru/" + obj.GetAttribute("href");
                }
            //Информация о процессе инвестирования средств ПН и ПР
            string processinvestsberbank = "";
            url = "https://npfsberbanka.ru/about/information-to-be-disclosed/policy/";
            htmlCode = downloadPage(url);
            dom = CQ.CreateDocument(htmlCode);
            foreach (IDomObject obj in dom.Find(".col.col_xs-4.col_md-8 a:eq(0)"))
            {
                processinvestsberbank = "https://npfsberbanka.ru/" + obj.GetAttribute("href");
            }



            //НПФ ЭВОЛЮЦИЯ
            displayChange("НПФ ЭВОЛЮЦИЯ");
           
                //Раздел с раскрытием информации на главном экране
                string razdelraskrevonpf = "";
                url = "https://www.evonpf.ru/";
                htmlCode = downloadPage(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".menu__categoty a:eq(6)"))
                {
                    razdelraskrevonpf = "https://www.evonpf.ru/" + obj.GetAttribute("href");
                }
                //телефон
                string televonpf = "";
                foreach (IDomObject obj in dom.Find(".phoneTime__phone a:eq(0)"))
                {
                    string tag = obj.TextContent;
                    televonpf = tag;
                }
                //устав фонда
                string ustavevonpf = "";
                url = "https://www.evonpf.ru/disclosure/";
                htmlCode = downloadPage(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".nav__submenuItem a:eq(0)"))
                {
                    ustavevonpf = "https://www.evonpf.ru/" + obj.GetAttribute("href");
                }
                //Финансовая отчётность и аудит
                string finotchetevonpf = "";
                foreach (IDomObject obj in dom.Find(".nav__submenuItem a:eq(1)"))
                {
                    finotchetevonpf = "https://www.evonpf.ru/" + obj.GetAttribute("href");
                }
                //Органы управления и контроль за деятельностью ПЕРЕДЕЛАТЬ
                string orguprevonpf = "";
                foreach (IDomObject obj in dom.Find(".nav__submenuItem a:eq(2)"))
                {
                    orguprevonpf = "https://www.evonpf.ru/" + obj.GetAttribute("href");
                }
                //Инвестирование
                string investirevonpf = "";
                foreach (IDomObject obj in dom.Find(".nav__submenuItem a:eq(3)"))
                {
                    investirevonpf = "https://www.evonpf.ru/" + obj.GetAttribute("href");
                }
                //Динамика показателей деятельности
                string dinamikaevonpf = "";
                foreach (IDomObject obj in dom.Find(".nav__submenuItem a:eq(5)"))
                {
                    dinamikaevonpf = "https://www.evonpf.ru/" + obj.GetAttribute("href");
                }

                 //Адреса подразделений
                string adrespodrazdelevonpf = "";
                foreach (IDomObject obj in dom.Find(".nav-menu__item a:eq(12)"))
                {
                    adrespodrazdelevonpf = "https://www.evonpf.ru/" + obj.GetAttribute("href");
                }

                //Сведения о договорах спец.депози.
                string specdepozevonpf = "";
                url = "https://www.evonpf.ru/disclosure/monitoring/";
                htmlCode = downloadPage(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".disclosure__inner a:eq(4)"))
                {
                    specdepozevonpf = "https://www.evonpf.ru/" + obj.GetAttribute("href");
                }
                //Управляющая компания
                string UKevonpf = "";
                url = "https://www.evonpf.ru/investments//";
                htmlCode = downloadPage(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".attachment__item a:eq(0)"))
                {
                    UKevonpf = "https://www.evonpf.ru/" + obj.GetAttribute("href");
                }

                //Состав и структура акционеров и информация о конечных владельцах фонда
                string sostavevonpf = "";
                url = "https://www.evonpf.ru/";
                displayChange(url);
                htmlCode = downloadPage(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".footer__copytight a:eq(0)"))
                {
                    sostavevonpf = obj.GetAttribute("href");
                }
                //адрес
                string adressnpfbmsk = "";
                url = "https://www.evonpf.ru/disclosure/requisites/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".requisites__itemRight:eq(1)"))
                {
                    adressnpfbmsk = obj.TextContent;
                }
            string adressnpfb = "";
            adressnpfb = "Головной офис:" + adressnpfbmsk + "\n" + "Адреса обособленных подразделений: " + adrespodrazdelevonpf;

                //Состав инвестиционного портфеля
                string sostavportfevonpf = "";
                url = "https://www.evonpf.ru/";
                htmlCode = downloadPage(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".menu__section a:eq(11)"))
                {
                    sostavportfevonpf = "https://www.evonpf.ru/" + obj.GetAttribute("href");
                }
            //Информация о процессе инвестирования средств
            string procinvestevonpf = "";
            url = "https://www.evonpf.ru/investments/invest_process/";
            htmlCode = downloadPage(url);

            dom = CQ.CreateDocument(htmlCode);
            foreach (IDomObject obj in dom.Find(".attachment__inner a:eq(0)"))
            {
                procinvestevonpf = "https://www.evonpf.ru/" + obj.GetAttribute("href");
            }

            //НПФ ТЕЛЕКОМ-СОЮЗ
            displayChange("НПФ ТЕЛЕКОМ-СОЮЗ");

            //контакты отделений
            string adrespodrazdelnpfts = "";
            url = "https://www.npfts.ru/contacts/centr/";
            htmlCode = downloadPage(url);

            dom = CQ.CreateDocument(htmlCode);
            foreach (IDomObject obj in dom.Find("#navleft a:eq(1)"))
            {
                adrespodrazdelnpfts = "https://www.npfts.ru/" + obj.GetAttribute("href");
            }

            //раздел с раскрытием на главной странице
            string razdelraskrnpfts = "";
                url = "https://www.npfts.ru/";
                htmlCode = downloadPage(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".conteiner a:eq(25)"))
                {
                    razdelraskrnpfts = "https://www.npfts.ru/" + obj.GetAttribute("href");
                }
                //телефон
                string telnpfts = "";
                url = "https://www.npfts.ru/disclosure/info/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".phone:eq(1)"))
                {
                    telnpfts = obj.TextContent;
                }
                //адрес
                htmlCode = htmlCode.Replace("<br>", "|");
                dom = CQ.CreateDocument(htmlCode);

                string adrestsmsk = "";
                foreach (IDomObject obj in dom.Find(".fil"))
                {
                    string[] arr = obj.TextContent.Split('|');
                    adrestsmsk = arr[1] + " " + arr[2];
                }
            string adrests = "";
            adrests = "Адрес головного офиса: " + adrestsmsk + "\n" + "Адреса обособленных подразделений: " + adrespodrazdelnpfts;

                //Управляющие компании
                string UKnpfts = "";
                foreach (IDomObject obj in dom.Find(".act a:eq(2)"))
                {
                    UKnpfts = "https://www.npfts.ru/" + obj.GetAttribute("href");
                }
                //Спец.депозитарий
                string SDnpfts = "";
                foreach (IDomObject obj in dom.Find(".act a:eq(3)"))
                {
                    SDnpfts = "https://www.npfts.ru/" + obj.GetAttribute("href");
                }
                //Бухгалтерская (финансовая) отчетность, Аудиторские, Актуарные заключения
                string otchetnpfts = "";
                foreach (IDomObject obj in dom.Find(".act a:eq(4)"))
                {
                    otchetnpfts = "https://www.npfts.ru/" + obj.GetAttribute("href");
                }
                //Показатели деятельности(Размер пенсионных накоплений, количество застрахованных лиц, (результат размещения ПР, результат инвестирования ПН, средневз.процент, размер дохода от размещения, отчёт о формировании ПН - только за 2020 год))
                string pokazatelnpfts = "";
                foreach (IDomObject obj in dom.Find(".act a:eq(6)"))
                {
                    pokazatelnpfts = "https://www.npfts.ru/" + obj.GetAttribute("href");
                }
                //Состав и структура инвестиционного портфеля фонда
                string investportfnpfts = "";
                foreach (IDomObject obj in dom.Find(".act a:eq(7)"))
                {
                    investportfnpfts = "https://www.npfts.ru/" + obj.GetAttribute("href");
                }
                //Устав
                string ustavnpfts = "";
                foreach (IDomObject obj in dom.Find(".act a:eq(8)"))
                {
                    ustavnpfts = "https://www.npfts.ru/" + obj.GetAttribute("href");
                }
                //Акционеры и конечные владельцы
                string sostavakcionernpfts = "";
                foreach (IDomObject obj in dom.Find(".act a:eq(10)"))
                {
                    sostavakcionernpfts = "https://www.npfts.ru/" + obj.GetAttribute("href");
                }
            //информация о процессе инвестирования
            string procinvestts = "";
            foreach (IDomObject obj in dom.Find(".content__ a:eq(4)"))
            {
                procinvestts = "https://www.npfts.ru/" + obj.GetAttribute("href");
            }
            //Информация о событиях (действиях), оказывающих существенное влияние на совокупную стоимость активов АО «НПФ «Телеком-Союз»
            string sobytnpfts = "";
                foreach (IDomObject obj in dom.Find(".content__ a:eq(5)"))
                {
                    sobytnpfts = "https://www.npfts.ru/" + obj.GetAttribute("href");
                }
                //Страховые правила
                string strahpravnpfts = "";
                url = "https://www.npfts.ru/disclosure/documents/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".content__ a:eq(33)"))
                {
                    strahpravnpfts = "https://www.npfts.ru/" + obj.GetAttribute("href");
                }
                //Пенсионные правила
                string penspravnpfts = "";
                foreach (IDomObject obj in dom.Find(".content__ a:eq(34)"))
                {
                    penspravnpfts = "https://www.npfts.ru/" + obj.GetAttribute("href");
                }
                //Лицензия
                string licenznpfts = "";
                foreach (IDomObject obj in dom.Find(".content__ a:eq(35)"))
                {
                    licenznpfts = "https://www.npfts.ru/" + obj.GetAttribute("href");
                }
                //Изменения и дополнения в пенсионные правила(только с архива)
                string doppenspravnpfts = "";
                foreach (IDomObject obj in dom.Find(".content__ a:eq(55)"))
                {
                    doppenspravnpfts = "https://www.npfts.ru/" + obj.GetAttribute("href");
                }

         

            //НПФ СУРГУТНЕФТЕГАЗ
            displayChange("НПФ СУРГУТНЕФТЕГАЗ");
          
                //раздел с раскрытием на главной странице
                string raskrytnpfsng = "";
                url = "https://npf-sng.ru/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".footer__item a:eq(2)"))
                {
                    raskrytnpfsng = "https://npf-sng.ru/" + obj.GetAttribute("href");
                }
                //лицензия
                string licenznpfsng = "";
                url = "https://npf-sng.ru/allinfo/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".allinfo a:eq(0)"))
                {
                    licenznpfsng = "https://npf-sng.ru/" + obj.GetAttribute("href");
                }
                //устав
                string ustavnpfsng = "";
                foreach (IDomObject obj in dom.Find(".allinfo a:eq(1)"))
                {
                    ustavnpfsng = "https://npf-sng.ru/" + obj.GetAttribute("href");
                }
                //пенсионные правила
                string penspravnpfsng = "";
                foreach (IDomObject obj in dom.Find(".allinfo a:eq(2)"))
                {
                    penspravnpfsng = "https://npf-sng.ru/" + obj.GetAttribute("href");
                }
                //страховые правила
                string strahpravnpfsng = "";
                foreach (IDomObject obj in dom.Find(".allinfo a:eq(3)"))
                {
                    strahpravnpfsng = "https://npf-sng.ru/" + obj.GetAttribute("href");
                }

                //Фонд должен раскрывать сведения о принятии Банком России решения о запрете на проведение фондом всех или части операций (с указанием перечня таких операций, даты введения запрета и срока, на который введен запрет
                string zapretnpfsng = "";
                foreach (IDomObject obj in dom.Find(".sng__section p:last"))
                {
                    zapretnpfsng = obj.TextContent;
                }
                //адрес
                string adressnpfsng = "";
                foreach (IDomObject obj in dom.Find(".sng__section p:first"))
                {
                    adressnpfsng = obj.TextContent;
                }





                //Аудиторские заключения (ЗАНЕСТИ В ЯЧЕЙКУ D)
                string audzaclnpfsng = "";
                foreach (IDomObject obj in dom.Find(".allinfo a:eq(11)"))
                {
                    audzaclnpfsng = "https://npf-sng.ru/" + obj.GetAttribute("href");
                }
                //Актуарные заключения(ЗАНЕСТИ В ЯЧЕЙКУ D)
                string actuarzaclnpfsng = "";
                foreach (IDomObject obj in dom.Find(".allinfo a:eq(12)"))
                {
                    actuarzaclnpfsng = "https://npf-sng.ru/" + obj.GetAttribute("href");
                }
                //Бухгалтерская(финанссовая) отченость
                string buhothetnpfsng = "";
                foreach (IDomObject obj in dom.Find(".allinfo a:eq(10)"))
                {
                    buhothetnpfsng = "https://npf-sng.ru/" + obj.GetAttribute("href");
                }

                string dlyaYacheykiD = buhothetnpfsng + " " + audzaclnpfsng + " " + actuarzaclnpfsng;






                //Финансовая отчетность по МСФО и аудиторское заключение 
                string MSFOnpfsng = "";
                foreach (IDomObject obj in dom.Find(".allinfo a:eq(18)"))
                {
                    MSFOnpfsng = "https://npf-sng.ru/" + obj.GetAttribute("href");
                }
                //Информация о процессе инвестирования средств пенсионных накоплений и размещения средств пенсионных резервов
                string procinvestnpfsng = "";
                foreach (IDomObject obj in dom.Find(".allinfo a:eq(19)"))
                {
                    procinvestnpfsng = "https://npf-sng.ru/" + obj.GetAttribute("href");
                }
                //Список акционеров фонда и лиц, под контролем либо значительным влиянием которых находится фонд
                string spisokakcionnpfsng = "";
                foreach (IDomObject obj in dom.Find(".allinfo a:eq(13)"))
                {
                    spisokakcionnpfsng = "https://npf-sng.ru/" + obj.GetAttribute("href");
                }
                //Показатели деятельности(кол-во застрахованных лиц, вкладчиков, участников, доход от размещения ПР, на страховой резерв, размер ПР) (отчёт о форммировании средств ПН находится в формате doc)
                string pokazatelnpfsng = "";
                foreach (IDomObject obj in dom.Find(".allinfo a:eq(14)"))
                {
                    pokazatelnpfsng = "https://npf-sng.ru/" + obj.GetAttribute("href");
                }
                //Сведения о депозитарии
                string SDnpfsng = "";
                foreach (IDomObject obj in dom.Find(".allinfo a:eq(17)"))
                {
                    SDnpfsng = "https://npf-sng.ru/" + obj.GetAttribute("href");
                }
                //Информация о событиях (действиях), оказывающих, по мнению фонда, существенное влияние на совокупную стоимость активов, в которые инвестированы средства пенсионных накоплений и размещены средства пенсионных резервов
                string sobytnpfsng = "";
                foreach (IDomObject obj in dom.Find(".allinfo a:eq(20)"))
                {
                    sobytnpfsng = "https://npf-sng.ru/" + obj.GetAttribute("href");
                }
                //Сведения об органах управления, членах совета директоров (наблюдательного совета), должностных лицах и работниках АО «НПФ «Сургутнефтегаз»
                string orguprnpfsng = "";
                foreach (IDomObject obj in dom.Find(".allinfo a:eq(21)"))
                {
                    orguprnpfsng = "https://npf-sng.ru/" + obj.GetAttribute("href");
                }
                //Информация о составе инвестиционного портфеля по обязательному пенсионному страхованию
                string investportfelnpfsng = "";
                foreach (IDomObject obj in dom.Find(".allinfo a:eq(24)"))
                {
                    investportfelnpfsng = "https://npf-sng.ru/" + obj.GetAttribute("href");
                }
                //Информация о составе средств пенсионных резервов (ВНЕСТИ В ЯЧЕЙКУ AB)
                string sostavsrPRnpfsng = "";
                foreach (IDomObject obj in dom.Find(".allinfo a:eq(25)"))
                {
                    sostavsrPRnpfsng = "https://npf-sng.ru/" + obj.GetAttribute("href");
                }
                //номер телефона
                string telnpfsng = "";
                foreach (IDomObject obj in dom.Find(".nav__phone-number:eq(0)"))
                {
                    telnpfsng = obj.TextContent;
                }
                //Результаты размещения ПН
                string rezOPSnpfsng = "";
                url = "https://npf-sng.ru/results/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".submenu_button a:eq(1)"))
                {
                    rezOPSnpfsng = "https://npf-sng.ru/" + obj.GetAttribute("href");
                }
                //Результат размещения ПР
                string srvzvprocnpfsng = "";
                foreach (IDomObject obj in dom.Find(".submenu_button a:eq(0)"))
                {
                    srvzvprocnpfsng = "https://npf-sng.ru/" + obj.GetAttribute("href");
                }
                //УК
                string UKnpfsng = "";
                foreach (IDomObject obj in dom.Find(".submenu_button a:eq(2)"))
                {
                    UKnpfsng = "https://npf-sng.ru/" + obj.GetAttribute("href");
                }

          

            //НПФ ТРАНСНЕФТЬ
            displayChange("НПФ ТРАНСНЕФТЬ");
          

                //раздел с раскрытием на главной странице
                string raskrytnpftransneft = "";
                url = "http://www.npf-transneft.ru/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".left-menu a:eq(10)"))
                {
                    raskrytnpftransneft = "http://www.npf-transneft.ru/" + obj.GetAttribute("href");
                }
                //устав
                string ustavtransneft = "";
                url = "http://www.npf-transneft.ru/Information/ustav/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".content a:eq(0)"))
                {
                    ustavtransneft = obj.GetAttribute("href");
                }
                //лицензия
                string licenztransneft = "";
                foreach (IDomObject obj in dom.Find(".content a:eq(1)"))
                {
                    licenztransneft = obj.GetAttribute("href");
                }
                //УК
                string UKtransneft = "";
                foreach (IDomObject obj in dom.Find(".sub a:eq(8)"))
                {
                    UKtransneft = "http://www.npf-transneft.ru/" + obj.GetAttribute("href");
                }
                //СД
                string SDtransneft = "";
                foreach (IDomObject obj in dom.Find(".sub a:eq(12)"))
                {
                    SDtransneft = "http://www.npf-transneft.ru/" + obj.GetAttribute("href");
                }
                //страховые правила
                string strahpravtransneft = "";
                foreach (IDomObject obj in dom.Find(".sub a:eq(10)"))
                {
                    strahpravtransneft = "http://www.npf-transneft.ru/" + obj.GetAttribute("href");
                }
                //пенсионные правила
                string penspravtransneft = "";
                foreach (IDomObject obj in dom.Find(".sub a:eq(11)"))
                {
                    penspravtransneft = "http://www.npf-transneft.ru/" + obj.GetAttribute("href");
                }
                //Показатели деятельности
                string pokazateltransneft = "";
                foreach (IDomObject obj in dom.Find(".sub a:eq(18)"))
                {
                    pokazateltransneft = "http://www.npf-transneft.ru/" + obj.GetAttribute("href");
                }
                //инвестирование портфелей структура и состав
                string portftransneft = "";
                foreach (IDomObject obj in dom.Find(".sub a:eq(19)"))
                {
                    portftransneft = "http://www.npf-transneft.ru/" + obj.GetAttribute("href");
                }
                //отчётность (отчет о формировании средств ПН только за 2020 год)
                string otchettransneft = "";
                foreach (IDomObject obj in dom.Find(".sub a:eq(20)"))
                {
                    otchettransneft = "http://www.npf-transneft.ru/" + obj.GetAttribute("href");
                }
                //Органы управления фондом
                string orguprtransneft = "";
                foreach (IDomObject obj in dom.Find(".sub a:eq(13)"))
                {
                    orguprtransneft = "http://www.npf-transneft.ru/" + obj.GetAttribute("href");
                }

                //адрес 
                string adresstransneft = "";
                url = "http://www.npf-transneft.ru/Information/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".content p:eq(5)"))
                {
                    adresstransneft = obj.TextContent;
                }
                //телефон
                string teltransneft = "";
                foreach (IDomObject obj in dom.Find(".content p:eq(6)"))
                {
                    teltransneft = obj.TextContent;
                }
                //структура и состав акционеров
                string sostavakctransneft = "";
                foreach (IDomObject obj in dom.Find(".content a:eq(8)"))
                {
                    sostavakctransneft = obj.GetAttribute("href");
                }
                //информация о событиях (действиях), оказывающих, по мнению фонда, существенное влияние на совокупную стоимость активов, в которые инвестированы средства пенсионных накоплений и размещены средства пенсионных резервов.
                string inftransneft = "";
                url = "http://www.npf-transneft.ru/Information/investirovanie-sredstv-pensionnih-nakoplenii-i/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".content p:last"))
                {
                    inftransneft = obj.TextContent;
                }



            //НПФ АЛЬЯНС
            displayChange("НПФ АЛЬЯНС");
          

                //раздел с раскрытием на главной странице
                string raskrallians = "";
                url = "https://www.npfalliance.ru/content/main";
                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".tabs a:eq(25)"))
                {
                    raskrallians = "https://www.npfalliance.ru/content/" + obj.GetAttribute("href");
                }
                //телефон
                string telalliance = "";
                foreach (IDomObject obj in dom.Find(".phone:eq(0)"))
                {
                    telalliance = obj.TextContent;
                }
                //Структура и состав акционеров
                string sostavakcionallians = "";
                url = "https://www.npfalliance.ru/content/disclosure_info";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".hierarchical-expanded a:eq(0)"))
                {
                    sostavakcionallians = "https://www.npfalliance.ru/content/" + obj.GetAttribute("href");
                }
                //Отчетность
                string othetalliance = "";
                foreach (IDomObject obj in dom.Find(".hierarchical-expanded a:eq(1)"))
                {
                    othetalliance = "https://www.npfalliance.ru/content/" + obj.GetAttribute("href");
                }
                //показатели
                string pokazatelalliance = "";
                foreach (IDomObject obj in dom.Find(".hierarchical-expanded a:eq(2)"))
                {
                    pokazatelalliance = "https://www.npfalliance.ru/content/" + obj.GetAttribute("href");
                }
                //инвестиционный портфель структура и состав
                string strportfalliance = "";
                foreach (IDomObject obj in dom.Find(".hierarchical-expanded a:eq(5)"))
                {
                    strportfalliance = "https://www.npfalliance.ru/content/" + obj.GetAttribute("href");
                }
                //УК
                string UKalliance = "";
                foreach (IDomObject obj in dom.Find(".hierarchical-expanded a:eq(6)"))
                {
                    UKalliance = "https://www.npfalliance.ru/content/" + obj.GetAttribute("href");
                }
                //СД
                string SDalliance = "";
                foreach (IDomObject obj in dom.Find(".hierarchical-expanded a:eq(7)"))
                {
                    SDalliance = "https://www.npfalliance.ru/content/" + obj.GetAttribute("href");
                }
                //лицензия
                string licenzalliance = "";
                url = "https://www.npfalliance.ru/content/disclosure_docs";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".highlight-row a:eq(0)"))
                {
                    licenzalliance = "https://www.npfalliance.ru/content/" + obj.GetAttribute("href");
                }
                //устав
                string ustavalliance = "";
                foreach (IDomObject obj in dom.Find(".highlight-row a:eq(1)"))
                {
                    ustavalliance = "https://www.npfalliance.ru/content/" + obj.GetAttribute("href");
                }
                //пенсионные правила
                string penspravalliance = "";
                foreach (IDomObject obj in dom.Find(".highlight-row a:eq(6)"))
                {
                    penspravalliance = "https://www.npfalliance.ru/content/" + obj.GetAttribute("href");
                }
                //страховые правила
                string strahpravalliance = "";
                foreach (IDomObject obj in dom.Find(".highlight-row a:eq(7)"))
                {
                    strahpravalliance = "https://www.npfalliance.ru/content/" + obj.GetAttribute("href");
                }
                //Руководство
                string rukovodalliance = "";
                url = "https://www.npfalliance.ru/content/about_fund";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".hierarchical-expanded a:eq(1)"))
                {
                    rukovodalliance = "https://www.npfalliance.ru/content/" + obj.GetAttribute("href");
                }
                //адрес 
                string adressalliance = "";
                url = "https://www.npfalliance.ru/content/contacts";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".contact_block p:last"))
                {
                    adressalliance = obj.TextContent;
                }


            //ПЕРВЫЙ ПРОМЫШЛЕННЫЙ АЛЬЯНС
            displayChange("ПЕРВЫЙ ПРОМЫШЛЕННЫЙ АЛЬЯНС");
       
                //Руководство
                string rukovodppafond = "";
                url = "https://ppafond.ru/about/controls.html";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".table a:eq(0)"))
                {
                    rukovodppafond = "https://ppafond.ru/" + obj.GetAttribute("href");
                }
                //Устав
                string ustavppafond = "";
                url = "https://ppafond.ru/information/documents.html";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".table a:eq(0)"))
                {
                    ustavppafond = "https://ppafond.ru/" + obj.GetAttribute("href");
                }
                //лицензия
                string licenzppafond = "";
                foreach (IDomObject obj in dom.Find(".table a:eq(4)"))
                {
                    licenzppafond = "https://ppafond.ru/" + obj.GetAttribute("href");
                }
                //пенсионные правила
                string penspravppafond = "";
                foreach (IDomObject obj in dom.Find(".table a:eq(25)"))
                {
                    penspravppafond = "https://ppafond.ru/" + obj.GetAttribute("href");
                }
                //страховые правила
                string strahpravppafond = "";
                foreach (IDomObject obj in dom.Find(".table a:eq(28)"))
                {
                    strahpravppafond = "https://ppafond.ru/" + obj.GetAttribute("href");
                }
                //адрес
                string adresppafondgolov = "";
                url = "https://ppafond.ru/information/opendata.html";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".about_text p:eq(7)"))
                {
                    adresppafondgolov = obj.TextContent;
                }
                //адрес филиала
            string adresppafondfillial = "";
            foreach (IDomObject obj in dom.Find(".about_text p:eq(9)"))
            {
                adresppafondfillial = obj.TextContent;
            }

            string adresppafond = "";
            adresppafond = adresppafondgolov + "\n" + adresppafondfillial;



            //телефон ДОДЕЛАЙ
            string telppafond = "";
                url = "https://ppafond.ru/contact.html";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".header-tel h2"))
                {
                    telppafond = obj.TextContent.Split('.')[1];
                }


                //инвестиционная политика
                string portfppafond = "";
                url = "https://ppafond.ru/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".dropdown-menu a:eq(18)"))
                {
                    portfppafond = obj.GetAttribute("href");
                }
                //Показатели деятельности
                string pokazatelppafond = "";
                foreach (IDomObject obj in dom.Find(".dropdown-menu a:eq(19)"))
                {
                    pokazatelppafond = obj.GetAttribute("href");
                }
                //Отчётность
                string otchetppafond = "";
                foreach (IDomObject obj in dom.Find(".dropdown-menu a:eq(20)"))
                {
                    otchetppafond = obj.GetAttribute("href");
                }
                //УК
                string UKppafond = "";
                foreach (IDomObject obj in dom.Find(".dropdown-menu a:eq(21)"))
                {
                    UKppafond = obj.GetAttribute("href");
                }
                //СД
                string SDppafond = "";
                foreach (IDomObject obj in dom.Find(".dropdown-menu a:eq(22)"))
                {
                    SDppafond = obj.GetAttribute("href");
                }
                //Акционеры и конечные владельцы
                string akcionerppafond = "";
                foreach (IDomObject obj in dom.Find(".dropdown-menu a:eq(25)"))
                {
                    akcionerppafond = obj.GetAttribute("href");
                }
            //Информация о событиях
            string sobytppafond = "";
            url = "https://ppafond.ru/information/invest-politics.html";
            htmlCode = downloadPageUTF8(url);

            dom = CQ.CreateDocument(htmlCode);
            foreach (IDomObject obj in dom.Find(".table > tr > a:eq(1)"))
            {
                sobytppafond = "https://ppafond.ru/" + obj.GetAttribute("href");
            }





            //НПФ АЛМАЗНАЯ ОСЕНЬ
            displayChange("НПФ АЛМАЗНАЯ ОСЕНЬ");
         

                //раздел с раскрытием на главной странице
                string raskrnpfao = "";
                url = "https://www.npfao.ru/";
                htmlCode = downloadPage(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".main-menu a:eq(4)"))
                {
                    raskrnpfao = "https://www.npfao.ru/" + obj.GetAttribute("href");
                }
                //лицензия
                string licenznpfao = "";
                url = "https://www.npfao.ru/information_disclosure/documentation-and-reporting/documents/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".blue-section a:eq(0)"))
                {
                    licenznpfao = "https://www.npfao.ru/" + obj.GetAttribute("href");
                }
                //Устав
                string ustavnpfao = "";
                foreach (IDomObject obj in dom.Find(".blue-section a:eq(2)"))
                {
                    ustavnpfao = "https://www.npfao.ru/" + obj.GetAttribute("href");
                }
                //Пенсионные правила
                string pensstrahpravnpfao = "";
                url = "https://www.npfao.ru/information_disclosure/documentation-and-reporting/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".help-section a:eq(0)"))
                {
                    pensstrahpravnpfao = "https://www.npfao.ru/" + obj.GetAttribute("href");
                }
                //Отчётность
                string otchetnpfao = "";
                foreach (IDomObject obj in dom.Find(".help-section a:eq(1)"))
                {
                    otchetnpfao = "https://www.npfao.ru/" + obj.GetAttribute("href");
                }
                //УК и СД
                string UKSDnpfao = "";
                foreach (IDomObject obj in dom.Find(".help-section a:eq(2)"))
                {
                    UKSDnpfao = "https://www.npfao.ru/" + obj.GetAttribute("href");
                }
                //телефон
                string telnpfao = "";
                url = "https://www.npfao.ru/information_disclosure/cb/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".container a:eq(1)"))
                {
                    telnpfao = obj.GetAttribute("href");
                }
                //структура и состав акционеров
                string sostavakcionernpfao = "";
                foreach (IDomObject obj in dom.Find(".container a:eq(26)"))
                {
                    sostavakcionernpfao = obj.GetAttribute("href");
                }
                //Размер дохода от размещения пенсионных резервов, направляемого на формирование страхового резерва фонда
               string razmerdohodaPRnpfao = "";
               foreach (IDomObject obj in dom.Find(".container a:eq(32)"))
               {
                razmerdohodaPRnpfao = "https://www.npfao.ru/" + obj.GetAttribute("href");
               }
               //Количество вкладчиков и участников фонда, а также участников фонда, получающих из фонда негосударственную пенсию
               string uchastnpfao = "";
                foreach (IDomObject obj in dom.Find(".container a:eq(34)"))
                {
                    uchastnpfao = "https://www.npfao.ru/" + obj.GetAttribute("href");
                }
                //Количество застрахованных лиц, осуществляющих формирование своих пенсионных накоплений в фонде
                string zastrahlicnpfao = "";
                foreach (IDomObject obj in dom.Find(".container a:eq(40)"))
                {
                    zastrahlicnpfao = "https://www.npfao.ru/" + obj.GetAttribute("href");
                }
                //Размер пенсионных резервов фонда (в том числе страхового резерва) и пенсионных накоплений фонда (в том числе резерва по ОПС, выплатного резерва и средств застрахованных лиц, которым установлена срочная пенсионная выплата)  
                string PRnpfao = "";
                foreach (IDomObject obj in dom.Find(".container a:eq(44)"))
                {
                    PRnpfao = "https://www.npfao.ru/" + obj.GetAttribute("href");
                }
                //Результаты инвестирования средств пенсионных накоплений
                string investPNnpfao = "";
                foreach (IDomObject obj in dom.Find(".container a:eq(67)"))
                {
                    investPNnpfao = "https://www.npfao.ru/" + obj.GetAttribute("href");
                }

                //Результат инвестирования пенсионных резервов
                string investPRnpfao = "";
                foreach (IDomObject obj in dom.Find(".container a:eq(69)"))
                {
                    investPRnpfao = "https://www.npfao.ru/" + obj.GetAttribute("href");
                }
               
                //Место нахождения фонда и его обособленных подразделений 
                string adresnpfao = "";
                foreach (IDomObject obj in dom.Find(".container a:eq(23)"))
                {
                    adresnpfao = "https://www.npfao.ru/" + obj.GetAttribute("href");
                }
                //Средневзвешенный процент
                string srvzvprocnpfao = "";
                foreach (IDomObject obj in dom.Find(".container a:eq(70)"))
                {
                    srvzvprocnpfao = "https://www.npfao.ru/" + obj.GetAttribute("href");
                }
                //Структура средств пенсионных резервов фонда
                string structuraPRnpfao = "";
                foreach (IDomObject obj in dom.Find(".container a:eq(73)"))
                {
                    structuraPRnpfao = "https://www.npfao.ru/" + obj.GetAttribute("href");
                }
                //Структура инвестиционного портфеля фонда ПН 
                string structuraPNnpfao = "";
                foreach (IDomObject obj in dom.Find(".container a:eq(77)"))
                {
                    structuraPNnpfao = "https://www.npfao.ru/" + obj.GetAttribute("href");
                }
                //Структура инвест.портвела ПР и ПН (в одну ячейку)
                string structuraPRPNnpfao = structuraPNnpfao + " " + structuraPRnpfao;

                //Информация о составе средств пенсионных резервов фонда
                string sostavPRnpfao = "";
                foreach (IDomObject obj in dom.Find(".container a:eq(80)"))
                {
                    sostavPRnpfao = "https://www.npfao.ru/" + obj.GetAttribute("href");
                }
                //Информация о составе инвестиционного портфеля фонда по обязательному пенсионному страхованию
                string sostavPNnpfao = "";
                foreach (IDomObject obj in dom.Find(".container a:eq(87)"))
                {
                    sostavPNnpfao = "https://www.npfao.ru/" + obj.GetAttribute("href");
                }
                //Состав инвест.портвеля ПР и ПН (в одну ячейку)
                string sostavPRPNnpfao = sostavPRnpfao + " " + sostavPNnpfao;
                //Решения о запрете
                string zaprnpfao = "";
                foreach (IDomObject obj in dom.Find(".container p:last"))
                {
                    zaprnpfao = obj.TextContent;
                }
            //Информация о процессе инвестирования
            //ПР
            string procinvestPRnpfao = "";
            foreach (IDomObject obj in dom.Find(".container a:eq(95)"))
            {
                procinvestPRnpfao = "https://www.npfao.ru/" + obj.GetAttribute("href");
            }
            //ПН
            string procinvestPNnpfao = "";
            foreach (IDomObject obj in dom.Find(".container a:eq(97)"))
            {
                procinvestPNnpfao = "https://www.npfao.ru/" + obj.GetAttribute("href");
            }

            string procinvestPRPNnpfao = "";
            procinvestPRPNnpfao = procinvestPRnpfao + " " + procinvestPNnpfao;


            //Информация о событиях
            string infsobnpfao = "";
                foreach (IDomObject obj in dom.Find(".container p:eq(73)"))
                {
                    infsobnpfao = obj.TextContent;
                }

   

            //НПФ СТРОЙКОМПЛЕКС
            displayChange("НПФ СТРОЙКОМПЛЕКС");

            //Лицензия
            string licenzstroykomplex = "";
            url = "https://npf-stroycomplex.ru/about/information/";
            htmlCode = downloadPageUTF8(url);

            dom = CQ.CreateDocument(htmlCode);
            foreach (IDomObject obj in dom.Find(".tab-pane:eq(12) a:eq(0)"))
            {
                licenzstroykomplex = "https://npf-stroycomplex.ru/" + obj.GetAttribute("href");
            }
            //Место нахождения Фонда и его обособленных подразделений
            string adresstroykomplex = "";
            foreach (IDomObject obj in dom.Find(".tab-pane:eq(12) a:eq(2)"))
            {
                adresstroykomplex = "https://npf-stroycomplex.ru/" + obj.GetAttribute("href");
            }
            //Бухгалтерская (финансовая) отчетность Фонда, аудиторское и актуарное заключение
            string otchetstroykomplex = "";
            foreach (IDomObject obj in dom.Find(".tab-pane:eq(12) a:eq(3)"))
            {
                otchetstroykomplex = "https://npf-stroycomplex.ru/" + obj.GetAttribute("href");
            }
            //Структура и состав акционеров Фонда
            string akcionstroykomplex = "";
            foreach (IDomObject obj in dom.Find(".tab-pane:eq(12) a:eq(4)"))
            {
                akcionstroykomplex = "https://npf-stroycomplex.ru/" + obj.GetAttribute("href");
            }
            //Результат инвестирования пенсионных резервов
            string rezinvestPRstroykomplex = "";
            foreach (IDomObject obj in dom.Find(".tab-pane:eq(12) a:eq(24)"))
            {
                rezinvestPRstroykomplex = "https://npf-stroycomplex.ru/" + obj.GetAttribute("href");
            }
            //	Результат инвестирования пенсионных накоплений
            string rezinvestPNstroykomplex = "";
            foreach (IDomObject obj in dom.Find(".tab-pane:eq(12) a:eq(23)"))
            {
                rezinvestPNstroykomplex = "https://npf-stroycomplex.ru/" + obj.GetAttribute("href");
            }
            //Размер дохода от размещения пенсионных резервов, направляемого на формирование страхового резерва Фонда
            string razmerdohodaPRstroykomplex = "";
            foreach (IDomObject obj in dom.Find(".tab-pane:eq(12) a:eq(7)"))
            {
                razmerdohodaPRstroykomplex = "https://npf-stroycomplex.ru/" + obj.GetAttribute("href");
            }
            //Количество вкладчиков и участников Фонда, а также участников Фонда, получающих из Фонда негосударственную пенсию
            string koluchstroykomplex = "";
            foreach (IDomObject obj in dom.Find(".tab-pane:eq(12) a:eq(8)"))
            {
                koluchstroykomplex = "https://npf-stroycomplex.ru/" + obj.GetAttribute("href");
            }
            //Количество застрахованных лиц, осуществляющих формирование своих пенсионных накоплений в Фонде 
            string kolzastrstroykomplex = "";
            foreach (IDomObject obj in dom.Find(".tab-pane:eq(12) a:eq(9)"))
            {
                kolzastrstroykomplex = "https://npf-stroycomplex.ru/" + obj.GetAttribute("href");
            }
            //Размер пенсионных резервов Фонда (в том числе страхового резерва) и пенсионных накоплений Фонда 
            string razmerPRstroykomplex = "";
            foreach (IDomObject obj in dom.Find(".tab-pane:eq(12) a:eq(10)"))
            {
                razmerPRstroykomplex = "https://npf-stroycomplex.ru/" + obj.GetAttribute("href");
            }
            //Информация о заключении и прекращении действия договора доверительного управления пенсионными резервами или пенсионными накоплениями с управляющей компанией с указанием ее фирменного наименования и номера лицензии
            string UKstroykomplex = "";
            foreach (IDomObject obj in dom.Find(".tab-pane:eq(12) a:eq(11)"))
            {
                UKstroykomplex = "https://npf-stroycomplex.ru/" + obj.GetAttribute("href");
            }
            //Информация о заключении и прекращении договора со специализированным депозитарием   
            string SDstroykomplex = "";
            foreach (IDomObject obj in dom.Find(".tab-pane:eq(12) a:eq(12)"))
            {
                SDstroykomplex = "https://npf-stroycomplex.ru/" + obj.GetAttribute("href");
            }
            //Пенсионные правила Фонда
            string penspravstroykomplex = "";
            foreach (IDomObject obj in dom.Find(".tab-pane:eq(12) a:eq(13)"))
            {
                penspravstroykomplex = "https://npf-stroycomplex.ru/" + obj.GetAttribute("href");
            }
            //Страховые правила Фонда 
            string strahpravstroykomplex = "";
            foreach (IDomObject obj in dom.Find(".tab-pane:eq(12) a:eq(14)"))
            {
                strahpravstroykomplex = "https://npf-stroycomplex.ru/" + obj.GetAttribute("href");
            }
            //Устав Фонда
            string ustavstroykomplex = "";
            foreach (IDomObject obj in dom.Find(".tab-pane:eq(12) a:eq(15)"))
            {
                ustavstroykomplex = "https://npf-stroycomplex.ru/" + obj.GetAttribute("href");
            }
            //Сведения об органах управления Фонда 
            string orgstroykomplex = "";
            foreach (IDomObject obj in dom.Find(".tab-pane:eq(12) a:eq(16)"))
            {
                orgstroykomplex = "https://npf-stroycomplex.ru/" + obj.GetAttribute("href");
            }
            //Сведения об опыте работы в кредитных организациях и некредитных финансовых организациях органов управления фонда, членах совета директоров фонда, должностных лицах и работниках Фонда
            string opytstroykomplex = "";
            foreach (IDomObject obj in dom.Find(".tab-pane:eq(12) a:eq(22)"))
            {
                opytstroykomplex = "https://npf-stroycomplex.ru/" + obj.GetAttribute("href");
            }
            //Органы управления
            string orguprstroykomplex = "";
            orguprstroykomplex = orgstroykomplex + " " + opytstroykomplex;
            //Средневзвешенный процент
            string srvzvprocstroykomplex = "";
            foreach (IDomObject obj in dom.Find(".tab-pane:eq(12) a:eq(25)"))
            {
                srvzvprocstroykomplex = "https://npf-stroycomplex.ru/" + obj.GetAttribute("href");
            }
            //Структура средств пенсионных резервов
            string structuraPRstroykomplex = "";
            foreach (IDomObject obj in dom.Find(".tab-pane:eq(12) a:eq(26)"))
            {
                structuraPRstroykomplex = "https://npf-stroycomplex.ru/" + obj.GetAttribute("href");
            }
            //Структура инвестиционного портфеля средств пенсионных накоплений
            string structuraPNstroykomplex = "";
            foreach (IDomObject obj in dom.Find(".tab-pane:eq(12) a:eq(27)"))
            {
                structuraPNstroykomplex = "https://npf-stroycomplex.ru/" + obj.GetAttribute("href");
            }
            string structuraPRPNstroykomplex = "";
            structuraPRPNstroykomplex = structuraPRstroykomplex + " " + structuraPNstroykomplex;
            //Состав инвестиционного портфеля
            string sostavportfstroykomplex = "";
            foreach (IDomObject obj in dom.Find(".tab-pane:eq(12) a:eq(28)"))
            {
                sostavportfstroykomplex = "https://npf-stroycomplex.ru/" + obj.GetAttribute("href");
            }
            //информация о процессе инвестирования
            string procinvestsroykomplex = "";
            foreach (IDomObject obj in dom.Find(".tab-pane:eq(12) a:eq(29)"))
            {
                procinvestsroykomplex = "https://npf-stroycomplex.ru/" + obj.GetAttribute("href");
            }
            //Иная информация
            string sobytstroykomplex = "";
            foreach (IDomObject obj in dom.Find(".tab-pane:eq(12) td:eq(100)"))
            {
                sobytstroykomplex = obj.TextContent;
            }
            //Запреты
            string zaprstroykomplex = "";
            foreach (IDomObject obj in dom.Find(".tab-pane:eq(12) td:eq(103)"))
            {
                zaprstroykomplex = obj.TextContent;
            }


            //Раздел с раскрытием на главной странице
            string raskrstroycomplex = "";
                url = "http://www.npf-stroycomplex.ru/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".col-md-8 a:eq(2)"))
                {
                    raskrstroycomplex = "http://www.npf-stroycomplex.ru/" + obj.GetAttribute("href");
                }



     
            
            //НПФ Достойное будущее
            displayChange("НПФ Достойное будущее");
        
                //раздел с раскрытием на главной странице
                string raskrdfnpf = "";
                url = "https://www.dfnpf.ru/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".uk-subnav a:eq(1) "))
                {
                    raskrdfnpf = "https://www.dfnpf.ru/" + obj.GetAttribute("href");
                }
                //УК
                string UKdfnpf = "";
                foreach (IDomObject obj in dom.Find(".uk-nav.uk-nav-default a:eq(17) "))
                {
                    UKdfnpf = "https://www.dfnpf.ru/" + obj.GetAttribute("href");
                }
                //СД
                string SDdfnpf = "";
                foreach (IDomObject obj in dom.Find(".uk-nav.uk-nav-default a:eq(18) "))
                {
                    SDdfnpf = "https://www.dfnpf.ru/" + obj.GetAttribute("href");
                }
            //Информация о процессе инвестирования
            string procinvestdfnpf = "";
            foreach (IDomObject obj in dom.Find(".uk-nav.uk-nav-default a:eq(19) "))
            {
                procinvestdfnpf = "https://www.dfnpf.ru/" + obj.GetAttribute("href");
            }
            //лицензия
            string teldfnpf = "";
                url = "https://www.dfnpf.ru/disclosure";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".uk-margin a:eq(0) "))
                {
                    teldfnpf = obj.GetAttribute("href");
                }
                //Список акционеров фонда и лиц, под контролем либо значительным влиянием которых находится фонд
                string spisokakciondfnpf = "";
                foreach (IDomObject obj in dom.Find(".uk-margin a:eq(2) "))
                {
                    spisokakciondfnpf = "https://www.dfnpf.ru/" + obj.GetAttribute("href");
                }
                //лицензия
                string licenzdfnpf = "";
                foreach (IDomObject obj in dom.Find(".el-item a:eq(4) "))
                {
                    licenzdfnpf = "https://www.dfnpf.ru/" + obj.GetAttribute("href");
                }
                //устав
                string ustavdfnpf = "";
                foreach (IDomObject obj in dom.Find(".el-item a:eq(6) "))
                {
                    ustavdfnpf = "https://www.dfnpf.ru/" + obj.GetAttribute("href");
                }
                //Сведения об органах управления фондом
                string orguprdfnpf = "";
                foreach (IDomObject obj in dom.Find(".el-item a:eq(10) "))
                {
                    orguprdfnpf = "https://www.dfnpf.ru/" + obj.GetAttribute("href");
                }
                //Уведомление о начале процедуры реорганизации АО «НПФ «Достойное БУДУЩЕЕ» в форме присоединения к АО «НПФ Эволюция»
                string reorgdfnpf1 = "";
                foreach (IDomObject obj in dom.Find(".el-item a:eq(17) "))
                {
                    reorgdfnpf1 = "https://www.dfnpf.ru/" + obj.GetAttribute("href");
                }
                string reorgdfnpf2 = "";
                foreach (IDomObject obj in dom.Find(".el-item a:eq(20) "))
                {
                    reorgdfnpf2 = "https://www.dfnpf.ru/" + obj.GetAttribute("href");
                }

                //в одну ячейку
                string reorgdfnpf = "";
                reorgdfnpf = reorgdfnpf1 + " " + reorgdfnpf2;

                //Информация о конечном владельце
                string vladelecdfnpf = "";
                foreach (IDomObject obj in dom.Find(".el-item a:eq(14) "))
                {
                    vladelecdfnpf = "https://www.dfnpf.ru/" + obj.GetAttribute("href");
                }
                //О решении о приостановлении привлечения новых застрахованных лиц по обязательному пенсионному страхованию и запрете на проведение АО НПФ «САФМАР» всех или части операций
                string zaprdfnpf = "";
                foreach (IDomObject obj in dom.Find(".el-item a:eq(22) "))
                {
                    zaprdfnpf = "https://www.dfnpf.ru/" + obj.GetAttribute("href");
                }
                //пенсионные правила
                string penspravfdnpf = "";
                foreach (IDomObject obj in dom.Find(".el-item a:eq(104) "))
                {
                    penspravfdnpf = "https://www.dfnpf.ru/" + obj.GetAttribute("href");
                }
                //страховые правила
                string strahpravdfnpf = "";
                foreach (IDomObject obj in dom.Find(".el-item a:eq(105) "))
                {
                    strahpravdfnpf = "https://www.dfnpf.ru/" + obj.GetAttribute("href");
                }
                //Отчёт о формировании средств пенсионных накоплений 
                string formirpndfnpf = "";
                foreach (IDomObject obj in dom.Find(".el-item a:eq(217) "))
                {
                    formirpndfnpf = "https://www.dfnpf.ru/" + obj.GetAttribute("href");
                }
                //бухгалтерская отчетность
                string buhdfnpf = "";
                foreach (IDomObject obj in dom.Find(".el-item a:eq(177) "))
                {
                    buhdfnpf = "https://www.dfnpf.ru/" + obj.GetAttribute("href");
                }
                //Аудиторское заключение 
                string auditdfnpf = "";
                foreach (IDomObject obj in dom.Find(".el-item a:eq(226) "))
                {
                    auditdfnpf = "https://www.dfnpf.ru/" + obj.GetAttribute("href");
                }
                //Актуарное заключение
                string actuardfnpf = "";
                foreach (IDomObject obj in dom.Find(".el-item a:eq(242) "))
                {
                    actuardfnpf = "https://www.dfnpf.ru/" + obj.GetAttribute("href");
                }
                //Отчётность
                string otchetdfnpf = "";
                otchetdfnpf = "Бухгалтерская отчетность:" + " " + buhdfnpf + " " + "Аудиторское заключение:" + " " + auditdfnpf + " " + "Актуарное заключение:" + " " + actuardfnpf;
                //Финансовая отчетность по МСФО и отчет независимого аудитора
                string MSFOdfnpf = "";
                foreach (IDomObject obj in dom.Find(".el-item a:eq(256) "))
                {
                    MSFOdfnpf = "https://www.dfnpf.ru/" + obj.GetAttribute("href");
                }

                //адрес
                
            string adressdfnpf = "";
            url = "https://www.dfnpf.ru/disclosure";
            htmlCode = downloadPageUTF8(url);

            dom = CQ.CreateDocument(htmlCode);
            foreach (IDomObject obj in dom.Find(".el-content.uk-panel.uk-margin-top li:eq(3)"))
            {
                adressdfnpf = obj.TextContent;
            }


            //МНПФ АКВИЛОН
            displayChange("МНПФ АКВИЛОН");
         
                //раздел на главной странице
                string raskrakvilon = "";
                url = "https://mnpf-akvilon.ru/content/mainakv/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".tabs a:eq(1)"))
                {
                    raskrakvilon = "https://mnpf-akvilon.ru/content/" + obj.GetAttribute("href");
                }
                //телефон
                string telakvilon = "";
                foreach (IDomObject obj in dom.Find(".phone "))
                {
                    telakvilon = obj.TextContent;
                }
                //адрес
                string adressakvilon = "";
                url = "https://mnpf-akvilon.ru/content/contacti";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".mapHint .desc"))
                {

                    adressakvilon = obj.TextContent.Split(':')[1];
                    adressakvilon = adressakvilon.Replace("Тел.,факс", "");

                }


                //Список акционеров фонда и лиц, под контролем либо значительным влиянием которых находится Фонд 
                string spisokakcionakvilon = "";
                url = "https://mnpf-akvilon.ru/content/struktura_upravlenija";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".cont a:eq(0) "))
                {
                    spisokakcionakvilon = obj.GetAttribute("href");
                }
                //УК
                string UKakvilon = "";
                url = "https://mnpf-akvilon.ru/content/raskrytie_informacii";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".hierarchical-expanded a:eq(3) "))
                {
                    UKakvilon = "https://mnpf-akvilon.ru/content/" + obj.GetAttribute("href");
                }
                //СД
                string SDakvilon = "";
                foreach (IDomObject obj in dom.Find(".hierarchical-expanded a:eq(4) "))
                {
                    SDakvilon = "https://mnpf-akvilon.ru/content/" + obj.GetAttribute("href");
                }
                //официальные документы
                string ofdocakvilon = "";
                foreach (IDomObject obj in dom.Find(".hierarchical-expanded a:eq(2) "))
                {
                    ofdocakvilon = "https://mnpf-akvilon.ru/content/" + obj.GetAttribute("href");
                }
                //Доходность
                string dohodakvilon = "";
                foreach (IDomObject obj in dom.Find(".hierarchical-expanded a:eq(5) "))
                {
                    dohodakvilon = "https://mnpf-akvilon.ru/content/" + obj.GetAttribute("href");
                }
                //Отчётность
                string otchetakvilon = "";
                foreach (IDomObject obj in dom.Find(".hierarchical-expanded a:eq(10) "))
                {
                    otchetakvilon = "https://mnpf-akvilon.ru/content/" + obj.GetAttribute("href");
                }



       

            //АПК ФОНД
            displayChange("АПК ФОНД");
         
                //раздел с раскрытием на главной странице
                string raskrapkfond = "";
                url = "http://www.apk-fond.ru/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);

                foreach (IDomObject obj in dom.Find("#menu div:eq(3) a"))
                {
                    raskrapkfond = "http://www.apk-fond.ru/" + obj.GetAttribute("href");
                }
                //отчетность
                string otchetapkfond = "";
                foreach (IDomObject obj in dom.Find("#left a:eq(1)"))
                {
                    otchetapkfond = "http://www.apk-fond.ru/" + obj.GetAttribute("href");
                }
                //Информация о структуре и составе акционеров
                string structuraakcionapkfond = "";
                foreach (IDomObject obj in dom.Find("#left a:eq(10)"))
                {
                    structuraakcionapkfond = "http://www.apk-fond.ru/" + obj.GetAttribute("href");
                }
                //Показатели деятельности
                string pokazatelapkfond = "";
                foreach (IDomObject obj in dom.Find("#left a:eq(14)"))
                {
                    pokazatelapkfond = "http://www.apk-fond.ru/" + obj.GetAttribute("href");
                }
                //устав
                string ustavapkfond = "";
                url = "http://www.apk-fond.ru/raskinfor/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);

                foreach (IDomObject obj in dom.Find("#text a:eq(0)"))
                {
                    ustavapkfond = "http://www.apk-fond.ru/" + obj.GetAttribute("href");
                }
                //Пенсионные правила
                string penspravapkfond = "";
                foreach (IDomObject obj in dom.Find("#text a:eq(4)"))
                {
                    penspravapkfond = "http://www.apk-fond.ru/" + obj.GetAttribute("href");
                }
                //Информация о заключенных и прекращенных договорах с управляющей компанией и специализированным депозитарием 
                string UKSDapkfond = "";
                foreach (IDomObject obj in dom.Find("#text a:eq(24)"))
                {
                    UKSDapkfond = "http://www.apk-fond.ru/" + obj.GetAttribute("href");
                }

                //Результат размещения ПР
                string rezrazmPRapkfond = "";
                foreach (IDomObject obj in dom.Find("#text a:eq(35)"))
                {
                    rezrazmPRapkfond = "http://www.apk-fond.ru/" + obj.GetAttribute("href");
                }
                //Информация о составе средств пенсионных резервов АО «НПФ «АПК-Фонд»
                string sostavPRapkfond = "";
                foreach (IDomObject obj in dom.Find("#text a:eq(36)"))
                {
                    sostavPRapkfond = "http://www.apk-fond.ru/" + obj.GetAttribute("href");
                }
                //Структура средств пенсионных резервов
                string structuraPRapkfond = "";
                foreach (IDomObject obj in dom.Find("#text a:eq(38)"))
                {
                    structuraPRapkfond = "http://www.apk-fond.ru/" + obj.GetAttribute("href");
                }
            //информация о процессе инвестирования
            string procinvestapkfond = "";
            foreach (IDomObject obj in dom.Find("#text a:eq(39)"))
            {
                procinvestapkfond = "http://www.apk-fond.ru/" + obj.GetAttribute("href");
            }

            //лицензия
            string licenzapkfond = "";
                url = "http://www.apk-fond.ru/about/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);

                foreach (IDomObject obj in dom.Find("#text a:eq(1)"))
                {
                    licenzapkfond = "http://www.apk-fond.ru/" + obj.GetAttribute("href");
                }
                //Контакты
                string contactapkfond = "";
                url = "http://www.apk-fond.ru/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find("#left a:eq(16)"))
                {
                    contactapkfond = "http://www.apk-fond.ru/" + obj.GetAttribute("href");
                }

        
            //ХАНТЫ-МАНСИЙСКИЙ НПФ
            displayChange("ХАНТЫ-МАНСИЙСКИЙ НПФ");
       
                //раздел с раскрытием на главной странице
                string raskrhmnpf = "";
                url = "https://www.hmnpf.ru/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);

                foreach (IDomObject obj in dom.Find(".expanded  a:eq(4)"))
                {
                    raskrhmnpf = "https://www.hmnpf.ru/" + obj.GetAttribute("href");
                }
                //лицензия
                string licenzhmnpf = "";
                url = "https://www.hmnpf.ru/about/disclosures/obshaya-informaciya/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);

                foreach (IDomObject obj in dom.Find(".page.clear a:eq(3)"))
                {
                    licenzhmnpf = obj.GetAttribute("href");
                }
                //устав
                string ustavhmnpf = "";
                foreach (IDomObject obj in dom.Find(".page.clear a:eq(4)"))
                {
                    ustavhmnpf = obj.GetAttribute("href");
                }
                //Информация о структуре и составе акционеров Фонда, в том числе о лицах, под контролем либо значительным влиянием которых находится Фонд
                string structuraakcionhmnpf = "";
                foreach (IDomObject obj in dom.Find(".page.clear a:eq(6)"))
                {
                    structuraakcionhmnpf = obj.GetAttribute("href");
                }
                //отчетность
                string otchethmnpf = "";
                foreach (IDomObject obj in dom.Find(".current a:eq(2)"))
                {
                    otchethmnpf = "https://www.hmnpf.ru/" + obj.GetAttribute("href");
                }
                //Руководство фонда
                string rucovodhmnpf = "";
                foreach (IDomObject obj in dom.Find(".current a:eq(3)"))
                {
                    rucovodhmnpf = "https://www.hmnpf.ru/" + obj.GetAttribute("href");
                }
                //Документы
                string dochmnpf = "";
                foreach (IDomObject obj in dom.Find(".current a:eq(4)"))
                {
                    dochmnpf = "https://www.hmnpf.ru/" + obj.GetAttribute("href");
                }
                //Показатели деятельности
                string pokazatelhmnpf = "";
                foreach (IDomObject obj in dom.Find(".current a:eq(5)"))
                {
                    pokazatelhmnpf = "https://www.hmnpf.ru/" + obj.GetAttribute("href");
                }
                //События оказывающие существенное влияние на стоимость активов фонда
                string sobythmnpf = "";
                foreach (IDomObject obj in dom.Find(".current a:eq(7)"))
                {
                    sobythmnpf = "https://www.hmnpf.ru/" + obj.GetAttribute("href");
                }
                
                //Структура и состав инвестиционного портфеля
                string investporthmnpf = "";
                foreach (IDomObject obj in dom.Find(".current a:eq(9)"))
                {
                    investporthmnpf = "https://www.hmnpf.ru/" + obj.GetAttribute("href");
                }
                //СД
                string SDUKhmnpf = "";
                url = "https://www.hmnpf.ru/about/disclosures/investitsionnaya-deyatelnost/archive_invest/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);

                foreach (IDomObject obj in dom.Find(".menu a:eq(13)"))
                {
                    SDUKhmnpf = "https://www.hmnpf.ru/" + obj.GetAttribute("href");
                }

            //Сведения о принятии Банком России решения о запрете на проведение фондом всех или части операций
            string zaprhmnpf = "";
            url = "https://www.hmnpf.ru/about/disclosures/ban_on_operations/";
            htmlCode = downloadPage(url);

            dom = CQ.CreateDocument(htmlCode);
            foreach (IDomObject obj in dom.Find(".p_title strong"))
            {
                zaprhmnpf = obj.TextContent;
            }
            
            


            //НПФ ОПФ
            displayChange("НПФ ОПФ");
         
                //раздел с раскрытием на главной странице
                string raskropf = "";
                url = "http://www.npfopf.ru/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);

                foreach (IDomObject obj in dom.Find(".additional-buttons a:eq(0)"))
                {
                    raskropf = "http://www.npfopf.ru/" + obj.GetAttribute("href");
                }
                //Органы управления и контроля, структура и состав акционеров
                string orgupropf = "";
                url = "http://www.npfopf.ru/?issue_id=88";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);

                foreach (IDomObject obj in dom.Find(".vmenu2_text a:eq(1)"))
                {
                    orgupropf = "http://www.npfopf.ru/" + obj.GetAttribute("href");
                }
                //УК
                string UKopf = "";
                foreach (IDomObject obj in dom.Find(".vmenu2_text a:eq(2)"))
                {
                    UKopf = "http://www.npfopf.ru/" + obj.GetAttribute("href");
                }
                //SD
                string SDopf = "";
                foreach (IDomObject obj in dom.Find(".vmenu2_text a:eq(3)"))
                {
                    SDopf = "http://www.npfopf.ru/" + obj.GetAttribute("href");
                }
                //филиалы
                string adressopf = "";
                foreach (IDomObject obj in dom.Find(".vmenu2_text a:eq(4)"))
                {
                    adressopf = "http://www.npfopf.ru/" + obj.GetAttribute("href");
                }
                //Устав
                string ustavopf = "";
                foreach (IDomObject obj in dom.Find(".vmenu2_text a:eq(5)"))
                {
                    ustavopf = "http://www.npfopf.ru/" + obj.GetAttribute("href");
                }
                //Страховые правила
                string strahpravopf = "";
                foreach (IDomObject obj in dom.Find(".vmenu2_text a:eq(6)"))
                {
                    strahpravopf = "http://www.npfopf.ru/" + obj.GetAttribute("href");
                }
                //Пенсионные правила
                string penspravopf = "";
                foreach (IDomObject obj in dom.Find(".vmenu2_text a:eq(7)"))
                {
                    penspravopf = "http://www.npfopf.ru/" + obj.GetAttribute("href");
                }
                //Акт.аудит. отчетности
                string actaudotchetopf = "";
                foreach (IDomObject obj in dom.Find(".vmenu2_text a:eq(8)"))
                {
                    actaudotchetopf = "http://www.npfopf.ru/" + obj.GetAttribute("href");
                }
                //Бух.отчет
                string buhotchetopf = "";
                foreach (IDomObject obj in dom.Find(".vmenu2_text a:eq(9)"))
                {
                    buhotchetopf = "http://www.npfopf.ru/" + obj.GetAttribute("href");
                }
                //Отчет
                string otchetopf = "Актуарный и аудиторский заключения:" + " " + actaudotchetopf + " " + "Бухгалтерский отчет:" + " " + buhotchetopf;
                //Основные показатели деятельности
                string pokazatelopf = "";
                foreach (IDomObject obj in dom.Find(".vmenu2_text a:eq(10)"))
                {
                    pokazatelopf = "http://www.npfopf.ru/" + obj.GetAttribute("href");
                }
                //Отчет о формировании средств ПН
                string formirPNopf = "";
                foreach (IDomObject obj in dom.Find(".vmenu2_text a:eq(12)"))
                {
                    formirPNopf = "http://www.npfopf.ru/" + obj.GetAttribute("href");
                }
                //Средневзвешенный процент
                string srvzvprocopf = "";
                foreach (IDomObject obj in dom.Find(".vmenu2_text a:eq(11)"))
                {
                    srvzvprocopf = "http://www.npfopf.ru/" + obj.GetAttribute("href");
                }
                //лицензия
                string licenzopf = "";
                foreach (IDomObject obj in dom.Find(".content  a:eq(1)"))
                {
                    licenzopf = "http://www.npfopf.ru/" + obj.GetAttribute("href");
                }
                //Руководство
                string rucovodopf = "";
                url = "http://www.npfopf.ru/?issue_id=145";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);

                foreach (IDomObject obj in dom.Find(".vmenu2_text a:eq(1)"))
                {
                    rucovodopf = "http://www.npfopf.ru/" + obj.GetAttribute("href");
                }
                //телефон
                string telopf = "";
                foreach (IDomObject obj in dom.Find(".vmenu2_text a:eq(2)"))
                {
                    telopf = "http://www.npfopf.ru/" + obj.GetAttribute("href");
                }
                //Структура и состав портфеля 
                string portfopf = "";
                foreach (IDomObject obj in dom.Find(".vmenu2_text a:eq(3)"))
                {
                    portfopf = "http://www.npfopf.ru/" + obj.GetAttribute("href");
                }
            //Информация о процессе инвестирования средств пенсионных накоплений и размещения средств пенсионных резервов 
            string procinvestopf = "";
            url = "http://www.npfopf.ru/?issue_id=126";
            htmlCode = downloadPageUTF8(url);

            dom = CQ.CreateDocument(htmlCode);

            foreach (IDomObject obj in dom.Find(".content a:eq(7)"))
            {
                procinvestopf = "http://www.npfopf.ru/" + obj.GetAttribute("href");
            }





            //НПФ ГЕФЕСТ
            displayChange("НПФ ГЕФЕСТ");
       
                //устав
                string ustavgefest = "";
                url = "https://npfgefest.ru/info/cbr/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);

                foreach (IDomObject obj in dom.Find(".tab_name a:eq(0)"))
                {
                    ustavgefest = "https://npfgefest.ru/" + obj.GetAttribute("href");
                }
                //Сведения об органах управления, членах совета директоров (наблюдательного совета) фонда, должностных лицах и работниках фонда
                string orguprgefest = "";
                foreach (IDomObject obj in dom.Find(".tab_name a:eq(1)"))
                {
                    orguprgefest = "https://npfgefest.ru/" + obj.GetAttribute("href");
                }
                //телефон
                string telgefest = "";
                foreach (IDomObject obj in dom.Find(".tab_name a:eq(2)"))
                {
                    telgefest = "https://npfgefest.ru/" + obj.GetAttribute("href");
                }
                //Результаты инвестирования средств пенсионных накоплений
                string rezinvestPNgefest = "";
                foreach (IDomObject obj in dom.Find(".tab_name a:eq(3)"))
                {
                    rezinvestPNgefest = "https://npfgefest.ru/" + obj.GetAttribute("href");
                }
                //Результаты размещения средств пенсионных резервов
                string rezinvestPRgefest = "";
                foreach (IDomObject obj in dom.Find(".tab_name a:eq(5)"))
                {
                    rezinvestPRgefest = "https://npfgefest.ru/" + obj.GetAttribute("href");
                }
                //Средневзвешенный процент
                string srvzvprocgefest = "";
                foreach (IDomObject obj in dom.Find(".tab_name a:eq(6)"))
                {
                    srvzvprocgefest = "https://npfgefest.ru/" + obj.GetAttribute("href");
                }
                //Структура инвестиционного портфеля фонда
                string structportfgefest = "";
                foreach (IDomObject obj in dom.Find(".tab_name a:eq(7)"))
                {
                    structportfgefest = "https://npfgefest.ru/" + obj.GetAttribute("href");
                }
            //Информация о процессе инвестирования
            string procinvestgefest = "";
            foreach (IDomObject obj in dom.Find(".tab_name a:eq(9)"))
            {
                procinvestgefest = "https://npfgefest.ru/" + obj.GetAttribute("href");
            }

            //Иная информация о событиях
            string sobytgefest = "";
                foreach (IDomObject obj in dom.Find(".tab_name a:eq(10)"))
                {
                    sobytgefest = "https://npfgefest.ru/" + obj.GetAttribute("href");
                }
                //Сведения о принятии Банком России решения о запрете
                string zaprgefest = "";
                foreach (IDomObject obj in dom.Find(".tab_name a:eq(11)"))
                {
                    zaprgefest = "https://npfgefest.ru/" + obj.GetAttribute("href");
                }
                //УК и СД 
                string UKSDgefest = "";
                url = "https://npfgefest.ru/info/investments/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);

                foreach (IDomObject obj in dom.Find(".index-news-block-big.anons-index-news a:eq(0)"))
                {
                    UKSDgefest = "https://npfgefest.ru/" + obj.GetAttribute("href");
                }
            //инвестиционный портфель
            string investportfgefest = "";
            foreach (IDomObject obj in dom.Find(".index-news-block-big.anons-index-news a:eq(1)"))
            {
                investportfgefest = "https://npfgefest.ru/" + obj.GetAttribute("href");
            }
            //Пенсионные и страховые правила
            string pensstrahgefest = "";
                foreach (IDomObject obj in dom.Find(".link a:eq(24)"))
                {
                    pensstrahgefest = "https://npfgefest.ru/" + obj.GetAttribute("href");
                }
                //Отчетность
                string otchetgefest = "";
                foreach (IDomObject obj in dom.Find(".link a:eq(25)"))
                {
                    otchetgefest = "https://npfgefest.ru/" + obj.GetAttribute("href");
                }
                //Показатели деятельности
                string pokazatelgefest = "";
                foreach (IDomObject obj in dom.Find(".link a:eq(26)"))
                {
                    pokazatelgefest = "https://npfgefest.ru/" + obj.GetAttribute("href");
                }
            //Расскрытие информации
            string raskrgefest = "";
            foreach (IDomObject obj in dom.Find(".link a:eq(28)"))
            {
                raskrgefest = "https://npfgefest.ru/" + obj.GetAttribute("href");
            }
            //Лицензия
            string licenzgefest = "";
                url = "https://npfgefest.ru/company/history/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);

                foreach (IDomObject obj in dom.Find(".row a:eq(6)"))
                {
                    licenzgefest = "https://npfgefest.ru/" + obj.GetAttribute("href");
                }
                //адрес
                string adressgefest = "";
                url = "https://npfgefest.ru/contacts/";
                htmlCode = downloadPage(url);

                dom = CQ.CreateDocument(htmlCode);

                foreach (IDomObject obj in dom.Find(".row p:eq(0)"))
                {
                    adressgefest = obj.TextContent + "\n" + "Контактная информация, касающуяся филиала или офиса НПФ «Гефест» " + telgefest;
                }

         

            //УГМК ПЕРСПЕКТИВА
            displayChange("УГМК ПЕРСПЕКТИВА");
          

                //раздел с раскрытием на главной странице
                string raskrnpfond = "";
                url = "http://www.npfond.ru/";
                htmlCode = downloadPage(url);

                dom = CQ.CreateDocument(htmlCode);

                foreach (IDomObject obj in dom.Find(".custom a:eq(0)"))
                {
                    raskrnpfond = "http://www.npfond.ru/" + obj.GetAttribute("href");
                }
                //Информация о структуре и составе акционеров фонда
                string sostavakcionnpfond = "";
                foreach (IDomObject obj in dom.Find(".custom a:eq(1)"))
                {
                    sostavakcionnpfond = obj.GetAttribute("href");
                }
                //лицензия
                string licenznpfond = "";
                url = "http://www.npfond.ru/index.php?option=com_content&view=article&id=381:informatsiya-raskrytivaemaya-po-5175&catid=88&Itemid=528";
                htmlCode = downloadPage(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".item-page a:eq(1)"))
                {
                    licenznpfond = "http://www.npfond.ru/" + obj.GetAttribute("href");
                }
                // Бухгалтерская (финансовая) отчетность Фонда, аудиторское и актуарное заключения
                string buhotchetnpfond = "";
                foreach (IDomObject obj in dom.Find(".item-page a:eq(4)"))
                {
                    buhotchetnpfond = "http://www.npfond.ru/" + obj.GetAttribute("href");
                }
                //Показатели
                string pokazatelnpfond = "";
                foreach (IDomObject obj in dom.Find(".item-page a:eq(6)"))
                {
                    pokazatelnpfond = "http://www.npfond.ru/" + obj.GetAttribute("href");
                }
                //О заключении и прекращении действия договора доверительного управления пенсионными резервами или пенсионными накоплениями с управляющей компанией          
                string UKnpfond = "";
                foreach (IDomObject obj in dom.Find(".item-page a:eq(7)"))
                {
                    UKnpfond = "http://www.npfond.ru/" + obj.GetAttribute("href");
                }
                //О заключении и прекращении договора со специализированным депозитарием
                string SDnpfond = "";
                foreach (IDomObject obj in dom.Find(".item-page a:eq(8)"))
                {
                    SDnpfond = "http://www.npfond.ru/" + obj.GetAttribute("href");
                }
                //Пенсионные правила, страховые правила, а также внесенные в них изменения и дополнения
                string pravilanpfond = "";
                foreach (IDomObject obj in dom.Find(".item-page a:eq(9)"))
                {
                    pravilanpfond = "http://www.npfond.ru/" + obj.GetAttribute("href");
                }
                //Информация о регистрации Банком России изменений и дополнений в пенсионные правила и в страховые правила Фонда             
                string izmpravnpfond = "";
                foreach (IDomObject obj in dom.Find(".item-page a:eq(10)"))
                {
                    izmpravnpfond = "http://www.npfond.ru/" + obj.GetAttribute("href");
                }
                //устав
                string ustavnpfond = "";
                foreach (IDomObject obj in dom.Find(".item-page a:eq(11)"))
                {
                    ustavnpfond = "http://www.npfond.ru/" + obj.GetAttribute("href");
                }
                //Сведения об органах управления
                string orguprnpfond = "";
                foreach (IDomObject obj in dom.Find(".item-page a:eq(12)"))
                {
                    orguprnpfond = "http://www.npfond.ru/" + obj.GetAttribute("href");
                }
                //Структура  и состав инвестиционного портфеля 
                string investportfnpfond = "";
                foreach (IDomObject obj in dom.Find(".item-page a:eq(13)"))
                {
                    investportfnpfond = "http://www.npfond.ru/" + obj.GetAttribute("href");
                }
                //Решения о запрете
                string zaprnpfond = "";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".item-page span:eq(43)"))
                {
                    zaprnpfond = obj.TextContent;
                }
                //Информация о событиях
                string sobytnpfond = "";
                foreach (IDomObject obj in dom.Find(".item-page span:eq(45)"))
                {
                    sobytnpfond = obj.TextContent;
                }
                //телефон
                string telnpfond = "";
                foreach (IDomObject obj in dom.Find(".item-page span:eq(16)"))
                {
                    telnpfond = obj.TextContent;
                }
                //адрес
                string adressnpfondekb = "";
                foreach (IDomObject obj in dom.Find(".item-page span:eq(14)"))
                {
                    adressnpfondekb = obj.TextContent;
                }
            //адреса филиалов
            string adressnpfondpodrazdel = "";
            url = "http://www.npfond.ru/index.php?option=com_content&view=article&id=47&Itemid=515";
            htmlCode = downloadPage(url);

            dom = CQ.CreateDocument(htmlCode);

            foreach (IDomObject obj in dom.Find(".MsoNormal a:eq(1)"))
            {
                adressnpfondpodrazdel = "http://www.npfond.ru/" + obj.GetAttribute("href");
            }
            string adressnpfond = "";
            adressnpfond = "Головной офис: " + adressnpfondekb + "\n" + "Обособленные подраздения: " + adressnpfondpodrazdel;





            //НПФ ФЕДЕРАЦИЯ
            displayChange("НПФ ФЕДЕРАЦИЯ");
          
                //раздел с раскрытием на главной странице
                string raskrfederation = "";
                url = "http://federation-npf.ru/";
                htmlCode = downloadPage(url);

                dom = CQ.CreateDocument(htmlCode);

                foreach (IDomObject obj in dom.Find(".menu a:eq(22)"))
                {
                    raskrfederation = "http://federation-npf.ru/" + obj.GetAttribute("href");
                }
                //Структура и состав инвестиционного портфеля
                string portffederation = "";
                foreach (IDomObject obj in dom.Find(".menu a:eq(23)"))
                {
                    portffederation = "http://federation-npf.ru/" + obj.GetAttribute("href");
                }
            //Информация о процессе инвестирования
            string procinvestfederation = "";
            foreach (IDomObject obj in dom.Find(".menu a:eq(24)"))
            {
                procinvestfederation = "http://federation-npf.ru/" + obj.GetAttribute("href");
            }
            //ОПС
            string OPSfederation = "";
                foreach (IDomObject obj in dom.Find(".menu a:eq(25)"))
                {
                    OPSfederation = "http://federation-npf.ru/" + obj.GetAttribute("href");
                }
                //Иная информация о событиях
                string sobytfederation = "";
                foreach (IDomObject obj in dom.Find(".menu a:eq(26)"))
                {
                    sobytfederation = "http://federation-npf.ru/" + obj.GetAttribute("href");
                }
                //Отчетность
                string otchetfederation = "";
                foreach (IDomObject obj in dom.Find(".menu a:eq(27)"))
                {
                    otchetfederation = "http://federation-npf.ru/" + obj.GetAttribute("href");
                }
                //Документы фонда
                string docfederation = "";
                foreach (IDomObject obj in dom.Find(".menu a:eq(28)"))
                {
                    docfederation = "http://federation-npf.ru/" + obj.GetAttribute("href");
                }
                //УК и СД
                string UKSDfederation = "";
                foreach (IDomObject obj in dom.Find(".menu a:eq(5)"))
                {
                    UKSDfederation = "http://federation-npf.ru/" + obj.GetAttribute("href");
                }
                //Органы управления
                string orguprfederation = "";
                foreach (IDomObject obj in dom.Find(".menu a:eq(2)"))
                {
                    orguprfederation = "http://federation-npf.ru/" + obj.GetAttribute("href");
                }
                //Акционеры 
                string akcionfederation = "";
                foreach (IDomObject obj in dom.Find(".menu a:eq(4)"))
                {
                    akcionfederation = "http://federation-npf.ru/" + obj.GetAttribute("href");
                }
                //телефон
                string telfederation = "";
                url = "http://federation-npf.ru/contact/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".block.phone p:eq(0)"))
                {
                    telfederation = obj.TextContent;
                }
                //адрес
                string adressfederationmsk = "";
                foreach (IDomObject obj in dom.Find(".block.address p:eq(0)"))
                {
                    adressfederationmsk = obj.TextContent;
                }
                //адрес филиала
                string adressfederationspb = "";
                foreach (IDomObject obj in dom.Find(".block.address p:eq(1)"))
                {
                    adressfederationspb = obj.TextContent;
                }

            string adressfederation = "";
            adressfederation = "Основной офис: " + adressfederationmsk + "\n" + "Уполномоченный представитель АО НПФ ФЕДЕРАЦИЯ по вопросам назначения пенсии – ООО «МОЯ ПЕНСИЯ»: " + adressfederationspb;

                //лицензия
            string licenzfederation = "";
                url = "http://federation-npf.ru/about/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".row a:eq(18)"))
                {
                    licenzfederation = obj.GetAttribute("href");
                }



            //НПФ БУДУЩЕЕ
            displayChange("НПФ БУДУЩЕЕ");
           //раздел с раскрытием на главной странице
                string raskrnpff = "";
                url = "https://npff.ru/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".h2-center a:eq(5)"))
                {
                    raskrnpff = "https://npff.ru/" + obj.GetAttribute("href");
                }
                //Инвестирование
                string investirnpff = "";
                foreach (IDomObject obj in dom.Find(".h2-center a:eq(6)"))
                {
                    investirnpff = "https://npff.ru/" + obj.GetAttribute("href");
                }
            //адрес 
            string adresnpff = "";
            foreach (IDomObject obj in dom.Find(".h2-center a:eq(1)"))
            {
                adresnpff = "https://npff.ru/" + obj.GetAttribute("href");
            }
            //телефон
            string telnpff = "";
            foreach (IDomObject obj in dom.Find(".h2-center a:eq(0)"))
            {
                telnpff = obj.GetAttribute("href");
            }
            //события
            string sobytnpff = "";
            url = "https://npff.ru/investment/";
            htmlCode = downloadPageUTF8(url);

            dom = CQ.CreateDocument(htmlCode);
            foreach (IDomObject obj in dom.Find("#others > div > p .file-link:eq(1)"))
            {
                sobytnpff = "https://npff.ru/" + obj.GetAttribute("href");
            }
            //Информация о процессе инвестирования
            string procinvestnpff = "";
            foreach (IDomObject obj in dom.Find("#others > div > p .file-link:eq(0)"))
            {
                procinvestnpff = "https://npff.ru/" + obj.GetAttribute("href");
            }





            //МНПФ БОЛЬШОЙ
            displayChange("МНПФ БОЛЬШОЙ");
        
                //раздел с раскрытием на главной странице
                string raskrbigpension = "";
                url = "https://www.bigpension.ru/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find("#h-nav a:eq(8)"))
                {
                    raskrbigpension = "https://www.bigpension.ru/" + obj.GetAttribute("href");
                }
            string telbigpension = "";
            foreach (IDomObject obj in dom.Find("#h-contacts p:eq(0)"))
            {
                telbigpension = obj.TextContent;
            }

            //структура и состав акционеров
            string structakcionbigpension = "";
            foreach (IDomObject obj in dom.Find(".fl a:eq(0)"))
            {
                structakcionbigpension = "https://www.bigpension.ru/" + obj.GetAttribute("href");
            }
            //акционеры фонда
            string akcionbigpension = "";
            foreach (IDomObject obj in dom.Find(".fl a:eq(1)"))
            {
                akcionbigpension = "https://www.bigpension.ru/" + obj.GetAttribute("href");
            }

            string acionbigpension = "";
            acionbigpension = structakcionbigpension + "\n" + akcionbigpension;

            //Документы
            string docbigpension = "";
                url = "https://www.bigpension.ru/docs/reporting/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find("#nav-faq a:eq(1)"))
                {
                    docbigpension = "https://www.bigpension.ru/" + obj.GetAttribute("href");
                }
            //Структура и состав акционеров 
            string sostavakcionbigpension = "";
            foreach (IDomObject obj in dom.Find(".nav-pager a:eq(4)"))
            {
                sostavakcionbigpension = "https://www.bigpension.ru/" + obj.GetAttribute("href");
            }
            //Результаты инвестирования
            string rezinvestbigpension = "";
            foreach (IDomObject obj in dom.Find(".nav-pager a:eq(5)"))
            {
                rezinvestbigpension = "https://www.bigpension.ru/" + obj.GetAttribute("href");
            }
            //Информация по основным показателям
            string pokazatelbigpension = "";
            url = "https://www.bigpension.ru/docs/reporting/";
            htmlCode = downloadPageUTF8(url);

            dom = CQ.CreateDocument(htmlCode);
            foreach (IDomObject obj in dom.Find(".c-box.c-box-2.c-box-14.cb a:eq(5)"))
            {
                pokazatelbigpension = "https://www.bigpension.ru/" + obj.GetAttribute("href");
            }
            //Отчетность
            string otchetbigpension = "";
            foreach (IDomObject obj in dom.Find(".c-box.c-box-2.c-box-14.cb a:eq(6)"))
            {
                otchetbigpension = "https://www.bigpension.ru/" + obj.GetAttribute("href");
            }
            //аудиторское
            string audbigpension = "";
            foreach (IDomObject obj in dom.Find(".c-box.c-box-2.c-box-14.cb a:eq(8)"))
            {
                audbigpension = "https://www.bigpension.ru/" + obj.GetAttribute("href");
            }
            //актуарное
            string actbigpension = "";
            foreach (IDomObject obj in dom.Find(".c-box.c-box-2.c-box-14.cb a:eq(9)"))
            {
                actbigpension = "https://www.bigpension.ru/" + obj.GetAttribute("href");
            }
            //Количество участников, вкладчиков и застрахованных лиц
            string vkladchbigpension = "";
            foreach (IDomObject obj in dom.Find(".c-box.c-box-2.c-box-14.cb a:eq(4)"))
            {
                vkladchbigpension = "https://www.bigpension.ru/" + obj.GetAttribute("href");
            }
            //Сведения об индексации
            string indecsbigpension = "";
            foreach (IDomObject obj in dom.Find(".nav-pager a:eq(2)"))
            {
                indecsbigpension = "https://www.bigpension.ru/" + obj.GetAttribute("href");
            }
            //Адрес
            string adresbigpension = "";
            url = "https://www.bigpension.ru/info-center/requisites/";
            htmlCode = downloadPageUTF8(url);

            dom = CQ.CreateDocument(htmlCode);
            foreach (IDomObject obj in dom.Find(".c-box.c-box-11.c-box-address.mb1 p:eq(5)"))
            {
                adresbigpension = obj.TextContent + "\n" + "Место нахождения обособленных подразделений фонда: " + raskrbigpension;
            }
            //лицензия 
            string licenzbigpension = "";
            url = "https://www.bigpension.ru/docs/ustav/";
            htmlCode = downloadPageUTF8(url);

            dom = CQ.CreateDocument(htmlCode);
            foreach (IDomObject obj in dom.Find(".c-box.c-box-2.c-box-14.cb a:eq(5)"))
            {
                licenzbigpension = "https://www.bigpension.ru/" + obj.GetAttribute("href");
            }
            //бухотчёт, ауд и акт заключения
            string buhotchetbigpension = "";
            buhotchetbigpension = "Бухгалтерская отчёность: " + otchetbigpension + "\n" + "Аудиторское заключение: " + audbigpension + "\n" + "Актуарное заключение: " + actbigpension;

            //Устав
            string ustavbigpension = "";
            foreach (IDomObject obj in dom.Find(".c-box.c-box-2.c-box-14.cb a:eq(2)"))
            {
                ustavbigpension = "https://www.bigpension.ru/" + obj.GetAttribute("href");
            }

            //Пенсионные правила 
            string penspravbigpension = "";
            foreach (IDomObject obj in dom.Find(".c-box.c-box-2.c-box-14.cb a:eq(3)"))
            {
                penspravbigpension = "https://www.bigpension.ru/" + obj.GetAttribute("href");
            }
            //Страховые правила
            string spravbigpension = "";
            foreach (IDomObject obj in dom.Find(".c-box.c-box-2.c-box-14.cb a:eq(4)"))
            {
                spravbigpension = "https://www.bigpension.ru/" + obj.GetAttribute("href");
            }

            //НПФ ВОЛГА-КАПИТАЛ
            displayChange("НПФ ВОЛГА-КАПИТАЛ");
          
                //раздел с раскрытием на главной странице
                string raskrvolgacapital = "";
                url = "https://www.volga-capital.ru/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find("#header a:eq(33)"))
                {
                    raskrvolgacapital = "https://www.volga-capital.ru/" + obj.GetAttribute("href");
                }
                //Страховые правила
                string strahpravvolgacapital = "";
                foreach (IDomObject obj in dom.Find("#header a:eq(14)"))
                {
                    strahpravvolgacapital = "https://www.volga-capital.ru/" + obj.GetAttribute("href");
                }
                //Пенсионные правила
                string penspravvolgacapital = "";
                foreach (IDomObject obj in dom.Find("#header a:eq(23)"))
                {
                    penspravvolgacapital = "https://www.volga-capital.ru/" + obj.GetAttribute("href");
                }
                //Документы фонда
                string docvolgacapital = "";
                foreach (IDomObject obj in dom.Find("#header a:eq(35)"))
                {
                    docvolgacapital = "https://www.volga-capital.ru/" + obj.GetAttribute("href");
                }
                //УК и СД
                string UKSDvolgacapital = "";
                foreach (IDomObject obj in dom.Find("#header a:eq(39)"))
                {
                    UKSDvolgacapital = "https://www.volga-capital.ru/" + obj.GetAttribute("href");
                }
                //Показатели деятельности
                string pokazatelvolgacapital = "";
                foreach (IDomObject obj in dom.Find("#header a:eq(41)"))
                {
                    pokazatelvolgacapital = "https://www.volga-capital.ru/" + obj.GetAttribute("href");
                }
                //Информация о структуре и составе акционеров
                string structuraakcionvolgacapital = "";
                foreach (IDomObject obj in dom.Find("#header a:eq(40)"))
                {
                    structuraakcionvolgacapital = "https://www.volga-capital.ru/" + obj.GetAttribute("href");
                }
                //Отчетность
                string otchetvolgacapital = "";
                foreach (IDomObject obj in dom.Find("#header a:eq(42)"))
                {
                    otchetvolgacapital = "https://www.volga-capital.ru/" + obj.GetAttribute("href");
                }
                //Сведения об органах управления, членах совета директоров, должностных лицах и работниках фонда
                string orguprvolgacapital = "";
                foreach (IDomObject obj in dom.Find("#header a:eq(43)"))
                {
                    orguprvolgacapital = "https://www.volga-capital.ru/" + obj.GetAttribute("href");
                }
                //о конечных владельцах фонда
                string konvladvolgacapital = "";
                foreach (IDomObject obj in dom.Find("#header a:eq(48)"))
                {
                    konvladvolgacapital = "https://www.volga-capital.ru/" + obj.GetAttribute("href");
                }
                //Сведения о месте нахождения Фонда и его обособленные подразделения
                string adressvolgacapital = "";
                foreach (IDomObject obj in dom.Find("#header a:eq(50)"))
                {
                    adressvolgacapital = "https://www.volga-capital.ru/" + obj.GetAttribute("href");
                }
                //Структура и состав инвестиционного портфеля
                string portfelvolgacapital = "";
                foreach (IDomObject obj in dom.Find("#header a:eq(51)"))
                {
                    portfelvolgacapital = "https://www.volga-capital.ru/" + obj.GetAttribute("href");
                }
                //телефон
                string telvolgacapital = "";
                htmlCode = downloadPage(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".menu-left a:eq(3)"))
                {
                    telvolgacapital = "https://www.volga-capital.ru/" + obj.GetAttribute("href");
                }
            //События
            string sobvolgacapital = "";
            url = "https://www.volga-capital.ru/disclosures/raskrytie-informatsii-v-sootvetstvii-s-ukazaniem-banka-rossii/";
            htmlCode = downloadPage(url);

            dom = CQ.CreateDocument(htmlCode);
            foreach (IDomObject obj in dom.Find("p > a"))
                {
                string findStr = "Иная информация о событиях (действиях), оказывающих, по мнению фонда, существенное влияние на совокупную стоимость активов, в которые инвестированы средства пенсионных накоплений и размещены средства пенсионных резервов";
                if (obj.TextContent.Equals(findStr))
                {
                    sobvolgacapital = "https://www.volga-capital.ru/" + obj.GetAttribute("href");
                }
                }
            //Информация о процессе инвестирования
            string procinvestvolgacapital = "";
            url = "https://www.volga-capital.ru/disclosures/raskrytie-informatsii-v-sootvetstvii-s-ukazaniem-banka-rossii/";
            htmlCode = downloadPage(url);

            dom = CQ.CreateDocument(htmlCode);
            foreach (IDomObject obj in dom.Find("p > a"))
            {
                string findStr = "Информация о процессе инвестирования средств пенсионных накоплений и размещения средств пенсионных резервов в объеме данных, предусмотренных абзацами вторым, третьим и пятым подпункта 3.1.6 пункта 3.1 Указания Банка России №4060-У";
                if (obj.TextContent.Equals(findStr))
                {
                    procinvestvolgacapital = "https://www.volga-capital.ru/" + obj.GetAttribute("href");
                }
            }






            //НПФ ГАЗФОНД
            displayChange("НПФ ГАЗФОНД");
         
                //раздел с раскрытием на главной странице
                string raskrgazfond = "";
                url = "https://gazfond.ru/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".header-menu-list a:eq(7)"))
                {
                    raskrgazfond = "https://gazfond.ru/" + obj.GetAttribute("href");
                }
                //лицензия
                string licenzgazfond = "";
                url = "https://gazfond.ru/fund/disclosure/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".row a:eq(5)"))
                {
                    licenzgazfond = "https://gazfond.ru/" + obj.GetAttribute("href");
                }
                //устав
                string ustavgazfond = "";
                foreach (IDomObject obj in dom.Find(".row a:eq(6)"))
                {
                    ustavgazfond = "https://gazfond.ru/" + obj.GetAttribute("href");
                }
                //Список акционеров
                string akcionergazfond = "";
                foreach (IDomObject obj in dom.Find(".row a:eq(10)"))
                {
                    akcionergazfond = "https://gazfond.ru/" + obj.GetAttribute("href");
                }
                //адреса
                string adressgazfond = "";
                foreach (IDomObject obj in dom.Find(".row a:eq(12)"))
                {
                    adressgazfond = "https://gazfond.ru/" + obj.GetAttribute("href");
                }
                //Органы управления
                string orguprgazfond = "";
                foreach (IDomObject obj in dom.Find(".row a:eq(13)"))
                {
                    orguprgazfond = "https://gazfond.ru/" + obj.GetAttribute("href");
                }
                //телефона 
                string telgazfond = "";
                foreach (IDomObject obj in dom.Find(".phone a:eq(0)"))
                {
                    telgazfond = obj.TextContent;
                }
            //информация о процессе инвестирования
            string procinvestgazfond = "";
            url = "https://gazfond.ru/fund/investment/";
            htmlCode = downloadPageUTF8(url);

            dom = CQ.CreateDocument(htmlCode);
            foreach (IDomObject obj in dom.Find(".warning-box.bg-lighten a:eq(0)"))
            {
                procinvestgazfond = "https://gazfond.ru/" + obj.GetAttribute("href");
            }


            //НПФ ПРОФЕССИОНАЛЬНЫЙ
            displayChange("НПФ ПРОФЕССИОНАЛЬНЫЙ");
         
                //раздел с раскрытием на главной странице
                string raskrprof = "";
                url = "https://www.npfprof.ru/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".nav__list a:eq(4)"))
                {
                    raskrprof = "https://www.npfprof.ru/" + obj.GetAttribute("href");
                }
                //Адрес 
                string adresnpfprof = "";
                foreach (IDomObject obj in dom.Find(".nav__list a:eq(5)"))
                {
                    adresnpfprof = "https://www.npfprof.ru/" + obj.GetAttribute("href");
                }
                //Лицензия
                string licenzprof = "";
                url = "https://www.npfprof.ru/info/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".section a:eq(2)"))
                {
                    licenzprof = "https://www.npfprof.ru/" + obj.GetAttribute("href");
                }
                //устав
                string ustavprof = "";
                foreach (IDomObject obj in dom.Find(".section a:eq(5)"))
                {
                    ustavprof = "https://www.npfprof.ru/" + obj.GetAttribute("href");
                }
            //Информация о процессе инвестирования
            string procinvestprof = "";
            url = "https://www.npfprof.ru/info/raskrytie-informatsii-v-sootvetstvii-s-5175-u/?special_version=Y";
            htmlCode = downloadPageUTF8(url);

            dom = CQ.CreateDocument(htmlCode);
            foreach (IDomObject obj in dom.Find(".catalog-list-content a:eq(8)"))
            {
                procinvestprof = "https://www.npfprof.ru/" + obj.GetAttribute("href");
            }


            //ГАЗПРОМБАНК-ФОНД
            displayChange("ГАЗПРОМБАНК-ФОНД");
            
                //раздел с раскрытием на главной странице
                string raskrgbf = "";
                url = "https://www.gpbf.ru/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".left-box a:eq(2)"))
                {
                    raskrgbf = "https://www.gpbf.ru/" + obj.GetAttribute("href");
                }
                //Структура и состав акционеров
                string sostavakciongbf = "";
                foreach (IDomObject obj in dom.Find(".col-xs-12.col-sm-5.col-lg-6 a:eq(2)"))
                {
                    sostavakciongbf = "https://www.gpbf.ru/" + obj.GetAttribute("href");
                }
                //Показатели деятельности
                string pokazatelgbf = "";
                url = "https://www.gpbf.ru/about/open-info/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".left-menu a:eq(1)"))
                {
                    pokazatelgbf = "https://www.gpbf.ru/" + obj.GetAttribute("href");
                }
                //Доходность
                string dohodgbf = "";
                foreach (IDomObject obj in dom.Find(".left-menu a:eq(2)"))
                {
                    dohodgbf = "https://www.gpbf.ru/" + obj.GetAttribute("href");
                }
                //Отчетность
                string otchetgbf = "";
                foreach (IDomObject obj in dom.Find(".left-menu a:eq(3)"))
                {
                    otchetgbf = "https://www.gpbf.ru/" + obj.GetAttribute("href");
                }
                //архив
                string arhivgbf = "";
                foreach (IDomObject obj in dom.Find(".left-menu a:eq(5)"))
                {
                    arhivgbf = "https://www.gpbf.ru/" + obj.GetAttribute("href");
                }
                //Структура портфеля
                string structportfgbf = "";
                foreach (IDomObject obj in dom.Find(".left-menu a:eq(8)"))
                {
                    structportfgbf = "https://www.gpbf.ru/" + obj.GetAttribute("href");
                }

                //УК
                string UKgbf = "";
                foreach (IDomObject obj in dom.Find(".left-menu a:eq(9)"))
                {
                    UKgbf = "https://www.gpbf.ru/" + obj.GetAttribute("href");
                }
                //СД
                string SDgbf = "";
                foreach (IDomObject obj in dom.Find(".left-menu a:eq(10)"))
                {
                    SDgbf = "https://www.gpbf.ru/" + obj.GetAttribute("href");
                }
                //устав
                string ustavgbf = "";
                foreach (IDomObject obj in dom.Find(".left-box a:eq(0)"))
                {
                    ustavgbf = "https://www.gpbf.ru/" + obj.GetAttribute("href");
                }
                //Лицензия
                string licenzgbf = "";
                foreach (IDomObject obj in dom.Find(".left-box a:eq(1)"))
                {
                    licenzgbf = "https://www.gpbf.ru/" + obj.GetAttribute("href");
                }
                //Пенсионные правила
                string penspravgbf = "";
                foreach (IDomObject obj in dom.Find(".left-box a:eq(3)"))
                {
                    penspravgbf = "https://www.gpbf.ru/" + obj.GetAttribute("href");
                }
            //информация о процессе инвестирования
            string procinvestgbf = "";
            foreach (IDomObject obj in dom.Find(".left-box a:eq(6)"))
            {
                procinvestgbf = "https://www.gpbf.ru/" + obj.GetAttribute("href");
            }
            //Органы управления
            string orguprgbf = "";
                foreach (IDomObject obj in dom.Find(".left-box a:eq(5)"))
                {
                    orguprgbf = "https://www.gpbf.ru/" + obj.GetAttribute("href");
                }

                //адрес
                string adressgbf = "";
                foreach (IDomObject obj in dom.Find(".col-xs-12.col-sm-5.col-lg-6 p:eq(0)"))
                {
                    adressgbf = obj.TextContent;
                }
                //телефон
                string telgbf = "";
                foreach (IDomObject obj in dom.Find(".phone a:eq(0)"))
                {
                    telgbf = obj.TextContent;
                }

          

            //НПФ ВТБ ПЕНСИОННЫЙ ФОНД
            displayChange("НПФ ВТБ ПЕНСИОННЫЙ ФОНД");
           
                //раздел с раскрытием на главной странице
                string raskrvtb = "";
                url = "https://www.vtbnpf.ru/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".header-submenu a:eq(15)"))
                {
                    raskrvtb = "https://www.vtbnpf.ru/" + obj.GetAttribute("href");
                }

                //Инвестиционная политика
                string investpolvtb = "";
                url = "https://www.vtbnpf.ru/achievment/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".container a:eq(3)"))
                {
                    investpolvtb = "https://www.vtbnpf.ru/" + obj.GetAttribute("href");
                }
                //адреса
                string adressvtb = "";
                foreach (IDomObject obj in dom.Find(".container a:eq(2)"))
                {
                    adressvtb = "https://www.vtbnpf.ru/" + obj.GetAttribute("href");
                }
                //СД
                string SDvtb = "";
                foreach (IDomObject obj in dom.Find(".container a:eq(4)"))
                {
                    SDvtb = "https://www.vtbnpf.ru/" + obj.GetAttribute("href");
                }
                //Результаты деятельности
                string rezultvtb = "";
                foreach (IDomObject obj in dom.Find(".container a:eq(5)"))
                {
                    rezultvtb = "https://www.vtbnpf.ru/" + obj.GetAttribute("href");
                }
                //реорганизация
                string reorgvtb = "";
                foreach (IDomObject obj in dom.Find(".container a:eq(6)"))
                {
                    reorgvtb = "https://www.vtbnpf.ru/" + obj.GetAttribute("href");
                }
                //Сообщения о существенных событиях
                string sobytvtb = "";
                foreach (IDomObject obj in dom.Find(".container a:eq(7)"))
                {
                    sobytvtb = "https://www.vtbnpf.ru/" + obj.GetAttribute("href");
                }
                //Отчетность
                string otchetvtb = "";
                foreach (IDomObject obj in dom.Find(".container a:eq(9)"))
                {
                    otchetvtb = "https://www.vtbnpf.ru/" + obj.GetAttribute("href");
                }
                //Органы управления
                string orguprvtb = "";
                foreach (IDomObject obj in dom.Find(".container a:eq(10)"))
                {
                    orguprvtb = "https://www.vtbnpf.ru/" + obj.GetAttribute("href");
                }
                //структура инвестиционного портфеля
                string structuraportfvtb = "";
                url = "https://www.vtbnpf.ru/achievment/result/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".container a:eq(4)"))
                {
                    structuraportfvtb = "https://www.vtbnpf.ru/" + obj.GetAttribute("href");
                }
                //устав
                string ustavvtb = "";
                url = "https://www.vtbnpf.ru/documents/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".container a:eq(2)"))
                {
                    ustavvtb = "https://www.vtbnpf.ru/" + obj.GetAttribute("href");
                }
                //Лицензия
                string licenzvtb = "";
                foreach (IDomObject obj in dom.Find(".container a:eq(3)"))
                {
                    licenzvtb = "https://www.vtbnpf.ru/" + obj.GetAttribute("href");
                }
               
                //Правила фонда
                string pravilavtb = "";
                foreach (IDomObject obj in dom.Find(".container a:eq(6)"))
                {
                    pravilavtb = "https://www.vtbnpf.ru/" + obj.GetAttribute("href");
                }
                //телефон
                string telvtb = "";
                foreach (IDomObject obj in dom.Find(".header__contacts a:eq(0)"))
                {
                    telvtb = obj.TextContent;
                }
                 //Структура и состав акционеров
                string sostavakcionvtb = "";
                url = "https://vtbnpf.ru/achievment/info/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".col-xl-10 a:eq(2)"))
                {
                    sostavakcionvtb = "https://www.vtbnpf.ru/" + obj.GetAttribute("href");
                }
            //Информация о конечных владельцах 
            string konvladvtb = "";
            foreach (IDomObject obj in dom.Find(".col-xl-10 a:eq(3)"))
            {
                konvladvtb = "https://www.vtbnpf.ru/" + obj.GetAttribute("href");
            }


            //ГАЗФОНД ПН
            displayChange("ГАЗФОНД ПН");
         
                //раздел с раскрытием на главной странице
                string raskrgazfondpn = "";
                url = "https://gazfond-pn.ru/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".aside-menu a:eq(5)"))
                {
                    raskrgazfondpn = "https://gazfond-pn.ru/" + obj.GetAttribute("href");
                }
                
                //Основные показатели
                string pokazatelgazfondpn = "";
                url = "https://gazfond-pn.ru/about/disclosure/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".information__wrapper a:eq(0)"))
                {
                    pokazatelgazfondpn = "https://gazfond-pn.ru/" + obj.GetAttribute("href");
                }
                //Официальные документы
                string ofdocgazfondpn = "";
                foreach (IDomObject obj in dom.Find(".information__wrapper a:eq(2)"))
                {
                    ofdocgazfondpn = "https://gazfond-pn.ru/" + obj.GetAttribute("href");
                }
                //отчетность
                string otchetgazfondpn = "";
                foreach (IDomObject obj in dom.Find(".information__wrapper a:eq(4)"))
                {
                    otchetgazfondpn = "https://gazfond-pn.ru/" + obj.GetAttribute("href");
                }
                //Структура инвестиционного портфеля
                string strinvestportfgazfondpn = "";
                foreach (IDomObject obj in dom.Find(".information__wrapper a:eq(10)"))
                {
                    strinvestportfgazfondpn = "https://gazfond-pn.ru/" + obj.GetAttribute("href");
                }
                //Состав инвестиционного портфеля
                string sostinvestportfgazfondpn = "";
                foreach (IDomObject obj in dom.Find(".information__wrapper a:eq(11)"))
                {
                    sostinvestportfgazfondpn = "https://gazfond-pn.ru/" + obj.GetAttribute("href");
                }
                //УК и СД
                string UKSDgazfondpn = "";
                foreach (IDomObject obj in dom.Find(".information__wrapper a:eq(12)"))
                {
                    UKSDgazfondpn = "https://gazfond-pn.ru/" + obj.GetAttribute("href");
                }
                //телефон
                string telgazfondpn = "";
                foreach (IDomObject obj in dom.Find(".aside-menu__body a:eq(0)"))
                {
                    telgazfondpn = obj.TextContent;
                }
                //адреса
                string adressgazfondpn = "";
                url = "https://gazfond-pn.ru/about/requisites/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".text-block__main a:eq(3)"))
                {
                    adressgazfondpn = "https://gazfond-pn.ru/" + obj.GetAttribute("href");
                }
            //информация о процессе инвестирования
            string procinvestgazfondpn = "";
            url = "https://gazfond-pn.ru/about/disclosure/investment_policy/";
            htmlCode = downloadPageUTF8(url);

            dom = CQ.CreateDocument(htmlCode);
            foreach (IDomObject obj in dom.Find(".documents__wrapper a:eq(1)"))
            {
                procinvestgazfondpn = "https://gazfond-pn.ru/" + obj.GetAttribute("href");
            }





            //НПФ НАЦИОНАЛЬНЫЙ
            displayChange("НПФ НАЦИОНАЛЬНЫЙ");
           
                string raskrnnpf = "";
                url = "https://www.nnpf.ru/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".header__burger-item a:eq(1)"))
                {
                    raskrnnpf = "https://www.nnpf.ru/" + obj.GetAttribute("href");
                }
                //Документы
                string docnnpf = "";
                url = "https://www.nnpf.ru/information-disclosure/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".aside-menu__links a:eq(1)"))
                {
                    docnnpf = "https://www.nnpf.ru/" + obj.GetAttribute("href");
                }
                //Акционеры фонда
                string akcionnnpf = "";
                foreach (IDomObject obj in dom.Find(".aside-menu__links a:eq(2)"))
                {
                    akcionnnpf = "https://www.nnpf.ru/" + obj.GetAttribute("href");
                }
                //О результате инвестирования
                string rezinvestnnpf = "";
                foreach (IDomObject obj in dom.Find(".aside-menu__links a:eq(4)"))
                {
                    rezinvestnnpf = "https://www.nnpf.ru/" + obj.GetAttribute("href");
                }
                //О количестве вкладчиков, участников и застрахованных лиц
                string kolvkladnnpf = "";
                foreach (IDomObject obj in dom.Find(".aside-menu__links a:eq(5)"))
                {
                    kolvkladnnpf = "https://www.nnpf.ru/" + obj.GetAttribute("href");
                }
                //О размерах пенсионных резервов и пенсионных накоплений
                string razmerprpnnnpf = "";
                foreach (IDomObject obj in dom.Find(".aside-menu__links a:eq(6)"))
                {
                    razmerprpnnnpf = "https://www.nnpf.ru/" + obj.GetAttribute("href");
                }
                //Бухгалтерская (финансовая отчётность)
                string buhotchetnnpf = "";
                foreach (IDomObject obj in dom.Find(".aside-menu__links a:eq(7)"))
                {
                    buhotchetnnpf = "https://www.nnpf.ru/" + obj.GetAttribute("href");
                }
                //Аудиторское заключение
                string audzaklnnpf = "";
                foreach (IDomObject obj in dom.Find(".aside-menu__links a:eq(8)"))
                {
                    audzaklnnpf = "https://www.nnpf.ru/" + obj.GetAttribute("href");
                }
                //Актуарное заключение
                string actuarzaklnnpf = "";
                foreach (IDomObject obj in dom.Find(".aside-menu__links a:eq(9)"))
                {
                    actuarzaklnnpf = "https://www.nnpf.ru/" + obj.GetAttribute("href");
                }
                //Отчетность 
                string otchetnostnnpf = "";
                otchetnostnnpf = "Бухгалтерская отчетность:" + " " + buhotchetnnpf + " " + "Аудиторское заключение:" + " " + audzaklnnpf + " " + "Актуарное заключение:" + " " + actuarzaklnnpf;
            //Отчет о формировании СПН
            string otchetSPNnnpf = "";
            foreach (IDomObject obj in dom.Find(".aside-menu__links a:eq(10)"))
            {
                otchetSPNnnpf = "https://www.nnpf.ru/" + obj.GetAttribute("href");
            }
            //Сведения об инвестировании
            string investirnnpf = "";
            foreach (IDomObject obj in dom.Find(".aside-menu__links a:eq(11)"))
            {
                investirnnpf = "https://www.nnpf.ru/" + obj.GetAttribute("href");
            }
            //Управляющие компании и специализированный депозитарий
            string UKSDnnpf = "";
            foreach (IDomObject obj in dom.Find(".aside-menu__links a:eq(12)"))
            {
                UKSDnnpf = "https://www.nnpf.ru/" + obj.GetAttribute("href");
            }
            //Место нахождения и телефоны Фонда и его обособленных подразделений
            string contactnnpf = "";
            foreach (IDomObject obj in dom.Find(".aside-menu__links a:eq(15)"))
            {
                contactnnpf = "https://www.nnpf.ru/" + obj.GetAttribute("href");
            }
            //Органы управления и должностные лица Фонда
            string orguprnnpf = "";
            foreach (IDomObject obj in dom.Find(".aside-menu__links a:eq(17)"))
            {
                orguprnnpf = "https://www.nnpf.ru/" + obj.GetAttribute("href");
            }
            //Информация о событиях 
            string sobnnpf = "";
            url = "https://www.nnpf.ru/information-disclosure/investment-details/";
            htmlCode = downloadPageUTF8(url);

            dom = CQ.CreateDocument(htmlCode);
            foreach (IDomObject obj in dom.Find(".article a:eq(0)"))
            {
                sobnnpf = "https://www.nnpf.ru/" + obj.GetAttribute("href");
            }







            //НПФ АТОМФОНД
            displayChange("НПФ АТОМФОНД");
          
                //Раздел с информацией о раскрытии на главной странице
                string raskratomfond = "";
                url = "https://www.atomfond.ru/";
                htmlCode = downloadPage(url);
                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".dropdown a:eq(5)"))
                {

                    raskratomfond = "https://www.atomfond.ru/" + obj.GetAttribute("href");
                }
                //Структура и состав акционеров
                string akcionatomfond = "";
                foreach (IDomObject obj in dom.Find(".dropdown a:eq(7)"))
                {

                    akcionatomfond = "https://www.atomfond.ru/" + obj.GetAttribute("href");
                }
                //Ораны управления и контроля
                string orgupratomfond = "";
                foreach (IDomObject obj in dom.Find(".dropdown a:eq(8)"))
                {

                    orgupratomfond = "https://www.atomfond.ru/" + obj.GetAttribute("href");
                }
                //Финансовая отченость
                string finotchetatomfond = "";
                foreach (IDomObject obj in dom.Find(".dropdown a:eq(9)"))
                {

                    finotchetatomfond = "https://www.atomfond.ru/" + obj.GetAttribute("href");
                }
                //Показатели деятельности
                string pokazatelatomfond = "";
                foreach (IDomObject obj in dom.Find(".dropdown a:eq(12)"))
                {
                    pokazatelatomfond = "https://www.atomfond.ru/" + obj.GetAttribute("href");
                }
                //УК
                string UKatomfond = "";
                foreach (IDomObject obj in dom.Find(".dropdown a:eq(13)"))
                {
                    UKatomfond = "https://www.atomfond.ru/" + obj.GetAttribute("href");
                }
                //СД
                string SDatomfond = "";
                foreach (IDomObject obj in dom.Find(".dropdown a:eq(14)"))
                {
                    SDatomfond = "https://www.atomfond.ru/" + obj.GetAttribute("href");
                }
                //Устав
                string ustavatomfond = "";
                url = "https://www.atomfond.ru/information_disclosure/dokumenty-dlya-raskrytiya/";
                htmlCode = downloadPage(url);
                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".col-md-9 a:eq(0)"))
                {

                    ustavatomfond = "https://www.atomfond.ru/" + obj.GetAttribute("href");
                }
                //лицензия
                string licenzatomfond = "";
                foreach (IDomObject obj in dom.Find(".col-md-9 a:eq(1)"))
                {

                    licenzatomfond = "https://www.atomfond.ru/" + obj.GetAttribute("href");
                }
                //Страховые правила
                string strahpravatomfond = "";
                foreach (IDomObject obj in dom.Find(".col-md-9 a:eq(2)"))
                {

                    strahpravatomfond = "https://www.atomfond.ru/" + obj.GetAttribute("href");
                }
                //Адрес
                string adresatomfond = "";
                url = "https://www.atomfond.ru/contacts/";
                htmlCode = downloadPage(url);
                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".col-md-9 li:eq(0)"))
                {

                    adresatomfond = obj.TextContent;
                }
                //телефон
                string telatomfond = "";
                foreach (IDomObject obj in dom.Find(".col-md-9 li:eq(1)"))
                {

                    telatomfond = obj.TextContent;
                }
                //Результаты инвестирования ПН
                string rezinvestPNatomfond = "";
                url = "https://www.atomfond.ru/information_disclosure/rezultaty-investirovaniya/";
                htmlCode = downloadPage(url);
                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".dop_menu a:eq(0)"))
                {

                    rezinvestPNatomfond = "https://www.atomfond.ru/" + obj.GetAttribute("href");
                }
                //Структура инвестиционного портфеля
                string structuraportfatomfond = "";
                foreach (IDomObject obj in dom.Find(".dop_menu a:eq(2)"))
                {

                    structuraportfatomfond = "https://www.atomfond.ru/" + obj.GetAttribute("href");
                }
                //Информация о событиях, существенно влияющих на стоимость активов, в которые инвестированы пенсионные накопления
                string sobytatomfond = "";
                foreach (IDomObject obj in dom.Find(".dop_menu a:eq(4)"))
                {

                    sobytatomfond = "https://www.atomfond.ru/" + obj.GetAttribute("href");
                }
            //Информация о процессе инвестирования
            string procinvestatomfond = "";
            foreach (IDomObject obj in dom.Find(".dop_menu a:eq(3)"))
            {

                procinvestatomfond = "https://www.atomfond.ru/" + obj.GetAttribute("href");
            }

            //нформация о составе инвестиционного портфеля по обязательному пенсионному страхованию
            string sostavportfatomfond = "";
                foreach (IDomObject obj in dom.Find(".dop_menu a:eq(6)"))
                {
                    sostavportfatomfond = "https://www.atomfond.ru/" + obj.GetAttribute("href");
                }
            //Решения БР о запрете проведения всех или части операций
            string zapratomfond = "";
            url = "https://www.atomfond.ru//information_disclosure/rezultaty-investirovaniya/resheniya-cbrf/";
            htmlCode = downloadPage(url);
            dom = CQ.CreateDocument(htmlCode);
            foreach (IDomObject obj in dom.Find(".inner ul"))
            {
                IDomObject tmp = obj.ParentNode;
                tmp.RemoveChild(obj);
                zapratomfond = tmp.InnerText;
            }



            //НПФ СБЕРФОНД
            displayChange("НПФ СБЕРФОНД");
         
                //раздел с раскрытием на главной странице
                string raskrsberfond = "";
                url = "http://www.sberfond.ru/";
                htmlCode = downloadPage(url);
                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".art-hmenu a:eq(4)"))
                {
                    raskrsberfond = "http://www.sberfond.ru/" + obj.GetAttribute("href");
                }
                //Устав
                string ustavsberfond = "";
                foreach (IDomObject obj in dom.Find(".art-box-body.art-blockcontent-body a:eq(3)"))
                {
                    ustavsberfond = "http://www.sberfond.ru/" + obj.GetAttribute("href");
                }
                //Лицензия
                string licenzsberfond = "";
                foreach (IDomObject obj in dom.Find(".art-box-body.art-blockcontent-body a:eq(1)"))
                {
                    licenzsberfond = "http://www.sberfond.ru/" + obj.GetAttribute("href");
                }
                //Пенсионные правила
                string penspravsberfond = "";
                foreach (IDomObject obj in dom.Find(".art-box-body.art-blockcontent-body a:eq(5)"))
                {
                    penspravsberfond = "http://www.sberfond.ru/" + obj.GetAttribute("href");
                }
                //телефон
                string telsberfond = "";
                foreach (IDomObject obj in dom.Find(".art-logo-name a:eq(0)"))
                {
                    telsberfond = obj.TextContent;
                }
                //адрес
                string adressberfond = "";
                url = "http://www.sberfond.ru/index.php?option=com_content&view=article&id=2&Itemid=2";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".art-article p:eq(8)"))
                {
                    adressberfond = obj.TextContent;
                }
            

            //НПФ ИНОГОССТРАХ ПЕНСИЯ
            displayChange("НПФ ИНОГОССТРАХ ПЕНСИЯ");
          
                //раздел с раскрытием информации на главной странице
                string raskringo = "";
                url = "https://ingospensiya.ru/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".tabs a:eq(2)"))
                {
                    raskringo = "https://ingospensiya.ru/" + obj.GetAttribute("href");
                }
                //Лицензия
                string licenzingo = "";
                url = "https://ingospensiya.ru/info/documents/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".downloads a:eq(0)"))
                {
                    licenzingo = "https://ingospensiya.ru/" + obj.GetAttribute("href");
                }
                //Устав
                string ustavingo = "";
                foreach (IDomObject obj in dom.Find(".downloads a:eq(8)"))
                {
                    ustavingo = "https://ingospensiya.ru/" + obj.GetAttribute("href");
                }
                //Пенсионные правила
                string penspravingo = "";
                foreach (IDomObject obj in dom.Find(".downloads a:eq(10)"))
                {
                    penspravingo = "https://ingospensiya.ru/" + obj.GetAttribute("href");
                }
                //Официальные документы и отчетность
                string otchetingo = "";
                url = "https://ingospensiya.ru/info/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".item-text a:eq(0)"))
                {
                    otchetingo = "https://ingospensiya.ru/" + obj.GetAttribute("href");
                }
                //Инвестиционная деятельность
                string investingo = "";
                foreach (IDomObject obj in dom.Find(".item-text a:eq(2)"))
                {
                    investingo = "https://ingospensiya.ru/" + obj.GetAttribute("href");
                }
                //Состав и структура акционеров
                string sostavakcioningo = "";
                foreach (IDomObject obj in dom.Find(".item-text a:eq(4)"))
                {
                    sostavakcioningo = "https://ingospensiya.ru/" + obj.GetAttribute("href");
                }
                //Органы управления
                string orgupringo = "";
                url = "https://ingospensiya.ru/about/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".item-text a:eq(2)"))
                {
                    orgupringo = "https://ingospensiya.ru/" + obj.GetAttribute("href");
                }
                //Адрес
                string adresingo = "";
                url = "https://ingospensiya.ru/about/contacts/moscow/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".column_one p:eq(0)"))
                {
                    adresingo = obj.TextContent;
                }
                //Телефон
                string telingo = "";
                foreach (IDomObject obj in dom.Find(".column_one p:eq(1)"))
                {
                    telingo = obj.TextContent;
                }

          

            //НПФ КОРАБЕЛ
            displayChange("НПФ КОРАБЕЛ");
         
                //раздел с раскрытием на главной странице
                string raskrkorabel = "";
                url = "https://www.npf-korabel.spb.ru/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".container a:eq(25)"))
                {
                    raskrkorabel = "https://www.npf-korabel.spb.ru/" + obj.GetAttribute("href");
                }
                //Бух отчет
                string buhotchetkorabel = "";
                foreach (IDomObject obj in dom.Find(".container a:eq(26)"))
                {
                    buhotchetkorabel = "https://www.npf-korabel.spb.ru/" + obj.GetAttribute("href");
                }
                //Показатели деятельности
                string pokazatelkorabel = "";
                foreach (IDomObject obj in dom.Find(".container a:eq(27)"))
                {
                    pokazatelkorabel = "https://www.npf-korabel.spb.ru/" + obj.GetAttribute("href");
                }
                //лицензия
                string licenzkorabel = "";
                foreach (IDomObject obj in dom.Find(".container a:eq(6)"))
                {
                    licenzkorabel = "https://www.npf-korabel.spb.ru/" + obj.GetAttribute("href");
                }
                //Нормативная документация
                string dockorabel = "";
                foreach (IDomObject obj in dom.Find(".container a:eq(7)"))
                {
                    dockorabel = "https://www.npf-korabel.spb.ru/" + obj.GetAttribute("href");
                }
                //Состав акционеров
                string akcionkorabel = "";
                foreach (IDomObject obj in dom.Find(".container a:eq(8)"))
                {
                    akcionkorabel = "https://www.npf-korabel.spb.ru/" + obj.GetAttribute("href");
                }
                //Органы управления
                string orguprkorabel = "";
                foreach (IDomObject obj in dom.Find(".container a:eq(9)"))
                {
                    orguprkorabel = "https://www.npf-korabel.spb.ru/" + obj.GetAttribute("href");
                }
                //УК
                string UKkorabel = "";
                foreach (IDomObject obj in dom.Find(".container a:eq(10)"))
                {
                    UKkorabel = "https://www.npf-korabel.spb.ru/" + obj.GetAttribute("href");
                }
                //СД
                string SDkorabel = "";
                foreach (IDomObject obj in dom.Find(".container a:eq(11)"))
                {
                    SDkorabel = "https://www.npf-korabel.spb.ru/" + obj.GetAttribute("href");
                }
                //адрес
                string adreskorabel = "";
                foreach (IDomObject obj in dom.Find(".custom:eq(0)"))
                {
                    adreskorabel = obj.TextContent;
                }
                //Телефон
                string telkorabel = "";
                foreach (IDomObject obj in dom.Find(".custom p:eq(1)"))
                {
                    telkorabel = obj.TextContent.Split('.')[1];
                    telkorabel = telkorabel.Replace("E-Mail:", "");
                }

          

            //НПФ СОЦИУМ
            displayChange("НПФ СОЦИУМ");
        
                //раздел с раскрытием информации на главной странице
                string raskrytsocium = "";
                url = "https://npfsocium.ru/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".tabs a:eq(2)"))
                {
                    raskrytsocium = "https://npfsocium.ru/" + obj.GetAttribute("href");
                }
                //Устав
                string ustavsocium = "";
                url = "https://npfsocium.ru/info/documents/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".downloads a:eq(0)"))
                {
                    ustavsocium = "https://npfsocium.ru/" + obj.GetAttribute("href");
                }
                //Лицензия
                string licenzsocium = "";
                foreach (IDomObject obj in dom.Find(".downloads a:eq(2)"))
                {
                    licenzsocium = "https://npfsocium.ru/" + obj.GetAttribute("href");
                }
                //Пенсионные правила
                string penspravsocium = "";
                foreach (IDomObject obj in dom.Find(".downloads a:eq(1)"))
                {
                    penspravsocium = "https://npfsocium.ru/" + obj.GetAttribute("href");
                }
                //Страховые правила
                string strahpravsocium = "";
                foreach (IDomObject obj in dom.Find(".downloads a:eq(13)"))
                {
                    strahpravsocium = "https://npfsocium.ru/" + obj.GetAttribute("href");
                }
                //Официальные документы и отчетность
                string docsocium = "";
                url = "https://npfsocium.ru/info/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".item-text a:eq(0)"))
                {
                    docsocium = "https://npfsocium.ru/" + obj.GetAttribute("href");
                }
                //Инвестиционная деятельность
                string investsocium = "";
                foreach (IDomObject obj in dom.Find(".item-text a:eq(1)"))
                {
                    investsocium = "https://npfsocium.ru/" + obj.GetAttribute("href");
                }
                //Органы управления
                string orguprsocium = "";
                url = "https://npfsocium.ru/about/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".item-text a:eq(2)"))
                {
                    orguprsocium = "https://npfsocium.ru/" + obj.GetAttribute("href");
                }
            //Адреса филлиалов
            string adressociumpodrazdel = "";
            foreach (IDomObject obj in dom.Find(".item-text a:eq(3)"))
            {
               adressociumpodrazdel = "https://npfsocium.ru/" + obj.GetAttribute("href");
            }
            //Адрес
            string adressociummsk = "";
                url = "https://npfsocium.ru/about/contacts/moscow/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".column_one p:eq(0)"))
                {
                    adressociummsk = obj.TextContent;
                }
            string adressocium = "";
            adressocium = "Головной офис:" + " " + adressociummsk + "\n" + "Место нахождение обособленных подразделений:" + " " + adressociumpodrazdel;


                //Телефон
            string telsocium = "";
                foreach (IDomObject obj in dom.Find(".column_one p:eq(1)"))
                {
                    telsocium = obj.TextContent;
                }

          

            //НПФ БЛАГОСОСТОЯНИЕ
            displayChange("НПФ БЛАГОСОСТОЯНИЕ");
          
                //Раздел с раскрытием информации на главном экране
                string razdelraskrnpfb = "";
                url = "https://npfb.ru/";
                htmlCode = downloadPage(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".menu a:eq(8)"))
                {
                    razdelraskrnpfb = "https://npfb.ru/" + obj.GetAttribute("href");
                }
                //Контакты
                string adresnpfb = "";
                foreach (IDomObject obj in dom.Find(".menu a:eq(12)"))
                {
                    adresnpfb = "https://npfb.ru/" + obj.GetAttribute("href");
                }
                //Лицензия
                string licenznpfb = "";
                url = "https://npfb.ru/o-fonde/raskrytie-informatsii/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".about-info a:eq(0)"))
                {
                    licenznpfb = "https://npfb.ru/" + obj.GetAttribute("href");
                }
                //Пенсионные правила
                string penspravnpfb = "";
                foreach (IDomObject obj in dom.Find(".about-info a:eq(3)"))
                {
                    penspravnpfb = "https://npfb.ru/" + obj.GetAttribute("href");
                }
                //Архив пенсионных правил
                string arhpenspravnpfb = "";
                foreach (IDomObject obj in dom.Find(".about-info a:eq(4)"))
                {
                    arhpenspravnpfb = "https://npfb.ru/" + obj.GetAttribute("href");
                }
                //Телефон
                string telnpfb = "";
                foreach (IDomObject obj in dom.Find(".accordion__itemlist a:eq(1)"))
                {
                    string tag = obj.TextContent;
                    telnpfb = tag;
                }
                //Устав
                string ustavnpfb = "";
                foreach (IDomObject obj in dom.Find(".documents__item a:eq(0)"))
                {
                    ustavnpfb = "https://npfb.ru/" + obj.GetAttribute("href");
                }
                //аудиторское заключение
                string audzaclnp = "";
                foreach (IDomObject obj in dom.Find("#reporting div:eq(0) .disclosure__reporting__list__item__link:last"))
                {
                    audzaclnp = "https://npfb.ru/" + obj.GetAttribute("href");
                }
           
            //Информация о деятельности: кол-во вкладчиков, участников, размер ПР
            string kolvovklnpfb = "";
            url = "https://npfb.ru/o-fonde/raskrytie-informatsii/";
            htmlCode = downloadPageUTF8(url);

            dom = CQ.CreateDocument(htmlCode);
            foreach (IDomObject obj in dom.Find(".disclosure__reporting__list__item a:eq(0)"))
            {
                kolvovklnpfb = "https://npfb.ru/" + obj.GetAttribute("href");
            }


            //НПФ АВИАПОЛИС
            displayChange("АВИАПОЛИС");
           
                //раздел с раскрытием на главной странице
                string raskraviapolis = "";
                url = "https://www.npf-aviapolis.ru/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find("#navigation a:eq(1)"))
                {
                    raskraviapolis = "https://www.npf-aviapolis.ru/" + obj.GetAttribute("href");
                }
                //Реорганизация
                string reorgaviapolis = "";
                foreach (IDomObject obj in dom.Find("#navigation a:eq(2)"))
                {
                    reorgaviapolis = "https://www.npf-aviapolis.ru/" + obj.GetAttribute("href");
                }
                //Адрес
                string adresaviapolis = "";
                htmlCode = downloadPage(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".style36"))
                {
                    adresaviapolis = obj.TextContent.Split('Т')[0].Split(':')[1];
                }

                //Список акционеров
                string akcionaviapolis = "";
                url = "https://www.npf-aviapolis.ru/index.files/Page023.html";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find("#navigation a:eq(8)"))
                {
                    akcionaviapolis = "https://www.npf-aviapolis.ru/index.files/" + obj.GetAttribute("href");
                }
                //УК
                string UKaviapolis = "";
                foreach (IDomObject obj in dom.Find("#navigation a:eq(10)"))
                {
                    UKaviapolis = "https://www.npf-aviapolis.ru/index.files/" + obj.GetAttribute("href");
                }
                //СД
                string SDaviapolis = "";
                foreach (IDomObject obj in dom.Find("#navigation a:eq(11)"))
                {
                    SDaviapolis = "https://www.npf-aviapolis.ru/index.files/" + obj.GetAttribute("href");
                }
                //Отчетность
                string otchetaviapolis = "";
                foreach (IDomObject obj in dom.Find("#navigation a:eq(12)"))
                {
                    otchetaviapolis = "https://www.npf-aviapolis.ru/index.files/" + obj.GetAttribute("href");
                }
                //Документы
                string docaviapolis = "";
                foreach (IDomObject obj in dom.Find("#navigation a:eq(13)"))
                {
                    docaviapolis = "https://www.npf-aviapolis.ru/index.files/" + obj.GetAttribute("href");
                }
                //Показатели деятельности
                string pokazatelaviapolis = "";
                foreach (IDomObject obj in dom.Find("#navigation a:eq(14)"))
                {
                    pokazatelaviapolis = "https://www.npf-aviapolis.ru/index.files/" + obj.GetAttribute("href");
                }
                //Состав средст ПР
                string sostavportfaviapolis = "";
                foreach (IDomObject obj in dom.Find("#navigation a:eq(18)"))
                {
                    sostavportfaviapolis = "https://www.npf-aviapolis.ru/index.files/" + obj.GetAttribute("href");
                }
                //Структура средств ПР
                string structuraportfaviapolis = "";
                foreach (IDomObject obj in dom.Find("#navigation a:eq(16)"))
                {
                    structuraportfaviapolis = "https://www.npf-aviapolis.ru/index.files/" + obj.GetAttribute("href");
                }
                //Органы управления
                string orgupraviapolis = "";
                foreach (IDomObject obj in dom.Find("#navigation a:eq(15)"))
                {
                    orgupraviapolis = "https://www.npf-aviapolis.ru/index.files/" + obj.GetAttribute("href");
                }

          
            //НПФ ВНЕШЭКОНОМФОНД
            displayChange("НПФ ВНЕШЭКОНОМФОНД");
           
                //раздел с раскрытием информации на главной странице
                string raskrnpfveb = "";
                url = "https://npfveb.ru/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);

                foreach (IDomObject obj in dom.Find(".mid a:eq(5)"))
                {
                    raskrnpfveb = "https://npfveb.ru/" + obj.GetAttribute("href");
                }
                //Устав
                string ustavnpfveb = "";
                url = "https://npfveb.ru/doc%20(subsections)/%D0%A0%D0%B0%D1%81%D0%BA%D1%80%D1%8B%D1%82%D0%B8%D0%B5%205175/index.php";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);

                foreach (IDomObject obj in dom.Find(".fon > span > table tr:eq(3) td:eq(2) a"))
                {
                    ustavnpfveb = "https://npfveb.ru/" + obj.GetAttribute("href");
                }
                //Лицензия
                string licenznpfveb = "";
                htmlCode = downloadPage(url);
                dom = CQ.CreateDocument(htmlCode);

                foreach (IDomObject obj in dom.Find(".fon > span > table tr:eq(2) td:eq(2) a"))
                {
                    licenznpfveb = "https://npfveb.ru/" + obj.GetAttribute("href");
                }
                //Место нахождения Фонда и его обособленных подразделений
                string adresnpfveb = "";
                foreach (IDomObject obj in dom.Find(".fon > span > table tr:eq(4) td:eq(2) a"))
                {
                    adresnpfveb = obj.GetAttribute("href");
                }
                //Номер телефона Фонда
                string telnpfveb = "";
                foreach (IDomObject obj in dom.Find(".fon > span > table tr:eq(5) td:eq(2) a"))
                {
                    telnpfveb = "https://npfveb.ru/" + obj.GetAttribute("href");
                }
                //Сведения об органах управления, членах совета директоров (наблюдательного совета) фонда, должностных лицах и работниках фонда: фамилия, имя, отчество, сведения об образовании, основном месте работы, опыте работы в кредитных организациях и некредитных финансовых организациях
                string orguprnpfveb = "";
                foreach (IDomObject obj in dom.Find(".fon > span > table tr:eq(6) td:eq(2) a"))
                {
                    orguprnpfveb = "https://npfveb.ru/" + obj.GetAttribute("href");
                }
                //Бухгалтерская (финансовая) отчетность фонда, аудиторское и актуарное заключения
                string buhotchetnpfveb = "";
                foreach (IDomObject obj in dom.Find(".fon > span > table tr:eq(7) td:eq(2) a"))
                {
                    buhotchetnpfveb = "https://npfveb.ru/" + obj.GetAttribute("href");
                }
                //Структура и состав акционеров
                string sostavakcionnpfveb = "";
                foreach (IDomObject obj in dom.Find(".fon > span > table tr:eq(8) td:eq(2) a"))
                {
                    sostavakcionnpfveb = obj.GetAttribute("href");
                }
                //Результат размещения средств пенсионных резервов.
                string rezrazmPRnpfveb = "";
                foreach (IDomObject obj in dom.Find(".fon > span > table tr:eq(9) td:eq(2) a"))
                {
                    rezrazmPRnpfveb = "https://npfveb.ru/" + obj.GetAttribute("href");
                }
                //Средневзвешенный процент
                string srvzvprocnpfveb = "";
                foreach (IDomObject obj in dom.Find(".fon > span > table tr:eq(10) td:eq(2) a"))
                {
                    srvzvprocnpfveb = "https://npfveb.ru/" + obj.GetAttribute("href");
                }
                //Размер дохода от размещения пенсионных резервов,
                string razmerdohodaPRnpfveb = "";
                foreach (IDomObject obj in dom.Find(".fon > span > table tr:eq(11) td:eq(2) a"))
                {
                    razmerdohodaPRnpfveb = "https://npfveb.ru/" + obj.GetAttribute("href");
                }
                //Информация о деятельности Фонда
                string infnpfveb = "";
                foreach (IDomObject obj in dom.Find(".fon > span > table tr:eq(12) td:eq(2) a"))
                {
                    infnpfveb = "https://npfveb.ru/" + obj.GetAttribute("href");
                }
                //Информация о заключении и прекращении действия договора доверительного управления пенсионными резервами с управляющей компанией с указанием ее фирменного наименования и номера лицензии
                string UKnpfveb = "";
                foreach (IDomObject obj in dom.Find(".fon > span > table tr:eq(13) td:eq(2) a"))
                {
                    UKnpfveb = "https://npfveb.ru/" + obj.GetAttribute("href");
                }
                //Информация о заключении и прекращении договора со специализированным депозитарием
                string SDnpfveb = "";
                foreach (IDomObject obj in dom.Find(".fon > span > table tr:eq(14) td:eq(2) a:eq(0)"))
                {
                    SDnpfveb = "https://npfveb.ru/" + obj.GetAttribute("href");
                }
                //Информация о структуре инвестиционного портфеля фонда
                string strportfnpfveb = "";
                foreach (IDomObject obj in dom.Find(".fon > span > table tr:eq(18) td:eq(2) a"))
                {
                    strportfnpfveb = "https://npfveb.ru/" + obj.GetAttribute("href");
                }
                //Информация о составе средств пенсионных резервов фонда.
                string sostportfnpfveb = "";
                foreach (IDomObject obj in dom.Find(".fon > span > table tr:eq(19) td:eq(2) a"))
                {
                    sostportfnpfveb = "https://npfveb.ru/" + obj.GetAttribute("href");
                }
            //Информация о процессе инвестирования
            string procinvestnpfveb = "";
            foreach (IDomObject obj in dom.Find(".fon > span > table tr:eq(23) td:eq(2) a"))
            {
                procinvestnpfveb = "https://npfveb.ru/" + obj.GetAttribute("href");
            }


            //НПФ РОСТЕХ 
            displayChange("НПФ РОСТЕХ");
          
                //раздел с раскрытием на главной странице
                string raskrrosteh = "";
                url = "https://rostecnpf.ru/";
                htmlCode = downloadPage(url);
                dom = CQ.CreateDocument(htmlCode);

                foreach (IDomObject obj in dom.Find("#cssmenu a:eq(1)"))
                {
                    raskrrosteh = "https://rostecnpf.ru/" + obj.GetAttribute("href");
                }
                //Структура и состав акционеров
                string akcionerrosteh = "";
                foreach (IDomObject obj in dom.Find("#cssmenu a:eq(4)"))
                {
                    akcionerrosteh = "https://rostecnpf.ru/" + obj.GetAttribute("href");
                }
                //Устав
                string ustavrosteh = "";
                url = "https://rostecnpf.ru/obschaya-informatsiya.htm";
                htmlCode = downloadPage(url);
                dom = CQ.CreateDocument(htmlCode);

                foreach (IDomObject obj in dom.Find("#td-main a:eq(0)"))
                {
                    ustavrosteh = "https://rostecnpf.ru/" + obj.GetAttribute("href");
                }
                //Лицензия
                string licenzrosteh = "";
                foreach (IDomObject obj in dom.Find("#td-main a:eq(1)"))
                {
                    licenzrosteh = "https://rostecnpf.ru/" + obj.GetAttribute("href");
                }
                //Контакты
                string contactrosteh = "";
                foreach (IDomObject obj in dom.Find("#td-main a:eq(3)"))
                {
                    contactrosteh = "https://rostecnpf.ru/" + obj.GetAttribute("href");
                }
                //Органы управления и контроля
                string orguprrosteh = "";
                url = "https://rostecnpf.ru/raskrytie-informatsii.htm";
                htmlCode = downloadPage(url);
                dom = CQ.CreateDocument(htmlCode);

                foreach (IDomObject obj in dom.Find("#td-main a:eq(1)"))
                {
                    orguprrosteh = "https://rostecnpf.ru/" + obj.GetAttribute("href");
                }
                //Пенсионные правила
                string penspravrosteh = "";
                foreach (IDomObject obj in dom.Find("#td-main a:eq(2)"))
                {
                    penspravrosteh = "https://rostecnpf.ru/" + obj.GetAttribute("href");
                }
                //Страховые правила
                string strahpravrosteh = "";
                foreach (IDomObject obj in dom.Find("#td-main a:eq(3)"))
                {
                    strahpravrosteh = "https://rostecnpf.ru/" + obj.GetAttribute("href");
                }
                //УК и СД
                string UKSDrosteh = "";
                foreach (IDomObject obj in dom.Find("#td-main a:eq(4)"))
                {
                    UKSDrosteh = "https://rostecnpf.ru/" + obj.GetAttribute("href");
                }
                //Бух отчет
                string buhotchetrosteh = "";
                foreach (IDomObject obj in dom.Find("#td-main a:eq(6)"))
                {
                    buhotchetrosteh = "https://rostecnpf.ru/" + obj.GetAttribute("href");
                }
                //Ауд заключение
                string audzaklrosteh = "";
                foreach (IDomObject obj in dom.Find("#td-main a:eq(7)"))
                {
                    audzaklrosteh = "https://rostecnpf.ru/" + obj.GetAttribute("href");
                }
                //Актуарн закл
                string actuarzaklrosteh = "";
                foreach (IDomObject obj in dom.Find("#td-main a:eq(8)"))
                {
                    actuarzaklrosteh = "https://rostecnpf.ru/" + obj.GetAttribute("href");
                }
                //Отчет
                buhotchetrosteh = buhotchetrosteh + " " + audzaklrosteh + " " + actuarzaklrosteh;
                //Показатели деятельности
                string pokazatelrosteh = "";
                foreach (IDomObject obj in dom.Find("#td-main a:eq(9)"))
                {
                    pokazatelrosteh = "https://rostecnpf.ru/" + obj.GetAttribute("href");
                }
                //Отчет о формировании средств ПН
                string formirPNrosteh = "";
                foreach (IDomObject obj in dom.Find("#td-main a:eq(10)"))
                {
                    formirPNrosteh = "https://rostecnpf.ru/" + obj.GetAttribute("href");
                }
                //МСФО
                string MSFOrosteh = "";
                foreach (IDomObject obj in dom.Find("#td-main a:eq(11)"))
                {
                    MSFOrosteh = "https://rostecnpf.ru/" + obj.GetAttribute("href");
                }
                //Инвестиционная деятельность
                string investrosteh = "";
                foreach (IDomObject obj in dom.Find("#td-main a:eq(14)"))
                {
                    investrosteh = "https://rostecnpf.ru/" + obj.GetAttribute("href");
                }

         

            //НПФ ТРАДИЦИЯ
            displayChange("НПФ ТРАДИЦИЯ");
           
                //раздел с раскрытием на главной странице
                string raskrtrad = "";
                url = "https://tradnpf.com/";
                htmlCode = downloadPage(url);
                dom = CQ.CreateDocument(htmlCode);

                foreach (IDomObject obj in dom.Find("#menu0 a:eq(1)"))
                {
                    raskrtrad = "https://tradnpf.com/" + obj.GetAttribute("href");
                }
                //Органы управления
                string orguprtrad = "";
                url = "https://tradnpf.com/raskrytie-informacii/o-fonde/";
                htmlCode = downloadPage(url);
                dom = CQ.CreateDocument(htmlCode);

                foreach (IDomObject obj in dom.Find("#content a:eq(1)"))
                {
                    orguprtrad = "https://tradnpf.com/" + obj.GetAttribute("href");
                }
                //Должностные лица
                string dolzhlicatrad = "";
                foreach (IDomObject obj in dom.Find("#content a:eq(2)"))
                {
                    dolzhlicatrad = "https://tradnpf.com/" + obj.GetAttribute("href");
                }
                //Сведения об органах управления
                string upravtrad = "";
                upravtrad = dolzhlicatrad + " " + orguprtrad;
                //Структура и состав акционеров
                string akciontrad = "";
                foreach (IDomObject obj in dom.Find("#content a:eq(3)"))
                {
                    akciontrad = "https://tradnpf.com/" + obj.GetAttribute("href");
                }
                //Результат размещения ПР
                string rezrazmPRtrad = "";
                url = "https://tradnpf.com/raskrytie-informacii/rezultaty-dejtelnosti/";
                htmlCode = downloadPage(url);
                dom = CQ.CreateDocument(htmlCode);

                foreach (IDomObject obj in dom.Find("#content a:eq(0)"))
                {
                    rezrazmPRtrad = "https://tradnpf.com/" + obj.GetAttribute("href");
                }
                //Средневзвешенный процент увеличения назначенных пенсий
                string srvzvproctrad = "";
                foreach (IDomObject obj in dom.Find("#content a:eq(1)"))
                {
                    srvzvproctrad = "https://tradnpf.com/" + obj.GetAttribute("href");
                }
                //Информация о размере дохода от размещения пенсионных резервов
                string razmerdohodatrad = "";
                foreach (IDomObject obj in dom.Find("#content a:eq(2)"))
                {
                    razmerdohodatrad = "https://tradnpf.com/" + obj.GetAttribute("href");
                }
                //Информация о количестве вкладчиков и участников Фонда
                string kolvkltrad = "";
                foreach (IDomObject obj in dom.Find("#content a:eq(3)"))
                {
                    kolvkltrad = "https://tradnpf.com/" + obj.GetAttribute("href");
                }
                //Страховой резерв
                string strahrezervtrad = "";
                foreach (IDomObject obj in dom.Find("#content a:eq(5)"))
                {
                    strahrezervtrad = "https://tradnpf.com/" + obj.GetAttribute("href");
                }
                //Информация о размере пенсионных резервов Фонда
                string razmerPRtrad = "";
                foreach (IDomObject obj in dom.Find("#content a:eq(4)"))
                {
                    razmerPRtrad = "https://tradnpf.com/" + obj.GetAttribute("href");
                }
                //Размер ПР и страхового резерва
                string razmerstrahPR = "";
                razmerstrahPR = strahrezervtrad + " " + razmerPRtrad;
                //Отчетность
                string otchettrad = "";
                url = "https://tradnpf.com/raskrytie-informacii/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find("#content a:eq(2)"))
                {
                    otchettrad = "https://tradnpf.com/" + obj.GetAttribute("href");
                }
                //Инвестиционная политика
                string investtrad = "";
                foreach (IDomObject obj in dom.Find("#content a:eq(3)"))
                {
                    investtrad = "https://tradnpf.com/" + obj.GetAttribute("href");
                }
                //Существенные события
                string sobyttrad = "";
                foreach (IDomObject obj in dom.Find("#content a:eq(6)"))
                {
                    sobyttrad = "https://tradnpf.com/" + obj.GetAttribute("href");
                }
                //УК
                string UKtrad = "";
                url = "https://tradnpf.com/upravljajushhie-kompanii-specdepozitarij/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find("#content a:eq(0)"))
                {
                    UKtrad = "https://tradnpf.com/" + obj.GetAttribute("href");
                }
                //СД
                string SDtrad = "";
                foreach (IDomObject obj in dom.Find("#content a:eq(1)"))
                {
                    SDtrad = "https://tradnpf.com/" + obj.GetAttribute("href");
                }
                //Устав
                string ustavtrad = "";
                url = "https://tradnpf.com/raskrytie-informacii/normativnaja-dokumentacija/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find("#content a:eq(0)"))
                {
                    ustavtrad = "https://tradnpf.com/" + obj.GetAttribute("href");
                }
                //Лицензия
                string licenztrad = "";
                foreach (IDomObject obj in dom.Find("#content a:eq(1)"))
                {
                    licenztrad = "https://tradnpf.com/" + obj.GetAttribute("href");
                }
                //Пенсионные правила
                string penspravtrad = "";
                foreach (IDomObject obj in dom.Find("#content a:eq(3)"))
                {
                    penspravtrad = "https://tradnpf.com/" + obj.GetAttribute("href");
                }
                //Телефон
                string teltrad = "";
                foreach (IDomObject obj in dom.Find(".contacts b:eq(0)"))
                {
                    teltrad = obj.TextContent;
                }
                //Адрес
                string adrestrad = "";
                url = "https://tradnpf.com/kontakty/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find("#content p:eq(2)"))
                {
                    adrestrad = obj.TextContent;
                }


            //НПФ ДОВЕРИЕ
            displayChange("НПФ ДОВЕРИЕ");
          
                //раздел с раскрытием на главной странице
                string raskrdoverie = "";
                url = "https://doverie56.ru/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".menu a:eq(31)"))
                {
                    raskrdoverie = "https://doverie56.ru/" + obj.GetAttribute("href");
                }
                //Контакты
                string contactdoverie = "";
                foreach (IDomObject obj in dom.Find(".menu a:eq(32)"))
                {
                    contactdoverie = "https://doverie56.ru/" + obj.GetAttribute("href");
                }
                //УК
                string UKdoverie = "";
                foreach (IDomObject obj in dom.Find(".menu a:eq(6)"))
                {
                    UKdoverie = "https://doverie56.ru/" + obj.GetAttribute("href");
                }
                //СД
                string SDdoverie = "";
                foreach (IDomObject obj in dom.Find(".menu a:eq(7)"))
                {
                    SDdoverie = "https://doverie56.ru/" + obj.GetAttribute("href");
                }
                //Лицензия
                string licenzdoverie = "";
                url = "https://doverie56.ru/raskrytie-informacii/oficialnye-dokumenty.html";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".content-description.innertCnt a:eq(0)"))
                {
                    licenzdoverie = "https://doverie56.ru/" + obj.GetAttribute("href");
                }
                //Устав
                string ustavdoverie = "";
                foreach (IDomObject obj in dom.Find(".content-description.innertCnt a:eq(1)"))
                {
                    ustavdoverie = "https://doverie56.ru/" + obj.GetAttribute("href");
                }
                //Официальные документы
                string ofdocdoverie = "";
                url = "https://doverie56.ru/raskrytie-informacii.html";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".list4a a:eq(0)"))
                {
                    ofdocdoverie = "https://doverie56.ru/" + obj.GetAttribute("href");
                }
                //Структура и состав акционеров
                string akciondoverie = "";
                foreach (IDomObject obj in dom.Find(".list4a a:eq(1)"))
                {
                    akciondoverie = "https://doverie56.ru/" + obj.GetAttribute("href");
                }
                //Информация о конечных владельцах
                string konechvladdoverie = "";
                foreach (IDomObject obj in dom.Find(".list4a a:eq(2)"))
                {
                    konechvladdoverie = "https://doverie56.ru/" + obj.GetAttribute("href");
                }
                //Сведения об органах управления, должностных лицах фонда
                string orguprdoverie = "";
                foreach (IDomObject obj in dom.Find(".list4a a:eq(3)"))
                {
                    orguprdoverie = "https://doverie56.ru/" + obj.GetAttribute("href");
                }
                //Отчетность
                string otchetdoverie = "";
                foreach (IDomObject obj in dom.Find(".list4a a:eq(10)"))
                {
                    otchetdoverie = "https://doverie56.ru/" + obj.GetAttribute("href");
                }


         

            //НПФ ОТКРЫТИЕ
            displayChange("НПФ ОТКРЫТИЕ");

            //Контакты
            string adresopen = "";
            url = "https://open-npf.ru/";
            htmlCode = downloadPageUTF8(url);

            dom = CQ.CreateDocument(htmlCode);
            foreach (IDomObject obj in dom.Find(".header__inner a:eq(5)"))
            {
                adresopen = "https://open-npf.ru/" + obj.GetAttribute("href");
            }

            //раздел с раскрытием на главной странице
            string raskropen = "";
                url = "https://open-npf.ru/about/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".op-cards-container a:eq(2)"))
                {
                    raskropen = "https://open-npf.ru/" + obj.GetAttribute("href");
                }
                //Органы управления
                string orgupropen = "";
                foreach (IDomObject obj in dom.Find(".op-cards-container a:eq(3)"))
                {
                    orgupropen = "https://open-npf.ru/" + obj.GetAttribute("href");
                }
                //Инвестирование
                string investopen = "";
                foreach (IDomObject obj in dom.Find(".op-cards-container a:eq(4)"))
                {
                    investopen = "https://open-npf.ru/" + obj.GetAttribute("href");
                }
                //Отчетность
                string otchetnostopen = "";
                url = "https://open-npf.ru/about/disclosure/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find("#disclosure-tab-2.tab-title.text-s-b.js-tab a:eq(0)"))
                {
                    otchetnostopen = "https://open-npf.ru/" + obj.GetAttribute("href"); 
                }
                



                
                //Устав
                string ustavopen = "";
                url = "https://open-npf.ru/about/disclosure/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".row a:eq(2)"))
                {
                    ustavopen = "https://open-npf.ru/" + obj.GetAttribute("href");
                }
                //Страховые правила
                string strahpravopen = "";
                foreach (IDomObject obj in dom.Find(".row a:eq(4)"))
                {
                    strahpravopen = "https://open-npf.ru/" + obj.GetAttribute("href");
                }
                //Пенсионные правила
                string penspravopen = "";
                foreach (IDomObject obj in dom.Find(".row a:eq(6)"))
                {
                    penspravopen = "https://open-npf.ru/" + obj.GetAttribute("href");
                }
                //Архив пенсионных правил
                string arhpenspravopen = "";
                foreach (IDomObject obj in dom.Find(".row a:eq(7)"))
                {
                    arhpenspravopen = "https://open-npf.ru/" + obj.GetAttribute("href");
                }
                //Лицнзия
                string licenzopen = "";
                foreach (IDomObject obj in dom.Find(".row a:eq(8)"))
                {
                    licenzopen = "https://open-npf.ru/" + obj.GetAttribute("href");
                }
                //Уведомление о начале процедуры реорганизации
                string uvedomnachaloreorgopen = "";
                foreach (IDomObject obj in dom.Find(".row a:eq(27)"))
                {
                    uvedomnachaloreorgopen = "https://open-npf.ru/" + obj.GetAttribute("href");
                }
                //Уведомление о согласовании реорганизации
                string uvedomsoglreorgopen = "";
                foreach (IDomObject obj in dom.Find(".row a:eq(26)"))
                {
                    uvedomsoglreorgopen = "https://open-npf.ru/" + obj.GetAttribute("href");
                }
                //Отчетность
                string otchetopen = "";
                foreach (IDomObject obj in dom.Find(".row a:eq(1)"))
                {
                    otchetopen = "https://open-npf.ru/" + obj.GetAttribute("href");
                }
                //Телефон
                string telopen = "";
                url = "https://open-npf.ru/contacts/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".col-sm-4.col-xs-12 a:eq(0)"))
                {
                    telopen = obj.TextContent;
                }

         

            //МОСПРОМСТРОЙ ФОНД
            displayChange("СТРОЙПРОМ-ФОНД");
         
                //раздел с раскрытием на главной странице
                string raskrmpsfond = "";
                url = "http://www.mpsfond.ru/";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find(".menu:eq(3)"))
                {
                    raskrmpsfond = "http://www.mpsfond.ru/" + obj.GetAttribute("href");
                }
                //Адрес
                string adresmpsfond = "";
                foreach (IDomObject obj in dom.Find(".menu:eq(9)"))
                {
                    adresmpsfond = "http://www.mpsfond.ru/" + obj.GetAttribute("href");
                }
                //Устав
                string ustavmpsfond = "";
                url = "http://www.mpsfond.ru/documents.php";
                htmlCode = downloadPageUTF8(url);

                dom = CQ.CreateDocument(htmlCode);
                foreach (IDomObject obj in dom.Find("li:eq(0) a"))
                {
                    ustavmpsfond = "http://www.mpsfond.ru/" + obj.GetAttribute("href");
                }
                //Пенсионные правила
                string penspravmpsfond = "";
                foreach (IDomObject obj in dom.Find("li:eq(2) a"))
                {
                    penspravmpsfond = "http://www.mpsfond.ru/" + obj.GetAttribute("href");
                }
                //Регистрация изменений и дополнений в пенсионные правила
                string izmpenspravmpsfond = "";
                foreach (IDomObject obj in dom.Find("li:eq(3) a"))
                {
                    izmpenspravmpsfond = "http://www.mpsfond.ru/" + obj.GetAttribute("href");
                }
                //Лицензия
                string licenzmpsfond = "";
                foreach (IDomObject obj in dom.Find("li:eq(4) a"))
                {
                    licenzmpsfond = "http://www.mpsfond.ru/" + obj.GetAttribute("href");
                }
                //Структура и состав акционеров
                string akcionmpsfond = "";
                foreach (IDomObject obj in dom.Find("li:eq(10) a"))
                {
                    akcionmpsfond = "http://www.mpsfond.ru/" + obj.GetAttribute("href");
                }
                //Показатели деятельности
                string pokazatelmpsfond = "";
                foreach (IDomObject obj in dom.Find("p:eq(1) a"))
                {
                    pokazatelmpsfond = "http://www.mpsfond.ru/" + obj.GetAttribute("href");
                }


          








            //ЗАПОЛНЕНИЕ EXCEL ТАБЛИЦЫ
            string sberbank = "НПФ Сбербанк";
            string npfatom = "АтомГарант";
            string npfb = "НПФ Благосостояние";
            string npfsocium = "НПФ Социум";
            string evonpf = "НПФ Эволюция";
            string npfts = "Телеком-Союз";
            string npfsng = "Сургутнефтегаз";
            string npftransneft = "НПФ Транснефть";
            string npfalliance = "НПФ Альянс";
            string ppafond = "Первый промышленный альянс";
            string npfao = "Алмазная осень";
            string npfstroycomplex = "Стройкомплекс";
            string mnpfakvilon = "МНПФ АКВИЛОН";
            string apkfond = "АПК-Фонд";
            string doverie = "НПФ Достойное будущее";
            string hmnpf = "Ханты-Мансийский НПФ";
            string npfopf = "НПФ ОПФ";
            string bigpension = "МНПФ Большой";
            string npfgefest = "НПФ Гефест";
            string npfond = "НПФ УГМК-Перспектива";
            string federation = "НПФ Федерация";
            string npff = "НПФ Будущее";
            string npfaviapolis = "НПФ Авиаполис";
            string volgacapital = "НПФ Волга-Капитал";
            string gazfond = "ГАЗФОНД";
            string npfprof = "НПФ Профессиональный";
            string gpbf = "НПФ Газпромбанк-фонд";
            string npfveb = "НПФ Внешэкономфонд";
            string vtbnpf = "ВТБ Пенсионный фонд";
            string rostecnpf = "Ростех";
            string gazfondpn = "НПФ ГАЗФОНД ПН";
            string tradnpf = "Традиция";
            string mpsfond = "НПФ Стройпром-Фонд";
            string nnpf = "Национальный НПФ";
            string atomfond = "Атомфонд";
            string sberfond = "НПФ СБЕРФОНД";
            string doverie56 = "НПФ Доверие";
            string ingospensiya = "НПФ Ингосстрах-Пенсия";
            string opennpf = "НПФ открытие";
            string npfkorabelspb = "НПФ Корабел";


            string txt1 = "Наименование и номер лицензии фонда";
            string txt2 = "Место нахождения фонда и его обособленных подразделений";
            string txt3 = "Бухгалтерская (финансовая) отчетность фонда, аудиторское и актуарное заключения";
            string txt4 = "Структура и состав акционеров";
            string txt5 = "Результат размещения пенсионных резервов %";
            string txt6 = "Результат инвестирования пенсионных накоплений %";
            string txt7 = "Размер дохода от размещения пенсионных резервов, направляемого на формирование страхового резерва фонда";
            string txt8 = "Количество вкладчиков и участников фонда, а также участников фонда, получающих из фонда негосударственную пенсию";
            string txt9 = "Количество застрахованных лиц, осуществляющих формирование своих пенсионных накоплений в фонде";
            string txt10 = "Размер пенсионных резервов фонда, в том числе страхового резерва, пенсионных накоплений, в том числе резерва по обязательному пенсионному страхованию, выплатного резерва, средств застрахованных лиц, которым установлена срочная пенсионная выплата";
            string txt11 = "Заключение и прекращение действия договора доверительного управления пенсионными резервами или пенсионными накоплениями с управляющей компанией с указанием ее фирменного наименования и номера лицензии";
            string txt12 = "Заключение и прекращение договора со специализированным депозитарием";
            string txt13 = "Пенсионные правила";
            string txt14 = "Страховые правила";
            string txt15 = "Регистрация изменений и дополнений в пенсионные и страховые правила";
            string txt16 = "Фонд должен раскрывать сведения о принятии Банком России решения о запрете на проведение фондом всех или части операций (с указанием перечня таких операций, даты введения запрета и срока, на который введен запрет";
            string txt17 = "Отчет о формировании средств пенсионных накоплений";
            string txt18 = "Уведомление о начале процедуры реорганизации";
            string txt19 = "Решение о согласовании проведения реорганизации";
            string txt20 = "Информация о конечных владельцах фонда (в соотв. с Указанием БР  № 441-П)";
            string txt21 = "Отчетность МСФО";
            string txt22 = "Номер телефона";
            string txt23 = "Устав (5175-У)";
            string txt24 = "Сведения об органах управления, членах совета директоров, должностных лиц и работников фонда (5175-У)";
            string txt25 = "Средневзвешенный процент, на который был увеличен размер назначенных негосударственных пенсий по итогам размещения средств пенсионных резервов за отчетный год (5175-У)";
            string txt26 = "Структура инвестиционного портфеля фонда (средств пенсионных резервов фонда) с указанием долей, приходящихся на виды активов (5175-У)";
            string txt27 = "Информация о составе инвестиционного портфеля фонда по обязательному пенсионному страхованию, а также информацию о составе средств пенсионных резервов фонда (5175-У)";
            string txt28 = "Информация о событиях (действиях), оказывающих, по мнению фонда, существенное влияние на совокупную стоимость активов, в которые инвестированы средства пенсионных накоплений и размещены средства пенсионных резервов (5175-У)";
            string txt29 = "Главная (начальная) страница сайта должна содержать раздел с информацией, подлежащей раскрытию (5175-У)";
            string txt30 = "Раскрытие информации о процессе инвестирования средств пенсионных накоплений и размещения средств пенсионных резервов";




            string fileName = "D:\\temp\\test.xlsx";

            try
            {
                var excel = new Excel.Application();

                var workBooks = excel.Workbooks;
                var workBook = workBooks.Add();
                var workSheet = (Excel.Worksheet)excel.ActiveSheet;
                workSheet.StandardWidth = 30;
                workSheet.Cells.RowHeight = 100;
                workSheet.Cells[1, 1].EntireRow.Font.Bold = true;
                workSheet.Cells[10, 1].EntireColumn.Font.Bold = true;
                workSheet.Cells[2, "A"] = sberbank;
                workSheet.Cells[3, "A"] = npfatom;
                workSheet.Cells[4, "A"] = npfb;
                workSheet.Cells[5, "A"] = npfsocium;
                workSheet.Cells[6, "A"] = evonpf;
                workSheet.Cells[7, "A"] = npfts;
                workSheet.Cells[8, "A"] = npfsng;
                workSheet.Cells[9, "A"] = npftransneft;
                workSheet.Cells[10, "A"] = npfalliance;
                workSheet.Cells[11, "A"] = ppafond;
                workSheet.Cells[12, "A"] = npfao;
                workSheet.Cells[13, "A"] = npfstroycomplex;
                workSheet.Cells[14, "A"] = mnpfakvilon;
                workSheet.Cells[15, "A"] = apkfond;
                workSheet.Cells[16, "A"] = doverie;
                workSheet.Cells[17, "A"] = hmnpf;
                workSheet.Cells[18, "A"] = npfopf;
                workSheet.Cells[19, "A"] = bigpension;
                workSheet.Cells[20, "A"] = npfgefest;
                workSheet.Cells[21, "A"] = npfond;
                workSheet.Cells[22, "A"] = federation;
                workSheet.Cells[23, "A"] = npff;
                workSheet.Cells[24, "A"] = npfaviapolis;
                workSheet.Cells[25, "A"] = volgacapital;
                workSheet.Cells[26, "A"] = gazfond;
                workSheet.Cells[27, "A"] = npfprof;
                workSheet.Cells[28, "A"] = gpbf;
                workSheet.Cells[29, "A"] = npfveb;
                workSheet.Cells[30, "A"] = vtbnpf;
                workSheet.Cells[31, "A"] = rostecnpf;
                workSheet.Cells[32, "A"] = gazfondpn;
                workSheet.Cells[33, "A"] = tradnpf;
                workSheet.Cells[34, "A"] = mpsfond;
                workSheet.Cells[35, "A"] = nnpf;
                workSheet.Cells[36, "A"] = atomfond;
                workSheet.Cells[37, "A"] = sberfond;
                workSheet.Cells[38, "A"] = doverie56;
                workSheet.Cells[39, "A"] = ingospensiya;
                workSheet.Cells[40, "A"] = opennpf;
                workSheet.Cells[41, "A"] = npfkorabelspb;




                workSheet.Cells[1, "B"] = txt1;
                workSheet.Cells[1, "C"] = txt2;
                workSheet.Cells[1, "D"] = txt3;
                workSheet.Cells[1, "E"] = txt4;
                workSheet.Cells[1, "F"] = txt5;
                workSheet.Cells[1, "G"] = txt6;
                workSheet.Cells[1, "H"] = txt7;
                workSheet.Cells[1, "I"] = txt8;
                workSheet.Cells[1, "J"] = txt9;
                workSheet.Cells[1, "K"] = txt10;
                workSheet.Cells[1, "L"] = txt11;
                workSheet.Cells[1, "M"] = txt12;
                workSheet.Cells[1, "N"] = txt13;
                workSheet.Cells[1, "O"] = txt14;
                workSheet.Cells[1, "P"] = txt15;
                workSheet.Cells[1, "Q"] = txt16;
                workSheet.Cells[1, "R"] = txt17;
                workSheet.Cells[1, "S"] = txt18;
                workSheet.Cells[1, "T"] = txt19;
                workSheet.Cells[1, "U"] = txt20;
                workSheet.Cells[1, "V"] = txt21;
                workSheet.Cells[1, "W"] = txt22;
                workSheet.Cells[1, "X"] = txt23;
                workSheet.Cells[1, "Y"] = txt24;
                workSheet.Cells[1, "Z"] = txt25;
                workSheet.Cells[1, "AA"] = txt26;
                workSheet.Cells[1, "AB"] = txt27;
                workSheet.Cells[1, "AC"] = txt28;
                workSheet.Cells[1, "AD"] = txt29;
                workSheet.Cells[1, "AE"] = txt30;

                //Фонды, которые не занимаются ПР
                workSheet.Cells[22, "H"].Interior.Color = Color.Gray;
                workSheet.Cells[22, "F"].Interior.Color = Color.Gray;
                workSheet.Cells[22, "I"].Interior.Color = Color.Gray;
                workSheet.Cells[22, "N"].Interior.Color = Color.Gray;
                workSheet.Cells[22, "P"].Interior.Color = Color.Gray;
                workSheet.Cells[22, "Z"].Interior.Color = Color.Gray;

                workSheet.Cells[36, "H"].Interior.Color = Color.Gray;
                workSheet.Cells[36, "F"].Interior.Color = Color.Gray;
                workSheet.Cells[36, "I"].Interior.Color = Color.Gray;
                workSheet.Cells[36, "N"].Interior.Color = Color.Gray;
                workSheet.Cells[36, "P"].Interior.Color = Color.Gray;
                workSheet.Cells[36, "Z"].Interior.Color = Color.Gray;

                //Не занимаются ПН
                workSheet.Cells[3, "G"].Interior.Color = Color.Gray;
                workSheet.Cells[3, "J"].Interior.Color = Color.Gray;
                workSheet.Cells[3, "O"].Interior.Color = Color.Gray;
                workSheet.Cells[3, "R"].Interior.Color = Color.Gray;

                workSheet.Cells[4, "G"].Interior.Color = Color.Gray;
                workSheet.Cells[4, "J"].Interior.Color = Color.Gray;
                workSheet.Cells[4, "O"].Interior.Color = Color.Gray;
                workSheet.Cells[4, "R"].Interior.Color = Color.Gray;

                workSheet.Cells[15, "G"].Interior.Color = Color.Gray;
                workSheet.Cells[15, "J"].Interior.Color = Color.Gray;
                workSheet.Cells[15, "O"].Interior.Color = Color.Gray;
                workSheet.Cells[15, "R"].Interior.Color = Color.Gray;

                workSheet.Cells[24, "G"].Interior.Color = Color.Gray;
                workSheet.Cells[24, "J"].Interior.Color = Color.Gray;
                workSheet.Cells[24, "O"].Interior.Color = Color.Gray;
                workSheet.Cells[24, "R"].Interior.Color = Color.Gray;

                workSheet.Cells[26, "G"].Interior.Color = Color.Gray;
                workSheet.Cells[26, "J"].Interior.Color = Color.Gray;
                workSheet.Cells[26, "O"].Interior.Color = Color.Gray;
                workSheet.Cells[26, "R"].Interior.Color = Color.Gray;

                workSheet.Cells[28, "G"].Interior.Color = Color.Gray;
                workSheet.Cells[28, "J"].Interior.Color = Color.Gray;
                workSheet.Cells[28, "O"].Interior.Color = Color.Gray;
                workSheet.Cells[28, "R"].Interior.Color = Color.Gray;

                workSheet.Cells[29, "G"].Interior.Color = Color.Gray;
                workSheet.Cells[29, "J"].Interior.Color = Color.Gray;
                workSheet.Cells[29, "O"].Interior.Color = Color.Gray;
                workSheet.Cells[29, "R"].Interior.Color = Color.Gray;

                workSheet.Cells[33, "G"].Interior.Color = Color.Gray;
                workSheet.Cells[33, "J"].Interior.Color = Color.Gray;
                workSheet.Cells[33, "O"].Interior.Color = Color.Gray;
                workSheet.Cells[33, "R"].Interior.Color = Color.Gray;

                workSheet.Cells[34, "G"].Interior.Color = Color.Gray;
                workSheet.Cells[34, "J"].Interior.Color = Color.Gray;
                workSheet.Cells[34, "O"].Interior.Color = Color.Gray;
                workSheet.Cells[34, "R"].Interior.Color = Color.Gray;

                workSheet.Cells[37, "G"].Interior.Color = Color.Gray;
                workSheet.Cells[37, "J"].Interior.Color = Color.Gray;
                workSheet.Cells[37, "O"].Interior.Color = Color.Gray;
                workSheet.Cells[37, "R"].Interior.Color = Color.Gray;

                workSheet.Cells[39, "G"].Interior.Color = Color.Gray;
                workSheet.Cells[39, "J"].Interior.Color = Color.Gray;
                workSheet.Cells[39, "O"].Interior.Color = Color.Gray;
                workSheet.Cells[39, "R"].Interior.Color = Color.Gray;

                workSheet.Cells[41, "G"].Interior.Color = Color.Gray;
                workSheet.Cells[41, "J"].Interior.Color = Color.Gray;
                workSheet.Cells[41, "O"].Interior.Color = Color.Gray;
                workSheet.Cells[41, "R"].Interior.Color = Color.Gray;




                //НПФ АВИАПОЛИС

                workSheet.Cells[24, "AD"] = raskraviapolis;
                workSheet.Cells[24, "E"] = akcionaviapolis;
                workSheet.Cells[24, "U"] = akcionaviapolis;
                workSheet.Cells[24, "L"] = UKaviapolis;
                workSheet.Cells[24, "M"] = SDaviapolis;
                workSheet.Cells[24, "D"] = otchetaviapolis;
                workSheet.Cells[24, "V"] = otchetaviapolis;
                workSheet.Cells[24, "B"] = docaviapolis;
                workSheet.Cells[24, "X"] = docaviapolis;
                workSheet.Cells[24, "N"] = docaviapolis;
                workSheet.Cells[24, "P"] = docaviapolis;
                workSheet.Cells[24, "I"] = pokazatelaviapolis;
                workSheet.Cells[24, "K"] = pokazatelaviapolis;
                workSheet.Cells[24, "F"] = pokazatelaviapolis;
                workSheet.Cells[24, "H"] = pokazatelaviapolis;
                workSheet.Cells[24, "Z"] = pokazatelaviapolis;
                workSheet.Cells[24, "AB"] = sostavportfaviapolis;
                workSheet.Cells[24, "AA"] = structuraportfaviapolis;
                workSheet.Cells[24, "Y"] = orgupraviapolis;
                workSheet.Cells[24, "S"] = reorgaviapolis;
                workSheet.Cells[24, "T"] = reorgaviapolis;
                workSheet.Cells[24, "C"] = adresaviapolis;
                workSheet.Cells[24, "AE"] = orgupraviapolis;














                //НПФ Сбербанк
                workSheet.Cells[2, "W"] = telnpfsberbanka;
                workSheet.Cells[2, "X"] = ustavnpfsberbanka;
                workSheet.Cells[2, "O"] = strahovpravilanpfsberbanka;
                workSheet.Cells[2, "P"] = izmenpravilnpfsberbanka;
                workSheet.Cells[2, "N"] = pensionpravilnpfsberbanka;
                workSheet.Cells[2, "B"] = licenznpfsberbanka;
                workSheet.Cells[2, "E"] = struktsostavnpfsberbanka;
                workSheet.Cells[2, "U"] = struktsostavnpfsberbanka;
                workSheet.Cells[2, "AD"] = razdelraskrnpfsberbanka;
                workSheet.Cells[2, "V"] = otchetnostnpfsber;
                workSheet.Cells[2, "D"] = otchetnostnpfsber;
                workSheet.Cells[2, "R"] = otchetnostnpfsber;
                workSheet.Cells[2, "F"] = rezuldeytnpfsberbanka;
                workSheet.Cells[2, "G"] = rezuldeytnpfsberbanka;
                workSheet.Cells[2, "I"] = rezuldeytnpfsberbanka;
                workSheet.Cells[2, "J"] = rezuldeytnpfsberbanka;
                workSheet.Cells[2, "H"] = rezuldeytnpfsberbanka;
                workSheet.Cells[2, "K"] = rezuldeytnpfsberbanka;
                workSheet.Cells[2, "Z"] = rezuldeytnpfsberbanka;
                workSheet.Cells[2, "L"] = rastordogovUKnpfsber;
                workSheet.Cells[2, "M"] = rastordogovSDnpfsber;
                workSheet.Cells[2, "AA"] = rastordogovUKnpfsber;
                workSheet.Cells[2, "AB"] = rastordogovUKnpfsber;
                workSheet.Cells[2, "C"] = addressnfsberbanka;
                workSheet.Cells[2, "Y"] = rukovodnpfsber;
                workSheet.Cells[2, "AE"] = processinvestsberbank;



                //АтомГарант
                workSheet.Cells[3, "W"] = telnpfatom;
                workSheet.Cells[3, "C"] = addressnpfatom;
                workSheet.Cells[3, "X"] = ustavnpfatom;
                workSheet.Cells[3, "E"] = sostavnpfatom;
                workSheet.Cells[3, "N"] = penspravilanpfatom;
                workSheet.Cells[3, "B"] = licenznpfatom;
                workSheet.Cells[3, "Y"] = orguprnpfatom;
                workSheet.Cells[3, "F"] = rezrazmPRnpfatom;
                workSheet.Cells[3, "Z"] = srvzvprocnpfatom;
                workSheet.Cells[3, "AA"] = structinvestportnpfatom;
                workSheet.Cells[3, "Q"] = reshozapretenpfatom;
                workSheet.Cells[3, "D"] = finotchetnpfatom;
                workSheet.Cells[3, "H"] = stahovrezervnpfatom;
                workSheet.Cells[3, "I"] = kolvkluchas;
                workSheet.Cells[3, "K"] = razmerprnpfatom;
                workSheet.Cells[3, "V"] = finotchetnpfatom;
                workSheet.Cells[3, "P"] = izmendopvpenspravnpfatom;
                workSheet.Cells[3, "L"] = prkrdogovUKnpfatom;
                workSheet.Cells[3, "AB"] = sostavsredstvPRnpfatom;
                workSheet.Cells[3, "U"] = infkonechvladnpfatom;
                workSheet.Cells[3, "AD"] = razdelraskrnpfatom;
                workSheet.Cells[3, "M"] = specdepoznpfatom;
                workSheet.Cells[3, "AE"] = procesrazmPRnpfatom;



                //НПФ БЛАГОСОСТОЯНИЕ
                workSheet.Cells[4, "AD"] = razdelraskrnpfb;
                workSheet.Cells[4, "B"] = licenznpfb;
                workSheet.Cells[4, "N"] = penspravnpfb;
                workSheet.Cells[4, "W"] = telnpfb;
                workSheet.Cells[4, "X"] = ustavnpfb;
                workSheet.Cells[4, "P"] = arhpenspravnpfb;
                workSheet.Cells[4, "C"] = adresnpfb;
                workSheet.Cells[4, "D"] = razdelraskrnpfb;
                workSheet.Cells[4, "E"] = razdelraskrnpfb;
                workSheet.Cells[4, "F"] = razdelraskrnpfb;
                workSheet.Cells[4, "H"] = razdelraskrnpfb;
                workSheet.Cells[4, "I"] = kolvovklnpfb;
                workSheet.Cells[4, "K"] = kolvovklnpfb;
                workSheet.Cells[4, "L"] = razdelraskrnpfb;
                workSheet.Cells[4, "M"] = razdelraskrnpfb;
                workSheet.Cells[4, "U"] = razdelraskrnpfb;
                workSheet.Cells[4, "V"] = razdelraskrnpfb;
                workSheet.Cells[4, "Y"] = razdelraskrnpfb;
                workSheet.Cells[4, "Z"] = razdelraskrnpfb;
                workSheet.Cells[4, "AA"] = razdelraskrnpfb;
                workSheet.Cells[4, "AB"] = razdelraskrnpfb;
                workSheet.Cells[4, "AE"] = razdelraskrnpfb;





                //НПФ СОЦИУМ
                workSheet.Cells[5, "AD"] = raskrytsocium;
                workSheet.Cells[5, "B"] = licenzsocium;
                workSheet.Cells[5, "X"] = ustavsocium;
                workSheet.Cells[5, "N"] = penspravsocium;
                workSheet.Cells[5, "O"] = strahpravsocium;
                workSheet.Cells[5, "D"] = docsocium;
                workSheet.Cells[5, "R"] = docsocium;
                workSheet.Cells[5, "I"] = docsocium;
                workSheet.Cells[5, "J"] = docsocium;
                workSheet.Cells[5, "K"] = docsocium;
                workSheet.Cells[5, "F"] = docsocium;
                workSheet.Cells[5, "G"] = docsocium;
                workSheet.Cells[5, "Z"] = docsocium;
                workSheet.Cells[5, "H"] = docsocium;
                workSheet.Cells[5, "P"] = docsocium;
                workSheet.Cells[5, "AA"] = investsocium;
                workSheet.Cells[5, "AB"] = investsocium;
                workSheet.Cells[5, "L"] = investsocium;
                workSheet.Cells[5, "M"] = investsocium;
                workSheet.Cells[5, "Y"] = orguprsocium;
                workSheet.Cells[5, "V"] = docsocium;
                workSheet.Cells[5, "C"] = adressocium;
                workSheet.Cells[5, "W"] = telsocium;
                workSheet.Cells[5, "AE"] = docsocium;



                //НПФ ЭВОЛЮЦИЯ
                workSheet.Cells[6, "AD"] = razdelraskrevonpf;
                workSheet.Cells[6, "W"] = televonpf;
                workSheet.Cells[6, "X"] = ustavevonpf;
                workSheet.Cells[6, "B"] = ustavevonpf;
                workSheet.Cells[6, "S"] = ustavevonpf;
                workSheet.Cells[6, "T"] = ustavevonpf;
                workSheet.Cells[6, "N"] = ustavevonpf;
                workSheet.Cells[6, "O"] = ustavevonpf;
                workSheet.Cells[6, "P"] = ustavevonpf;
                workSheet.Cells[6, "D"] = finotchetevonpf;
                workSheet.Cells[6, "V"] = finotchetevonpf;
                workSheet.Cells[6, "R"] = finotchetevonpf;
                workSheet.Cells[6, "Z"] = finotchetevonpf;
                workSheet.Cells[6, "Y"] = orguprevonpf;
                workSheet.Cells[6, "M"] = specdepozevonpf;
                workSheet.Cells[6, "E"] = sostavevonpf;
                workSheet.Cells[6, "U"] = sostavevonpf;
                workSheet.Cells[6, "F"] = investirevonpf;
                workSheet.Cells[6, "G"] = investirevonpf;
                workSheet.Cells[6, "AA"] = investirevonpf;
                workSheet.Cells[6, "L"] = UKevonpf;
                workSheet.Cells[6, "H"] = dinamikaevonpf;
                workSheet.Cells[6, "I"] = dinamikaevonpf;
                workSheet.Cells[6, "J"] = dinamikaevonpf;
                workSheet.Cells[6, "K"] = dinamikaevonpf;
                workSheet.Cells[6, "C"] = adressnpfb;
                workSheet.Cells[6, "AB"] = sostavportfevonpf;
                workSheet.Cells[6, "AE"] = procinvestevonpf;


                //НПФ ТЕЛЕКОМ-СОЮЗ
                workSheet.Cells[7, "AD"] = razdelraskrnpfts;
                workSheet.Cells[7, "Y"] = razdelraskrnpfts;
                workSheet.Cells[7, "W"] = telnpfts;
                workSheet.Cells[7, "C"] = adrests;
                workSheet.Cells[7, "L"] = UKnpfts;
                workSheet.Cells[7, "M"] = SDnpfts;
                workSheet.Cells[7, "D"] = otchetnpfts;
                workSheet.Cells[7, "V"] = otchetnpfts;
                workSheet.Cells[7, "K"] = pokazatelnpfts;
                workSheet.Cells[7, "J"] = pokazatelnpfts;
                workSheet.Cells[7, "I"] = pokazatelnpfts;
                workSheet.Cells[7, "Z"] = pokazatelnpfts;
                workSheet.Cells[7, "F"] = pokazatelnpfts;
                workSheet.Cells[7, "G"] = pokazatelnpfts;
                workSheet.Cells[7, "H"] = pokazatelnpfts;
                workSheet.Cells[7, "R"] = pokazatelnpfts;
                workSheet.Cells[7, "AA"] = investportfnpfts;
                workSheet.Cells[7, "AB"] = investportfnpfts;
                workSheet.Cells[7, "X"] = ustavnpfts;
                workSheet.Cells[7, "E"] = sostavakcionernpfts;
                workSheet.Cells[7, "U"] = sostavakcionernpfts;
                workSheet.Cells[7, "O"] = strahpravnpfts;
                workSheet.Cells[7, "N"] = penspravnpfts;
                workSheet.Cells[7, "B"] = licenznpfts;
                workSheet.Cells[7, "AC"] = sobytnpfts;
                workSheet.Cells[7, "P"] = doppenspravnpfts;
                workSheet.Cells[7, "AE"] = procinvestts;



                //НПФ СУРГУТНЕФТЕГАЗ
                workSheet.Cells[8, "AD"] = raskrytnpfsng;
                workSheet.Cells[8, "B"] = licenznpfsng;
                workSheet.Cells[8, "C"] = adressnpfsng;
                workSheet.Cells[8, "X"] = ustavnpfsng;
                workSheet.Cells[8, "N"] = penspravnpfsng;
                workSheet.Cells[8, "P"] = penspravnpfsng;
                workSheet.Cells[8, "O"] = strahpravnpfsng;
                workSheet.Cells[8, "D"] = dlyaYacheykiD;
                workSheet.Cells[8, "V"] = MSFOnpfsng;
                workSheet.Cells[8, "E"] = spisokakcionnpfsng;
                workSheet.Cells[8, "U"] = spisokakcionnpfsng;
                workSheet.Cells[8, "J"] = pokazatelnpfsng;
                workSheet.Cells[8, "I"] = pokazatelnpfsng;
                workSheet.Cells[8, "H"] = pokazatelnpfsng;
                workSheet.Cells[8, "K"] = pokazatelnpfsng;
                workSheet.Cells[8, "R"] = pokazatelnpfsng;
                workSheet.Cells[8, "M"] = SDnpfsng;
                workSheet.Cells[8, "AC"] = sobytnpfsng;
                workSheet.Cells[8, "Y"] = orguprnpfsng;
                workSheet.Cells[8, "AB"] = investportfelnpfsng;
                workSheet.Cells[8, "W"] = telnpfsng;
                workSheet.Cells[8, "F"] = rezOPSnpfsng;
                workSheet.Cells[8, "Z"] = rezOPSnpfsng;
                workSheet.Cells[8, "AA"] = rezOPSnpfsng;
                workSheet.Cells[8, "G"] = srvzvprocnpfsng;
                workSheet.Cells[8, "L"] = UKnpfsng;
                workSheet.Cells[8, "Q"] = zapretnpfsng;
                workSheet.Cells[8, "AE"] = procinvestnpfsng;



                //НПФ ТРАНСНЕФТЬ 
                workSheet.Cells[9, "AD"] = raskrytnpftransneft;
                workSheet.Cells[9, "X"] = ustavtransneft;
                workSheet.Cells[9, "C"] = adresstransneft;
                workSheet.Cells[9, "W"] = teltransneft;
                workSheet.Cells[9, "B"] = licenztransneft;
                workSheet.Cells[9, "O"] = strahpravtransneft;
                workSheet.Cells[9, "N"] = penspravtransneft;
                workSheet.Cells[9, "P"] = penspravtransneft;
                workSheet.Cells[9, "L"] = UKtransneft;
                workSheet.Cells[9, "M"] = SDtransneft;
                workSheet.Cells[9, "Y"] = orguprtransneft;
                workSheet.Cells[9, "Z"] = pokazateltransneft;
                workSheet.Cells[9, "I"] = pokazateltransneft;
                workSheet.Cells[9, "J"] = pokazateltransneft;
                workSheet.Cells[9, "K"] = pokazateltransneft;
                workSheet.Cells[9, "F"] = pokazateltransneft;
                workSheet.Cells[9, "G"] = pokazateltransneft;
                workSheet.Cells[9, "H"] = pokazateltransneft;
                workSheet.Cells[9, "AA"] = portftransneft;
                workSheet.Cells[9, "AB"] = portftransneft;
                workSheet.Cells[9, "AE"] = portftransneft;
                workSheet.Cells[9, "AC"] = inftransneft;
                workSheet.Cells[9, "D"] = otchettransneft;
                workSheet.Cells[9, "V"] = otchettransneft;
                workSheet.Cells[9, "R"] = otchettransneft;
                workSheet.Cells[9, "E"] = sostavakctransneft;
                workSheet.Cells[9, "U"] = sostavakctransneft;


                //НПФ АЛЬЯНС 
                workSheet.Cells[10, "AD"] = raskrallians;
                workSheet.Cells[10, "E"] = sostavakcionallians;
                workSheet.Cells[10, "U"] = sostavakcionallians;
                workSheet.Cells[10, "D"] = othetalliance;
                workSheet.Cells[10, "R"] = othetalliance;
                workSheet.Cells[10, "V"] = othetalliance;
                workSheet.Cells[10, "I"] = pokazatelalliance;
                workSheet.Cells[10, "J"] = pokazatelalliance;
                workSheet.Cells[10, "K"] = pokazatelalliance;
                workSheet.Cells[10, "H"] = pokazatelalliance;
                workSheet.Cells[10, "F"] = pokazatelalliance;
                workSheet.Cells[10, "G"] = pokazatelalliance;
                workSheet.Cells[10, "Z"] = pokazatelalliance;
                workSheet.Cells[10, "B"] = licenzalliance;
                workSheet.Cells[10, "W"] = telalliance;
                workSheet.Cells[10, "X"] = ustavalliance;
                workSheet.Cells[10, "L"] = UKalliance;
                workSheet.Cells[10, "M"] = SDalliance;
                workSheet.Cells[10, "N"] = penspravalliance;
                workSheet.Cells[10, "O"] = strahpravalliance;
                workSheet.Cells[10, "AA"] = strportfalliance;
                workSheet.Cells[10, "AB"] = strportfalliance;
                workSheet.Cells[10, "Y"] = rukovodalliance;
                workSheet.Cells[10, "C"] = adressalliance;



                //ПЕРВЫЙ ПРОМЫШЛЕННЫЙ АЛЬЯНС
                workSheet.Cells[11, "Y"] = rukovodppafond;
                workSheet.Cells[11, "X"] = ustavppafond;
                workSheet.Cells[11, "B"] = licenzppafond;
                workSheet.Cells[11, "N"] = penspravppafond;
                workSheet.Cells[11, "O"] = strahpravppafond;
                workSheet.Cells[11, "C"] = adresppafond;
                workSheet.Cells[11, "AA"] = portfppafond;
                workSheet.Cells[11, "AB"] = portfppafond;
                workSheet.Cells[11, "H"] = pokazatelppafond;
                workSheet.Cells[11, "I"] = pokazatelppafond;
                workSheet.Cells[11, "J"] = pokazatelppafond;
                workSheet.Cells[11, "K"] = pokazatelppafond;
                workSheet.Cells[11, "F"] = pokazatelppafond;
                workSheet.Cells[11, "G"] = pokazatelppafond;
                workSheet.Cells[11, "Z"] = pokazatelppafond;
                workSheet.Cells[11, "D"] = otchetppafond;
                workSheet.Cells[11, "R"] = otchetppafond;
                workSheet.Cells[11, "V"] = otchetppafond;
                workSheet.Cells[11, "L"] = UKppafond;
                workSheet.Cells[11, "M"] = SDppafond;
                workSheet.Cells[11, "E"] = akcionerppafond;
                workSheet.Cells[11, "U"] = akcionerppafond;
                workSheet.Cells[11, "W"] = telppafond;
                workSheet.Cells[11, "AC"] = portfppafond;
                workSheet.Cells[11, "AE"] = portfppafond;



                //НПФ АЛМАЗНАЯ ОСЕНЬ
                workSheet.Cells[12, "AD"] = raskrnpfao;
                workSheet.Cells[12, "B"] = licenznpfao;
                workSheet.Cells[12, "X"] = ustavnpfao;
                workSheet.Cells[12, "N"] = pensstrahpravnpfao;
                workSheet.Cells[12, "O"] = pensstrahpravnpfao;
                workSheet.Cells[12, "P"] = pensstrahpravnpfao;
                workSheet.Cells[12, "D"] = otchetnpfao;
                workSheet.Cells[12, "V"] = otchetnpfao;
                workSheet.Cells[12, "R"] = otchetnpfao;
                workSheet.Cells[12, "L"] = UKSDnpfao;
                workSheet.Cells[12, "M"] = UKSDnpfao;
                workSheet.Cells[12, "C"] = adresnpfao;
                workSheet.Cells[12, "W"] = telnpfao;
                workSheet.Cells[12, "E"] = sostavakcionernpfao;
                workSheet.Cells[12, "U"] = sostavakcionernpfao;
                workSheet.Cells[12, "Y"] = sostavakcionernpfao;
                workSheet.Cells[12, "F"] = investPRnpfao;
                workSheet.Cells[12, "H"] = razmerdohodaPRnpfao;
                workSheet.Cells[12, "G"] = investPNnpfao;
                workSheet.Cells[12, "I"] = uchastnpfao;
                workSheet.Cells[12, "J"] = zastrahlicnpfao;
                workSheet.Cells[12, "K"] = PRnpfao;
                workSheet.Cells[12, "Q"] = zaprnpfao;
                workSheet.Cells[12, "Z"] = srvzvprocnpfao;
                workSheet.Cells[12, "AA"] = structuraPRPNnpfao;
                workSheet.Cells[12, "AB"] = sostavPRPNnpfao;
                workSheet.Cells[12, "AC"] = infsobnpfao;
                workSheet.Cells[12, "AE"] = procinvestPRPNnpfao;




                //НПФ СТРОЙКОМПЛЕКС
                workSheet.Cells[13, "AD"] = raskrstroycomplex;
                workSheet.Cells[13, "B"] = licenzstroykomplex;
                workSheet.Cells[13, "C"] = adresstroykomplex;
                workSheet.Cells[13, "D"] = otchetstroykomplex;
                workSheet.Cells[13, "V"] = otchetstroykomplex;
                workSheet.Cells[13, "E"] = akcionstroykomplex;
                workSheet.Cells[13, "U"] = akcionstroykomplex;
                workSheet.Cells[13, "F"] = rezinvestPRstroykomplex;
                workSheet.Cells[13, "G"] = rezinvestPNstroykomplex;
                workSheet.Cells[13, "H"] = razmerdohodaPRstroykomplex;
                workSheet.Cells[13, "I"] = koluchstroykomplex;
                workSheet.Cells[13, "J"] = kolzastrstroykomplex;
                workSheet.Cells[13, "K"] = razmerPRstroykomplex;
                workSheet.Cells[13, "L"] = UKstroykomplex;
                workSheet.Cells[13, "M"] = SDstroykomplex;
                workSheet.Cells[13, "N"] = penspravstroykomplex;
                workSheet.Cells[13, "O"] = strahpravstroykomplex;
                workSheet.Cells[13, "X"] = ustavstroykomplex;
                workSheet.Cells[13, "Y"] = orguprstroykomplex;
                workSheet.Cells[13, "Z"] = srvzvprocstroykomplex;
                workSheet.Cells[13, "AA"] = structuraPRPNstroykomplex;
                workSheet.Cells[13, "AB"] = sostavportfstroykomplex;
                workSheet.Cells[13, "AC"] = sobytstroykomplex;
                workSheet.Cells[13, "Q"] = zaprstroykomplex;
                workSheet.Cells[13, "AE"] = procinvestsroykomplex;




                //МНПФ АКВИЛОН 
                workSheet.Cells[14, "AD"] = raskrakvilon;
                workSheet.Cells[14, "W"] = telakvilon;
                workSheet.Cells[14, "C"] = adressakvilon;
                workSheet.Cells[14, "E"] = spisokakcionakvilon;
                workSheet.Cells[14, "U"] = spisokakcionakvilon;
                workSheet.Cells[14, "L"] = UKakvilon;
                workSheet.Cells[14, "M"] = SDakvilon;
                workSheet.Cells[14, "X"] = ofdocakvilon;
                workSheet.Cells[14, "AA"] = ofdocakvilon;
                workSheet.Cells[14, "AB"] = ofdocakvilon;
                workSheet.Cells[14, "AE"] = ofdocakvilon;
                workSheet.Cells[14, "B"] = ofdocakvilon;
                workSheet.Cells[14, "H"] = otchetakvilon;
                workSheet.Cells[14, "F"] = dohodakvilon;
                workSheet.Cells[14, "G"] = dohodakvilon;
                workSheet.Cells[14, "I"] = otchetakvilon;
                workSheet.Cells[14, "J"] = otchetakvilon;
                workSheet.Cells[14, "K"] = otchetakvilon;
                workSheet.Cells[14, "D"] = otchetakvilon;
                workSheet.Cells[14, "N"] = ofdocakvilon;
                workSheet.Cells[14, "P"] = ofdocakvilon;
                workSheet.Cells[14, "R"].Interior.Color = Color.Red;
                workSheet.Cells[14, "O"].Interior.Color = Color.Red;
                workSheet.Cells[14, "Y"].Interior.Color = Color.Red;
                workSheet.Cells[14, "V"].Interior.Color = Color.Red;
                workSheet.Cells[14, "Z"].Interior.Color = Color.Red;




                //АПК ФОНД
                workSheet.Cells[15, "AD"] = raskrapkfond;
                workSheet.Cells[15, "Y"] = raskrapkfond;
                workSheet.Cells[15, "P"] = raskrapkfond;
                workSheet.Cells[15, "X"] = ustavapkfond;
                workSheet.Cells[15, "N"] = penspravapkfond;
                workSheet.Cells[15, "L"] = UKSDapkfond;
                workSheet.Cells[15, "M"] = UKSDapkfond;
                workSheet.Cells[15, "C"] = contactapkfond;
                workSheet.Cells[15, "W"] = contactapkfond;
                workSheet.Cells[15, "F"] = pokazatelapkfond;
                workSheet.Cells[15, "AB"] = sostavPRapkfond;
                workSheet.Cells[15, "AA"] = structuraPRapkfond;
                workSheet.Cells[15, "Z"] = raskrapkfond;
                workSheet.Cells[15, "D"] = otchetapkfond;
                workSheet.Cells[15, "V"] = otchetapkfond;
                workSheet.Cells[15, "E"] = structuraakcionapkfond;
                workSheet.Cells[15, "U"] = structuraakcionapkfond;
                workSheet.Cells[15, "K"] = pokazatelapkfond;
                workSheet.Cells[15, "I"] = pokazatelapkfond;
                workSheet.Cells[15, "H"] = pokazatelapkfond;
                workSheet.Cells[15, "B"] = licenzapkfond;
                workSheet.Cells[15, "F"] = rezrazmPRapkfond;
                workSheet.Cells[15, "AE"] = procinvestapkfond;




                //НПФ Достойное будущее
                workSheet.Cells[16, "AD"] = raskrdfnpf;
                workSheet.Cells[16, "W"] = teldfnpf;
                workSheet.Cells[16, "B"] = licenzdfnpf;
                workSheet.Cells[16, "E"] = spisokakciondfnpf;
                workSheet.Cells[16, "X"] = ustavdfnpf;
                workSheet.Cells[16, "Y"] = orguprdfnpf;
                workSheet.Cells[16, "U"] = vladelecdfnpf;
                workSheet.Cells[16, "S"] = reorgdfnpf;
                workSheet.Cells[16, "Q"] = zaprdfnpf;
                workSheet.Cells[16, "N"] = penspravfdnpf;
                workSheet.Cells[16, "O"] = strahpravdfnpf;
                workSheet.Cells[16, "D"] = otchetdfnpf;
                workSheet.Cells[16, "R"] = formirpndfnpf;
                workSheet.Cells[16, "V"] = MSFOdfnpf;
                workSheet.Cells[16, "C"] = adressdfnpf;
                workSheet.Cells[16, "AA"] = raskrdfnpf;
                workSheet.Cells[16, "AB"] = raskrdfnpf;
                workSheet.Cells[16, "K"] = raskrdfnpf;
                workSheet.Cells[16, "F"] = raskrdfnpf;
                workSheet.Cells[16, "G"] = raskrdfnpf;
                workSheet.Cells[16, "I"] = raskrdfnpf;
                workSheet.Cells[16, "J"] = raskrdfnpf;
                workSheet.Cells[16, "P"] = raskrdfnpf;
                workSheet.Cells[16, "Z"] = raskrdfnpf;
                workSheet.Cells[16, "H"] = raskrdfnpf;
                workSheet.Cells[16, "T"] = raskrdfnpf;
                workSheet.Cells[16, "L"] = UKdfnpf;
                workSheet.Cells[16, "M"] = SDdfnpf;
                workSheet.Cells[16, "AE"] = raskrdfnpf;






                //ХАНТЫ-МАНСИЙСКИЙ НПФ 
                workSheet.Cells[17, "AD"] = raskrhmnpf;
                workSheet.Cells[17, "B"] = licenzhmnpf;
                workSheet.Cells[17, "X"] = ustavhmnpf;
                workSheet.Cells[17, "E"] = structuraakcionhmnpf;
                workSheet.Cells[17, "U"] = structuraakcionhmnpf;
                workSheet.Cells[17, "C"] = raskrhmnpf;
                workSheet.Cells[17, "W"] = raskrhmnpf;
                workSheet.Cells[17, "D"] = otchethmnpf;
                workSheet.Cells[17, "R"] = otchethmnpf;
                workSheet.Cells[17, "V"] = otchethmnpf;
                workSheet.Cells[17, "Y"] = rucovodhmnpf;
                workSheet.Cells[17, "N"] = dochmnpf;
                workSheet.Cells[17, "O"] = dochmnpf;
                workSheet.Cells[17, "P"] = dochmnpf;
                workSheet.Cells[17, "I"] = pokazatelhmnpf;
                workSheet.Cells[17, "J"] = pokazatelhmnpf;
                workSheet.Cells[17, "K"] = pokazatelhmnpf;
                workSheet.Cells[17, "F"] = pokazatelhmnpf;
                workSheet.Cells[17, "G"] = pokazatelhmnpf;
                workSheet.Cells[17, "H"] = pokazatelhmnpf;
                workSheet.Cells[17, "Z"] = pokazatelhmnpf;
                workSheet.Cells[17, "AC"] = sobythmnpf;
                workSheet.Cells[17, "Q"] = zaprhmnpf;
                workSheet.Cells[17, "AA"] = investporthmnpf;
                workSheet.Cells[17, "AB"] = investporthmnpf;
                workSheet.Cells[17, "M"] = SDUKhmnpf;
                workSheet.Cells[17, "L"] = SDUKhmnpf;


                //НПФ ОПФ
                workSheet.Cells[18, "AD"] = raskropf;
                workSheet.Cells[18, "W"] = telopf;
                workSheet.Cells[18, "E"] = orgupropf;
                workSheet.Cells[18, "U"] = orgupropf;
                workSheet.Cells[18, "L"] = UKopf;
                workSheet.Cells[18, "M"] = SDopf;
                workSheet.Cells[18, "X"] = ustavopf;
                workSheet.Cells[18, "O"] = strahpravopf;
                workSheet.Cells[18, "N"] = penspravopf;
                workSheet.Cells[18, "P"] = penspravopf;
                workSheet.Cells[18, "D"] = otchetopf;
                workSheet.Cells[18, "I"] = srvzvprocopf;
                workSheet.Cells[18, "J"] = srvzvprocopf;
                workSheet.Cells[18, "F"] = pokazatelopf;
                workSheet.Cells[18, "G"] = pokazatelopf;
                workSheet.Cells[18, "H"] = pokazatelopf;
                workSheet.Cells[18, "K"] = pokazatelopf;
                workSheet.Cells[18, "Z"] = srvzvprocopf;
                workSheet.Cells[18, "V"] = actaudotchetopf;
                workSheet.Cells[18, "C"] = adressopf;
                workSheet.Cells[18, "R"] = formirPNopf;
                workSheet.Cells[18, "B"] = licenzopf;
                workSheet.Cells[18, "Y"] = rucovodopf;
                workSheet.Cells[18, "AA"] = portfopf;
                workSheet.Cells[18, "AB"] = portfopf;
                workSheet.Cells[18, "AE"] = procinvestopf;




                //МНПФ БОЛЬШОЙ 
                workSheet.Cells[19, "AD"] = raskrbigpension;
                workSheet.Cells[19, "AA"] = raskrbigpension;
                workSheet.Cells[19, "AB"] = raskrbigpension;
                workSheet.Cells[19, "Y"] = raskrbigpension;
                workSheet.Cells[19, "L"] = raskrbigpension;
                workSheet.Cells[19, "M"] = raskrbigpension;
                workSheet.Cells[19, "C"] = adresbigpension;
                workSheet.Cells[19, "S"] = rezinvestbigpension;
                workSheet.Cells[19, "D"] = buhotchetbigpension;
                workSheet.Cells[19, "V"] = audbigpension;
                workSheet.Cells[19, "I"] = vkladchbigpension;
                workSheet.Cells[19, "J"] = vkladchbigpension;
                workSheet.Cells[19, "X"] = ustavbigpension;
                workSheet.Cells[19, "N"] = penspravbigpension;
                workSheet.Cells[19, "O"] = spravbigpension;
                workSheet.Cells[19, "B"] = licenzbigpension;
                workSheet.Cells[19, "P"] = penspravbigpension + "\n" + spravbigpension;
                workSheet.Cells[19, "W"] = telbigpension;
                workSheet.Cells[19, "E"] = acionbigpension;
                workSheet.Cells[19, "F"] = rezinvestbigpension;
                workSheet.Cells[19, "G"] = rezinvestbigpension;
                workSheet.Cells[19, "H"] = sostavakcionbigpension;
                workSheet.Cells[19, "K"] = vkladchbigpension;
                workSheet.Cells[19, "R"] = otchetbigpension;
                workSheet.Cells[19, "U"] = acionbigpension;
                workSheet.Cells[19, "Z"] = indecsbigpension;
                workSheet.Cells[19, "AE"] = rezinvestbigpension;
                workSheet.Cells[19, "AC"] = "Не расккрыто, в связи с отсутствием данных событий";
                workSheet.Cells[19, "Q"] = "Не расккрыто, в связи с отсутствием данных событий";
                workSheet.Cells[19, "T"] = "Не расккрыто, в связи с отсутствием данных событий";






                //НПФ ГЕФЕСТ
                workSheet.Cells[20, "X"] = ustavgefest;
                workSheet.Cells[20, "Y"] = orguprgefest;
                workSheet.Cells[20, "W"] = telgefest;
                workSheet.Cells[20, "G"] = rezinvestPNgefest;
                workSheet.Cells[20, "F"] = rezinvestPRgefest;
                workSheet.Cells[20, "Z"] = srvzvprocgefest;
                workSheet.Cells[20, "AA"] = investportfgefest;
                workSheet.Cells[20, "AB"] = investportfgefest;
                workSheet.Cells[20, "AC"] = sobytgefest;
                workSheet.Cells[20, "Q"] = zaprgefest;
                workSheet.Cells[20, "L"] = UKSDgefest;
                workSheet.Cells[20, "M"] = UKSDgefest;
                workSheet.Cells[20, "N"] = pensstrahgefest;
                workSheet.Cells[20, "O"] = pensstrahgefest;
                workSheet.Cells[20, "P"] = pensstrahgefest;
                workSheet.Cells[20, "D"] = otchetgefest;
                workSheet.Cells[20, "V"] = otchetgefest;
                workSheet.Cells[20, "H"] = pokazatelgefest;
                workSheet.Cells[20, "I"] = pokazatelgefest;
                workSheet.Cells[20, "J"] = pokazatelgefest;
                workSheet.Cells[20, "K"] = pokazatelgefest;
                workSheet.Cells[20, "E"] = pokazatelgefest;
                workSheet.Cells[20, "U"] = pokazatelgefest;
                workSheet.Cells[20, "R"] = pokazatelgefest;
                workSheet.Cells[20, "B"] = licenzgefest;
                workSheet.Cells[20, "C"] = adressgefest;
                workSheet.Cells[20, "AE"] = procinvestgefest;
                workSheet.Cells[20, "AD"] = raskrgefest;



                //УГМК ПЕРСПЕКТИВА
                workSheet.Cells[21, "AD"] = raskrnpfond;
                workSheet.Cells[21, "E"] = sostavakcionnpfond;
                workSheet.Cells[21, "B"] = licenznpfond;
                workSheet.Cells[21, "D"] = buhotchetnpfond;
                workSheet.Cells[21, "V"] = buhotchetnpfond;
                workSheet.Cells[21, "R"] = buhotchetnpfond;
                workSheet.Cells[21, "F"] = pokazatelnpfond;
                workSheet.Cells[21, "G"] = pokazatelnpfond;
                workSheet.Cells[21, "H"] = pokazatelnpfond;
                workSheet.Cells[21, "I"] = pokazatelnpfond;
                workSheet.Cells[21, "J"] = pokazatelnpfond;
                workSheet.Cells[21, "K"] = pokazatelnpfond;
                workSheet.Cells[21, "S"] = pokazatelnpfond;
                workSheet.Cells[21, "Z"] = pokazatelnpfond;
                workSheet.Cells[21, "L"] = UKnpfond;
                workSheet.Cells[21, "M"] = SDnpfond;
                workSheet.Cells[21, "N"] = pravilanpfond;
                workSheet.Cells[21, "O"] = pravilanpfond;
                workSheet.Cells[21, "P"] = izmpravnpfond;
                workSheet.Cells[21, "X"] = ustavnpfond;
                workSheet.Cells[21, "Y"] = orguprnpfond;
                workSheet.Cells[21, "U"] = orguprnpfond;
                workSheet.Cells[21, "AA"] = investportfnpfond;
                workSheet.Cells[21, "AB"] = investportfnpfond;
                workSheet.Cells[21, "AE"] = investportfnpfond;
                workSheet.Cells[21, "Q"] = zaprnpfond;
                workSheet.Cells[21, "W"] = telnpfond;
                workSheet.Cells[21, "C"] = adressnpfond;
                workSheet.Cells[21, "AC"] = sobytnpfond;


                //НПФ ФЕДЕРАЦИЯ
                workSheet.Cells[22, "AD"] = raskrfederation;
                workSheet.Cells[22, "AA"] = portffederation;
                workSheet.Cells[22, "AB"] = portffederation;
                workSheet.Cells[22, "J"] = OPSfederation;
                workSheet.Cells[22, "K"] = OPSfederation;
                workSheet.Cells[22, "G"] = OPSfederation;
                workSheet.Cells[22, "AC"] = sobytfederation;
                workSheet.Cells[22, "D"] = otchetfederation;
                workSheet.Cells[22, "V"] = otchetfederation;
                workSheet.Cells[22, "R"] = otchetfederation;
                workSheet.Cells[22, "X"] = docfederation;
                workSheet.Cells[22, "O"] = docfederation;
                workSheet.Cells[22, "L"] = UKSDfederation;
                workSheet.Cells[22, "M"] = UKSDfederation;
                workSheet.Cells[22, "Y"] = orguprfederation;
                workSheet.Cells[22, "E"] = akcionfederation;
                workSheet.Cells[22, "U"] = akcionfederation;
                workSheet.Cells[22, "W"] = telfederation;
                workSheet.Cells[22, "C"] = adressfederation;
                workSheet.Cells[22, "B"] = licenzfederation;
                workSheet.Cells[22, "AE"] = procinvestfederation;




                //НПФ БУДУЩЕЕ
                workSheet.Cells[23, "AD"] = raskrnpff;
                workSheet.Cells[23, "E"] = raskrnpff;
                workSheet.Cells[23, "U"] = raskrnpff;
                workSheet.Cells[23, "Y"] = raskrnpff;
                workSheet.Cells[23, "B"] = raskrnpff;
                workSheet.Cells[23, "I"] = raskrnpff;
                workSheet.Cells[23, "J"] = raskrnpff;
                workSheet.Cells[23, "K"] = raskrnpff;
                workSheet.Cells[23, "R"] = raskrnpff;
                workSheet.Cells[23, "F"] = raskrnpff;
                workSheet.Cells[23, "G"] = raskrnpff;
                workSheet.Cells[23, "H"] = raskrnpff;
                workSheet.Cells[23, "X"] = raskrnpff;
                workSheet.Cells[23, "N"] = raskrnpff;
                workSheet.Cells[23, "O"] = raskrnpff;
                workSheet.Cells[23, "D"] = raskrnpff;
                workSheet.Cells[23, "V"] = raskrnpff;
                workSheet.Cells[23, "S"] = raskrnpff;
                workSheet.Cells[23, "W"] = telnpff;
                workSheet.Cells[23, "C"] = adresnpff;
                workSheet.Cells[23, "P"] = raskrnpff;
                workSheet.Cells[23, "Z"] = raskrnpff;
                workSheet.Cells[23, "AA"] = investirnpff;
                workSheet.Cells[23, "AB"] = investirnpff;
                workSheet.Cells[23, "AC"] = sobytnpff;
                workSheet.Cells[23, "L"] = investirnpff;
                workSheet.Cells[23, "M"] = investirnpff;
                workSheet.Cells[23, "AE"] = procinvestnpff;





                //НПФ ВОЛГА-КАПИТАЛ
                workSheet.Cells[25, "AD"] = raskrvolgacapital;
                workSheet.Cells[25, "O"] = strahpravvolgacapital;
                workSheet.Cells[25, "N"] = penspravvolgacapital;
                workSheet.Cells[25, "B"] = docvolgacapital;
                workSheet.Cells[25, "X"] = docvolgacapital;
                workSheet.Cells[25, "P"] = docvolgacapital;
                workSheet.Cells[25, "Y"] = orguprvolgacapital;
                workSheet.Cells[25, "L"] = UKSDvolgacapital;
                workSheet.Cells[25, "M"] = UKSDvolgacapital;
                workSheet.Cells[25, "I"] = pokazatelvolgacapital;
                workSheet.Cells[25, "J"] = pokazatelvolgacapital;
                workSheet.Cells[25, "H"] = pokazatelvolgacapital;
                workSheet.Cells[25, "K"] = pokazatelvolgacapital;
                workSheet.Cells[25, "F"] = pokazatelvolgacapital;
                workSheet.Cells[25, "G"] = pokazatelvolgacapital;
                workSheet.Cells[25, "Z"] = pokazatelvolgacapital;
                workSheet.Cells[25, "E"] = structuraakcionvolgacapital;
                workSheet.Cells[25, "D"] = otchetvolgacapital;
                workSheet.Cells[25, "R"] = otchetvolgacapital;
                workSheet.Cells[25, "V"] = otchetvolgacapital;
                workSheet.Cells[25, "U"] = konvladvolgacapital;
                workSheet.Cells[25, "C"] = adressvolgacapital;
                workSheet.Cells[25, "AA"] = portfelvolgacapital;
                workSheet.Cells[25, "AB"] = portfelvolgacapital;
                workSheet.Cells[25, "AC"] = sobvolgacapital;
                workSheet.Cells[25, "W"] = telvolgacapital;
                workSheet.Cells[25, "AE"] = procinvestvolgacapital;



                //НПФ ГАЗФОНД
                workSheet.Cells[26, "AD"] = raskrgazfond;
                workSheet.Cells[26, "N"] = raskrgazfond;
                workSheet.Cells[26, "P"] = raskrgazfond;
                workSheet.Cells[26, "D"] = raskrgazfond;
                workSheet.Cells[26, "V"] = raskrgazfond;
                workSheet.Cells[26, "I"] = raskrgazfond;
                workSheet.Cells[26, "K"] = raskrgazfond;
                workSheet.Cells[26, "H"] = raskrgazfond;
                workSheet.Cells[26, "F"] = raskrgazfond;
                workSheet.Cells[26, "AA"] = raskrgazfond;
                workSheet.Cells[26, "AB"] = raskrgazfond;
                workSheet.Cells[26, "Z"] = raskrgazfond;
                workSheet.Cells[26, "L"] = raskrgazfond;
                workSheet.Cells[26, "M"] = raskrgazfond;
                workSheet.Cells[26, "B"] = licenzgazfond;
                workSheet.Cells[26, "X"] = ustavgazfond;
                workSheet.Cells[26, "E"] = akcionergazfond;
                workSheet.Cells[26, "U"] = akcionergazfond;
                workSheet.Cells[26, "C"] = adressgazfond;
                workSheet.Cells[26, "Y"] = orguprgazfond;
                workSheet.Cells[26, "W"] = telgazfond;
                workSheet.Cells[26, "AE"] = procinvestgazfond;



                //НПФ ПРОФЕССИОНАЛЬНЫЙ
                workSheet.Cells[27, "AD"] = raskrprof;
                workSheet.Cells[27, "I"] = raskrprof;
                workSheet.Cells[27, "J"] = raskrprof;
                workSheet.Cells[27, "H"] = raskrprof;
                workSheet.Cells[27, "Z"] = raskrprof;
                workSheet.Cells[27, "K"] = raskrprof;
                workSheet.Cells[27, "I"] = raskrprof;
                workSheet.Cells[27, "J"] = raskrprof;
                workSheet.Cells[27, "E"] = raskrprof;
                workSheet.Cells[27, "U"] = raskrprof;
                workSheet.Cells[27, "B"] = licenzprof;
                workSheet.Cells[27, "X"] = ustavprof;
                workSheet.Cells[27, "N"] = raskrprof;
                workSheet.Cells[27, "O"] = raskrprof;
                workSheet.Cells[27, "P"] = raskrprof;
                workSheet.Cells[27, "L"] = raskrprof;
                workSheet.Cells[27, "M"] = raskrprof;
                workSheet.Cells[27, "D"] = raskrprof;
                workSheet.Cells[27, "R"] = raskrprof;
                workSheet.Cells[27, "V"] = raskrprof;
                workSheet.Cells[27, "AA"] = raskrprof;
                workSheet.Cells[27, "AB"] = raskrprof;
                workSheet.Cells[27, "F"] = raskrprof;
                workSheet.Cells[27, "G"] = raskrprof;
                workSheet.Cells[27, "Y"] = raskrprof;
                workSheet.Cells[27, "W"] = adresnpfprof;
                workSheet.Cells[27, "C"] = adresnpfprof;
                workSheet.Cells[27, "AE"] = procinvestprof;





                //ГАЗПРОМБАНК-ФОНД
                workSheet.Cells[28, "AD"] = raskrgbf;
                workSheet.Cells[28, "Z"] = pokazatelgbf;
                workSheet.Cells[28, "AB"] = pokazatelgbf;
                workSheet.Cells[28, "K"] = pokazatelgbf;
                workSheet.Cells[28, "I"] = pokazatelgbf;
                workSheet.Cells[28, "H"] = pokazatelgbf;
                workSheet.Cells[28, "F"] = dohodgbf;
                workSheet.Cells[28, "D"] = otchetgbf;
                workSheet.Cells[28, "V"] = otchetgbf;
                workSheet.Cells[28, "P"] = arhivgbf;
                workSheet.Cells[28, "L"] = UKgbf;
                workSheet.Cells[28, "M"] = SDgbf;
                workSheet.Cells[28, "X"] = ustavgbf;
                workSheet.Cells[28, "B"] = licenzgbf;
                workSheet.Cells[28, "N"] = penspravgbf;
                workSheet.Cells[28, "Y"] = raskrgbf;
                workSheet.Cells[28, "E"] = sostavakciongbf;
                workSheet.Cells[28, "U"] = sostavakciongbf;
                workSheet.Cells[28, "C"] = adressgbf;
                workSheet.Cells[28, "W"] = telgbf;
                workSheet.Cells[28, "AA"] = structportfgbf;
                workSheet.Cells[28, "AE"] = procinvestgbf;



                //НПФ ВНЕШЭКОНОМФОНД
                workSheet.Cells[29, "AD"] = raskrnpfveb;
                workSheet.Cells[29, "X"] = ustavnpfveb;
                workSheet.Cells[29, "B"] = licenznpfveb;
                workSheet.Cells[29, "C"] = adresnpfveb;
                workSheet.Cells[29, "W"] = telnpfveb;
                workSheet.Cells[29, "Y"] = orguprnpfveb;
                workSheet.Cells[29, "D"] = buhotchetnpfveb;
                workSheet.Cells[29, "V"] = buhotchetnpfveb;
                workSheet.Cells[29, "E"] = sostavakcionnpfveb;
                workSheet.Cells[29, "U"] = sostavakcionnpfveb;
                workSheet.Cells[29, "F"] = rezrazmPRnpfveb;
                workSheet.Cells[29, "Z"] = srvzvprocnpfveb;
                workSheet.Cells[29, "H"] = razmerdohodaPRnpfveb;
                workSheet.Cells[29, "I"] = infnpfveb;
                workSheet.Cells[29, "K"] = infnpfveb;
                workSheet.Cells[29, "L"] = UKnpfveb;
                workSheet.Cells[29, "M"] = SDnpfveb;
                workSheet.Cells[29, "N"] = ustavnpfveb;
                workSheet.Cells[29, "P"] = ustavnpfveb;
                workSheet.Cells[29, "AA"] = strportfnpfveb;
                workSheet.Cells[29, "AB"] = sostportfnpfveb;
                workSheet.Cells[29, "AE"] = procinvestnpfveb;



                //НПФ ВТБ ПЕНСИОННЫЙ ФОНД
                workSheet.Cells[30, "AD"] = raskrvtb;
                workSheet.Cells[30, "AA"] = investpolvtb;
                workSheet.Cells[30, "M"] = SDvtb;
                workSheet.Cells[30, "I"] = rezultvtb;
                workSheet.Cells[30, "J"] = rezultvtb;
                workSheet.Cells[30, "K"] = rezultvtb;
                workSheet.Cells[30, "H"] = rezultvtb;
                workSheet.Cells[30, "F"] = rezultvtb;
                workSheet.Cells[30, "G"] = rezultvtb;
                workSheet.Cells[30, "Z"] = rezultvtb;
                workSheet.Cells[30, "AB"] = structuraportfvtb;
                workSheet.Cells[30, "S"] = reorgvtb;
                workSheet.Cells[30, "T"] = reorgvtb;
                workSheet.Cells[30, "AC"] = sobytvtb;
                workSheet.Cells[30, "D"] = otchetvtb;
                workSheet.Cells[30, "V"] = otchetvtb;
                workSheet.Cells[30, "R"] = otchetvtb;
                workSheet.Cells[30, "Y"] = orguprvtb;
                workSheet.Cells[30, "X"] = ustavvtb;
                workSheet.Cells[30, "B"] = licenzvtb;
                workSheet.Cells[30, "E"] = sostavakcionvtb;
                workSheet.Cells[30, "U"] = konvladvtb;
                workSheet.Cells[30, "N"] = pravilavtb;
                workSheet.Cells[30, "O"] = pravilavtb;
                workSheet.Cells[30, "P"] = pravilavtb;
                workSheet.Cells[30, "C"] = adressvtb;
                workSheet.Cells[30, "W"] = telvtb;
                workSheet.Cells[30, "L"].Interior.Color = Color.Red;




                //НПФ РОСТЕХ
                workSheet.Cells[31, "AD"] = raskrrosteh;
                workSheet.Cells[31, "X"] = ustavrosteh;
                workSheet.Cells[31, "B"] = licenzrosteh;
                workSheet.Cells[31, "W"] = contactrosteh;
                workSheet.Cells[31, "C"] = contactrosteh;
                workSheet.Cells[31, "Y"] = orguprrosteh;
                workSheet.Cells[31, "N"] = penspravrosteh;
                workSheet.Cells[31, "P"] = penspravrosteh;
                workSheet.Cells[31, "O"] = strahpravrosteh;
                workSheet.Cells[31, "L"] = UKSDrosteh;
                workSheet.Cells[31, "M"] = UKSDrosteh;
                workSheet.Cells[31, "D"] = buhotchetrosteh;
                workSheet.Cells[31, "F"] = pokazatelrosteh;
                workSheet.Cells[31, "G"] = pokazatelrosteh;
                workSheet.Cells[31, "I"] = pokazatelrosteh;
                workSheet.Cells[31, "J"] = pokazatelrosteh;
                workSheet.Cells[31, "K"] = pokazatelrosteh;
                workSheet.Cells[31, "Z"] = pokazatelrosteh;
                workSheet.Cells[31, "H"] = pokazatelrosteh;
                workSheet.Cells[31, "R"] = formirPNrosteh;
                workSheet.Cells[31, "V"] = MSFOrosteh;
                workSheet.Cells[31, "AA"] = investrosteh;
                workSheet.Cells[31, "AB"] = investrosteh;
                workSheet.Cells[31, "AE"] = investrosteh;
                workSheet.Cells[31, "E"] = akcionerrosteh;
                workSheet.Cells[31, "U"] = akcionerrosteh;





                //ГАЗФОНД ПН
                workSheet.Cells[32, "AD"] = raskrgazfondpn;
                workSheet.Cells[32, "J"] = pokazatelgazfondpn;
                workSheet.Cells[32, "I"] = pokazatelgazfondpn;
                workSheet.Cells[32, "K"] = pokazatelgazfondpn;
                workSheet.Cells[32, "H"] = pokazatelgazfondpn;
                workSheet.Cells[32, "Z"] = pokazatelgazfondpn;
                workSheet.Cells[32, "F"] = pokazatelgazfondpn;
                workSheet.Cells[32, "G"] = pokazatelgazfondpn;
                workSheet.Cells[32, "X"] = ofdocgazfondpn;
                workSheet.Cells[32, "B"] = ofdocgazfondpn;
                workSheet.Cells[32, "N"] = ofdocgazfondpn;
                workSheet.Cells[32, "O"] = ofdocgazfondpn;
                workSheet.Cells[32, "E"] = ofdocgazfondpn;
                workSheet.Cells[32, "P"] = ofdocgazfondpn;
                workSheet.Cells[32, "AA"] = strinvestportfgazfondpn;
                workSheet.Cells[32, "AB"] = sostinvestportfgazfondpn;
                workSheet.Cells[32, "L"] = UKSDgazfondpn;
                workSheet.Cells[32, "M"] = UKSDgazfondpn;
                workSheet.Cells[32, "R"] = otchetgazfondpn;
                workSheet.Cells[32, "D"] = otchetgazfondpn;
                workSheet.Cells[32, "V"] = otchetgazfondpn;
                workSheet.Cells[32, "U"] = otchetgazfondpn;
                workSheet.Cells[32, "Y"] = otchetgazfondpn;
                workSheet.Cells[32, "W"] = telgazfondpn;
                workSheet.Cells[32, "C"] = adressgazfondpn;
                workSheet.Cells[32, "AE"] = procinvestgazfondpn;




                //НПФ ТРАДИЦИЯ
                workSheet.Cells[33, "AD"] = raskrtrad;
                workSheet.Cells[33, "Y"] = upravtrad;
                workSheet.Cells[33, "E"] = akciontrad;
                workSheet.Cells[33, "U"] = akciontrad;
                workSheet.Cells[33, "F"] = rezrazmPRtrad;
                workSheet.Cells[33, "Z"] = srvzvproctrad;
                workSheet.Cells[33, "I"] = kolvkltrad;
                workSheet.Cells[33, "H"] = razmerdohodatrad;
                workSheet.Cells[33, "K"] = razmerstrahPR;
                workSheet.Cells[33, "D"] = otchettrad;
                workSheet.Cells[33, "V"] = otchettrad;
                workSheet.Cells[33, "AA"] = investtrad;
                workSheet.Cells[33, "AB"] = investtrad;
                workSheet.Cells[33, "AE"] = investtrad;
                workSheet.Cells[33, "L"] = UKtrad;
                workSheet.Cells[33, "M"] = SDtrad;
                workSheet.Cells[33, "X"] = ustavtrad;
                workSheet.Cells[33, "B"] = licenztrad;
                workSheet.Cells[33, "N"] = penspravtrad;
                workSheet.Cells[33, "P"] = penspravtrad;
                workSheet.Cells[33, "AC"] = sobyttrad;
                workSheet.Cells[33, "W"] = teltrad;
                workSheet.Cells[33, "C"] = adrestrad;



                //СТРОЙПРОМ ФОНД
                workSheet.Cells[34, "AD"] = raskrmpsfond;
                workSheet.Cells[34, "X"] = ustavmpsfond;
                workSheet.Cells[34, "N"] = penspravmpsfond;
                workSheet.Cells[34, "P"] = izmpenspravmpsfond;
                workSheet.Cells[34, "B"] = licenzmpsfond;
                workSheet.Cells[34, "E"] = akcionmpsfond;
                workSheet.Cells[34, "U"] = akcionmpsfond;
                workSheet.Cells[34, "I"] = raskrmpsfond;
                workSheet.Cells[34, "K"] = raskrmpsfond;
                workSheet.Cells[34, "Z"] = pokazatelmpsfond;
                workSheet.Cells[34, "H"] = raskrmpsfond;
                workSheet.Cells[34, "F"] = pokazatelmpsfond;
                workSheet.Cells[34, "D"] = pokazatelmpsfond;
                workSheet.Cells[34, "V"] = pokazatelmpsfond;
                workSheet.Cells[34, "L"] = raskrmpsfond;
                workSheet.Cells[34, "M"] = raskrmpsfond;
                workSheet.Cells[34, "Y"] = raskrmpsfond;
                workSheet.Cells[34, "AA"] = raskrmpsfond;
                workSheet.Cells[34, "AB"] = raskrmpsfond;
                workSheet.Cells[34, "C"] = adresmpsfond;
                workSheet.Cells[34, "W"] = adresmpsfond;





                //НПФ НАЦИОНАЛЬНЫЙ
                workSheet.Cells[35, "AD"] = raskrnnpf;
                workSheet.Cells[35, "B"] = docnnpf;
                workSheet.Cells[35, "X"] = docnnpf;
                workSheet.Cells[35, "N"] = docnnpf;
                workSheet.Cells[35, "O"] = docnnpf;
                workSheet.Cells[35, "P"] = docnnpf;
                workSheet.Cells[35, "E"] = akcionnnpf;
                workSheet.Cells[35, "U"] = akcionnnpf;
                workSheet.Cells[35, "F"] = rezinvestnnpf;
                workSheet.Cells[35, "G"] = rezinvestnnpf;
                workSheet.Cells[35, "Z"] = rezinvestnnpf;
                workSheet.Cells[35, "I"] = kolvkladnnpf;
                workSheet.Cells[35, "J"] = kolvkladnnpf;
                workSheet.Cells[35, "H"] = rezinvestnnpf;
                workSheet.Cells[35, "K"] = razmerprpnnnpf;
                workSheet.Cells[35, "D"] = otchetnostnnpf;
                workSheet.Cells[35, "R"] = otchetSPNnnpf;
                workSheet.Cells[35, "V"] = audzaklnnpf;
                workSheet.Cells[35, "AA"] = investirnnpf;
                workSheet.Cells[35, "AB"] = investirnnpf;
                workSheet.Cells[35, "AE"] = investirnnpf;
                workSheet.Cells[35, "AC"] = sobnnpf;
                workSheet.Cells[35, "L"] = UKSDnnpf;
                workSheet.Cells[35, "M"] = UKSDnnpf;
                workSheet.Cells[35, "C"] = contactnnpf;
                workSheet.Cells[35, "W"] = contactnnpf;
                workSheet.Cells[35, "Y"] = orguprnnpf;






                //НПФ АТОМФОНД
                workSheet.Cells[36, "AD"] = raskratomfond;
                workSheet.Cells[36, "E"] = akcionatomfond;
                workSheet.Cells[36, "U"] = akcionatomfond;
                workSheet.Cells[36, "Y"] = orgupratomfond;
                workSheet.Cells[36, "D"] = finotchetatomfond;
                workSheet.Cells[36, "V"] = finotchetatomfond;
                workSheet.Cells[36, "G"] = rezinvestPNatomfond;
                workSheet.Cells[36, "AA"] = structuraportfatomfond;
                workSheet.Cells[36, "Q"] = zapratomfond;
                workSheet.Cells[36, "AB"] = sostavportfatomfond;
                workSheet.Cells[36, "J"] = pokazatelatomfond;
                workSheet.Cells[36, "K"] = pokazatelatomfond;
                workSheet.Cells[36, "L"] = UKatomfond;
                workSheet.Cells[36, "M"] = SDatomfond;
                workSheet.Cells[36, "X"] = ustavatomfond;
                workSheet.Cells[36, "B"] = licenzatomfond;
                workSheet.Cells[36, "O"] = strahpravatomfond;
                workSheet.Cells[36, "C"] = adresatomfond;
                workSheet.Cells[36, "W"] = telatomfond;
                workSheet.Cells[36, "AE"] = procinvestatomfond;



                //НПФ СБЕРФОНД
                workSheet.Cells[37, "AD"] = raskrsberfond;
                workSheet.Cells[37, "X"] = ustavsberfond;
                workSheet.Cells[37, "B"] = licenzsberfond;
                workSheet.Cells[37, "N"] = penspravsberfond;
                workSheet.Cells[37, "AA"] = raskrsberfond;
                workSheet.Cells[37, "AB"] = raskrsberfond;
                workSheet.Cells[37, "L"] = raskrsberfond;
                workSheet.Cells[37, "M"] = raskrsberfond;
                workSheet.Cells[37, "D"] = raskrsberfond;
                workSheet.Cells[37, "V"] = raskrsberfond;
                workSheet.Cells[37, "Y"] = raskrsberfond;
                workSheet.Cells[37, "F"] = raskrsberfond;
                workSheet.Cells[37, "H"] = raskrsberfond;
                workSheet.Cells[37, "I"] = raskrsberfond;
                workSheet.Cells[37, "K"] = raskrsberfond;
                workSheet.Cells[37, "P"] = raskrsberfond;
                workSheet.Cells[37, "Z"] = raskrsberfond;
                workSheet.Cells[37, "AC"] = raskrsberfond;
                workSheet.Cells[37, "E"] = raskrsberfond;
                workSheet.Cells[37, "W"] = telsberfond;
                workSheet.Cells[37, "C"] = adressberfond;
                workSheet.Cells[37, "AE"] = raskrsberfond;




                //НПФ ДОВЕРИЕ
                workSheet.Cells[38, "AD"] = raskrdoverie;
                workSheet.Cells[38, "B"] = licenzdoverie;
                workSheet.Cells[38, "X"] = ustavdoverie;
                workSheet.Cells[38, "N"] = ofdocdoverie;
                workSheet.Cells[38, "O"] = ofdocdoverie;
                workSheet.Cells[38, "P"] = ofdocdoverie;
                workSheet.Cells[38, "E"] = akciondoverie;
                workSheet.Cells[38, "U"] = konechvladdoverie;
                workSheet.Cells[38, "Y"] = orguprdoverie;
                workSheet.Cells[38, "AA"] = otchetdoverie;
                workSheet.Cells[38, "K"] = otchetdoverie;
                workSheet.Cells[38, "I"] = otchetdoverie;
                workSheet.Cells[38, "J"] = otchetdoverie;
                workSheet.Cells[38, "D"] = otchetdoverie;
                workSheet.Cells[38, "AB"] = otchetdoverie;
                workSheet.Cells[38, "F"] = otchetdoverie;
                workSheet.Cells[38, "G"] = otchetdoverie;
                workSheet.Cells[38, "H"] = otchetdoverie;
                workSheet.Cells[38, "R"] = otchetdoverie;
                workSheet.Cells[38, "V"] = otchetdoverie;
                workSheet.Cells[38, "Z"] = otchetdoverie;
                workSheet.Cells[38, "L"] = UKdoverie;
                workSheet.Cells[38, "M"] = SDdoverie;
                workSheet.Cells[38, "C"] = contactdoverie;
                workSheet.Cells[38, "W"] = contactdoverie;


                //НПФ ИНГОССТРАХ ПЕНСИЯ
                workSheet.Cells[39, "AD"] = raskringo;
                workSheet.Cells[39, "B"] = licenzingo;
                workSheet.Cells[39, "X"] = ustavingo;
                workSheet.Cells[39, "N"] = penspravingo;
                workSheet.Cells[39, "D"] = otchetingo;
                workSheet.Cells[39, "P"] = otchetingo;
                workSheet.Cells[39, "I"] = otchetingo;
                workSheet.Cells[39, "H"] = otchetingo;
                workSheet.Cells[39, "Z"] = otchetingo;
                workSheet.Cells[39, "AE"] = otchetingo;
                workSheet.Cells[39, "K"] = otchetingo;
                workSheet.Cells[39, "F"] = otchetingo;
                workSheet.Cells[39, "AA"] = investingo;
                workSheet.Cells[39, "AB"] = investingo;
                workSheet.Cells[39, "L"] = investingo;
                workSheet.Cells[39, "M"] = investingo;
                workSheet.Cells[39, "U"] = sostavakcioningo;
                workSheet.Cells[39, "Y"] = orgupringo;
                workSheet.Cells[39, "C"] = adresingo;
                workSheet.Cells[39, "W"] = telingo;
                workSheet.Cells[39, "V"] = otchetingo;





                //НПФ ОТКРЫТИЕ
                workSheet.Cells[40, "AD"] = raskropen;
                workSheet.Cells[40, "Y"] = orgupropen;
                workSheet.Cells[40, "E"] = orgupropen;
                workSheet.Cells[40, "U"] = orgupropen;
                workSheet.Cells[40, "X"] = raskropen;
                workSheet.Cells[40, "N"] = raskropen;
                workSheet.Cells[40, "P"] = raskropen;
                workSheet.Cells[40, "B"] = raskropen;
                workSheet.Cells[40, "T"] = raskropen;
                workSheet.Cells[40, "S"] = raskropen;
                workSheet.Cells[40, "D"] = raskropen;
                workSheet.Cells[40, "V"] = raskropen;
                workSheet.Cells[40, "O"] = raskropen;
                workSheet.Cells[40, "I"] = raskropen;
                workSheet.Cells[40, "J"] = raskropen;
                workSheet.Cells[40, "K"] = raskropen;
                workSheet.Cells[40, "F"] = raskropen;
                workSheet.Cells[40, "G"] = raskropen;
                workSheet.Cells[40, "H"] = raskropen;
                workSheet.Cells[40, "R"] = raskropen;
                workSheet.Cells[40, "Z"] = investopen;
                workSheet.Cells[40, "AA"] = investopen;
                workSheet.Cells[40, "AB"] = investopen;
                workSheet.Cells[40, "L"] = investopen;
                workSheet.Cells[40, "M"] = investopen;
                workSheet.Cells[40, "W"] = adresopen;
                workSheet.Cells[40, "C"] = adresopen;


                //НПФ КОРАБЕЛ
                workSheet.Cells[41, "AD"] = raskrkorabel;
                workSheet.Cells[41, "B"] = licenzkorabel;
                workSheet.Cells[41, "D"] = buhotchetkorabel;
                workSheet.Cells[41, "V"] = buhotchetkorabel;
                workSheet.Cells[41, "X"] = dockorabel;
                workSheet.Cells[41, "N"] = dockorabel;
                workSheet.Cells[41, "P"] = dockorabel;
                workSheet.Cells[41, "E"] = akcionkorabel;
                workSheet.Cells[41, "U"] = akcionkorabel;
                workSheet.Cells[41, "Y"] = orguprkorabel;
                workSheet.Cells[41, "L"] = UKkorabel;
                workSheet.Cells[41, "M"] = SDkorabel;
                workSheet.Cells[41, "AA"] = pokazatelkorabel;
                workSheet.Cells[41, "AB"] = pokazatelkorabel;
                workSheet.Cells[41, "K"] = pokazatelkorabel;
                workSheet.Cells[41, "H"] = pokazatelkorabel;
                workSheet.Cells[41, "I"] = pokazatelkorabel;
                workSheet.Cells[41, "F"] = pokazatelkorabel;
                workSheet.Cells[41, "Z"] = pokazatelkorabel;
                workSheet.Cells[41, "C"] = adreskorabel;
                workSheet.Cells[41, "W"] = telkorabel;



                workBook.SaveAs(fileName);
                workBook.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.ToString());
            }

            MessageBox.Show("Файл " + fileName + " записан успешно!");

           


            displayChange(url);

        }

        private void button1_Click(object sender, EventArgs e)
        {
            backgroundWorker1.RunWorkerAsync();
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            getAllInfo();
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage <= 100)
                progressBar1.Value = e.ProgressPercentage;
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            MessageBox.Show("Операция завершена!");
        }
    }
}
