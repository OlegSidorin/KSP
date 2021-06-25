namespace KSP
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Reflection;
    using System.Windows.Media.Imaging;
    using Autodesk.Revit.UI;
    class Main : IExternalApplication
    {
        public static string tabName { get; set; } = "КПСП";
        public static string panelBIM360Name { get; set; } = "Публикация в BIM360";
        public static string panelCheckName { get; set; } = "Подготовка к МГЭ";
        #region external application public methods
        public Result OnStartup(UIControlledApplication application)
        {
            List<RibbonPanel> panelList = new List<RibbonPanel>();
            try
            {
                panelList = application.GetRibbonPanels(tabName);
            }
            catch
            {

            }

            if (panelList.Count == 0)
            {
                application.CreateRibbonTab(tabName);
            }

            string path = Assembly.GetExecutingAssembly().Location;
            var panelBIM360 = application.CreateRibbonPanel(tabName, panelBIM360Name);


            var CreateBIM360ViewBtnData = new PushButtonData("CreateBIM360BtnData", "Создать вид\nNavisworks", path, "KSP.CreateBIM360ViewCommand")
            {
                ToolTipImage = new BitmapImage(new Uri(Path.GetDirectoryName(path) + "\\res\\ao-32.png", UriKind.Absolute)),
                ToolTip = "Создает вид Экспорт в Navisworks, \nвыгружает ссылки,\nготовит набор к публикации в BIM360"
            };
            var CreateBIM360ViewBtn = panelBIM360.AddItem(CreateBIM360ViewBtnData) as PushButton;
            CreateBIM360ViewBtn.LargeImage = new BitmapImage(new Uri(Path.GetDirectoryName(path) + "\\res\\ao-32.png", UriKind.Absolute));

            var WorkSetRVTLinksBtnData = new PushButtonData("WorkSetRVTLinksBtnData", "Рабочие наборы\nдля RVT связей", path, "KSP.WorkSetRVTLinksCommand")
            {
                ToolTipImage = new BitmapImage(new Uri(Path.GetDirectoryName(path) + "\\res\\ws-32.png", UriKind.Absolute)),
                ToolTip = "Создает рабочие наборы для RVT связей"
            };
            var WorkSetRVTLinksBtn = panelBIM360.AddItem(WorkSetRVTLinksBtnData) as PushButton;
            WorkSetRVTLinksBtn.LargeImage = new BitmapImage(new Uri(Path.GetDirectoryName(path) + "\\res\\ws-32.png", UriKind.Absolute));


            var panelCheck = application.CreateRibbonPanel(tabName, panelCheckName);

            var BasePointsBtnData = new PushButtonData("BasePointsBtnData", "Базовые точки\nИнформация", path, "KSP.BasePointsCommand")
            {
                ToolTipImage = new BitmapImage(new Uri(Path.GetDirectoryName(path) + "\\res\\bp-32.png", UriKind.Absolute)),
                ToolTip = "Создает .xlsx файл с информацией о базовых точках"
            };
            var BasePointsBtn = panelCheck.AddItem(BasePointsBtnData) as PushButton;
            BasePointsBtn.LargeImage = new BitmapImage(new Uri(Path.GetDirectoryName(path) + "\\res\\bp-32.png", UriKind.Absolute));

            var CheckMSKArBtnData = new PushButtonData("CheckMSKArBtnData", "Проверка\nМСК_АР", path, "KSP.CheckMSKARwGCommand")
            {
                ToolTipImage = new BitmapImage(new Uri(Path.GetDirectoryName(path) + "\\res\\check-ar-32.png", UriKind.Absolute)),
                ToolTip = "Проверяет наличие МСК_ параметров в модели АР"
            };
            var CheckMSKArBtn = panelCheck.AddItem(CheckMSKArBtnData) as PushButton;
            CheckMSKArBtn.LargeImage = new BitmapImage(new Uri(Path.GetDirectoryName(path) + "\\res\\check-ar-32.png", UriKind.Absolute));

            var CheckMSKKrBtnData = new PushButtonData("CheckMSKKrBtnData", "Проверка\nМСК_KР", path, "KSP.CheckMSKKRwGCommand")
            {
                ToolTipImage = new BitmapImage(new Uri(Path.GetDirectoryName(path) + "\\res\\check-kr-32.png", UriKind.Absolute)),
                ToolTip = "Проверяет наличие МСК_ параметров в модели КР"
            };
            var CheckMSKKrBtn = panelCheck.AddItem(CheckMSKKrBtnData) as PushButton;
            CheckMSKKrBtn.LargeImage = new BitmapImage(new Uri(Path.GetDirectoryName(path) + "\\res\\check-kr-32.png", UriKind.Absolute));

            var CheckMSKIosBtnData = new PushButtonData("CheckMSKIosBtnData", "Проверка\nМСК_ИОС", path, "KSP.CheckMSKIOSwGCommand")
            {
                ToolTipImage = new BitmapImage(new Uri(Path.GetDirectoryName(path) + "\\res\\check-ios-32.png", UriKind.Absolute)),
                ToolTip = "Проверяет наличие МСК_ параметров в модели инженерного раздела"
            };
            var CheckMSKIosBtn = panelCheck.AddItem(CheckMSKIosBtnData) as PushButton;
            CheckMSKIosBtn.LargeImage = new BitmapImage(new Uri(Path.GetDirectoryName(path) + "\\res\\check-ios-32.png", UriKind.Absolute));

            var CodData = new PushButtonData("CodData", "Заполнение\nКлассификатора", path, "KSP.CheckCodCommand")
            {
                ToolTipImage = new BitmapImage(new Uri(Path.GetDirectoryName(path) + "\\res\\kk-32.png", UriKind.Absolute)),
                ToolTip = "Проверяет заполнение Кода по классификатору"
            };
            var CodDataBtn = panelCheck.AddItem(CodData) as PushButton;
            CodDataBtn.LargeImage = new BitmapImage(new Uri(Path.GetDirectoryName(path) + "\\res\\kk-32.png", UriKind.Absolute));

            /*
            string panelTestName = "В разработке";
            var panelTest = application.CreateRibbonPanel(tabName, panelTestName);

            var JoshTest01Data = new PushButtonData("JoshTest01Data", "Тест 1\nДжоша", Assembly.GetExecutingAssembly().Location, "KSP.JoshTest01Command")
            {
                ToolTipImage = new BitmapImage(new Uri(user + @"\source\repos\KSP\res\img\red-m-128.png")),
                ToolTip = "Фигня какая-то "
            };
            var JoshTest01Btn = panelTest.AddItem(JoshTest01Data) as PushButton;
            JoshTest01Btn.LargeImage = new BitmapImage(new Uri(user + @"\source\repos\KSP\res\img\cool-32.png"));

            var SumData = new PushButtonData("SumData", "Складывает\nобъемы", Assembly.GetExecutingAssembly().Location, "KSP.SumCommand")
            {
                ToolTipImage = new BitmapImage(new Uri(user + @"\source\repos\KSP\res\img\batman-32.png")),
                ToolTip = "Суммирует объемы элементов модели"
            };
            var SumDataBtn = panelTest.AddItem(SumData) as PushButton;
            SumDataBtn.LargeImage = new BitmapImage(new Uri(user + @"\source\repos\KSP\res\img\batman-32.png"));
            */

            string panelKIRname = "КИРГИЗИЯ";
            var panelKIR = application.CreateRibbonPanel(tabName, panelKIRname);

            string ButtonADGCName = "ADGCBtn";
            string ButtonADGCText = "Добавить\nКлассификатор";
            PushButtonData ButtonADGCData = new PushButtonData(ButtonADGCName, ButtonADGCText, path, "KSP.AddADGClass")
            {
                ToolTipImage = new BitmapImage(new Uri(Path.GetDirectoryName(path) + "\\res\\kir-class.png", UriKind.Absolute)),
                ToolTip = "Добавляет классификатор"
            };
            ButtonADGCData.LargeImage = new BitmapImage(new Uri(Path.GetDirectoryName(path) + "\\res\\kir-class.png", UriKind.Absolute));

            string ButtonADG1Name = "ADG1Btn";
            string ButtonADG1Text = "Добавить\nв ОВИК";
            PushButtonData ButtonADG1Data = new PushButtonData(ButtonADG1Name, ButtonADG1Text, path, "KSP.AddGPParameterOVIK")
            {
                ToolTipImage = new BitmapImage(new Uri(Path.GetDirectoryName(path) + "\\res\\kir-ovik.png", UriKind.Absolute)),
                ToolTip = "Добавляет параметры в ОВ по Киргизии"
            };
            ButtonADG1Data.LargeImage = new BitmapImage(new Uri(Path.GetDirectoryName(path) + "\\res\\kir-ovik.png", UriKind.Absolute));
            //PushButton ButtonADG1 = panelCheck.AddItem(ButtonADG1Data) as PushButton;

            string ButtonADG2Name = "ADG2Btn";
            string ButtonADG2Text = "Добавить\nв ВК";
            PushButtonData ButtonADG2Data = new PushButtonData(ButtonADG2Name, ButtonADG2Text, path, "KSP.AddGPParameterVK")
            {
                ToolTipImage = new BitmapImage(new Uri(Path.GetDirectoryName(path) + "\\res\\kir-vk.png", UriKind.Absolute)),
                ToolTip = "Добавляет параметры в ВК по Киргизии"
            };
            ButtonADG2Data.LargeImage = new BitmapImage(new Uri(Path.GetDirectoryName(path) + "\\res\\kir-vk.png", UriKind.Absolute));
            //PushButton ButtonADG1 = panelCheck.AddItem(ButtonADG1Data) as PushButton;

            SplitButtonData sBtnData = new SplitButtonData("splitButton", "Split");
            SplitButton sBtn = panelKIR.AddItem(sBtnData) as SplitButton;
            sBtn.AddPushButton(ButtonADGCData);
            sBtn.AddPushButton(ButtonADG1Data);
            sBtn.AddPushButton(ButtonADG2Data);

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        #endregion



    }
}
