namespace KSP
{
    using System;
    using System.Reflection;
    using System.Windows.Media.Imaging;
    using Autodesk.Revit.UI;
    using static KSP.CommonMethods;
    class Main : IExternalApplication
    {
        #region external application public methods
        public Result OnStartup(UIControlledApplication application)
        {
            string tabName = "КПСП";
            application.CreateRibbonTab(tabName);

            string panelBIM360Name = "Подготовка к BIM360";
            var panelBIM360 = application.CreateRibbonPanel(tabName, panelBIM360Name);
            var CreateBIM360ViewBtnData = new PushButtonData("CreateBIM360BtnData", "Создать вид\nЭкспорт в Navisworks", Assembly.GetExecutingAssembly().Location, "KSP.CreateBIM360ViewCommand")
            {
                ToolTipImage = new BitmapImage(new Uri(user + @"\source\repos\KSP\src\KSP\res\ao-32.png")),
                ToolTip = "Создает вид Экспорт в Navisworks, \nвыгружает ссылки,\nготовит набор к публикации в BIM360"
            };
            var CreateBIM360ViewBtn = panelBIM360.AddItem(CreateBIM360ViewBtnData) as PushButton;
            CreateBIM360ViewBtn.LargeImage = new BitmapImage(new Uri(user + @"\source\repos\KSP\src\KSP\res\ao-32.png"));

            var WorkSetRVTLinksBtnData = new PushButtonData("WorkSetRVTLinksBtnData", "WorkSet\nRVTLinks", Assembly.GetExecutingAssembly().Location, "KSP.WorkSetRVTLinksCommand")
            {
                ToolTipImage = new BitmapImage(new Uri(user + @"\source\repos\KSP\src\KSP\res\ws-32.png")),
                ToolTip = "Создает рабочие набопы для RVT связей"
            };
            var WorkSetRVTLinksBtn = panelBIM360.AddItem(WorkSetRVTLinksBtnData) as PushButton;
            WorkSetRVTLinksBtn.LargeImage = new BitmapImage(new Uri(user + @"\source\repos\KSP\src\KSP\res\ws-32.png"));

            string panelCheckName = "Подготовка к МГЭ";
            var panelCheck = application.CreateRibbonPanel(tabName, panelCheckName);

            var BasePointsBtnData = new PushButtonData("BasePointsBtnData", "BasePoints\nInfo", Assembly.GetExecutingAssembly().Location, "KSP.BasePointsCommand")
            {
                ToolTipImage = new BitmapImage(new Uri(user + @"\source\repos\KSP\src\KSP\res\bp-32.png")),
                ToolTip = "Создает Exl файл о базовых точках"
            };
            var BasePointsBtn = panelCheck.AddItem(BasePointsBtnData) as PushButton;
            BasePointsBtn.LargeImage = new BitmapImage(new Uri(user + @"\source\repos\KSP\src\KSP\res\bp-32.png"));

            var CheckMSKArBtnData = new PushButtonData("CheckMSKArBtnData", "Проверка\nМСК_АР", Assembly.GetExecutingAssembly().Location, "KSP.CheckARGuidCommand")
            {
                ToolTipImage = new BitmapImage(new Uri(user + @"\source\repos\KSP\res\img\check-ar-150.png")),
                ToolTip = "Проверяет наличие МСК_ параметров в модели АР"
            };
            var CheckMSKArBtn = panelCheck.AddItem(CheckMSKArBtnData) as PushButton;
            CheckMSKArBtn.LargeImage = new BitmapImage(new Uri(user + @"\source\repos\KSP\res\img\check-ar-32.png"));

            var CheckMSKKrBtnData = new PushButtonData("CheckMSKKrBtnData", "Проверка\nМСК_KР", Assembly.GetExecutingAssembly().Location, "KSP.CheckKRGuidCommand")
            {
                ToolTipImage = new BitmapImage(new Uri(user + @"\source\repos\KSP\res\img\check-kr-150.png")),
                ToolTip = "Проверяет наличие МСК_ параметров в модели КР"
            };
            var CheckMSKKrBtn = panelCheck.AddItem(CheckMSKKrBtnData) as PushButton;
            CheckMSKKrBtn.LargeImage = new BitmapImage(new Uri(user + @"\source\repos\KSP\res\img\check-kr-32.png"));

            var CheckMSKIosBtnData = new PushButtonData("CheckMSKIosBtnData", "Проверка\nМСК_ИОС", Assembly.GetExecutingAssembly().Location, "KSP.CheckIOSGuidCommand")
            {
                ToolTipImage = new BitmapImage(new Uri(user + @"\source\repos\KSP\res\img\check-ios-150.png")),
                ToolTip = "Проверяет наличие МСК_ параметров в модели инженерного раздела"
            };
            var CheckMSKIosBtn = panelCheck.AddItem(CheckMSKIosBtnData) as PushButton;
            CheckMSKIosBtn.LargeImage = new BitmapImage(new Uri(user + @"\source\repos\KSP\res\img\check-ios-32.png"));

            var CodData = new PushButtonData("CodData", "Проверка заполнения\nКода по классификатору", Assembly.GetExecutingAssembly().Location, "KSP.CheckCodCommand")
            {
                ToolTipImage = new BitmapImage(new Uri(user + @"\source\repos\KSP\res\img\blue-m-32.png")),
                ToolTip = "Проверяет заполнение Кода по классификатору"
            };
            var CodDataBtn = panelCheck.AddItem(CodData) as PushButton;
            CodDataBtn.LargeImage = new BitmapImage(new Uri(user + @"\source\repos\KSP\res\img\blue-m-32.png"));

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

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        #endregion



    }
}
