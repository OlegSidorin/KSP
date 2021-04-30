namespace KSP
{
    using System;
    using System.Reflection;
    using System.Windows.Media.Imaging;
    using Autodesk.Revit.UI;
    class Main : IExternalApplication
    {
        #region external application public methods
        public Result OnStartup(UIControlledApplication application)
        {
            string tabName = "КПСП";
            string panelAnnotationName = "Весьма полезные методы";
            application.CreateRibbonTab(tabName);
            var panelAnnotation = application.CreateRibbonPanel(tabName, panelAnnotationName);
            var CreateBIM360ViewBtnData = new PushButtonData("CreateBIM360BtnData", "Create\nNavisView", Assembly.GetExecutingAssembly().Location, "KSP.CreateBIM360ViewCommand")
            {
                ToolTipImage = new BitmapImage(new Uri(@"C:\Users\Sidorin_O\source\repos\KSP\res\img\navis-3d-100.png")),
                ToolTip = "Создает вид Экспорт в Navisworks"
            };
            var CreateBIM360ViewBtn = panelAnnotation.AddItem(CreateBIM360ViewBtnData) as PushButton;
            CreateBIM360ViewBtn.LargeImage = new BitmapImage(new Uri(@"C:\Users\Sidorin_O\source\repos\KSP\res\img\navis-3d-32.png"));

            var WorkSetRVTLinksBtnData = new PushButtonData("WorkSetRVTLinksBtnData", "WorkSet\nRVTLinks", Assembly.GetExecutingAssembly().Location, "KSP.WorkSetRVTLinksCommand")
            {
                ToolTipImage = new BitmapImage(new Uri(@"C:\Users\Sidorin_O\source\repos\KSP\res\img\wset-100.png")),
                ToolTip = "Создает рабочие набопы для RVT связей"
            };
            var WorkSetRVTLinksBtn = panelAnnotation.AddItem(WorkSetRVTLinksBtnData) as PushButton;
            WorkSetRVTLinksBtn.LargeImage = new BitmapImage(new Uri(@"C:\Users\Sidorin_O\source\repos\KSP\res\img\wset-32.png"));

            var BasePointsBtnData = new PushButtonData("BasePointsBtnData", "BasePoints\nInfo", Assembly.GetExecutingAssembly().Location, "KSP.BasePointsCommand")
            {
                ToolTipImage = new BitmapImage(new Uri(@"C:\Users\Sidorin_O\source\repos\KSP\res\img\base-point-150.png")),
                ToolTip = "Создает Exl файл о базовых точках"
            };
            var BasePointsBtn = panelAnnotation.AddItem(BasePointsBtnData) as PushButton;
            BasePointsBtn.LargeImage = new BitmapImage(new Uri(@"C:\Users\Sidorin_O\source\repos\KSP\res\img\base-point-32.png"));

            var CheckMSKArBtnData = new PushButtonData("CheckMSKArBtnData", "Проверка\nМСК_АР", Assembly.GetExecutingAssembly().Location, "KSP.CheckMSKArCommand")
            {
                ToolTipImage = new BitmapImage(new Uri(@"C:\Users\Sidorin_O\source\repos\KSP\res\img\check-ar-150.png")),
                ToolTip = "Проверяет наличие МСК_ параметров"
            };
            var CheckMSKArBtn = panelAnnotation.AddItem(CheckMSKArBtnData) as PushButton;
            CheckMSKArBtn.LargeImage = new BitmapImage(new Uri(@"C:\Users\Sidorin_O\source\repos\KSP\res\img\check-ar-32.png"));

            var CheckMSKKrBtnData = new PushButtonData("CheckMSKKrBtnData", "Проверка\nМСК_KР", Assembly.GetExecutingAssembly().Location, "KSP.CheckMSKKrCommand")
            {
                ToolTipImage = new BitmapImage(new Uri(@"C:\Users\Sidorin_O\source\repos\KSP\res\img\check-kr-150.png")),
                ToolTip = "Проверяет наличие МСК_ параметров"
            };
            var CheckMSKKrBtn = panelAnnotation.AddItem(CheckMSKKrBtnData) as PushButton;
            CheckMSKKrBtn.LargeImage = new BitmapImage(new Uri(@"C:\Users\Sidorin_O\source\repos\KSP\res\img\check-kr-32.png"));

            var CheckMSKIosBtnData = new PushButtonData("CheckMSKIosBtnData", "Проверка\nМСК_ИОС", Assembly.GetExecutingAssembly().Location, "KSP.CheckMSKIosCommand")
            {
                ToolTipImage = new BitmapImage(new Uri(@"C:\Users\Sidorin_O\source\repos\KSP\res\img\check-ios-150.png")),
                ToolTip = "Проверяет наличие МСК_ параметров"
            };
            var CheckMSKIosBtn = panelAnnotation.AddItem(CheckMSKIosBtnData) as PushButton;
            CheckMSKIosBtn.LargeImage = new BitmapImage(new Uri(@"C:\Users\Sidorin_O\source\repos\KSP\res\img\check-ios-32.png"));

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        #endregion



    }
}
