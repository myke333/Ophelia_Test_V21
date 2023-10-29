using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_3
{
    class CrudKnmliyca
    {
        protected static WindowsDriver<WindowsElement> rootSession;
        public static WindowsDriver<WindowsElement> RootSession()
        {
            var appCapabilities = new AppiumOptions();
            appCapabilities.AddAdditionalCapability("app", "Root");
            var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);
            return ApplicationSession;
        }

        static public WindowsDriver<WindowsElement> ReloadSession(WindowsDriver<WindowsElement> session, string Clase)
        {
            try
            {
                var ApplicationWindow = session.FindElementByClassName(Clase);
                var ApplicationSessionHandle = ApplicationWindow.GetAttribute("NativeWindowHandle");
                ApplicationSessionHandle = (int.Parse(ApplicationSessionHandle)).ToString("x"); //Convert to hex


                var appCapabilities = new AppiumOptions();
                appCapabilities.AddAdditionalCapability("appTopLevelWindow", ApplicationSessionHandle);
                var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);

                string ses = ApplicationSession.PageSource;
                return ApplicationSession;
            }
            catch (Exception ex) { return null; }
        }
        public static void CrudKNmliyca(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            List<string> lupas = new List<string> { data[0], data[1], data[2] };
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementList1 = desktopSession.FindElementsByClassName("TDBMemo");
            if (Tipo == 1)
            {
                PruebaCRUD.LupaDinamica(desktopSession, lupas);
                ElementList[15].SendKeys(data[3]);
                Thread.Sleep(500);
                ElementList1[0].SendKeys(data[4]);
                Thread.Sleep(500);
            }
            else if (Tipo == 2)
            {
                Thread.Sleep(500);
                ElementList1[0].Clear();
                Thread.Sleep(500);
                ElementList1[0].SendKeys(data[5]);
                Thread.Sleep(500);
            }
        }
        public static List<string> PreliminarKNmliyca(WindowsDriver<WindowsElement> desktopSession, string BanderaPreli, string file, string pdf1)
        {
            List<string> errors = new List<string> { };
            List<string> ErrorValidacion = new List<string> { };
            for(int i = 0; i < 2; i++)
            {
                FuncionesGlobales.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1);
                Thread.Sleep(1000);
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TFrmOpc");
                if(i == 1)
                {
                    rootSession.Mouse.MouseMove(rootSession.FindElementByName("Totalizado").Coordinates);
                    rootSession.Mouse.Click(null);
                    Thread.Sleep(500);
                }
                rootSession.Mouse.MouseMove(rootSession.FindElementByName("Aceptar").Coordinates);
                rootSession.Mouse.Click(null);
                Thread.Sleep(500);
                errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
                if (errors.Count > 0) { foreach (string er in errors) { ErrorValidacion.Add(er); } }
            }
            return ErrorValidacion;
        }
    }
}
