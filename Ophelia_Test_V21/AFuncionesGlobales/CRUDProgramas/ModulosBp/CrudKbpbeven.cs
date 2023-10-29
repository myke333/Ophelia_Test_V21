using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosBp
{
    class CrudKbpbeven
    {
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
        public static void CrudKBpbeven(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementList1 = desktopSession.FindElementsByClassName("TDBMemo");
            if (Tipo == 1)
            {
                ElementList[4].SendKeys(data[0]);
                Thread.Sleep(500);
                ElementList1[0].SendKeys(data[1]);
                Thread.Sleep(500);
            }
            else if (Tipo == 2)
            {
                ElementList1[0].Clear();
                Thread.Sleep(500);
                ElementList1[0].SendKeys(data[2]);
            }
        }
        public static void CrudKBpbevenDetalle(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, 66, 26);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(500);
            if (Tipo == 1)
            {
                ElementList[1].SendKeys(data[0]);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                ElementList[1].SendKeys(data[1]);
                Thread.Sleep(500);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                ElementList[1].SendKeys(data[2]);
            }
            else if (Tipo == 2)
            {
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(200);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(200);
                ElementList[1].Clear();
                Thread.Sleep(500);
                ElementList[1].SendKeys(data[3]);
            }
        }
        public static List<string> KBpbevenPreliminar(WindowsDriver<WindowsElement> desktopSession, string BanderaPreli, string pdf1, string file)
        {
            List<string> errors = new List<string> { };
            List<string> listErrors = new List<string> { };
            WindowsDriver<WindowsElement> rootSession;
            FuncionesGlobales.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1);
            Thread.Sleep(2000);
            rootSession = RootSession();
            rootSession = ReloadSession(rootSession, "TFrmRepOpc");
            var ElementList = rootSession.FindElementsByClassName("TGroupButton");
            for (int i = 1; i < ElementList.Count; i++)
            {
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TFrmRepOpc");
                Thread.Sleep(1000);
                ElementList = rootSession.FindElementsByClassName("TGroupButton");
                var aceptar = rootSession.FindElementByName("Aceptar");
                Thread.Sleep(1500);
                rootSession.Mouse.MouseMove(ElementList[i].Coordinates);
                Thread.Sleep(500);
                rootSession.Mouse.Click(null);
                Thread.Sleep(500);
                if(i == 4)
                {
                    var check = rootSession.FindElementsByClassName("TCheckBox");
                    for(int j = 0; j < check.Count; j++)
                    {
                        rootSession.Mouse.MouseMove(check[j].Coordinates);
                        Thread.Sleep(500);
                        rootSession.Mouse.Click(null);
                        Thread.Sleep(500);
                    }
                }
                rootSession.Mouse.MouseMove(aceptar.Coordinates);
                Thread.Sleep(500);
                rootSession.Mouse.Click(null);
                Thread.Sleep(2000);
                try
                {
                    errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
                    Thread.Sleep(1000);
                    if (errors.Count > 0) { foreach (string er in errors) { listErrors.Add(er); } }
                }
                catch
                {
                    PruebaCRUD.cerrarBorrar(desktopSession);
                    Thread.Sleep(2000);
                }
                if(i == ElementList.Count - 1)
                {
                    break;
                }
                FuncionesGlobales.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1);
                Thread.Sleep(2000);
            }
            return listErrors;
        }
    }
}
