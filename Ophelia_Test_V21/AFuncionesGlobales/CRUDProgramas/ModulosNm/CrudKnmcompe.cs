using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Ophelia_Test_V21.AFuncionesGlobales;
using OpenQA.Selenium;
using System.Drawing;
using System.Windows.Forms;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm
{
    class CrudKnmcompe : PruebaCRUD
    {
        public static void CRUDKnmcompe(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data)
        {
            List<string> errorMessages = new List<string>();
            var ElementList1 = desktopSession.FindElementsByClassName("TDBEdit");
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            if (tipo == 1)
            {
                LupaDinamica(rootSession, data);
                ElementList1[9].SendKeys(data[2]);
                ElementList1[8].SendKeys(data[3]);
                ElementList1[6].SendKeys(data[4]);
                ElementList1[7].SendKeys(data[5]);
            }
            else
            {
                ElementList1[6].Clear();
                ElementList1[6].SendKeys(data[5]);
            }
        }

        public static List<string> PreliminarKnmcompe(WindowsDriver<WindowsElement> desktopSession, string file, string pdf1, string BanderaPreli)
        {
            List<string> errorMessages = new List<string>();
            List<string> errors = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            WindowsDriver<WindowsElement> RootSession()
            {
                var appCapabilities = new AppiumOptions();
                appCapabilities.AddAdditionalCapability("app", "Root");
                var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);
                return ApplicationSession;
            }
            WindowsDriver<WindowsElement> ReloadSession(WindowsDriver<WindowsElement> session, string parametro)
            {
                try
                {
                    var ApplicationWindow = session.FindElementByClassName(parametro);
                    var ApplicationSessionHandle = ApplicationWindow.GetAttribute("NativeWindowHandle");
                    ApplicationSessionHandle = (int.Parse(ApplicationSessionHandle)).ToString("x"); //Convert to hex


                    var appCapabilities = new AppiumOptions();
                    //appCapabilities.AddAdditionalCapability("app", "Root");
                    appCapabilities.AddAdditionalCapability("appTopLevelWindow", ApplicationSessionHandle);
                    var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);

                    string ses = ApplicationSession.PageSource;
                    return ApplicationSession;
                }
                catch (Exception ex) { return null; }
            }
            rootSession = RootSession();
            rootSession = ReloadSession(rootSession, "TFrmOpc");
            var bot = rootSession.FindElementsByClassName("TGroupButton");
            bot[bot.Count - 1].Click();
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Opción Preliminar", true, file);
            Thread.Sleep(500);
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
            Thread.Sleep(500);
            if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
            for (int i = bot.Count - 2; i >= 0; i--)
            {
                Thread.Sleep(1000);
                pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                FuncionesGlobales.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1);
                if (i == 0)
                {
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                }else
                {
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                    //rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                }
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(500);
                newFunctions_4.ScreenshotNuevo("Opción Preliminar", true, file);
                Thread.Sleep(1000);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1000);
                errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
                if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
            }
            return errorMessages;
        }
    }
}
