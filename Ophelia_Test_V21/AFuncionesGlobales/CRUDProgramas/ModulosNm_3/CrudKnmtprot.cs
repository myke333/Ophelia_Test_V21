using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_3
{
    class CrudKnmtprot : PruebaCRUD
    {
        public static void CRUDKnmtprot(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data)
        {
            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            if (tipo == 1)
            {
                ElementList[4].SendKeys(data[0]);
                ElementList[3].SendKeys(data[1]);
            }
            else
            {
                ElementList[3].Clear();
                ElementList[3].SendKeys(data[2]);
            }
        }

        public static void CRUDDetalle(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data)
        {

            var ElementList = desktopSession.FindElementsByClassName("TKactusDBgrid");
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            if (tipo == 1)
            {
                desktopSession.Mouse.MouseMove(ElementList[1].Coordinates, 18, 12);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(data[3]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(data[4]);

            }
            else
            {
                desktopSession.Keyboard.SendKeys(data[5]);
            }
        }

        public static List<string> PreliminarKnmtprot(WindowsDriver<WindowsElement> desktopSession, string file, string pdf1, string BanderaPreli)
        {
            List<string> errorMessages = new List<string>();
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

            try
            {
            rootSession = RootSession();
            rootSession = ReloadSession(rootSession, "TFrmOpciones");
            var bot = rootSession.FindElementsByClassName("TGroupButton");
            bot[bot.Count - 1].Click();
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1000);
                newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
            for (int i = bot.Count - 2; i >= 0; i--)
            {
                pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                FuncionesGlobales.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1);
                bot[i].Click();
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    Thread.Sleep(1000);
                    newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
                }
            }
            catch (Exception ex)
            {
                errorMessages.Add($"Error de Appium: {ex.ToString()}");
            }


            return errorMessages;
        }        
    }
}
