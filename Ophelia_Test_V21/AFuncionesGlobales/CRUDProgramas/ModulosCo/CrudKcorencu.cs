using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosCo
{
    class CrudKcorencu : PruebaCRUD
    {
        public static void CRUDKcorencu(WindowsDriver<WindowsElement> desktopSession, List<string> data)
        {
            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementsByClassName("TEdit");
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            ElementList[2].SendKeys(data[0]);
            ElementList[1].SendKeys(data[1]);
        }

        public static void PreliminarKcorencu(WindowsDriver<WindowsElement> desktopSession, string file, string pdf1, string BanderaPreli, List<string> data2, List<string> data3)
        {
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
            rootSession = ReloadSession(rootSession, "TFrmMenu");

            var bot = rootSession.FindElementsByClassName("TGroupButton");
            bot[bot.Count - 1].Click();
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
            for (int i = bot.Count - 2; i >= 8; i--)
            {
                pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                FuncionesGlobales.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1);

                if (i == 14 || i == 13)
                {
                    bot[i].Click();
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                    newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
                }
                else if (i == 12)
                {
                    bot[i].Click();
                    for (int j = 7; j >= 2; j--)
                    {
                        pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                        bot[j].Click();
                        Thread.Sleep(1000);
                        if (j == 6)
                        {
                            LupaDinamica2(rootSession, data2);
                        }
                        rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                        newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
                        FuncionesGlobales.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1);
                    }
                }
                else if (i == 11 || i == 10 || i == 9)
                {
                    bot[i].Click();
                    if (i == 9)
                    {
                        LupaDinamica2(rootSession, data3);
                    }
                    for (int j = 1; j >= 0; j--)
                    {
                        pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                        bot[j].Click();
                        rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                        newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
                        FuncionesGlobales.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1);
                    }


                }
                else if (i == 8)
                {
                    bot[i].Click();
                    int cont = 0;
                    for (int j = 1; j >= 0; j--)
                    {
                        pdf1 = @"C:\Reportes\ReportePDF1_" + "Preliminar" + "_" + Hora();
                        bot[j].Click();
                        rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                        rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                        rootSession = RootSession();
                        rootSession = ReloadSession(rootSession, "TFrmMenu");
                        var eleme = rootSession.FindElementsByClassName("#32770");
                        eleme[0].Click();
                        rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                        newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
                        if (cont == 0)
                        {
                            FuncionesGlobales.ReportePreliminar(desktopSession, BanderaPreli, file, pdf1);
                            cont++;
                        }
                        
                    }

                }
            }
        }


    }
}
