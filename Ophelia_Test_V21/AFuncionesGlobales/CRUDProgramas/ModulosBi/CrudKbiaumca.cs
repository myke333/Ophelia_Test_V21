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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosBi
{
    class CrudKbiaumca : PruebaCRUD
    {
        public static List<string> BotonAceptar(WindowsDriver<WindowsElement> desktopSession, string file)
        {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> botones = desktopSession.FindElementsByClassName("TBitBtn");
            List<string> errors = new List<string>();
            List<string> errorMessages = new List<string>();
            foreach (var elem in botones)
            {
                if (elem.Text == "Aceptar")
                {
                    elem.Click();
                    string navigator = AbrirPrograma.SelectNavigator();
                    if (navigator == "IE")
                    {
                        rootSession = RootSession();
                        rootSession = ReloadSession(rootSession, "IEFrame");
                        var allFrame = rootSession.FindElementsByClassName("IEFrame");
                        new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset(allFrame[0].Size.Width / 2, allFrame[0].Size.Height / 2).DoubleClick().Click().Perform();
                        rootSession = RootSession();
                        for (int i = 0; i < 3; i++)
                        {
                            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                            Thread.Sleep(200);
                        }
                    }
                    else
                    {
                        Thread.Sleep(2000);
                        rootSession = RootSession();
                        Thread.Sleep(2000);
                        newFunctions_4.ScreenshotNuevo("Aceptar", true, file);
                        Thread.Sleep(1000);
                        rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                        //new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset(allFrame[0].Size.Width / 2, allFrame[0].Size.Height / 2).DoubleClick().Click().Perform();
                        rootSession = RootSession();
                        Thread.Sleep(5000);
                        for (int i = 0; i < 3; i++)
                        {
                            Thread.Sleep(1000);
                            newFunctions_4.ScreenshotNuevo("Aceptar", true, file);
                            Thread.Sleep(1000);
                            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                            Thread.Sleep(2000);
                        }
                        /*Debugger.Launch();
                        //rootSession = RootSession();
                        //rootSession = ReloadSession(rootSession, "#32769");
                        var allFrame = rootSession.FindElementsByClassName("Chrome_WidgetWin_1");
                        new Actions(rootSession).MoveToElement(allFrame[0], 0, 0).MoveByOffset(allFrame[0].Size.Width / 2, allFrame[0].Size.Height / 2).DoubleClick().Click().Perform();
                        Thread.Sleep(2000);
                        rootSession = RootSession();
                        for (int i = 0; i < 3; i++)
                        {
                            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                            Thread.Sleep(200);
                        }*/
                    }


                }
            }

            return errorMessages;

        }
    }
}
