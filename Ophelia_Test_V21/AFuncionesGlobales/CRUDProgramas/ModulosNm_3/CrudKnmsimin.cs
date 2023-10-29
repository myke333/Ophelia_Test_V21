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
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosNm_3
{
    class CrudKnmsimin : PruebaCRUD
    {
        public static void CRUDKnmsimin(WindowsDriver<WindowsElement> desktopSession, List<string> data)
        {
            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementsByClassName("TEdit");
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            ElementList[0].SendKeys(data[2]);
            ElementList[2].SendKeys(data[0]);
            ElementList[1].SendKeys(data[1]);
        }

        public static List<string> BotonAceptar(WindowsDriver<WindowsElement> desktopSession)
        {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));

            ////////////////////////////////////////////////
            WebDriverWait printScreen;
            WindowsElement printerField;
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> botones = desktopSession.FindElementsByClassName("TBitBtn");
            List<string> errors = new List<string>();
            List<string> errorMessages = new List<string>();

            foreach (var elem in botones)
            {
                if (elem.Text == "Aceptar")
                {
                    elem.Click();
                    Thread.Sleep(2000);
                    PruebaCRUD.cerrarBorrar(desktopSession);
                    desktopSession.Keyboard.SendKeys(Keys.Enter);
                    PruebaCRUD.cerrarBorrar(desktopSession);
                    desktopSession.Keyboard.SendKeys(Keys.Enter);
                }
            }

            return errorMessages;

        }


    }
}
