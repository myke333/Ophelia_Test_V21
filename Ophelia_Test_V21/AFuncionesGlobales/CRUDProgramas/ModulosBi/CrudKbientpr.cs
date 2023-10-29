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
    class CrudKbientpr : PruebaCRUD

    {
        public static void CRUDKbientpr(WindowsDriver<WindowsElement> desktopSession, int tipo, string fecEntrega, string observa, List<string> data)
        {
            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementByClassName("TDBMemo");
            var ElementList1 = desktopSession.FindElementsByClassName("TDBEdit");
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            if (tipo == 1)
            {
                LupaDinamica(rootSession, data);
                ElementList.Clear();
                ElementList.SendKeys(fecEntrega);
                ElementList1[20].SendKeys(fecEntrega);
            }
            else
            {
                ElementList.Clear();
                ElementList.SendKeys(observa);
            }
        }
        public static List<string> PreliminarKbientpr(WindowsDriver<WindowsElement> desktopSessio, string file, string pdf1)
        {
            List<string> errorMessages = new List<string>();
            List<string> errors = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            var vent = rootSession.FindElementByClassName("TFrmReports");
            vent.Click();
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Opción Preliminar", true, file);
            Thread.Sleep(500);
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
            if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
            return errorMessages;
        }
    }
}
