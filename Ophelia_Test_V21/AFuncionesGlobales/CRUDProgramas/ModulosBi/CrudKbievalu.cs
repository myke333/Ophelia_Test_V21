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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosBi
{
    class CrudKbievalu : PruebaCRUD
    {
        public static void PreliminarKbievalu(WindowsDriver<WindowsElement> desktopSession, string file, string pdf1, int opc)
        {
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            var vent = rootSession.FindElementByClassName("TFrmImpRep");
            vent.Click();
            //Thread.Sleep(7000);
            if (opc == 1)
            {
                Thread.Sleep(7000);
                var bot = rootSession.FindElementsByClassName("TGroupButton");
                bot[0].Click();
            }
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            //Thread.Sleep(7000);
            newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
        }
    }
}
