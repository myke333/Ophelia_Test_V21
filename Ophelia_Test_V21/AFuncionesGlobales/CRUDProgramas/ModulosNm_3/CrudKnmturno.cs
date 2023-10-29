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
    class CrudKnmturno : PruebaCRUD
    {
        public static void CRUDKnmturno(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data)
        {
            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            if (tipo == 1)
            {
                ElementList[6].SendKeys(data[0]);
                ElementList[5].SendKeys(data[1]);
            }
            else
            {
                ElementList[5].Clear();
                ElementList[5].SendKeys(data[2]);
            }
        }

        public static void CRUDDetalle(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data)
        {
            
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            if (tipo == 1)
            {
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 24, 13);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(data[3]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(data[4]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(data[5]);

            }
            else
            {
                desktopSession.Keyboard.SendKeys(data[6]);
            }
        }
    }
}
