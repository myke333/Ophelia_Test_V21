using OpenQA.Selenium.Appium.Windows;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosCo
{
    class CrudKcovaria
    {
        public static void CrudKCovaria(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementList1 = desktopSession.FindElementsByClassName("TDBMemo");
            if (Tipo == 1)
            {
                ElementList[5].SendKeys(data[0]);
                ElementList[4].SendKeys(data[1]);
                ElementList[3].SendKeys(data[2]);
                Thread.Sleep(500);
                desktopSession.Mouse.MouseMove(desktopSession.FindElementByClassName("TDBLookupComboBox").Coordinates);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(200);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Down);
                Thread.Sleep(200);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(500);
                ElementList1[1].SendKeys(data[3]);
            }
            else if (Tipo == 2)
            {
                ElementList[4].Clear();
                Thread.Sleep(500);
                ElementList[4].SendKeys(data[4]);
                Thread.Sleep(500);
            }
        }
        public static void CrudKCovariaDetalle1(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            var ElementList1 = desktopSession.FindElementsByClassName("TDBMemo");
            if (Tipo == 1)
            {
                ElementList1[0].SendKeys(data[0]);
            }
            else if (Tipo == 2)
            {
                ElementList1[0].Clear();
                Thread.Sleep(500);
                ElementList1[0].SendKeys(data[1]);
                Thread.Sleep(500);
            }
        }
        public static void CrudKCovariaDetalle2(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            var ElementList1 = desktopSession.FindElementsByClassName("TDBGrid");
            desktopSession.Mouse.MouseMove(ElementList1[0].Coordinates, 66, 26);
            desktopSession.Mouse.Click(null);
            if (Tipo == 1)
            {
                ElementList1[0].SendKeys(data[0]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                ElementList1[0].SendKeys(data[1]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                ElementList1[0].SendKeys(data[2]);
            }
            else if (Tipo == 2)
            {
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                ElementList1[0].Clear();
                Thread.Sleep(500);
                ElementList1[0].SendKeys(data[3]);
                Thread.Sleep(500);
            }
        }
    }
}
