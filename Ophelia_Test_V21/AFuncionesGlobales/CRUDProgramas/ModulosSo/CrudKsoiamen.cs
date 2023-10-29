using OpenQA.Selenium.Appium.Windows;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosSo
{
    class CrudKsoiamen
    {
        public static void CrudKSoiamen(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            if (Tipo == 1)
            {
                ElementList[2].SendKeys(data[0]);
                Thread.Sleep(500);
                ElementList[3].SendKeys(data[1]);
                Thread.Sleep(500);
            }
            else if (Tipo == 2)
            {
                ElementList[3].Clear();
                Thread.Sleep(500);
                ElementList[3].SendKeys(data[2]);
                Thread.Sleep(500);
            }
        }
        public static void CrudKSoiamenDetalle(WindowsDriver<WindowsElement> desktopSession, int Tipo, List<string> data)
        {
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 66, 26);
            desktopSession.Mouse.Click(null);
            Thread.Sleep(200);

            if (Tipo == 1)
            {
                ElementList[0].SendKeys(data[0]);
                Thread.Sleep(200);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(200);
                ElementList[0].SendKeys(data[1]);
                Thread.Sleep(200);
            }
            else if (Tipo == 2)
            {
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(200);
                ElementList[0].Clear();
                Thread.Sleep(200);
                ElementList[0].SendKeys(data[2]);
            }
        }
        public static Dictionary<string, Point> findnameDetalle(WindowsDriver<WindowsElement> desktopSession, int offset, int indiceBarra)
        {
            WindowsDriver<WindowsElement> rootSession = null;
            List<string> errorMessages = new List<string>();
            Dictionary<string, Point> coordenadasBotones = new Dictionary<string, Point>();
            List<string> botones = new List<string>() { "Nuevo", "Borrar", "Editar", "Aplicar", "Descartar" };
            Point coord = new Point();
            var ElementList = NavClass(desktopSession);
            ////Debugger.Launch();
            int x = ElementList[indiceBarra].Size.Width / 2;
            int y = ElementList[indiceBarra].Size.Height / 2;
            int nameCounter = 0;
            WindowsElement boton;
            desktopSession.Mouse.MouseMove(ElementList[indiceBarra].Coordinates, x + 20, y + 50);
            while (nameCounter < botones.Count && x < ElementList[indiceBarra].Size.Width)
            {
                desktopSession.Mouse.MouseMove(ElementList[indiceBarra].Coordinates, x, y);

                Thread.Sleep(300);

                boton = desktopSession.FindElementByName(botones[nameCounter]);

                if (boton != null)
                {
                    coord.X = x;
                    coord.Y = ElementList[indiceBarra].Size.Height / 2;
                    if (!coordenadasBotones.ContainsKey(botones[nameCounter]))
                    {
                        coordenadasBotones.Add(botones[nameCounter], coord);
                    }
                    if (nameCounter < botones.Count)
                    {
                        nameCounter++;
                    }
                }

                x += offset;
            }


            return coordenadasBotones;
        }
        public static System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> NavClass(WindowsDriver<WindowsElement> desktopSession)
        {
            var ElementList = desktopSession.FindElementsByClassName("TKNavegador");
            Debug.WriteLine($"Navigators find: {ElementList.Count}");
            return ElementList;
        }
    }
}
