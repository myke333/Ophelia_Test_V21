using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using System.Drawing;
using OpenQA.Selenium;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosSo
{
    class CrudKsoespec
    {
        public static void KSoEspecCRUD(WindowsDriver<WindowsElement> desktopSession, List<string> variables, int flag) {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1),
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> grids = desktopSession.FindElementsByClassName("TDBGrid");
            if (flag == 0)
            {
                desktopSession.Mouse.MouseMove(grids[0].Coordinates, grids[0].Size.Width/25, grids[0].Size.Height/6);
                Thread.Sleep(5000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);

                for (int i = 0; i < variables.Count - 1; i++) {
                    desktopSession.Keyboard.SendKeys(variables[i]);
                    Thread.Sleep(1000);
                    if (i < variables.Count - 2) desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                }

            }
            else {
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(variables[variables.Count-1]);
                Thread.Sleep(1000);
            }

        }

        public static Dictionary<string, Point> findname(WindowsDriver<WindowsElement> desktopSession, int offset, int indiceBarra)
        {
            WindowsDriver<WindowsElement> rootSession = null;
            List<string> errorMessages = new List<string>();
            Dictionary<string, Point> coordenadasBotones = new Dictionary<string, Point>();
            List<string> botones = new List<string>() { "Nuevo", "Borrar", "Editar", "Aplicar", "Descartar" };
            Point coord = new Point();
            var ElementList = PruebaCRUD.NavClass(desktopSession);
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

    }
}
