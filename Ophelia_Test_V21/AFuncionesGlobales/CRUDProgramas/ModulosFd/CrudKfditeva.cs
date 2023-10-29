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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosFd
{
    class CrudKfditeva
    {

        public static void KFdItevaCRUD(WindowsDriver<WindowsElement> desktopSession, List<string> variables, string edit, int flag) {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> editFields = desktopSession.FindElementsByClassName("TDBEdit");
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> radioButtons = desktopSession.FindElementsByClassName("TGroupButton");
            

            if (flag == 0)
            {

                desktopSession.Mouse.MouseMove(editFields[2].Coordinates);
                desktopSession.Mouse.Click(null);
                editFields[2].Clear();
                desktopSession.Keyboard.SendKeys(variables[0]);

                foreach (var rad in radioButtons)
                {
                    if (variables.Contains(rad.Text))
                    {
                        desktopSession.Mouse.MouseMove(rad.Coordinates);
                        desktopSession.Mouse.Click(null);
                        Thread.Sleep(1000);
                    }
                }
            }
            else {
                foreach (var rad in radioButtons) {
                    if (rad.Text == edit) {
                        desktopSession.Mouse.MouseMove(rad.Coordinates);
                        Thread.Sleep(1000);
                        desktopSession.Mouse.Click(null);
                        Thread.Sleep(1000);
                    }
                }
            }
        }


        public static void KFdItevaCRUDDet1(WindowsDriver<WindowsElement> desktopSession, List<string> variables, int flag, int descartar) {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> grids = desktopSession.FindElementsByClassName("TDBGrid");
            
            
            desktopSession.Mouse.MouseMove(grids[1].Coordinates, grids[1].Size.Width/4, grids[1].Size.Height/5);
            Thread.Sleep(1000);
            desktopSession.Mouse.DoubleClick(null);
            rootSession = newFunctions_4.RootSessionNew();
            Thread.Sleep(2000);
            WebDriverWait error = new WebDriverWait(rootSession, TimeSpan.FromSeconds(3));
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> memos = rootSession.FindElementsByClassName("TDBMemo");
            rootSession.Mouse.MouseMove(memos[0].Coordinates);
            rootSession.Mouse.Click(null);
            memos[0].Clear();
            if (flag == 0)
            {
                rootSession.Keyboard.SendKeys(variables[0]);
            }
            else {
                rootSession.Keyboard.SendKeys(variables[variables.Count-1]);
            }
            Thread.Sleep(2000);
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> menus = rootSession.FindElementsByClassName("TKNavegador");
            WindowsElement boton = null;
            for (int i = 0; i < menus[0].Size.Width; i += 20) {
                rootSession.Mouse.MouseMove(menus[0].Coordinates, menus[0].Size.Width/3+i, menus[0].Size.Height / 2);
                try {
                    if (descartar == 0)
                    {
                        boton = rootSession.FindElementByName("Aplicar");
                    }
                }
                catch (Exception e) { 
                
                }
                if (boton != null || descartar != 0)
                {
                    Thread.Sleep(1000);
                    WindowsElement cerrar = rootSession.FindElementByName("Cerrar");
                    rootSession.Mouse.MouseMove(cerrar.Coordinates);
                    rootSession.Mouse.Click(null);
                    break;
                }
            }

        }

        public static void KFdItevaCRUDDet2(WindowsDriver<WindowsElement> desktopSession, List<string> variables, int flag) {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> grids = desktopSession.FindElementsByClassName("TDBGrid");


            if (flag == 0)
            {
                desktopSession.Mouse.MouseMove(grids[0].Coordinates, grids[0].Size.Width / 4, grids[0].Size.Height / 5);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(variables[0]);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(variables[1]);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);

                desktopSession.Mouse.MouseMove(grids[0].Coordinates, Convert.ToInt32(grids[0].Size.Width / 1.9), grids[0].Size.Height / 5);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(grids[0].Coordinates, grids[0].Size.Width / 2, grids[0].Size.Height / 5);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(variables[2]);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1000);
            }
            else {
                desktopSession.Mouse.MouseMove(grids[0].Coordinates, Convert.ToInt32(grids[0].Size.Width / 1.9), grids[0].Size.Height / 5);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(grids[0].Coordinates, grids[0].Size.Width / 2, grids[0].Size.Height / 5);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(variables[variables.Count-1]);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1000);
            }

        }

    }
}
