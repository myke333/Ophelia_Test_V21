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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosRl
{
    class CrudKrledfor
    {
        public static void KRlEdforCRUD(WindowsDriver<WindowsElement> desktopSession, List<string> variables, int flag) {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> memoFields = desktopSession.FindElementsByClassName("TDBMemo");
            
            if (flag == 0)
            {
                List<int> lupasOff = new List<int>() { 2, 3, 4};
                newFunctions_4.LupaDinamicaDiscriminatoria(desktopSession, variables, lupasOff);
                Thread.Sleep(2000);

                memoFields[0].Click();
                memoFields[0].Clear();
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(variables[2]);
                Thread.Sleep(1000);

                newFunctions_4.changeWindow(desktopSession, "Datos Adicionales");
                Thread.Sleep(1000);

                rootSession = newFunctions_4.RootSessionNew();
                WebDriverWait close = new WebDriverWait(rootSession, TimeSpan.FromSeconds(5));
                Thread.Sleep(2000);
                System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> editFields = desktopSession.FindElementsByClassName("TDBEdit");

                editFields[10].Click();
                editFields[10].Clear();
                desktopSession.Keyboard.SendKeys(variables[variables.Count - 2]);

                newFunctions_4.changeWindow(desktopSession, "Datos Básicos");
                Thread.Sleep(1000);

            }
            else {
                memoFields[0].Click();
                memoFields[0].Clear();
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(variables[variables.Count-1]);
                Thread.Sleep(1000);
            }
        }

        public static void KRlEdforCRUDDet1(WindowsDriver<WindowsElement> desktopSession, List<string> variables, int flag) {
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

                desktopSession.Mouse.MouseMove(grids[0].Coordinates, grids[0].Size.Width / 15, grids[0].Size.Height / 5);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
                rootSession = PruebaCRUD.RootSession();
                WebDriverWait close = new WebDriverWait(rootSession, TimeSpan.FromSeconds(5));
                Thread.Sleep(1000);
                System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> fieldDots = rootSession.FindElementsByClassName("TDBGridInplaceEdit");
                close.Until(res => fieldDots[0].Displayed);
                rootSession.Mouse.MouseMove(fieldDots[1].Coordinates, fieldDots[1].Size.Width - 10, fieldDots[1].Size.Height / 2);
                Thread.Sleep(1000);
                rootSession.Mouse.Click(null);
                Thread.Sleep(1000);


                rootSession = PruebaCRUD.RootSession();
                rootSession = PruebaCRUD.ReloadSession(rootSession, "TCaptura");
                var campo = rootSession.FindElementsByClassName("TEdit");
                ////Debugger.Launch();
                rootSession.Mouse.MouseMove(campo[1].Coordinates, 10, 10);
                rootSession.Mouse.Click(null);
                rootSession.Keyboard.SendKeys(variables[0]);
                //Aqui se puede enviar un valor

                var btn = rootSession.FindElementsByClassName("TBitBtn");
                ////Debugger.Launch();
                foreach (var boton in btn)
                {
                    if (boton.Text == "Aceptar")
                    {
                        rootSession.Mouse.MouseMove(boton.Coordinates, 10, 10);
                        rootSession.Mouse.Click(null);
                    }
                }
                Thread.Sleep(3000);
                rootSession = PruebaCRUD.RootSession();
                close = new WebDriverWait(rootSession, TimeSpan.FromSeconds(5));
                desktopSession.Mouse.MouseMove(grids[0].Coordinates, grids[0].Size.Width / 20, grids[0].Size.Height / 5);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);

                for (int i = 1; i < variables.Count - 1; i++) {
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    Thread.Sleep(1000);
                    desktopSession.Keyboard.SendKeys(variables[i]);
                    Thread.Sleep(1000);
                }
            }
            else {
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(variables[variables.Count-1]);
                Thread.Sleep(1000);
            }

        }

    }
}
