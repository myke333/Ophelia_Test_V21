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


namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosCo
{
    class CrudKcoparam
    {
        public static void KCoParamCRUD(WindowsDriver<WindowsElement> desktopSession, List<string> variables, int flag) {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> comboBoxes = desktopSession.FindElementsByClassName("TComboBox");
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> editFields = desktopSession.FindElementsByClassName("TDBEdit");

            if (flag == 0)
            {
                Thread.Sleep(6000);
                desktopSession.Mouse.MouseMove(comboBoxes[0].Coordinates);
                Thread.Sleep(2000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(variables[variables.Count-4]);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                Thread.Sleep(1000);

                //desktopSession.Mouse.MouseMove(editFields[5].Coordinates);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(1000);
                //editFields[5].Clear();
                //desktopSession.Keyboard.SendKeys(variables[variables.Count - 2]);
                //Thread.Sleep(1000);

                //desktopSession.Mouse.MouseMove(editFields[4].Coordinates);
                //desktopSession.Mouse.Click(null);
                //Thread.Sleep(1000);
                //editFields[4].Clear();
                //desktopSession.Keyboard.SendKeys(variables[variables.Count - 3]);
                //Thread.Sleep(1000);

                PruebaCRUD.LupaDinamica(desktopSession, variables);
                Thread.Sleep(1000);


            }
            else {
                List<int> lupasOff = new List<int>() { 0,2 };
                List<string> editValue = new List<string>() { variables[variables.Count - 1] };
                newFunctions_4.LupaDinamicaDiscriminatoria(desktopSession, editValue, lupasOff);
                Thread.Sleep(1000);
            }

        }


        public static void CheckBoxes(WindowsDriver<WindowsElement> desktopSession) {
           
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> checkBoxes = desktopSession.FindElementsByClassName("TDBCheckBox");
            List<string> boxNames = new List<string>() { "Género", "Cargo" };
            foreach (var box in checkBoxes) {
                if (boxNames.Contains(box.Text)) {
                    desktopSession.Mouse.MouseMove(box.Coordinates);
                    Thread.Sleep(1000);
                    desktopSession.Mouse.Click(null);
                    Thread.Sleep(1000);
                }
            }
            
        }


        public static void KCoParamCRUDDet1(WindowsDriver<WindowsElement> desktopSession, List<string> variables, int flag) {
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
                desktopSession.Mouse.MouseMove(grids[0].Coordinates, grids[0].Size.Width / 60, grids[0].Size.Height / 5);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(variables[0]);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(variables[1]);
                Thread.Sleep(3000);
            }
            else {
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(variables[variables.Count - 1]);
                Thread.Sleep(1000);
            }
        }

        public static void KCoParamCRUDDet2(WindowsDriver<WindowsElement> desktopSession, List<string> variables, int flag) {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1)
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> grids = desktopSession.FindElementsByClassName("TKactusDBgrid");
            if (flag == 0)
            {
                desktopSession.Mouse.MouseMove(grids[1].Coordinates, grids[1].Size.Width / 15, grids[1].Size.Height / 5);
                Thread.Sleep(3000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);

                rootSession = PruebaCRUD.RootSession();
                //var campo = rootSession.FindElementsByXPath("//Window[@ClassName=\"TCaptura\"][@Name=\" LISTA DE EMPLEADOS\"]/Edit[@ClassName=\"TEdit\"]");
                WebDriverWait close = new WebDriverWait(rootSession, TimeSpan.FromSeconds(5));
                Thread.Sleep(2000);
                System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> fieldDots = rootSession.FindElementsByClassName("TDBGridInplaceEdit");
                close.Until(res => fieldDots[0].Displayed);
                rootSession.Mouse.MouseMove(fieldDots[1].Coordinates, fieldDots[1].Size.Width - 10, fieldDots[1].Size.Height / 2);
                Thread.Sleep(3000);
                rootSession.Mouse.Click(null);
                Thread.Sleep(3000);


                rootSession = PruebaCRUD.RootSession();
                rootSession = PruebaCRUD.ReloadSession(rootSession, "TCaptura");
                //var campo = rootSession.FindElementsByXPath("//Window[@ClassName=\"TCaptura\"][@Name=\" LISTA DE EMPLEADOS\"]/Edit[@ClassName=\"TEdit\"]");
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

                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowLeft);
                Thread.Sleep(1000);
                rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);

                Thread.Sleep(3000);
                desktopSession.Mouse.MouseMove(grids[1].Coordinates, grids[1].Size.Width / 6, grids[1].Size.Height / 5);
                Thread.Sleep(3000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);

                rootSession = null;
                rootSession = PruebaCRUD.RootSession();
                //var campo = rootSession.FindElementsByXPath("//Window[@ClassName=\"TCaptura\"][@Name=\" LISTA DE EMPLEADOS\"]/Edit[@ClassName=\"TEdit\"]");
                close = new WebDriverWait(rootSession, TimeSpan.FromSeconds(5));
                Thread.Sleep(2000);
                fieldDots = rootSession.FindElementsByClassName("TDBGridInplaceEdit");
                close.Until(res => fieldDots[0].Displayed);
                rootSession.Mouse.MouseMove(fieldDots[1].Coordinates, fieldDots[1].Size.Width - 10, fieldDots[1].Size.Height / 2);
                Thread.Sleep(1000);
                rootSession.Mouse.Click(null);
                Thread.Sleep(3000);

                rootSession = PruebaCRUD.RootSession();
                rootSession = PruebaCRUD.ReloadSession(rootSession, "TCaptura");
                //var campo = rootSession.FindElementsByXPath("//Window[@ClassName=\"TCaptura\"][@Name=\" LISTA DE EMPLEADOS\"]/Edit[@ClassName=\"TEdit\"]");
                campo = rootSession.FindElementsByClassName("TEdit");
                ////Debugger.Launch();
                rootSession.Mouse.MouseMove(campo[1].Coordinates, 10, 10);
                rootSession.Mouse.Click(null);
                rootSession.Keyboard.SendKeys(variables[1]);
                //Aqui se puede enviar un valor

                btn = rootSession.FindElementsByClassName("TBitBtn");
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


            }
            else {


                Thread.Sleep(3000);
                desktopSession.Mouse.MouseMove(grids[1].Coordinates, grids[1].Size.Width / 6, grids[1].Size.Height / 5);
                Thread.Sleep(3000);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);

                rootSession = null;
                rootSession = PruebaCRUD.RootSession();
                //var campo = rootSession.FindElementsByXPath("//Window[@ClassName=\"TCaptura\"][@Name=\" LISTA DE EMPLEADOS\"]/Edit[@ClassName=\"TEdit\"]");
                WebDriverWait close = new WebDriverWait(rootSession, TimeSpan.FromSeconds(5));
                Thread.Sleep(2000);
                var fieldDots = rootSession.FindElementsByClassName("TDBGridInplaceEdit");
                close.Until(res => fieldDots[0].Displayed);
                rootSession.Mouse.MouseMove(fieldDots[1].Coordinates, fieldDots[1].Size.Width - 10, fieldDots[1].Size.Height / 2);
                Thread.Sleep(1000);
                rootSession.Mouse.Click(null);
                Thread.Sleep(3000);

                rootSession = PruebaCRUD.RootSession();
                rootSession = PruebaCRUD.ReloadSession(rootSession, "TCaptura");
                //var campo = rootSession.FindElementsByXPath("//Window[@ClassName=\"TCaptura\"][@Name=\" LISTA DE EMPLEADOS\"]/Edit[@ClassName=\"TEdit\"]");
                var campo = rootSession.FindElementsByClassName("TEdit");
                ////Debugger.Launch();
                rootSession.Mouse.MouseMove(campo[1].Coordinates, 10, 10);
                rootSession.Mouse.Click(null);
                rootSession.Keyboard.SendKeys(variables[variables.Count-1]);
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

            }


        }

    }
}
