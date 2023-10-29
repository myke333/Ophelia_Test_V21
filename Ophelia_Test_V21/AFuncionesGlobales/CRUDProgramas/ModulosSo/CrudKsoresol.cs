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
using System.IO;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosSo
{
    class CrudKsoresol
    {
        public static void KSoResolCRUD(WindowsDriver<WindowsElement> desktopSession, List<string> variables, int flag) {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
            {
                Timeout = TimeSpan.FromSeconds(120),
                PollingInterval = TimeSpan.FromSeconds(1),
            };
            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
            WindowsDriver<WindowsElement> rootSession = null;
            System.Collections.ObjectModel.ReadOnlyCollection<WindowsElement> editFields = desktopSession.FindElementsByClassName("TDBEdit");
            if (flag == 0)
            {
                List<int> fieldIndex = new List<int>() { 4, 9 };
                for (int i = 0; i < fieldIndex.Count; i++) {
                    editFields[fieldIndex[i]].Click();
                    editFields[fieldIndex[i]].Clear();
                    Thread.Sleep(1000);
                    editFields[fieldIndex[i]].SendKeys(variables[i]);
                }
                Thread.Sleep(1000);
            }
            else {
                Thread.Sleep(1000);
                editFields[9].Click();
                editFields[9].Clear();
                Thread.Sleep(1000);
                editFields[9].SendKeys(variables[variables.Count-1]);
                Thread.Sleep(1000);
            }

        }

        public static void KSoResolCRUDDet1(WindowsDriver<WindowsElement> desktopSession, List<string> variables, int flag) {
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
                desktopSession.Mouse.MouseMove(grids[0].Coordinates, 25, grids[0].Size.Height/6);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                for (int i = 0; i < variables.Count - 1; i++) {
                    desktopSession.Keyboard.SendKeys(variables[i]);
                    if (i < variables.Count - 2) desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    Thread.Sleep(1000);
                }
            }
            else {
                Thread.Sleep(1000);
                desktopSession.Mouse.MouseMove(grids[0].Coordinates, 25, grids[0].Size.Height / 6);
                Thread.Sleep(1000);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                Thread.Sleep(1000);
                desktopSession.Keyboard.SendKeys(variables[variables.Count - 1]);
                Thread.Sleep(1000);
            }

        }

		static public void InsertarDatos(WindowsDriver<WindowsElement> desktopSession, int bandera, List<string> crudVariables, string file)
		{

			var btnTDVI = desktopSession.FindElementsByClassName("TScrollBox");
			if (bandera == 0)
			{

				///lupa secuencial

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 666, 35);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 780, 72);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);


				WindowsDriver<WindowsElement> RootSession1 = null;
				RootSession1 = PruebaCRUD.RootSession();



				RootSession1.Keyboard.SendKeys(crudVariables[0]);
				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1076, 108);
				desktopSession.Mouse.Click(null);

				///lupa seccional 

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 551, 132);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 780, 72);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);


				WindowsDriver<WindowsElement> RootSession2 = null;
				RootSession2 = PruebaCRUD.RootSession();



				RootSession2.Keyboard.SendKeys(crudVariables[1]);
				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1076, 108);
				desktopSession.Mouse.Click(null);

				///lupa dependencia

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 334, 174);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 780, 72);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);


				WindowsDriver<WindowsElement> RootSession3 = null;
				RootSession3 = PruebaCRUD.RootSession();



				RootSession3.Keyboard.SendKeys(crudVariables[2]);
				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1076, 108);
				desktopSession.Mouse.Click(null);

				//centro de costos 

				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 663, 174);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 780, 72);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);


				WindowsDriver<WindowsElement> RootSession4 = null;
				RootSession4 = PruebaCRUD.RootSession();



				RootSession4.Keyboard.SendKeys(crudVariables[3]);
				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1076, 108);
				desktopSession.Mouse.Click(null);
			}


			else
			{
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 663, 174);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 780, 72);
				desktopSession.Mouse.Click(null);
				Thread.Sleep(1000);


				WindowsDriver<WindowsElement> RootSession4 = null;
				RootSession4 = PruebaCRUD.RootSession();



				RootSession4.Keyboard.SendKeys(crudVariables[4]);
				Thread.Sleep(1000);
				desktopSession.Mouse.MouseMove(btnTDVI[0].Coordinates, 1076, 108);
				desktopSession.Mouse.Click(null);

			}
		}

	}
}
