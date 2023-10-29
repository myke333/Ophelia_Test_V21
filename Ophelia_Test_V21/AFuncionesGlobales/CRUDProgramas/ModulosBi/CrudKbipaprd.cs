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
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosBi
{
    class CrudKbipaprd : FuncionesVitales
    {
        static public WindowsDriver<WindowsElement> RootSession()
        {

            var appCapabilities = new AppiumOptions();
            appCapabilities.AddAdditionalCapability("app", "Root");
            var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);
            return ApplicationSession;
        }
        static public WindowsDriver<WindowsElement> ReloadSession(WindowsDriver<WindowsElement> session, string parametro)
        {
            try
            {

                var ApplicationWindow = session.FindElementByClassName(parametro);
                var ApplicationSessionHandle = ApplicationWindow.GetAttribute("NativeWindowHandle");
                ApplicationSessionHandle = (int.Parse(ApplicationSessionHandle)).ToString("x"); //Convert to hex


                var appCapabilities = new AppiumOptions();
                //appCapabilities.AddAdditionalCapability("app", "Root");
                appCapabilities.AddAdditionalCapability("appTopLevelWindow", ApplicationSessionHandle);
                var ApplicationSession = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:4723"), appCapabilities);

                string ses = ApplicationSession.PageSource;
                return ApplicationSession;
            }
            catch (Exception ex) { return null; }
        }
        public static void CRUD(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> crudVariables, String file)
        {
            List<string> errorMessages = new List<string>();

            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElementButton = desktopSession.FindElementsByClassName("TGroupButton");
            var ElementEdit = desktopSession.FindElementsByClassName("TDBGrid");

            if (tipo == 1)
            {

                {
                    ElementList[4].SendKeys(crudVariables[0]);
                    ElementList[3].SendKeys(crudVariables[1]);
                    ElementList[2].SendKeys(crudVariables[2]);
                    foreach (var elem in ElementButton)
                    {
                        if (elem.Text.Contains(crudVariables[3])  )
                        {
                            elem.Click();
                        }
                        Thread.Sleep(500);
                    }
                    
                }

            }
            
            else if(tipo==2)
            {
                desktopSession.Mouse.MouseMove(ElementEdit[1].Coordinates, 27, 27);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                desktopSession.Keyboard.SendKeys(crudVariables[7]);
                              
            }
            else if (tipo == 3)
            {
                desktopSession.Mouse.MouseMove(ElementEdit[1].Coordinates, 54, 29);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(6000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                desktopSession.Keyboard.SendKeys(crudVariables[6]);

               
            }
            else if(tipo==4)
            {

                desktopSession.Mouse.MouseMove(ElementEdit[0].Coordinates, 27, 27);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(6000);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);

                desktopSession.Keyboard.SendKeys(crudVariables[5]);

            }

            else if (tipo == 5)
            {
                desktopSession.Mouse.MouseMove(ElementEdit[0].Coordinates, 27,27);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                desktopSession.Keyboard.SendKeys(crudVariables[8]);
             
            }
            else 
            {
                ElementList[3].Clear();
                ElementList[3].SendKeys(crudVariables[4]);
            }
        }

        
    }
}
