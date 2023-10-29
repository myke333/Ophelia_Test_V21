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
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosAc
{
    class CrudKacpamaf:FuncionesVitales
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

        public static void CRUD(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data,string file)
        {
            var ElemenList = desktopSession.FindElementsByClassName("TDBEdit");
            var ElemenList2 = desktopSession.FindElementsByClassName("TDBCheckBox");
            if (tipo == 1) 
            {
                newFunctions_4.ScreenshotNuevo("Agregar Registro", true, file);
                foreach (var elem2 in ElemenList2)
                {
                    if(
                        elem2.Text!="Criterios de Desempeño" && 
                        elem2.Text!="Autoridad" &&
                        elem2.Text!="Leer Perfil de Planta" && 
                        elem2.Text!="Elementos de protección" && 
                        elem2.Text!="Actualización todos los Cargos Versión" && 
                        elem2.Text!="Compromisos" &&
                        elem2.Text!="No Requieren ser Consultadas" &&
                        elem2.Text!= "Deben ser Consultadas" &&
                        elem2.Text!= "Agregar Nombre de la  Variable" &&
                        elem2.Text!="Autonomía" &&
                        elem2.Text!="Conocimientos Básicos o Esenciales" &&
                        elem2.Text != "Leer Perfil de Cargo"
                      )
                    {
                        elem2.Click();
                    }
                }
                ElemenList[22].SendKeys(data[0]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(data[1]);

                ElemenList[15].SendKeys(data[2]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(data[3]);

                ElemenList[14].SendKeys(data[4]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(data[5]);

                ElemenList[7].SendKeys(data[6]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(data[7]);

                ElemenList[6].SendKeys(data[8]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(data[9]);
            }
            else
            {
                newFunctions_4.ScreenshotNuevo("Editar Registro", true, file);
                ElemenList[14].Click();
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(data[10]);
            }
        }
    }
}
