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
    class CrudKbiarbol:FuncionesVitales
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

        public static void CRUD(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string>crudVariables, string file)
        {
            List<string> errorMessages = new List<string>();

            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            Debug.WriteLine($"la cantidad de elementos es: {ElementList.Count}");
            int cont = 0;
           // List<string> variables = new List<string>() {"PRUEBACRUD", "3939" };
            int contadorcampos = 0;
            if (tipo == 1)
            {
                newFunctions_4.ScreenshotNuevo("Agregar Registro", true, file);
                foreach (var elem in ElementList)
                {
                    if (elem.Enabled == true)
                    {
                        elem.SendKeys(crudVariables[cont]);
                        cont += 1;
                    }
                    Thread.Sleep(500);
                    //Debug.WriteLine($"campos {contadorcampos} campo habilitado {cont}");
                }
                newFunctions_4.ScreenshotNuevo("Regstro Agregado", true, file);
            }
            else
            {
                newFunctions_4.ScreenshotNuevo("Edtar Registroi", true, file);
                ElementList[3].Clear();
                ElementList[3].SendKeys(crudVariables[crudVariables.Count-1]);
                newFunctions_4.ScreenshotNuevo("Registro Editado", true, file);
            }
        }

        public static void CrudDetalleArbol(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string>crudDetalle, string file)
        {
            List<string> errorMessages = new List<string>();

            var ElementList = desktopSession.FindElementsByClassName("TDBGridInplaceEdit");
            Debug.WriteLine($"la cantidad de elementos es: {ElementList.Count}");
            desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 40, 150);
            desktopSession.Mouse.Click(null);
            int cont = 0;
            //List<string> variables = new List<string>() { "1", "PRUEBA"};
            if (tipo == 1)
            {
                newFunctions_4.ScreenshotNuevo("Agregar Detalle", true, file);
                for (int i = 0; i < crudDetalle.Count; i++)
                {
                    
                    desktopSession.Keyboard.SendKeys(crudDetalle[cont]);
                    cont += 1;
                    if(i< 1)
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    Thread.Sleep(500);
                }
                newFunctions_4.ScreenshotNuevo("Detalle Agregado", true, file);
            }
            else
            {
                newFunctions_4.ScreenshotNuevo("Editar Detalle", true, file);
                ElementList[0].Clear();
                ElementList[0].SendKeys(crudDetalle[crudDetalle.Count-1]);
                newFunctions_4.ScreenshotNuevo("Detalle Editado", true, file);
            }

        }
    }
}
