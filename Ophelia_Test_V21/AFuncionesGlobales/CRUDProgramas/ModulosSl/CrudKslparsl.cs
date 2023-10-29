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

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosSl
{
    class CrudKslparsl : FuncionesVitales
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

        public static void CRUDSQL(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data)
        {
            List<string> errorMessages = new List<string>();
            if (tipo == 1 && data[1] == "N")
            {
                desktopSession.FindElementByName("Filtrar Vacantes X Centro de Costo").Click();
            }
            else if (tipo == 2 && data[2] == "S")
            {
                desktopSession.FindElementByName("Filtrar Vacantes X Centro de Costo").Click();
            }
        }

        public static void ClickButton(WindowsDriver<WindowsElement> desktopSession, int tipo)
        {
            List<string> errorMessages = new List<string>();
            if (tipo == 1)
            {
                desktopSession.FindElementByName("Parámetros Convocatorias").Click();
            }
            else if (tipo == 2)
            {
                desktopSession.FindElementByName("Parámetros Requisiciones").Click();
            }
        }

        public static void ClickButtonInterno(WindowsDriver<WindowsElement> desktopSession, int tipo)
        {
            List<string> errorMessages = new List<string>();
            if (tipo == 1)
            {
                desktopSession.FindElementByName("Generales").Click();
            }
            else if (tipo == 2)
            {
                desktopSession.FindElementByName("Parámetros de Requisición Web").Click();
            }
            else
            {
                desktopSession.FindElementByName("Parámetros Reporte Validación de Perfil").Click();
            }
        }

        public static void CRUDORA(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data, List<string> data1, string file)
        {
            List<string> errorMessages = new List<string>();
            List<string> variablesButtonGenerales = new List<string> { "Automatizar Consecutivo Requisición de Personal", "Validación de Perfiles de Planta", "Permitir Aspirantes en Varios Procesos Activos", "Activar revisión hoja de vida ", "Usa Etapas de Selección", "Usa Kit de Ingreso" };
            List<string> variablesButtonReqWeb = new List<string> { "Activar Forma de cobertura de la vacante", "Activar  Sueldo", "Activar  Tipo de contrato" };
            List<string> variablesButtonParamConvocatorias = new List<string> { "Maneja Anexos", "Valida Documentos Adicionales", "Sedes Vacantes", "Convocatoria por Proceso", "Vigencia Lista de Elegidos", "Cerrar Prueba", "Causales Tipo RUIC", "Activar link convocatoria en KactusRL", "Filtrar Vacantes X Centro de Costo" };

            if (tipo == 1)
            {
                LupaDinamica(desktopSession, data);
                //CRUD GENERALES
                int cont = 1;
                foreach (var list1 in variablesButtonGenerales)
                {
                    if (data[cont] == "S")
                    {
                        desktopSession.FindElementByName(list1).Click();
                    }
                    cont = cont + 1;
                }
                newFunctions_4.ScreenshotNuevo("Crud Generales", true, file);

                //CRUD PARAMETROS DE REQUISICION WEB
                // 1 ->Generales 2 -> Parámetros de Requisición Web 3 -> Parámetros Reporte Validación de Perfil
                ClickButtonInterno(desktopSession, 2);

                int cont1 = 7;
                foreach (var list2 in variablesButtonReqWeb)
                {
                    if (data[cont1] == "S" && cont1 <= 9)
                    {
                        desktopSession.FindElementByName(list2).Click();
                    }
                    cont1 = cont1 + 1;
                }

                desktopSession.FindElementByClassName("TDBLookupComboBox").Click();
                for (int i = 0; i < 3; i++)
                {
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                }
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                desktopSession.FindElementByClassName("TDBMemo").SendKeys(data[10]);

                newFunctions_4.ScreenshotNuevo("Crud Parametros de Requisición Web", true, file);

                //CRUD PARAMETROS CONVOCATORIAS
                // 2 -> CLick en Parámetros Requisiciones 1 -> Parámetros Convocatorias
                ClickButton(desktopSession, 1);

                if (data[11] == "S")
                {
                    desktopSession.FindElementByName("Validar Tipo Cargo").Click();
                }
                LupaDinamica(desktopSession, data1);

                int cont2 = 12;
                foreach (var list3 in variablesButtonParamConvocatorias)
                {
                    if (data[cont2] == "S" && cont2 <= 20)
                    {
                        desktopSession.FindElementByName(list3).Click();
                    }
                    cont2 = cont2 + 1;
                }
                newFunctions_4.ScreenshotNuevo("Crud Parametros Convocatorias", true, file);

            }
            else if (tipo == 2)
            {
                if (data[21] == "S")
                {
                    desktopSession.FindElementByName("Filtrar Vacantes X Centro de Costo").Click();
                }

            }
        }


        public static void CRUDDetalle1ORA(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data)
        {

            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            if (tipo == 1)
            {
                Thread.Sleep(2000);
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 53, 26);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(data[0]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(data[1]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(data[2]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }
            else
            {
                desktopSession.Keyboard.SendKeys(data[3]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }
        }

        public static void CRUDDetalle2ORA(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data)
        {

            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            if (tipo == 1)
            {
                Thread.Sleep(2000);
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 27, 29);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(data[0]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(data[1]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(data[2]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }
            else
            {
                desktopSession.Keyboard.SendKeys(data[3]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }
        }


        public static void LupaDinamica(WindowsDriver<WindowsElement> desktopSession, List<string> data)
        {
            List<string> errorMessages = new List<string>();
            List<Point> coordenadas = PruebaCRUD.CoordinatesFinder(desktopSession, 82, 150, 75);
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
            int contLupa = 1;
            int indexVal = 0;
            foreach (Point coord in coordenadas)
            {
                desktopSession.Mouse.MouseMove(parentElement.Coordinates, coord.X + 10, coord.Y);
                desktopSession.Mouse.Click(null);
                Thread.Sleep(2000);
                if (contLupa == 1)
                {
                    rootSession = RootSession();
                    rootSession = ReloadSession(rootSession, "TCaptura");
                    Thread.Sleep(1000);
                    var campo = rootSession.FindElementsByClassName("TEdit");
                    ////Debugger.Launch();
                    rootSession.Mouse.MouseMove(campo[1].Coordinates, 10, 10);
                    rootSession.Mouse.Click(null);
                    campo[1].Clear();
                    rootSession.Keyboard.SendKeys(data[indexVal]);
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
                    indexVal++;
                }
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
