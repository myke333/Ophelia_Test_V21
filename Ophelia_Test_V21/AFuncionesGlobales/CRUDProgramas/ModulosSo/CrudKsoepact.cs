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
using System.IO;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosSo
{
    class CrudKsoepact : FuncionesVitales
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

        public static void CRUD(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data, string file)
        {
            List<string> errorMessages = new List<string>();
            ClickButtonExterno(desktopSession, 1);
            var ElementList = desktopSession.FindElementsByClassName("TDBEdit");
            if (tipo == 1)
            {
                LupaDinamica2(desktopSession, data);
                //Click en Datos del Evento
                ClickButtonExterno(desktopSession, 2);
                var ElementList1 = desktopSession.FindElementsByClassName("TDBEdit");
                ElementList1[17].SendKeys(data[2]);
                ElementList1[16].SendKeys(data[3]);
                //Click en Datos Iniciales
                ClickButtonExterno(desktopSession, 1);
                desktopSession.Keyboard.SendKeys(data[1]);

            }
            else
            {
                //Click en Datos Iniciales
                ClickButtonExterno(desktopSession, 1);
                for (int i = 0; i < 3; i++)
                {
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowRight);
                }
                for (int i = 0; i < 3; i++)
                {
                    desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Backspace);
                }
                desktopSession.Keyboard.SendKeys(data[4]);
            }
        }

        public static void CrudDetalle(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data, string file)
        {
            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            if (tipo == 1)
            {
                Thread.Sleep(2000);
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 158, 30);
                desktopSession.Mouse.Click(null);
                for (int i = 0; i < 3; i++)
                {
                    desktopSession.Keyboard.SendKeys(data[i]);
                    if (i < 2)
                    {
                        desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                    }
                }
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }
            else
            {
                desktopSession.Keyboard.SendKeys(data[3]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }
        }

        public static void CrudDetalle0(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data, string file)
        {
            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            if (tipo == 1)
            {
                Thread.Sleep(2000);
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 39, 29);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(data[0]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(data[1]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }
            else
            {
                desktopSession.Keyboard.SendKeys(data[2]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }
        }

        public static void CrudDetalle1(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data, string file)
        {
            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            if (tipo == 1)
            {
                Thread.Sleep(2000);
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 33, 31);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(data[0]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(data[1]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }
            else
            {
                desktopSession.Keyboard.SendKeys(data[2]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }
        }

        public static void CrudDetalle2(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data, string file)
        {
            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            if (tipo == 1)
            {
                Thread.Sleep(2000);
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 59, 31);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(data[0]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                desktopSession.Keyboard.SendKeys(data[1]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }
            else
            {
                desktopSession.Keyboard.SendKeys(data[2]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }
        }

        public static void CrudDetalle3(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data, string file)
        {
            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementsByClassName("TDBGrid");
            if (tipo == 1)
            {
                Thread.Sleep(2000);
                desktopSession.Mouse.MouseMove(ElementList[0].Coordinates, 67, 31);
                desktopSession.Mouse.Click(null);
                desktopSession.Keyboard.SendKeys(data[0]);
                desktopSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            }
        }


        public static void CrudDetalle4(WindowsDriver<WindowsElement> desktopSession, int tipo, List<string> data, string file)
        {
            List<string> errorMessages = new List<string>();
            var ElementList = desktopSession.FindElementsByClassName("TDBMemo");
            if (tipo == 1)
            {
                ElementList[0].SendKeys(data[0]);
            }
            else
            {
                ElementList[0].Clear();
                ElementList[0].SendKeys(data[1]);
            } 
        }

        public static void ClickButtonExterno(WindowsDriver<WindowsElement> desktopSession, int tipo)
        {
            //1 -> Datos Iniciales 2 -> Datos del Evento 3 -> Adicionales 4 -> Recomendaciones
            List<string> errorMessages = new List<string>();
            if (tipo == 1)
            {
                desktopSession.FindElementByName("Datos Iniciales").Click();
            }
            else if (tipo == 2)
            {
                desktopSession.FindElementByName("Datos del Evento").Click();
            }
            else if (tipo == 3)
            {
                desktopSession.FindElementByName("Adicionales").Click();
            }
            else if (tipo == 4)
            {
                desktopSession.FindElementByName("Recomendaciones").Click();
            }
        }

        public static void ClickButtonInterno(WindowsDriver<WindowsElement> desktopSession, int tipo)
        {
            //1 -> Datos Iniciales 2 -> Sitio y Lesión 3 -> Factores 4 -> Partes Afectadas 5 -> Actos Inseguros y/o Causas Accidente 
            List<string> errorMessages = new List<string>();
            if (tipo == 1)
            {
                desktopSession.FindElementByName("Parte de Cuerpo Afectada / Agente / Mecanismo del Accidente").Click();
            }
            else if (tipo == 2)
            {
                desktopSession.FindElementByName("Sitio y Lesión").Click();
            }
            else if (tipo == 3)
            {
                desktopSession.FindElementByName("Factores").Click();
            }
            else if (tipo == 4)
            {
                desktopSession.FindElementByName("Partes Afectadas").Click();
            }
            else
            {
                desktopSession.FindElementByName("Actos Inseguros y/o Causas Accidente ").Click();
            }
        }

        public static void LupaDinamica2(WindowsDriver<WindowsElement> desktopSession, List<string> data)
        {
            List<string> errorMessages = new List<string>();
            List<Point> coordenadas = PruebaCRUD.CoordinatesFinder(desktopSession, 82, 150, 75);
            WindowsDriver<WindowsElement> rootSession = null;
            rootSession = RootSession();
            var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
            int contLupa = 1;
            bool IsDisplayedQbe = false;
            int indexVal = 0;
            foreach (Point coord in coordenadas)
            {
                Thread.Sleep(2000);
                //Actions noteClicks = new Actions(desktopSession);
                //noteClicks.MoveToElement(parentElement).MoveByOffset(coord.X + 10, coord.Y).Click().Perform();

                desktopSession.Mouse.MouseMove(parentElement.Coordinates, coord.X + 10, coord.Y);
                desktopSession.Mouse.Click(null);
                if (contLupa == 1)
                {
                    try
                    {
                        rootSession = RootSession();
                        rootSession = PruebaCRUD.ReloadSession(rootSession, "TKQBE");
                    }
                    catch
                    {

                    }
                    if (rootSession != null)
                    {
                        IsDisplayedQbe = true;
                    }
                    else
                    {
                        errorMessages.Add($"No puede encontrar la ventana de QBE");
                    }
                    if (IsDisplayedQbe)
                    {
                        var elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/*"));
                        if (!string.IsNullOrEmpty(data[indexVal]))
                        {
                            elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/Pane[contains(@ClassName, 'TStringGrid')]/*"));
                            elements[0].SendKeys(data[indexVal]);
                        }
                        elements = rootSession.FindElements(By.XPath("//Window[contains(@ClassName, 'TKQBE')]/*"))[0].FindElements(By.XPath("Tab[contains(@ClassName, 'TPageControl')]/Pane[contains(@Name, 'General')]/*"));
                        new Actions(rootSession).MoveToElement(elements[8], 0, 0).MoveByOffset(20, 10).DoubleClick().Perform();
                        Thread.Sleep(500);
                        IsDisplayedQbe = false;

                    }

                    rootSession = RootSession();
                    rootSession = PruebaCRUD.ReloadSession(rootSession, "TCaptura");
                    //var campo = rootSession.FindElementsByXPath("//Window[@ClassName=\"TCaptura\"][@Name=\" LISTA DE EMPLEADOS\"]/Edit[@ClassName=\"TEdit\"]");
                    var campo = rootSession.FindElementsByClassName("TEdit");
                    ////Debugger.Launch();
                    rootSession.Mouse.MouseMove(campo[1].Coordinates, 10, 10);
                    rootSession.Mouse.Click(null);
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
            List<string> botones = new List<string>() { "Insertar", "Borrar", "Editar", "Aplicar", "Cancelar" };
            Point coord = new Point();
            var ElementList = PruebaCRUD.NavClass(desktopSession);
            ////Debugger.Launch();
            int x = ElementList[indiceBarra].Size.Width / 3;
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

        public static List<string> ReportePreliminar(WindowsDriver<WindowsElement> desktopSession, string bandera, string file, string pdf1,int tipo ,string nomPrograma = null)
        {
            List<string> errorMessages = new List<string>();
            List<string> errors = new List<string>();
            if (bandera == "0" || bandera == "1")
            {
                var wait = new DefaultWait<WindowsDriver<WindowsElement>>(desktopSession)
                {
                    Timeout = TimeSpan.FromSeconds(60),
                    PollingInterval = TimeSpan.FromSeconds(1)
                };
                wait.IgnoreExceptionTypes(typeof(InvalidOperationException));
                WindowsDriver<WindowsElement> rootSession = null;
                WindowsElement codPanel = null;
                bool brk = false;
                bool IsDisplayedPreWin = false;
                wait.Until(driver =>
                {
                    List<Point> coordenadas = FuncionesGlobales.coordinatesFinder(desktopSession, 227, 117, 0);
                    int coordX = coordenadas[0].X;
                    int coordY = coordenadas[0].Y;
                    var parentElement = desktopSession.FindElement(By.Name(desktopSession.Title));
                    Thread.Sleep(1000);

                    try
                    {
                        desktopSession.Mouse.MouseMove(parentElement.Coordinates, coordX, coordY);
                        desktopSession.Mouse.Click(null);
                        Thread.Sleep(2000);
                        if (bandera == "0")
                        {
                            
                            if (tipo == 1)
                            {
                                WindowsDriver<WindowsElement> rootSession1 = null;
                                rootSession1 = RootSession();
                                rootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                                rootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                                newFunctions_4.ScreenshotNuevo("Opción Preliminar", true, file);
                                rootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                                Thread.Sleep(1000);
                                errors = GenerarReportes("Preliminar", file, pdf1, nomPrograma);
                                if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                            }
                            else if (tipo == 2)
                            {
                                WindowsDriver<WindowsElement> rootSession1 = null;
                                rootSession1 = RootSession();
                                rootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                                rootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                                rootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                                rootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                                rootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                                rootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                                newFunctions_4.ScreenshotNuevo("Opción Preliminar", true, file);
                                rootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                                Thread.Sleep(1000);
                                GenerarReportes1("Preliminar", file, pdf1, nomPrograma);
                            }
                            else if (tipo == 3 )
                            {
                                WindowsDriver<WindowsElement> rootSession1 = null;
                                rootSession1 = RootSession();
                                rootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                                rootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                                rootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                                rootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                                rootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                                newFunctions_4.ScreenshotNuevo("Opción Preliminar", true, file);
                                rootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                                Thread.Sleep(1000);
                                GenerarReportes1("Preliminar", file, pdf1, nomPrograma);
                            }
                            else
                            {
                                WindowsDriver<WindowsElement> rootSession1 = null;
                                rootSession1 = RootSession();
                                rootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                                rootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                                rootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                                rootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.ArrowDown);
                                rootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Tab);
                                newFunctions_4.ScreenshotNuevo("Opción Preliminar", true, file);
                                rootSession1.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                                Thread.Sleep(1000);
                                errors = GenerarReportes("Preliminar", file, pdf1, nomPrograma);
                                if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
                            }
                            
                        }
                    }
                    catch (Exception ex)
                    {
                        errorMessages.Add($"Error de Appium: {ex.ToString()}");
                    }
                    codPanel = parentElement;
                    return codPanel != null;
                });
            }
            return errorMessages;
        }


        public static List<string> GenerarReportes(string TipoReporte, string file, string ruta, string nomPrograma = null)
        {
            List<string> errorMessages = new List<string>();
            WindowsDriver<WindowsElement> rootSession;
            Thread.Sleep(4000);
            string ReportePdf = "";
            string ReporteW = "";
            string ReporteE = "";
            string Ruta1 = "";
            string Extension1 = "";
            string Ruta2 = "";
            string Extension2 = "";
            string Ruta3 = "";
            string Extension3 = "";
            bool Pdf = false;
            bool Word = false;
            bool Excel = false;

            // Revisar que se abrio el preview
            bool IsDisplayedReview = false;

            try
            {
                for (int i = 0; i < 20; i++)
                {
                    try
                    {
                        rootSession = RootSession();
                        rootSession = ReloadSession(rootSession, "TfrxPreviewForm");
                        var validar = rootSession.FindElementsByClassName("TfrxPreviewForm");
                        break;
                    }
                    catch
                    {
                        Thread.Sleep(1000);
                    }
                }
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TfrxPreviewForm");

                int count = 0;
                while (count < 300)
                {
                    try
                    {
                        List<Point> carga = PruebaCRUD.CoordinatesFinder(rootSession, 224, 32, 64);
                        break;
                    }
                    catch
                    {
                        Thread.Sleep(1000);
                        count++;
                    }
                }

                List<Point> coordenadas = PruebaCRUD.CoordinatesFinder(rootSession, 247, 223, 99);
                int CoordX = coordenadas[0].X;
                int CoordY = coordenadas[0].Y;

                var PanelElements = rootSession.FindElement(By.Name(rootSession.Title));
                bool IsDisplayedReport = false;

                // Clic en el icono de editar reporte
                rootSession.Mouse.MouseMove(PanelElements.Coordinates, CoordX, CoordY);
                rootSession.Mouse.Click(null);
                Thread.Sleep(8000);
                newFunctions_4.ScreenshotNuevo("La opción 'Editar Reporte' " + TipoReporte + " está activada", true, file);

                errorMessages.Add("El Preview del reporte " + TipoReporte + " tiene habilitado la opcion 'Editar Reporte'");

                // Cerrar la ventana de editar Reporte
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TfrxDesignerForm");

                var CerrarEditar = rootSession.FindElementByName("Cerrar");
                CerrarEditar.Click();
                Thread.Sleep(2000);
                newFunctions_4.ScreenshotNuevo("Ventana 'Editar Reporte' " + TipoReporte + " Cerrada", true, file);
            }
            catch
            {

            }

            rootSession = RootSession();
            rootSession = ReloadSession(rootSession, "TfrxPreviewForm");

            var ElementsPreview = rootSession.FindElementsByClassName("TToolBar");

            if (ElementsPreview.Count > 0)
            {
                IsDisplayedReview = true;
            }
            else
            {
                errorMessages.Add("No se puede visualizar la ventana Preview del Reporte " + TipoReporte);
            }

            string fecha = Hora();
            if (IsDisplayedReview == true)
            {
                ElementsPreview = rootSession.FindElementsByClassName("TToolBar");
                int cerrar = 0;
                try
                {
                    Thread.Sleep(2000);
                    List<Point> coordenadas1 = PruebaCRUD.CoordinatesFinder(rootSession, 224, 32, 64);
                    int CoordX = coordenadas1[0].X;
                    int CoordY = coordenadas1[0].Y;

                    var PanelElements = rootSession.FindElement(By.Name(rootSession.Title));

                    // Clic en el icono de exportar PDF
                    rootSession.Mouse.MouseMove(PanelElements.Coordinates, CoordX, CoordY);
                    rootSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    rootSession = RootSession();
                    rootSession = ReloadSession(rootSession, "TfrxPDFExportDialog");

                    var ExportPDF = rootSession.FindElementByName("Aceptar");

                    if (ExportPDF != null)
                    {
                        //Clic en Aceptar exportar PDF
                        ExportPDF.Click();
                        Thread.Sleep(2000);
                        rootSession = RootSession();

                        try
                        {
                            //ingresar la ruta
                            rootSession.Keyboard.SendKeys(ruta);
                            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                            Thread.Sleep(2000);
                            //Abrir PDF
                            Process.Start(ruta + ".pdf");
                            Thread.Sleep(4000);
                            rootSession = RootSession();
                            rootSession = ReloadSession(rootSession, "AcrobatSDIWindow");
                            //var TraerPantalla = rootSession.FindElementsByName("AVL_AVView");
                            newFunctions_4.ScreenshotNuevo("Reporte " + TipoReporte + " exportado a PDF 1", true, file);
                            LimpiarProcesos();
                        }
                        catch
                        {
                            var Ok = rootSession.FindElementByName("OK");
                            Ok.Click();
                            Thread.Sleep(1000);
                            errorMessages.Add("No existe el botón exportar reporte a PDF");
                        }
                    }
                }
                catch
                {
                    errorMessages.Add("En el Preview  del Reporte " + TipoReporte + " NO Existe el Botón Exportar Reporte a PDF");
                    cerrar = 1;
                }

                Thread.Sleep(2000);
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TfrxPreviewForm");
                ElementsPreview = rootSession.FindElementsByClassName("TToolBar");
                Thread.Sleep(1000);

                //Clic en Disquet
                rootSession.Mouse.MouseMove(ElementsPreview[0].Coordinates, 57, 17);
                rootSession.Mouse.Click(null);
                Thread.Sleep(2000);

                //Reportes
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "#32768");
                var MenuItem = rootSession.FindElementsByXPath("/Menu[@Name=\"Contexto\"][@ClassName=\"#32768\"]/*");

                for (int i = 0; i < MenuItem.Count; i++)
                {
                    if (MenuItem[i].GetAttribute("Name").Contains("PDF"))
                    {
                        Pdf = true;
                        ReportePdf = MenuItem[i].GetAttribute("Name");
                        Ruta1 = @"C:\Reportes\Reporte_PDF_" + TipoReporte + "_" + fecha;
                        Extension1 = ".pdf";
                    }

                    if (MenuItem[i].GetAttribute("Name").Contains("Word"))
                    {
                        Word = true;
                        ReporteW = MenuItem[i].GetAttribute("Name");
                        Ruta2 = @"C:\Reportes\Reporte_Word_" + TipoReporte + "_" + fecha;
                        Extension2 = ".docx";
                    }

                    if (MenuItem[i].GetAttribute("Name").Contains("Excel") && MenuItem[i].GetAttribute("Name").Contains("OLE"))
                    {

                    }
                    else if (MenuItem[i].GetAttribute("Name").Contains("Excel"))
                    {
                        Excel = true;
                        ReporteE = MenuItem[i].GetAttribute("Name");
                        Ruta3 = @"C:\Reportes\Reporte_Excel_" + TipoReporte + "_" + fecha;
                        Extension3 = ".xlsx";
                    }
                }

                if (Pdf == true)
                {
                    try
                    {
                        ExportarReporte(ReportePdf, Ruta1, Extension1, TipoReporte, file);
                    }
                    catch
                    {
                        errorMessages.Add("El Proceso de exportar el reporte " + TipoReporte + " a PDF Falló");
                    }
                }

                if (Word == true)
                {
                    try
                    {
                        ExportarReporte(ReporteW, Ruta2, Extension2, TipoReporte, file);
                    }
                    catch
                    {
                        errorMessages.Add("El Proceso de exportar el reporte " + TipoReporte + " a Word Falló");
                    }
                }

                if (Excel == true)
                {
                    try
                    {
                        ExportarReporte(ReporteE, Ruta3, Extension3, TipoReporte, file, nomPrograma);
                    }
                    catch
                    {
                        errorMessages.Add("El Proceso de exportar el reporte " + TipoReporte + " a Excel Falló");
                    }
                }

                Thread.Sleep(2000);
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TfrxPreviewForm");
                ElementsPreview = rootSession.FindElementsByClassName("TToolBar");
                Thread.Sleep(1000);

                if (cerrar == 0)
                {
                    //Cerrar Preview
                    rootSession.Mouse.MouseMove(ElementsPreview[0].Coordinates, 600, 17);
                    rootSession.Mouse.Click(null);
                    Thread.Sleep(2000);
                    try
                    {
                        ElementsPreview = rootSession.FindElementsByClassName("TToolBar");
                        //Cerrar Preview
                        rootSession.Mouse.MouseMove(ElementsPreview[0].Coordinates, 543, 17);
                        rootSession.Mouse.Click(null);
                        Thread.Sleep(2000);
                        try
                        {
                            rootSession.Mouse.MouseMove(rootSession.FindElementByName("Cerrar").Coordinates);
                            rootSession.Mouse.Click(null);
                            errorMessages.Add("El Botón 'Cerrar' no Existe en el Preview del Reporte " + TipoReporte);
                        }
                        catch
                        {

                        }
                    }
                    catch
                    {

                    }
                }
                else
                {
                    //Cerrar Preview
                    rootSession.Mouse.MouseMove(ElementsPreview[0].Coordinates, 520, 17);
                    rootSession.Mouse.Click(null);
                    Thread.Sleep(2000);
                    try
                    {
                        ElementsPreview = rootSession.FindElementsByClassName("TToolBar");
                        //Cerrar Preview
                        rootSession.Mouse.MouseMove(ElementsPreview[0].Coordinates, 543, 17);
                        rootSession.Mouse.Click(null);
                        Thread.Sleep(2000);
                        try
                        {
                            rootSession.Mouse.MouseMove(rootSession.FindElementByName("Cerrar").Coordinates);
                            rootSession.Mouse.Click(null);
                            errorMessages.Add("El Botón 'Cerrar' no Existe en el Preview del Reporte " + TipoReporte);
                        }
                        catch
                        {

                        }
                    }
                    catch
                    {

                    }
                }
            }
            return errorMessages;
        }



        public static void GenerarReportes1(string TipoReporte, string file, string ruta, string nomPrograma = null)
        {
            List<string> errorMessages = new List<string>();
            WindowsDriver<WindowsElement> rootSession;
            rootSession = RootSession();

            newFunctions_4.ScreenshotNuevo("Reporte Preliminar" , true, file);
            var ElementsPreview = rootSession.FindElementsByClassName("TToolBar");
            //Cerrar Preview
            rootSession.Mouse.MouseMove(ElementsPreview[0].Coordinates, 543, 17);
            rootSession.Mouse.Click(null);
            Thread.Sleep(2000);
        }


        private static void ExportarReporte(string NomReporte, string Ruta, string Extension, string TipoReporte, string file, string nomPrograma = null)
        {
            WindowsDriver<WindowsElement> rootSession;

            if (NomReporte.Contains("Word") || NomReporte.Contains("Excel"))
            {
                rootSession = RootSession();
                rootSession = ReloadSession(rootSession, "TfrxPreviewForm");
                var ElementsPreview = rootSession.FindElementsByClassName("TToolBar");

                Thread.Sleep(1000);
                //Clic en Disquet
                rootSession.Mouse.MouseMove(ElementsPreview[0].Coordinates, 57, 17);
                rootSession.Mouse.Click(null);
                Thread.Sleep(1000);
            }

            rootSession = RootSession();
            //clic en Tipo Reporte
            var ClicReport = rootSession.FindElementByName(NomReporte);
            ClicReport.Click();
            Thread.Sleep(1000);
            rootSession = RootSession();
            var AcceptPDF = rootSession.FindElementByName("Aceptar");

            //clic en aceptar exportar Reporte
            AcceptPDF.Click();
            Thread.Sleep(2000);

            //ingresar la ruta
            rootSession.Keyboard.SendKeys(Ruta);
            Thread.Sleep(1000);
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            Thread.Sleep(3000);

            if (NomReporte.Contains("Excel"))
            {
                if (File.Exists(Ruta + Extension))
                {
                    if (nomPrograma == null)
                    {
                        //Abrir reporte
                        Process.Start(Ruta + Extension);
                        Thread.Sleep(4000);
                        rootSession = RootSession();
                        rootSession = ReloadSession(rootSession, "XLMAIN");
                        newFunctions_4.ScreenshotNuevo("Reporte " + TipoReporte + " Exportado a Excel", true, file);
                        LimpiarProcesos();
                    }
                }
                else
                {
                    if (nomPrograma == null)
                    {
                        Extension = ".xls";
                        //Abrir reporte
                        Process.Start(Ruta + Extension);
                        Thread.Sleep(4000);
                        rootSession = RootSession();
                        rootSession = ReloadSession(rootSession, "XLMAIN");
                        newFunctions_4.ScreenshotNuevo("Reporte " + TipoReporte + " Exportado a Excel", true, file);
                        LimpiarProcesos();
                    }
                }
            }
            else
            {
                //Abrir reporte
                Process.Start(Ruta + Extension);
                Thread.Sleep(4000);
                if (Extension == ".docx")
                {
                    rootSession = RootSession();
                    rootSession = ReloadSession(rootSession, "OpusApp");
                    newFunctions_4.ScreenshotNuevo("Reporte " + TipoReporte + " Exportado a Word", true, file);
                }
                else if (Extension == ".pdf")
                {
                    rootSession = RootSession();
                    rootSession = ReloadSession(rootSession, "AcrobatSDIWindow");
                    newFunctions_4.ScreenshotNuevo("Reporte " + TipoReporte + " Exportado a Pdf", true, file);
                }
                LimpiarProcesos();
            }
        }
    }
}

    