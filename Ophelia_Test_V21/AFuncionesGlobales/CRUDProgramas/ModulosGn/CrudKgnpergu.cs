using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Ophelia_Test_V21.AFuncionesGlobales;
using Ophelia_Test_V21.AFuncionesGlobales.newFunctionsAppium;
using OpenQA.Selenium;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Ophelia_Test_V21.AFuncionesGlobales.newFolder;

namespace Ophelia_Test_V21.AFuncionesGlobales.CRUDProgramas.ModulosGn
{
    class CrudKgnpergu : FuncionesVitales
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

        public static List<string> PreliKGnPergu(WindowsDriver<WindowsElement> desktopSession, string file, string pdf1, string BanderaPreli, List<String> data)
        {
            List<string> errorMessages = new List<string>();
            List<string> errors = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            //Thread.Sleep(5000);
            var winElem_RightClickPane = desktopSession.FindElementsByClassName("TGridPanel"); //Encuentra el panel donde esta el preliminar
            Thread.Sleep(1000);
            desktopSession.Mouse.MouseMove(winElem_RightClickPane[0].Coordinates, 130, 20); //El puntero se mueve a las coordenadas indicadas
            desktopSession.Mouse.Click(null); //Da click
            Thread.Sleep(1000);
            rootSession = RootSession();
            rootSession = ReloadSession(rootSession, "TFrmOpc"); //Se para en la ventanita que se abre al dar click
            Thread.Sleep(1000);
            var bot = rootSession.FindElementsByClassName("TGroupButton"); //Encuentra los radio buttons                 
            Thread.Sleep(500);
            //---------------------Preliminar opción que pide usuario---------------------------------------------//
            bot[0].Click(); //Le da click al ultimo radio button (el cual necesita ingresar un usuario por un textedit)
            Thread.Sleep(500);
            var Elementlist = rootSession.FindElementsByClassName("TEdit"); //Encuentra el elemento para ingresar el usuario          
            Elementlist[0].SendKeys(data[0]); //Ingresa el usuario
            newFunctions_4.ScreenshotNuevo("Opción Preliminar 2", true, file);
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter); //Le da aceptar a traves de un "Enter"
            Thread.Sleep(1000);
            errors = GenerarReportes("Preliminar", file, pdf1);
            if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }

            //-----------------------Preliminar para la otra opción------------------------//
            Thread.Sleep(500);
            desktopSession.Mouse.MouseMove(winElem_RightClickPane[0].Coordinates, 130, 20);
            desktopSession.Mouse.Click(null);
            bot[1].Click();
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Opción Preliminar 1", true, file);
            rootSession.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
            errors = newFunctions_5.GenerarReportes("Preliminar", file, pdf1);
            if (errors.Count > 0) { foreach (string er in errors) { errorMessages.Add(er); } }
            Thread.Sleep(1000);
            return errorMessages;
        }
        public static List<string> generarExcel(WindowsDriver<WindowsElement> Session, string file, string codProgram, string ruta)
        {
            List<string> errorMessages = new List<string>();
            WindowsDriver<WindowsElement> rootSession = null;
            Thread.Sleep(5000);
            var winPanel = Session.FindElementsByClassName("TGridPanel"); //Encuentra el panel donde esta el excel
            Thread.Sleep(1000);
            Session.Mouse.MouseMove(winPanel[0].Coordinates, 212, 20); //El puntero se mueve a las coordenadas indicadas
            Session.Mouse.Click(null); //Da click
            Thread.Sleep(1000);
            try
            {
                Session = RootSession();
                Session = ReloadSession(Session, "TFrmExcelex");
                Thread.Sleep(1000);
                var element = Session.FindElementsByXPath("//Window[contains(@ClassName, 'TFrmExcelex')]");
                if (element[0] != null)
                {
                    bool IsDisplayedExcel = false;
                    string Fecha = Hora();
                    Thread.Sleep(500);
                    if (element[0] != null)
                    {
                        IsDisplayedExcel = true;
                    }
                    else
                    {
                        errorMessages.Add($"No se puede visualizar la ventana Exportar a excel");
                    }

                    if (IsDisplayedExcel)
                    {
                        Session = RootSession();
                        Session = ReloadSession(Session, "TFrmExcelex");
                        var ReportElements = Session.FindElementsByXPath("//Window[contains(@ClassName, 'TFrmExcelex')]");
                        Session.Mouse.MouseMove(ReportElements[0].Coordinates, 222, 207);
                        Session.Mouse.Click(null);
                        newFunctions_4.ScreenshotNuevo("Exportar datos Excel", true, file);
                        var buttonAccept = Session.FindElementByName("Aceptar");
                        buttonAccept.Click();
                        Thread.Sleep(5000);
                        try
                        {
                            Thread.Sleep(1000);
                            Session = RootSession();
                            Thread.Sleep(1000);
                            Session = ReloadSession(Session, "XLMAIN");
                            Thread.Sleep(1000);

                            if (Session != null)
                            {
                                int count = 0;
                                while (count < 240)
                                {
                                    try
                                    {
                                        Session = RootSession();
                                        Session = ReloadSession(Session, "XLMAIN");
                                        Thread.Sleep(1000);
                                        var saveExcel1 = Session.FindElementsByName("Maximizar");
                                        if (saveExcel1.Count > 0)
                                        {
                                            break;
                                        }
                                    }
                                    catch
                                    {
                                        Thread.Sleep(1000);
                                        count++;
                                    }
                                }
                                Thread.Sleep(500);
                                var saveExcel = Session.FindElementsByName("Maximizar");
                                Thread.Sleep(500);
                                saveExcel[1].Click();
                                Thread.Sleep(500);
                                Session.FindElementByName("Guardar").Click();
                                Session.FindElementByName("Examinar").Click();
                                Session.Keyboard.SendKeys(ruta);
                                Session.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                                Thread.Sleep(1000);
                                newFunctions_4.ScreenshotNuevo("Reporte Excel Guardado", true, file);
                                LimpiarProcesos();

                            }
                        }
                        catch
                        {
                            Thread.Sleep(1000);
                        }

                        if (Session == null)
                        {
                            Thread.Sleep(1000);
                            newFunctions_4.ScreenshotNuevo("Opción de Excel no trae registros", true, file);
                            Thread.Sleep(1000);
                            Session = RootSession();
                            PruebaCRUD.VentanaAzul(Session);
                            Thread.Sleep(500);
                        }

                    }
                }
            }
            catch
            {
                Session.Keyboard.SendKeys(OpenQA.Selenium.Keys.Enter);
                errorMessages.Add($"Error al intentar generar el excel");
            }
            return errorMessages;
        }

        //-------------------------GUARDAR REPORTES PRELIMINAR--------------------------------------------//
        private static List<Point> CoordinatesFinder(WindowsDriver<WindowsElement> session, int R, int G, int B)
        {
            Screenshot image = ((ITakesScreenshot)session).GetScreenshot();
            string name = FuncionesVitales.Hora();
            string path = "C:\\imagenesTest\\" + name + ".Png";
            Image imgSource;
            try
            {
                image.SaveAsFile(string.Format("C:\\imagenesTest\\{0}.Png", name), ScreenshotImageFormat.Png);
                Thread.Sleep(1000);
                imgSource = Image.FromFile(path);
            }
            catch (Exception e)
            {
                image.SaveAsFile(string.Format("C:\\imagenesTest\\{0}.Png", name), ScreenshotImageFormat.Png);
                Thread.Sleep(1000);
                imgSource = Image.FromFile(path);
            }

            Point puntos = new Point();
            int lastx = 0;
            int lasty = 0;
            List<Point> coordenadas = new List<Point>();
            Bitmap bmpSource = new Bitmap(imgSource);
            Bitmap bmpDest = new Bitmap(bmpSource.Width, bmpSource.Height);
            int Romper = 0;

            for (int x = 0; x < bmpSource.Width; x++)
            {
                for (int y = 0; y < bmpSource.Height; y++)
                {
                    Color clrPixel = bmpSource.GetPixel(x, y);
                    ////Debugger.Launch();
                    if (clrPixel.R == R && clrPixel.G == G && clrPixel.B == B)
                    {
                        puntos.X = x;
                        puntos.Y = y;
                        if (x > lastx + 20 || y > lasty + 20)
                        {
                            coordenadas.Add(puntos);
                            lastx = x;
                            lasty = y;
                            Romper++;
                            break;
                        }
                    }
                }
                if (Romper != 0)
                {
                    break;
                }
            }
            ////Debugger.Launch();
            return coordenadas;
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
        public static List<string> GenerarReportes(string TipoReporte, string file, string ruta, string nomPrograma = null)
        {
            List<string> errorMessages = new List<string>();
            WindowsDriver<WindowsElement> rootSession;
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

                //var sleep = 0;
                //if (x=="0")
                //{
                //    sleep = 100;
                //} else { sleep = 1140000; }
                //Thread.Sleep(sleep);

                int count = 0;
                while (count < 300)
                {
                    try
                    {
                        List<Point> carga = CoordinatesFinder(rootSession, 224, 32, 64);
                        break;
                    }
                    catch
                    {
                        Thread.Sleep(1000);
                        count++;
                    }
                }

                List<Point> coordenadas = CoordinatesFinder(rootSession, 247, 223, 99);
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
            Thread.Sleep(500);
            newFunctions_4.ScreenshotNuevo("Ventana Preview Reporte", true, file);

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
                    List<Point> coordenadas1 = CoordinatesFinder(rootSession, 255, 0, 0);
                    int CoordX = coordenadas1[0].X;
                    int CoordY = coordenadas1[0].Y;

                    var PanelElements = rootSession.FindElement(By.Name(rootSession.Title));

                    // Clic en el icono de exportar PDF
                    rootSession.Mouse.MouseMove(PanelElements.Coordinates, CoordX, CoordY);
                    rootSession.Mouse.Click(null);
                    Thread.Sleep(1000);

                    rootSession = RootSession();
                    rootSession = ReloadSession(rootSession, "TfrxPDFExportDialog");
                    //Thread.Sleep(2000);
                    var bot = rootSession.FindElementsByClassName("TRadioButton"); //Encuentra los radio buttons
                    Thread.Sleep(500);
                    bot[1].Click(); //Le da click a Pagina Actual
                    Thread.Sleep(500);

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
                            Thread.Sleep(1000);
                            //Abrir PDF
                            Process.Start(ruta + ".pdf");
                            //Thread.Sleep(4000);
                            rootSession = RootSession();
                            rootSession = ReloadSession(rootSession, "AcrobatSDIWindow");
                            //var TraerPantalla = rootSession.FindElementsByName("AVL_AVView");
                            //Thread.Sleep(4000);
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
                //Thread.Sleep(2000);

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
        //-------------------------GUARDAR REPORTES PRELIMINAR--------------------------------------------//
    }
}
