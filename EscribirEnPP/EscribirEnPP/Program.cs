using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EscribirEnPP
{
    public class Program
    {
        private static string fileName = @"\EjemploCreado.pptx";

        public static void Main(string[] args)
        {
            if (!ExisteArchivo())
                CreatePresentation();

            if (WriteOnSlide())
                Console.WriteLine("La aplicacion escribio correctamente en la diapositiva");
            Console.ReadKey();
        }


        private static bool WriteOnSlide()
        {
            var resultado = false;
            string filePath = CurrentDirectory() + fileName;
            try
            {
                Application pptApplication = new Application();
                Presentations multi_presentations = pptApplication.Presentations;
                Presentation presentation = multi_presentations.Open(filePath, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);
                CustomLayout customLayout = presentation.SlideMaster.CustomLayouts[PpSlideLayout.ppLayoutText];

                Slides slides = presentation.Slides;
                Microsoft.Office.Interop.PowerPoint.Shapes shapes = presentation.Slides[1].Shapes;
                TextRange objText;

                slides = presentation.Slides;

                var text1 = "Esto es una prueba de escribir en el titulo";
                var text2 = "Estoy escribiendo en la seccion de contenido de la diapositiva PPT";

                objText = shapes[1].TextFrame.TextRange;
                objText.Text = text1;
                objText.Font.Name = "Arial";
                objText.Font.Size = 32;

                objText = shapes[2].TextFrame.TextRange;
                objText.Text = text2;
                objText.Font.Name = "Arial";
                objText.Font.Size = 28;

                ReadWriteTxt(filePath);
                presentation.SaveAs(filePath, PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
                presentation.Close();
                pptApplication.Quit();

                resultado = true;
            }
            catch (Exception ex)
            {
                WriteException(ex.ToString());
            }
            return resultado;
        }

        private static bool CreatePresentation()
        {
            var resultado = false;
            try
            {
                Application pptApplication = new Application();

                Slides slides;
                _Slide slide;
                TextRange objText;

                Presentation pptPresentation = pptApplication.Presentations.Add(MsoTriState.msoFalse);
                CustomLayout customLayout = pptPresentation.SlideMaster.CustomLayouts[PpSlideLayout.ppLayoutText];

                slides = pptPresentation.Slides;
                slide = slides.AddSlide(1, customLayout);

                objText = slide.Shapes[1].TextFrame.TextRange;
                objText.Text = "Escriba un titulo";
                objText.Font.Name = "Arial";
                objText.Font.Size = 32;

                objText = slide.Shapes[2].TextFrame.TextRange;
                objText.Text = "Escriba un contenido";
                objText.Font.Size = 28;

                slide.NotesPage.Shapes[2].TextFrame.TextRange.Text = "Presentacion creada desde C#, Efrain Mejias C";

                pptPresentation.SaveAs(CurrentDirectory() + fileName, PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
                pptPresentation.Close();
                pptApplication.Quit();
                resultado = true;
            }
            catch(Exception ex)
            {
                WriteException(ex.ToString());
            }
      
            return resultado;
        }

        public static void ReadSlide()
        {
            try
            {
                string filePath = CurrentDirectory() + fileName;

                Application pptApplication = new Application();
                Presentations multi_presentations = pptApplication.Presentations;
                Presentation presentation = multi_presentations.Open(filePath, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);

                string presentationText = string.Empty;
                foreach (var item in presentation.Slides[1].Shapes)
                {
                    var shape = (Microsoft.Office.Interop.PowerPoint.Shape)item;
                    if (shape.HasTextFrame == MsoTriState.msoTrue)
                    {
                        if (shape.TextFrame.HasText == MsoTriState.msoTrue)
                        {
                            var textRange = shape.TextFrame.TextRange;
                            var text = textRange.Text;

                            presentationText += text + " ";
                        }
                    }
                }

                Console.WriteLine(presentationText);
            }
            catch (Exception ex)
            {
                WriteException(ex.ToString());
            }

        }

        private static string CurrentDirectory()
        {
            return System.IO.Directory.GetCurrentDirectory();
        }

        private static void ReadWriteTxt(string pathArchivo)
        {
            FileAttributes atr = File.GetAttributes(pathArchivo);
            File.SetAttributes(pathArchivo,atr & ~FileAttributes.ReadOnly);
        }

        private static bool ExisteArchivo()
        {
            bool result = false;
            if (File.Exists(CurrentDirectory() + fileName))
                result = true;

            return result;
        }

        private static void WriteException(string exception)
        {
            Console.WriteLine(exception);
        }
    }
}
