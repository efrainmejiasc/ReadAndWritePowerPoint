using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EscribirEnPP
{
    public class Program
    {
        public static void Main(string[] args)
        {

            ReadSlide();
            Console.ReadKey();
           
        }

        private static bool WriteOnSlide()
        {
            var resultado = false;
            string filePath = CurrentDirectory() + @"\EjemploCreado.pptx";
            try
            {
                Application pptApplication = new Application();
                Presentations multi_presentations = pptApplication.Presentations;
                Presentation presentation = multi_presentations.Open(filePath, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);
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
                objText.Text = "Escribiendo en PPT desde cero";
                objText.Font.Name = "Arial";
                objText.Font.Size = 32;

                objText = slide.Shapes[2].TextFrame.TextRange;
                objText.Text = "Otra Linea en PPT";

                slide.NotesPage.Shapes[2].TextFrame.TextRange.Text = "Presentacion Creada";

                pptPresentation.SaveAs(CurrentDirectory() + @"\EjemploCreado.pptx", PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
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
                string filePath = CurrentDirectory() + @"\EjemploCreado.pptx";

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

                            presentationText+= text + " ";
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

        private static void WriteException(string exception)
        {
            Console.WriteLine(exception);
        }
    }
}
