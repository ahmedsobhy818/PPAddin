using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;

namespace PPAddin
{
   public partial class ThisAddIn
   {
        public static PowerPoint.Application application;
      
      protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
      {
            
         return new Ribbon2();
           

        }
      private void ThisAddIn_Startup(object sender, System.EventArgs e)
      {
          //  MessageBox.Show("hi");
            application = this.Application;
            this.Application.PresentationNewSlide +=
   new PowerPoint.EApplication_PresentationNewSlideEventHandler(
   Application_PresentationNewSlide);
            //PowerPoint.p
            //foreach (PresentationSlide o in Application.ActivePresentation.Slides.GetEnumerator())
            
        }
        void Application_PresentationNewSlide(PowerPoint.Slide Sld)
        {
            PowerPoint.Shape textBox = Sld.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 50, 50);
            textBox.TextFrame.TextRange.InsertAfter("This text was added by using code.");
        }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
      {
           // MessageBox.Show("bye");
        }

      

      #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
      {
         this.Startup += new System.EventHandler(ThisAddIn_Startup);
         this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
      }

      #endregion
   }
}
