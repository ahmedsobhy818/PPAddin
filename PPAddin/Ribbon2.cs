using Microsoft.Office.Core;
using PPAddin.Forms;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using System.Threading;
using Application = System.Windows.Forms.Application;
using TextFrame2 = Microsoft.Office.Interop.PowerPoint.TextFrame2;
using System.Drawing;
using PPAddin.Properties;
// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon2();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace PPAddin
{
    [ComVisible(true)]
    public class Ribbon2 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon2()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("PPAddin.Ribbon2.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }
        //        public  void AnonymizeWithLoremIpsum(IRibbonControl control )
        //        {
        //            // MessageBox.Show("aaaaa");
        //            /*
        //           Sub AnonymizeShapeWithLoremIpsum(SlideShape)

        //    If SlideShape.Type = msoGroup Then

        //        Set SlideShapeGroup = SlideShape.GroupItems

        //        For Each Shape SlideShapeChild In SlideShapeGroup


        //            AnonymizeShapeWithLoremIpsum SlideShapeChild
        //        Next

        //    Else

        //        If SlideShape.HasTextFrame Then

        //            For Each Paragraph In SlideShape.TextFrame2.TextRange.Paragraphs
        //                If Paragraph.Length > 1 Then
        //                Paragraph.text = GetLoremIpsum(Paragraph.words.Count, Paragraph.Length)
        //                End If
        //            Next

        //        End If

        //        If SlideShape.HasTable Then
        //            For TableRow = 1 To SlideShape.Table.Rows.Count
        //                For TableColumn = 1 To SlideShape.Table.Columns.Count

        //                    For Each Paragraph In SlideShape.Table.Cell(TableRow, TableColumn).Shape.TextFrame2.TextRange.Paragraphs
        //                        If Paragraph.Length > 1 Then
        //                        Paragraph.text = GetLoremIpsum(Paragraph.words.Count, Paragraph.Length)
        //                        End If
        //                    Next

        //                Next
        //            Next
        //        End If

        //        If SlideShape.HasSmartArt Then

        //            For SlideShapeSmartArtNode = 1 To SlideShape.SmartArt.AllNodes.Count

        //                For Each Paragraph In SlideShape.SmartArt.AllNodes(SlideShapeSmartArtNode).TextFrame2.TextRange.Paragraphs
        //                    If Paragraph.Length > 1 Then
        //                    Paragraph.text = GetLoremIpsum(Paragraph.words.Count, Paragraph.Length)
        //                    End If
        //                Next

        //            Next

        //        End If

        //    End If

        //End Sub
        //             */
        //            //MessageBox.Show(  ThisAddIn.application.ActivePresentation.Slides.Count.ToString());

        //            //foreach (Slide x in ThisAddIn.application.ActivePresentation.Slides)
        //            //{

        //            //}
        //            ProgressForm f = new ProgressForm();
        //            f.Show();
        //            foreach (Slide PresentationSlide in ThisAddIn.application.ActivePresentation.Slides)
        //            {
        //                //                SetProgress(PresentationSlide.SlideNumber / ActivePresentation.Slides.Count * 100)
        //                f.SetProgress(PresentationSlide.SlideNumber * 100/ ThisAddIn.application.ActivePresentation.Slides.Count );
        //                //System.Windows.Forms.Application.DoEvents();

        //                foreach(Shape  SlideShape in PresentationSlide.Shapes)
        //                {
        //                    AnonymizeShapeWithLoremIpsum (SlideShape);
        //                }
        //                Application.DoEvents();
        //            }
        //            f.Hide();
        //        }
        //       void AnonymizeShapeWithLoremIpsum(Shape SlideShape)
        //        {
        //            if (SlideShape.Type == MsoShapeType.msoGroup)
        //            {
        //                var SlideShapeGroup = SlideShape.GroupItems;
        //                foreach (Shape SlideShapeChild in SlideShapeGroup)
        //                {
        //                    AnonymizeShapeWithLoremIpsum(SlideShapeChild);
        //                }
        //            }
        //            else
        //            {
        //                if (SlideShape.HasTextFrame == MsoTriState.msoCTrue || SlideShape.HasTextFrame == MsoTriState.msoTrue) {
        //                    foreach (TextRange2 Paragraph in SlideShape.TextFrame2.TextRange.Paragraphs)
        //                    {
        //                        if (Paragraph.Length > 1) {
        //                            Paragraph.Text = GetLoremIpsum(Paragraph.Words.Count, Paragraph.Length);
        //                        }
        //                    }
        //                }

        //                ////
        //                if (SlideShape.HasTable == MsoTriState.msoCTrue|| SlideShape.HasTable == MsoTriState.msoTrue) {
        //                    for (int TableRow = 1; TableRow <= SlideShape.Table.Rows.Count; TableRow++) {
        //                        for (int TableColumn = 1; TableColumn <= SlideShape.Table.Columns.Count; TableColumn++) {

        //                            foreach (TextRange2 Paragraph in SlideShape.Table.Cell(TableRow, TableColumn).Shape.TextFrame2.TextRange.Paragraphs) {
        //                                if (Paragraph.Length > 1) {
        //                                    Paragraph.Text = GetLoremIpsum(Paragraph.Words.Count, Paragraph.Length);
        //                         }
        //                       }

        //                       }
        //                   }
        //        }

        //                ///
        //                if (SlideShape.HasSmartArt==MsoTriState.msoCTrue || SlideShape.HasSmartArt == MsoTriState.msoTrue) {


        //                    for (int SlideShapeSmartArtNode = 1; SlideShapeSmartArtNode <= SlideShape.SmartArt.AllNodes.Count; SlideShapeSmartArtNode++) {

        //                        foreach (TextRange2 Paragraph in SlideShape.SmartArt.AllNodes[SlideShapeSmartArtNode].TextFrame2.TextRange.Paragraphs) {
        //                            if (Paragraph.Length > 1) {
        //                                Paragraph.Text = GetLoremIpsum(Paragraph.Words.Count, Paragraph.Length);
        //                    }
        //                   }

        //                  }


        //        }


        //            }






        //        }
        //        public string GetLoremIpsum(int NumberOfWords ,int MaxLength ) {
        //            if (NumberOfWords <= 0)
        //                return "";

        //            string[] LoremIpsumWords;
        //            string LoremResult="";
        //            int WordCount;

        //            string LoremIpsum = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Phasellus ac finibus purus. Phasellus et ultricies erat. Nullam maximus risus est, a pulvinar lectus pulvinar ut. Integer dictum malesuada sapien ac vulputate. Nam leo mauris, tincidunt quis dictum vel, semper nec est. Sed et dignissim tortor. Phasellus bibendum elit posuere erat malesuada ornare a sed odio. Integer purus lectus, gravida ac porttitor in, volutpat dictum sem. Pellentesque fermentum ante euismod dolor pellentesque, vitae vestibulum odio sagittis. In et massa massa.";
        //            LoremIpsum = LoremIpsum + "Mauris maximus sem eget semper sollicitudin. Nullam gravida eros non scelerisque cursus. Sed non sem iaculis diam lacinia fermentum id vitae neque. Nulla facilisi. Vestibulum interdum ex non lorem tristique condimentum. Vestibulum facilisis tincidunt nulla at commodo. Ut pretium rhoncus lacus eget porttitor. Etiam quis euismod risus. Maecenas vel porta ante. Curabitur at rutrum eros, et vehicula ligula. Duis in maximus ante. Duis sed est in diam finibus venenatis.";
        //            LoremIpsum = LoremIpsum + "Morbi sollicitudin felis sed scelerisque congue. Nullam vitae urna facilisis, consectetur urna non, ultricies mauris. Vivamus leo tortor, cursus vitae lacinia eget, varius et libero. Fusce luctus nec lectus sed dignissim. Donec malesuada ipsum in sagittis dictum. Nam vel augue id nulla porttitor consectetur. Duis nec enim id enim sagittis aliquam. Curabitur at nulla mi.";
        //    LoremIpsum = LoremIpsum + "Praesent ac turpis eu elit auctor rhoncus. Mauris quis vehicula purus. Morbi sed neque leo. Sed ornare, ipsum et vulputate mattis, augue nisl feugiat magna, nec consequat elit risus eu est. Fusce viverra, urna vel porttitor vehicula, nulla nunc efficitur nunc, quis dapibus nulla ex quis ante. Nunc auctor iaculis sodales. Nunc vitae diam scelerisque, pretium ante vel, tincidunt velit. Sed nec congue arcu. Vestibulum vestibulum dolor sed nulla consequat vulputate. Donec nec dolor sed massa facilisis hendrerit. Curabitur dignissim vestibulum orci, sed facilisis neque condimentum id. Pellentesque erat nibh, euismod at dui quis, rutrum consectetur dolor.";
        //            LoremIpsum = LoremIpsum + "Duis non ex nec lorem venenatis pellentesque. Ut euismod luctus tortor, sed consequat ipsum luctus sed. Duis at velit consectetur, commodo justo id, viverra tellus. Phasellus eu turpis non nisl porta suscipit et at ipsum. Mauris sodales purus vitae dolor hendrerit feugiat. Sed sit amet semper urna, a egestas ex. Phasellus mollis sodales augue at fermentum. Quisque aliquam scelerisque congue. In vitae hendrerit orci. Quisque ut luctus nisi. Donec sit amet mollis neque. Suspendisse vulputate tempus elit. Mauris quis turpis pellentesque, bibendum lectus eu, aliquam leo. Duis congue magna ac erat iaculis, eu bibendum orci finibus.";
        //            LoremIpsum = LoremIpsum + "Ut volutpat maximus orci, vel ultrices turpis consequat in. Cras eu euismod odio, quis dapibus neque. Mauris ut dui id lacus tincidunt dapibus a eget lacus. Aenean imperdiet fringilla justo, in pellentesque sapien placerat a. Donec nisi augue, tempor eu blandit sed, efficitur et mi. Donec efficitur lectus non eros placerat, at egestas diam iaculis. Integer sodales turpis congue sagittis tempor. Donec nec orci sit amet augue sagittis gravida id vitae massa. Donec nec tincidunt velit. Integer nisl dolor, mollis ut ultrices quis, fermentum sed nisi. Ut aliquam nisi at orci ullamcorper, at malesuada orci sodales. Nunc ut molestie mauris. Donec rutrum aliquet velit, nec maximus urna tincidunt sed.";
        //            LoremIpsum = LoremIpsum + "Donec rhoncus massa leo, sit amet tempus dui rutrum ac. Suspendisse at rutrum libero. Proin pharetra maximus mollis. Morbi molestie quis tortor sed consectetur. Aenean ullamcorper iaculis pharetra. Maecenas et blandit nisl, quis scelerisque nisl. Donec vel tempor sem, ac consequat justo. Pellentesque quis libero euismod, feugiat lacus et, finibus eros. Aenean finibus sit amet massa consectetur semper. Ut hendrerit euismod ipsum. Pellentesque lorem leo, vulputate non orci ut, convallis semper ex. Nunc fermentum tempor sagittis. Aliquam erat volutpat. Vivamus fringilla finibus ex sed pharetra. Quisque pharetra dictum lectus, sit amet dapibus eros accumsan eu. Pellentesque at lectus eu ipsum congue mollis.";
        //            LoremIpsum = LoremIpsum + "Nunc ac condimentum justo. Phasellus vel massa aliquet, pulvinar ligula in, ornare enim. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Nulla molestie nisi nec posuere tincidunt. Cras eget bibendum ante, id facilisis augue. Class aptent taciti sociosqu ad litora torquent per conubia nostra, per inceptos himenaeos. Donec id turpis maximus, semper orci ac, tristique arcu. Sed euismod sapien sed nisl scelerisque suscipit. Pellentesque mollis volutpat orci quis eleifend. Curabitur et nisi est. Integer finibus commodo pretium.";
        //            LoremIpsum = LoremIpsum + "Nunc dignissim tincidunt blandit. Sed quis arcu a lacus cursus mollis vitae nec eros. Ut dignissim cursus massa, nec elementum leo pellentesque ut. Aenean nec nunc scelerisque dui maximus consequat. Morbi diam augue, ullamcorper eget dictum id, venenatis vitae ipsum. Nulla facilisi. Aliquam mollis leo sed leo tempus aliquam. Donec a erat at justo rhoncus commodo ut eu erat. Ut vitae nisl rutrum, consectetur leo quis, laoreet diam. Sed metus leo, semper sit amet volutpat ut, placerat eu diam. Donec malesuada nunc ac pretium hendrerit.";
        //            LoremIpsum = LoremIpsum + "Integer viverra pulvinar augue. Nulla et erat sed ante suscipit vulputate. Proin a iaculis nisl. Pellentesque convallis lorem sit amet euismod tincidunt. Pellentesque nisl mauris, dignissim sed imperdiet vel, tristique a orci. Integer ut scelerisque quam. Sed scelerisque lectus ut convallis malesuada. Morbi vehicula hendrerit magna in placerat.";
        //            LoremIpsum = LoremIpsum + "Integer non interdum sapien. Praesent dictum risus erat, non iaculis dolor bibendum accumsan. Fusce fermentum ultricies ultrices. Ut condimentum elit vitae scelerisque euismod. Suspendisse massa ante, interdum in nisl quis, blandit.";
        //            LoremIpsum = LoremIpsum + LoremIpsum + LoremIpsum + LoremIpsum + LoremIpsum;
        //            LoremIpsumWords = LoremIpsum.Split(new string[] { " " }, StringSplitOptions.None);
        //            if (NumberOfWords > LoremIpsumWords.Length - 1)
        //                return LoremIpsum;

        //            LoremResult = LoremIpsumWords[0];
        //            WordCount = 1;

        //            while (WordCount < NumberOfWords) {


        //                if ((LoremResult + " " + LoremIpsumWords[WordCount]).Length <= MaxLength || NumberOfWords <= 2)
        //                    LoremResult = LoremResult + " " + LoremIpsumWords[WordCount];



        //        WordCount = WordCount + 1;
        //            }

        //    return LoremResult;
        //        }
        /*
         Public Function GetLoremIpsum(NumberOfWords As Long, MaxLength As Long) As String
    
    If (NumberOfWords <= 0) Then
        GetLoremIpsum = ""
        Exit Function
    End If
    
    Dim LoremIpsumWords() As String
    Dim LoremResult As String
    Dim WordCount   As Long
    
    LoremIpsum = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Phasellus ac finibus purus. Phasellus et ultricies erat. Nullam maximus risus est, a pulvinar lectus pulvinar ut. Integer dictum malesuada sapien ac vulputate. Nam leo mauris, tincidunt quis dictum vel, semper nec est. Sed et dignissim tortor. Phasellus bibendum elit posuere erat malesuada ornare a sed odio. Integer purus lectus, gravida ac porttitor in, volutpat dictum sem. Pellentesque fermentum ante euismod dolor pellentesque, vitae vestibulum odio sagittis. In et massa massa."
    LoremIpsum = LoremIpsum + "Mauris maximus sem eget semper sollicitudin. Nullam gravida eros non scelerisque cursus. Sed non sem iaculis diam lacinia fermentum id vitae neque. Nulla facilisi. Vestibulum interdum ex non lorem tristique condimentum. Vestibulum facilisis tincidunt nulla at commodo. Ut pretium rhoncus lacus eget porttitor. Etiam quis euismod risus. Maecenas vel porta ante. Curabitur at rutrum eros, et vehicula ligula. Duis in maximus ante. Duis sed est in diam finibus venenatis."
    LoremIpsum = LoremIpsum + "Morbi sollicitudin felis sed scelerisque congue. Nullam vitae urna facilisis, consectetur urna non, ultricies mauris. Vivamus leo tortor, cursus vitae lacinia eget, varius et libero. Fusce luctus nec lectus sed dignissim. Donec malesuada ipsum in sagittis dictum. Nam vel augue id nulla porttitor consectetur. Duis nec enim id enim sagittis aliquam. Curabitur at nulla mi."
    LoremIpsum = LoremIpsum + "Praesent ac turpis eu elit auctor rhoncus. Mauris quis vehicula purus. Morbi sed neque leo. Sed ornare, ipsum et vulputate mattis, augue nisl feugiat magna, nec consequat elit risus eu est. Fusce viverra, urna vel porttitor vehicula, nulla nunc efficitur nunc, quis dapibus nulla ex quis ante. Nunc auctor iaculis sodales. Nunc vitae diam scelerisque, pretium ante vel, tincidunt velit. Sed nec congue arcu. Vestibulum vestibulum dolor sed nulla consequat vulputate. Donec nec dolor sed massa facilisis hendrerit. Curabitur dignissim vestibulum orci, sed facilisis neque condimentum id. Pellentesque erat nibh, euismod at dui quis, rutrum consectetur dolor."
    LoremIpsum = LoremIpsum + "Duis non ex nec lorem venenatis pellentesque. Ut euismod luctus tortor, sed consequat ipsum luctus sed. Duis at velit consectetur, commodo justo id, viverra tellus. Phasellus eu turpis non nisl porta suscipit et at ipsum. Mauris sodales purus vitae dolor hendrerit feugiat. Sed sit amet semper urna, a egestas ex. Phasellus mollis sodales augue at fermentum. Quisque aliquam scelerisque congue. In vitae hendrerit orci. Quisque ut luctus nisi. Donec sit amet mollis neque. Suspendisse vulputate tempus elit. Mauris quis turpis pellentesque, bibendum lectus eu, aliquam leo. Duis congue magna ac erat iaculis, eu bibendum orci finibus."
    LoremIpsum = LoremIpsum + "Ut volutpat maximus orci, vel ultrices turpis consequat in. Cras eu euismod odio, quis dapibus neque. Mauris ut dui id lacus tincidunt dapibus a eget lacus. Aenean imperdiet fringilla justo, in pellentesque sapien placerat a. Donec nisi augue, tempor eu blandit sed, efficitur et mi. Donec efficitur lectus non eros placerat, at egestas diam iaculis. Integer sodales turpis congue sagittis tempor. Donec nec orci sit amet augue sagittis gravida id vitae massa. Donec nec tincidunt velit. Integer nisl dolor, mollis ut ultrices quis, fermentum sed nisi. Ut aliquam nisi at orci ullamcorper, at malesuada orci sodales. Nunc ut molestie mauris. Donec rutrum aliquet velit, nec maximus urna tincidunt sed."
    LoremIpsum = LoremIpsum + "Donec rhoncus massa leo, sit amet tempus dui rutrum ac. Suspendisse at rutrum libero. Proin pharetra maximus mollis. Morbi molestie quis tortor sed consectetur. Aenean ullamcorper iaculis pharetra. Maecenas et blandit nisl, quis scelerisque nisl. Donec vel tempor sem, ac consequat justo. Pellentesque quis libero euismod, feugiat lacus et, finibus eros. Aenean finibus sit amet massa consectetur semper. Ut hendrerit euismod ipsum. Pellentesque lorem leo, vulputate non orci ut, convallis semper ex. Nunc fermentum tempor sagittis. Aliquam erat volutpat. Vivamus fringilla finibus ex sed pharetra. Quisque pharetra dictum lectus, sit amet dapibus eros accumsan eu. Pellentesque at lectus eu ipsum congue mollis."
    LoremIpsum = LoremIpsum + "Nunc ac condimentum justo. Phasellus vel massa aliquet, pulvinar ligula in, ornare enim. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Nulla molestie nisi nec posuere tincidunt. Cras eget bibendum ante, id facilisis augue. Class aptent taciti sociosqu ad litora torquent per conubia nostra, per inceptos himenaeos. Donec id turpis maximus, semper orci ac, tristique arcu. Sed euismod sapien sed nisl scelerisque suscipit. Pellentesque mollis volutpat orci quis eleifend. Curabitur et nisi est. Integer finibus commodo pretium."
    LoremIpsum = LoremIpsum + "Nunc dignissim tincidunt blandit. Sed quis arcu a lacus cursus mollis vitae nec eros. Ut dignissim cursus massa, nec elementum leo pellentesque ut. Aenean nec nunc scelerisque dui maximus consequat. Morbi diam augue, ullamcorper eget dictum id, venenatis vitae ipsum. Nulla facilisi. Aliquam mollis leo sed leo tempus aliquam. Donec a erat at justo rhoncus commodo ut eu erat. Ut vitae nisl rutrum, consectetur leo quis, laoreet diam. Sed metus leo, semper sit amet volutpat ut, placerat eu diam. Donec malesuada nunc ac pretium hendrerit."
    LoremIpsum = LoremIpsum + "Integer viverra pulvinar augue. Nulla et erat sed ante suscipit vulputate. Proin a iaculis nisl. Pellentesque convallis lorem sit amet euismod tincidunt. Pellentesque nisl mauris, dignissim sed imperdiet vel, tristique a orci. Integer ut scelerisque quam. Sed scelerisque lectus ut convallis malesuada. Morbi vehicula hendrerit magna in placerat."
    LoremIpsum = LoremIpsum + "Integer non interdum sapien. Praesent dictum risus erat, non iaculis dolor bibendum accumsan. Fusce fermentum ultricies ultrices. Ut condimentum elit vitae scelerisque euismod. Suspendisse massa ante, interdum in nisl quis, blandit."
    LoremIpsum = LoremIpsum + LoremIpsum + LoremIpsum + LoremIpsum + LoremIpsum
    
    LoremIpsumWords = Split(LoremIpsum, " ")
    
    If (NumberOfWords > UBound(LoremIpsumWords)) Then
        GetLoremIpsum = LoremIpsum
        Exit Function
    End If
    
    LoremResult = LoremIpsumWords(0)
    WordCount = 1
    
    Do While (WordCount < NumberOfWords)
    
        If (Len(LoremResult & " " & LoremIpsumWords(WordCount)) <= MaxLength) Or NumberOfWords <= 2 Then
            LoremResult = LoremResult & " " & LoremIpsumWords(WordCount)
        End If
        
        WordCount = WordCount + 1
    Loop
    
    GetLoremIpsum = LoremResult
    
End Function

         */

        #region ObjectsToggleAutoSize
        public void ObjectsToggleAutoSize(IRibbonControl control)
        {
            var myDocument = ThisAddIn.application.ActiveWindow;
            if (!(myDocument.Selection.Type == PpSelectionType.ppSelectionShapes || myDocument.Selection.Type == PpSelectionType.ppSelectionText))
            {
                MessageBox.Show("No shapes selected.");
            }
            else
            {
                if (myDocument.Selection.ShapeRange.HasTextFrame == MsoTriState.msoCTrue || myDocument.Selection.ShapeRange.HasTextFrame == MsoTriState.msoTrue)
                {
                    TextFrame2 obj = myDocument.Selection.ShapeRange.TextFrame2;
                    switch ((int)obj.AutoSize)
                    {
                        case 0:
                            obj.AutoSize = (MsoAutoSize)1;
                            break;
                        case 1:
                            obj.AutoSize = (MsoAutoSize)2;
                            break;
                        case 2:
                            obj.AutoSize = (MsoAutoSize)0;
                            break;
                        case -2:
                            obj.AutoSize = (MsoAutoSize)0;
                            break;
                    }
                }
                else
                {
                    MessageBox.Show("No text capable shape selected.");
                }
            }
        }

        /*
         Sub ObjectsToggleAutoSize()
    
    Set myDocument = Application.ActiveWindow
    
    If Not (myDocument.Selection.Type = ppSelectionShapes Or myDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No shapes selected."
    Else
    
    If myDocument.Selection.ShapeRange.HasTextFrame Then
    
    With myDocument.Selection.ShapeRange.TextFrame2
        If .AutoSize = 0 Then
            .AutoSize = 1
        ElseIf .AutoSize = 1 Then
            .AutoSize = 2
        ElseIf .AutoSize = 2 Then
            .AutoSize = 0
        ElseIf .AutoSize = -2 Then
            .AutoSize = 0
        End If
        
    End With
    
    Else
    
    MsgBox "No text capable shape selected."
    
    End If
    
    End If
End Sub

         */
        #endregion
        #region ObjectsAutoSizeNone
        public void ObjectsAutoSizeNone(IRibbonControl control)
        {
            var myDocument = ThisAddIn.application.ActiveWindow;
            if (!(myDocument.Selection.Type == PpSelectionType.ppSelectionShapes || myDocument.Selection.Type == PpSelectionType.ppSelectionText))
            {
                MessageBox.Show("No shapes selected.");
            }
            else
            {
                if(myDocument.Selection.ShapeRange.HasTextFrame ==MsoTriState.msoCTrue || myDocument.Selection.ShapeRange.HasTextFrame == MsoTriState.msoTrue)
                {
                    myDocument.Selection.ShapeRange.TextFrame2.AutoSize =(MsoAutoSize ) 0;
                }
                else
                {
                    MessageBox.Show("No text capable shape selected.");
                }
            }
        }
        /*
         Sub ObjectsAutoSizeNone()
    
    Set myDocument = Application.ActiveWindow
    
    If Not (myDocument.Selection.Type = ppSelectionShapes Or myDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No shapes selected."
    Else
    
    If myDocument.Selection.ShapeRange.HasTextFrame Then
    
    With myDocument.Selection.ShapeRange.TextFrame2
        .AutoSize = 0
    End With
    
    Else
    
    MsgBox "No text capable shape selected."
    
    End If
    
    End If
End Sub
         */
        #endregion
        #region ObjectsAutoSizeTextToFitShape

        public void ObjectsAutoSizeTextToFitShape(IRibbonControl control)
        {
            var myDocument = ThisAddIn.application.ActiveWindow;
            if (!(myDocument.Selection.Type == PpSelectionType.ppSelectionShapes || myDocument.Selection.Type == PpSelectionType.ppSelectionText))
            {
                MessageBox.Show("No shapes selected.");
            }
            else
            {
                if (myDocument.Selection.ShapeRange.HasTextFrame == MsoTriState.msoCTrue || myDocument.Selection.ShapeRange.HasTextFrame == MsoTriState.msoTrue)
                {
                    myDocument.Selection.ShapeRange.TextFrame2.AutoSize = (MsoAutoSize)2;
                }
                else
                {
                    MessageBox.Show("No text capable shape selected.");
                }
            }
        }
        /*
         Sub ObjectsAutoSizeTextToFitShape()
    
    Set myDocument = Application.ActiveWindow
    
    If Not (myDocument.Selection.Type = ppSelectionShapes Or myDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No shapes selected."
    Else
    
    If myDocument.Selection.ShapeRange.HasTextFrame Then
    
    With myDocument.Selection.ShapeRange.TextFrame2
        .AutoSize = 2
    End With
    
    Else
    
    MsgBox "No text capable shape selected."
    
    End If
    
    End If
End Sub
         */
        #endregion
        #region ObjectsAutoSizeShapeToFitText
        public void ObjectsAutoSizeShapeToFitText(IRibbonControl control)
        {
            var myDocument = ThisAddIn.application.ActiveWindow;
            if (!(myDocument.Selection.Type == PpSelectionType.ppSelectionShapes || myDocument.Selection.Type == PpSelectionType.ppSelectionText))
            {
                MessageBox.Show("No shapes selected.");
            }
            else
            {
                if (myDocument.Selection.ShapeRange.HasTextFrame == MsoTriState.msoCTrue || myDocument.Selection.ShapeRange.HasTextFrame == MsoTriState.msoTrue)
                {
                    myDocument.Selection.ShapeRange.TextFrame2.AutoSize = (MsoAutoSize)1;
                }
                else
                {
                    MessageBox.Show("No text capable shape selected.");
                }
            }
        }
        /*
         Sub ObjectsAutoSizeShapeToFitText()
    
    Set myDocument = Application.ActiveWindow
    
    If Not (myDocument.Selection.Type = ppSelectionShapes Or myDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No shapes selected."
    Else
    
    If myDocument.Selection.ShapeRange.HasTextFrame Then
    
    With myDocument.Selection.ShapeRange.TextFrame2
        .AutoSize = 1
    End With
    
    Else
    
    MsgBox "No text capable shape selected."
    
    End If
    
    End If
End Sub
         */
        #endregion
        #region TextInsertEuro
        public void TextInsertEuro(IRibbonControl control)
        {
            ObjectsTextInsertSpecialCharacter(8364);
        }
        /*
         Sub TextInsertEuro()
    ObjectsTextInsertSpecialCharacter 8364
End Sub
         */
        #endregion
        #region TextInsertCopyright
        public void TextInsertCopyright(IRibbonControl control)
        {
            ObjectsTextInsertSpecialCharacter(169);
        }
        /*
         Sub TextInsertCopyright()
    ObjectsTextInsertSpecialCharacter 169
End Sub
         */
        #endregion
        #region TextBulletsTicks
        public void TextBulletsTicks(IRibbonControl control)
        {
            var myDocument = ThisAddIn.application.ActiveWindow;
            try
            {
                var bullet = myDocument.Selection.TextRange.ParagraphFormat.Bullet;
                bullet.Character = 252;
                bullet.Visible = MsoTriState.msoTrue ;
                bullet.Font.Name = "Wingdings";
                bullet.Font.Color.RGB = Color.FromArgb(0, 128, 0).ToArgb();
            }
            catch(Exception ex) { }
        }
        /*
         Sub TextBulletsTicks()
 
    Set myDocument = Application.ActiveWindow
     
    On Error Resume Next
    With myDocument.Selection.TextRange.ParagraphFormat.Bullet
        
        .Character = 252
        .Visible = True
        .Font.Name = "Wingdings"
        .Font.Color = RGB(0, 128, 0)
        
    End With
    On Error GoTo 0
    
End Sub
         */
        #endregion

        #region TextBulletsCrosses
        public void TextBulletsCrosses(IRibbonControl control)
        {
            var myDocument = ThisAddIn.application.ActiveWindow;
            try
            {
                var bullet = myDocument.Selection.TextRange.ParagraphFormat.Bullet;
                bullet.Character = 215;
                bullet.Visible = MsoTriState.msoTrue;
                bullet.Font.Name = "Calibri";
                bullet.Font.Color.RGB = Color.FromArgb(255, 0, 0).ToArgb();
            }
            catch (Exception ex) { }
        }
        /*
         Sub TextBulletsCrosses()
    
    Set myDocument = Application.ActiveWindow
   
    On Error Resume Next
    With myDocument.Selection.TextRange.ParagraphFormat.Bullet
        
        .Character = 215
        .Visible = True
        .Font.Name = "Calibri"
        .Font.Color = RGB(255, 0, 0)
        
    End With
    On Error GoTo 0
    
End Subv
         */
        #endregion

        #region ObjectsIncreaseLineSpacing
        public void ObjectsIncreaseLineSpacing(IRibbonControl control)
        {
            var myDocument = ThisAddIn.application.ActiveWindow;
            if (!(myDocument.Selection.Type == PpSelectionType.ppSelectionShapes || myDocument.Selection.Type == PpSelectionType.ppSelectionText))
            {
                MessageBox.Show("No shapes selected.");
            }
            else
            {
                if (myDocument.Selection.ShapeRange.HasTextFrame == MsoTriState.msoCTrue || myDocument.Selection.ShapeRange.HasTextFrame == MsoTriState.msoTrue)
                {
                    myDocument.Selection.ShapeRange.TextFrame.TextRange.ParagraphFormat.SpaceWithin +=(float) 0.1;
                }
                else
                {
                    MessageBox.Show("No text capable shapes selected.");
                }
            }
        }
        /*
         
Sub ObjectsIncreaseLineSpacing()
    
    Set myDocument = Application.ActiveWindow
      
    If Not (myDocument.Selection.Type = ppSelectionShapes Or myDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No shapes selected."
    Else
    
    If myDocument.Selection.ShapeRange.HasTextFrame Then
    
    With myDocument.Selection.ShapeRange.TextFrame.TextRange.ParagraphFormat
        .SpaceWithin = .SpaceWithin + 0.1
    End With
    
    Else
    
    MsgBox "No text capable shapes selected."
    
    End If
    
    End If
    
End Sub
         */
        #endregion

        #region ObjectsDecreaseLineSpacing

        public void ObjectsDecreaseLineSpacing(IRibbonControl control)
        {
            var myDocument = ThisAddIn.application.ActiveWindow;
            if (!(myDocument.Selection.Type == PpSelectionType.ppSelectionShapes || myDocument.Selection.Type == PpSelectionType.ppSelectionText))
            {
                MessageBox.Show("No shapes selected.");
            }
            else
            {
                if (myDocument.Selection.ShapeRange.HasTextFrame == MsoTriState.msoCTrue || myDocument.Selection.ShapeRange.HasTextFrame == MsoTriState.msoTrue)
                {
                    if (myDocument.Selection.ShapeRange.TextFrame.TextRange.ParagraphFormat.SpaceWithin <= 0.1)
                        myDocument.Selection.ShapeRange.TextFrame.TextRange.ParagraphFormat.SpaceWithin = 0;
                    else
                        myDocument.Selection.ShapeRange.TextFrame.TextRange.ParagraphFormat.SpaceWithin -= (float)0.1;
                }
                else
                {
                    MessageBox.Show("No text capable shapes selected.");
                }
            }
        }

        /*
         Sub ObjectsDecreaseLineSpacing()
    
    Set myDocument = Application.ActiveWindow
    
    If Not (myDocument.Selection.Type = ppSelectionShapes Or myDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No shapes selected."
    Else
    
    If myDocument.Selection.ShapeRange.HasTextFrame Then
    
    With myDocument.Selection.ShapeRange.TextFrame.TextRange.ParagraphFormat
    
        If .SpaceWithin <= 0.1 Then
            .SpaceWithin = 0
        Else
            .SpaceWithin = .SpaceWithin - 0.1
        End If
        
    End With
    
    Else
    
    MsgBox "No text capable shapes selected."
    
    End If
    
    End If
    
End Sub

         */
        #endregion
        #region ObjectsRemoveText
        public void ObjectsRemoveText(IRibbonControl control)
        {
            var myDocument = ThisAddIn.application.ActiveWindow;
            if (!(myDocument.Selection.Type == PpSelectionType.ppSelectionShapes || myDocument.Selection.Type == PpSelectionType.ppSelectionText))
            {
                MessageBox.Show("No shapes selected.");
            }
            else
            {
                if (myDocument.Selection.ShapeRange.HasTextFrame == MsoTriState.msoCTrue || myDocument.Selection.ShapeRange.HasTextFrame == MsoTriState.msoTrue)
                {
                    myDocument.Selection.ShapeRange.TextFrame.TextRange.Text = "";
                }
                else
                {
                    MessageBox.Show("No text capable shape selected.");
                }
            }
        }
        /*
         
Sub ObjectsRemoveText()
    Set myDocument = Application.ActiveWindow
    
    If Not (myDocument.Selection.Type = ppSelectionShapes Or myDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No shapes selected."
    Else
    
    If myDocument.Selection.ShapeRange.HasTextFrame Then
    
    myDocument.Selection.ShapeRange.TextFrame.TextRange.Text = ""
    
    Else
    
    MsgBox "No text capable shape selected."
    
    End If
    
    End If
    
End Sub
         */
        #endregion
        #region ObjectsSwapTextNoFormatting
        public void ObjectsSwapTextNoFormatting(IRibbonControl control)
        {
            string text1, text2;
            var myDocument = ThisAddIn.application.ActiveWindow;
            if(myDocument.Selection.Type!=PpSelectionType.ppSelectionShapes)
            {
                MessageBox.Show("No shapes selected.");
            }
            else
            {
                if (myDocument.Selection.ShapeRange.Count == 2)
                {
                    if((myDocument.Selection.ShapeRange[1].HasTextFrame ==MsoTriState.msoCTrue|| myDocument.Selection.ShapeRange[1].HasTextFrame == MsoTriState.msoTrue) && (myDocument.Selection.ShapeRange[2].HasTextFrame == MsoTriState.msoCTrue || myDocument.Selection.ShapeRange[2].HasTextFrame == MsoTriState.msoTrue))
                    {
                        text1 = myDocument.Selection.ShapeRange[1].TextFrame.TextRange.Text;
                        text2 = myDocument.Selection.ShapeRange[2].TextFrame.TextRange.Text;
                        myDocument.Selection.ShapeRange[1].TextFrame.TextRange.Text = text2;
                        myDocument.Selection.ShapeRange[2].TextFrame.TextRange.Text = text1;
                    }
                    else
                    {
                        MessageBox.Show("Select two shapes that (can) have text.");
                    }
                }
                else
                {
                    MessageBox.Show("Select two shapes to swap their text.");
                }
            }
        }
        /*
         Sub ObjectsSwapTextNoFormatting()

    Dim text1, text2 As String
    Set myDocument = Application.ActiveWindow

    If Not myDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
    Else

    If myDocument.Selection.ShapeRange.Count = 2 Then

    If myDocument.Selection.ShapeRange(1).HasTextFrame And myDocument.Selection.ShapeRange(2).HasTextFrame Then

    text1 = myDocument.Selection.ShapeRange(1).TextFrame.TextRange.Text
    text2 = myDocument.Selection.ShapeRange(2).TextFrame.TextRange.Text
    myDocument.Selection.ShapeRange(1).TextFrame.TextRange.Text = text2
    myDocument.Selection.ShapeRange(2).TextFrame.TextRange.Text = text1

    Else

    MsgBox "Select two shapes that (can) have text."

    End If


    Else

    MsgBox "Select two shapes to swap their text."

    End If

    End If

End Sub
         */
        #endregion

        #region ObjectsSwapText
        public void ObjectsSwapText(IRibbonControl control)
        {

            var myDocument = ThisAddIn.application.ActiveWindow;
            if (myDocument.Selection.Type != PpSelectionType.ppSelectionShapes)
            {
                MessageBox.Show("No shapes selected.");
            }
            else
            {
                if (myDocument.Selection.ShapeRange.Count == 2)
                {
                    if ((myDocument.Selection.ShapeRange[1].HasTextFrame == MsoTriState.msoCTrue || myDocument.Selection.ShapeRange[1].HasTextFrame == MsoTriState.msoTrue) && (myDocument.Selection.ShapeRange[2].HasTextFrame == MsoTriState.msoCTrue || myDocument.Selection.ShapeRange[2].HasTextFrame == MsoTriState.msoTrue))
                    {
                        Shape SlidePlaceHolder;
                        SlidePlaceHolder = ThisAddIn.application.ActivePresentation.Slides[1].Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 0, 0, 100, 100);
                        try
                        {
                            myDocument.Selection.ShapeRange[1].TextFrame.TextRange.Cut();
                            SlidePlaceHolder.TextFrame.TextRange.Paste();

                            myDocument.Selection.ShapeRange[2].TextFrame.TextRange.Cut();
                            myDocument.Selection.ShapeRange[1].TextFrame.TextRange.Paste();

                            SlidePlaceHolder.TextFrame.TextRange.Cut();
                            myDocument.Selection.ShapeRange[2].TextFrame.TextRange.Paste();
                            SlidePlaceHolder.Delete();
                        }
                        catch(Exception ex)
                        {

                        }
                    }
                    else
                    {
                        MessageBox.Show("Select two shapes that (can) have text.");
                    }
                }
                else
                {
                    MessageBox.Show("Select two shapes to swap their text.");
                }
            }
        }
        /*
         Sub ObjectsSwapText()

    Set myDocument = Application.ActiveWindow

    If Not myDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
    Else

    If myDocument.Selection.ShapeRange.Count = 2 Then

    If myDocument.Selection.ShapeRange(1).HasTextFrame And myDocument.Selection.ShapeRange(2).HasTextFrame Then

    Dim SlidePlaceHolder As PowerPoint.Shape
    Set SlidePlaceHolder = ActivePresentation.Slides(1).Shapes.AddShape(Type:=msoShapeRectangle, Left:=0, Top:=0, Width:=100, Height:=100)

    myDocument.Selection.ShapeRange(1).TextFrame.TextRange.Cut
    SlidePlaceHolder.TextFrame.TextRange.Paste

    myDocument.Selection.ShapeRange(2).TextFrame.TextRange.Cut
    myDocument.Selection.ShapeRange(1).TextFrame.TextRange.Paste

    SlidePlaceHolder.TextFrame.TextRange.Cut
    myDocument.Selection.ShapeRange(2).TextFrame.TextRange.Paste

    SlidePlaceHolder.Delete

    Else

    MsgBox "Select two shapes that (can) have text."

    End If


    Else

    MsgBox "Select two shapes to swap their text."

    End If

    End If

End Sub

         */
        #endregion

        #region ShowChangeSpellCheckLanguageForm
        public void ShowChangeSpellCheckLanguageForm(IRibbonControl control)
        {
            string[] LanguageNames=new string[217];
            LanguageNames[1] = "Afrikaans";
    LanguageNames[2] = "Albanian";
    LanguageNames[3] = "Amharic" ;
            LanguageNames[4] = "Arabic";
    LanguageNames[5] = "Arabic Algeria" ;
    LanguageNames[6] = "Arabic Bahrain" ;
    LanguageNames[7] = "Arabic Egypt"   ;
    LanguageNames[8] = "Arabic Iraq"    ;
    LanguageNames[9] = "Arabic Jordan"  ;
    LanguageNames[10] = "Arabic Kuwait" ;
    LanguageNames[11] = "Arabic Lebanon";
    LanguageNames[12] = "Arabic Libya"  ;
    LanguageNames[13] = "Arabic Morocco";
    LanguageNames[14] = "Arabic Oman"   ;
    LanguageNames[15] = "Arabic Qatar"  ;
    LanguageNames[16] = "Arabic Syria"  ;
    LanguageNames[17] = "Arabic Tunisia";
    LanguageNames[18] = "Arabic UAE"    ;
    LanguageNames[19] = "Arabic Yemen"  ;
    LanguageNames[20] = "Armenian"      ;
            LanguageNames[21] = "Assamese"; 
    LanguageNames[22] = "Azerbaijani Cyrillic"       ;
    LanguageNames[23] = "Azerbaijani Latin"          ;
    LanguageNames[24] = "Basque (Basque)"            ;
    LanguageNames[25] = "Belgian Dutch"              ;
    LanguageNames[26] = "Belgian French"             ;
    LanguageNames[27] = "Bengali"                    ;
            LanguageNames[28] = "Bosnian";
    LanguageNames[29] = "Bosnian Bosnia Herzegovina Cyrillic";
    LanguageNames[30] = "Bosnian Bosnia Herzegovina Latin"   ;
    LanguageNames[31] = "Portuguese (Brazil)"                ;
    LanguageNames[32] = "Bulgarian"                          ;
    LanguageNames[33] = "Burmese"                            ;
    LanguageNames[34] = "Belarusian"                         ;
    LanguageNames[35] = "Catalan"                            ;
    LanguageNames[36] = "Cherokee"                           ;
    LanguageNames[37] = "Chinese Hong Kong SAR"              ;
    LanguageNames[38] = "Chinese Macao SAR"                  ;
    LanguageNames[39] = "Chinese Singapore"                  ;
    LanguageNames[40] = "Croatian"                           ;
    LanguageNames[41] = "Czech"                              ;
    LanguageNames[42] = "Danish"                             ;
    LanguageNames[43] = "Divehi"                             ;
    LanguageNames[44] = "Dutch"                              ;
    LanguageNames[45] = "Edo"                                ;
    LanguageNames[46] = "English AUS"                        ;
    LanguageNames[47] = "English Belize"                     ;
    LanguageNames[48] = "English Canadian"                   ;
    LanguageNames[49] = "English Caribbean"                  ;
    LanguageNames[50] = "English Indonesia"                  ;
    LanguageNames[51] = "English Ireland"                    ;
    LanguageNames[52] = "English Jamaica"                    ;
    LanguageNames[53] = "English NewZealand"                 ;
    LanguageNames[54] = "English Philippines"                ;
    LanguageNames[55] = "English South Africa"               ;
    LanguageNames[56] = "English Trinidad Tobago"            ;
    LanguageNames[57] = "English UK"                         ;
    LanguageNames[58] = "English US"                         ;
    LanguageNames[59] = "English Zimbabwe"                   ;
    LanguageNames[60] = "Estonian"                           ;
    LanguageNames[61] = "Faeroese"                           ;
    LanguageNames[62] = "Farsi"                              ;
    LanguageNames[63] = "Filipino"                           ;
    LanguageNames[64] = "Finnish"                            ;
    LanguageNames[65] = "French"                             ;
    LanguageNames[66] = "French Cameroon"                    ;
    LanguageNames[67] = "French Canadian"                    ;
    LanguageNames[68] = "French Coted Ivoire"                ;
    LanguageNames[69] = "French Haiti"                       ;
    LanguageNames[70] = "French Luxembourg"                  ;
    LanguageNames[71] = "French Mali"                        ;
    LanguageNames[72] = "French Monaco"                      ;
    LanguageNames[73] = "French Morocco"                     ;
    LanguageNames[74] = "French Reunion"                     ;
    LanguageNames[75] = "French Senegal"                     ;
    LanguageNames[76] = "French West Indies"                 ;
    LanguageNames[77] = "French Congo DRC"                   ;
    LanguageNames[78] = "Frisian Netherlands"                ;
    LanguageNames[79] = "Fulfulde"                           ;
    LanguageNames[80] = "Irish (Ireland)"                    ;
    LanguageNames[81] = "Scottish Gaelic"                    ;
    LanguageNames[82] = "Galician"                           ;
    LanguageNames[83] = "Georgian"                           ;
    LanguageNames[84] = "German"                             ;
    LanguageNames[85] = "German Austria"                     ;
    LanguageNames[86] = "German Liechtenstein"               ;
    LanguageNames[87] = "German Luxembourg"                  ;
    LanguageNames[88] = "Greek"                              ;
    LanguageNames[89] = "Guarani"                            ;
    LanguageNames[90] = "Gujarati"                           ;
    LanguageNames[91] = "Hausa"                              ;
    LanguageNames[92] = "Hawaiian"                           ;
    LanguageNames[93] = "Hebrew"                             ;
    LanguageNames[94] = "Hindi"                              ;
    LanguageNames[95] = "Hungarian"                          ;
    LanguageNames[96] = "Ibibio"                             ;
    LanguageNames[97] = "Icelandic"                          ;
    LanguageNames[98] = "Igbo"                               ;
    LanguageNames[99] = "Indonesian"                         ;
    LanguageNames[100] = "Inuktitut"                         ;
    LanguageNames[101] = "Italian"                           ;
    LanguageNames[102] = "Japanese"                          ;
    LanguageNames[103] = "Kannada"                           ;
    LanguageNames[104] = "Kanuri"                            ;
    LanguageNames[105] = "Kashmiri"                          ;
    LanguageNames[106] = "Kashmiri Devanagari"               ;
    LanguageNames[107] = "Kazakh"                            ;
    LanguageNames[108] = "Khmer"                             ;
    LanguageNames[109] = "Kirghiz"                           ;
    LanguageNames[110] = "Konkani"                           ;
    LanguageNames[111] = "Korean"                            ;
    LanguageNames[112] = "Kyrgyz"                            ;
    LanguageNames[113] = "Lao"                               ;
    LanguageNames[114] = "Latin"                             ;
    LanguageNames[115] = "Latvian"                           ;
    LanguageNames[116] = "Lithuanian"                        ;
    LanguageNames[117] = "Macedonian FYROM"                  ;
    LanguageNames[118] = "Malayalam"                         ;
    LanguageNames[119] = "Malay Brunei Darussalam"           ;
    LanguageNames[120] = "Malaysian"                         ;
    LanguageNames[121] = "Maltese"                           ;
    LanguageNames[122] = "Manipuri"                          ;
    LanguageNames[123] = "Maori"                             ;
    LanguageNames[124] = "Marathi"                           ;
    LanguageNames[125] = "Mexican Spanish"                   ;
    LanguageNames[126] = "Mixed"                             ;
    LanguageNames[127] = "Mongolian"                         ;
    LanguageNames[128] = "Nepali"                            ;
    LanguageNames[129] = "No specified"                      ;
    LanguageNames[130] = "No proofing"                       ;
    LanguageNames[131] = "Norwegian Bokmol"                  ;
    LanguageNames[132] = "Norwegian Nynorsk"                 ;
    LanguageNames[133] = "Odia"                              ;
    LanguageNames[134] = "Oromo"                             ;
    LanguageNames[135] = "Pashto"                            ;
    LanguageNames[136] = "Polish"                            ;
    LanguageNames[137] = "Portuguese"                        ;
    LanguageNames[138] = "Punjabi"                           ;
    LanguageNames[139] = "Quechua Bolivia"                   ;
    LanguageNames[140] = "Quechua Ecuador"                   ;
    LanguageNames[141] = "Quechua Peru"                      ;
    LanguageNames[142] = "Rhaeto Romanic"                    ;
    LanguageNames[143] = "Romanian"                          ;
    LanguageNames[144] = "Romanian Moldova"                  ;
    LanguageNames[145] = "Russian"                           ;
    LanguageNames[146] = "Russian Moldova"                   ;
    LanguageNames[147] = "Sami Lappish"                      ;
    LanguageNames[148] = "Sanskrit"                          ;
            LanguageNames[149] = "Sepedi";
    LanguageNames[150] = "Serbian Bosnia Herzegovina Cyrillic";
    LanguageNames[151] = "Serbian Bosnia Herzegovina Latin"   ;
    LanguageNames[152] = "Serbian Cyrillic"                   ;
    LanguageNames[153] = "Serbian Latin"                      ;
    LanguageNames[154] = "Sesotho"                            ;
    LanguageNames[155] = "Simplified Chinese"                 ;
    LanguageNames[156] = "Sindhi"                             ;
    LanguageNames[157] = "Sindhi Pakistan"                    ;
    LanguageNames[158] = "Sinhalese"                          ;
    LanguageNames[159] = "Slovak"                             ;
    LanguageNames[160] = "Slovenian"                          ;
    LanguageNames[161] = "Somali"                             ;
    LanguageNames[162] = "Sorbian"                            ;
    LanguageNames[163] = "Spanish"                            ;
    LanguageNames[164] = "Spanish Argentina"                  ;
    LanguageNames[165] = "Spanish Bolivia"                    ;
    LanguageNames[166] = "Spanish Chile"                      ;
    LanguageNames[167] = "Spanish Colombia"                   ;
    LanguageNames[168] = "Spanish Costa Rica"                 ;
    LanguageNames[169] = "Spanish Dominican Republic"         ;
    LanguageNames[170] = "Spanish Ecuador"                    ;
    LanguageNames[171] = "Spanish El Salvador"                ;
    LanguageNames[172] = "Spanish Guatemala"                  ;
    LanguageNames[173] = "Spanish Honduras"                   ;
    LanguageNames[174] = "Spanish Modern Sort"                ;
    LanguageNames[175] = "Spanish Nicaragua"                  ;
    LanguageNames[176] = "Spanish Panama"                     ;
    LanguageNames[177] = "Spanish Paraguay"                   ;
    LanguageNames[178] = "Spanish Peru"                       ;
    LanguageNames[179] = "Spanish Puerto Rico"                ;
    LanguageNames[180] = "Spanish Uruguay"                    ;
    LanguageNames[181] = "Spanish Venezuela"                  ;
    LanguageNames[182] = "Sutu"                               ;
    LanguageNames[183] = "Swahili"                            ;
    LanguageNames[184] = "Swedish"                            ;
    LanguageNames[185] = "Swedish Finland"                    ;
    LanguageNames[186] = "Swiss French"                       ;
    LanguageNames[187] = "Swiss German"                       ;
    LanguageNames[188] = "Swiss Italian"                      ;
    LanguageNames[189] = "Syriac"                             ;
    LanguageNames[190] = "Tajik"                              ;
    LanguageNames[191] = "Tamazight"                          ;
    LanguageNames[192] = "Tamazight Latin"                    ;
    LanguageNames[193] = "Tamil"                              ;
    LanguageNames[194] = "Tatar"                              ;
    LanguageNames[195] = "Telugu"                             ;
    LanguageNames[196] = "Thai"                               ;
    LanguageNames[197] = "Tibetan"                            ;
    LanguageNames[198] = "Tigrigna Eritrea"                   ;
    LanguageNames[199] = "Tigrigna Ethiopic"                  ;
    LanguageNames[200] = "Traditional Chinese"                ;
    LanguageNames[201] = "Tsonga"                             ;
    LanguageNames[202] = "Tswana"                             ;
    LanguageNames[203] = "Turkish"                            ;
    LanguageNames[204] = "Turkmen"                            ;
    LanguageNames[205] = "Ukrainian"                          ;
    LanguageNames[206] = "Urdu"                               ;
    LanguageNames[207] = "Uzbek Cyrillic"                     ;
    LanguageNames[208] = "Uzbek Latin"                        ;
    LanguageNames[209] = "Venda"                              ;
    LanguageNames[210] = "Vietnamese"                         ;
    LanguageNames[211] = "Welsh"                              ;
    LanguageNames[212] = "Xhosa"                              ;
    LanguageNames[213] = "Yi"                                 ;
    LanguageNames[214] = "Yiddish"                            ;
    LanguageNames[215] = "Yoruba"                             ;
            LanguageNames[216] = "Zulu";


            ChangeSpellCheckLanguageForm f = new ChangeSpellCheckLanguageForm();
            f.ComboBox1.Items.Clear();
           for (int i = 1; i <= 216; i++)
                f.ComboBox1.Items.Add ( LanguageNames[i]);



            f.Show();
        }
        /*
         Sub ShowChangeSpellCheckLanguageForm()
    
    Dim LanguageNames(1 To 216) As String
    LanguageNames(1) = "Afrikaans"
    LanguageNames(2) = "Albanian"
    LanguageNames(3) = "Amharic"
    LanguageNames(4) = "Arabic"
    LanguageNames(5) = "Arabic Algeria"
    LanguageNames(6) = "Arabic Bahrain"
    LanguageNames(7) = "Arabic Egypt"
    LanguageNames(8) = "Arabic Iraq"
    LanguageNames(9) = "Arabic Jordan"
    LanguageNames(10) = "Arabic Kuwait"
    LanguageNames(11) = "Arabic Lebanon"
    LanguageNames(12) = "Arabic Libya"
    LanguageNames(13) = "Arabic Morocco"
    LanguageNames(14) = "Arabic Oman"
    LanguageNames(15) = "Arabic Qatar"
    LanguageNames(16) = "Arabic Syria"
    LanguageNames(17) = "Arabic Tunisia"
    LanguageNames(18) = "Arabic UAE"
    LanguageNames(19) = "Arabic Yemen"
    LanguageNames(20) = "Armenian"
    LanguageNames(21) = "Assamese"
    LanguageNames(22) = "Azerbaijani Cyrillic"
    LanguageNames(23) = "Azerbaijani Latin"
    LanguageNames(24) = "Basque (Basque)"
    LanguageNames(25) = "Belgian Dutch"
    LanguageNames(26) = "Belgian French"
    LanguageNames(27) = "Bengali"
    LanguageNames(28) = "Bosnian"
    LanguageNames(29) = "Bosnian Bosnia Herzegovina Cyrillic"
    LanguageNames(30) = "Bosnian Bosnia Herzegovina Latin"
    LanguageNames(31) = "Portuguese (Brazil)"
    LanguageNames(32) = "Bulgarian"
    LanguageNames(33) = "Burmese"
    LanguageNames(34) = "Belarusian"
    LanguageNames(35) = "Catalan"
    LanguageNames(36) = "Cherokee"
    LanguageNames(37) = "Chinese Hong Kong SAR"
    LanguageNames(38) = "Chinese Macao SAR"
    LanguageNames(39) = "Chinese Singapore"
    LanguageNames(40) = "Croatian"
    LanguageNames(41) = "Czech"
    LanguageNames(42) = "Danish"
    LanguageNames(43) = "Divehi"
    LanguageNames(44) = "Dutch"
    LanguageNames(45) = "Edo"
    LanguageNames(46) = "English AUS"
    LanguageNames(47) = "English Belize"
    LanguageNames(48) = "English Canadian"
    LanguageNames(49) = "English Caribbean"
    LanguageNames(50) = "English Indonesia"
    LanguageNames(51) = "English Ireland"
    LanguageNames(52) = "English Jamaica"
    LanguageNames(53) = "English NewZealand"
    LanguageNames(54) = "English Philippines"
    LanguageNames(55) = "English South Africa"
    LanguageNames(56) = "English Trinidad Tobago"
    LanguageNames(57) = "English UK"
    LanguageNames(58) = "English US"
    LanguageNames(59) = "English Zimbabwe"
    LanguageNames(60) = "Estonian"
    LanguageNames(61) = "Faeroese"
    LanguageNames(62) = "Farsi"
    LanguageNames(63) = "Filipino"
    LanguageNames(64) = "Finnish"
    LanguageNames(65) = "French"
    LanguageNames(66) = "French Cameroon"
    LanguageNames(67) = "French Canadian"
    LanguageNames(68) = "French Coted Ivoire"
    LanguageNames(69) = "French Haiti"
    LanguageNames(70) = "French Luxembourg"
    LanguageNames(71) = "French Mali"
    LanguageNames(72) = "French Monaco"
    LanguageNames(73) = "French Morocco"
    LanguageNames(74) = "French Reunion"
    LanguageNames(75) = "French Senegal"
    LanguageNames(76) = "French West Indies"
    LanguageNames(77) = "French Congo DRC"
    LanguageNames(78) = "Frisian Netherlands"
    LanguageNames(79) = "Fulfulde"
    LanguageNames(80) = "Irish (Ireland)"
    LanguageNames(81) = "Scottish Gaelic"
    LanguageNames(82) = "Galician"
    LanguageNames(83) = "Georgian"
    LanguageNames(84) = "German"
    LanguageNames(85) = "German Austria"
    LanguageNames(86) = "German Liechtenstein"
    LanguageNames(87) = "German Luxembourg"
    LanguageNames(88) = "Greek"
    LanguageNames(89) = "Guarani"
    LanguageNames(90) = "Gujarati"
    LanguageNames(91) = "Hausa"
    LanguageNames(92) = "Hawaiian"
    LanguageNames(93) = "Hebrew"
    LanguageNames(94) = "Hindi"
    LanguageNames(95) = "Hungarian"
    LanguageNames(96) = "Ibibio"
    LanguageNames(97) = "Icelandic"
    LanguageNames(98) = "Igbo"
    LanguageNames(99) = "Indonesian"
    LanguageNames(100) = "Inuktitut"
    LanguageNames(101) = "Italian"
    LanguageNames(102) = "Japanese"
    LanguageNames(103) = "Kannada"
    LanguageNames(104) = "Kanuri"
    LanguageNames(105) = "Kashmiri"
    LanguageNames(106) = "Kashmiri Devanagari"
    LanguageNames(107) = "Kazakh"
    LanguageNames(108) = "Khmer"
    LanguageNames(109) = "Kirghiz"
    LanguageNames(110) = "Konkani"
    LanguageNames(111) = "Korean"
    LanguageNames(112) = "Kyrgyz"
    LanguageNames(113) = "Lao"
    LanguageNames(114) = "Latin"
    LanguageNames(115) = "Latvian"
    LanguageNames(116) = "Lithuanian"
    LanguageNames(117) = "Macedonian FYROM"
    LanguageNames(118) = "Malayalam"
    LanguageNames(119) = "Malay Brunei Darussalam"
    LanguageNames(120) = "Malaysian"
    LanguageNames(121) = "Maltese"
    LanguageNames(122) = "Manipuri"
    LanguageNames(123) = "Maori"
    LanguageNames(124) = "Marathi"
    LanguageNames(125) = "Mexican Spanish"
    LanguageNames(126) = "Mixed"
    LanguageNames(127) = "Mongolian"
    LanguageNames(128) = "Nepali"
    LanguageNames(129) = "No specified"
    LanguageNames(130) = "No proofing"
    LanguageNames(131) = "Norwegian Bokmol"
    LanguageNames(132) = "Norwegian Nynorsk"
    LanguageNames(133) = "Odia"
    LanguageNames(134) = "Oromo"
    LanguageNames(135) = "Pashto"
    LanguageNames(136) = "Polish"
    LanguageNames(137) = "Portuguese"
    LanguageNames(138) = "Punjabi"
    LanguageNames(139) = "Quechua Bolivia"
    LanguageNames(140) = "Quechua Ecuador"
    LanguageNames(141) = "Quechua Peru"
    LanguageNames(142) = "Rhaeto Romanic"
    LanguageNames(143) = "Romanian"
    LanguageNames(144) = "Romanian Moldova"
    LanguageNames(145) = "Russian"
    LanguageNames(146) = "Russian Moldova"
    LanguageNames(147) = "Sami Lappish"
    LanguageNames(148) = "Sanskrit"
    LanguageNames(149) = "Sepedi"
    LanguageNames(150) = "Serbian Bosnia Herzegovina Cyrillic"
    LanguageNames(151) = "Serbian Bosnia Herzegovina Latin"
    LanguageNames(152) = "Serbian Cyrillic"
    LanguageNames(153) = "Serbian Latin"
    LanguageNames(154) = "Sesotho"
    LanguageNames(155) = "Simplified Chinese"
    LanguageNames(156) = "Sindhi"
    LanguageNames(157) = "Sindhi Pakistan"
    LanguageNames(158) = "Sinhalese"
    LanguageNames(159) = "Slovak"
    LanguageNames(160) = "Slovenian"
    LanguageNames(161) = "Somali"
    LanguageNames(162) = "Sorbian"
    LanguageNames(163) = "Spanish"
    LanguageNames(164) = "Spanish Argentina"
    LanguageNames(165) = "Spanish Bolivia"
    LanguageNames(166) = "Spanish Chile"
    LanguageNames(167) = "Spanish Colombia"
    LanguageNames(168) = "Spanish Costa Rica"
    LanguageNames(169) = "Spanish Dominican Republic"
    LanguageNames(170) = "Spanish Ecuador"
    LanguageNames(171) = "Spanish El Salvador"
    LanguageNames(172) = "Spanish Guatemala"
    LanguageNames(173) = "Spanish Honduras"
    LanguageNames(174) = "Spanish Modern Sort"
    LanguageNames(175) = "Spanish Nicaragua"
    LanguageNames(176) = "Spanish Panama"
    LanguageNames(177) = "Spanish Paraguay"
    LanguageNames(178) = "Spanish Peru"
    LanguageNames(179) = "Spanish Puerto Rico"
    LanguageNames(180) = "Spanish Uruguay"
    LanguageNames(181) = "Spanish Venezuela"
    LanguageNames(182) = "Sutu"
    LanguageNames(183) = "Swahili"
    LanguageNames(184) = "Swedish"
    LanguageNames(185) = "Swedish Finland"
    LanguageNames(186) = "Swiss French"
    LanguageNames(187) = "Swiss German"
    LanguageNames(188) = "Swiss Italian"
    LanguageNames(189) = "Syriac"
    LanguageNames(190) = "Tajik"
    LanguageNames(191) = "Tamazight"
    LanguageNames(192) = "Tamazight Latin"
    LanguageNames(193) = "Tamil"
    LanguageNames(194) = "Tatar"
    LanguageNames(195) = "Telugu"
    LanguageNames(196) = "Thai"
    LanguageNames(197) = "Tibetan"
    LanguageNames(198) = "Tigrigna Eritrea"
    LanguageNames(199) = "Tigrigna Ethiopic"
    LanguageNames(200) = "Traditional Chinese"
    LanguageNames(201) = "Tsonga"
    LanguageNames(202) = "Tswana"
    LanguageNames(203) = "Turkish"
    LanguageNames(204) = "Turkmen"
    LanguageNames(205) = "Ukrainian"
    LanguageNames(206) = "Urdu"
    LanguageNames(207) = "Uzbek Cyrillic"
    LanguageNames(208) = "Uzbek Latin"
    LanguageNames(209) = "Venda"
    LanguageNames(210) = "Vietnamese"
    LanguageNames(211) = "Welsh"
    LanguageNames(212) = "Xhosa"
    LanguageNames(213) = "Yi"
    LanguageNames(214) = "Yiddish"
    LanguageNames(215) = "Yoruba"
    LanguageNames(216) = "Zulu"
    
    ChangeSpellCheckLanguageForm.ComboBox1.Clear
    For i = 1 To 216
        ChangeSpellCheckLanguageForm.ComboBox1.AddItem LanguageNames(i)
    Next
    
    ChangeSpellCheckLanguageForm.Show
    
End Sub

         */
        #endregion

        #region ObjectsCloneRight
        public void ObjectsCloneRight(IRibbonControl control)
        {
            var myDocument = ThisAddIn.application.ActiveWindow;
            double  OldTop, OldLeft;
            if (myDocument.Selection.Type!=PpSelectionType.ppSelectionShapes )
            {
                MessageBox.Show("No shapes selected.");
            }
            else
            {
                if (myDocument.Selection.ShapeRange.Count == 1)
                {
                    OldTop = myDocument.Selection.ShapeRange.Top;
                    OldLeft = myDocument.Selection.ShapeRange.Left;
                    var SlideShape = myDocument.Selection.ShapeRange.Duplicate();
                    SlideShape.Top =(float ) OldTop;
                    SlideShape.Left =(float )( OldLeft + SlideShape.Width);
                    SlideShape.Select();
                }
                else
                {
                    var ShapesToDuplicate = myDocument.Selection.ShapeRange.Group();
                    OldTop = ShapesToDuplicate.Top;
                    OldLeft = ShapesToDuplicate.Left;


                    var SlideShape = ShapesToDuplicate.Duplicate();
                    SlideShape.Top =(float ) OldTop;
                   SlideShape.Left =(float ) (OldLeft + SlideShape.Width);
                    ShapesToDuplicate.Ungroup();
                    SlideShape.Ungroup().Select();
                }
            }
        
        
        }
        /*
         Sub ObjectsCloneRight()
    
    Set myDocument = Application.ActiveWindow
    
    
    Dim OldTop, OldLeft As Double
    
    If Not myDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
        
    ElseIf myDocument.Selection.ShapeRange.Count = 1 Then
        
        OldTop = myDocument.Selection.ShapeRange.Top
        OldLeft = myDocument.Selection.ShapeRange.Left
        
        Set SlideShape = myDocument.Selection.ShapeRange.Duplicate
        
        With SlideShape
            .Top = OldTop
            .Left = OldLeft + SlideShape.Width
        End With
        
        SlideShape.Select
        
    Else
        
        Set ShapesToDuplicate = myDocument.Selection.ShapeRange.Group
        
        OldTop = ShapesToDuplicate.Top
        OldLeft = ShapesToDuplicate.Left
        
        Set SlideShape = ShapesToDuplicate.Duplicate
        
        With SlideShape
            .Top = OldTop
            .Left = OldLeft + SlideShape.Width
        End With
        
        ShapesToDuplicate.Ungroup
        SlideShape.Ungroup.Select
        
    End If
    
End Sub

         */
        #endregion

        #region ObjectsCloneDown
        public void ObjectsCloneDown(IRibbonControl control)
        {
            var myDocument = ThisAddIn.application.ActiveWindow;
            double OldTop, OldLeft;
            if (myDocument.Selection.Type != PpSelectionType.ppSelectionShapes)
            {
                MessageBox.Show("No shapes selected.");
            }
            else
            {
                if (myDocument.Selection.ShapeRange.Count == 1)
                {
                    OldTop = myDocument.Selection.ShapeRange.Top;
                    OldLeft = myDocument.Selection.ShapeRange.Left;
                    var SlideShape = myDocument.Selection.ShapeRange.Duplicate();
                    SlideShape.Top = (float)OldTop + SlideShape.Height;
                    SlideShape.Left = (float)(OldLeft );
                    SlideShape.Select();
                }
                else
                {
                    var ShapesToDuplicate = myDocument.Selection.ShapeRange.Group();
                    OldTop = ShapesToDuplicate.Top;
                    OldLeft = ShapesToDuplicate.Left;


                    var SlideShape = ShapesToDuplicate.Duplicate();
                    SlideShape.Top = (float)OldTop + SlideShape.Height;
                    SlideShape.Left = (float)(OldLeft);
                    ShapesToDuplicate.Ungroup();
                    SlideShape.Ungroup().Select();
                }
            }

        }

        /*
         Sub ObjectsCloneDown()
    
    Set myDocument = Application.ActiveWindow
    Dim OldTop, OldLeft As Double
    
    If Not myDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
        
    ElseIf myDocument.Selection.ShapeRange.Count = 1 Then
        
        OldTop = myDocument.Selection.ShapeRange.Top
        OldLeft = myDocument.Selection.ShapeRange.Left
        
        Set SlideShape = myDocument.Selection.ShapeRange.Duplicate
        
        With SlideShape
            .Top = OldTop + SlideShape.Height
            .Left = OldLeft
        End With
        
        SlideShape.Select
        
    Else
        
        Set ShapesToDuplicate = myDocument.Selection.ShapeRange.Group
        
        OldTop = ShapesToDuplicate.Top
        OldLeft = ShapesToDuplicate.Left
        
        Set SlideShape = ShapesToDuplicate.Duplicate
        
        With SlideShape
            .Top = OldTop + SlideShape.Height
            .Left = OldLeft
        End With
        
        ShapesToDuplicate.Ungroup
        SlideShape.Ungroup.Select
        
    End If
    
End Sub

         */
        #endregion

        #region ObjectsCopyRoundedCorner
        public void ObjectsCopyRoundedCorner(IRibbonControl control)
        {
            Microsoft.Office.Interop.PowerPoint.Shape SlideShape;
            var myDocument = ThisAddIn.application.ActiveWindow;
            if(myDocument.Selection.Type != PpSelectionType.ppSelectionShapes)
            {
                MessageBox.Show("No shapes selected.");
            }
            else
            {
                float  ShapeRadius, ShapeRadius2=0;
                if(ThisAddIn.application.ActiveWindow.Selection.ShapeRange[1].Adjustments.Count > 0)
                {
                    ShapeRadius = myDocument.Selection.ShapeRange[1].Adjustments[1] / (1 / (myDocument.Selection.ShapeRange[1].Height + myDocument.Selection.ShapeRange[1].Width));
                    if (myDocument.Selection.ShapeRange[1].Adjustments.Count > 1)
                        ShapeRadius2 = myDocument.Selection.ShapeRange[1].Adjustments[2] / (1 / (myDocument.Selection.ShapeRange[1].Height + myDocument.Selection.ShapeRange[1].Width));

                    foreach (Shape ss in ThisAddIn.application.ActiveWindow.Selection.ShapeRange) {
                     ss.AutoShapeType = myDocument.Selection.ShapeRange[1].AutoShapeType;
                        ss.Adjustments[1] = (1 / (ss.Height + ss.Width)) * ShapeRadius;
                        if (myDocument.Selection.ShapeRange[1].Adjustments.Count > 1)
                            ss.Adjustments[2] = (1 / (ss.Height + ss.Width)) * ShapeRadius2;
            
                     }

                }
            }
        }

        /*
         Sub ObjectsCopyRoundedCorner()
    Dim SlideShape  As PowerPoint.Shape
    Set myDocument = Application.ActiveWindow
    
    If Not myDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
    Else
    
    Dim ShapeRadius As Single
    If Application.ActiveWindow.Selection.ShapeRange(1).Adjustments.Count > 0 Then
    
    ShapeRadius = myDocument.Selection.ShapeRange(1).Adjustments(1) / (1 / (myDocument.Selection.ShapeRange(1).Height + myDocument.Selection.ShapeRange(1).Width))
    
    If myDocument.Selection.ShapeRange(1).Adjustments.Count > 1 Then
        ShapeRadius2 = myDocument.Selection.ShapeRange(1).Adjustments(2) / (1 / (myDocument.Selection.ShapeRange(1).Height + myDocument.Selection.ShapeRange(1).Width))
    End If
    
    For Each SlideShape In ActiveWindow.Selection.ShapeRange
        With SlideShape
            .AutoShapeType = myDocument.Selection.ShapeRange(1).AutoShapeType
            .Adjustments(1) = (1 / (SlideShape.Height + SlideShape.Width)) * ShapeRadius
            If myDocument.Selection.ShapeRange(1).Adjustments.Count > 1 Then
                .Adjustments(2) = (1 / (SlideShape.Height + SlideShape.Width)) * ShapeRadius2
            End If
        End With
    Next
    
    End If
    
    End If
    
End Sub

         */
        #endregion

        #region ObjectsCopyShapeTypeAndAdjustments

        public void ObjectsCopyShapeTypeAndAdjustments(IRibbonControl control)
        {
            Microsoft.Office.Interop.PowerPoint.Shape SlideShape;
            var myDocument = ThisAddIn.application.ActiveWindow;
            if (myDocument.Selection.Type != PpSelectionType.ppSelectionShapes)
            {
                MessageBox.Show("No shapes selected.");
            }
            else
            {
                long AdjustmentsCount, ShapeCount;
                for (ShapeCount = 2; ShapeCount <= ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Count; ShapeCount++)
                {
                    myDocument.Selection.ShapeRange[ShapeCount].AutoShapeType = myDocument.Selection.ShapeRange[1].AutoShapeType;
                    for (AdjustmentsCount = 1; AdjustmentsCount <= myDocument.Selection.ShapeRange[1].Adjustments.Count; AdjustmentsCount++)
                        myDocument.Selection.ShapeRange[ShapeCount].Adjustments[(int)AdjustmentsCount] = myDocument.Selection.ShapeRange[1].Adjustments[(int)AdjustmentsCount];

                }
                
            }
        }
        /*
         Sub ObjectsCopyShapeTypeAndAdjustments()
            Dim SlideShape  As PowerPoint.Shape
            Set myDocument = Application.ActiveWindow

            If Not myDocument.Selection.Type = ppSelectionShapes Then
                MsgBox "No shapes selected."
            Else

            Dim AdjustmentsCount As Long
            Dim ShapeCount  As Long

            For ShapeCount = 2 To ActiveWindow.Selection.ShapeRange.Count

                myDocument.Selection.ShapeRange(ShapeCount).AutoShapeType = myDocument.Selection.ShapeRange(1).AutoShapeType

                For AdjustmentsCount = 1 To myDocument.Selection.ShapeRange(1).Adjustments.Count

                    myDocument.Selection.ShapeRange(ShapeCount).Adjustments(AdjustmentsCount) = myDocument.Selection.ShapeRange(1).Adjustments(AdjustmentsCount)

                Next AdjustmentsCount

            Next ShapeCount

            End If

        End Sub

         */
        #endregion

        #region RectifyLines
        public void RectifyLines(IRibbonControl control)
        {
           
            var myDocument = ThisAddIn.application.ActiveWindow;
            if(myDocument.Selection.Type==PpSelectionType.ppSelectionShapes)
            {
                foreach (Microsoft.Office.Interop.PowerPoint.Shape LineShape in myDocument.Selection.ShapeRange )
                {
                   if (LineShape.Fill.Type ==(MsoFillType )(-2) && LineShape.AutoShapeType == (MsoAutoShapeType ) (-2)) {


                        if (LineShape.Width > LineShape.Height) 
                            LineShape.Height = 0;
                        else
                            LineShape.Width = 0;
                        
                    }
                }
            }
            else
            {
                MessageBox.Show( "No shape selected.");
            }

        }
        /*
         Sub RectifyLines()
    
    Dim LineShape   As Shape
    Set myDocument = Application.ActiveWindow
    
    If myDocument.Selection.Type = ppSelectionShapes Then
    
    For Each LineShape In myDocument.Selection.ShapeRange
        
        With LineShape
            
            If .Fill.Type = -2 And .AutoShapeType = -2 Then
                
                If .Width > .Height Then
                    .Height = 0
                Else
                    .Width = 0
                End If
            End If
        End With
        
    Next
    
    Else
        MsgBox "No shape selected."
    End If
    
End Sub

         */
        #endregion

        #region ConnectRectangleShapes
        public void ConnectRectangleShapesRightToLeft(IRibbonControl control)
        {
            ConnectRectangleShapes ("RightToLeft");
        }
        public void ConnectRectangleShapesLeftToRight(IRibbonControl control)
        {
            ConnectRectangleShapes("LeftToRight");
        }
        public void ConnectRectangleShapesBottomToTop(IRibbonControl control)
        {
            ConnectRectangleShapes("BottomToTop");
        }
        public void ConnectRectangleShapesTopToBottom(IRibbonControl control)
        {
            ConnectRectangleShapes("TopToBottom");
        }
        /*
         
         */
        #endregion

        #region ObjectsSelectBySameFillColor
        public void ObjectsSelectBySameFillColor(IRibbonControl control)
        {
            var myDocument = ThisAddIn.application.ActiveWindow;
            if(myDocument.Selection.Type !=PpSelectionType.ppSelectionShapes)
            {
                MessageBox.Show("No shapes selected.");
            }
            else
            {
                Microsoft.Office.Interop.PowerPoint.Shape SlideShape;
                List<string> SelectedShapes = new List<string>();
                long ShapeCount = 1;
                 SlideShape =ThisAddIn.application. ActiveWindow.Selection.ShapeRange[1];
                SelectedShapes.Add(SlideShape.Name );


                foreach(Microsoft.Office.Interop.PowerPoint.Shape   SlideShapeToCheck in myDocument.View.Slide.Shapes)
                {
                   if(SlideShapeToCheck .Fill.ForeColor.RGB==SlideShape .Fill.ForeColor.RGB  && SlideShapeToCheck.Type !=MsoShapeType.msoPlaceholder && (SlideShapeToCheck .Fill.Visible==MsoTriState.msoCTrue || SlideShapeToCheck.Fill.Visible == MsoTriState.msoTrue))
                    {
                        if(SlideShapeToCheck.Name !=SlideShape.Name)
                        {
                            SelectedShapes.Add(SlideShapeToCheck.Name);
                        }
                    }
                    
                }
                myDocument.View.Slide.Shapes.Range(SelectedShapes.ToArray()).Select();
            }
        }
        /*
         Sub ObjectsSelectBySameFillColor()
    
    Set myDocument = Application.ActiveWindow
    
    If Not myDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
    Else
    
    Dim SlideShape, SlideShapeToCheck As PowerPoint.Shape
    Dim SelectedShapes() As String
    Dim ShapeCount  As Long
    ShapeCount = 1
    
    Set SlideShape = ActiveWindow.Selection.ShapeRange(1)
    
    ReDim Preserve SelectedShapes(0)
    SelectedShapes(0) = SlideShape.Name
    
    For Each SlideShapeToCheck In myDocument.View.Slide.Shapes
        
        If (SlideShapeToCheck.Fill.ForeColor.RGB = SlideShape.Fill.ForeColor.RGB) And (SlideShapeToCheck.Type <> msoPlaceholder) And (SlideShapeToCheck.Fill.Visible = True) Then
            
            If (SlideShapeToCheck.Name <> SlideShape.Name) Then
                ReDim Preserve SelectedShapes(ShapeCount + 1)
                SelectedShapes(ShapeCount) = SlideShapeToCheck.Name
                ShapeCount = ShapeCount + 1
            End If
        End If
        
    Next SlideShapeToCheck
    myDocument.View.Slide.Shapes.Range(SelectedShapes).Select
    
    End If
    
End Sub
         */
        #endregion

        #region ObjectsSelectBySameLineColor
        public void ObjectsSelectBySameLineColor(IRibbonControl control)
        {
            var myDocument = ThisAddIn.application.ActiveWindow;
            if (myDocument.Selection.Type != PpSelectionType.ppSelectionShapes)
            {
                MessageBox.Show("No shapes selected.");
            }
            else
            {
                Microsoft.Office.Interop.PowerPoint.Shape SlideShape;
                List<string> SelectedShapes = new List<string>();
                long ShapeCount = 1;
                SlideShape = ThisAddIn.application.ActiveWindow.Selection.ShapeRange[1];
                
                SelectedShapes.Add(SlideShape.Name);


                foreach (Microsoft.Office.Interop.PowerPoint.Shape SlideShapeToCheck in myDocument.View.Slide.Shapes)
                {
                    if (SlideShapeToCheck.Line .ForeColor.RGB == SlideShape.Line .ForeColor.RGB && SlideShapeToCheck.Type != MsoShapeType.msoPlaceholder && (SlideShapeToCheck.Line .Visible == MsoTriState.msoCTrue || SlideShapeToCheck.Line .Visible == MsoTriState.msoTrue))
                    {
                        if (SlideShapeToCheck.Name != SlideShape.Name)
                        {
                            SelectedShapes.Add(SlideShapeToCheck.Name);
                        }
                    }

                }
                myDocument.View.Slide.Shapes.Range(SelectedShapes.ToArray()).Select();
            }
        }
        /*
         Sub ObjectsSelectBySameLineColor()
    
    Set myDocument = Application.ActiveWindow
    
    If Not myDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
    Else
    
    Dim SlideShape, SlideShapeToCheck As PowerPoint.Shape
    Dim SelectedShapes() As String
    Dim ShapeCount  As Long
    ShapeCount = 1
    
    Set SlideShape = ActiveWindow.Selection.ShapeRange(1)
    
    ReDim Preserve SelectedShapes(0)
    SelectedShapes(0) = SlideShape.Name
    
    For Each SlideShapeToCheck In myDocument.View.Slide.Shapes
        
        If (SlideShapeToCheck.Line.ForeColor.RGB = SlideShape.Line.ForeColor.RGB) And (SlideShapeToCheck.Type <> msoPlaceholder) And (SlideShapeToCheck.Line.Visible = True) Then
            
            If (SlideShapeToCheck.Name <> SlideShape.Name) Then
                ReDim Preserve SelectedShapes(ShapeCount + 1)
                SelectedShapes(ShapeCount) = SlideShapeToCheck.Name
                ShapeCount = ShapeCount + 1
            End If
        End If
        
    Next SlideShapeToCheck
    myDocument.View.Slide.Shapes.Range(SelectedShapes).Select
    
    End If
End Sub

         */
        #endregion

        #region ObjectsSelectBySameFillAndLineColor
        public void ObjectsSelectBySameFillAndLineColor(IRibbonControl control)
        {
            var myDocument = ThisAddIn.application.ActiveWindow;
            if (myDocument.Selection.Type != PpSelectionType.ppSelectionShapes)
            {
                MessageBox.Show("No shapes selected.");
            }
            else
            {
                Microsoft.Office.Interop.PowerPoint.Shape SlideShape;
                List<string> SelectedShapes = new List<string>();
                long ShapeCount = 1;
                SlideShape = ThisAddIn.application.ActiveWindow.Selection.ShapeRange[1];

                SelectedShapes.Add(SlideShape.Name);


                foreach (Microsoft.Office.Interop.PowerPoint.Shape SlideShapeToCheck in myDocument.View.Slide.Shapes)
                {
                    if (SlideShapeToCheck.Line.ForeColor.RGB == SlideShape.Line.ForeColor.RGB && SlideShapeToCheck.Fill .ForeColor.RGB == SlideShape.Fill .ForeColor.RGB && SlideShapeToCheck.Type != MsoShapeType.msoPlaceholder && (SlideShapeToCheck.Line.Visible == MsoTriState.msoCTrue || SlideShapeToCheck.Line.Visible == MsoTriState.msoTrue) && (SlideShapeToCheck.Fill.Visible == MsoTriState.msoCTrue || SlideShapeToCheck.Fill .Visible == MsoTriState.msoTrue))
                    {
                        if (SlideShapeToCheck.Name != SlideShape.Name)
                        {
                            SelectedShapes.Add(SlideShapeToCheck.Name);
                        }
                    }

                }
                myDocument.View.Slide.Shapes.Range(SelectedShapes.ToArray()).Select();
            }
        }

        /*
         Sub ObjectsSelectBySameFillAndLineColor()
    
    Set myDocument = Application.ActiveWindow
    
    If Not myDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
    Else
    
    Dim SlideShape, SlideShapeToCheck As PowerPoint.Shape
    Dim SelectedShapes() As String
    Dim ShapeCount  As Long
    ShapeCount = 1
    
    Set SlideShape = ActiveWindow.Selection.ShapeRange(1)
    
    ReDim Preserve SelectedShapes(0)
    SelectedShapes(0) = SlideShape.Name
    
    For Each SlideShapeToCheck In myDocument.View.Slide.Shapes
        
        If (SlideShapeToCheck.Fill.ForeColor.RGB = SlideShape.Fill.ForeColor.RGB) And (SlideShapeToCheck.Line.ForeColor.RGB = SlideShape.Line.ForeColor.RGB) And (SlideShapeToCheck.Type <> msoPlaceholder) And (SlideShapeToCheck.Fill.Visible = True) And (SlideShapeToCheck.Line.Visible = True) Then
            
            If (SlideShapeToCheck.Name <> SlideShape.Name) Then
                ReDim Preserve SelectedShapes(ShapeCount + 1)
                SelectedShapes(ShapeCount) = SlideShapeToCheck.Name
                ShapeCount = ShapeCount + 1
            End If
        End If
        
    Next SlideShapeToCheck
    myDocument.View.Slide.Shapes.Range(SelectedShapes).Select
    
    End If
End Sub

         */
        #endregion

        #region ObjectsSelectBySameWidthAndHeight
        public void ObjectsSelectBySameWidthAndHeight(IRibbonControl control)
        {
            var myDocument = ThisAddIn.application.ActiveWindow;
            if (myDocument.Selection.Type != PpSelectionType.ppSelectionShapes)
            {
                MessageBox.Show("No shapes selected.");
            }
            else
            {
                Microsoft.Office.Interop.PowerPoint.Shape SlideShape;
                List<string> SelectedShapes = new List<string>();
                long ShapeCount = 1;
                SlideShape = ThisAddIn.application.ActiveWindow.Selection.ShapeRange[1];

                SelectedShapes.Add(SlideShape.Name);


                foreach (Microsoft.Office.Interop.PowerPoint.Shape SlideShapeToCheck in myDocument.View.Slide.Shapes)
                {
                    if (SlideShapeToCheck.Width  == SlideShape.Width  && SlideShapeToCheck.Height  == SlideShape.Height  && SlideShapeToCheck.Type != MsoShapeType.msoPlaceholder && (SlideShapeToCheck.Fill.Visible == MsoTriState.msoCTrue || SlideShapeToCheck.Fill.Visible == MsoTriState.msoTrue))
                    {
                        if (SlideShapeToCheck.Name != SlideShape.Name)
                        {
                            SelectedShapes.Add(SlideShapeToCheck.Name);
                        }
                    }

                }
                myDocument.View.Slide.Shapes.Range(SelectedShapes.ToArray()).Select();
            }
        }
        /*
         
         Sub ObjectsSelectBySameWidthAndHeight()
    
    Set myDocument = Application.ActiveWindow
    
    If Not myDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
    Else
    
    Dim SlideShape, SlideShapeToCheck As PowerPoint.Shape
    Dim SelectedShapes() As String
    Dim ShapeCount  As Long
    ShapeCount = 1
    
    Set SlideShape = ActiveWindow.Selection.ShapeRange(1)
    
    ReDim Preserve SelectedShapes(0)
    SelectedShapes(0) = SlideShape.Name
    
    For Each SlideShapeToCheck In myDocument.View.Slide.Shapes
        
        If (SlideShapeToCheck.Width = SlideShape.Width) And (SlideShapeToCheck.Height = SlideShape.Height) And (SlideShapeToCheck.Type <> msoPlaceholder) And (SlideShapeToCheck.Fill.Visible = True) Then
            
            If (SlideShapeToCheck.Name <> SlideShape.Name) Then
                ReDim Preserve SelectedShapes(ShapeCount + 1)
                SelectedShapes(ShapeCount) = SlideShapeToCheck.Name
                ShapeCount = ShapeCount + 1
            End If
        End If
        
    Next SlideShapeToCheck
    myDocument.View.Slide.Shapes.Range(SelectedShapes).Select
    
    End If
End Sub

         */
        #endregion
        #region ObjectsSelectBySameWidth
        public void ObjectsSelectBySameWidth(IRibbonControl control)
        {
            var myDocument = ThisAddIn.application.ActiveWindow;
            if (myDocument.Selection.Type != PpSelectionType.ppSelectionShapes)
            {
                MessageBox.Show("No shapes selected.");
            }
            else
            {
                Microsoft.Office.Interop.PowerPoint.Shape SlideShape;
                List<string> SelectedShapes = new List<string>();
                long ShapeCount = 1;
                SlideShape = ThisAddIn.application.ActiveWindow.Selection.ShapeRange[1];

                SelectedShapes.Add(SlideShape.Name);


                foreach (Microsoft.Office.Interop.PowerPoint.Shape SlideShapeToCheck in myDocument.View.Slide.Shapes)
                {
                    if (SlideShapeToCheck.Width == SlideShape.Width  && SlideShapeToCheck.Type != MsoShapeType.msoPlaceholder && (SlideShapeToCheck.Fill.Visible == MsoTriState.msoCTrue || SlideShapeToCheck.Fill.Visible == MsoTriState.msoTrue))
                    {
                        if (SlideShapeToCheck.Name != SlideShape.Name)
                        {
                            SelectedShapes.Add(SlideShapeToCheck.Name);
                        }
                    }

                }
                myDocument.View.Slide.Shapes.Range(SelectedShapes.ToArray()).Select();
            }
        }
        /*
         Sub ObjectsSelectBySameWidth()
    
    Set myDocument = Application.ActiveWindow
    
    If Not myDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
    Else
    
    Dim SlideShape, SlideShapeToCheck As PowerPoint.Shape
    Dim SelectedShapes() As String
    Dim ShapeCount  As Long
    ShapeCount = 1
    
    Set SlideShape = ActiveWindow.Selection.ShapeRange(1)
    
    ReDim Preserve SelectedShapes(0)
    SelectedShapes(0) = SlideShape.Name
    
    For Each SlideShapeToCheck In myDocument.View.Slide.Shapes
        
        If (SlideShapeToCheck.Width = SlideShape.Width) And (SlideShapeToCheck.Type <> msoPlaceholder) And (SlideShapeToCheck.Fill.Visible = True) Then
            
            If (SlideShapeToCheck.Name <> SlideShape.Name) Then
                ReDim Preserve SelectedShapes(ShapeCount + 1)
                SelectedShapes(ShapeCount) = SlideShapeToCheck.Name
                ShapeCount = ShapeCount + 1
            End If
        End If
        
    Next SlideShapeToCheck
    myDocument.View.Slide.Shapes.Range(SelectedShapes).Select
    
    End If
End Sub

         */
        #endregion

        #region ObjectsSelectBySameHeight
        public void ObjectsSelectBySameHeight(IRibbonControl control)
        {
            var myDocument = ThisAddIn.application.ActiveWindow;
            if (myDocument.Selection.Type != PpSelectionType.ppSelectionShapes)
            {
                MessageBox.Show("No shapes selected.");
            }
            else
            {
                Microsoft.Office.Interop.PowerPoint.Shape SlideShape;
                List<string> SelectedShapes = new List<string>();
                long ShapeCount = 1;
                SlideShape = ThisAddIn.application.ActiveWindow.Selection.ShapeRange[1];

                SelectedShapes.Add(SlideShape.Name);


                foreach (Microsoft.Office.Interop.PowerPoint.Shape SlideShapeToCheck in myDocument.View.Slide.Shapes)
                {
                    if (SlideShapeToCheck.Height  == SlideShape.Height  && SlideShapeToCheck.Type != MsoShapeType.msoPlaceholder && (SlideShapeToCheck.Fill.Visible == MsoTriState.msoCTrue || SlideShapeToCheck.Fill.Visible == MsoTriState.msoTrue))
                    {
                        if (SlideShapeToCheck.Name != SlideShape.Name)
                        {
                            SelectedShapes.Add(SlideShapeToCheck.Name);
                        }
                    }

                }
                myDocument.View.Slide.Shapes.Range(SelectedShapes.ToArray()).Select();
            }
        }

        /*
         Sub ObjectsSelectBySameHeight()
    
    Set myDocument = Application.ActiveWindow
    
    If Not myDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
    Else
    
    Dim SlideShape, SlideShapeToCheck As PowerPoint.Shape
    Dim SelectedShapes() As String
    Dim ShapeCount  As Long
    ShapeCount = 1
    
    Set SlideShape = ActiveWindow.Selection.ShapeRange(1)
    
    ReDim Preserve SelectedShapes(0)
    SelectedShapes(0) = SlideShape.Name
    
    For Each SlideShapeToCheck In myDocument.View.Slide.Shapes
        
        If (SlideShapeToCheck.Height = SlideShape.Height) And (SlideShapeToCheck.Type <> msoPlaceholder) And (SlideShapeToCheck.Fill.Visible = True) Then
            
            If (SlideShapeToCheck.Name <> SlideShape.Name) Then
                ReDim Preserve SelectedShapes(ShapeCount + 1)
                SelectedShapes(ShapeCount) = SlideShapeToCheck.Name
                ShapeCount = ShapeCount + 1
            End If
        End If
        
    Next SlideShapeToCheck
    myDocument.View.Slide.Shapes.Range(SelectedShapes).Select
    
    End If
End Sub

         */
        #endregion
        #region ObjectsSelectBySameType
        public void ObjectsSelectBySameType(IRibbonControl control)
        {
            var myDocument = ThisAddIn.application.ActiveWindow;
            if (myDocument.Selection.Type != PpSelectionType.ppSelectionShapes)
            {
                MessageBox.Show("No shapes selected.");
            }
            else
            {
                Microsoft.Office.Interop.PowerPoint.Shape SlideShape;
                List<string> SelectedShapes = new List<string>();
                long ShapeCount = 1;
                SlideShape = ThisAddIn.application.ActiveWindow.Selection.ShapeRange[1];

                SelectedShapes.Add(SlideShape.Name);


                foreach (Microsoft.Office.Interop.PowerPoint.Shape SlideShapeToCheck in myDocument.View.Slide.Shapes)
                {
                    if (SlideShapeToCheck.AutoShapeType == SlideShape.AutoShapeType)
                    {
                        if (SlideShapeToCheck.Name != SlideShape.Name)
                        {
                            SelectedShapes.Add(SlideShapeToCheck.Name);
                        }
                    }

                }
                myDocument.View.Slide.Shapes.Range(SelectedShapes.ToArray()).Select();
            }
        }


        /*
         Sub ObjectsSelectBySameType()
    
    Set myDocument = Application.ActiveWindow
    
    If Not myDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
    Else
    
    Dim SlideShape, SlideShapeToCheck As PowerPoint.Shape
    Dim SelectedShapes() As String
    Dim ShapeCount  As Long
    ShapeCount = 1
    
    Set SlideShape = ActiveWindow.Selection.ShapeRange(1)
    
    ReDim Preserve SelectedShapes(0)
    SelectedShapes(0) = SlideShape.Name
    
    For Each SlideShapeToCheck In myDocument.View.Slide.Shapes
        
        If (SlideShapeToCheck.AutoShapeType = SlideShape.AutoShapeType) Then
            
            If (SlideShapeToCheck.Name <> SlideShape.Name) Then
                ReDim Preserve SelectedShapes(ShapeCount + 1)
                SelectedShapes(ShapeCount) = SlideShapeToCheck.Name
                ShapeCount = ShapeCount + 1
            End If
        End If
        
    Next SlideShapeToCheck
    myDocument.View.Slide.Shapes.Range(SelectedShapes).Select
    
    End If
    
End Sub

         */
        #endregion

        #region CopyPosition
        public void CopyPosition(IRibbonControl control)
        {
            var myDocument = ThisAddIn.application.ActiveWindow;
            if(myDocument.Selection.Type !=PpSelectionType.ppSelectionShapes)
            {
                MessageBox.Show  ("No shapes selected.");
            }
            else
            {
              ThisAddIn.TopToCopy = myDocument.Selection.ShapeRange[1].Top      ;
              ThisAddIn.LeftToCopy = myDocument.Selection.ShapeRange[1].Left    ;
              ThisAddIn.WidthToCopy = myDocument.Selection.ShapeRange[1].Width  ;
                ThisAddIn.HeightToCopy = myDocument.Selection.ShapeRange[1].Height;
              ThisAddIn.PositionCopied = true;


            }
        }
        /*
         Sub CopyPosition()
    
    Set myDocument = Application.ActiveWindow
    
    If Not myDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
    Else
    
    TopToCopy = myDocument.Selection.ShapeRange(1).Top
    LeftToCopy = myDocument.Selection.ShapeRange(1).Left
    WidthToCopy = myDocument.Selection.ShapeRange(1).Width
    HeightToCopy = myDocument.Selection.ShapeRange(1).Height
    PositionCopied = True
    
    End If
End Sub
         */
        #endregion

        #region PastePosition
        public void PastePosition(IRibbonControl control)
        {
            var myDocument = ThisAddIn.application.ActiveWindow;
            if(myDocument.Selection.Type !=PpSelectionType.ppSelectionShapes)
            {
                MessageBox.Show("No shapes selected.");
            }
            else
            {
                if(ThisAddIn.PositionCopied)
                {
                    myDocument.Selection.ShapeRange[1].Top = (float )ThisAddIn.TopToCopy;
                    myDocument.Selection.ShapeRange[1].Left =(float )ThisAddIn. LeftToCopy;
                }
                else
                {
                    MessageBox.Show("No dimensions available. First copy position / dimension of a shape.");
                }
            }
        }
        /*
         Sub PastePosition()
    Set myDocument = Application.ActiveWindow
    
    If Not myDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
    Else
    
    If PositionCopied = True Then
        myDocument.Selection.ShapeRange(1).Top = TopToCopy
        myDocument.Selection.ShapeRange(1).Left = LeftToCopy
    Else
        MsgBox "No dimensions available. First copy position / dimension of a shape."
    End If
    
    End If
End Sub
         */
        #endregion

        #region PastePositionAndDimensions
        public void PastePositionAndDimensions(IRibbonControl control)
        {
            var myDocument = ThisAddIn.application.ActiveWindow;
            if (myDocument.Selection.Type != PpSelectionType.ppSelectionShapes)
            {
                MessageBox.Show("No shapes selected.");
            }
            else
            {
                if (ThisAddIn.PositionCopied)
                {
                    myDocument.Selection.ShapeRange[1].Top = (float)ThisAddIn.TopToCopy;
                    myDocument.Selection.ShapeRange[1].Left = (float)ThisAddIn.LeftToCopy;
                    myDocument.Selection.ShapeRange[1].Width = (float)ThisAddIn.WidthToCopy;
                    myDocument.Selection.ShapeRange[1].Height = (float)ThisAddIn.HeightToCopy;
                }
                else
                {
                    MessageBox.Show("No dimensions available. First copy position / dimension of a shape.");
                }
            }

        }
        /*
         Sub PastePositionAndDimensions()
    
    Set myDocument = Application.ActiveWindow
    
    If Not myDocument.Selection.Type = ppSelectionShapes Then
        MsgBox "No shapes selected."
    Else
    
    If PositionCopied = True Then
        myDocument.Selection.ShapeRange(1).Top = TopToCopy
        myDocument.Selection.ShapeRange(1).Left = LeftToCopy
        myDocument.Selection.ShapeRange(1).Width = WidthToCopy
        myDocument.Selection.ShapeRange(1).Height = HeightToCopy
    Else
        MsgBox "No dimensions available. First copy position / dimension of a shape."
    End If
    
    End If
End Sub
         */
        #endregion
        #region ShowFormCopyShapeToMultipleSlides
        public void ShowFormCopyShapeToMultipleSlides(IRibbonControl control)
        {
            var myDocument = ThisAddIn.application.ActiveWindow;
            Random r = new Random();
            var RandomNumber = r.Next(1000000);
            CopyShapeToMultipleSlidesForm frm = new CopyShapeToMultipleSlidesForm();
            frm.AllSlidesListBox.Items.Clear();
            if(myDocument.Selection.Type ==PpSelectionType.ppSelectionShapes)
            {
                frm.ShapeIdentifierTextBox.Text = "NewShape" + RandomNumber.ToString();
                if (ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Tags["INSTRUMENTA CROSS-SLIDE SHAPE"].ToString() != "") {
                    frm.ShapeIdentifierTextBox.Text = ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Tags["INSTRUMENTA CROSS-SLIDE SHAPE"].ToString();
                    //frm.ShapeIdentifierTextBox.Text = ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Tags["INSTRUMENTA CROSS-SLIDE SHAPE"].ToString();
                 }
                string StorylineText;
                float  CurrentSlide;
                CurrentSlide = 0;
                try
                {
                    for(int SlideCount = 1; SlideCount<= ThisAddIn.application.ActivePresentation.Slides.Count; SlideCount++)
                    {
                        if(ThisAddIn.application. ActivePresentation.Slides[SlideCount].SlideNumber != ThisAddIn.application.ActiveWindow.Selection.SlideRange.SlideNumber)
                        {
                             StorylineText = "Untitled";


                
                foreach (Shape   SlidePlaceHolder in ThisAddIn .application . ActivePresentation.Slides[SlideCount].Shapes.Placeholders) {
                                if (SlidePlaceHolder.PlaceholderFormat.Type ==PpPlaceholderType.ppPlaceholderTitle) {
                                    StorylineText = SlidePlaceHolder.TextFrame.TextRange.Text;
                                    break;
                                }
                }



                            frm.AllSlidesListBox.Items.Add(new { });
                            //CopyShapeToMultipleSlidesForm.AllSlidesListBox.List(SlideCount - 1 - CurrentSlide, 0) = ActivePresentation.Slides(SlideCount).SlideNumber
                            frm.AllSlidesListBox.Items [(int)SlideCount - 1 - (int)CurrentSlide] = StorylineText;
                //CopyShapeToMultipleSlidesForm.AllSlidesListBox.List(SlideCount - 1 - CurrentSlide, 2) = ActivePresentation.Slides(SlideCount).SlideID


                        }
                        else
                        {
                            CurrentSlide = 1;
                        }
                    }
                    frm.Show();
                }
                catch(Exception ex)
                {

                }


            }
            else
            {
                MessageBox.Show("No shapes selected.");
            }
        }
        /*
         Sub ShowFormCopyShapeToMultipleSlides()
    Set myDocument = Application.ActiveWindow
    
    RandomNumber = Round(Rnd() * 1000000, 0)
    
    CopyShapeToMultipleSlidesForm.AllSlidesListBox.Clear
    CopyShapeToMultipleSlidesForm.AllSlidesListBox.ColumnCount = 3
    CopyShapeToMultipleSlidesForm.AllSlidesListBox.ColumnWidths = "15;300;0"
    
    If myDocument.Selection.Type = ppSelectionShapes Then
        
        CopyShapeToMultipleSlidesForm.ShapeIdentifierTextBox.Value = "NewShape" + Str(RandomNumber)
        
        If Not Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA CROSS-SLIDE SHAPE") = "" Then
            CopyShapeToMultipleSlidesForm.ShapeIdentifierTextBox.Value = Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA CROSS-SLIDE SHAPE")
            CopyShapeToMultipleSlidesForm.ShapeIdentifierTextBox.Text = Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA CROSS-SLIDE SHAPE")
        End If
        
        Dim StorylineText As String
        Dim CurrentSlide As Long
        CurrentSlide = 0
        
        On Error Resume Next
        
        For SlideCount = 1 To ActivePresentation.Slides.Count
            
            If Not ActivePresentation.Slides(SlideCount).SlideNumber = Application.ActiveWindow.Selection.SlideRange.SlideNumber Then
                
                StorylineText = "Untitled"
                
                On Error Resume Next
                For Each SlidePlaceHolder In ActivePresentation.Slides(SlideCount).Shapes.Placeholders
                    
                    If SlidePlaceHolder.PlaceholderFormat.Type = ppPlaceholderTitle Then
                        StorylineText = SlidePlaceHolder.TextFrame.TextRange.Text
                        Exit For
                    End If
                Next SlidePlaceHolder
                On Error GoTo 0
                
                CopyShapeToMultipleSlidesForm.AllSlidesListBox.AddItem
                CopyShapeToMultipleSlidesForm.AllSlidesListBox.List(SlideCount - 1 - CurrentSlide, 0) = ActivePresentation.Slides(SlideCount).SlideNumber
                CopyShapeToMultipleSlidesForm.AllSlidesListBox.List(SlideCount - 1 - CurrentSlide, 1) = StorylineText
                CopyShapeToMultipleSlidesForm.AllSlidesListBox.List(SlideCount - 1 - CurrentSlide, 2) = ActivePresentation.Slides(SlideCount).SlideID
                
            Else
                CurrentSlide = 1
                
            End If
            
        Next SlideCount
        On Error GoTo 0
        
        CopyShapeToMultipleSlidesForm.Show
        
    Else
        MsgBox "No shapes selected."
    End If
End Sub
         */
        #endregion

        #region UpdateTaggedShapePositionAndDimensions
        public void UpdateTaggedShapePositionAndDimensions(IRibbonControl control)
        {
            var myDocument = ThisAddIn.application.ActiveWindow;
            string CrossSlideShapeId;
            if(myDocument.Selection.Type ==PpSelectionType.ppSelectionShapes)
            {
                if (ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Tags["INSTRUMENTA CROSS-SLIDE SHAPE"].ToString() != "")
                {
                    CrossSlideShapeId = ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Tags["INSTRUMENTA CROSS-SLIDE SHAPE"];
                    for (int SlideCount = 1; SlideCount <=ThisAddIn.application. ActivePresentation.Slides.Count; SlideCount++) {
                        foreach (Shape Shape in ThisAddIn.application. ActivePresentation.Slides[SlideCount].Shapes) {


                            if (Shape.Tags["INSTRUMENTA CROSS-SLIDE SHAPE"] == CrossSlideShapeId) {


                                Shape.Top = ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Top;
                                    Shape.Left = ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Left;
                                    Shape.Width = ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Width;
                                    Shape.Height = ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Height;

                    }


                }
            }


                }
                else
                {
                    MessageBox.Show("This shape does Not have a tag.");
                }
            }
            else
            {
                MessageBox.Show( "No shape selected.");
            }
        }

        /*
         Sub UpdateTaggedShapePositionAndDimensions()
    Set myDocument = Application.ActiveWindow
    Dim CrossSlideShapeId As String
    
    If myDocument.Selection.Type = ppSelectionShapes Then
        
        If Not Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA CROSS-SLIDE SHAPE") = "" Then
            CrossSlideShapeId = Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA CROSS-SLIDE SHAPE")
            
            For SlideCount = 1 To ActivePresentation.Slides.Count
                For Each Shape In ActivePresentation.Slides(SlideCount).Shapes
                    
                    If Shape.Tags("INSTRUMENTA CROSS-SLIDE SHAPE") = CrossSlideShapeId Then
                        
                        With Shape
                            .Top = Application.ActiveWindow.Selection.ShapeRange.Top
                            .Left = Application.ActiveWindow.Selection.ShapeRange.Left
                            .Width = Application.ActiveWindow.Selection.ShapeRange.Width
                            .Height = Application.ActiveWindow.Selection.ShapeRange.Height
                            
                        End With
                        
                    End If
                    
                Next
            Next
            
        Else
            MsgBox "This shape does Not have a tag."
        End If
        
    Else
        MsgBox "No shape selected."
    End If
    
End Sub
         */
        #endregion

        #region DeleteTaggedShapes
        public void DeleteTaggedShapes(IRibbonControl control)
        {
            var myDocument = ThisAddIn.application.ActiveWindow;
            string CrossSlideShapeId;
               if(myDocument.Selection.Type==PpSelectionType.ppSelectionShapes)
            {
                if (ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Tags["INSTRUMENTA CROSS-SLIDE SHAPE"] != "")
                {
                    CrossSlideShapeId = ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Tags["INSTRUMENTA CROSS-SLIDE SHAPE"];


                    for (int SlideCount = 1; SlideCount <= ThisAddIn.application.ActivePresentation.Slides.Count; SlideCount++) {
                        foreach(Shape  Shape in ThisAddIn.application.ActivePresentation.Slides[SlideCount].Shapes) {



                            if (Shape.Tags["INSTRUMENTA CROSS-SLIDE SHAPE"] == CrossSlideShapeId)
                                Shape.Delete();

                    

                }
                    }


                }
                else
                {
                    MessageBox.Show("This shape does Not have a tag.");
                }
            }
            else
            {
                MessageBox.Show("No shape selected.");
            }

        }
        /*
         Sub DeleteTaggedShapes()
    Set myDocument = Application.ActiveWindow
    Dim CrossSlideShapeId As String
    
    If myDocument.Selection.Type = ppSelectionShapes Then
        
        If Not Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA CROSS-SLIDE SHAPE") = "" Then
            CrossSlideShapeId = Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA CROSS-SLIDE SHAPE")
            
            For SlideCount = 1 To ActivePresentation.Slides.Count
                For Each Shape In ActivePresentation.Slides(SlideCount).Shapes
                    
                    If Shape.Tags("INSTRUMENTA CROSS-SLIDE SHAPE") = CrossSlideShapeId Then
                        
                        Shape.Delete
                        
                    End If
                    
                Next
            Next
            
        Else
            MsgBox "This shape does Not have a tag."
        End If
        
    Else
        MsgBox "No shape selected."
    End If
    
End Sub
         */
        #endregion

        #region ObjectsMarginsToZero
        public void ObjectsMarginsToZero(IRibbonControl control)
        {
            var myDocument =ThisAddIn.application.ActiveWindow;
            if(!(myDocument.Selection.Type==PpSelectionType.ppSelectionShapes || myDocument.Selection.Type == PpSelectionType.ppSelectionText))
            {
                MessageBox.Show("No shapes selected.");
            }
            else
            {
                if(myDocument.Selection.ShapeRange.HasTextFrame==MsoTriState.msoCTrue || myDocument.Selection.ShapeRange.HasTextFrame == MsoTriState.msoTrue)
                {

                    myDocument.Selection.ShapeRange.TextFrame.MarginBottom = 0;
                    myDocument.Selection.ShapeRange.TextFrame.MarginLeft = 0       ;
                    myDocument.Selection.ShapeRange.TextFrame.MarginRight = 0      ;
                    myDocument.Selection.ShapeRange.TextFrame.MarginTop = 0        ;

                }
                else
                {
                    MessageBox.Show("No text capable shape selected.");
                }
            }
        }
        /*
         Sub ObjectsMarginsToZero()
    
    Set myDocument = Application.ActiveWindow
    
    If Not (myDocument.Selection.Type = ppSelectionShapes Or myDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No shapes selected."
    Else
    
    If myDocument.Selection.ShapeRange.HasTextFrame Then
    
        With myDocument.Selection.ShapeRange.TextFrame
        .MarginBottom = 0
        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = 0
        
    End With
    
    Else
    
    MsgBox "No text capable shape selected."
    
    End If
    
    End If
    
End Sub
         */
        #endregion


        #region ObjectsMarginsIncrease
        public void ObjectsMarginsIncrease(IRibbonControl control)
        {
            var myDocument = ThisAddIn.application.ActiveWindow;
            if (!(myDocument.Selection.Type == PpSelectionType.ppSelectionShapes || myDocument.Selection.Type == PpSelectionType.ppSelectionText))
            {
                MessageBox.Show("No shapes selected.");
            }
            else
            {
                if (myDocument.Selection.ShapeRange.HasTextFrame == MsoTriState.msoCTrue || myDocument.Selection.ShapeRange.HasTextFrame == MsoTriState.msoTrue)
                {

                    myDocument.Selection.ShapeRange.TextFrame.MarginBottom += (float)0.2;
                    myDocument.Selection.ShapeRange.TextFrame.MarginLeft   += (float)0.2;
                    myDocument.Selection.ShapeRange.TextFrame.MarginRight  += (float)0.2;
                    myDocument.Selection.ShapeRange.TextFrame.MarginTop    += (float)0.2;

                }
                else
                {
                    MessageBox.Show("No text capable shape selected.");
                }
            }
        }

        /*
         Sub ObjectsMarginsIncrease()
    
    Set myDocument = Application.ActiveWindow
    
    If Not (myDocument.Selection.Type = ppSelectionShapes Or myDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No shapes selected."
    Else
    
    If myDocument.Selection.ShapeRange.HasTextFrame Then
    
    With myDocument.Selection.ShapeRange.TextFrame
         .MarginBottom = .MarginBottom + 0.2
        .MarginLeft = .MarginLeft + 0.2
        .MarginRight = .MarginRight + 0.2
        .MarginTop = .MarginTop + 0.2
        
    End With
    
    Else
    
    MsgBox "No text capable shape selected."
    
    End If
    
    End If
End Sub

         */
        #endregion

        #region ObjectsMarginsDecrease
        public void ObjectsMarginsDecrease(IRibbonControl control)
        {
            var myDocument = ThisAddIn.application.ActiveWindow;
            if (!(myDocument.Selection.Type == PpSelectionType.ppSelectionShapes || myDocument.Selection.Type == PpSelectionType.ppSelectionText))
            {
                MessageBox.Show("No shapes selected.");
            }
            else
            {
                if (myDocument.Selection.ShapeRange.HasTextFrame == MsoTriState.msoCTrue || myDocument.Selection.ShapeRange.HasTextFrame == MsoTriState.msoTrue)
                {

                    if(myDocument.Selection.ShapeRange.TextFrame.MarginBottom>=0.2)
                        myDocument.Selection.ShapeRange.TextFrame.MarginBottom -= (float)0.2;
                    if(myDocument.Selection.ShapeRange.TextFrame.MarginLeft>=0.2)
                        myDocument.Selection.ShapeRange.TextFrame.MarginLeft -= (float)0.2;
                    if(myDocument.Selection.ShapeRange.TextFrame.MarginRight>=0.2)
                       myDocument.Selection.ShapeRange.TextFrame.MarginRight -= (float)0.2;
                    if(myDocument.Selection.ShapeRange.TextFrame.MarginTop>=0.2)
                       myDocument.Selection.ShapeRange.TextFrame.MarginTop -= (float)0.2;

                }
                else
                {
                    MessageBox.Show("No text capable shape selected.");
                }
            }
        }
        /*
         Sub ObjectsMarginsDecrease()
    
    Set myDocument = Application.ActiveWindow
    
    If Not (myDocument.Selection.Type = ppSelectionShapes Or myDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No shapes selected."
    Else
    
    If myDocument.Selection.ShapeRange.HasTextFrame Then
    
    With myDocument.Selection.ShapeRange.TextFrame
        If .MarginBottom >= 0.2 Then
            .MarginBottom = .MarginBottom - 0.2
        End If
        If .MarginLeft >= 0.2 Then
            .MarginLeft = .MarginLeft - 0.2
        End If
        If .MarginRight >= 0.2 Then
            .MarginRight = .MarginRight - 0.2
        End If
        If .MarginTop >= 0.2 Then
            .MarginTop = .MarginTop - 0.2
        End If
        
    End With
    
    Else
    
    MsgBox "No text capable shape selected."
    
    End If
    
    End If
End Sub

         */
        #endregion


        #region test
        public void doTest(IRibbonControl control)
        {
            MessageBox.Show("testtttttttt");
            ///var presentation=ThisAddIn.application.ActivePresentation;

            // IAutoShape shape = (IAutoShape)presentation.Slides[0].Shapes.AppendShape(ShapeType.DoubleWave, new RectangleF(100, 100, 400, 200));

            Microsoft.Office.Interop.PowerPoint.Shapes shapes = ThisAddIn.application.ActiveWindow.Selection.SlideRange[1].Shapes;

            Shape shape = shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 50, 50, 100, 100);
            string picUrl = @"C:\Users\Administrator\Desktop\image.jpg";
            
//shape.Fill. = FillFormatType.Picture;
            
//shape.Fill.PictureFill.Picture.Url = picUrl;

        }
        public Bitmap getTestImage(IRibbonControl control)
        {
            return Resource1.test ;
        }

        #endregion

        #region TableDistributeRowsWithGaps
        public void TableDistributeRowsWithGaps(IRibbonControl control)
        {
            var myDocument = ThisAddIn.application.ActiveWindow;
            double TotalHeight;
            double NumberOfRowsToDistribute;
            TotalHeight = 0;
            NumberOfRowsToDistribute = 0;

            if (!(myDocument.Selection.Type == PpSelectionType.ppSelectionShapes || myDocument.Selection.Type == PpSelectionType.ppSelectionText))
            {
                MessageBox.Show("No table or cells selected.");
            }
            else
            {
                if(ThisAddIn.application.ActiveWindow.Selection.ShapeRange.HasTable==MsoTriState.msoCTrue || ThisAddIn.application.ActiveWindow.Selection.ShapeRange.HasTable == MsoTriState.msoTrue)
                {
                   var TypeOfGaps = ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Tags["INSTRUMENTA ROWGAPS"];
                    for( int RowsCount = 1;RowsCount<= ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Rows.Count; RowsCount++)
                    {
                        for(int ColsCount = 1; ColsCount <=ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Columns.Count; ColsCount++)
                        {
                            if (ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Selected) 
                            {
                                if (!((RowsCount % 2 == 0 && TypeOfGaps == "even") ||( RowsCount % 2 != 0 && TypeOfGaps == "odd")))
                                {
                                    TotalHeight = TotalHeight + ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Rows[RowsCount].Height;
                                    NumberOfRowsToDistribute = NumberOfRowsToDistribute + 1;
                                    break;
                                }
                            }
                        }
                    }
                    if (NumberOfRowsToDistribute > 0)
                    {
                        for (int RowsCount = 1; RowsCount <= ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Rows.Count; RowsCount++)
                        {
                            for (int ColsCount = 1; ColsCount <= ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Columns.Count; ColsCount++)
                            {



                                if (ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Selected) 
                                {

                                    if (!((RowsCount % 2 == 0 && TypeOfGaps == "even") || (RowsCount % 2 != 0 && TypeOfGaps == "odd"))) 
                                    {

                                        ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Rows[RowsCount].Height = (float)(TotalHeight / NumberOfRowsToDistribute);
                                        break;

                                     }


                                                }
 

                            }
                        }
                     }




                        }
                else
                {
                    MessageBox.Show("No table or cells selected.");
                }
            }

        }
        /*
         Sub TableDistributeRowsWithGaps()

    Set myDocument = Application.ActiveWindow
    Dim TotalHeight As Double
    Dim NumberOfRowsToDistribute As Long
    TotalHeight = 0
    NumberOfRowsToDistribute = 0

    If Not (myDocument.Selection.Type = ppSelectionShapes Or myDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No table or cells selected."
    Else


    If Application.ActiveWindow.Selection.ShapeRange.HasTable Then

    With Application.ActiveWindow.Selection.ShapeRange.Table

        TypeOfGaps = Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS")

        For RowsCount = 1 To .Rows.Count

            For ColsCount = 1 To .Columns.Count

                If .Cell(RowsCount, ColsCount).Selected Then

                If Not ((RowsCount Mod 2 = 0 And TypeOfGaps = "even") Or (Not RowsCount Mod 2 = 0 And TypeOfGaps = "odd")) Then

                TotalHeight = TotalHeight + .Rows(RowsCount).Height
                NumberOfRowsToDistribute = NumberOfRowsToDistribute + 1
                Exit For

                End If


                End If

            Next ColsCount
        Next RowsCount


        If NumberOfRowsToDistribute > 0 Then

        For RowsCount = 1 To .Rows.Count

            For ColsCount = 1 To .Columns.Count

                If .Cell(RowsCount, ColsCount).Selected Then

                If Not ((RowsCount Mod 2 = 0 And TypeOfGaps = "even") Or (Not RowsCount Mod 2 = 0 And TypeOfGaps = "odd")) Then

                .Rows(RowsCount).Height = TotalHeight / NumberOfRowsToDistribute
                Exit For

                End If


                End If

            Next ColsCount
        Next RowsCount
        End If

    End With

    Else

    MsgBox "No table or cells selected."

    End If

    End If

End Sub

         */
        #endregion

        #region TableDistributeColumnsWithGaps
        public void TableDistributeColumnsWithGaps(IRibbonControl control)
        {
            var myDocument = ThisAddIn.application.ActiveWindow;
            double TotalWidth=0;
            float NumberOfColumnsToDistribute = 0;
            if (!(myDocument.Selection.Type == PpSelectionType.ppSelectionShapes || myDocument.Selection.Type == PpSelectionType.ppSelectionText))
            {
                MessageBox.Show("No table or cells selected.");
            }
            else
            {
                if (ThisAddIn.application.ActiveWindow.Selection.ShapeRange.HasTable == MsoTriState.msoCTrue || ThisAddIn.application.ActiveWindow.Selection.ShapeRange.HasTable == MsoTriState.msoTrue)
                {
                    var TypeOfGaps = ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Tags["INSTRUMENTA COLUMNGAPS"];

                    for (int ColsCount = 1; ColsCount <= ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Columns.Count; ColsCount++)
                    
                    {
                        for (int RowsCount = 1; RowsCount <= ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Rows.Count; RowsCount++)
                        {
                            if (ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Selected)
                            {
                                if (!((ColsCount % 2 == 0 && TypeOfGaps == "even") || (ColsCount % 2 != 0 && TypeOfGaps == "odd")))
                                {
                                    TotalWidth = TotalWidth + ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Columns[ColsCount].Width ;
                                    NumberOfColumnsToDistribute = NumberOfColumnsToDistribute + 1;
                                    break;
                                }

                            
                            }
                        }
                    }
                    if (NumberOfColumnsToDistribute > 0)
                    {
                        for (int ColsCount = 1; ColsCount <= ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Columns.Count; ColsCount++)
                        {
                            for (int RowsCount = 1; RowsCount <= ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Rows.Count; RowsCount++)
                            {


                                if (ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Selected)
                                {

                                    if (!((ColsCount % 2 == 0 && TypeOfGaps == "even") || (ColsCount % 2 != 0 && TypeOfGaps == "odd")))
                                    {

                                        ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Columns[ColsCount].Width  = (float)(TotalWidth / NumberOfColumnsToDistribute);
                                        break;

                                    }


                                }
                            }
                        }
                    }




                }
                else
                {
                    MessageBox.Show("No table or cells selected.");
                }
            }
        }

        /*
         Sub TableDistributeColumnsWithGaps()

    Set myDocument = Application.ActiveWindow
    Dim TotalWidth As Double
    Dim NumberOfColumnsToDistribute As Long
    TotalWidth = 0
    NumberOfColumnsToDistribute = 0
     
    If Not (myDocument.Selection.Type = ppSelectionShapes Or myDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No table or cells selected."
    Else
    
        
    If Application.ActiveWindow.Selection.ShapeRange.HasTable Then
        
    With Application.ActiveWindow.Selection.ShapeRange.Table
        
        TypeOfGaps = Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS")
        
        For ColsCount = 1 To .Columns.Count
        
            For RowsCount = 1 To .Rows.Count
                
                If .Cell(RowsCount, ColsCount).Selected Then
                    
                If Not ((ColsCount Mod 2 = 0 And TypeOfGaps = "even") Or (Not ColsCount Mod 2 = 0 And TypeOfGaps = "odd")) Then
                
                TotalWidth = TotalWidth + .Columns(ColsCount).Width
                NumberOfColumnsToDistribute = NumberOfColumnsToDistribute + 1
                Exit For
                    
                End If
                    
                
                End If
                
            Next RowsCount
        Next ColsCount
        
        
        If NumberOfColumnsToDistribute > 0 Then
        For ColsCount = 1 To .Columns.Count
        
            For RowsCount = 1 To .Rows.Count
                
                If .Cell(RowsCount, ColsCount).Selected Then
                    
                If Not ((ColsCount Mod 2 = 0 And TypeOfGaps = "even") Or (Not ColsCount Mod 2 = 0 And TypeOfGaps = "odd")) Then
                
                .Columns(ColsCount).Width = TotalWidth / NumberOfColumnsToDistribute
                Exit For
                    
                End If
                    
                
                End If
                
            Next RowsCount
        Next ColsCount
        End If
        
    End With
    
    Else
    
    MsgBox "No table or cells selected."
    
    End If
    
    End If

End Sub

         
         
         */
        #endregion
        #region TableQuickFormat
        public void TableQuickFormat(IRibbonControl control)
        {
            var myDocument = ThisAddIn.application.ActiveWindow;
            if (!(myDocument.Selection.Type == PpSelectionType.ppSelectionShapes || myDocument.Selection.Type == PpSelectionType.ppSelectionText))
            {
                MessageBox.Show("No table selected.");
            }
            else
            {
                if(ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Count==1 && (ThisAddIn.application.ActiveWindow.Selection.ShapeRange.HasTable==MsoTriState.msoCTrue || ThisAddIn.application.ActiveWindow.Selection.ShapeRange.HasTable == MsoTriState.msoTrue))
                {
                    TableRemoveBackgrounds(control);
                    TableRemoveBorders(control);
                    //With ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table

                    if (ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Tags["INSTRUMENTA COLUMNGAPS"] != "")
                        TableColumnRemoveGaps(control );



                    if (ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Tags["INSTRUMENTA ROWGAPS"] != "")
                        TableRowRemoveGaps(control);

                    for(int RowCount=1; RowCount<= ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Rows.Count; RowCount++)
                    {
                        for (int ColumnCount = 1; ColumnCount <= ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Columns.Count; ColumnCount++)
                        {
                            ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Cell(RowCount, ColumnCount).Shape.TextFrame.TextRange.Font.Color.RGB = Color.Black.ToArgb();//;
                        }
                    }

                    for(int CellCount=1; CellCount<= ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Rows[1].Cells.Count; CellCount++)
                    {
                        ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Rows[1].Cells[CellCount].Borders[PpBorderType.ppBorderTop].Weight = 0;
                        ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Rows[1].Cells[CellCount].Borders[PpBorderType.ppBorderBottom ].Weight = 2;
                        ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Rows[1].Cells[CellCount].Borders[PpBorderType.ppBorderBottom].ForeColor.RGB = Color.Black.ToArgb();//;
                        ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Rows[1].Cells[CellCount].Shape.Fill.Visible = MsoTriState.msoFalse;
                        ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Rows[1].Cells[CellCount].Shape.TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorBottom;
                        ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Rows[1].Cells[CellCount].Shape.TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
                        ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Rows[1].Cells[CellCount].Shape.TextFrame.TextRange.Font.Color.RGB = Color.Black.ToArgb();//;

                    }

                    TableColumnGaps ("even", 20);
                    /*
                     */


                }
                else
                {
                    MessageBox.Show("No table selected or too many shapes selected. Select one table.");
                }
            }
        }
        /*
         Sub TableQuickFormat()

Set myDocument = Application.ActiveWindow

If Not (myDocument.Selection.Type = ppSelectionShapes Or myDocument.Selection.Type = ppSelectionText) Then
    MsgBox "No table selected."
Else

    If (Application.ActiveWindow.Selection.ShapeRange.Count = 1) And Application.ActiveWindow.Selection.ShapeRange.HasTable Then

        With Application.ActiveWindow.Selection.ShapeRange.Table

            TableRemoveBackgrounds
            TableRemoveBorders

            If Not Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS") = "" Then
                TableColumnRemoveGaps
            End If

            If Not Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS") = "" Then
                TableRowRemoveGaps
            End If

            For RowCount = 1 To .Rows.Count

                For ColumnCount = 1 To .Columns.Count

                    .Cell(RowCount, ColumnCount).Shape.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)

                Next

            Next

            For CellCount = 1 To .Rows(1).Cells.Count

                .Rows(1).Cells(CellCount).Borders(ppBorderTop).Weight = 0
                .Rows(1).Cells(CellCount).Borders(ppBorderBottom).Weight = 2
                .Rows(1).Cells(CellCount).Borders(ppBorderBottom).ForeColor.RGB = RGB(0, 0, 0)
                .Rows(1).Cells(CellCount).Shape.Fill.Visible = msoFalse
                .Rows(1).Cells(CellCount).Shape.TextFrame.VerticalAnchor = msoAnchorBottom
                .Rows(1).Cells(CellCount).Shape.TextFrame.TextRange.Font.Bold = msoTrue
                .Rows(1).Cells(CellCount).Shape.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)

            Next CellCount

            TableColumnGaps "even", 20

        End With

    Else

        MsgBox "No table selected or too many shapes selected. Select one table."

    End If

End If

End Sub

         */

        #endregion

        #region TableRemoveBackgrounds
        public void TableRemoveBackgrounds(IRibbonControl control)
        {
            var myDocument = ThisAddIn.application.ActiveWindow;
            if (!(myDocument.Selection.Type == PpSelectionType.ppSelectionShapes || myDocument.Selection.Type == PpSelectionType.ppSelectionText))
            {
                MessageBox.Show("No table selected.");
            }
            else
            {
                if (ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Count == 1 && (ThisAddIn.application.ActiveWindow.Selection.ShapeRange.HasTable == MsoTriState.msoCTrue || ThisAddIn.application.ActiveWindow.Selection.ShapeRange.HasTable == MsoTriState.msoTrue))
                {
                    //ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table
                    ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.HorizBanding = false;
                    ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.VertBanding = false;


                    ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Fill.Solid();
                     ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Fill.ForeColor.RGB = Color.White.ToArgb();// RGB(255, 255, 255)
                     ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Fill.Visible = MsoTriState.msoFalse;

                    ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Background.Fill.Solid();
                    ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Background.Fill.ForeColor.RGB = Color.White.ToArgb();//RGB(255, 255, 255)
                    ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Background.Fill.Visible = MsoTriState.msoFalse;

                    ProgressForm pf = new ProgressForm();
                    pf.Show();
                    for(int RowCount=1; RowCount<= ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Rows.Count; RowCount++)
                    {
                        pf.SetProgress((RowCount * 100) / ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Rows.Count);
                        for (int ColumnCount = 1; ColumnCount <= ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Columns.Count; ColumnCount++)
                        {
                            ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Cell(RowCount, ColumnCount).Shape.Fill.Solid();
                            ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Cell(RowCount, ColumnCount).Shape.Fill.ForeColor.RGB = Color.White.ToArgb();
                            ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Cell(RowCount, ColumnCount).Shape.Fill.Visible = MsoTriState.msoFalse;
                        }
                        pf.Close();
                    }
                
                }
                else
                {
                    MessageBox.Show("No table selected or too many shapes selected. Select one table.");
                }
            }
        }
        /*
         Sub TableRemoveBackgrounds()
    
    Set myDocument = Application.ActiveWindow
    
    If Not (myDocument.Selection.Type = ppSelectionShapes Or myDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No table selected."
    Else
        
        If (Application.ActiveWindow.Selection.ShapeRange.Count = 1) And Application.ActiveWindow.Selection.ShapeRange.HasTable Then
            
            With Application.ActiveWindow.Selection.ShapeRange.Table

                .HorizBanding = False
                .VertBanding = False
                
                Application.ActiveWindow.Selection.ShapeRange.Fill.Solid
                Application.ActiveWindow.Selection.ShapeRange.Fill.ForeColor.RGB = RGB(255, 255, 255)
                Application.ActiveWindow.Selection.ShapeRange.Fill.Visible = msoFalse
                
                .Background.Fill.Solid
                .Background.Fill.ForeColor.RGB = RGB(255, 255, 255)
                .Background.Fill.Visible = msoFalse
                
                ProgressForm.Show
                
                For RowCount = 1 To .Rows.Count
                    
                SetProgress (RowCount / .Rows.Count * 100)
                    
                    For ColumnCount = 1 To .Columns.Count
                        
                        .Cell(RowCount, ColumnCount).Shape.Fill.Solid
                        .Cell(RowCount, ColumnCount).Shape.Fill.ForeColor.RGB = RGB(255, 255, 255)
                        .Cell(RowCount, ColumnCount).Shape.Fill.Visible = msoFalse
                    Next
                    
                Next
                
                ProgressForm.Hide
                
            End With
            
        Else
            
            MsgBox "No table selected or too many shapes selected. Select one table."
            
        End If
        
    End If
    
End Sub

         */
        #endregion
        #region TableRemoveBorders
        public void TableRemoveBorders(IRibbonControl control)
        {
            var myDocument = ThisAddIn.application.ActiveWindow;
            if (!(myDocument.Selection.Type == PpSelectionType.ppSelectionShapes || myDocument.Selection.Type == PpSelectionType.ppSelectionText))
            {
                MessageBox.Show("No table selected.");
            }
            else
            {
                if (ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Count == 1 && (ThisAddIn.application.ActiveWindow.Selection.ShapeRange.HasTable == MsoTriState.msoCTrue || ThisAddIn.application.ActiveWindow.Selection.ShapeRange.HasTable == MsoTriState.msoTrue))
                {
                    //ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table
                    ProgressForm pf = new ProgressForm();
                    pf.Show();
                    for (int RowCount = 1; RowCount <= ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Rows.Count; RowCount++)
                    {
                        pf.SetProgress((RowCount * 100) / ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Rows.Count);
                        for (int ColumnCount = 1; ColumnCount <= ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Columns.Count; ColumnCount++)
                        {
                             ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Cell(RowCount, ColumnCount).Borders[PpBorderType.ppBorderLeft].ForeColor.RGB = Color.White.ToArgb();
                             ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Cell(RowCount, ColumnCount).Borders[PpBorderType.ppBorderRight].ForeColor.RGB =Color.White.ToArgb();
                             ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Cell(RowCount, ColumnCount).Borders[PpBorderType.ppBorderTop].ForeColor.RGB = Color.White.ToArgb();
                             ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Cell(RowCount, ColumnCount).Borders[PpBorderType.ppBorderBottom].ForeColor.RGB= Color.White.ToArgb();
                             
                             ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Cell(RowCount, ColumnCount).Borders[PpBorderType.ppBorderLeft].Weight = 0;
                             ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Cell(RowCount, ColumnCount).Borders[PpBorderType.ppBorderRight].Weight = 0 ;
                             ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Cell(RowCount, ColumnCount).Borders[PpBorderType.ppBorderTop].Weight = 0   ;
                             ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Cell(RowCount, ColumnCount).Borders[PpBorderType.ppBorderBottom].Weight = 0;
                             
                             ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Cell(RowCount, ColumnCount).Borders[PpBorderType.ppBorderLeft].Visible = MsoTriState.msoFalse;
                            ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Cell(RowCount, ColumnCount).Borders[PpBorderType.ppBorderRight].Visible = MsoTriState.msoFalse;
                             ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Cell(RowCount, ColumnCount).Borders[PpBorderType.ppBorderTop].Visible =   MsoTriState.msoFalse ;
                             ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Cell(RowCount, ColumnCount).Borders[PpBorderType.ppBorderBottom].Visible = MsoTriState.msoFalse;
                        }
                    }
                    pf.Close();
                           
                        }
                else
                {
                    MessageBox.Show("No table selected or too many shapes selected. Select one table.");
                }
            }
        }


        /*
         Sub TableRemoveBorders()
    
    Set myDocument = Application.ActiveWindow
    
    If Not (myDocument.Selection.Type = ppSelectionShapes Or myDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No table selected."
    Else
        
        If (Application.ActiveWindow.Selection.ShapeRange.Count = 1) And Application.ActiveWindow.Selection.ShapeRange.HasTable Then
            
            With Application.ActiveWindow.Selection.ShapeRange.Table
            
                ProgressForm.Show
                
                For RowCount = 1 To .Rows.Count
                    
                SetProgress (RowCount / .Rows.Count * 100)
                    
                    For ColumnCount = 1 To .Columns.Count
                        
                        .Cell(RowCount, ColumnCount).Borders(ppBorderLeft).ForeColor.RGB = RGB(255, 255, 255)
                        .Cell(RowCount, ColumnCount).Borders(ppBorderRight).ForeColor.RGB = RGB(255, 255, 255)
                        .Cell(RowCount, ColumnCount).Borders(ppBorderTop).ForeColor.RGB = RGB(255, 255, 255)
                        .Cell(RowCount, ColumnCount).Borders(ppBorderBottom).ForeColor.RGB = RGB(255, 255, 255)
                        
                        .Cell(RowCount, ColumnCount).Borders(ppBorderLeft).Weight = 0
                        .Cell(RowCount, ColumnCount).Borders(ppBorderRight).Weight = 0
                        .Cell(RowCount, ColumnCount).Borders(ppBorderTop).Weight = 0
                        .Cell(RowCount, ColumnCount).Borders(ppBorderBottom).Weight = 0
                        
                        .Cell(RowCount, ColumnCount).Borders(ppBorderLeft).Visible = msoFalse
                        .Cell(RowCount, ColumnCount).Borders(ppBorderRight).Visible = msoFalse
                        .Cell(RowCount, ColumnCount).Borders(ppBorderTop).Visible = msoFalse
                        .Cell(RowCount, ColumnCount).Borders(ppBorderBottom).Visible = msoFalse
                    Next
                    
                Next
                
                ProgressForm.Hide
                
            End With
            
        Else
            
            MsgBox "No table selected or too many shapes selected. Select one table."
            
        End If
        
    End If
    
End Sub

         */
        #endregion region

        #region ConvertTableToShapes
        public void ConvertTableToShapes(IRibbonControl control)
        {
            var myDocument = ThisAddIn.application.ActiveWindow;
            if(myDocument.Selection.Type != PpSelectionType.ppSelectionShapes)
            {
                MessageBox.Show("Please select a table.");
            }

            else if (myDocument.Selection.ShapeRange.HasTable==MsoTriState.msoCTrue|| myDocument.Selection.ShapeRange.HasTable == MsoTriState.msoTrue)
            {
                float TableTop = myDocument.Selection.ShapeRange.Top;
                float TableLeft = myDocument.Selection.ShapeRange.Left;


              var  TypeOfColumnGaps = ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Tags["INSTRUMENTA COLUMNGAPS"];
              var TypeOfRowGaps = ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Tags["INSTRUMENTA ROWGAPS"];

                ProgressForm pf = new ProgressForm();
                pf.Show();
                for (int RowsCount = 1; RowsCount <= ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Rows.Count; RowsCount++)
                {
                    pf.SetProgress((RowsCount * 100) / ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Rows.Count);
                    for (int ColsCount = 1; ColsCount <= ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Columns.Count; ColsCount++)
                    {
                        if (!((ColsCount%2==0 && TypeOfColumnGaps == "even")||(ColsCount%2!=0 && TypeOfColumnGaps=="odd")||(RowsCount%2==0 && TypeOfRowGaps=="even")||(RowsCount%2!=0 && TypeOfRowGaps=="odd")))
                        {
                            var NewShape = myDocument.Selection.SlideRange.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, TableLeft, TableTop, myDocument.Selection.ShapeRange.Table.Columns[ColsCount].Width, myDocument.Selection.ShapeRange.Table.Rows[RowsCount].Height);
                                        NewShape.TextFrame.MarginBottom = myDocument.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame.MarginBottom;
                                        NewShape.TextFrame.MarginLeft = myDocument.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame.MarginLeft    ;
                                        NewShape.TextFrame.MarginRight = myDocument.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame.MarginRight  ;
                                        NewShape.TextFrame.MarginTop = myDocument.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame.MarginTop;
                            if (myDocument.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame.TextRange.Text != "")
                            {
                                myDocument.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame.TextRange.Cut();
                                NewShape.TextFrame.TextRange.Paste();
                            }
                            NewShape.TextFrame.TextRange.ParagraphFormat.Alignment = myDocument.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame.TextRange.ParagraphFormat.Alignment;
                            NewShape.TextFrame.TextRange.ParagraphFormat.BaseLineAlignment = myDocument.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame.TextRange.ParagraphFormat.BaseLineAlignment;
                            NewShape.Fill.ForeColor.RGB = myDocument.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.Fill.ForeColor.RGB;
                            NewShape.Line.ForeColor.RGB = myDocument.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Borders[PpBorderType.ppBorderBottom].ForeColor.RGB;
                        }
                        TableLeft = TableLeft + ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Columns[ColsCount].Width;
                        
                    }

                    TableLeft = ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Left;
                    TableTop = TableTop + ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Rows[RowsCount].Height;

                }
                pf.Close();
                ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Delete();

            }
            else
            {
                MessageBox.Show("No table selected.");
            }
        }

        /*
         Sub ConvertTableToShapes()
    
    Set myDocument = Application.ActiveWindow
            
    If Not myDocument.Selection.Type = ppSelectionShapes Then
    MsgBox "Please select a table."
    
    ElseIf myDocument.Selection.ShapeRange.HasTable Then
    
    TableTop = myDocument.Selection.ShapeRange.Top
    TableLeft = myDocument.Selection.ShapeRange.Left
    
    TypeOfColumnGaps = Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS")
    TypeOfRowGaps = Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS")
    
    ProgressForm.Show
    
    For RowsCount = 1 To myDocument.Selection.ShapeRange.Table.Rows.Count
    
    SetProgress (RowsCount / myDocument.Selection.ShapeRange.Table.Rows.Count * 100)
    
        For ColsCount = 1 To myDocument.Selection.ShapeRange.Table.Columns.Count
            
            If Not ((ColsCount Mod 2 = 0 And TypeOfColumnGaps = "even") Or (Not ColsCount Mod 2 = 0 And TypeOfColumnGaps = "odd") Or (RowsCount Mod 2 = 0 And TypeOfRowGaps = "even") Or (Not RowsCount Mod 2 = 0 And TypeOfRowGaps = "odd")) Then
            
            Set NewShape = myDocument.Selection.SlideRange.Shapes.AddShape(Type:=msoShapeRectangle, Left:=TableLeft, Top:=TableTop, Width:=myDocument.Selection.ShapeRange.Table.Columns(ColsCount).Width, Height:=myDocument.Selection.ShapeRange.Table.Rows(RowsCount).Height)
            
            With NewShape
                .TextFrame.MarginBottom = myDocument.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame.MarginBottom
                .TextFrame.MarginLeft = myDocument.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame.MarginLeft
                .TextFrame.MarginRight = myDocument.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame.MarginRight
                .TextFrame.MarginTop = myDocument.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame.MarginTop
                
                If Not myDocument.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame.TextRange.Text = "" Then
                    myDocument.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame.TextRange.Cut
                    .TextFrame.TextRange.Paste
                End If
                
                .TextFrame.TextRange.ParagraphFormat.Alignment = myDocument.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame.TextRange.ParagraphFormat.Alignment
                .TextFrame.TextRange.ParagraphFormat.BaseLineAlignment = myDocument.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame.TextRange.ParagraphFormat.BaseLineAlignment
                .Fill.ForeColor.RGB = myDocument.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.Fill.ForeColor.RGB
                .Line.ForeColor.RGB = myDocument.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Borders(ppBorderBottom).ForeColor.RGB
            End With
            
            End If
            
            TableLeft = TableLeft + Application.ActiveWindow.Selection.ShapeRange.Table.Columns(ColsCount).Width
            
        Next ColsCount
        
        
        TableLeft = Application.ActiveWindow.Selection.ShapeRange.Left
        TableTop = TableTop + Application.ActiveWindow.Selection.ShapeRange.Table.Rows(RowsCount).Height
        
    Next RowsCount
    
    ProgressForm.Hide
    
    Application.ActiveWindow.Selection.ShapeRange.Delete
    
    Else
    
    MsgBox "No table selected."
    
    End If
       
End Sub

         */
        #endregion

        #region TableTranspose
        public void TableTranspose(IRibbonControl control)
        {
            var myDocument = ThisAddIn.application.ActiveWindow;
            if(!(myDocument.Selection.Type== PpSelectionType.ppSelectionShapes || myDocument.Selection.Type == PpSelectionType.ppSelectionText))
            {
                MessageBox.Show("No table selected.");
            }
            else
            {
                if(ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Count == 1 && (ThisAddIn.application.ActiveWindow.Selection.ShapeRange.HasTable==MsoTriState.msoCTrue|| ThisAddIn.application.ActiveWindow.Selection.ShapeRange.HasTable == MsoTriState.msoTrue))
                {

                    // ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table

                    var CopyTable = ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Duplicate();
                    for(int RowsCount= ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Rows.Count; RowsCount >= 2; RowsCount--)
                    {
                        CopyTable.Table.Rows[RowsCount].Delete();
                    }
                    for (int ColsCount = ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Columns.Count; ColsCount >= 2; ColsCount--)
                    {
                        CopyTable.Table.Columns[ColsCount].Delete();
                    }
                    for (int RowsCount = ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Rows.Count; RowsCount >= 2; RowsCount--)
                    {
                        CopyTable.Table.Columns.Add();
                    }
                    for (int ColsCount = ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Columns.Count; ColsCount >= 2; ColsCount--)
                    {
                        CopyTable.Table.Rows.Add();
                    }
                    CopyTable.Width =ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Width;
                    CopyTable.Top =  ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Top;
                    CopyTable.Left = ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Left;

                    for(int RowsCount=1; RowsCount<= ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Rows.Count; RowsCount++)
                    {
                        for (int ColsCount = 1; ColsCount <= ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Columns.Count; ColsCount++)
                        {
                            ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame2.TextRange.Cut();
                            CopyTable.Table.Cell(ColsCount, RowsCount).Shape.TextFrame2.TextRange.Paste();
                        }
                    }
                    ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Delete();
                    CopyTable.Select();
                }
                else
                {
                    MessageBox.Show("No table selected or too many shapes selected. Select one table.");
                }
            }
        }
        /*
         Sub TableTranspose()
    
    Set myDocument = Application.ActiveWindow
    
    
    If Not (myDocument.Selection.Type = ppSelectionShapes Or myDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No table selected."
    Else
        
        If (Application.ActiveWindow.Selection.ShapeRange.Count = 1) And Application.ActiveWindow.Selection.ShapeRange.HasTable Then
            
            With Application.ActiveWindow.Selection.ShapeRange.Table
                
                Set CopyTable = Application.ActiveWindow.Selection.ShapeRange.Duplicate
                
                For RowsCount = .Rows.Count To 2 Step -1
                    CopyTable.Table.Rows(RowsCount).Delete
                Next RowsCount
                
                For ColsCount = .Columns.Count To 2 Step -1
                    CopyTable.Table.Columns(ColsCount).Delete
                Next ColsCount
                
                For RowsCount = .Rows.Count To 2 Step -1
                    CopyTable.Table.Columns.Add
                Next RowsCount
                
                For ColsCount = .Columns.Count To 2 Step -1
                    CopyTable.Table.Rows.Add
                Next ColsCount
                
                CopyTable.Width = Application.ActiveWindow.Selection.ShapeRange.Width
                CopyTable.Top = Application.ActiveWindow.Selection.ShapeRange.Top
                CopyTable.Left = Application.ActiveWindow.Selection.ShapeRange.Left
                
                For RowsCount = 1 To .Rows.Count
                    For ColsCount = 1 To .Columns.Count
                        
                        .Cell(RowsCount, ColsCount).Shape.TextFrame2.TextRange.Cut
                        CopyTable.Table.Cell(ColsCount, RowsCount).Shape.TextFrame2.TextRange.Paste
                        
                    Next ColsCount
                Next RowsCount
                
            End With
            
            Application.ActiveWindow.Selection.ShapeRange.Delete
            CopyTable.Select
            
        Else
            
            MsgBox "No table selected or too many shapes selected. Select one table."
            
        End If
        
    End If
    
End Sub

         */
        #endregion
        #region SplitTableByRow
        public void SplitTableByRow(IRibbonControl control)
        {
            var myDocument = ThisAddIn.application.ActiveWindow;
            if (!(myDocument.Selection.Type == PpSelectionType.ppSelectionShapes || myDocument.Selection.Type == PpSelectionType.ppSelectionText))
            {
                MessageBox.Show("No table or cells selected.");
            }
            else
            {
                if (ThisAddIn.application.ActiveWindow.Selection.ShapeRange[1].HasTable==MsoTriState.msoCTrue || ThisAddIn.application.ActiveWindow.Selection.ShapeRange[1].HasTable == MsoTriState.msoTrue)
                {
                    //Application.ActiveWindow.Selection.ShapeRange(1).Table
                    for (int RowsCount = 1; RowsCount <= ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Rows.Count; RowsCount++)
                    {
                        for (int ColsCount = 1; ColsCount <= ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Columns.Count; ColsCount++)
                        {
                            if (ThisAddIn.application.ActiveWindow.Selection.ShapeRange[1].Table.Cell(RowsCount, ColsCount).Selected)
                            {
                                if (RowsCount != 1)
                                {
                                    var ThisTable = ThisAddIn.application.ActiveWindow.Selection.ShapeRange[1];
                                    var DuplicatedTable = ThisTable.Duplicate();
                                    DuplicatedTable.Left = ThisTable.Left;
                                    DuplicatedTable.Top = ThisTable.Top;

                                    DuplicatedTable.Table.FirstRow = false;

                                    for (int DeleteRows = 1; DeleteRows <= RowsCount - 1; DeleteRows++) {
                                        float DuplicatedHeight = DuplicatedTable.Table.Rows[1].Height;
                                        DuplicatedTable.Table.Rows[1].Delete();
                                        DuplicatedTable.Top = DuplicatedTable.Top + DuplicatedHeight;

                                    }

                                    DuplicatedTable.Top = DuplicatedTable.Top + 5;

                                    for (int DeleteRows = ThisAddIn.application.ActiveWindow.Selection.ShapeRange[1].Table.Rows.Count; DeleteRows >= RowsCount; DeleteRows--) {
                                        ThisTable.Table.Rows[DeleteRows].Delete();
                                    }

                                    return;

                                }
                                else
                                {
                                    MessageBox.Show("Will not work on the first row.");
                                }
                            }

                        }
                    }

                        }
                else
                {
                    MessageBox.Show("No table or cells selected.");
                }
            }
        }

        /*
         Sub SplitTableByRow()

    Set myDocument = Application.ActiveWindow

    If Not (myDocument.Selection.Type = ppSelectionShapes Or myDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No table or cells selected."
    Else

        If Application.ActiveWindow.Selection.ShapeRange(1).HasTable Then

            With Application.ActiveWindow.Selection.ShapeRange(1).Table

                For RowsCount = 1 To .Rows.Count
                    For ColsCount = 1 To .Columns.Count

                        If .Cell(RowsCount, ColsCount).Selected Then

                            If Not RowsCount = 1 Then

                                Set ThisTable = Application.ActiveWindow.Selection.ShapeRange(1)
                                Set DuplicatedTable = ThisTable.Duplicate
                                DuplicatedTable.Left = ThisTable.Left
                                DuplicatedTable.Top = ThisTable.Top

                                DuplicatedTable.Table.FirstRow = False

                                For DeleteRows = 1 To RowsCount - 1
                                    DuplicatedHeight = DuplicatedTable.Table.Rows(1).Height
                                    DuplicatedTable.Table.Rows(1).Delete
                                    DuplicatedTable.Top = DuplicatedTable.Top + DuplicatedHeight

                                Next

                                DuplicatedTable.Top = DuplicatedTable.Top + 5

                                For DeleteRows = .Rows.Count To RowsCount Step -1
                                    ThisTable.Table.Rows(DeleteRows).Delete
                                Next

                                Exit Sub

                            Else

                                MsgBox "Will not work on the first row."

                            End If

                        End If

                    Next ColsCount
                Next RowsCount

            End With

        Else

            MsgBox "No table or cells selected."

        End If

    End If

End Sub

         */
        #endregion
        #region SplitTableByColumn
        public void SplitTableByColumn(IRibbonControl control)
        {
            var myDocument = ThisAddIn.application.ActiveWindow;
            if (!(myDocument.Selection.Type == PpSelectionType.ppSelectionShapes || myDocument.Selection.Type == PpSelectionType.ppSelectionText))
            {
                MessageBox.Show("No table or cells selected.");
            }
            else
            {
                if (ThisAddIn.application.ActiveWindow.Selection.ShapeRange[1].HasTable == MsoTriState.msoCTrue || ThisAddIn.application.ActiveWindow.Selection.ShapeRange[1].HasTable == MsoTriState.msoTrue)
                {
                    //Application.ActiveWindow.Selection.ShapeRange(1).Table
                    for (int RowsCount = 1; RowsCount <= ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Rows.Count; RowsCount++)
                    {
                        for (int ColsCount = 1; ColsCount <= ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Columns.Count; ColsCount++)
                        {
                            if (ThisAddIn.application.ActiveWindow.Selection.ShapeRange[1].Table.Cell(RowsCount, ColsCount).Selected)
                            {
                                if (ColsCount != 1)
                                {
                                    var ThisTable = ThisAddIn.application.ActiveWindow.Selection.ShapeRange[1];
                                    var DuplicatedTable = ThisTable.Duplicate();
                                    DuplicatedTable.Left = ThisTable.Left;
                                    DuplicatedTable.Top = ThisTable.Top;


                                    DuplicatedTable.Table.FirstCol = false;


                                    for (int DeleteColumns = 1; DeleteColumns <= ColsCount - 1; DeleteColumns++) 
                                    {
                                       float DuplicatedWidth = DuplicatedTable.Table.Columns[1].Width;
                                        DuplicatedTable.Table.Columns[1].Delete();
                                        DuplicatedTable.Left = DuplicatedTable.Left + DuplicatedWidth;
    
                                    }


                                    DuplicatedTable.Left = DuplicatedTable.Left + 5;


                                    for (int DeleteColumns = ThisAddIn.application.ActiveWindow.Selection.ShapeRange[1].Table.Columns.Count; DeleteColumns >= ColsCount; DeleteColumns--) 
                                    {
                                        ThisTable.Table.Columns[DeleteColumns].Delete();
                                    }

                                    return;

                                }
                                else
                                {
                                    MessageBox.Show("Will not work on the first column.");
                                }
                            }

                        }
                    }

                }
                else
                {
                    MessageBox.Show("No table or cells selected.");
                }
            }
        }

        /*
         Sub SplitTableByColumn()
    
    Set myDocument = Application.ActiveWindow
    
    If Not (myDocument.Selection.Type = ppSelectionShapes Or myDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No table or cells selected."
    Else
        
        If Application.ActiveWindow.Selection.ShapeRange(1).HasTable Then
            
            With Application.ActiveWindow.Selection.ShapeRange(1).Table
                
                For RowsCount = 1 To .Rows.Count
                    For ColsCount = 1 To .Columns.Count
                        
                        If .Cell(RowsCount, ColsCount).Selected Then
                            
                            If Not ColsCount = 1 Then
                                
                                Set ThisTable = Application.ActiveWindow.Selection.ShapeRange(1)
                                Set DuplicatedTable = ThisTable.Duplicate
                                DuplicatedTable.Left = ThisTable.Left
                                DuplicatedTable.Top = ThisTable.Top
                                
                                DuplicatedTable.Table.FirstCol = False
                                
                                For DeleteColumns = 1 To ColsCount - 1
                                    DuplicatedWidth = DuplicatedTable.Table.Columns(1).Width
                                    DuplicatedTable.Table.Columns(1).Delete
                                    DuplicatedTable.Left = DuplicatedTable.Left + DuplicatedWidth
                                    
                                Next
                                
                                DuplicatedTable.Left = DuplicatedTable.Left + 5
                                
                                For DeleteColumns = .Columns.Count To ColsCount Step -1
                                    ThisTable.Table.Columns(DeleteColumns).Delete
                                Next
                                
                                Exit Sub
                                
                            Else
                                
                                MsgBox "Will not work on the first column."
                                
                            End If
                            
                        End If
                        
                    Next ColsCount
                Next RowsCount
                
            End With
            
        Else
            
            MsgBox "No table or cells selected."
            
        End If
        
    End If
    
End Sub

         */
        #endregion
        #region TableSum
        public void TableSum(IRibbonControl control)
        {
            var myDocument = ThisAddIn.application.ActiveWindow;
            double TotalSum = 0;
            if (!(myDocument.Selection.Type == PpSelectionType.ppSelectionShapes || myDocument.Selection.Type == PpSelectionType.ppSelectionText))
            {
                MessageBox.Show("No table or cells selected.");
            }
            else
            {
               if (ThisAddIn.application.ActiveWindow.Selection.ShapeRange.HasTable==MsoTriState.msoCTrue || ThisAddIn.application.ActiveWindow.Selection.ShapeRange.HasTable == MsoTriState.msoTrue)
                {
                    //With Application.ActiveWindow.Selection.ShapeRange.Table

                    var TypeOfGaps = ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Tags["INSTRUMENTA COLUMNGAPS"];
                    for (int RowsCount = 1; RowsCount <= ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Rows.Count; RowsCount++)
                    {
                        for (int ColsCount = 1; ColsCount <= ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Columns.Count; ColsCount++)
                        {
                            if (ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Selected)
                            {
                                if (!((ColsCount % 2 == 0 && TypeOfGaps == "even") ||( ColsCount % 2 != 0 && TypeOfGaps == "odd")))
                                {
                                    for(int SumCount=1; SumCount<= RowsCount-1; SumCount++)
                                    {
                                        try
                                        {
                                            TotalSum = TotalSum + double.Parse(ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Cell(SumCount, ColsCount).Shape.TextFrame.TextRange.Text);
                                        }
                                        catch(Exception ex)
                                        {

                                        }
                                    }
                                   ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame.TextRange.Text = TotalSum.ToString();
                                  
                                }
                                TotalSum = 0;
                            }
                        }
                    }

                }
                else
                {
                    MessageBox.Show("No table or cells selected.");
                }
            }
        }
        /*
         Sub TableSum()

    Set myDocument = Application.ActiveWindow
    Dim TotalSum As Double
    TotalSum = 0
    
       
    If Not (myDocument.Selection.Type = ppSelectionShapes Or myDocument.Selection.Type = ppSelectionText) Then
    MsgBox "No table or cells selected."
    Else
    
        
    If Application.ActiveWindow.Selection.ShapeRange.HasTable Then
        
    With Application.ActiveWindow.Selection.ShapeRange.Table
        
        TypeOfGaps = Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS")
        
        For RowsCount = 1 To .Rows.Count
            For ColsCount = 1 To .Columns.Count
                
                If .Cell(RowsCount, ColsCount).Selected Then
                    
                If Not ((ColsCount Mod 2 = 0 And TypeOfGaps = "even") Or (Not ColsCount Mod 2 = 0 And TypeOfGaps = "odd")) Then
                
                    For SumCount = 1 To RowsCount - 1
                    
                    On Error Resume Next
                    TotalSum = TotalSum + CDbl(.Cell(SumCount, ColsCount).Shape.TextFrame.TextRange.Text)
                    On Error GoTo 0
                    
                    Next SumCount
                        
                    .Cell(RowsCount, ColsCount).Shape.TextFrame.TextRange.Text = TotalSum
                    
                    End If
                    
                    TotalSum = 0
                
                End If
                
            Next ColsCount
        Next RowsCount
        
    End With
    
    Else
    
    MsgBox "No table or cells selected."
    
    End If
    
    End If

End Sub

         */
        #endregion

        #region TableRowSum
        public void TableRowSum(IRibbonControl control)
        {
            var myDocument = ThisAddIn.application.ActiveWindow;
            double TotalSum = 0;
            if (!(myDocument.Selection.Type == PpSelectionType.ppSelectionShapes || myDocument.Selection.Type == PpSelectionType.ppSelectionText))
            {
                MessageBox.Show("No table or cells selected.");
            }
            else
            {
                if (ThisAddIn.application.ActiveWindow.Selection.ShapeRange.HasTable == MsoTriState.msoCTrue || ThisAddIn.application.ActiveWindow.Selection.ShapeRange.HasTable == MsoTriState.msoTrue)
                {
                    //With Application.ActiveWindow.Selection.ShapeRange.Table

                    var TypeOfGaps = ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Tags["INSTRUMENTA COLUMNGAPS"];
                    for (int RowsCount = 1; RowsCount <= ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Rows.Count; RowsCount++)
                    {
                        for (int ColsCount = 1; ColsCount <= ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Columns.Count; ColsCount++)
                        {
                           
                            if (ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Selected)
                            {
                                if (!((RowsCount % 2 == 0 && TypeOfGaps == "even") || (RowsCount % 2 != 0 && TypeOfGaps == "odd")))
                                {
                                    for (int SumCount = 1; SumCount <= ColsCount - 1; SumCount++)
                                    {
                                        try
                                        {
                                            TotalSum = TotalSum + double.Parse(ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Cell(RowsCount,SumCount).Shape.TextFrame.TextRange.Text);
                                        }
                                        catch (Exception ex)
                                        {

                                        }
                                    }
                                    ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame.TextRange.Text = TotalSum.ToString();
                                }
                                TotalSum = 0;
                            }
                        }
                    }

                }
                else
                {
                    MessageBox.Show("No table or cells selected.");
                }
            }
        }

        /*
         Sub TableRowSum()

    Set myDocument = Application.ActiveWindow
    Dim TotalSum As Double
    TotalSum = 0
    
    If Not (myDocument.Selection.Type = ppSelectionShapes Or myDocument.Selection.Type = ppSelectionText) Then
    MsgBox "No table or cells selected."
    Else
    
    
    If Application.ActiveWindow.Selection.ShapeRange.HasTable Then
        
    TypeOfGaps = Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS")
        
    With Application.ActiveWindow.Selection.ShapeRange.Table
        
        For RowsCount = 1 To .Rows.Count
            For ColsCount = 1 To .Columns.Count
                
                If .Cell(RowsCount, ColsCount).Selected Then
                
                If Not ((RowsCount Mod 2 = 0 And TypeOfGaps = "even") Or (Not RowsCount Mod 2 = 0 And TypeOfGaps = "odd")) Then
                
                    For SumCount = 1 To ColsCount - 1
                    
                    On Error Resume Next
                    TotalSum = TotalSum + CDbl(.Cell(RowsCount, SumCount).Shape.TextFrame.TextRange.Text)
                    On Error GoTo 0
                    
                    Next SumCount
                        
                    .Cell(RowsCount, ColsCount).Shape.TextFrame.TextRange.Text = TotalSum
                    
                    End If
                    
                    TotalSum = 0
                
                End If
                
            Next ColsCount
        Next RowsCount
        
    End With
    
    Else
    
    MsgBox "No table or cells selected."
    
    End If
    
    End If

End Sub

         */
        #endregion

        #region TableColumnGaps
        public void TableColumnGapsEven(IRibbonControl control)
        {
            TableColumnGaps("even", 5);
        }
        public void TableColumnGapsOdd(IRibbonControl control)
        {
            TableColumnGaps ("odd", 5);
        }
        #endregion

        #region TableColumnRemoveGaps
        public void TableColumnRemoveGaps(IRibbonControl control)
        {
            var myDocument = ThisAddIn.application.ActiveWindow;
            if (!(myDocument.Selection.Type == PpSelectionType.ppSelectionShapes || myDocument.Selection.Type == PpSelectionType.ppSelectionText))
            {
                MessageBox.Show("No table  selected.");
            }
            else
            {
                if (ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Count == 1 && (ThisAddIn.application.ActiveWindow.Selection.ShapeRange.HasTable == MsoTriState.msoCTrue || ThisAddIn.application.ActiveWindow.Selection.ShapeRange.HasTable == MsoTriState.msoTrue))
                {

                    if (!(ThisAddIn.application.ActiveWindow .Selection.ShapeRange.Tags["INSTRUMENTA COLUMNGAPS"] == "odd" || ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Tags["INSTRUMENTA COLUMNGAPS"] == "even"))
                    {
                        if (MessageBox.Show("No column gaps found, are you sure you want to continue?","",MessageBoxButtons.YesNo,MessageBoxIcon.Question) == DialogResult.No)
                        {
                            return;
                        }
                    }
                    var TypeOfGaps = ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Tags["INSTRUMENTA COLUMNGAPS"];

                    ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Tags.Delete("INSTRUMENTA COLUMNGAPS");
                    //ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table

                    for(int ColumnCount= ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Columns.Count; ColumnCount >= 1; ColumnCount--)
                    {
                        if ((ColumnCount % 2 == 0 && TypeOfGaps == "even") || (ColumnCount % 2 != 0 && TypeOfGaps == "odd"))
                            ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Columns[ColumnCount].Delete();
                        
                    }
                }
                else
                {
                    MessageBox.Show("No table selected or too many shapes selected. Select one table.");
                }

             }
        }
        /*
         Sub TableColumnRemoveGaps()

    Set myDocument = Application.ActiveWindow

    If Not (myDocument.Selection.Type = ppSelectionShapes Or myDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No table selected."
    Else

        If (Application.ActiveWindow.Selection.ShapeRange.Count = 1) And Application.ActiveWindow.Selection.ShapeRange.HasTable Then

            If Not (Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS") = "odd" Or Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS") = "even") Then

                If MsgBox("No column gaps found, are you sure you want to continue?", vbYesNo) = vbNo Then
                    Exit Sub
                End If
            End If

            TypeOfGaps = Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS")

            Application.ActiveWindow.Selection.ShapeRange.Tags.Delete "INSTRUMENTA COLUMNGAPS"

            With Application.ActiveWindow.Selection.ShapeRange.Table

                For ColumnCount = .Columns.Count To 1 Step -1

                    If (ColumnCount Mod 2 = 0 And TypeOfGaps = "even") Or (Not ColumnCount Mod 2 = 0 And TypeOfGaps = "odd") Then
                        .Columns(ColumnCount).Delete
                    End If

                Next ColumnCount

            End With

        Else

            MsgBox "No table selected or too many shapes selected. Select one table."

        End If

    End If

End Sub

         */
        #endregion

        #region TableColumnIncreaseGaps
        public void TableColumnIncreaseGaps(IRibbonControl control)
        {
            var myDocument = ThisAddIn.application.ActiveWindow;
            if (!(myDocument.Selection.Type == PpSelectionType.ppSelectionShapes || myDocument.Selection.Type == PpSelectionType.ppSelectionText))
            {
                MessageBox.Show("No table  selected.");
            }
            else
            {
                if (ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Count == 1 && (ThisAddIn.application.ActiveWindow.Selection.ShapeRange.HasTable == MsoTriState.msoCTrue || ThisAddIn.application.ActiveWindow.Selection.ShapeRange.HasTable == MsoTriState.msoTrue))
                {

                    if (!(ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Tags["INSTRUMENTA COLUMNGAPS"] == "odd" || ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Tags["INSTRUMENTA COLUMNGAPS"] == "even"))
                    {
                        if (MessageBox.Show("No column gaps found, are you sure you want to continue?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                        {
                            return;
                        }
                    }
               

                    var TypeOfGaps = ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Tags["INSTRUMENTA COLUMNGAPS"];

                    
                    //ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table

                    for (int ColumnCount = 1; ColumnCount <= ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Columns.Count; ColumnCount++)
                    {
                        if ((ColumnCount % 2 == 0 && TypeOfGaps == "even") || (ColumnCount % 2 != 0 && TypeOfGaps == "odd"))
                            ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Columns[ColumnCount].Width = ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Columns[ColumnCount].Width + 1;

                    }

                }
                else
                {
                    MessageBox.Show("No table selected or too many shapes selected. Select one table.");
                }

            }
        }

        /*
         Sub TableColumnIncreaseGaps()
    
    Set myDocument = Application.ActiveWindow
    
    If Not (myDocument.Selection.Type = ppSelectionShapes Or myDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No table selected."
    Else
        
        If (Application.ActiveWindow.Selection.ShapeRange.Count = 1) And Application.ActiveWindow.Selection.ShapeRange.HasTable Then
            
            If Not (Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS") = "odd" Or Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS") = "even") Then
                
                If MsgBox("No column gaps found, are you sure you want to continue?", vbYesNo) = vbNo Then
                    Exit Sub
                End If
            End If
            TypeOfGaps = Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS")
            
            With Application.ActiveWindow.Selection.ShapeRange.Table
                
                For ColumnCount = 1 To .Columns.Count
                    
                    If (ColumnCount Mod 2 = 0 And TypeOfGaps = "even") Or (Not ColumnCount Mod 2 = 0 And TypeOfGaps = "odd") Then
                        .Columns(ColumnCount).Width = .Columns(ColumnCount).Width + 1
                    End If
                    
                Next ColumnCount
                
            End With
            
        Else
            
            MsgBox "No table selected or too many shapes selected. Select one table."
            
        End If
        
    End If
    
End Sub

         */
        #endregion
        #region TableColumnDecreaseGaps
        public void TableColumnDecreaseGaps(IRibbonControl control)
        {
            var myDocument = ThisAddIn.application.ActiveWindow;
            if (!(myDocument.Selection.Type == PpSelectionType.ppSelectionShapes || myDocument.Selection.Type == PpSelectionType.ppSelectionText))
            {
                MessageBox.Show("No table  selected.");
            }
            else
            {
                if (ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Count == 1 && (ThisAddIn.application.ActiveWindow.Selection.ShapeRange.HasTable == MsoTriState.msoCTrue || ThisAddIn.application.ActiveWindow.Selection.ShapeRange.HasTable == MsoTriState.msoTrue))
                {

                    if (!(ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Tags["INSTRUMENTA COLUMNGAPS"] == "odd" || ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Tags["INSTRUMENTA COLUMNGAPS"] == "even"))
                    {
                        if (MessageBox.Show("No column gaps found, are you sure you want to continue?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                        {
                            return;
                        }
                    }


                    var TypeOfGaps = ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Tags["INSTRUMENTA COLUMNGAPS"];


                    //ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table

                    for (int ColumnCount = 1; ColumnCount <= ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Columns.Count; ColumnCount++)
                    {
                        if ((ColumnCount % 2 == 0 && TypeOfGaps == "even") || (ColumnCount % 2 != 0 && TypeOfGaps == "odd"))
                        {
                            if(ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Columns[ColumnCount].Width - 1 >= 0)
                              ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Columns[ColumnCount].Width = ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Columns[ColumnCount].Width - 1;
                        }
                    }

                }
                else
                {
                    MessageBox.Show("No table selected or too many shapes selected. Select one table.");
                }

            }
        }

        /*
         Sub TableColumnDecreaseGaps()
    
    Set myDocument = Application.ActiveWindow
    
    If Not (myDocument.Selection.Type = ppSelectionShapes Or myDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No table selected."
    Else
        
        If (Application.ActiveWindow.Selection.ShapeRange.Count = 1) And Application.ActiveWindow.Selection.ShapeRange.HasTable Then
            
            If Not (Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS") = "odd" Or Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS") = "even") Then
                
                If MsgBox("No column gaps found, are you sure you want to continue?", vbYesNo) = vbNo Then
                    Exit Sub
                End If
            End If
            TypeOfGaps = Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS")
            
            With Application.ActiveWindow.Selection.ShapeRange.Table
                
                For ColumnCount = 1 To .Columns.Count
                    
                    If ((ColumnCount Mod 2 = 0 And TypeOfGaps = "even") Or (Not ColumnCount Mod 2 = 0 And TypeOfGaps = "odd") And ((.Columns(ColumnCount).Width - 1) >= 0)) Then
                        .Columns(ColumnCount).Width = .Columns(ColumnCount).Width - 1
                    End If
                    
                Next ColumnCount
                
            End With
            
        Else
            
            MsgBox "No table selected or too many shapes selected. Select one table."
            
        End If
        
    End If
    
End Sub

         */
        #endregion
        #region TableRowGaps
        public void TableRowGapsEven(IRibbonControl control)
        {
            TableRowGaps ("even", 5);
        }
        public void TableRowGapsOdd(IRibbonControl control)
        {
            TableRowGaps( "odd", 5);
        }
        #endregion

        #region TableRowRemoveGaps

        public void TableRowRemoveGaps(IRibbonControl control)
        {
            var myDocument = ThisAddIn.application.ActiveWindow;
            if (!(myDocument.Selection.Type == PpSelectionType.ppSelectionShapes || myDocument.Selection.Type == PpSelectionType.ppSelectionText))
            {
                MessageBox.Show("No table  selected.");
            }
            else
            {
                if (ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Count == 1 && (ThisAddIn.application.ActiveWindow.Selection.ShapeRange.HasTable == MsoTriState.msoCTrue || ThisAddIn.application.ActiveWindow.Selection.ShapeRange.HasTable == MsoTriState.msoTrue))
                {

                    if (!(ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Tags["INSTRUMENTA ROWGAPS"] == "odd" || ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Tags["INSTRUMENTA ROWGAPS"] == "even"))
                    {
                        if (MessageBox.Show("No row gaps found, are you sure you want to continue?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                        {
                            return;
                        }
                    }
  
                    var TypeOfGaps = ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Tags["INSTRUMENTA ROWGAPS"];

                    ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Tags.Delete("INSTRUMENTA ROWGAPS");

                    //ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table

                    for (int RowCount = ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Rows.Count; RowCount >= 1; RowCount--)
                    {
                        if ((RowCount % 2 == 0 && TypeOfGaps == "even") || (RowCount % 2 != 0 && TypeOfGaps == "odd"))
                            ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Rows[RowCount].Delete();

                    }
                }
                else
                {
                    MessageBox.Show("No table selected or too many shapes selected. Select one table.");
                }

            }
        }

        /*
         Sub TableRowRemoveGaps()
    
    Set myDocument = Application.ActiveWindow
    
    If Not (myDocument.Selection.Type = ppSelectionShapes Or myDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No table selected."
    Else
        
        If (Application.ActiveWindow.Selection.ShapeRange.Count = 1) And Application.ActiveWindow.Selection.ShapeRange.HasTable Then
            
            If Not (Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS") = "odd" Or Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS") = "even") Then
                
                If MsgBox("No row gaps found, are you sure you want to continue?", vbYesNo) = vbNo Then
                    Exit Sub
                End If
            End If
            
            TypeOfGaps = Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS")
            
            Application.ActiveWindow.Selection.ShapeRange.Tags.Delete "INSTRUMENTA ROWGAPS"
            
            With Application.ActiveWindow.Selection.ShapeRange.Table
                
                For RowCount = .Rows.Count To 1 Step -1
                    
                    If (RowCount Mod 2 = 0 And TypeOfGaps = "even") Or (Not RowCount Mod 2 = 0 And TypeOfGaps = "odd") Then
                        .Rows(RowCount).Delete
                    End If
                    
                Next RowCount
                
            End With
            
        Else
            
            MsgBox "No table selected or too many shapes selected. Select one table."
            
        End If
        
    End If
    
End Sub

         */

        #endregion

        #region TableRowIncreaseGaps
        public void TableRowIncreaseGaps(IRibbonControl control)
        {
            var myDocument = ThisAddIn.application.ActiveWindow;
            if (!(myDocument.Selection.Type == PpSelectionType.ppSelectionShapes || myDocument.Selection.Type == PpSelectionType.ppSelectionText))
            {
                MessageBox.Show("No table  selected.");
            }
            else
            {
                if (ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Count == 1 && (ThisAddIn.application.ActiveWindow.Selection.ShapeRange.HasTable == MsoTriState.msoCTrue || ThisAddIn.application.ActiveWindow.Selection.ShapeRange.HasTable == MsoTriState.msoTrue))
                {
                   
                    if (!(ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Tags["INSTRUMENTA ROWGAPS"] == "odd" || ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Tags["INSTRUMENTA ROWGAPS"] == "even"))
                    {
                        if (MessageBox.Show("No row gaps found, are you sure you want to continue?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                        {
                            return;
                        }
                    }


                    var TypeOfGaps = ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Tags["INSTRUMENTA ROWGAPS"];


                    //ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table

                    for (int RowCount = 1; RowCount <= ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Rows.Count; RowCount++)
                    {
                        if ((RowCount % 2 == 0 && TypeOfGaps == "even") || (RowCount % 2 != 0 && TypeOfGaps == "odd"))
                            ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Rows[RowCount].Height = ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Rows [RowCount].Height + 1;

                    }
                 
                }
                else
                {
                    MessageBox.Show("No table selected or too many shapes selected. Select one table.");
                }

            }
        }

        /*
         
         Sub TableRowIncreaseGaps()
    
    Set myDocument = Application.ActiveWindow
    
    If Not (myDocument.Selection.Type = ppSelectionShapes Or myDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No table selected."
    Else
        
        If (Application.ActiveWindow.Selection.ShapeRange.Count = 1) And Application.ActiveWindow.Selection.ShapeRange.HasTable Then
            
            If Not (Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS") = "odd" Or Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS") = "even") Then
                
                If MsgBox("No row gaps found, are you sure you want to continue?", vbYesNo) = vbNo Then
                    Exit Sub
                End If
            End If
            TypeOfGaps = Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS")
            
            With Application.ActiveWindow.Selection.ShapeRange.Table
                
                For RowCount = 1 To .Rows.Count
                    
                    If (RowCount Mod 2 = 0 And TypeOfGaps = "even") Or (Not RowCount Mod 2 = 0 And TypeOfGaps = "odd") Then
                        .Rows(RowCount).Height = .Rows(RowCount).Height + 1
                    End If
                    
                Next RowCount
                
            End With
            
        Else
            
            MsgBox "No table selected or too many shapes selected. Select one table."
            
        End If
        
    End If
    
End Sub

         */
        #endregion

        #region TableRowDecreaseGaps
        public void TableRowDecreaseGaps(IRibbonControl control)
        {
            var myDocument = ThisAddIn.application.ActiveWindow;
            if (!(myDocument.Selection.Type == PpSelectionType.ppSelectionShapes || myDocument.Selection.Type == PpSelectionType.ppSelectionText))
            {
                MessageBox.Show("No table  selected.");
            }
            else
            {
                if (ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Count == 1 && (ThisAddIn.application.ActiveWindow.Selection.ShapeRange.HasTable == MsoTriState.msoCTrue || ThisAddIn.application.ActiveWindow.Selection.ShapeRange.HasTable == MsoTriState.msoTrue))
                {

                    if (!(ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Tags["INSTRUMENTA ROWGAPS"] == "odd" || ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Tags["INSTRUMENTA ROWGAPS"] == "even"))
                    {
                        if (MessageBox.Show("No row gaps found, are you sure you want to continue?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                        {
                            return;
                        }
                    }


                    var TypeOfGaps = ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Tags["INSTRUMENTA ROWGAPS"];


                    //ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table

                    for (int RowCount = 1; RowCount <= ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Rows.Count; RowCount++)
                    {
                        if ((RowCount % 2 == 0 && TypeOfGaps == "even") || (RowCount % 2 != 0 && TypeOfGaps == "odd"))
                        {
                            if (ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Rows[RowCount].Height - 1 >= 0)
                             ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Rows[RowCount].Height = ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Rows[RowCount].Height - 1;
                        }

                    }

                }
                else
                {
                    MessageBox.Show("No table selected or too many shapes selected. Select one table.");
                }

            }
        }
        /*
         
Sub TableRowDecreaseGaps()
    
    Set myDocument = Application.ActiveWindow
    
    If Not (myDocument.Selection.Type = ppSelectionShapes Or myDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No table selected."
    Else
        
        If (Application.ActiveWindow.Selection.ShapeRange.Count = 1) And Application.ActiveWindow.Selection.ShapeRange.HasTable Then
            
            If Not (Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS") = "odd" Or Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS") = "even") Then
                
                If MsgBox("No row gaps found, are you sure you want to continue?", vbYesNo) = vbNo Then
                    Exit Sub
                End If
            End If
            TypeOfGaps = Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS")
            
            With Application.ActiveWindow.Selection.ShapeRange.Table
                
                For RowCount = 1 To .Rows.Count
                    
                    If ((RowCount Mod 2 = 0 And TypeOfGaps = "even") Or (Not RowCount Mod 2 = 0 And TypeOfGaps = "odd") And ((.Rows(RowCount).Height - 1) >= 0)) Then
                        .Rows(RowCount).Height = .Rows(RowCount).Height - 1
                    End If
                    
                Next RowCount
                
            End With
            
        Else
            
            MsgBox "No table selected or too many shapes selected. Select one table."
            
        End If
        
    End If
    
End Sub
         */
        #endregion

        #region TablesMarginsToZero
        public void TablesMarginsToZero(IRibbonControl control)
        {
            var myDocument = ThisAddIn.application.ActiveWindow;
            if (!(myDocument.Selection.Type == PpSelectionType.ppSelectionShapes || myDocument.Selection.Type == PpSelectionType.ppSelectionText))
            {
                MessageBox.Show("No table or cells selected.");
            }
            else
            {
                if ( (ThisAddIn.application.ActiveWindow.Selection.ShapeRange.HasTable == MsoTriState.msoCTrue || ThisAddIn.application.ActiveWindow.Selection.ShapeRange.HasTable == MsoTriState.msoTrue))
                {
                    //ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table
                    for(int RowsCount=1;RowsCount <= ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Rows.Count; RowsCount++)
                    {
                        for (int ColsCount = 1; ColsCount <= ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Columns.Count; ColsCount++)
                        {
                            if (ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Selected)
                            {
                                var tf = ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame;
                                tf.MarginBottom = 0;
                                tf.MarginLeft = 0     ;
                                tf.MarginRight = 0    ;
                                tf.MarginTop = 0      ;

                            }
                        }
                    }
                    /*
                       With Application.ActiveWindow.Selection.ShapeRange.Table

                        For RowsCount = 1 To .Rows.Count
                            For ColsCount = 1 To .Columns.Count

                                If .Cell(RowsCount, ColsCount).Selected Then

                                    With .Cell(RowsCount, ColsCount).Shape.TextFrame

                                        .MarginBottom = 0
                                        .MarginLeft = 0
                                        .MarginRight = 0
                                        .MarginTop = 0

                                    End With

                                End If

                            Next ColsCount
                        Next RowsCount

                    End With
                     */
                }
                else
                {
                    MessageBox.Show("No table or cells selected.");
                }
            }
        }
        /*
         Sub TablesMarginsToZero()

    Set myDocument = Application.ActiveWindow

    If Not (myDocument.Selection.Type = ppSelectionShapes Or myDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No table or cells selected."
    Else

    If Application.ActiveWindow.Selection.ShapeRange.HasTable Then

    With Application.ActiveWindow.Selection.ShapeRange.Table

        For RowsCount = 1 To .Rows.Count
            For ColsCount = 1 To .Columns.Count

                If .Cell(RowsCount, ColsCount).Selected Then

                    With .Cell(RowsCount, ColsCount).Shape.TextFrame

                        .MarginBottom = 0
                        .MarginLeft = 0
                        .MarginRight = 0
                        .MarginTop = 0

                    End With

                End If

            Next ColsCount
        Next RowsCount

    End With

    Else

    MsgBox "No table or cells selected."

    End If

    End If

End Sub

         */
        #endregion

        #region TablesMarginsIncrease
        public void TablesMarginsIncrease(IRibbonControl control)
        {
            var myDocument = ThisAddIn.application.ActiveWindow;
            if (!(myDocument.Selection.Type == PpSelectionType.ppSelectionShapes || myDocument.Selection.Type == PpSelectionType.ppSelectionText))
            {
                MessageBox.Show("No table or cells selected.");
            }
            else
            {
                if ((ThisAddIn.application.ActiveWindow.Selection.ShapeRange.HasTable == MsoTriState.msoCTrue || ThisAddIn.application.ActiveWindow.Selection.ShapeRange.HasTable == MsoTriState.msoTrue))
                {
                    //ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table
                    for (int RowsCount = 1; RowsCount <= ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Rows.Count; RowsCount++)
                    {
                        for (int ColsCount = 1; ColsCount <= ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Columns.Count; ColsCount++)
                        {
                            if (ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Selected)
                            {
                                var tf = ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame;
                                tf.MarginBottom +=(float) 0.2;
                                tf.MarginLeft += (float)0.2;
                                tf.MarginRight +=(float) 0.2;
                                tf.MarginTop   +=(float) 0.2;

                            }
                        }
                    }
                    /*
                       
                        .MarginBottom = .MarginBottom + 0.2
                        .MarginLeft = .MarginLeft + 0.2
                        .MarginRight = .MarginRight + 0.2
                        .MarginTop = .MarginTop + 0.2
                     */
                }
                else
                {
                    MessageBox.Show("No table or cells selected.");
                }
            }
        }
        /*
         Sub TablesMarginsIncrease()
    
    Set myDocument = Application.ActiveWindow
    
    If Not (myDocument.Selection.Type = ppSelectionShapes Or myDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No table or cells selected."
    Else
    
    If Application.ActiveWindow.Selection.ShapeRange.HasTable Then
    
    With Application.ActiveWindow.Selection.ShapeRange.Table
        
        For RowsCount = 1 To .Rows.Count
            For ColsCount = 1 To .Columns.Count
                
                If .Cell(RowsCount, ColsCount).Selected Then
                    
                    With .Cell(RowsCount, ColsCount).Shape.TextFrame
                        
                        .MarginBottom = .MarginBottom + 0.2
                        .MarginLeft = .MarginLeft + 0.2
                        .MarginRight = .MarginRight + 0.2
                        .MarginTop = .MarginTop + 0.2
                        
                    End With
                    
                End If
                
            Next ColsCount
        Next RowsCount
        
    End With
    
    Else
    
    MsgBox "No table or cells selected."
    
    End If
    
    End If
    
End Sub

         */
        #endregion

        #region TablesMarginsDecrease
        public void TablesMarginsDecrease(IRibbonControl control)
        {
            var myDocument = ThisAddIn.application.ActiveWindow;
            if (!(myDocument.Selection.Type == PpSelectionType.ppSelectionShapes || myDocument.Selection.Type == PpSelectionType.ppSelectionText))
            {
                MessageBox.Show("No table or cells selected.");
            }
            else
            {
                if ((ThisAddIn.application.ActiveWindow.Selection.ShapeRange.HasTable == MsoTriState.msoCTrue || ThisAddIn.application.ActiveWindow.Selection.ShapeRange.HasTable == MsoTriState.msoTrue))
                {
                    //ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table
                    for (int RowsCount = 1; RowsCount <= ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Rows.Count; RowsCount++)
                    {
                        for (int ColsCount = 1; ColsCount <= ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Columns.Count; ColsCount++)
                        {
                            if (ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Selected)
                            {
                                var tf = ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table.Cell(RowsCount, ColsCount).Shape.TextFrame;
                               
                                if(tf.MarginBottom>=0.2)
                                  tf.MarginBottom -= (float)0.2;
                                if (tf.MarginLeft >= 0.2)
                                    tf.MarginLeft -= (float)0.2;
                                if (tf.MarginRight >= 0.2)
                                    tf.MarginRight -= (float)0.2;
                                if (tf.MarginTop >= 0.2)
                                    tf.MarginTop -= (float)0.2;

                            }
                        }
                    }
                    /*
                       
                           If .Cell(RowsCount, ColsCount).Selected Then
                    
                    With .Cell(RowsCount, ColsCount).Shape.TextFrame
                        
                        If .MarginBottom >= 0.2 Then
                            .MarginBottom = .MarginBottom - 0.2
                        End If
                        If .MarginLeft >= 0.2 Then
                            .MarginLeft = .MarginLeft - 0.2
                        End If
                        If .MarginRight >= 0.2 Then
                            .MarginRight = .MarginRight - 0.2
                        End If
                        If .MarginTop >= 0.2 Then
                            .MarginTop = .MarginTop - 0.2
                        End If
                        
                    End With
                    
                End If
                     */
                }
                else
                {
                    MessageBox.Show("No table or cells selected.");
                }
            }
        }
        #endregion
        #region Misc
        void TableColumnGaps(string TypeOfGaps, double GapSize)
        {
            var myDocument = ThisAddIn.application.ActiveWindow;
            if (!(myDocument.Selection.Type == PpSelectionType.ppSelectionShapes || myDocument.Selection.Type == PpSelectionType.ppSelectionText))
            {
                MessageBox.Show("No table  selected.");
            }
            else
            {
                if (ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Count == 1 && (ThisAddIn.application.ActiveWindow.Selection.ShapeRange.HasTable == MsoTriState.msoCTrue || ThisAddIn.application.ActiveWindow.Selection.ShapeRange.HasTable == MsoTriState.msoTrue))
                {

                    if (!(ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Tags["INSTRUMENTA COLUMNGAPS"] == "odd" || ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Tags["INSTRUMENTA COLUMNGAPS"] == "even"))
                    {
                        if (MessageBox.Show("Existing column gaps found in table, do you want to remove those first?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        {
                            TableColumnRemoveGaps(null);
                        }
                    }
                    ////////////
                    if (TypeOfGaps == "odd")
                        ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Tags.Add("INSTRUMENTA COLUMNGAPS", "odd");
                    else
                        ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Tags.Add("INSTRUMENTA COLUMNGAPS", "even");
                    //////
                    var tbl = ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table;
                    var NumberOfColumns = tbl.Columns.Count;
                    List<double> ColumnWidthArray = new List<double>();
                    var NumberOfNewColumns = 0;
                    if (TypeOfGaps == "odd")
                    {
                        NumberOfNewColumns = NumberOfColumns + NumberOfColumns  + 1;
                        ColumnWidthArray.Clear();
                        for (int ColumnCount = 1; ColumnCount <= NumberOfColumns; ColumnCount++)
                        {
                            ColumnWidthArray.Add(GapSize);
                            ColumnWidthArray.Add(tbl.Columns[ColumnCount].Width );
                            if (ColumnCount == NumberOfColumns )
                            {
                                ColumnWidthArray.Add(GapSize);
                            }
                        }
                    }
                    else
                    {
                        NumberOfNewColumns = NumberOfColumns + NumberOfColumns - 1;
                        for (int ColumnCount = 1; ColumnCount <= NumberOfColumns; ColumnCount++)
                        {
                            if (ColumnCount != 1)
                            {

                                ColumnWidthArray.Add(GapSize);
                                ColumnWidthArray.Add(tbl.Columns[ColumnCount].Width);

                            }
                            else
                            {
                                ColumnWidthArray.Clear();
                                ColumnWidthArray.Add(tbl.Columns[ColumnCount].Width);
                            }
                        }

                    }
                    ///////////
                    for (int ColumnCount = NumberOfColumns; ColumnCount >= 1; ColumnCount--)
                    {

                        if (TypeOfGaps == "odd")
                        {

                            var AddedColumn = tbl.Columns.Add(ColumnCount);

                            for (int CellCount = 1; CellCount <= AddedColumn.Cells.Count; CellCount++)
                            {
                                AddedColumn.Cells[CellCount].Shape.Fill.Visible = MsoTriState.msoFalse;
                                AddedColumn.Cells[CellCount].Borders[PpBorderType.ppBorderTop].Weight = 0;
                                AddedColumn.Cells[CellCount].Borders[PpBorderType.ppBorderBottom ].Weight = 0;
                                AddedColumn.Cells[CellCount].Shape.TextFrame.TextRange.Font.Size = 1;

                                AddedColumn.Cells[CellCount].Shape.TextFrame.MarginBottom = 0;
                                AddedColumn.Cells[CellCount].Shape.TextFrame.MarginLeft = 0;
                                AddedColumn.Cells[CellCount].Shape.TextFrame.MarginRight = 0;
                                AddedColumn.Cells[CellCount].Shape.TextFrame.MarginTop = 0;

                            }

                            if (ColumnCount  == NumberOfColumns)
                            {

                                AddedColumn = tbl.Columns.Add();

                                for (int CellCount = 1; CellCount <= AddedColumn .Cells.Count; CellCount++)
                                {
                                    AddedColumn.Cells[CellCount].Shape.Fill.Visible = MsoTriState.msoFalse;
                                    AddedColumn.Cells[CellCount].Borders[PpBorderType.ppBorderTop].Weight = 0;
                                    AddedColumn.Cells[CellCount].Borders[PpBorderType.ppBorderBottom].Weight = 0;
                                    AddedColumn.Cells[CellCount].Shape.TextFrame.TextRange.Font.Size = 1;
                                    
                                    AddedColumn.Cells[CellCount].Shape.TextFrame.MarginBottom = 0;
                                    AddedColumn.Cells[CellCount].Shape.TextFrame.MarginLeft = 0;
                                    AddedColumn.Cells[CellCount].Shape.TextFrame.MarginRight = 0;
                                    AddedColumn.Cells[CellCount].Shape.TextFrame.MarginTop = 0;

                                }

                            }

                        }
                        else
                        {

                            if (ColumnCount != 1)
                            {

                                var AddedColumn = tbl.Columns.Add(ColumnCount);

                                for (int CellCount = 1; CellCount <= AddedColumn .Cells.Count; CellCount++)
                                {
                                    AddedColumn.Cells[CellCount].Shape.Fill.Visible = MsoTriState.msoFalse;
                                    AddedColumn.Cells[CellCount].Borders[PpBorderType.ppBorderTop].Weight = 0;
                                    AddedColumn.Cells[CellCount].Borders[PpBorderType.ppBorderBottom].Weight = 0;
                                    AddedColumn.Cells[CellCount].Shape.TextFrame.TextRange.Font.Size = 1;
                                    
                                    AddedColumn.Cells[CellCount].Shape.TextFrame.MarginBottom = 0;
                                    AddedColumn.Cells[CellCount].Shape.TextFrame.MarginLeft = 0;
                                    AddedColumn.Cells[CellCount].Shape.TextFrame.MarginRight = 0;
                                    AddedColumn.Cells[CellCount].Shape.TextFrame.MarginTop = 0;

                                }

                            }

                        }

                    }
                    /////////
                    for (int ColumnCount = 1; ColumnCount <= NumberOfNewColumns; ColumnCount++)
                    {
                        tbl.Columns[ColumnCount].Width = (float)ColumnWidthArray[ColumnCount - 1];

                    }

                }
                else
                {
                    MessageBox.Show("No table selected or too many shapes selected. Select one table.");
                }
            }
        }
        /*
         Sub TableColumnGaps(TypeOfGaps As String, GapSize As Double, Optional GapColor As RGBColor)
    
    Set myDocument = Application.ActiveWindow
    
    If Not (myDocument.Selection.Type = ppSelectionShapes Or myDocument.Selection.Type = ppSelectionText) Then
        MsgBox "No table selected."
    Else
        
        If (Application.ActiveWindow.Selection.ShapeRange.Count = 1) And Application.ActiveWindow.Selection.ShapeRange.HasTable Then
            
            If Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS") = "odd" Or Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA COLUMNGAPS") = "even" Then
                
                If MsgBox("Existing column gaps found in table, do you want to remove those first?", vbYesNo) = vbYes Then
                    TableColumnRemoveGaps
                End If
                
            End If
            
            If TypeOfGaps = "odd" Then
                Application.ActiveWindow.Selection.ShapeRange.Tags.Add "INSTRUMENTA COLUMNGAPS", "odd"
            Else
                Application.ActiveWindow.Selection.ShapeRange.Tags.Add "INSTRUMENTA COLUMNGAPS", "even"
            End If
            
            With Application.ActiveWindow.Selection.ShapeRange.Table
                
                NumberOfColumns = .Columns.Count
                Dim ColumnWidthArray() As Double
                
                If TypeOfGaps = "odd" Then
                    
                    NumberOfNewColumns = NumberOfColumns + NumberOfColumns + 1
                    ReDim ColumnWidthArray(0)
                    
                    For ColumnCount = 1 To NumberOfColumns
                        ReDim Preserve ColumnWidthArray(UBound(ColumnWidthArray) + 2)
                        ColumnWidthArray(UBound(ColumnWidthArray) - 1) = .Columns(ColumnCount).Width
                        ColumnWidthArray(UBound(ColumnWidthArray) - 2) = GapSize
                        
                        If ColumnCount = NumberOfColumns Then
                            ReDim Preserve ColumnWidthArray(UBound(ColumnWidthArray) + 1)
                            ColumnWidthArray(UBound(ColumnWidthArray) - 1) = GapSize
                        End If
                        
                    Next ColumnCount
                    
                Else
                    
                    NumberOfNewColumns = NumberOfColumns + NumberOfColumns - 1
                    
                    For ColumnCount = 1 To NumberOfColumns
                        
                        If Not ColumnCount = 1 Then
                            ReDim Preserve ColumnWidthArray(UBound(ColumnWidthArray) + 2)
                            ColumnWidthArray(UBound(ColumnWidthArray) - 1) = .Columns(ColumnCount).Width
                            ColumnWidthArray(UBound(ColumnWidthArray) - 2) = GapSize
                            
                        Else
                            ReDim ColumnWidthArray(1)
                            ColumnWidthArray(UBound(ColumnWidthArray) - 1) = .Columns(ColumnCount).Width
                        End If
                        
                    Next ColumnCount
                    
                End If
                
                For ColumnCount = NumberOfColumns To 1 Step -1
                    
                    If TypeOfGaps = "odd" Then
                        
                        Set AddedColumn = .Columns.Add(ColumnCount)
                        
                        For CellCount = 1 To AddedColumn.Cells.Count
                            AddedColumn.Cells(CellCount).Shape.Fill.Visible = msoFalse
                            AddedColumn.Cells(CellCount).Borders(ppBorderTop).Weight = 0
                            AddedColumn.Cells(CellCount).Borders(ppBorderBottom).Weight = 0
                            AddedColumn.Cells(CellCount).Shape.TextFrame.TextRange.Font.Size = 1
                            
                            AddedColumn.Cells(CellCount).Shape.TextFrame.MarginBottom = 0
                            AddedColumn.Cells(CellCount).Shape.TextFrame.MarginLeft = 0
                            AddedColumn.Cells(CellCount).Shape.TextFrame.MarginRight = 0
                            AddedColumn.Cells(CellCount).Shape.TextFrame.MarginTop = 0
                            
                        Next CellCount
                        
                        If ColumnCount = NumberOfColumns Then
                            
                            Set AddedColumn = .Columns.Add
                            
                            For CellCount = 1 To AddedColumn.Cells.Count
                                AddedColumn.Cells(CellCount).Shape.Fill.Visible = msoFalse
                                AddedColumn.Cells(CellCount).Borders(ppBorderTop).Weight = 0
                                AddedColumn.Cells(CellCount).Borders(ppBorderBottom).Weight = 0
                                AddedColumn.Cells(CellCount).Shape.TextFrame.TextRange.Font.Size = 1
                                
                                AddedColumn.Cells(CellCount).Shape.TextFrame.MarginBottom = 0
                                AddedColumn.Cells(CellCount).Shape.TextFrame.MarginLeft = 0
                                AddedColumn.Cells(CellCount).Shape.TextFrame.MarginRight = 0
                                AddedColumn.Cells(CellCount).Shape.TextFrame.MarginTop = 0
                                
                            Next CellCount
                            
                        End If
                        
                    Else
                        
                        If Not ColumnCount = 1 Then
                            
                            Set AddedColumn = .Columns.Add(ColumnCount)
                            
                            For CellCount = 1 To AddedColumn.Cells.Count
                                AddedColumn.Cells(CellCount).Shape.Fill.Visible = msoFalse
                                AddedColumn.Cells(CellCount).Borders(ppBorderTop).Weight = 0
                                AddedColumn.Cells(CellCount).Borders(ppBorderBottom).Weight = 0
                                AddedColumn.Cells(CellCount).Shape.TextFrame.TextRange.Font.Size = 1
                                
                                AddedColumn.Cells(CellCount).Shape.TextFrame.MarginBottom = 0
                                AddedColumn.Cells(CellCount).Shape.TextFrame.MarginLeft = 0
                                AddedColumn.Cells(CellCount).Shape.TextFrame.MarginRight = 0
                                AddedColumn.Cells(CellCount).Shape.TextFrame.MarginTop = 0
                                
                            Next CellCount
                            
                        End If
                        
                    End If
                    
                Next ColumnCount
                
                For ColumnCount = 1 To NumberOfNewColumns
                    
                    .Columns(ColumnCount).Width = ColumnWidthArray(ColumnCount - 1)
                    
                Next ColumnCount
                
            End With
            
        Else
            
            MsgBox "No table selected or too many shapes selected. Select one table."
            
        End If
        
    End If
    
End Sub

         */
        void TableRowGaps(string TypeOfGaps,double GapSize)
        {
            var myDocument = ThisAddIn.application.ActiveWindow;
            if (!(myDocument.Selection.Type == PpSelectionType.ppSelectionShapes || myDocument.Selection.Type == PpSelectionType.ppSelectionText))
            {
                MessageBox.Show("No table  selected.");
            }
            else
            {
                if (ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Count == 1 && (ThisAddIn.application.ActiveWindow.Selection.ShapeRange.HasTable == MsoTriState.msoCTrue || ThisAddIn.application.ActiveWindow.Selection.ShapeRange.HasTable == MsoTriState.msoTrue))
                {

                    if (!(ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Tags["INSTRUMENTA ROWGAPS"] == "odd" || ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Tags["INSTRUMENTA ROWGAPS"] == "even"))
                    {
                        if (MessageBox.Show("Existing row gaps found in table, do you want to remove those first?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes )
                        {
                            TableRowRemoveGaps(null);
                        }
                    }
                    ////////////
                    if (TypeOfGaps == "odd")
                        ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Tags.Add("INSTRUMENTA ROWGAPS", "odd");
                    else
                        ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Tags.Add("INSTRUMENTA ROWGAPS", "even");
                    //////
                    var tbl =ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Table;
                    var NumberOfRows =tbl .Rows.Count;
                    List<double> RowHeightArray = new List<double>();
                    var NumberOfNewRows = 0;
                    if (TypeOfGaps == "odd")
                    {
                         NumberOfNewRows = NumberOfRows + NumberOfRows + 1;
                        RowHeightArray.Clear();
                        for (int RowCount = 1; RowCount <= NumberOfRows; RowCount++)
                        {
                            RowHeightArray.Add(GapSize);
                            RowHeightArray.Add(tbl.Rows[RowCount].Height);
                            if (RowCount == NumberOfRows)
                            {
                                RowHeightArray.Add(GapSize);
                            }
                        }
                    }
                    else
                    {
                         NumberOfNewRows = NumberOfRows + NumberOfRows - 1;
                        for (int RowCount = 1; RowCount <= NumberOfRows; RowCount++)
                        {
                            if (RowCount != 1) {
                                
                                      RowHeightArray.Add(GapSize);
                                      RowHeightArray.Add(tbl.Rows[RowCount].Height);

                            }
                            else {
                                RowHeightArray.Clear();
                                RowHeightArray.Add(tbl.Rows[RowCount].Height);
                                }
                        }

                    }
                    ///////////
                    for (int RowCount = NumberOfRows; RowCount >= 1; RowCount--) {

                        if (TypeOfGaps == "odd") {

                            var AddedRow = tbl.Rows.Add(RowCount);

                            for (int CellCount = 1; CellCount <= AddedRow.Cells.Count; CellCount++) {
                                AddedRow.Cells[CellCount].Shape.Fill.Visible = MsoTriState.msoFalse;
                                AddedRow.Cells[CellCount].Borders[PpBorderType.ppBorderLeft].Weight = 0;
                                    AddedRow.Cells[CellCount].Borders[PpBorderType.ppBorderRight].Weight = 0;
                                AddedRow.Cells[CellCount].Shape.TextFrame.TextRange.Font.Size = 1;

                                AddedRow.Cells[CellCount].Shape.TextFrame.MarginBottom = 0;
                                AddedRow.Cells[CellCount].Shape.TextFrame.MarginLeft = 0;
                                AddedRow.Cells[CellCount].Shape.TextFrame.MarginRight = 0;
                                AddedRow.Cells[CellCount].Shape.TextFrame.MarginTop = 0;

                            }

                            if (RowCount == NumberOfRows) {

                                AddedRow = tbl.Rows.Add();

                                for (int CellCount = 1; CellCount <= AddedRow.Cells.Count; CellCount++) {
                                    AddedRow.Cells[CellCount].Shape.Fill.Visible = MsoTriState.msoFalse;
                                    AddedRow.Cells[CellCount].Borders[PpBorderType.ppBorderLeft].Weight = 0;
                                        AddedRow.Cells[CellCount].Borders[PpBorderType.ppBorderRight].Weight = 0;
                                        AddedRow.Cells[CellCount].Shape.TextFrame.TextRange.Font.Size = 1;

                                    AddedRow.Cells[CellCount].Shape.TextFrame.MarginBottom = 0;
                                    AddedRow.Cells[CellCount].Shape.TextFrame.MarginLeft = 0;
                                    AddedRow.Cells[CellCount].Shape.TextFrame.MarginRight = 0;
                                    AddedRow.Cells[CellCount].Shape.TextFrame.MarginTop = 0;

                                }

                            }

                        } else {

                            if (RowCount != 1) {

                               var AddedRow = tbl.Rows.Add(RowCount);

                                for (int CellCount = 1; CellCount <= AddedRow.Cells.Count; CellCount++) {
                                    AddedRow.Cells[CellCount].Shape.Fill.Visible = MsoTriState.msoFalse;
                                    AddedRow.Cells[CellCount].Borders[PpBorderType.ppBorderLeft].Weight = 0;
                                    AddedRow.Cells[CellCount].Borders[PpBorderType.ppBorderRight].Weight = 0;
                                    AddedRow.Cells[CellCount].Shape.TextFrame.TextRange.Font.Size = 1;

                                    AddedRow.Cells[CellCount].Shape.TextFrame.MarginBottom = 0;
                                    AddedRow.Cells[CellCount].Shape.TextFrame.MarginLeft = 0;
                                    AddedRow.Cells[CellCount].Shape.TextFrame.MarginRight = 0;
                                    AddedRow.Cells[CellCount].Shape.TextFrame.MarginTop = 0;

                                }

                            }

                        }

                        }
                    /////////
                    for (int RowCount = 1; RowCount <= NumberOfNewRows; RowCount++)
                    {
                        tbl.Rows[RowCount].Height =(float) RowHeightArray[RowCount - 1];

                    }

                }
                else
                {
                    MessageBox.Show("No table selected or too many shapes selected. Select one table.");
                }
            }
        }
                //TableRowGaps
                /*
                 Sub TableRowGaps(TypeOfGaps As String, GapSize As Double, Optional GapColor As RGBColor)

            Set myDocument = Application.ActiveWindow

            If Not (myDocument.Selection.Type = ppSelectionShapes Or myDocument.Selection.Type = ppSelectionText) Then
                MsgBox "No table selected."
            Else

                If (Application.ActiveWindow.Selection.ShapeRange.Count = 1) And Application.ActiveWindow.Selection.ShapeRange.HasTable Then

                    If Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS") = "odd" Or Application.ActiveWindow.Selection.ShapeRange.Tags("INSTRUMENTA ROWGAPS") = "even" Then

                        If MsgBox("Existing row gaps found in table, do you want to remove those first?", vbYesNo) = vbYes Then
                            TableRowRemoveGaps
                        End If

                    End If

                    If TypeOfGaps = "odd" Then
                        Application.ActiveWindow.Selection.ShapeRange.Tags.Add "INSTRUMENTA ROWGAPS", "odd"
                    Else
                        Application.ActiveWindow.Selection.ShapeRange.Tags.Add "INSTRUMENTA ROWGAPS", "even"
                    End If

                    With Application.ActiveWindow.Selection.ShapeRange.Table

                        NumberOfRows = .Rows.Count
                        Dim RowHeightArray() As Double

                        If TypeOfGaps = "odd" Then

                            NumberOfNewRows = NumberOfRows + NumberOfRows + 1
                            ReDim RowHeightArray(0)

                            For RowCount = 1 To NumberOfRows
                                ReDim Preserve RowHeightArray(UBound(RowHeightArray) + 2)
                                RowHeightArray(UBound(RowHeightArray) - 1) = .Rows(RowCount).Height
                                RowHeightArray(UBound(RowHeightArray) - 2) = GapSize

                                If RowCount = NumberOfRows Then
                                    ReDim Preserve RowHeightArray(UBound(RowHeightArray) + 1)
                                    RowHeightArray(UBound(RowHeightArray) - 1) = GapSize
                                End If

                            Next RowCount

                        Else

                            NumberOfNewRows = NumberOfRows + NumberOfRows - 1

                            For RowCount = 1 To NumberOfRows

                                If Not RowCount = 1 Then
                                    ReDim Preserve RowHeightArray(UBound(RowHeightArray) + 2)
                                    RowHeightArray(UBound(RowHeightArray) - 1) = .Rows(RowCount).Height
                                    RowHeightArray(UBound(RowHeightArray) - 2) = GapSize

                                Else
                                    ReDim RowHeightArray(1)
                                    RowHeightArray(UBound(RowHeightArray) - 1) = .Rows(RowCount).Height
                                End If

                            Next RowCount

                        End If
                ***************1***************
                        For RowCount = NumberOfRows To 1 Step -1

                            If TypeOfGaps = "odd" Then

                                Set AddedRow = .Rows.Add(RowCount)

                                For CellCount = 1 To AddedRow.Cells.Count
                                    AddedRow.Cells(CellCount).Shape.Fill.Visible = msoFalse
                                    AddedRow.Cells(CellCount).Borders(ppBorderLeft).Weight = 0
                                    AddedRow.Cells(CellCount).Borders(ppBorderRight).Weight = 0
                                    AddedRow.Cells(CellCount).Shape.TextFrame.TextRange.Font.Size = 1

                                    AddedRow.Cells(CellCount).Shape.TextFrame.MarginBottom = 0
                                    AddedRow.Cells(CellCount).Shape.TextFrame.MarginLeft = 0
                                    AddedRow.Cells(CellCount).Shape.TextFrame.MarginRight = 0
                                    AddedRow.Cells(CellCount).Shape.TextFrame.MarginTop = 0

                                Next CellCount

                                If RowCount = NumberOfRows Then

                                    Set AddedRow = .Rows.Add

                                    For CellCount = 1 To AddedRow.Cells.Count
                                        AddedRow.Cells(CellCount).Shape.Fill.Visible = msoFalse
                                        AddedRow.Cells(CellCount).Borders(ppBorderLeft).Weight = 0
                                        AddedRow.Cells(CellCount).Borders(ppBorderRight).Weight = 0
                                        AddedRow.Cells(CellCount).Shape.TextFrame.TextRange.Font.Size = 1

                                        AddedRow.Cells(CellCount).Shape.TextFrame.MarginBottom = 0
                                        AddedRow.Cells(CellCount).Shape.TextFrame.MarginLeft = 0
                                        AddedRow.Cells(CellCount).Shape.TextFrame.MarginRight = 0
                                        AddedRow.Cells(CellCount).Shape.TextFrame.MarginTop = 0

                                    Next CellCount

                                End If

                            Else

                                If Not RowCount = 1 Then

                                    Set AddedRow = .Rows.Add(RowCount)

                                    For CellCount = 1 To AddedRow.Cells.Count
                                        AddedRow.Cells(CellCount).Shape.Fill.Visible = msoFalse
                                        AddedRow.Cells(CellCount).Borders(ppBorderLeft).Weight = 0
                                        AddedRow.Cells(CellCount).Borders(ppBorderRight).Weight = 0
                                        AddedRow.Cells(CellCount).Shape.TextFrame.TextRange.Font.Size = 1

                                        AddedRow.Cells(CellCount).Shape.TextFrame.MarginBottom = 0
                                        AddedRow.Cells(CellCount).Shape.TextFrame.MarginLeft = 0
                                        AddedRow.Cells(CellCount).Shape.TextFrame.MarginRight = 0
                                        AddedRow.Cells(CellCount).Shape.TextFrame.MarginTop = 0

                                    Next CellCount

                                End If

                            End If

                        Next RowCount
                ******************2****************
                        For RowCount = 1 To NumberOfNewRows

                            .Rows(RowCount).Height = RowHeightArray(RowCount - 1)

                        Next RowCount

                    End With

                Else

                    MsgBox "No table selected or too many shapes selected. Select one table."

                End If

            End If

        End Sub

                 */
                void ConnectRectangleShapes(string ShapeDirection)
        {
            if(ThisAddIn.application.ActiveWindow.Selection.Type==PpSelectionType.ppSelectionShapes)
            {
                if (ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Count == 1)
                {
                    float Left1, Right1, Top1, Bottom1, Left2, Right2, Top2, Bottom2;
                    Left1 = ThisAddIn.application.ActiveWindow.Selection.ShapeRange[1].Left;
                    Right1 = Left1 + ThisAddIn.application.ActiveWindow.Selection.ShapeRange[1].Width;
                    Top1 = ThisAddIn.application.ActiveWindow.Selection.ShapeRange[1].Top;
                    Bottom1 = Top1 + ThisAddIn.application.ActiveWindow.Selection.ShapeRange[1].Height;

                    Left2 = ThisAddIn.application.ActiveWindow.Selection.ShapeRange[2].Left;
                    Right2 = Left2 + ThisAddIn.application.ActiveWindow.Selection.ShapeRange[2].Width;
                    Top2 = ThisAddIn.application.ActiveWindow.Selection.ShapeRange[2].Top;
                    Bottom2 = Top2 + ThisAddIn.application.ActiveWindow.Selection.ShapeRange[2].Height;

                    var myDocument = ThisAddIn.application.ActiveWindow.Selection.SlideRange;

                    switch (ShapeDirection)
                    {
                      case "RightToLeft":
                            var obj= myDocument.Shapes.BuildFreeform(MsoEditingType.msoEditingCorner, Right1, Top1);
                            obj.AddNodes(MsoSegmentType.msoSegmentLine, MsoEditingType. msoEditingAuto,Right1,  Bottom1);
                            obj.AddNodes(MsoSegmentType.msoSegmentLine, MsoEditingType.msoEditingAuto, Left2 , Bottom2);
                            obj.AddNodes(MsoSegmentType.msoSegmentLine, MsoEditingType.msoEditingAuto, Left2, Top2 );
                            obj.AddNodes(MsoSegmentType.msoSegmentLine, MsoEditingType.msoEditingAuto, Right1 , Top1);
                            obj.ConvertToShape().Line.Visible = MsoTriState.msoFalse;
                            break;
                        case "LeftToRight":
                            var obj2 = myDocument.Shapes.BuildFreeform(MsoEditingType.msoEditingCorner, Right2, Top2);
                            obj2.AddNodes(MsoSegmentType.msoSegmentLine, MsoEditingType.msoEditingAuto, Right2, Bottom2);
                            obj2.AddNodes(MsoSegmentType.msoSegmentLine, MsoEditingType.msoEditingAuto, Left1, Bottom1);
                            obj2.AddNodes(MsoSegmentType.msoSegmentLine, MsoEditingType.msoEditingAuto, Left1, Top1);
                            obj2.AddNodes(MsoSegmentType.msoSegmentLine, MsoEditingType.msoEditingAuto, Right2, Top2);
                            obj2.ConvertToShape().Line.Visible = MsoTriState.msoFalse;
                            break;
                        case "BottomToTop":
                            var obj3 = myDocument.Shapes.BuildFreeform(MsoEditingType.msoEditingCorner, Left1, Bottom1);
                            obj3.AddNodes(MsoSegmentType.msoSegmentLine, MsoEditingType.msoEditingAuto, Right1, Bottom1);
                            obj3.AddNodes(MsoSegmentType.msoSegmentLine, MsoEditingType.msoEditingAuto, Right2, Top2);
                            obj3.AddNodes(MsoSegmentType.msoSegmentLine, MsoEditingType.msoEditingAuto, Left2, Top2);
                            obj3.AddNodes(MsoSegmentType.msoSegmentLine, MsoEditingType.msoEditingAuto, Left1, Bottom1);
                            obj3.ConvertToShape().Line.Visible = MsoTriState.msoFalse;
                            break;

                        case "TopToBottom":
                            var obj4 = myDocument.Shapes.BuildFreeform(MsoEditingType.msoEditingCorner, Left2 , Bottom2 );
                            obj4.AddNodes(MsoSegmentType.msoSegmentLine, MsoEditingType.msoEditingAuto, Right2, Bottom2);
                            obj4.AddNodes(MsoSegmentType.msoSegmentLine, MsoEditingType.msoEditingAuto, Right1 , Top1 );
                            obj4.AddNodes(MsoSegmentType.msoSegmentLine, MsoEditingType.msoEditingAuto, Left1, Top1);
                            obj4.AddNodes(MsoSegmentType.msoSegmentLine, MsoEditingType.msoEditingAuto, Left2 , Bottom2 );
                            obj4.ConvertToShape().Line.Visible = MsoTriState.msoFalse;
                            break;

                    }
                }
                else
                {
                    MessageBox.Show("Select two shapes.");
                }
            }
            else
            {
                MessageBox.Show("Select two shapes.");
            }
        }
        /*
         Sub ConnectRectangleShapes(ShapeDirection As String)
    
    If ActiveWindow.Selection.Type = ppSelectionShapes Then
    
    If ActiveWindow.Selection.ShapeRange.Count = 2 Then
    
    Dim Left1, Right1, Top1, Bottom1, Left2, Right2, Top2, Bottom2 As Double
    
    
    Left1 = ActiveWindow.Selection.ShapeRange(1).Left
    Right1 = Left1 + ActiveWindow.Selection.ShapeRange(1).Width
    Top1 = ActiveWindow.Selection.ShapeRange(1).Top
    Bottom1 = Top1 + ActiveWindow.Selection.ShapeRange(1).Height
    
    Left2 = ActiveWindow.Selection.ShapeRange(2).Left
    Right2 = Left2 + ActiveWindow.Selection.ShapeRange(2).Width
    Top2 = ActiveWindow.Selection.ShapeRange(2).Top
    Bottom2 = Top2 + ActiveWindow.Selection.ShapeRange(2).Height
    
    Set myDocument = Application.ActiveWindow.Selection.SlideRange
    
    Select Case ShapeDirection
    
    Case "RightToLeft"
        With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, X1:=Right1, Y1:=Top1)
            .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Right1, Y1:=Bottom1
            .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Left2, Y1:=Bottom2
            .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Left2, Y1:=Top2
            .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Right1, Y1:=Top1
            '.ConvertToShape
            .ConvertToShape.Line.Visible = msoFalse
        End With
        
    Case "LeftToRight"
        With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, X1:=Right2, Y1:=Top2)
            .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Right2, Y1:=Bottom2
            .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Left1, Y1:=Bottom1
            .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Left1, Y1:=Top1
            .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Right2, Y1:=Top2
            '.ConvertToShape
            .ConvertToShape.Line.Visible = msoFalse
        End With
        
     Case "BottomToTop"
        With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, X1:=Left1, Y1:=Bottom1)
            .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Right1, Y1:=Bottom1
            .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Right2, Y1:=Top2
            .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Left2, Y1:=Top2
            .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Left1, Y1:=Bottom1
            '.ConvertToShape
            .ConvertToShape.Line.Visible = msoFalse
        End With
        
     Case "TopToBottom"
        With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, X1:=Left2, Y1:=Bottom2)
            .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Right2, Y1:=Bottom2
            .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Right1, Y1:=Top1
            .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Left1, Y1:=Top1
            .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, X1:=Left2, Y1:=Bottom2
            '.ConvertToShape
            .ConvertToShape.Line.Visible = msoFalse
        End With
        
        
    End Select
    
    Else
    MsgBox "Select two shapes."
    End If
    
    Else
    MsgBox "Select two shapes."
    End If
    
    
End Sub

         */
        void ObjectsTextInsertSpecialCharacter(int SpecialCharacter)
        {
            if(ThisAddIn.application.ActiveWindow.Selection.Type==PpSelectionType.ppSelectionText)
            {
                ThisAddIn.application.ActiveWindow.Selection.TextRange.InsertSymbol(
                    ThisAddIn.application.ActiveWindow.Selection.TextRange.Font.Name ,
                    SpecialCharacter,
                    MsoTriState.msoTrue
                    );
            }
        }
        /*
         Sub ObjectsTextInsertSpecialCharacter(SpecialCharacter As Long)

If ActiveWindow.Selection.Type = ppSelectionText Then

Application.ActiveWindow.Selection.TextRange.InsertSymbol Application.ActiveWindow.Selection.TextRange.Font.Name, SpecialCharacter, MsoTriState.msoTrue

End If
        Endn sub
         */
        #endregion
    
        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
      {
         Assembly asm = Assembly.GetExecutingAssembly();
         string[] resourceNames = asm.GetManifestResourceNames();
         for (int i = 0; i < resourceNames.Length; ++i)
         {
            if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
            {
               using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
               {
                  if (resourceReader != null)
                  {
                     return resourceReader.ReadToEnd();
                  }
               }
            }
         }
         return null;
      }

      #endregion
   }
}
