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
        #region Misc
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
