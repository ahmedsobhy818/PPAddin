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
        #region Misc

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
