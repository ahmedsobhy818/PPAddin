using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PPAddin.Forms
{
    public partial class CopyShapeToMultipleSlidesForm : Form
    {
        public CopyShapeToMultipleSlidesForm()
        {
            InitializeComponent();
        }

        private void CopyShapesToSelectedSlidesButton_Click(object sender, EventArgs e)
        {
            CopyShapeToMultipleSlides();
        }

        void CopyShapeToMultipleSlides()
        {
            //Microsoft.Office.Interop.PowerPoint.Shape Shape;
            Random r = new Random();
            int RandomNumber = r.Next(1000000);
            string CrossSlideShapeId;
            bool OverwriteExisting, SkipSlide;

            OverwriteExisting = this.OptionExistingShapes1.Checked;
            CrossSlideShapeId = this.ShapeIdentifierTextBox.Text;

            ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Tags.Add("INSTRUMENTA CROSS-SLIDE SHAPE", CrossSlideShapeId);

            for (int SelectedCount = 0; SelectedCount <= this.AllSlidesListBox.Items.Count - 1; SelectedCount++)
            {
                if (this.AllSlidesListBox.SelectedItems.Contains(AllSlidesListBox.Items[SelectedCount]))
                {
                    SkipSlide = false;
                    foreach (Microsoft.Office.Interop.PowerPoint.Shape Shape in ThisAddIn.application.ActivePresentation.Slides[SelectedCount].Shapes)
                    {
                        if (Shape.Tags["INSTRUMENTA CROSS-SLIDE SHAPE"] == CrossSlideShapeId)
                        {

                            if (OverwriteExisting)
                            {
                                Shape.Delete();
                            }
                            else
                                SkipSlide = true;


                        }
                    }

                    if (!SkipSlide) {
                        ThisAddIn.application.ActiveWindow.Selection.ShapeRange.Copy();
                        var PastedShape =ThisAddIn.application. ActivePresentation.Slides[SelectedCount ].Shapes.Paste();
                        PastedShape.Name = ShapeIdentifierTextBox.Text + RandomNumber.ToString();
                        PastedShape.Tags.Add("INSTRUMENTA CROSS-SLIDE SHAPE", CrossSlideShapeId);
            }
                }
            }
            Hide();
        }


        /*
         Sub CopyShapeToMultipleSlides()
    
    Dim Shape       As Shape
    RandomNumber = Round(Rnd() * 1000000, 0)
    
    Dim OverwriteExisting As Boolean
    Dim CrossSlideShapeId As String
    Dim SkipSlide   As Boolean
    
    OverwriteExisting = CopyShapeToMultipleSlidesForm.OptionExistingShapes1.Value
    CrossSlideShapeId = CopyShapeToMultipleSlidesForm.ShapeIdentifierTextBox.Value
    
    Application.ActiveWindow.Selection.ShapeRange.Tags.Add "INSTRUMENTA CROSS-SLIDE SHAPE", CrossSlideShapeId
    
    For SelectedCount = 0 To CopyShapeToMultipleSlidesForm.AllSlidesListBox.ListCount - 1
        If (CopyShapeToMultipleSlidesForm.AllSlidesListBox.Selected(SelectedCount) = True) Then
            
            SkipSlide = False
            
            For Each Shape In ActivePresentation.Slides(CLng(CopyShapeToMultipleSlidesForm.AllSlidesListBox.List(SelectedCount))).Shapes
                
                If Shape.Tags("INSTRUMENTA CROSS-SLIDE SHAPE") = CrossSlideShapeId Then
                    
                    If OverwriteExisting = True Then
                        
                        Shape.Delete
                        
                    Else
                        
                        SkipSlide = True
                        
                    End If
                    
                End If
                
            Next
            
            If SkipSlide = False Then
                Application.ActiveWindow.Selection.ShapeRange.Copy
                Set PastedShape = ActivePresentation.Slides(CLng(CopyShapeToMultipleSlidesForm.AllSlidesListBox.List(SelectedCount))).Shapes.Paste
                PastedShape.Name = CopyShapeToMultipleSlidesForm.ShapeIdentifierTextBox.Value + Str(RandomNumber)
                PastedShape.Tags.Add "INSTRUMENTA CROSS-SLIDE SHAPE", CrossSlideShapeId
            End If
            
        End If
    Next SelectedCount
    
    CopyShapeToMultipleSlidesForm.Hide
    
End Sub

         */
    }
}
