using Microsoft.Office.Core;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using System.Diagnostics;
using GroupShapes = Microsoft.Office.Interop.PowerPoint.GroupShapes;

namespace TestVSTO
{
    public partial class Ribbon1
    {
        public static string sourceFilePath = Environment.ExpandEnvironmentVariables("%HOMEDRIVE%%HOMEPATH%") + "\\DarwinSource\\master-visualisation-source.pptx";
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            //Microsoft.Office.Interop.PowerPoint.Application pptApp = new Microsoft.Office.Interop.PowerPoint.Application();
            //Microsoft.Office.Interop.PowerPoint.Presentations pptPresentations = pptApp.Presentations;
            //Microsoft.Office.Interop.PowerPoint.Presentation pptPresentation = pptPresentations.Open(
            //    sourceFilePath,
            //    MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse
            //);

            //Microsoft.Office.Interop.PowerPoint.Slide pptSlide = pptPresentation.Slides[15];
            ////<darwin><data><keys><key><name>charts</name><index></index></key></keys></data></darwin>
            //foreach (Shape shape in pptSlide.Shapes)
            //{
            //    if(shape.Type == MsoShapeType.msoGroup)
            //    {
            //        GroupShapes visualisationShapes = shape.GroupItems;

            //        foreach (Shape s in visualisationShapes)
            //        {
            //            if (s.Type == MsoShapeType.msoTextBox)
            //            {
            //                var addXML = s.CustomerData.Add();
            //                addXML.LoadXML("<darwin><data><keys><key><name>charts</name><index></index></key></keys></data></darwin>");

            //                int count = s.CustomerData.Count;

            //                for (int i = 1; i <= count; i++)
            //                {
            //                    string ss = s.CustomerData._Index(i).XML.ToString();
            //                    Debug.WriteLine(ss);
            //                }
            //            }
            //        }
            //    }
            //}

            //pptPresentation.Save();
        }
    }
}
