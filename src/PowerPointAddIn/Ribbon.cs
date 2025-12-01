using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace PowerPointAddIn
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("PowerPointAddIn.Ribbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        // Called when Ribbon loads
        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;

            var app = Globals.ThisAddIn.Application;
            app.WindowSelectionChange += App_WindowSelectionChange;
        }

        // EditBox dynamic text
        public string GetEditBox(IRibbonControl control)
        {
            return ConvertRadiusToPercentage(GetCurrentRadius());
        }

        // EditBox changed by user
        public void EditCornerRadius_Changed(IRibbonControl control, string text)
        {
            string input = text.Replace("%", "").Trim();
            if (float.TryParse(input, out float percent))
            {
                float radius = ConvertPercentageToRadius(percent);
                ApplyCornerRadius(radius);

                // Reformat text with % sign
                ribbon.InvalidateControl("editCornerRadius");
            }
        }

        // Increase button clicked
        public void BtnDecreaseCornerRadius_Click(IRibbonControl control)
        {
            float currentRadius = GetCurrentRadius();
            float newRadius = Math.Max(currentRadius - 0.01f, 0.0f);
            ApplyCornerRadius(newRadius);
            ribbon.InvalidateControl("editCornerRadius");
        }

        // Decrease button clicked
        public void BtnIncreaseCornerRadius_Click(IRibbonControl control)
        {
            float currentRadius = GetCurrentRadius();
            float newRadius = Math.Min(currentRadius + 0.01f, 0.5f);
            ApplyCornerRadius(newRadius);
            ribbon.InvalidateControl("editCornerRadius");
        }

        public void BtnMinCornerRadius_Click(IRibbonControl control)
        {
            ApplyCornerRadius(0.0f); // 0% radius
            ribbon.InvalidateControl("editCornerRadius");
        }

        public void BtnMaxCornerRadius_Click(IRibbonControl control)
        {
            ApplyCornerRadius(0.5f); // 100% radius
            ribbon.InvalidateControl("editCornerRadius");
        }

        #endregion

        #region Event handlers

        private void App_WindowSelectionChange(PowerPoint.Selection Sel)
        {
            // Refresh the editBox when selection changes
            ribbon.InvalidateControl("editCornerRadius");
        }

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

        private float GetCurrentRadius()
        {
            var app = Globals.ThisAddIn.Application;
            var sel = app.ActiveWindow.Selection;

            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                var shape = sel.ShapeRange[1];
                if (shape.AutoShapeType == MsoAutoShapeType.msoShapeRoundedRectangle)
                {
                    return shape.Adjustments[1];
                }
            }
            return 0.0f;
        }

        private void ApplyCornerRadius(float radius)
        {
            var app = Globals.ThisAddIn.Application;
            var sel = app.ActiveWindow.Selection;

            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                foreach (PowerPoint.Shape shape in sel.ShapeRange)
                {
                    if (shape.AutoShapeType == MsoAutoShapeType.msoShapeRoundedRectangle)
                    {
                        shape.Adjustments[1] = radius;
                    }
                }
            }
        }

        private string ConvertRadiusToPercentage(float radius)
        {
            return ((radius / 0.5f) * 100f).ToString("0") + "%";
        }

        private float ConvertPercentageToRadius(float percent)
        {
            return Math.Min(Math.Max((percent / 100f) * 0.5f, 0.0f), 0.5f);
        }

        #endregion
    }
}

