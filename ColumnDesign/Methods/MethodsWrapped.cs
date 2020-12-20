using System;
using Autodesk.Revit.UI;
using ColumnDesign.UI;

namespace ColumnDesign.Methods
{
    /// <summary>
    /// This is an example of of wrapping a method with an ExternalEventHandler using an instance of WPF
    /// as an argument. Any type of argument can be passed to the RevitEventWrapper, and therefore be used in
    /// the execution of a method which has to take place within a "Valid Revit API Context". This specific
    /// pattern can be useful for smaller applications, where it is convenient to access the WPF properties
    /// directly, but can become cumbersome in larger application architectures. At that point, it is suggested
    /// to use more "low-level" wrapping, as with the string-argument-wrapped method above.
    /// </summary>
    public class EventHandlerWithWpfArg : RevitEventWrapper<ColumnCreatorView>
    {
        /// <summary>
        /// The Execute override void must be present in all methods wrapped by the RevitEventWrapper.
        /// This defines what the method will do when raised externally.
        /// </summary>
        protected override void Execute(UIApplication uiApp, ColumnCreatorView ui, DrawingTypes drawingType)
        {
            try
            {
                switch (drawingType)
                {
                    case DrawingTypes.Gates:
                        Methods.CreateGates(uiApp.ActiveUIDocument.Document, ui);
                        break;
                    case DrawingTypes.Scissors:
                        Methods.CreateScissors(uiApp.ActiveUIDocument.Document, ui);
                        break;
                    default:
                        throw new ArgumentOutOfRangeException(nameof(drawingType), drawingType, null);
                }
            }
            catch (Exception e)
            {
                TaskDialog.Show("Error", e.Message + e.StackTrace);
            }
        }
    }
}
