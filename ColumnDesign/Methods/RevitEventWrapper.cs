using Autodesk.Revit.UI;

namespace ColumnDesign.Methods
{
    /// <summary>
    /// Class for creating Argument (Wrapped) External Events
    /// </summary>
    /// <typeparam name="TType">The Class type being wrapped for the External Event Handler.</typeparam>
    /// <typeparam name="TType2">The Class type being wrapped for the External Event Handler.</typeparam>
    public abstract class RevitEventWrapper<TType, TType2> : IExternalEventHandler
    {
        private readonly object _lock;
        private TType _savedArgs;
        private TType2 _savedArgs2;
        private readonly ExternalEvent _revitEvent;
        private DrawingTypes _drawingType;
        /// <summary>
        /// Class for wrapping methods for execution within a "valid" Revit API context.
        /// </summary>
        protected RevitEventWrapper()
        {
            _revitEvent = ExternalEvent.Create(this);
            _lock = new object();
        }

        /// <summary>
        /// Wraps the "Execution" method in a valid Revit API context.
        /// </summary>
        /// <param name="app">Revit UI Application to use as the "wrapper" API context.</param>
        public void Execute(UIApplication app)
        {
            TType args;
            TType2 args2;
            DrawingTypes drawingType;
            lock (_lock)
            {
                args = _savedArgs;
                args2 = _savedArgs2;
                drawingType = _drawingType;
                _savedArgs = default;
            }

            Execute(app, args, args2, drawingType);
        }

        /// <summary>
        /// Get the name of the operation.
        /// </summary>
        /// <returns>Operation Name.</returns>
        public string GetName()
        {
            return GetType().Name;
        }

        /// <summary>
        /// Execute the wrapped external event in a valid Revit API context.
        /// </summary>
        /// <param name="args">Arguments that could be passed to the execution method.</param>
        /// <param name="args2">Arguments that could be passed to the execution method.</param>
        /// <param name="drawingTypes"></param>
        public void Raise(TType args, TType2 args2, DrawingTypes drawingTypes)
        {
            lock (_lock)
            {
                _savedArgs = args;
                _savedArgs2 = args2;
                _drawingType = drawingTypes;
            }

            _revitEvent.Raise();
        }

        /// <summary>
        /// Override void which wraps the "Execution" method in a valid Revit API context.
        /// </summary>
        /// <param name="app">Revit UI Application to use as the "wrapper" API context.</param>
        /// <param name="view">View</param>
        /// <param name="vm">ViewModel</param>
        /// <param name="drawingType">Type of sheet to create</param>
        protected abstract void Execute(UIApplication app, TType view, TType2 vm, DrawingTypes drawingType);
    }
}