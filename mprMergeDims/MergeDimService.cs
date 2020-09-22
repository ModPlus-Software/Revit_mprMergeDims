namespace mprMergeDims
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using Autodesk.Revit.DB;
    using Autodesk.Revit.UI;
    using ModPlusAPI;
    using ModPlusAPI.Windows;

    /// <summary>
    /// Сервис слияния размеров
    /// </summary>
    public class MergeDimService
    {
        private readonly UIApplication _application;

        /// <summary>
        /// Initializes a new instance of the <see cref="MergeDimService"/> class.
        /// </summary>
        /// <param name="application">Current Revit <see cref="UIApplication"/></param>
        public MergeDimService(UIApplication application)
        {
            _application = application;
        }

        /// <summary>
        /// Выполнить процедуру слияния выбранных размеров
        /// </summary>
        public void MergeExecute()
        {
            var selection = _application.ActiveUIDocument.Selection;
            var doc = _application.ActiveUIDocument.Document;
            var prompt = Language.GetItem(new ModPlusConnector().Name, "h1");
            const double tol = 0.0001;

            while (true)
            {
                try
                {
                    var dimensions = selection
                        .PickElementsByRectangle(new DimensionsFilter(), prompt)
                        .OfType<Dimension>().ToList();

                    var byDimLineDictionary = new Dictionary<Line, List<Dimension>>();
                    foreach (var dimension in dimensions)
                    {
                        if (!(dimension.Curve is Line line))
                            continue;

                        var isMatch = false;
                        foreach (var pair in byDimLineDictionary)
                        {
                            if ((Math.Abs(pair.Key.Origin.DistanceTo(line.Origin)) < tol &&
                                 Math.Abs(Math.Abs(pair.Key.Direction.DotProduct(line.Direction)) - 1) < tol) ||
                                Math.Abs(Math.Abs(pair.Key.Direction.DotProduct(Line.CreateBound(pair.Key.Origin, line.Origin).Direction)) - 1) < tol)
                            {
                                isMatch = true;
                                pair.Value.Add(dimension);
                                break;
                            }
                        }

                        if (!isMatch)
                        {
                            byDimLineDictionary.Add(line, new List<Dimension> { dimension });
                        }
                    }

                    var transactionName = Language.GetFunctionLocalName(new ModPlusConnector());
                    if (string.IsNullOrEmpty(transactionName))
                        transactionName = "Merge dimensions";

                    using (var t = new TransactionGroup(doc, transactionName))
                    {
                        t.Start();
                        foreach (var pair in byDimLineDictionary)
                        {
                            if (pair.Value.Count <= 1)
                                continue;
                            try
                            {
                                using (var tr = new Transaction(doc, "Merge dims"))
                                {
                                    tr.Start();

                                    var referenceArray = new ReferenceArray();
                                    var view = pair.Value.First().View;
                                    var type = pair.Value.First().DimensionType;
                                    foreach (var d in pair.Value)
                                    {
                                        foreach (Reference reference in d.References)
                                        {
                                            if (reference.ElementId != ElementId.InvalidElementId &&
                                                doc.GetElement(reference.ElementId) is Grid grid)
                                            {
                                                var fromGrid = GetReferenceFromGrid(grid);
                                                if (fromGrid != null)
                                                    referenceArray.Append(fromGrid);
                                            }
                                            else
                                            {
                                                referenceArray.Append(reference);
                                            }
                                        }
                                    }

                                    if (doc.Create.NewDimension(view, pair.Key, referenceArray, type) != null)
                                    {
                                        doc.Delete(pair.Value.Select(d => d.Id).ToList());
                                    }

                                    tr.Commit();
                                }
                            }
                            catch (Exception exception)
                            {
                                ExceptionBox.Show(exception);
                            }
                        }

                        t.Assimilate();
                    }
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                    break;
                }
            }
        }

        private static Reference GetReferenceFromGrid(Grid grid)
        {
            var optionsAllGeometry = new Options
            {
                ComputeReferences = true,
                View = grid.Document.ActiveView,
                IncludeNonVisibleObjects = true
            };
            var wasException = false;
            try
            {
                var geometry = grid.get_Geometry(optionsAllGeometry).GetTransformed(Transform.Identity);
                if (geometry != null)
                {
                    foreach (var geometryObject in geometry)
                    {
                        if (geometryObject is Line line && line.Reference != null)
                        {
                            return line.Reference;
                        }
                    }
                }
            }
            catch
            {
                wasException = true;
            }

            if (wasException)
            {
                try
                {
                    var gridLine = grid.Curve as Line;
                    if (gridLine != null && gridLine.Reference != null)
                    {
                        return gridLine.Reference;
                    }
                }
                catch (Exception exception)
                {
                    ExceptionBox.Show(exception);
                }
            }

            return null;
        }
    }
}
