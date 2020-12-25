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
            var elementsIds = new FilteredElementCollector(doc, doc.ActiveView.Id)
                .WhereElementIsNotElementType()
                .Select(e => e.Id.IntegerValue)
                .ToList();

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
                            if (pair.Value.Count < 1)
                                continue;
                            try
                            {
                                using (var tr = new Transaction(doc, "Merge dims"))
                                {
                                    tr.Start();

                                    var referenceArray = new ReferenceArray();
                                    var view = pair.Value.First().View;
                                    var type = pair.Value.First().DimensionType;
                                    var dimsData = new List<DimData>();
                                    foreach (var d in pair.Value)
                                    {
                                        if (d.NumberOfSegments > 0)
                                            dimsData.AddRange(from DimensionSegment segment in d.Segments select new DimData(segment));
                                        else
                                            dimsData.Add(new DimData(d));

                                        foreach (Reference reference in d.References)
                                        {
                                            if (reference.ElementId != ElementId.InvalidElementId &&
                                                !elementsIds.Contains(reference.ElementId.IntegerValue))
                                                continue;

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

                                    if (doc.Create.NewDimension(view, pair.Key, referenceArray, type) is Dimension createdDimension)
                                    {
                                        if (ModPlus_Revit.Utils.Dimensions.TryRemoveZeroes(createdDimension, out referenceArray))
                                        {
                                            if (doc.Create.NewDimension(view, pair.Key, referenceArray, type) is Dimension reCreatedDimension)
                                            {
                                                doc.Delete(createdDimension.Id);
                                                RestoreTextFields(reCreatedDimension, dimsData);
                                            }
                                        }
                                        else
                                        {
                                            RestoreTextFields(createdDimension, dimsData);
                                        }

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

        private void RestoreTextFields(Dimension dimension, IReadOnlyCollection<DimData> dimsData)
        {
            foreach (DimensionSegment dimensionSegment in dimension.Segments)
            {
                var dimData = dimsData.FirstOrDefault(d => 
                    d.IsMatchValue(dimensionSegment.Value) &&
                    Math.Abs(d.Origin.DistanceTo(dimensionSegment.Origin)) < 0.0001);
                
                if (dimData == null) 
                    continue;

                dimensionSegment.Prefix = dimData.Prefix;
                dimensionSegment.Suffix = dimData.Suffix;
                dimensionSegment.Above = dimData.Above;
                dimensionSegment.Below = dimData.Below;
                dimensionSegment.TextPosition = dimData.TextPosition;
            }
        }

        /// <summary>
        /// <see cref="Dimension"/>/<see cref="DimensionSegment"/> data
        /// </summary>
        internal class DimData
        {
            /// <summary>
            /// Initializes a new instance of the <see cref="DimData"/> class.
            /// </summary>
            /// <param name="dimension">Instance of <see cref="Dimension"/> without segments</param>
            public DimData(Dimension dimension)
            {
                Value = dimension.Value;
                Origin = dimension.Origin;
                TextPosition = dimension.TextPosition;
                Prefix = dimension.Prefix;
                Suffix = dimension.Suffix;
                Above = dimension.Above;
                Below = dimension.Below;
            }

            /// <summary>
            /// Initializes a new instance of the <see cref="DimData"/> class.
            /// </summary>
            /// <param name="dimension">Instance of <see cref="DimensionSegment"/></param>
            public DimData(DimensionSegment dimension)
            {
                Value = dimension.Value;
                Origin = dimension.Origin;
                TextPosition = dimension.TextPosition;
                Prefix = dimension.Prefix;
                Suffix = dimension.Suffix;
                Above = dimension.Above;
                Below = dimension.Below;
            }

            /// <summary>
            /// <see cref="Dimension.Value"/>
            /// </summary>
            public double? Value { get; }

            /// <summary>
            /// <see cref="Dimension.Origin"/>
            /// </summary>
            public XYZ Origin { get; }

            /// <summary>
            /// <see cref="Dimension.TextPosition"/>
            /// </summary>
            public XYZ TextPosition { get; }

            /// <summary>
            /// <see cref="Dimension.Prefix"/>
            /// </summary>
            public string Prefix { get; }

            /// <summary>
            /// <see cref="Dimension.Suffix"/>
            /// </summary>
            public string Suffix { get; }

            /// <summary>
            /// <see cref="Dimension.Above"/>
            /// </summary>
            public string Above { get; }

            /// <summary>
            /// <see cref="Dimension.Below"/>
            /// </summary>
            public string Below { get; }

            /// <summary>
            /// Is match <see cref="Value"/> to other value
            /// </summary>
            /// <param name="value">other value</param>
            public bool IsMatchValue(double? value)
            {
                if (Value.HasValue && value.HasValue)
                    return Math.Abs(Value.Value - value.Value) < 0.0001;
                
                return !Value.HasValue && !value.HasValue;
            }
        }
    }
}
