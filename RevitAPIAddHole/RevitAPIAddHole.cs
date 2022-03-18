using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPIAddHole
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class RevitAPIAddHole : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            #region Получение данных из проекта Revit

            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document ardoc = uiDoc.Document;
            Document ovDoc = ardoc.Application.Documents.OfType<Document>().Where(x => x.Title.Contains("ОВ")).FirstOrDefault();
            if (ovDoc == null)
            {
                TaskDialog.Show("Ошибка", "Не найден ОВ файл");
                return Result.Cancelled;
            }

            FamilySymbol familySymbol = new FilteredElementCollector(ardoc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .OfType<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("Отверстие"))
                .FirstOrDefault();
            if (familySymbol == null)
            {
                TaskDialog.Show("Ошибка", "Не найден семейство \"Отверстие\"");
                return Result.Cancelled;
            }

            List<Duct> ducts = new FilteredElementCollector(ovDoc)
                .OfClass(typeof(Duct))
                .OfType<Duct>()
                .ToList();

            List<Pipe> pipes = new FilteredElementCollector(ovDoc)
                .OfClass(typeof(Pipe))
                .OfType<Pipe>()
                .ToList();

            View3D view3D = new FilteredElementCollector(ardoc)
                .OfClass(typeof(View3D))
                .OfType<View3D>()
                .Where(x => !x.IsTemplate)
                .FirstOrDefault();
            if (view3D == null)
            {
                TaskDialog.Show("Ошибка", "Не найден 3D вид");
                return Result.Cancelled;
            }

            ReferenceIntersector referenceIntersector = new ReferenceIntersector(new ElementClassFilter(typeof(Wall)), FindReferenceTarget.Element, view3D);
            #endregion



            #region Команды, выполняющие построение в Revit

            using (var ts = new Transaction(ardoc, "Создание отверстий"))
            {
                ts.Start();

                if (!familySymbol.IsActive)
                    familySymbol.Activate();


                foreach (Duct duct in ducts)
                {
                    Line line = (duct.Location as LocationCurve).Curve as Line;
                    XYZ point = line.GetEndPoint(0);
                    XYZ direction = line.Direction;

                    List<ReferenceWithContext> intersections = referenceIntersector.Find(point, direction)
                        .Where(x => x.Proximity <= line.Length)
                        .Distinct(new ReferenceWithContextElementEqualityComparer())
                        .ToList();

                    foreach (ReferenceWithContext refer in intersections)
                    {
                        double proximity = refer.Proximity;
                        Reference reference = refer.GetReference();
                        Wall wall = ardoc.GetElement(reference.ElementId) as Wall;
                        Level level = ardoc.GetElement(wall.LevelId) as Level;
                        XYZ pointHole = point + (direction * proximity);

                        FamilyInstance hole = ardoc.Create.NewFamilyInstance(pointHole, familySymbol, wall, level, StructuralType.NonStructural);

                        Parameter width = hole.LookupParameter("ширина");
                        Parameter height = hole.LookupParameter("высота");
                        width.Set(duct.Diameter);
                        height.Set(duct.Diameter);
                    }
                }

                foreach (Pipe pipe in pipes)
                {
                    Line line = (pipe.Location as LocationCurve).Curve as Line;
                    XYZ point = line.GetEndPoint(0);
                    XYZ direction = line.Direction;

                    List<ReferenceWithContext> intersections = referenceIntersector.Find(point, direction)
                        .Where(x => x.Proximity <= line.Length)
                        .Distinct(new ReferenceWithContextElementEqualityComparer())
                        .ToList();

                    foreach (ReferenceWithContext refer in intersections)
                    {
                        double proximity = refer.Proximity;
                        Reference reference = refer.GetReference();
                        Wall wall = ardoc.GetElement(reference.ElementId) as Wall;
                        Level level = ardoc.GetElement(wall.LevelId) as Level;
                        XYZ pointHole = point + (direction * proximity);

                        FamilyInstance hole = ardoc.Create.NewFamilyInstance(pointHole, familySymbol, wall, level, StructuralType.NonStructural);

                        Parameter width = hole.LookupParameter("ширина");
                        Parameter height = hole.LookupParameter("высота");
                        width.Set(pipe.Diameter);
                        height.Set(pipe.Diameter);
                    }
                }
                ts.Commit();
            }
            return Result.Succeeded;
            #endregion
        }


        #region Вспомогательные классы
        public class ReferenceWithContextElementEqualityComparer : IEqualityComparer<ReferenceWithContext>
        {
            public bool Equals(ReferenceWithContext x, ReferenceWithContext y)
            {
                if (ReferenceEquals(x, y)) return true;
                if (ReferenceEquals(null, x)) return false;
                if (ReferenceEquals(null, y)) return false;

                var xReference = x.GetReference();

                var yReference = y.GetReference();

                return xReference.LinkedElementId == yReference.LinkedElementId
                           && xReference.ElementId == yReference.ElementId;
            }

            public int GetHashCode(ReferenceWithContext obj)
            {
                var reference = obj.GetReference();

                unchecked
                {
                    return (reference.LinkedElementId.GetHashCode() * 397) ^ reference.ElementId.GetHashCode();
                }
            }
        }
        #endregion
    }
}
