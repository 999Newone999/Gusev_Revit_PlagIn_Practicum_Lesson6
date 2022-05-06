using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.Revit.ApplicationServices;

namespace Gusev_Revit_PlagIn_Practicum_Lesson6
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class CreationModel : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            List<String> levelNames = new List<String>();
            levelNames.Add("Уровень 1");
            levelNames.Add("Уровень 2");

            List<Level> levels = GetLevels(doc, levelNames);

            List<Wall> walls = CreateWalls(doc, 10000, 5000, 0, 0, levels.ElementAt(0), levels.ElementAt(1));


            Transaction transaction = new Transaction(doc, "Установка двери");
            transaction.Start();
            AddDoor(doc, levels.ElementAt(0), walls[0]);
            transaction.Commit();

            Transaction transaction1 = new Transaction(doc, "Установка окон");
            transaction.Start();
            AddWindow(doc, levels.ElementAt(0), walls[1], 1000);
            AddWindow(doc, levels.ElementAt(0), walls[2], 1000);
            AddWindow(doc, levels.ElementAt(0), walls[3], 1000);
            transaction.Commit();

            Transaction transaction2 = new Transaction(doc, "Создание крыши");
            transaction.Start();
            //AddRoof(doc, levels.ElementAt(1), walls);
            AddRoof2(doc, levels.ElementAt(1), walls);
            transaction.Commit();

            return Result.Succeeded;
        }

        private void AddRoof2(Document doc, Level level, List<Wall> walls)
        {
            
            RoofType roofType = new FilteredElementCollector(doc)
                                    .OfClass(typeof(RoofType))
                                    .OfType<RoofType>()
                                    .Where(x => x.Name.Equals("Типовой - 400мм"))
                                    .Where(x => x.FamilyName.Equals("Базовая крыша"))
                                    .FirstOrDefault();
            
            int indexOfShortestWall = 0;
            int indexOfLongestWall = 1;
            if ((walls[indexOfShortestWall].Location as LocationCurve).Curve.Length <
                (walls[indexOfShortestWall + 1].Location as LocationCurve).Curve.Length)
            {
                indexOfShortestWall = 0;
                indexOfLongestWall = 1;
            }
            else
            {
                indexOfShortestWall = 1;
                indexOfLongestWall = 0;
            }
            LocationCurve curve = walls[indexOfShortestWall].Location as LocationCurve;
            XYZ _p1 = curve.Curve.GetEndPoint(0);
            XYZ _p2 = curve.Curve.GetEndPoint(1);

            double wallWidth = walls[0].Width;
            double dt = wallWidth / 2;

            List<XYZ> points = new List<XYZ>();
            points.Add(new XYZ(-dt, -dt, 0));
            points.Add(new XYZ(dt, -dt, 0));
            points.Add(new XYZ(dt, dt, 0));

            XYZ p1 = _p1 + points[indexOfShortestWall];
            XYZ p2 = _p2 + points[indexOfShortestWall + 1];

            double zCoord =0.258819*(curve.Curve.Length/2 + dt); //угол крыши 15 градусов
            double height = walls[0].get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).AsDouble();
            XYZ pHeight = new XYZ(0, 0, height);
            XYZ pZcoord = new XYZ(0, 0, zCoord);
            p1 = p1 + pHeight;
            p2 = p2 + pHeight;
            XYZ p3 = (p1+p2)/2 + pZcoord;


            CurveArray curveArray = new CurveArray();
            curveArray.Append(Line.CreateBound(p1, p3));
            curveArray.Append(Line.CreateBound(p3, p2));
            //curveArray.Append(Line.CreateBound(new XYZ(0,p1.Y,p1.Z), new XYZ(0,p3.Y,p3.Z)));
            //curveArray.Append(Line.CreateBound(new XYZ(0, p3.Y, p3.Z), new XYZ(0, p2.Y, p2.Z)));

            // Line line2 = Line.CreateBound(p3, p1);
            // curveArray.Append(line2);
            Autodesk.Revit.DB.View view = new FilteredElementCollector(doc)
                                    .OfClass(typeof(Autodesk.Revit.DB.View))
                                    .OfType<Autodesk.Revit.DB.View>()
                                    .Where(x => x.Id.Equals(level.FindAssociatedPlanViewId()))
                                    .FirstOrDefault();

            ReferencePlane plane = doc.Create.NewReferencePlane2(p1, p3, p2, view);
            //ReferencePlane plane = doc.Create.NewReferencePlane(new XYZ(0, 0, 0), new XYZ(0, 0, 20), new XYZ(0, 20, 0), doc.ActiveView);

            doc.Create.NewExtrusionRoof(curveArray, plane, level, roofType, 0,
                ((walls[indexOfLongestWall].Location as LocationCurve).Curve.Length+2*dt));

        }

        private void AddRoof(Document doc, Level level, List<Wall> walls)
        {
            RoofType roofType = new FilteredElementCollector(doc)
                                    .OfClass(typeof(RoofType))
                                    .OfType<RoofType>()
                                    .Where(x => x.Name.Equals("Типовой - 400мм"))
                                    .Where(x => x.FamilyName.Equals("Базовая крыша"))
                                    .FirstOrDefault();

            double wallWidth = walls[0].Width;
            double dt = wallWidth / 2;
            List<XYZ> points = new List<XYZ>();
            points.Add(new XYZ(-dt, -dt, 0));
            points.Add(new XYZ(dt, -dt, 0));
            points.Add(new XYZ(dt, dt, 0));
            points.Add(new XYZ(-dt, dt, 0));
            points.Add(new XYZ(-dt, -dt, 0));


            Autodesk.Revit.ApplicationServices.Application application = doc.Application;
            CurveArray footprint = application.Create.NewCurveArray();
            for (int i = 0; i < 4; i++)
            {
                LocationCurve curve = walls[i].Location as LocationCurve;
                XYZ p1 = curve.Curve.GetEndPoint(0);
                XYZ p2 = curve.Curve.GetEndPoint(1);
                Line line = Line.CreateBound(p1 + points[i], p2 + points[i + 1]);
                footprint.Append(line);
            }
            ModelCurveArray footPrintToModelCurveMapping = new ModelCurveArray();
            FootPrintRoof footPrintRoof = doc.Create.NewFootPrintRoof(footprint, level, roofType,
                                          out footPrintToModelCurveMapping);
            ModelCurveArrayIterator iterator = footPrintToModelCurveMapping.ForwardIterator();
            iterator.Reset();
            while (iterator.MoveNext())
            {
                ModelCurve modelCurve = iterator.Current as ModelCurve;
                footPrintRoof.set_DefinesSlope(modelCurve, true);
                footPrintRoof.set_SlopeAngle(modelCurve, 0.5);
            }
            
        }

        private void AddWindow(Document doc, Level level, Wall wall, double _height)
        {
            double height = UnitUtils.ConvertToInternalUnits(_height, UnitTypeId.Millimeters);

            FamilySymbol windowType = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_Windows)
                .OfType<FamilySymbol>()
                .Where(x => x.Name.Equals("0915 x 1220 мм"))
                .Where(x => x.FamilyName.Equals("Фиксированные"))
                .FirstOrDefault();
            LocationCurve hostCurve = wall.Location as LocationCurve;
            XYZ point1 = hostCurve.Curve.GetEndPoint(0);
            XYZ point2 = hostCurve.Curve.GetEndPoint(1);
            XYZ pointHeight = new XYZ(0, 0, height);
            XYZ point = (point1 + point2) / 2;
            point = point + pointHeight;

            if (!windowType.IsActive)
                windowType.Activate();

            doc.Create.NewFamilyInstance(point, windowType, wall, level, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

        }

        private void AddDoor(Document doc, Level level, Wall wall)
        {
            FamilySymbol doorType = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_Doors)
                .OfType<FamilySymbol>()
                .Where(x => x.Name.Equals("0915 x 2134 мм"))
                .Where(x => x.FamilyName.Equals("Одиночные-Щитовые"))
                .FirstOrDefault();
            LocationCurve hostCurve = wall.Location as LocationCurve;
            XYZ point1 = hostCurve.Curve.GetEndPoint(0);
            XYZ point2 = hostCurve.Curve.GetEndPoint(1);
            XYZ point = (point1 + point2)/2;

            if (!doorType.IsActive)
                doorType.Activate();

            doc.Create.NewFamilyInstance(point, doorType, wall, level, Autodesk.Revit.DB.Structure.StructuralType.NonStructural) ;

        }

        // Метод получающий существующие уровни из документа, имена которых перечислены в входном списке

        public List<Level> GetLevels(Document doc, List<String> levelNames)
        {
            List<Level> listNamedlevel = new List<Level>();
            List<Level> listlevel = new FilteredElementCollector(doc)
                            .OfClass(typeof(Level))
                            .OfType<Level>()
                            .ToList();
            foreach (String levelName in levelNames)
            {
                try
                {
                    listNamedlevel.Add(listlevel.FirstOrDefault(x => x.Name.Equals(levelName)));
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            return listNamedlevel;
        }

        public List<Wall> CreateWalls(Document doc, double _width, double _depth, double x, double y,
                                    Level baseLevel, Level upperLevel)
        {
            double width = UnitUtils.ConvertToInternalUnits(_width, UnitTypeId.Millimeters);
            double depth = UnitUtils.ConvertToInternalUnits(_depth, UnitTypeId.Millimeters);
            double dx = width / 2;
            double dy = depth / 2;

            List<XYZ> points = new List<XYZ>();
            points.Add(new XYZ(x-dx, y-dy, 0));
            points.Add(new XYZ(x+dx, y-dy, 0));
            points.Add(new XYZ(x+dx, y+dy, 0));
            points.Add(new XYZ(x-dx, y+dy, 0));
            points.Add(new XYZ(x-dx, y-dy, 0));

            List<Wall> walls = new List<Wall>();

            Transaction transaction = new Transaction(doc, "Построение стен");
            transaction.Start();
            for (int i = 0; i < 4; i++)
            {
                Line line = Line.CreateBound(points[i], points[i + 1]);
                Wall wall = Wall.Create(doc, line, baseLevel.Id, false);
                wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE).Set(upperLevel.Id);
                walls.Add(wall);
            }

            transaction.Commit();

            return walls;
        }
    }
}
