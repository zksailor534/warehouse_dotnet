using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Colors;

namespace ASI_DOTNET
{
    public static class Utils
    {
        public static void SetBlockObjectProperties(Entity ent)
        {
            // Set layer and properties
            ent.Layer = "0";
            ent.Color = ChooseColor("ByBlock");
            ent.Linetype = "ByBlock";
            ent.LineWeight = LineWeight.ByBlock;
        }

        public static void CreateLayer(Database db,
            string name,
            Color color)
        {
            // Start a transaction
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                // Returns the layer table for the current database
                LayerTable acLyrTbl;
                acLyrTbl = acTrans.GetObject(db.LayerTableId,
                                            OpenMode.ForWrite) as LayerTable;
 
                if (acLyrTbl.Has(name) == false)
                {
                    LayerTableRecord acLyrTblRec = new LayerTableRecord();
 
                    // Assign the layer the color and name
                    acLyrTblRec.Name = name;
                    acLyrTblRec.Color = color;
 
                    // Upgrade the Layer table for write
                    acLyrTbl.UpgradeOpen();
 
                    // Append the new layer to the Layer table and the transaction
                    acLyrTbl.Add(acLyrTblRec);
                    acTrans.AddNewlyCreatedDBObject(acLyrTblRec, true);
                }
 
                // Save changes and dispose of the transaction
                acTrans.Commit();
            }
        }

        public static Color ChooseColor(string name)
        {
            short id = 7;

            switch (name)
            {
                case "red":
                    id = 1;
                    break;
                case "yellow":
                    id = 52;
                    break;
                case "green":
                    id = 3;
                    break;
                case "cyan":
                    id = 4;
                    break;
                case "blue":
                    id = 5;
                    break;
                case "magenta":
                    id = 6;
                    break;
                case "white":
                    id = 7;
                    break;
                case "orange":
                    id = 30;
                    break;
                case "teal":
                    id = 134;
                    break;
                case "black":
                    id = 250;
                    break;
                case "ByBlock":
                    id = 0;
                    break;
                case "ByLayer":
                    id = 256;
                    break;
                default:
                    Application.ShowAlertDialog("Invalid Color Name: " + name);
                    break;
            }
            
            return Color.FromColorIndex(ColorMethod.ByAci, id);
        }

        public static Solid3d SweepPolylineOverLine(Polyline p,
            Line l)
        {
            Solid3d solSweep = new Solid3d();
            try
            {
                // Create a region from the polyline
                DBObjectCollection acDBObjColl = new DBObjectCollection();
                acDBObjColl.Add(p);
                DBObjectCollection myRegionColl = new DBObjectCollection();
                myRegionColl = Region.CreateFromCurves(acDBObjColl);
                Region region = myRegionColl[0] as Region;

                // Create 3D solid and sweep existing region along path
                SweepOptionsBuilder sob =
                    new SweepOptionsBuilder();
                solSweep.CreateSweptSolid(region, l, sob.ToSweepOptions());
            }
            catch (Autodesk.AutoCAD.Runtime.Exception Ex)
            {
                Application.ShowAlertDialog("Unable to sweep the following polyline:\n" +
                    "polyline: " + p.ToString() + "\n" +
                    "line: " + l.ToString() + "\n" +
                    Ex.Message);
            }
            return solSweep;
        }

        public static Solid3d ExtrudePolyline(Polyline p,
            double height,
            double taperAngle = 0)
        {
            Solid3d solExt = new Solid3d();
            try
            {
                // Create a region from the polyline
                DBObjectCollection acDBObjColl = new DBObjectCollection();
                acDBObjColl.Add(p);
                DBObjectCollection myRegionColl = new DBObjectCollection();
                myRegionColl = Region.CreateFromCurves(acDBObjColl);
                Region region = myRegionColl[0] as Region;

                // Create 3D solid and sweep existing region along path
                solExt.Extrude(region, height, taperAngle);
            }
            catch (Autodesk.AutoCAD.Runtime.Exception Ex)
            {
                Application.ShowAlertDialog("Unable to extrude the following polyline:\n" +
                    "polyline: " + p.ToString() + "\n" +
                    "height: " + height.ToString() + "\n" +
                    Ex.Message);
            }
            return solExt;
        }

        public static Solid3d ExtrudeCircle(Circle c,
            double height,
            double taperAngle = 0)
        {
            Solid3d solExt = new Solid3d();
            try
            {
                // Create a region from the polyline
                DBObjectCollection acDBObjColl = new DBObjectCollection();
                acDBObjColl.Add(c);
                DBObjectCollection myRegionColl = new DBObjectCollection();
                myRegionColl = Region.CreateFromCurves(acDBObjColl);
                Region region = myRegionColl[0] as Region;

                // Create 3D solid and sweep existing region along path
                solExt.Extrude(region, height, taperAngle);
            }
            catch (Autodesk.AutoCAD.Runtime.Exception Ex)
            {
                Application.ShowAlertDialog("Unable to extrude the following circle:\n" +
                    "circle: " + c.ToString() + "\n" +
                    "height: " + height.ToString() + "\n" +
                    Ex.Message);
            }
            return solExt;
        }

        public static double ToRadians(this double val)
        {
            return (Math.PI / 180) * val;
        }

        public static double ToDegrees(this double val)
        {
            return (180 / Math.PI) * val;
        }

        public static double PolarAngleTheta(Line l)
        {
            double theta;
            double x = l.EndPoint.X - l.StartPoint.X;
            double y = l.EndPoint.Y - l.StartPoint.Y;
            double z = l.EndPoint.Z - l.StartPoint.Z;
            theta = Math.Acos(z / Math.Sqrt(x * x + y * y + z * z));

            return theta;
        }

        public static double PolarAnglePhi(Line l)
        {
            double x = l.EndPoint.X - l.StartPoint.X;
            double y = l.EndPoint.Y - l.StartPoint.Y;
            return Math.Atan2(y, x);
        }

        // Creates exactly orthogonal angles (removes small variations)
        public static double OrthogonalAngle(double angle)
        {
            double orthoAngle = Math.PI / 2;
            double multiple = Math.Round(angle / orthoAngle, 0);
            return orthoAngle * multiple;
        }

        public static Vector3d UnitVector(this Vector3d val)
        {
            return val / val.Length;
        }
    }
}
