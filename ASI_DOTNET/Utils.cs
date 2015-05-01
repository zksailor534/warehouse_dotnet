using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Windows;

namespace ASI_DOTNET
{
    public static class Utils
    {
        public static Solid3d SweepPolylineOverLine(Polyline p,
            Line l)
        {
            /// Create a region from the polyline
            DBObjectCollection acDBObjColl = new DBObjectCollection();
            acDBObjColl.Add(p);
            DBObjectCollection myRegionColl = new DBObjectCollection();
            myRegionColl = Region.CreateFromCurves(acDBObjColl);
            Region region = myRegionColl[0] as Region;

            /// Create 3D solid and sweep existing region along path
            Solid3d solSweep = new Solid3d();
            SweepOptionsBuilder sob =
                new SweepOptionsBuilder();
            solSweep.CreateSweptSolid(region, l, sob.ToSweepOptions());

            return solSweep;
        }

        public static Solid3d ExtrudePolyline(Polyline p,
            double height,
            double taperAngle = 0)
        {
            /// Create a region from the polyline
            DBObjectCollection acDBObjColl = new DBObjectCollection();
            acDBObjColl.Add(p);
            DBObjectCollection myRegionColl = new DBObjectCollection();
            myRegionColl = Region.CreateFromCurves(acDBObjColl);
            Region region = myRegionColl[0] as Region;

            /// Create 3D solid and sweep existing region along path
            Solid3d solExt = new Solid3d();
            solExt.Extrude(region, height, taperAngle);

            return solExt;
        }

        public static Solid3d ExtrudeCircle(Circle c,
            double height,
            double taperAngle = 0)
        {
            /// Create a region from the polyline
            DBObjectCollection acDBObjColl = new DBObjectCollection();
            acDBObjColl.Add(c);
            DBObjectCollection myRegionColl = new DBObjectCollection();
            myRegionColl = Region.CreateFromCurves(acDBObjColl);
            Region region = myRegionColl[0] as Region;

            /// Create 3D solid and sweep existing region along path
            Solid3d solExt = new Solid3d();
            solExt.Extrude(region, height, taperAngle);

            return solExt;
        }

        public static double ToRadians(this double val)
        {
            return (Math.PI / 180) * val;
        }
    }
}
