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

[assembly: CommandClass(typeof(ASI_DOTNET.Column))]
[assembly: CommandClass(typeof(ASI_DOTNET.Truss))]

namespace ASI_DOTNET
{
    class Column
    {
        public static void CreateColumn(Database acCurDb,
            double cHeight,
            double cDiameter,
            double bDiameter)
        {
            // Column block name
            string cName = "Column - " + cHeight + " - " + cDiameter + "x" + cDiameter;

            // Baseplate default height (could be argument)
            double bHeight = 0.75;

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // Open the Block table record for read
                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                                             OpenMode.ForRead) as BlockTable;

                // Blank objectid reference
                ObjectId fId = ObjectId.Null;

                if (acBlkTbl.Has(cName))
                {
                    // Retrieve object id
                    fId = acBlkTbl[cName];
                }
                else
                {
                    using (BlockTableRecord acBlkTblRec = new BlockTableRecord())
                    {
                        acBlkTblRec.Name = cName;

                        // Set the insertion point for the block
                        acBlkTblRec.Origin = new Point3d(0, 0, 0);
                        
                        // Calculate baseplate locations
                        Point3d bMid = new Point3d(bDiameter / 2, bDiameter / 2, bHeight / 2);
                        Vector3d bVec = Point3d.Origin.GetVectorTo(bMid);
                        var cBottom = cHeight / 2 + bHeight;
                        Vector3d cVec;
                        Point3d cMid;

                        //// Calculate column location relative to baseplate
                        //if (cPlaceRes.StringResult == "Corner")
                        //{
                        //    cMid = new Point3d(cWidth / 2, cWidth / 2, cBottom);
                        //}
                        //else if (cPlaceRes.StringResult == "Side")
                        //{
                        //    cMid = new Point3d(bWidth / 2, cWidth / 2, cBottom);
                        //}
                        //else
                        //{
                            cMid = new Point3d(bDiameter / 2, bDiameter / 2, cBottom);
                        //}

                        cVec = Point3d.Origin.GetVectorTo(cMid);
                
                        // Create the 3D solid baseplate
                        Solid3d bPlate = new Solid3d();
                        bPlate.SetDatabaseDefaults();
                        bPlate.CreateBox(bDiameter, bDiameter, bHeight);

                        // Position the baseplate 
                        bPlate.TransformBy(Matrix3d.Displacement(bVec));

                        // Add baseplate to block
                        acBlkTblRec.AppendEntity(bPlate);

                        // Create the 3D solid column
                        Solid3d column = new Solid3d();
                        column.SetDatabaseDefaults();
                        column.CreateBox(cDiameter, cDiameter, cHeight - bHeight);

                        // Position the column 
                        column.TransformBy(Matrix3d.Displacement(cVec));

                        // Add column to block
                        acBlkTblRec.AppendEntity(column);

                        /// Add Block to Block Table and close Transaction
                        acBlkTbl.UpgradeOpen();
                        acBlkTbl.Add(acBlkTblRec);
                        acTrans.AddNewlyCreatedDBObject(acBlkTblRec, true);
                    }

                }

                // Save the new object to the database
                acTrans.Commit();

            }

        }


        [CommandMethod("ColumnPrompt")]
        public static void ColumnPrompt()
        {
            // Frame default characteristics
            double cHeight = 96;
            double cDiameter = 6;
            double bDiameter = 16;

            // Get the current document and database, and start a transaction
            // !!! Move this outside function
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            // Prompt for the column height
            PromptDoubleResult cHeightRes;
            PromptDistanceOptions cHeightOpts = new PromptDistanceOptions("");
            cHeightOpts.Message = "\nEnter the total column height: ";
            cHeightRes = acDoc.Editor.GetDistance(cHeightOpts);
            cHeight = cHeightRes.Value;

            // Exit if the user presses ESC or cancels the command
            if (cHeightRes.Status == PromptStatus.Cancel) return;

            // Prompt for the column width
            PromptDoubleResult cDiameterRes;
            PromptDistanceOptions cDiameterOpts = new PromptDistanceOptions("");
            cDiameterOpts.Message = "\nEnter the column diameter: ";
            cDiameterRes = acDoc.Editor.GetDistance(cDiameterOpts);
            cDiameter = cDiameterRes.Value;

            // Exit if the user presses ESC or cancels the command
            if (cDiameterRes.Status == PromptStatus.Cancel) return;

            // Prompt for the base width
            PromptDoubleResult bDiameterRes;
            PromptDistanceOptions bDiameterOpts = new PromptDistanceOptions("");
            bDiameterOpts.Message = "\nEnter the baseplate diameter: ";
            bDiameterRes = acDoc.Editor.GetDistance(bDiameterOpts);
            bDiameter = bDiameterRes.Value;

            // Exit if the user presses ESC or cancels the command
            if (bDiameterRes.Status == PromptStatus.Cancel) return;

            //// Prompt for the base height
            //PromptDoubleResult bHeightRes;
            //PromptDistanceOptions bHeightOpts = new PromptDistanceOptions("");
            //bHeightOpts.Message = "\nEnter the baseplate height: ";
            //bHeightOpts.DefaultValue = 0.75;
            //bHeightRes = acDoc.Editor.GetDistance(bHeightOpts);
            //double bHeight = bHeightRes.Value;

            //// Exit if the user presses ESC or cancels the command
            //if (bHeightRes.Status == PromptStatus.Cancel) return;

            //// Prompt for the column orientation on the baseplate
            //PromptKeywordOptions cPlaceOpts = new PromptKeywordOptions("");
            //cPlaceOpts.Message = "\nEnter baseplate offset: ";
            //cPlaceOpts.Keywords.Add("Center");
            //cPlaceOpts.Keywords.Add("Side");
            //cPlaceOpts.Keywords.Add("Corner");
            //cPlaceOpts.Keywords.Default = "Center";
            //cPlaceOpts.AllowArbitraryInput = false;
            //PromptResult cPlaceRes = acDoc.Editor.GetKeywords(cPlaceOpts);

            //// Exit if the user presses ESC or cancels the command
            //if (cPlaceRes.Status == PromptStatus.Cancel) return;

            CreateColumn(acCurDb, cHeight, cDiameter, bDiameter);

            //// Open the active viewport
            //ViewportTableRecord acVportTblRec;
            //acVportTblRec = acTrans.GetObject(acDoc.Editor.ActiveViewportId,
            //                                  OpenMode.ForWrite) as ViewportTableRecord;

            //// Save the new objects to the database
            //acTrans.Commit();

        }

    }

    class Truss
    {
        public static Solid3d CreateTrussChord(double x1,
            double x2,
            double y1,
            double y2,
            double length,
            double rotateAngle,
            Vector3d rotateAxis,
            Vector3d chordVector,
            bool mirror = false)
        {           
            // Create polyline of chord profile
            Polyline chordPoly = new Polyline();
            chordPoly.AddVertexAt(0, new Point2d(0, 0), 0, 0, 0);
            chordPoly.AddVertexAt(1, new Point2d(x1, 0), 0, 0, 0);
            chordPoly.AddVertexAt(2, new Point2d(x1, y1), 0, 0, 0);
            chordPoly.AddVertexAt(3, new Point2d(x2, y1), 0, 0, 0);
            chordPoly.AddVertexAt(4, new Point2d(x2, y2), 0, 0, 0);
            chordPoly.AddVertexAt(5, new Point2d(0, y2), 0, 0, 0);
            chordPoly.Closed = true;

            // Create the 3D solid top chord(s)
            Solid3d chord = Utils.ExtrudePolyline(chordPoly, length);

            // Position the top chord
            chord.TransformBy(Matrix3d.Rotation(rotateAngle, rotateAxis, Point3d.Origin));
            chord.TransformBy(Matrix3d.Displacement(chordVector));
            if (mirror == true)
            {
                chord.TransformBy(Matrix3d.Mirroring(new Plane(Point3d.Origin, rotateAxis)));
            }

            return chord;
        }

        public static Solid3d TrussTopChord(Dictionary<string, double> data,
            string orientation,
            bool mirror = false)
        {

            // Variable stubs
            Vector3d rotateAxis;
            double rotateAngle;
            Vector3d chordVec;
            double cX1;
            double cX2;
            double cY1;
            double cY2;
                        
            // Calculate truss values
            if (orientation == "X-Axis")
            {
                rotateAxis = Vector3d.YAxis;
                rotateAngle = Math.PI / 2;
                chordVec = new Vector3d(0, data["rodDiameter"] / 2, data["chordWidth"] * 2);
                cX1 = data["chordWidth"];
                cX2 = data["chordThickness"];
                cY1 = data["chordThickness"];
                cY2 = data["chordWidth"];
            }
            else // assume y orientation
            {
                rotateAxis = Vector3d.XAxis;
                rotateAngle = - Math.PI / 2;
                chordVec = new Vector3d(data["rodDiameter"] / 2, 0, data["chordWidth"] * 2);
                cX1 = data["chordWidth"];
                cX2 = data["chordThickness"];
                cY1 = data["chordThickness"];
                cY2 = data["chordWidth"];
            }

            return CreateTrussChord(cX1, cX2, cY1, cY2, data["trussLength"], rotateAngle, 
                rotateAxis, chordVec, mirror: mirror);
        }

        public static Solid3d TrussEndChord(Dictionary<string, double> data,
            string placement = "near",
            string orientation = "Y-Axis",
            bool mirror = false)
        {

            // Variable stubs
            double chordLength;
            double chordLocation;
            Vector3d rotateAxis;
            double rotateAngle;
            Vector3d chordVec;
            double cX1;
            double cX2;
            double cY1;
            double cY2;

            // Calculate truss values
            if (placement == "far")
            {
                chordLength = data["chordFarEndLength"];
                chordLocation = data["trussLength"] - chordLength;
            }
            else
            {
                chordLength = data["chordNearEndLength"];
                chordLocation = 0;
            }

            if (orientation == "X-Axis")
            {
                rotateAxis = Vector3d.YAxis;
                rotateAngle = Math.PI / 2;
                chordVec = new Vector3d(chordLocation, data["rodDiameter"] / 2, 0);
                cX1 = -data["chordWidth"];
                cX2 = -data["chordThickness"];
                cY1 = data["chordThickness"];
                cY2 = data["chordWidth"];
            }
            else // assume y orientation
            {
                rotateAxis = Vector3d.XAxis;
                rotateAngle = -Math.PI / 2;
                chordVec = new Vector3d(data["rodDiameter"] / 2, chordLocation, 0);
                cX1 = data["chordWidth"];
                cX2 = data["chordThickness"];
                cY1 = -data["chordThickness"];
                cY2 = -data["chordWidth"];
            }

            return CreateTrussChord(cX1, cX2, cY1, cY2, chordLength, rotateAngle,
                rotateAxis, chordVec, mirror: mirror);
        }

        public static Solid3d TrussBottomChord(Dictionary<string, double> data,
            string orientation = "Y-Axis",
            bool mirror = false)
        {

            // Variable stubs
            Vector3d rotateAxis;
            double rotateAngle;
            Vector3d chordVec;
            double cX1;
            double cX2;
            double cY1;
            double cY2;

            // Calculate truss values
            if (!data.ContainsKey("bottomChordLength"))
            {
                data.Add("bottomChordLength", data["trussLength"]
                    - data["chordNearEndLength"]
                    - data["chordFarEndLength"]
                    - 2 * (data["trussHeight"] - 3 * data["chordWidth"]) * Math.Tan(data["rodEndAngle"])
                    + 2 * data["rodDiameter"] / Math.Cos(data["rodEndAngle"]));
            }

            if (orientation == "X-Axis")
            {
                rotateAxis = Vector3d.YAxis;
                rotateAngle = Math.PI / 2;
                chordVec = new Vector3d((data["trussLength"] - data["bottomChordLength"]) / 2,
                    data["rodDiameter"] / 2,
                    -(data["trussHeight"] - 2 * data["chordWidth"]));
                cX1 = -data["chordWidth"];
                cX2 = -data["chordThickness"];
                cY1 = data["chordThickness"];
                cY2 = data["chordWidth"];
            }
            else // assume y orientation
            {
                rotateAxis = Vector3d.XAxis;
                rotateAngle = -Math.PI / 2;
                chordVec = new Vector3d(data["rodDiameter"] / 2,
                    (data["trussLength"] - data["bottomChordLength"]) / 2,
                    -(data["trussHeight"] - 2 * data["chordWidth"]));
                cX1 = data["chordWidth"];
                cX2 = data["chordThickness"];
                cY1 = -data["chordThickness"];
                cY2 = -data["chordWidth"];
            }

            return CreateTrussChord(cX1, cX2, cY1, cY2, data["bottomChordLength"], rotateAngle,
                rotateAxis, chordVec, mirror: mirror);
        }

        private static Solid3d CreateTrussRod(double rodDiameter,
            double rodLength,
            Vector3d rodLocation,
            double rodAngle,
            Vector3d rotateAxis)
        {
            // Create rod
            Solid3d rod = new Solid3d();

            try
            {
                // Draw rod
                rod.CreateFrustum(rodLength, rodDiameter / 2, rodDiameter / 2, rodDiameter / 2);

                // Rotate rod
                rod.TransformBy(Matrix3d.Rotation(rodAngle, rotateAxis, Point3d.Origin));
                rod.TransformBy(Matrix3d.Displacement(rodLocation));
            }
            catch (Autodesk.AutoCAD.Runtime.Exception Ex)
            {
                Application.ShowAlertDialog("There was a problem creating the rod:\n" +
                    "Rod Length = " + rodLength + "\n" +
                    "Rod Diameter = " + rodDiameter + "\n" +
                    Ex.Message + "\n" +
                    Ex.StackTrace);
            }

            return rod;
        }

        private static Tuple<double,double> rodIntersectPoint(Solid3d rod,
            double rodDiameter,
            double rodAngle,
            string orient,
            string location)
        {
            // Declare variable stubs
            double dPoint = 0;
            double hPoint = 0;

            // Choose distance point
            if (location == "near")
            {
                if (orient == "X-Axis")
                {
                    dPoint = rod.Bounds.Value.MaxPoint.X;
                }
                else if (orient == "Y-Axis")
                {
                    dPoint = rod.Bounds.Value.MaxPoint.Y;
                }
            }
            else if (location == "far")
            {
                if (orient == "X-Axis")
                {
                    dPoint = rod.Bounds.Value.MinPoint.X;
                }
                else if (orient == "Y-Axis")
                {
                    dPoint = rod.Bounds.Value.MinPoint.Y;
                }
            }

            // Calculate Z (height) coordinate
            hPoint = rod.Bounds.Value.MinPoint.Z + rodDiameter * Math.Sin(Math.Abs(rodAngle));

            return Tuple.Create(dPoint, hPoint);
            
        }

        public static void TrussMidRods(BlockTableRecord btr,
            Dictionary<string, double> data,
            string orient = "Y-Axis")
        {
            // Declare variable stubs
            Vector3d rodNearEndVec;
            Vector3d rodFarEndVec;
            Vector3d rotateAxis;
            double rotateAngle;


            // Calculate mid rod length and horizontal location (does not change)
            data.Add("rodMidLength", (data["trussHeight"] - data["chordWidth"])
                / Math.Cos(data["rodMidAngle"]));
            double rodVerticalLocation = 2 * data["chordWidth"] - data["trussHeight"] / 2;

            // Calculate mid rod horizontal length (for loop)
            double rodHorizontalLength = data["rodMidLength"] * Math.Sin(data["rodMidAngle"])
                + data["rodDiameter"] / Math.Cos(data["rodMidAngle"]);

            // Calculate near and far mid rod locations
            double rodNearHorizontalLocation = (data["rodEndNearMaxPtD"] + data["rodDiameter"] / (2 * Math.Cos(data["rodMidAngle"])))
                - Math.Tan(-data["rodMidAngle"]) * (data["rodEndNearMaxPtH"] - rodVerticalLocation);
            double rodFarHorizontalLocation = (data["rodEndFarMinPtD"] - data["rodDiameter"] / (2 * Math.Cos(data["rodMidAngle"])))
                + Math.Tan(data["rodMidAngle"]) * (rodVerticalLocation - data["rodEndFarMinPtH"]);

            if (orient == "X-Axis")
            {
                rodNearEndVec = new Vector3d(rodNearHorizontalLocation,
                    0,
                    rodVerticalLocation);
                rodFarEndVec = new Vector3d(rodFarHorizontalLocation,
                    0,
                    rodVerticalLocation);
                rotateAxis = Vector3d.YAxis;
                rotateAngle = -data["rodMidAngle"];
            }
            else
            {
                rodNearEndVec = new Vector3d(0,
                    rodNearHorizontalLocation,
                    rodVerticalLocation);
                rodFarEndVec = new Vector3d(0,
                    rodFarHorizontalLocation,
                    rodVerticalLocation);
                rotateAxis = Vector3d.XAxis;
                rotateAngle = data["rodMidAngle"];
            }

            // Create mid end rods
            Solid3d rodNear = CreateTrussRod(data["rodDiameter"], data["rodMidLength"],
                rodNearEndVec, rotateAngle, rotateAxis);
            Solid3d rodFar = CreateTrussRod(data["rodDiameter"], data["rodMidLength"],
                rodFarEndVec, -rotateAngle, rotateAxis);

            // Add mid end rods to block
            btr.AppendEntity(rodNear);
            btr.AppendEntity(rodFar);

            // Loop to create remaining rods
            double space = rodIntersectPoint(rodFar, data["rodDiameter"], rotateAngle, orient, "far").Item1
                - rodIntersectPoint(rodNear, data["rodDiameter"], rotateAngle, orient, "near").Item1;
            while (space >= rodHorizontalLength)
            {
                // Switch rotate angle (alternates rod angles)
                rotateAngle = -rotateAngle;

                // Calculate near and far rod locations
                rodNearHorizontalLocation = rodIntersectPoint(rodNear, data["rodDiameter"], rotateAngle, orient, "near").Item1
                    + rodHorizontalLength / 2;
                rodFarHorizontalLocation = rodIntersectPoint(rodFar, data["rodDiameter"], rotateAngle, orient, "far").Item1
                    - rodHorizontalLength / 2;

                if (orient == "X-Axis")
                {
                    rodNearEndVec = new Vector3d(rodNearHorizontalLocation,
                        0,
                        rodVerticalLocation);
                    rodFarEndVec = new Vector3d(rodFarHorizontalLocation,
                        0,
                        rodVerticalLocation);
                }
                else
                {
                    rodNearEndVec = new Vector3d(0,
                        rodNearHorizontalLocation,
                        rodVerticalLocation);
                    rodFarEndVec = new Vector3d(0,
                        rodFarHorizontalLocation,
                        rodVerticalLocation);
                }

                // Create near mid rod
                rodNear = CreateTrussRod(data["rodDiameter"], data["rodMidLength"],
                    rodNearEndVec, rotateAngle, rotateAxis);
                btr.AppendEntity(rodNear);
                space -= rodHorizontalLength;

                // Re-check available space (break if no more space)
                if (space < rodHorizontalLength) { break; }

                // Create far mid rod (if space allows)
                rodFar = CreateTrussRod(data["rodDiameter"], data["rodMidLength"],
                    rodFarEndVec, -rotateAngle, rotateAxis);
                btr.AppendEntity(rodFar);
                space -= rodHorizontalLength;
            }
        }

        public static void TrussEndRods(BlockTableRecord btr,
            Dictionary<string, double> data,
            string orient = "Y-Axis")
        {
            // Declare variable stubs
            Vector3d rodNearEndVec;
            Vector3d rodFarEndVec;
            Vector3d rotateAxis;
            double rotateAngle;

            // Calculate end rod locations
            data.Add("rodEndLength", (data["trussHeight"] - 2 * data["chordWidth"])
                / Math.Cos(data["rodEndAngle"]));
            double rodHorizontalLocation = ((data["trussHeight"] - 3 * data["chordWidth"]) * Math.Tan(data["rodEndAngle"])
                - data["rodDiameter"] / Math.Cos(data["rodEndAngle"])) / 2;
            double rodVerticalLocation = data["chordWidth"] - (data["trussHeight"] - data["chordWidth"]) / 2;

            if (orient == "X-Axis")
            {
                rodNearEndVec = new Vector3d(data["chordNearEndLength"] + rodHorizontalLocation,
                    0,
                    rodVerticalLocation);
                rodFarEndVec = new Vector3d(data["trussLength"] - (data["chordFarEndLength"] + rodHorizontalLocation),
                    0,
                    rodVerticalLocation);
                rotateAxis = Vector3d.YAxis;
                rotateAngle = -data["rodEndAngle"];
            }
            else
            {
                rodNearEndVec = new Vector3d(0,
                    data["chordNearEndLength"] + rodHorizontalLocation,
                    rodVerticalLocation);
                rodFarEndVec = new Vector3d(0,
                    data["trussLength"] - (data["chordFarEndLength"] + rodHorizontalLocation),
                    rodVerticalLocation);
                rotateAxis = Vector3d.XAxis;
                rotateAngle = data["rodEndAngle"];
            }

            Solid3d rodNear = CreateTrussRod(data["rodDiameter"], data["rodEndLength"],
                rodNearEndVec, rotateAngle, rotateAxis);
            Solid3d rodFar = CreateTrussRod(data["rodDiameter"], data["rodEndLength"],
                rodFarEndVec, -rotateAngle, rotateAxis);

            // Add end rods to block
            btr.AppendEntity(rodNear);
            btr.AppendEntity(rodFar);

            // Calculate and record limits of end rods
            data.Add("rodEndNearMaxPtD", rodIntersectPoint(rodNear, data["rodDiameter"],
                rotateAngle, orient, "near").Item1);
            data.Add("rodEndNearMaxPtH", rodIntersectPoint(rodNear, data["rodDiameter"],
                rotateAngle, orient, "near").Item2);
            data.Add("rodEndFarMinPtD", rodIntersectPoint(rodFar, data["rodDiameter"],
                rotateAngle, orient, "far").Item1);
            data.Add("rodEndFarMinPtH", rodIntersectPoint(rodFar, data["rodDiameter"],
                rotateAngle, orient, "far").Item2);
        }

        public static void CreateTruss(Database acCurDb,
            Dictionary<string, double> data,
            string orient = "X Axis")
        {
            // Some error checking
            if (!data.ContainsKey("trussLength"))
            { Application.ShowAlertDialog("Error: Truss Length not found."); return; } 
            else if (!data.ContainsKey("trussHeight"))
            { Application.ShowAlertDialog("Error: Truss Height not found."); return; }
            else if (!data.ContainsKey("chordWidth"))
            { Application.ShowAlertDialog("Error: Chord width not found."); return; }

            // Truss block name
            string tName = "Truss - " + data["trussLength"] + "x" + data["trussHeight"];
                
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // Open the Block table record for read
                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                                             OpenMode.ForRead) as BlockTable;

                // Blank objectid reference
                ObjectId fId = ObjectId.Null;

                if (acBlkTbl.Has(tName))
                {
                    // Retrieve object id
                    fId = acBlkTbl[tName];
                    //!!! Add options here to save or overwrite
                }
                else
                {
                    using (BlockTableRecord acBlkTblRec = new BlockTableRecord())
                    {
                        acBlkTblRec.Name = tName;

                        // Set the insertion point for the block
                        acBlkTblRec.Origin = Point3d.Origin;

                        // Create top chords
                        Solid3d cTop1 = TrussTopChord(data, orient);
                        Solid3d cTop2 = TrussTopChord(data, orient, mirror: true);

                        // Add top chords to block
                        acBlkTblRec.AppendEntity(cTop1);
                        acBlkTblRec.AppendEntity(cTop2);

                        // Create end chords
                        Solid3d cEndNear1 = TrussEndChord(data, "near", orient);
                        Solid3d cEndNear2 = TrussEndChord(data, "near", orient, mirror: true);
                        Solid3d cEndFar1 = TrussEndChord(data, "far", orient);
                        Solid3d cEndFar2 = TrussEndChord(data, "far", orient, mirror: true);

                        // Add top chords to block
                        acBlkTblRec.AppendEntity(cEndNear1);
                        acBlkTblRec.AppendEntity(cEndNear2);
                        acBlkTblRec.AppendEntity(cEndFar1);
                        acBlkTblRec.AppendEntity(cEndFar2);

                        // Create bottom chords
                        Solid3d cBottom1 = TrussBottomChord(data, orientation: orient);
                        Solid3d cBottom2 = TrussBottomChord(data, orientation: orient, mirror: true);

                        // Add top chords to block
                        acBlkTblRec.AppendEntity(cBottom1);
                        acBlkTblRec.AppendEntity(cBottom2);

                        // Create and add truss end rods
                        TrussEndRods(acBlkTblRec, data, orient);

                        // Create and add truss end rods
                        TrussMidRods(acBlkTblRec, data, orient);

                        /// Add Block to Block Table and close Transaction
                        acBlkTbl.UpgradeOpen();
                        acBlkTbl.Add(acBlkTblRec);
                        acTrans.AddNewlyCreatedDBObject(acBlkTblRec, true);

                    }

                }

                // Save the new object to the database
                acTrans.Commit();

            }

        }

        [CommandMethod("TrussPrompt")]
        public static void TrussPrompt()
        {
            // Initialize data dictionary
            Dictionary<string, double> trussData = new Dictionary<string, double>();

            // Initialize prompt check variable
            bool viable = false;

            // Some default values (could be arguments)
            trussData.Add("trussLength", 240); // truss length
            trussData.Add("trussHeight", 1); // truss height
            trussData.Add("chordWidth", 1.25); // chord width
            trussData.Add("chordThickness", 0.125); // chord thickness
            trussData.Add("chordNearEndLength", 8); // chord end piece length (sits on column/beam)
            trussData.Add("chordFarEndLength", 8); // chord end piece length (sits on column/beam)
            trussData.Add("rodDiameter", 0.875); // truss rod diameter
            double rEndAngle = 60; // end rod angle (from vertical)
            double rMidAngle = 35; // middle rod angles (from vertical)
            trussData.Add("rodEndAngle", rEndAngle.ToRadians());
            trussData.Add("rodMidAngle", rMidAngle.ToRadians());
            string orientation = "X-Axis"; // truss orientation

            // Declare prompt variables
            PromptDoubleResult tLengthRes;
            PromptDoubleResult tHeightRes;
            PromptResult tOthersRes;
            PromptResult tOrientRes;
            PromptDoubleResult cWidthRes;
            PromptDistanceOptions tLengthOpts = new PromptDistanceOptions("");
            PromptDistanceOptions tHeightOpts = new PromptDistanceOptions("");
            PromptKeywordOptions tOthersOpts = new PromptKeywordOptions("");
            PromptKeywordOptions tOrientOpts = new PromptKeywordOptions("");
            PromptDistanceOptions cWidthOpts = new PromptDistanceOptions("");

            // Get the current document and database, and start a transaction
            // !!! Move this outside function
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            // Prepare prompt for the truss length
            tLengthOpts.Message = "\nEnter the truss length: ";
            tLengthOpts.DefaultValue = 240;

            // Prepare prompt for the truss height
            tHeightOpts.Message = "\nEnter the truss height: ";
            tHeightOpts.DefaultValue = 18;

            // Prepare prompt for other options
            tOthersOpts.Message = "\nTruss Options: ";
            tOthersOpts.Keywords.Add("Orientation");
            tOthersOpts.Keywords.Add("ChordWidth");
            //tOthersOpts.Keywords.Add("NearChordLength");
            //tOthersOpts.Keywords.Add("FarChordLength");
            //tOthersOpts.Keywords.Add("NearTab");
            //tOthersOpts.Keywords.Add("FarTab");
            //tOthersOpts.Keywords.Add("EndRodAngle");
            //tOthersOpts.Keywords.Add("MiddleRodAngle");
            tOthersOpts.AllowArbitraryInput = false;
            tOthersOpts.AllowNone = true;

            // Prepare prompt for the truss orientation
            tOrientOpts.Message = "\nEnter truss orientation: ";
            tOrientOpts.Keywords.Add("X-Axis");
            tOrientOpts.Keywords.Add("Y-Axis");
            tOrientOpts.Keywords.Default = "X-Axis";
            tOrientOpts.AllowArbitraryInput = false;

            // Prepare prompt for the chord width
            cWidthOpts.Message = "\nEnter the chord width: ";
            cWidthOpts.DefaultValue = 1.25;

            while (viable == false)
            {

                // Prompt for the truss length
                tLengthRes = acDoc.Editor.GetDistance(tLengthOpts);
                trussData["trussLength"] = tLengthRes.Value;
                if (tLengthRes.Status == PromptStatus.Cancel) return;

                // Prompt for the truss height
                tHeightRes = acDoc.Editor.GetDistance(tHeightOpts);
                trussData["trussHeight"] = tHeightRes.Value;
                if (tHeightRes.Status == PromptStatus.Cancel) return;

                // Prompt for other options
                tOthersRes = acDoc.Editor.GetKeywords(tOthersOpts);

                // Exit if the user presses ESC or cancels the command
                if (tOthersRes.Status == PromptStatus.Cancel) return;

                while (tOthersRes.Status == PromptStatus.OK)
                {
                    Application.ShowAlertDialog("Keyword: " + tOthersRes.StringResult);

                    switch (tOthersRes.StringResult)
                    {
                        case "Orientation":
                            tOrientRes = acDoc.Editor.GetKeywords(tOrientOpts);
                            orientation = tOrientRes.StringResult;
                            if (tOrientRes.Status != PromptStatus.OK) return;
                            break;
                        case "ChordWidth":
                            cWidthRes = acDoc.Editor.GetDistance(cWidthOpts);
                            trussData["chordWidth"] = cWidthRes.Value;
                            if (cWidthRes.Status == PromptStatus.Cancel) return;
                            break;
                        default:
                            Application.ShowAlertDialog("Invalid Keyword");
                            break;
                    }

                    // Re-prompt for keywords
                    tOthersRes = acDoc.Editor.GetKeywords(tOthersOpts);

                    // Exit if the user presses ESC or cancels the command
                    if (tOthersRes.Status == PromptStatus.Cancel) return;
                }

                // Check viability of truss length
                double rodEndLengthHorizontal = (trussData["trussHeight"] - 2 * trussData["chordWidth"])
                    / Math.Cos(trussData["rodEndAngle"]);
                double chordBreakHorizontal = ((trussData["trussHeight"] - 3 * trussData["chordWidth"])
                    * Math.Tan(trussData["rodEndAngle"]) - trussData["rodDiameter"]
                    / Math.Cos(trussData["rodEndAngle"]));
                double rodMidWidthHorizontal = trussData["rodDiameter"] / (2 * Math.Cos(trussData["rodMidAngle"]));

                if (trussData["trussLength"] < trussData["chordNearEndLength"] + trussData["chordFarEndLength"]
                    + (chordBreakHorizontal + rodEndLengthHorizontal) + rodMidWidthHorizontal)
                {
                    // Display error
                    Application.ShowAlertDialog("Warning: Truss length is too short\n" +
                        "Must be greater than " + 
                        (trussData["chordNearEndLength"] + trussData["chordFarEndLength"]
                        + (chordBreakHorizontal + rodEndLengthHorizontal) + rodMidWidthHorizontal)
                        + " inches");
                }
                else if (trussData["trussHeight"] <= trussData["chordWidth"] * 4)
                {
                    // Display error
                    Application.ShowAlertDialog("Warning: Truss height is too short.");
                }
                else
                {
                    viable = true;
                }

            }
            
            CreateTruss(acCurDb, trussData, orient: orientation);
        }

    }
    
}
