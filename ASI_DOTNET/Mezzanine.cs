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
using Autodesk.AutoCAD.Colors;

[assembly: CommandClass(typeof(ASI_DOTNET.Column))]
[assembly: CommandClass(typeof(ASI_DOTNET.Truss))]

namespace ASI_DOTNET
{
    class Column
    {
        // Auto-impl class properties
        private Database db;
        public double height { get; private set; }
        public double width { get; private set; }
        public double baseHeight { get; private set; }
        public double baseWidth { get; private set; }
        public string name { get; private set; }
        public string style { get; private set; }
        public string layerName { get; private set; }
        public ObjectId id { get; private set; }

        // Public constructor
        public Column(Database db,
            double height,
            double width,
            double baseWidth,
            double baseHeight = 0.75)
        {
            this.db = db;
            this.height = height;
            this.width = width;
            this.baseHeight = baseHeight;
            this.baseWidth = baseWidth;
            this.name = "Column - " + height + "in - " + width + "x" + baseWidth;

            // Create beam layer (if necessary)
            this.layerName = "2D-Mezz-Column";
            Color layerColor = Utils.ChooseColor("black");
            Utils.CreateLayer(db, layerName, layerColor);
        }

        public void Build()
        {
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                // Open the Block table for read
                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;

                if (acBlkTbl.Has(name))
                {
                    // Retrieve object id
                    this.id = acBlkTbl[name];
                }
                else
                {
                    // Create new block (record)
                    using (BlockTableRecord acBlkTblRec = new BlockTableRecord())
                    {
                        acBlkTblRec.Name = name;

                        // Set the insertion point for the block
                        acBlkTblRec.Origin = new Point3d(0, 0, 0);
                        
                        // Calculate baseplate locations
                        Point3d bMid = new Point3d(baseWidth / 2, baseWidth / 2, baseHeight / 2);
                        Vector3d bVec = Point3d.Origin.GetVectorTo(bMid);
                        var cBottom = height / 2 + baseHeight;
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
                        cMid = new Point3d(baseWidth / 2, baseWidth / 2, cBottom);
                        //}

                        cVec = Point3d.Origin.GetVectorTo(cMid);
                
                        // Create the 3D solid baseplate
                        Solid3d bPlate = new Solid3d();
                        bPlate.SetDatabaseDefaults();
                        bPlate.CreateBox(baseWidth, baseWidth, baseHeight);

                        // Position the baseplate 
                        bPlate.TransformBy(Matrix3d.Displacement(bVec));

                        // Set block object properties
                        Utils.SetBlockObjectProperties(bPlate);

                        // Add baseplate to block
                        acBlkTblRec.AppendEntity(bPlate);

                        // Create the 3D solid column
                        Solid3d column = new Solid3d();
                        column.SetDatabaseDefaults();
                        column.CreateBox(width, width, height - baseHeight);

                        // Position the column 
                        column.TransformBy(Matrix3d.Displacement(cVec));

                        // Set block object properties
                        Utils.SetBlockObjectProperties(column);

                        // Add column to block
                        acBlkTblRec.AppendEntity(column);

                        /// Add Block to Block Table and close Transaction
                        acBlkTbl.UpgradeOpen();
                        acBlkTbl.Add(acBlkTblRec);
                        acTrans.AddNewlyCreatedDBObject(acBlkTblRec, true);
                    }

                    // Set block id property
                    this.id = acBlkTbl[name];

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
            double bDiameter = 14;

            // Get the current document and database, and start a transaction
            // !!! Move this outside function
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            // Prompt for the column height
            PromptDoubleResult cHeightRes;
            PromptDistanceOptions cHeightOpts = new PromptDistanceOptions("");
            cHeightOpts.Message = "\nEnter the total column height: ";
            cHeightOpts.DefaultValue = 96;
            cHeightRes = acDoc.Editor.GetDistance(cHeightOpts);
            cHeight = cHeightRes.Value;

            // Exit if the user presses ESC or cancels the command
            if (cHeightRes.Status == PromptStatus.Cancel) return;

            // Prompt for the column width
            PromptDoubleResult cDiameterRes;
            PromptDistanceOptions cDiameterOpts = new PromptDistanceOptions("");
            cDiameterOpts.Message = "\nEnter the column diameter: ";
            cDiameterOpts.DefaultValue = 6;
            cDiameterRes = acDoc.Editor.GetDistance(cDiameterOpts);
            cDiameter = cDiameterRes.Value;

            // Exit if the user presses ESC or cancels the command
            if (cDiameterRes.Status == PromptStatus.Cancel) return;

            // Prompt for the base width
            PromptDoubleResult bDiameterRes;
            PromptDistanceOptions bDiameterOpts = new PromptDistanceOptions("");
            bDiameterOpts.Message = "\nEnter the baseplate diameter: ";
            bDiameterOpts.DefaultValue = 14;
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

            // Create Column block
            Column mezzColumn = new Column(db: acCurDb,
                height: cHeight,
                width: cDiameter,
                baseWidth: bDiameter);

            mezzColumn.Build();

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
        // Auto-impl class properties
        private Database db;
        public double trussLength { get; private set; }
        public double trussHeight { get; private set; }
        public double chordWidth { get; private set; }
        public double chordThickness { get; private set; }
        public double chordNearEndLength { get; private set; }
        public double chordFarEndLength { get; private set; }
        public double bottomChordLength { get; private set; }
        public double rodDiameter { get; private set; }
        public double rodEndLength { get; private set; }
        public double rodMidLength { get; private set; }
        public double rodEndAngle { get; private set; }
        public double rodMidAngle { get; private set; }
        public string name { get; private set; }
        public string orientation { get; private set; }
        public string layerName { get; private set; }
        public ObjectId id { get; private set; }

        // Public constructor
        public Truss(Database db,
            Dictionary<string, double> data,
            string orient = "X Axis")
        {
            this.db = db;
            this.trussLength = data["trussLength"];
            this.trussHeight = data["trussHeight"];
            this.chordWidth = data["chordWidth"];
            this.chordThickness = data["chordThickness"];
            this.chordNearEndLength = data["chordNearEndLength"];
            this.chordFarEndLength = data["chordFarEndLength"];
            this.rodDiameter = data["rodDiameter"];
            this.rodEndAngle = data["rodEndAngle"];
            this.rodMidAngle = data["rodMidAngle"];
            this.orientation = orient;
            this.name = "Truss - " + trussLength + "x" + trussHeight;

            // Calculate truss bottom length
            this.bottomChordLength = trussLength - chordNearEndLength - chordFarEndLength
                    - 2 * (trussHeight - 3 * chordWidth) * Math.Tan(rodEndAngle)
                    + 2 * rodDiameter / Math.Cos(rodEndAngle);

            // Calculate end rod length
            this.rodEndLength = (trussHeight - 2 * chordWidth) / Math.Cos(rodEndAngle);

            // Calculate mid rod length
            this.rodMidLength = (trussHeight - chordWidth) / Math.Cos(rodMidAngle);

            // Create beam layer (if necessary)
            this.layerName = "2D-Mezz-Truss";
            Color layerColor = Utils.ChooseColor("black");
            Utils.CreateLayer(db, layerName, layerColor);
        }

        public void Build()
        {
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                // Open the Block table for read
                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;

                if (acBlkTbl.Has(name))
                {
                    // Retrieve object id
                    this.id = acBlkTbl[name];
                }
                else
                {
                    using (BlockTableRecord acBlkTblRec = new BlockTableRecord())
                    {
                        acBlkTblRec.Name = name;

                        // Set the insertion point for the block
                        acBlkTblRec.Origin = new Point3d(0, 0, 0);

                        // Create top chords
                        Solid3d cTop1 = TrussTopChord(trussLength, chordWidth,
                            chordThickness, rodDiameter, orientation);
                        Solid3d cTop2 = TrussTopChord(trussLength, chordWidth,
                            chordThickness, rodDiameter, orientation, mirror: true);

                        // Add top chords to block
                        acBlkTblRec.AppendEntity(cTop1);
                        acBlkTblRec.AppendEntity(cTop2);

                        // Create end chords
                        Solid3d cEndNear1 = TrussEndChord(trussLength, chordNearEndLength,
                            chordWidth, chordThickness, rodDiameter, "near", orientation);
                        Solid3d cEndNear2 = TrussEndChord(trussLength, chordNearEndLength,
                            chordWidth, chordThickness, rodDiameter, "near", orientation, mirror: true);
                        Solid3d cEndFar1 = TrussEndChord(trussLength, chordFarEndLength,
                            chordWidth, chordThickness, rodDiameter, "far", orientation);
                        Solid3d cEndFar2 = TrussEndChord(trussLength, chordFarEndLength,
                            chordWidth, chordThickness, rodDiameter, "far", orientation, mirror: true);

                        // Add end chords to block
                        acBlkTblRec.AppendEntity(cEndNear1);
                        acBlkTblRec.AppendEntity(cEndNear2);
                        acBlkTblRec.AppendEntity(cEndFar1);
                        acBlkTblRec.AppendEntity(cEndFar2);

                        // Create bottom chords
                        Solid3d cBottom1 = TrussBottomChord(trussLength, trussHeight, chordWidth,
                            chordThickness, bottomChordLength, rodDiameter, orientation);
                        Solid3d cBottom2 = TrussBottomChord(trussLength, trussHeight, chordWidth,
                            chordThickness, bottomChordLength, rodDiameter, orientation, mirror: true);

                        // Add top chords to block
                        acBlkTblRec.AppendEntity(cBottom1);
                        acBlkTblRec.AppendEntity(cBottom2);

                        // Create truss end rods
                        Solid3d rEndNear = TrussEndRod(trussHeight, chordWidth, chordNearEndLength,
                            rodEndLength, rodDiameter, rodEndAngle, orientation);
                        Solid3d rEndFar = TrussEndRod(trussHeight, chordWidth, trussLength - chordFarEndLength,
                            rodEndLength, rodDiameter, -rodEndAngle, orientation);

                        // Add top chords to block
                        acBlkTblRec.AppendEntity(rEndNear);
                        acBlkTblRec.AppendEntity(rEndFar);

                        // Create and add truss end rods
                        TrussMidRods(acBlkTblRec, trussHeight, chordWidth,
                            rodDiameter, rodMidLength, rodMidAngle,
                            rodIntersectPoint(rEndNear, rodDiameter, rodEndAngle, orientation, "near"),
                            rodIntersectPoint(rEndFar, rodDiameter, rodEndAngle, orientation, "far"),
                            orientation);

                        /// Add Block to Block Table and close Transaction
                        acBlkTbl.UpgradeOpen();
                        acBlkTbl.Add(acBlkTblRec);
                        acTrans.AddNewlyCreatedDBObject(acBlkTblRec, true);

                    }

                    // Set block id property
                    this.id = acBlkTbl[name];

                }

                // Save the new object to the database
                acTrans.Commit();

            }
        }

        private static Solid3d CreateTrussChord(double x1,
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

            // Set block object properties
            Utils.SetBlockObjectProperties(chord);

            return chord;
        }

        private static Solid3d TrussTopChord(double trussLength,
            double chordWidth,
            double chordThickness,
            double rodDiameter,
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
                chordVec = new Vector3d(0, rodDiameter / 2, chordWidth * 2);
                cX1 = chordWidth;
                cX2 = chordThickness;
                cY1 = chordThickness;
                cY2 = chordWidth;
            }
            else // assume y orientation
            {
                rotateAxis = Vector3d.XAxis;
                rotateAngle = - Math.PI / 2;
                chordVec = new Vector3d(rodDiameter / 2, 0, chordWidth * 2);
                cX1 = chordWidth;
                cX2 = chordThickness;
                cY1 = chordThickness;
                cY2 = chordWidth;
            }

            return CreateTrussChord(cX1, cX2, cY1, cY2, trussLength, rotateAngle, 
                rotateAxis, chordVec, mirror: mirror);
        }

        private static Solid3d TrussEndChord(double trussLength,
            double chordLength,
            double chordWidth,
            double chordThickness,
            double rodDiameter,
            string placement = "near",
            string orientation = "Y-Axis",
            bool mirror = false)
        {

            // Variable stubs
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
                chordLocation = trussLength - chordLength;
            }
            else
            {
                chordLocation = 0;
            }

            if (orientation == "X-Axis")
            {
                rotateAxis = Vector3d.YAxis;
                rotateAngle = Math.PI / 2;
                chordVec = new Vector3d(chordLocation, rodDiameter / 2, 0);
                cX1 = -chordWidth;
                cX2 = -chordThickness;
                cY1 = chordThickness;
                cY2 = chordWidth;
            }
            else // assume y orientation
            {
                rotateAxis = Vector3d.XAxis;
                rotateAngle = -Math.PI / 2;
                chordVec = new Vector3d(rodDiameter / 2, chordLocation, 0);
                cX1 = chordWidth;
                cX2 = chordThickness;
                cY1 = -chordThickness;
                cY2 = -chordWidth;
            }

            return CreateTrussChord(cX1, cX2, cY1, cY2, chordLength, rotateAngle,
                rotateAxis, chordVec, mirror: mirror);
        }

        private static Solid3d TrussBottomChord(double trussLength,
            double trussHeight,
            double chordWidth,
            double chordThickness,
            double chordLength,
            double rodDiameter,
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

            if (orientation == "X-Axis")
            {
                rotateAxis = Vector3d.YAxis;
                rotateAngle = Math.PI / 2;
                chordVec = new Vector3d((trussLength - chordLength) / 2,
                    rodDiameter / 2,
                    -(trussHeight - 2 * chordWidth));
                cX1 = -chordWidth;
                cX2 = -chordThickness;
                cY1 = chordThickness;
                cY2 = chordWidth;
            }
            else // assume y orientation
            {
                rotateAxis = Vector3d.XAxis;
                rotateAngle = -Math.PI / 2;
                chordVec = new Vector3d(rodDiameter / 2,
                    (trussLength - chordLength) / 2,
                    -(trussHeight - 2 * chordWidth));
                cX1 = chordWidth;
                cX2 = chordThickness;
                cY1 = -chordThickness;
                cY2 = -chordWidth;
            }

            return CreateTrussChord(cX1, cX2, cY1, cY2, chordLength, rotateAngle,
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

            // Set block object properties
            Utils.SetBlockObjectProperties(rod);

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

        private static void TrussMidRods(BlockTableRecord btr,
            double trussHeight,
            double chordWidth,
            double rodDiameter,
            double rodLength,
            double rodAngle,
            Tuple<double, double> nearPt,
            Tuple<double, double> farPt,
            string orient = "Y-Axis")
        {
            // Declare variable stubs
            Vector3d rodNearEndVec;
            Vector3d rodFarEndVec;
            Vector3d rotateAxis;
            double rotateAngle;

            // Calculate mid rod vertical location
            double rodVerticalLocation = 2 * chordWidth - trussHeight / 2;

            // Calculate mid rod horizontal length (for loop)
            double rodHorizontalLength = rodLength * Math.Sin(rodAngle)
                + rodDiameter / Math.Cos(rodAngle);

            // Calculate near and far mid rod locations
            double rodNearHorizontalLocation = (nearPt.Item1 + rodDiameter / (2 * Math.Cos(rodAngle)))
                - Math.Tan(-rodAngle) * (nearPt.Item2 - rodVerticalLocation);
            double rodFarHorizontalLocation = (farPt.Item1 - rodDiameter / (2 * Math.Cos(rodAngle)))
                + Math.Tan(rodAngle) * (rodVerticalLocation - farPt.Item2);

            if (orient == "X-Axis")
            {
                rodNearEndVec = new Vector3d(rodNearHorizontalLocation,
                    0,
                    rodVerticalLocation);
                rodFarEndVec = new Vector3d(rodFarHorizontalLocation,
                    0,
                    rodVerticalLocation);
                rotateAxis = Vector3d.YAxis;
                rotateAngle = -rodAngle;
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
                rotateAngle = rodAngle;
            }

            // Create mid end rods
            Solid3d rodNear = CreateTrussRod(rodDiameter, rodLength,
                rodNearEndVec, rotateAngle, rotateAxis);
            Solid3d rodFar = CreateTrussRod(rodDiameter, rodLength,
                rodFarEndVec, -rotateAngle, rotateAxis);

            // Add mid end rods to block
            btr.AppendEntity(rodNear);
            btr.AppendEntity(rodFar);

            // Loop to create remaining rods
            double space = rodIntersectPoint(rodFar, rodDiameter, rotateAngle, orient, "far").Item1
                - rodIntersectPoint(rodNear, rodDiameter, rotateAngle, orient, "near").Item1;
            while (space >= rodHorizontalLength)
            {
                // Switch rotate angle (alternates rod angles)
                rotateAngle = -rotateAngle;

                // Calculate near and far rod locations
                rodNearHorizontalLocation = rodIntersectPoint(rodNear, rodDiameter, rotateAngle, orient, "near").Item1
                    + rodHorizontalLength / 2;
                rodFarHorizontalLocation = rodIntersectPoint(rodFar, rodDiameter, rotateAngle, orient, "far").Item1
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
                rodNear = CreateTrussRod(rodDiameter, rodLength,
                    rodNearEndVec, rotateAngle, rotateAxis);
                btr.AppendEntity(rodNear);
                space -= rodHorizontalLength;

                // Re-check available space (break if no more space)
                if (space < rodHorizontalLength) { break; }

                // Create far mid rod (if space allows)
                rodFar = CreateTrussRod(rodDiameter, rodLength,
                    rodFarEndVec, -rotateAngle, rotateAxis);
                btr.AppendEntity(rodFar);
                space -= rodHorizontalLength;
            }
        }

        private static Solid3d TrussEndRod(double trussHeight,
            double chordWidth,
            double rodStartLocation,
            double rodLength,
            double rodDiameter,
            double rodAngle,
            string orient = "Y-Axis")
        {
            // Declare variable stubs
            Vector3d rodVec;
            Vector3d rotateAxis;
            double rotateAngle;

            // Calculate end rod locations
            double rodHorizontalOffset = Math.Sign(rodAngle) * ((trussHeight - 3 * chordWidth)
                * Math.Tan(Math.Abs(rodAngle)) - rodDiameter / Math.Cos(Math.Abs(rodAngle))) / 2;
            double rodVerticalLocation = chordWidth - (trussHeight - chordWidth) / 2;
            double rodHorizontalLocation = rodStartLocation + rodHorizontalOffset;

            if (orient == "X-Axis")
            {
                rodVec = new Vector3d(rodHorizontalLocation, 0, rodVerticalLocation);
                rotateAxis = Vector3d.YAxis;
                rotateAngle = -rodAngle;
            }
            else
            {
                rodVec = new Vector3d(0, rodHorizontalLocation, rodVerticalLocation);
                rotateAxis = Vector3d.XAxis;
                rotateAngle = rodAngle;
            }

            Solid3d rod = CreateTrussRod(rodDiameter, rodLength,
                rodVec, rotateAngle, rotateAxis);

            return rod;
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

            // Create Truss block
            Truss mezzTruss = new Truss(db: acCurDb,
                data: trussData,
                orient: orientation);

            mezzTruss.Build();
        }

    }
    
}
