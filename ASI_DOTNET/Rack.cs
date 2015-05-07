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

[assembly: CommandClass(typeof(ASI_DOTNET.Frame))]
[assembly: CommandClass(typeof(ASI_DOTNET.Beam))]

namespace ASI_DOTNET
{
    class Frame
    {
        // Auto-impl class properties
        private Database db;
        public double height { get; private set; }
        public double width { get; private set; }
        public double diameter { get; private set; }
        public string name { get; private set; }
        public string layerName { get; private set; }
        public ObjectId id { get; private set; }

        // Public constructor
        public Frame (Database db,
            double height,
            double width,
            double diameter)
        {
            this.db = db;
            this.height = height;
            this.width = width;
            this.diameter = diameter;
            this.name = "Rack Frame - " + height + "x" + width;

            // Create beam layer (if necessary)
            this.layerName = "2D-Rack-Frame";
            Color layerColor = Utils.ChooseColor("teal");
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

                        // Add frame posts to the block table record
                        for (double dist = diameter; dist <= width; dist += width - diameter)
                        {
                            Solid3d Post = framePost(height, diameter, dist);
                            acBlkTblRec.AppendEntity(Post);
                        }

                        // Calculate cross-bracing
                        Dictionary<string, double> braceData = calcBraces(height, diameter);

                        // Add horizontal cross-braces
                        horizontalBraces(acBlkTblRec, height, diameter, braceData);

                        // Add angled cross-braces to the block
                        angleBraces(acBlkTblRec, height, diameter, braceData);

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

        // Build frame post
        private Solid3d framePost (double height,
            double diameter,
            double distance)
        {
            Solid3d Post = new Solid3d();
            Post.SetDatabaseDefaults();

            // Create the 3D solid frame post
            Post.CreateBox(diameter, diameter, height);

            // Find location for post
            Point3d pMid = new Point3d(diameter / 2, distance - diameter / 2, height / 2);
            Vector3d pVec = Point3d.Origin.GetVectorTo(pMid);

            // Position the post
            Post.TransformBy(Matrix3d.Displacement(pVec));

            // Set block object properties
            Utils.SetBlockObjectProperties(Post);

            return Post;
        }

        private Dictionary<string, double> calcBraces(double height,
            double diameter)
        {
            // Initialize dictionary
            Dictionary<string, double> data = new Dictionary<string, double>();

            // Calculate cross-bracing
            double bNum;
            double bSpace = 0;
            double bAngle = Math.PI;
            double bSize = 1;
            double bEndSpace = 8;

            for (bNum = 1; bNum <= Math.Floor((height - bEndSpace) / 12); bNum++)
            {
                bSpace = Math.Floor(((height - bEndSpace) - (bNum + 1) * bSize) / bNum);
                bAngle = Math.Atan2(bSpace - bSize, width - (2 * diameter));
                if ((bAngle < Math.PI / 3))
                {
                    break;
                }
            }

            data.Add("num", bNum);
            data.Add("size", bSize);
            data.Add("space", bSpace);
            data.Add("angle", bAngle);
            data.Add("leftover", height - ((bNum + 1) * bSize) - (bNum * bSpace));

            return data;
        }

        // Build frame post
        private void horizontalBraces(BlockTableRecord btr,
            double height,
            double diameter,
            Dictionary<string, double> data)
        {

            /// Add horizontal cross-braces to the block
            for (int i = 1; i <= Convert.ToInt32(data["num"]) + 1; i++)
            {
                Solid3d brace = new Solid3d();
                brace.SetDatabaseDefaults();

                // Create the 3D solid horizontal cross brace
                brace.CreateBox(data["size"], width - 2 * diameter, data["size"]);

                // Find location for brace
                Point3d bMid = new Point3d(diameter / 2, width / 2,
                    data["leftover"] / 2 + ((data["space"] + data["size"]) * (i - 1)) + data["size"] / 2);
                Vector3d bVec = Point3d.Origin.GetVectorTo(bMid);

                // Position the brace
                brace.TransformBy(Matrix3d.Displacement(bVec));

                // Set block object properties
                Utils.SetBlockObjectProperties(brace);

                btr.AppendEntity(brace);
            }
        }

        private void angleBraces(BlockTableRecord btr,
            double height,
            double diameter,
            Dictionary<string, double> data)
        {
            // Create polyline of brace profile and rotate to correct orienation
            Polyline bPoly = new Polyline();
            bPoly.AddVertexAt(0, new Point2d(0, 0), 0, 0, 0);
            bPoly.AddVertexAt(1, new Point2d(data["size"], 0), 0, 0, 0);
            bPoly.AddVertexAt(2, new Point2d(data["size"], data["size"]), 0, 0, 0);
            bPoly.AddVertexAt(3, new Point2d(0, data["size"]), 0, 0, 0);
            bPoly.Closed = true;
            bPoly.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.XAxis, Point3d.Origin));

            /// Create sweep path
            Line bPath = new Line(Point3d.Origin, new Point3d(0, width - 2 * diameter, data["space"] - data["size"]));

            /// Create swept cross braces
            for (int i = 1; i <= Convert.ToInt32(data["num"]); i++)
            {
                // Create swept brace at origin
                Solid3d aBrace = Utils.SweepPolylineOverLine(bPoly, bPath);

                // Calculate location
                double aBxLoc = (diameter / 2) - (data["size"] / 2);
                double aByLoc = diameter;
                double aBzLoc = data["leftover"] / 2 + data["size"] * i + data["space"] * (i - 1);

                /// Even braces get rotated 180 deg and different x and y coordinates
                if (i % 2 == 0)
                {
                    aBrace.TransformBy(Matrix3d.Rotation(Math.PI, Vector3d.ZAxis, Point3d.Origin));
                    aBxLoc = (diameter / 2) + (data["size"] / 2);
                    aByLoc = width - diameter;
                }

                // Create location for brace
                Point3d aMid = new Point3d(aBxLoc, aByLoc, aBzLoc);
                Vector3d aVec = Point3d.Origin.GetVectorTo(aMid);

                // Position the brace
                aBrace.TransformBy(Matrix3d.Displacement(aVec));

                // Set block object properties
                Utils.SetBlockObjectProperties(aBrace);

                // Add brace to block
                btr.AppendEntity(aBrace);
            }
        }

        [CommandMethod("FramePrompt")]
        public static void FramePrompt()
        {
            // Get the current document and database, and start a transaction
            // !!! Move this outside function
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            // Prompt for the frame height
            PromptDoubleResult heightRes;
            PromptDistanceOptions heightOpts = new PromptDistanceOptions("");
            heightOpts.Message = "\nEnter the frame height: ";
            heightOpts.DefaultValue = 96;
            heightRes = acDoc.Editor.GetDistance(heightOpts);
            double height = heightRes.Value;

            // Exit if the user presses ESC or cancels the command
            if (heightRes.Status != PromptStatus.OK) return;

            // Prompt for the frame width
            PromptDoubleResult widthRes;
            PromptDistanceOptions widthOpts = new PromptDistanceOptions("");
            widthOpts.Message = "\nEnter the frame width: ";
            widthOpts.DefaultValue = 36;
            widthRes = acDoc.Editor.GetDistance(widthOpts);
            double width = widthRes.Value;

            // Exit if the user presses ESC or cancels the command
            if (widthRes.Status != PromptStatus.OK) return;

            // Prompt for the frame diameter
            PromptDoubleResult diameterRes;
            PromptDistanceOptions diameterOpts = new PromptDistanceOptions("");
            diameterOpts.Message = "\nEnter the frame width: ";
            diameterOpts.DefaultValue = 3.0;
            diameterRes = acDoc.Editor.GetDistance(diameterOpts);
            double diameter = diameterRes.Value;

            // Exit if the user presses ESC or cancels the command
            if (diameterRes.Status != PromptStatus.OK) return;

            // Create Frame block
            Frame rackFrame = new Frame(db: acCurDb,
                height: height,
                width: width,
                diameter: diameter);

            rackFrame.Build();
        }

    }

    class Beam
    {
        // Auto-impl class properties
        private Database db;
        public double length { get; private set; }
        public double height { get; private set; }
        public double width { get; private set; }
        public string name { get; private set; }
        public string orientation { get; private set; }
        public string style { get; private set; }
        public string layerName { get; private set; }
        public ObjectId id { get; private set; }

        // Public constructor
        public Beam (Database db,
            double length,
            double height = 3,
            double width = 2,
            string orient = "X-Axis",
            string style = "Step")
        {
            this.db = db;
            this.length = length;
            this.height = height;
            this.width = width;
            this.orientation = orient;
            this.style = style;
            this.name = "Rack Beam - " + length + "x" + height;

            // Create beam layer (if necessary)
            this.layerName = "2D-Rack-Beam";
            Color layerColor = Utils.ChooseColor("blue");
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

                        // Create beam profile
                        Polyline beamPoly = beamProfile(height,
                            width,
                            style: style,
                            orient: orientation);

                        // Create beam path
                        Line beamLine = beamPath(length, orientation);

                        // Create beam
                        Solid3d beamSolid = Utils.SweepPolylineOverLine(beamPoly, beamLine);

                        // Set block properties
                        Utils.SetBlockObjectProperties(beamSolid);

                        // Add entity to Block Table Record
                        acBlkTblRec.AppendEntity(beamSolid);

                        /// Add Block to Block Table and Transaction
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

        // Draw beam profile
        private Polyline beamProfile (double height,
            double width,
            string style = "Step",
            double stepSize = 0.75,
            string orient = "X-Axis")
        {
            int i = 0;

            /// Create polyline of brace profile and rotate to correct orienation
            Polyline poly = new Polyline();
            poly.AddVertexAt(i, new Point2d(0, 0), 0, 0, 0);
            poly.AddVertexAt(i++, new Point2d(-height, 0), 0, 0, 0);

            // Add vertices for beam shape
            if (style == "Step")
            {
                poly.AddVertexAt(i++, new Point2d(-height, width - stepSize), 0, 0, 0);
                poly.AddVertexAt(i++, new Point2d(-height + stepSize, width - stepSize), 0, 0, 0);
                poly.AddVertexAt(i++, new Point2d(-height + stepSize, width), 0, 0, 0);
            }
            else if (style == "Box")
            {
                poly.AddVertexAt(i++, new Point2d(-height, width), 0, 0, 0);
            }

            poly.AddVertexAt(i++, new Point2d(0, width), 0, 0, 0);
            poly.Closed = true;
            poly.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.YAxis, Point3d.Origin));

            // Rotate beam profile if necessary
            if (orient == "Y-Axis")
            {
                poly.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, Point3d.Origin));
            }

            return poly;
        }

        private Line beamPath (double length,
            string orient)
        {
            Line path;

            if (orient == "Y-Axis")
            {
                path = new Line(Point3d.Origin, new Point3d(0, length, 0));
            }
            else // X-Axis (default)
            {
                path = new Line(Point3d.Origin, new Point3d(length, 0, 0));
            }

            return path;
        }

        [CommandMethod("BeamPrompt")]
        public static void BeamPrompt()
        {
            // Get the current document and database, and start a transaction
            // !!! Move this outside function
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            // Prepare prompt for the beam length
            PromptDoubleResult lengthRes;
            PromptDistanceOptions lengthOpts = new PromptDistanceOptions("");
            lengthOpts.Message = "\nEnter the beam length: ";
            lengthOpts.DefaultValue = 96;

            // Prepare prompt for the beam height
            PromptDoubleResult heightRes;
            PromptDistanceOptions heightOpts = new PromptDistanceOptions("");
            heightOpts.Message = "\nEnter the beam height: ";
            heightOpts.DefaultValue = 3;

            // Prepare prompt for the beam width
            PromptDoubleResult widthRes;
            PromptDistanceOptions widthOpts = new PromptDistanceOptions("");
            widthOpts.Message = "\nEnter the beam width: ";
            widthOpts.DefaultValue = 2;

            // Prepare prompt for other options
            PromptResult bOthersRes;
            PromptKeywordOptions bOthersOpts = new PromptKeywordOptions("");
            bOthersOpts.Message = "\nBeam Options: ";
            bOthersOpts.Keywords.Add("Orientation");
            bOthersOpts.Keywords.Add("Style");
            bOthersOpts.AllowArbitraryInput = false;
            bOthersOpts.AllowNone = true;

            // Prepare prompt for the beam orientation
            string orientation = "X-Axis";
            PromptResult bOrientRes;
            PromptKeywordOptions bOrientOpts = new PromptKeywordOptions("");
            bOrientOpts.Message = "\nEnter beam orientation: ";
            bOrientOpts.Keywords.Add("X-Axis");
            bOrientOpts.Keywords.Add("Y-Axis");
            bOrientOpts.Keywords.Default = "X-Axis";
            bOrientOpts.AllowArbitraryInput = false;

            // Prepare prompt for the beam style
            string style = "Step";
            PromptResult bStyleRes;
            PromptKeywordOptions bStyleOpts = new PromptKeywordOptions("");
            bStyleOpts.Message = "\nEnter beam style: ";
            bStyleOpts.Keywords.Add("Step");
            bStyleOpts.Keywords.Add("Box");
            bStyleOpts.Keywords.Default = "Step";
            bStyleOpts.AllowArbitraryInput = false;

            // Prompt for beam length
            lengthRes = acDoc.Editor.GetDistance(lengthOpts);
            if (lengthRes.Status != PromptStatus.OK) return;
            double length = lengthRes.Value;

            // Prompt for beam height
            heightRes = acDoc.Editor.GetDistance(heightOpts);
            if (heightRes.Status != PromptStatus.OK) return;
            double height = heightRes.Value;

            // Prompt for beam width
            widthRes = acDoc.Editor.GetDistance(widthOpts);
            if (widthRes.Status != PromptStatus.OK) return;
            double width = widthRes.Value;
            
            // Prompt for other options
            bOthersRes = acDoc.Editor.GetKeywords(bOthersOpts);

            // Exit if the user presses ESC or cancels the command
            if (bOthersRes.Status == PromptStatus.Cancel) return;

            while (bOthersRes.Status == PromptStatus.OK)
            {
                switch (bOthersRes.StringResult)
                {
                    case "Orientation":
                        bOrientRes = acDoc.Editor.GetKeywords(bOrientOpts);
                        orientation = bOrientRes.StringResult;
                        if (bOrientRes.Status != PromptStatus.OK) return;
                        break;
                    case "Style":
                        bStyleRes = acDoc.Editor.GetKeywords(bStyleOpts);
                        style = bStyleRes.StringResult;
                        if (bStyleRes.Status == PromptStatus.Cancel) return;
                        break;
                    default:
                        Application.ShowAlertDialog("Invalid Keyword");
                        break;
                }

                // Re-prompt for keywords
                bOthersRes = acDoc.Editor.GetKeywords(bOthersOpts);

                // Exit if the user presses ESC or cancels the command
                if (bOthersRes.Status == PromptStatus.Cancel) return;
            }

            // Create beam
            Beam rackBeam = new Beam(db: acCurDb,
                length: length,
                height: height,
                width: width,
                orient: orientation,
                style: style);

            rackBeam.Build();
        }
                        
    }

}
