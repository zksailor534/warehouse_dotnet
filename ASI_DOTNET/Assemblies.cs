using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.PlottingServices;
using Autodesk.AutoCAD.Colors;

[assembly: CommandClass(typeof(ASI_DOTNET.Rack))]

namespace ASI_DOTNET
{
    class Rack
    {
        // Core class properties
        private Database db;
        private int deep;
        private int high;
        private int bays;
        private string type;

        // Get/Set for Core Properties
        public int rackCountDeep
        {
            get { return deep; }
            set
            {
                this.deep = value;
                OnRackPropertyChange();
            }
        }
        public int rackCountHigh
        {
            get { return high; }
            set
            {
                this.high = value;
                OnRackPropertyChange();
            }
        }
        public int rackCountBays
        {
            get { return bays; }
            set
            {
                this.bays = value;
                OnRackPropertyChange();
            }
        }
        public string rackType
        {
            get { return type; }
            set { this.type = value.ToLower(); }
        }

        // Auto-impl other class properties
        public Point3d rackOrigin { get; set; }
        public double frameHeight { get; set; }
        public double frameWidth { get; set; }
        public double frameDiameter { get; set; }
        public double beamLength { get; set; }
        public double beamHeight { get; set; }
        public double beamWidth { get; set; }
        public string beamType { get; set; }
        public double beamBottom { get; set; }
        public double spacerLength { get; set; }
        public double spacerHeight { get; set; }
        public double spacerWidth { get; set; }
        public string spacerType { get; set; }
        public double columnHeight { get; set; }
        public double columnWidth { get; set; }
        public double tieLength { get; set; }
        public double tieHeight { get; set; }
        public double tieWidth { get; set; }
        public string tieType { get; set; }
        private Point3d[,] frameLocations { get; set; }
        private Point3d[] columnLocations { get; set; }
        private Point3d[,,] beamLocations { get; set; }

        // Some other class variables
        private ObjectIdCollection objList = new ObjectIdCollection();
        private string beamOrientation;
        private string spacerOrientation;
        private string tieOrientation;

        // Public constructor
        public Rack(Database db,
            int rackCountDeep = 1,
            int rackCountHigh = 1,
            int rackCountBays = 1,
            string rackType = "pallet")
        {
            this.db = db;
            this.rackCountDeep = rackCountDeep;
            this.rackCountHigh = rackCountHigh;
            this.rackCountBays = rackCountBays;
            this.rackType = rackType;

            // Initialize locations and variables
            OnRackPropertyChange();

            // Set some defaults
            this.rackOrigin = Point3d.Origin;
            this.frameHeight = 96;
            this.frameWidth = 42;
            this.frameDiameter = 3;
            this.beamLength = 96;
            this.beamHeight = 3;
            this.beamWidth = 2;
            this.beamType = "Step";
            this.beamOrientation = "X-Axis";
            this.beamBottom = 0;
            this.spacerLength = 42;
            this.spacerHeight = 1;
            this.spacerWidth = 1;
            this.spacerType = "Box";
            this.spacerOrientation = "Y-Axis";
            this.columnHeight = 96;
            this.columnWidth = frameDiameter;
            this.tieLength = beamLength;
            this.tieHeight = frameDiameter;
            this.tieWidth  = frameDiameter;
            this.tieType = "Box";
            this.tieOrientation = "X-Axis";
        }

        // Call this function whenever core properties change
        private void OnRackPropertyChange()
        {
            // Frames
            this.frameLocations = new Point3d[bays + 1,
                Convert.ToInt32(Math.Ceiling(Convert.ToDouble(deep) / 2.0))];

            // Columns
            if (deep % 2 == 0) { this.columnLocations = new Point3d[bays + 1]; }
            else { this.columnLocations = new Point3d[0]; }

            // Beams
            this.beamLocations = new Point3d[bays,
                frameLocations.GetLength(1) * 2 + columnLocations.GetLength(0),
                high];
        }

        // Readable list of key properties
        public string RackDetails()
        {
            string details = "Rack:\n" +
                String.Format("\tRack Type: {0}\n", this.type) +
                String.Format("\t# of Bays: {0}\n", this.bays) +
                String.Format("\tRack Depth: {0}\n", this.deep) +
                String.Format("\t# of Rack Levels: {0}\n", this.high);

            return details;
        }

        public void Build()
        {
            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                // Open the Block table for read
                BlockTable acBlkTbl = acTrans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;

                // Open the Block table record Model space for write
                BlockTableRecord modelBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                    OpenMode.ForWrite) as BlockTableRecord;

                // Create frame block
                Frame rackFrame = new Frame(db,
                    height: frameHeight,
                    width: frameWidth,
                    diameter: frameDiameter);
                rackFrame.Build();

                // Create beam block
                Beam rackBeam = new Beam(db,
                    length: beamLength,
                    height: beamHeight,
                    width: beamWidth,
                    orient: beamOrientation,
                    style: beamType);
                rackBeam.Build();

                // Create spacer block
                Beam rackSpacer = new Beam(db,
                    length: spacerLength,
                    height: spacerHeight,
                    width: spacerWidth,
                    orient: spacerOrientation,
                    style: spacerType);
                rackSpacer.layerColor = Utils.ChooseColor("black");
                rackSpacer.layerName = "3D-Rack-Spacer";
                rackSpacer.Build();

                // Create column block
                Column rackColumn = new Column(db,
                    height: columnHeight,
                    width: columnWidth,
                    baseWidth: columnWidth);
                rackColumn.baseHeight = 0;
                rackColumn.Build();
                
                // Calculate frame locations and add to model space
                for (int x = 0; x < frameLocations.GetLength(0); x++)
                {
                    for (int y = 0; y < frameLocations.GetLength(1); y++)
                    {
                        frameLocations[x, y] = new Point3d(x * (rackFrame.diameter + rackBeam.length),
                            y * (rackFrame.width + rackSpacer.length), 0);
                        using (BlockReference frameRef = new BlockReference(frameLocations[x, y], rackFrame.id))
                        {
                            frameRef.Layer = rackFrame.layerName;
                            objList.Add(frameRef.Id);
                            modelBlkTblRec.AppendEntity(frameRef);
                            acTrans.AddNewlyCreatedDBObject(frameRef, true);
                        }
                    }
                }

                // Calculate column locations and add to model space
                for (int x = 0; x < columnLocations.GetLength(0); x++)
                {
                    columnLocations[x] = new Point3d(x * (rackFrame.diameter + rackBeam.length),
                        frameLocations[x, frameLocations.GetLength(1) - 1].Y + (rackFrame.width + rackSpacer.length),
                        0);
                    using (BlockReference columnRef = new BlockReference(columnLocations[x], rackColumn.id))
                    {
                        columnRef.Layer = rackColumn.layerName;
                        objList.Add(columnRef.Id);
                        modelBlkTblRec.AppendEntity(columnRef);
                        acTrans.AddNewlyCreatedDBObject(columnRef, true);
                    }
                }

                // Calculate beam locations and add to model space (Pallet Rack)
                for (int x = 0; x < beamLocations.GetLength(0); x++)
                {
                    for (int y = 0; y < beamLocations.GetLength(1); y++)
                    {
                        for (int z = 0; z < beamLocations.GetLength(2); z++)
                        {
                            beamLocations[x, y, z] = new Point3d(x * (rackFrame.diameter + rackBeam.length) +
                                rackFrame.diameter + (Convert.ToDouble(y) % 2) * rackBeam.length,
                                Math.Ceiling(Convert.ToDouble(y) / 2) * rackFrame.width +
                                Math.Floor(Convert.ToDouble(y) / 2) * rackSpacer.length,
                                beamBottom + z * (rackFrame.height - beamBottom - rackBeam.height) / (high - 1));
                            using (BlockReference beamRef = new BlockReference(beamLocations[x, y, z], rackBeam.id))
                            {
                                beamRef.Layer = rackBeam.layerName;
                                if (Convert.ToDouble(y) % 2 != 0) { beamRef.Rotation = Math.PI; }
                                objList.Add(beamRef.Id);
                                modelBlkTblRec.AppendEntity(beamRef);
                                acTrans.AddNewlyCreatedDBObject(beamRef, true);
                            }
                        }
                    }
                }

                // Save the new object to the database
                acTrans.Commit();

            }

        }

    }

}
