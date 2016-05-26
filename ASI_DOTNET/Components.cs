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
using Autodesk.AutoCAD.Colors;

[assembly: CommandClass(typeof(ASI_DOTNET.Column))]
[assembly: CommandClass(typeof(ASI_DOTNET.Truss))]
[assembly: CommandClass(typeof(ASI_DOTNET.Frame))]
[assembly: CommandClass(typeof(ASI_DOTNET.Beam))]
[assembly: CommandClass(typeof(ASI_DOTNET.Rail))]
[assembly: CommandClass(typeof(ASI_DOTNET.Stair))]
[assembly: CommandClass(typeof(ASI_DOTNET.Ladder))]

namespace ASI_DOTNET
{
  class Column
  {
    // Auto-impl class properties
    private Database db;
    public double height { get; private set; }
    public double width { get; private set; }
    public double baseHeight { get; set; }
    public double baseWidth { get; private set; }
    public string name { get; private set; }
    public string layerName { get; set; }
    public ObjectId id { get; private set; }

    // Other class properties
    private string baseOrient;

    public string baseplateOrientation
    {
      get { return baseOrient; }
      set
      {
        this.baseOrient = value;
        if (baseHeight > 0) this.name = "Column - " + height + "in - " + width + "x" + baseWidth + " (" + value + ")";
        else this.name = "Column - " + height + "x" + width + " (" + value + ")";
      }
    }

    // Public constructor
    public Column(Database db,
        double height,
        double width,
        double baseWidth)
    {
      Color layerColor;
      this.db = db;
      this.height = height;
      this.width = width;
      this.baseWidth = baseWidth;

      // Set defaults
      this.baseHeight = 0.75;

      // Set name and column layer
      if (baseHeight > 0)
      {
        this.name = "Column - " + height + "in - " + width + "x" + baseWidth;
        this.layerName = "3D-Mezz-Column";
        layerColor = Utils.ChooseColor("black");
      }
      else
      {
        this.name = "Column - " + height + "x" + width;
        this.layerName = "3D-Rack-Column";
        layerColor = Utils.ChooseColor("teal");
      }

      // Create beam layer (if necessary)
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

            // Calculate locations
            Vector3d cVec = Point3d.Origin.GetVectorTo(
                new Point3d(width / 2, width / 2, (height + baseHeight) / 2));
            Vector3d bVec = BaseplateLocation(baseOrient, baseWidth, width, baseHeight);

            if (baseHeight > 0)
            {
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
            }

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

    private static Vector3d BaseplateLocation(string baseOrient,
        double baseWidth,
        double columnWidth,
        double baseHeight)
    {
      Point3d middle;

      switch (baseOrient)
      {
        case "+X":
          middle = new Point3d(baseWidth / 2, columnWidth / 2, baseHeight / 2);
          break;
        case "-X":
          middle = new Point3d(columnWidth - (baseWidth / 2), columnWidth / 2, baseHeight / 2);
          break;
        case "+Y":
          middle = new Point3d(columnWidth / 2, baseWidth / 2, baseHeight / 2);
          break;
        case "-Y":
          middle = new Point3d(columnWidth / 2, columnWidth - (baseWidth / 2), baseHeight / 2);
          break;
        case "+X+Y":
          middle = new Point3d(baseWidth / 2, baseWidth / 2, baseHeight / 2);
          break;
        case "+X-Y":
          middle = new Point3d(baseWidth / 2, columnWidth - (baseWidth / 2), baseHeight / 2);
          break;
        case "-X+Y":
          middle = new Point3d(columnWidth - (baseWidth / 2), baseWidth / 2, baseHeight / 2);
          break;
        case "-X-Y":
          middle = new Point3d(columnWidth - (baseWidth / 2), columnWidth - (baseWidth / 2), baseHeight / 2);
          break;
        default:
          middle = new Point3d(columnWidth / 2, columnWidth / 2, baseHeight / 2);
          break;
      }


      return Point3d.Origin.GetVectorTo(middle);
    }

    [CommandMethod("ColumnPrompt")]
    public static void ColumnPrompt()
    {
      // Get the current document and database, and start a transaction
      Document acDoc = Application.DocumentManager.MdiActiveDocument;
      Database acCurDb = acDoc.Database;

      // Prepare prompt for the column height
      PromptDoubleResult cHeightRes;
      PromptDistanceOptions cHeightOpts = new PromptDistanceOptions("");
      cHeightOpts.Message = "\nEnter the column height: ";
      cHeightOpts.DefaultValue = 96;
      cHeightOpts.AllowZero = false;
      cHeightOpts.AllowNegative = false;

      // Prepare prompt for the column width
      PromptDoubleResult cWidthRes;
      PromptDistanceOptions cWidthOpts = new PromptDistanceOptions("");
      cWidthOpts.Message = "\nEnter the column diameter: ";
      cWidthOpts.DefaultValue = 6;
      cWidthOpts.AllowZero = false;
      cWidthOpts.AllowNegative = false;

      // Prepare prompt for the baseplate diameter
      PromptDoubleResult bWidthRes;
      PromptDistanceOptions bWidthOpts = new PromptDistanceOptions("");
      bWidthOpts.Message = "\nEnter the baseplate diameter: ";
      bWidthOpts.DefaultValue = 14;
      bWidthOpts.AllowZero = false;
      bWidthOpts.AllowNegative = false;

      // Prepare prompt for other options
      PromptResult cOthersRes;
      PromptKeywordOptions cOthersOpts = new PromptKeywordOptions("");
      cOthersOpts.Message = "\nOptions: ";
      cOthersOpts.Keywords.Add("BaseplateHeight");
      cOthersOpts.Keywords.Add("BaseplateOrientation");
      cOthersOpts.AllowArbitraryInput = false;
      cOthersOpts.AllowNone = true;

      // Prepare prompt for the baseplate height
      PromptDoubleResult bHeightRes;
      PromptDistanceOptions bHeightOpts = new PromptDistanceOptions("");
      bHeightOpts.Message = "\nEnter the baseplate height: ";
      bHeightOpts.DefaultValue = 0.75;
      bHeightOpts.AllowNegative = false;

      // Prepare first prompt for the baseplate orientation
      PromptResult bMainOrientRes;
      PromptKeywordOptions bMainOrientOpts = new PromptKeywordOptions("");
      bMainOrientOpts.Message = "\nChoose baseplate orientation: ";
      bMainOrientOpts.Keywords.Add("CEnter");
      bMainOrientOpts.Keywords.Add("COrner");
      bMainOrientOpts.Keywords.Add("Side");
      bMainOrientOpts.Keywords.Default = "CEnter";
      bMainOrientOpts.AllowArbitraryInput = false;
      bMainOrientOpts.AllowNone = false;

      // Prepare secondary prompt for the side baseplate orientation
      PromptResult bSideOrientRes;
      PromptKeywordOptions bSideOrientOpts = new PromptKeywordOptions("");
      bSideOrientOpts.Message = "\nChoose baseplate side orientation: ";
      bSideOrientOpts.Keywords.Add("+X");
      bSideOrientOpts.Keywords.Add("-X");
      bSideOrientOpts.Keywords.Add("+Y");
      bSideOrientOpts.Keywords.Add("-Y");
      bSideOrientOpts.Keywords.Default = "+X";
      bSideOrientOpts.AllowArbitraryInput = false;
      bSideOrientOpts.AllowNone = false;

      // Prepare secondary prompt for the side baseplate orientation
      PromptResult bCornerOrientRes;
      PromptKeywordOptions bCornerOrientOpts = new PromptKeywordOptions("");
      bCornerOrientOpts.Message = "\nChoose baseplate corner orientation: ";
      bCornerOrientOpts.Keywords.Add("+X+Y");
      bCornerOrientOpts.Keywords.Add("+X-Y");
      bCornerOrientOpts.Keywords.Add("-X+Y");
      bCornerOrientOpts.Keywords.Add("-X-Y");
      bCornerOrientOpts.Keywords.Default = "+X+Y";
      bCornerOrientOpts.AllowArbitraryInput = false;
      bCornerOrientOpts.AllowNone = false;

      // Prompt for column height
      cHeightRes = acDoc.Editor.GetDistance(cHeightOpts);
      if (cHeightRes.Status != PromptStatus.OK) return;
      double cHeight = cHeightRes.Value;

      // Prompt for column diameter
      cWidthRes = acDoc.Editor.GetDistance(cWidthOpts);
      if (cWidthRes.Status != PromptStatus.OK) return;
      double cWidth = cWidthRes.Value;

      // Prompt for baseplate diameter
      bWidthRes = acDoc.Editor.GetDistance(bWidthOpts);
      if (bWidthRes.Status != PromptStatus.OK) return;
      double bWidth = bWidthRes.Value;

      // Create column block
      Column columnBlock = new Column(db: acCurDb,
          height: cHeight,
          width: cWidth,
          baseWidth: bWidth);

      // Prompt for other options
      cOthersRes = acDoc.Editor.GetKeywords(cOthersOpts);

      // Exit if the user presses ESC or cancels the command
      if (cOthersRes.Status == PromptStatus.Cancel) return;

      while (cOthersRes.Status == PromptStatus.OK)
      {
        switch (cOthersRes.StringResult)
        {
          case "BaseplateHeight":
            bHeightRes = acDoc.Editor.GetDistance(bHeightOpts);
            if (bHeightRes.Status != PromptStatus.OK) return;
            columnBlock.baseHeight = bHeightRes.Value;
            break;
          case "BaseplateOrientation":
            bMainOrientRes = acDoc.Editor.GetKeywords(bMainOrientOpts);
            if (bMainOrientRes.Status == PromptStatus.Cancel) return;
            else
            {
              switch (bMainOrientRes.StringResult)
              {
                case "CEnter":
                  break;
                case "Side":
                  bSideOrientRes = acDoc.Editor.GetKeywords(bSideOrientOpts);
                  if (bSideOrientRes.Status == PromptStatus.Cancel) return;
                  columnBlock.baseplateOrientation = bSideOrientRes.StringResult;
                  break;
                case "COrner":
                  bCornerOrientRes = acDoc.Editor.GetKeywords(bCornerOrientOpts);
                  if (bCornerOrientRes.Status == PromptStatus.Cancel) return;
                  columnBlock.baseplateOrientation = bCornerOrientRes.StringResult;
                  break;
                default:
                  Application.ShowAlertDialog("Invalid Keyword: " + bMainOrientRes.StringResult);
                  break;
              }
            }
            break;
          default:
            Application.ShowAlertDialog("Invalid Keyword: " + cOthersRes.StringResult);
            break;
        }

        // Re-prompt for keywords
        cOthersRes = acDoc.Editor.GetKeywords(cOthersOpts);

        // Exit if the user presses ESC or cancels the command
        if (cOthersRes.Status == PromptStatus.Cancel) return;
      }

      // Build column
      columnBlock.Build();
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
      this.layerName = "3D-Mezz-Truss";
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
        rotateAngle = -Math.PI / 2;
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

    private static Tuple<double, double> rodIntersectPoint(Solid3d rod,
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

  class Frame
  {
    // Auto-impl class properties
    private Database db;
    public double height { get; private set; }
    public double width { get; private set; }
    public double diameter { get; private set; }
    public string name { get; private set; }
    public string layerName { get; set; }
    public ObjectId id { get; private set; }

    // Public constructor
    public Frame(Database db,
        double height = 96,
        double width = 42,
        double diameter = 3)
    {
      this.db = db;
      this.height = height;
      this.width = width;
      this.diameter = diameter;
      this.name = "Frame - " + height + "x" + width;

      // Create beam layer (if necessary)
      this.layerName = "3D-Rack-Frame";
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
    private Solid3d framePost(double height,
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
    public double step { get; private set; }
    public double thickness { get; private set; }
    public string name { get; private set; }
    public string orientation { get; private set; }
    public string style { get; private set; }
    public string layerName { get; set; }
    public Color layerColor { get; set; }
    public ObjectId id { get; private set; }

    // Public constructor
    public Beam(Database db,
        double length = 96,
        double height = 3,
        double width = 2,
        string orient = "X-Axis",
        string style = "Box")
    {
      this.db = db;
      this.length = length;
      this.height = height;
      this.width = width;
      this.orientation = orient;
      this.style = style;
      this.name = "Beam - " + length + "x" + height + "x" + width;
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
                beamThickness: thickness,
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
    private Polyline beamProfile(double height,
        double width,
        string style = "Box",
        double stepSize = 0.75,
        double beamThickness = 0.75,
        string orient = "X-Axis")
    {
      int i = 0;

      /// Create polyline of brace profile and rotate to correct orienation
      Polyline poly = new Polyline();
      if (style == "Step")
      {
        poly.AddVertexAt(i, new Point2d(0, 0), 0, 0, 0);
        poly.AddVertexAt(i++, new Point2d(-height, 0), 0, 0, 0);
        poly.AddVertexAt(i++, new Point2d(-height, width - stepSize), 0, 0, 0);
        poly.AddVertexAt(i++, new Point2d(-height + stepSize, width - stepSize), 0, 0, 0);
        poly.AddVertexAt(i++, new Point2d(-height + stepSize, width), 0, 0, 0);
        poly.AddVertexAt(i++, new Point2d(0, width), 0, 0, 0);
      }
      else if (style == "Box")
      {
        poly.AddVertexAt(i, new Point2d(0, 0), 0, 0, 0);
        poly.AddVertexAt(i++, new Point2d(-height, 0), 0, 0, 0);
        poly.AddVertexAt(i++, new Point2d(-height, width), 0, 0, 0);
        poly.AddVertexAt(i++, new Point2d(0, width), 0, 0, 0);
      }
      else if (style == "IBeam")
      {
        poly.AddVertexAt(i, new Point2d(0, 0), 0, 0, 0);
        poly.AddVertexAt(i++, new Point2d(-beamThickness, 0), 0, 0, 0);
        poly.AddVertexAt(i++, new Point2d(-beamThickness, (width - beamThickness) / 2), 0, 0, 0);
        poly.AddVertexAt(i++, new Point2d(-height + beamThickness, (width - beamThickness) / 2), 0, 0, 0);
        poly.AddVertexAt(i++, new Point2d(-height + beamThickness, 0), 0, 0, 0);
        poly.AddVertexAt(i++, new Point2d(-height, 0), 0, 0, 0);
        poly.AddVertexAt(i++, new Point2d(-height, width), 0, 0, 0);
        poly.AddVertexAt(i++, new Point2d(-height + beamThickness, width), 0, 0, 0);
        poly.AddVertexAt(i++, new Point2d(-height + beamThickness, (width + beamThickness) / 2), 0, 0, 0);
        poly.AddVertexAt(i++, new Point2d(-beamThickness, (width + beamThickness) / 2), 0, 0, 0);
        poly.AddVertexAt(i++, new Point2d(-beamThickness, width), 0, 0, 0);
        poly.AddVertexAt(i++, new Point2d(0, width), 0, 0, 0);
      }
      else if (style == "CChannel")
      {
        poly.AddVertexAt(i, new Point2d(0, 0), 0, 0, 0);
        poly.AddVertexAt(i++, new Point2d(-height, 0), 0, 0, 0);
        poly.AddVertexAt(i++, new Point2d(-height, width), 0, 0, 0);
        poly.AddVertexAt(i++, new Point2d(-height + beamThickness, width), 0, 0, 0);
        poly.AddVertexAt(i++, new Point2d(-height + beamThickness, beamThickness), 0, 0, 0);
        poly.AddVertexAt(i++, new Point2d(-beamThickness, beamThickness), 0, 0, 0);
        poly.AddVertexAt(i++, new Point2d(-beamThickness, width), 0, 0, 0);
        poly.AddVertexAt(i++, new Point2d(0, width), 0, 0, 0);
      }

      poly.Closed = true;
      poly.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.YAxis, Point3d.Origin));

      // Rotate beam profile if necessary
      if (orient == "Y-Axis")
      {
        poly.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.ZAxis, Point3d.Origin));
      }

      return poly;
    }

    private Line beamPath(double length,
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
      bOthersOpts.Keywords.Add("Thickness");
      bOthersOpts.Keywords.Add("Style");
      bOthersOpts.AllowArbitraryInput = false;
      bOthersOpts.AllowNone = true;

      // Prepare prompt for the beam orientation
      PromptResult bOrientRes;
      PromptKeywordOptions bOrientOpts = new PromptKeywordOptions("");
      bOrientOpts.Message = "\nEnter beam orientation: ";
      bOrientOpts.Keywords.Add("X-Axis");
      bOrientOpts.Keywords.Add("Y-Axis");
      bOrientOpts.Keywords.Default = "X-Axis";
      bOrientOpts.AllowArbitraryInput = false;

      // Prepare prompt for the beam thickness
      PromptDoubleResult thickRes;
      PromptDistanceOptions thickOpts = new PromptDistanceOptions("");
      thickOpts.Message = "\nEnter the beam thickness: ";
      thickOpts.DefaultValue = 0.75;

      // Prepare prompt for the beam style
      PromptResult bStyleRes;
      PromptKeywordOptions bStyleOpts = new PromptKeywordOptions("");
      bStyleOpts.Message = "\nEnter beam style: ";
      bStyleOpts.Keywords.Add("Step");
      bStyleOpts.Keywords.Add("Box");
      bStyleOpts.Keywords.Add("IBeam");
      bStyleOpts.Keywords.Add("CChannel");
      bStyleOpts.Keywords.Default = "Box";
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

      // Create beam
      Beam solidBeam = new Beam(db: acCurDb,
          length: length,
          height: height,
          width: width);

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
            solidBeam.orientation = bOrientRes.StringResult;
            if (bOrientRes.Status != PromptStatus.OK) return;
            break;
          case "Thickness":
            thickRes = acDoc.Editor.GetDistance(thickOpts);
            solidBeam.thickness = thickRes.Value;
            if (thickRes.Status != PromptStatus.OK) return;
            break;
          case "Style":
            bStyleRes = acDoc.Editor.GetKeywords(bStyleOpts);
            solidBeam.style = bStyleRes.StringResult;
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

      // Build beam
      solidBeam.Build();
    }

  }

  class Rail
  {
    // Auto-impl class properties
    public Database db { get; private set; }
    public Point3dCollection points { get; private set; }
    public double height { get; private set; }
    public int numSections { get; private set; }
    public int tiers { get; set; }
    public double[] tierHeight { get; set; }
    public double postWidth { get; private set; }
    public double defaultRailLength { get; set; }
    public double railWidth { get; private set; }
    public string layerName { get; set; }
    public List<bool> firstPost { get; set; }
    public List<bool> lastPost { get; set; }
    public bool kickPlate { get; set; }
    public List<Line> pathSections { get; private set; }
    public List<double> pathOrientation { get; private set; }
    public double postOrientation { get; private set; }
    public double kickplateHeight { get; set; }
    public double kickplateWidth { get; set; }

    // Public constructor
    public Rail(Database db,
        Point3dCollection pts)
    {
      Color layerColor;
      this.db = db;
      this.points = pts;
      this.numSections = points.Count - 1;

      // Set some defaults
      // postOrientation = 1 --> left of section line
      // postOrientation = -1 --> right of section line
      this.postOrientation = 1;
      this.postWidth = 1.5;
      this.kickPlate = true;

      // Create beam layer (if necessary)
      this.layerName = "3D-Mezz-Rail";
      layerColor = Utils.ChooseColor("yellow");
      Utils.CreateLayer(db, layerName, layerColor);

      // Initialize lists
      this.pathSections = new List<Line>();
      this.pathOrientation = new List<double>();
      this.firstPost = new List<bool>();
      this.lastPost = new List<bool>();

    }

    private void Calculations()
    {

      // Loop over number of railing sections for initial calculations
      for (int s = 0; s < numSections; s++)
      {

        // Get current path section
        this.pathSections.Add(new Line(points[s], points[s + 1]));

        // Path horizontal orientation (phi in polar coordinates)
        this.pathOrientation.Add(Utils.PolarAnglePhi(pathSections[s]));

        // Set first post defaults
        // All subsequent sections do not need first post (from last post of previous section)
        if (s > 0) this.firstPost.Add(false);
        else this.firstPost.Add(true);

        // Set last post defaults
        this.lastPost.Add(true);

      }

      // Loop over number of railing sections for detailed section calculations
      for (int s = 0; s < numSections; s++)
      {

        // Move section start & end points to correct point on last post
        if (s != 0)
        {

          // Straight sections
          if (pathOrientation[s] == pathOrientation[s - 1])
          {
            this.pathSections[s].StartPoint = pathSections[s].StartPoint + new Vector3d(
              -postWidth * Math.Cos(pathOrientation[s]),
              -postWidth * Math.Sin(pathOrientation[s]),
              0);
          }
          // Right hand turn
          else if (Math.Sin(pathOrientation[s] - pathOrientation[s - 1]) < 0)
          {
            this.pathSections[s - 1].EndPoint = pathSections[s - 1].EndPoint + new Vector3d(
              postWidth * Math.Cos(pathOrientation[s - 1]),
              postWidth * Math.Sin(pathOrientation[s - 1]),
              0);
            this.pathSections[s].StartPoint = pathSections[s].StartPoint + new Vector3d(
              -postWidth * Math.Cos(pathOrientation[s]),
              -postWidth * Math.Sin(pathOrientation[s]),
              0);
            Application.ShowAlertDialog("Section " + (s + 1) + "\nMoved startpoint " +
              new Vector3d(-postWidth * Math.Cos(pathOrientation[s]), -postWidth * Math.Sin(pathOrientation[s]), 0));
          }

        }
      }
    }

    public void Build()
    {

      // Declare variables
      RailSection section;

      // Run calculations
      Calculations();

      using (Transaction acTrans = db.TransactionManager.StartTransaction())
      {
        // Open the Block table for read
        BlockTable acBlkTbl;
        acBlkTbl = acTrans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;

        // Open the Block table record Model space for write
        BlockTableRecord modelBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
            OpenMode.ForWrite) as BlockTableRecord;

        // Loop over number of railing sections
        for (int s = 0; s < numSections; s++)
        {

          section = new RailSection(modelBlkTblRec,
            acTrans,
            pathSections[s],
            postOrientation,
            layerName);

          // First and last post options
          if (!firstPost[s]) section.firstPost = false;
          if (!lastPost[s]) section.lastPost = false;

          // Kickplate options
          if (!kickPlate) section.kickPlate = false;

          section.Build();

        }

        // Save the transaction
        acTrans.Commit();

      }

    }

    [CommandMethod("RailPrompt")]
    public static void RailPrompt()
    {
      // Declare variables
      Point3dCollection pts = new Point3dCollection();
      bool? kickPlate = null;
      double? orientation = null;

      // Get the current document and database, and start a transaction
      Document doc = Application.DocumentManager.MdiActiveDocument;
      Database db = doc.Database;

      // Prepare prompt for points
      PromptPointResult ptRes;
      PromptPointOptions ptOpts = new PromptPointOptions("");
      ptOpts.SetMessageAndKeywords("\nEnter point or [Kickplate/PostOrientation]",
          "Kickplate PostOrientation");
      ptOpts.AllowNone = true;

      // Prepare prompt for kickplates
      PromptResult kickplateRes;
      PromptKeywordOptions kickplateOpts = new PromptKeywordOptions("");
      kickplateOpts.Message = "\nInclude kickplate? ";
      kickplateOpts.Keywords.Add("True");
      kickplateOpts.Keywords.Add("False");
      kickplateOpts.Keywords.Default = "True";
      kickplateOpts.AllowArbitraryInput = false;

      // Prepare prompt for kickplates
      PromptResult orientRes;
      PromptKeywordOptions orientOpts = new PromptKeywordOptions("");
      orientOpts.Message = "\nPost on which side? ";
      orientOpts.Keywords.Add("Left");
      orientOpts.Keywords.Add("Right");
      orientOpts.Keywords.Default = "Left";
      orientOpts.AllowArbitraryInput = false;

      do
      {
        ptRes = doc.Editor.GetPoint(ptOpts);
        if (ptRes.Status == PromptStatus.Cancel) return;
        if (ptRes.Status == PromptStatus.Keyword)
          switch (ptRes.StringResult)
          {
            case "KickPlate":
              kickplateRes = doc.Editor.GetKeywords(kickplateOpts);
              if (kickplateRes.Status != PromptStatus.OK) return;
              if (kickplateRes.StringResult == "False") kickPlate = false;
              break;
            case "PostOrientation":
              orientRes = doc.Editor.GetKeywords(orientOpts);
              if (orientRes.Status != PromptStatus.OK) return;
              if (orientRes.StringResult == "Left") orientation = 1;
              else orientation = -1;
              break;
            default:
              break;
          }
        else if (ptRes.Status == PromptStatus.OK)
        {
          pts.Add(ptRes.Value);
          ptOpts.BasePoint = ptRes.Value;
          ptOpts.UseBasePoint = true;
          ptOpts.UseDashedLine = true;
        }
      } while (ptRes.Status != PromptStatus.None);

      if (ptRes.Status != PromptStatus.None) return;

      // Initialize rail system
      Rail rail = new Rail(db, pts);

      // Deal with kickplate input
      if (kickPlate.HasValue) rail.kickPlate = (bool)kickPlate;

      // Deal with post orientation input
      if (orientation.HasValue) rail.postOrientation = (double)orientation;

      // Build rail
      rail.Build();
    }

  }

  class RailSection
  {
    // Auto-impl class properties
    public BlockTableRecord blkTblRec { get; private set; }
    public Transaction acTrans { get; private set; }
    public Line pathSection { get; private set; }
    public double height { get; private set; }
    public int tiers { get; set; }
    public double[] tierHeight { get; set; }
    public double postWidth { get; private set; }
    public double defaultRailLength { get; set; }
    public double defaultOrientation { get; set; }
    public double railWidth { get; private set; }
    public string layerName { get; private set; }
    public bool firstPost { get; set; }
    public bool lastPost { get; set; }
    public bool kickPlate { get; set; }
    public double pathOrientation { get; private set; }
    public double pathVerticalAngle { get; private set; }
    public double postOrientation { get; private set; }
    public double sectionRails { get; private set; }
    public double railLength { get; private set; }
    public double kickplateHeight { get; private set; }
    public double kickplateWidth { get; private set; }

    // Public constructor
    public RailSection(BlockTableRecord rec,
      Transaction tr,
      Line section,
      double side,
      string layerName)
    {
      // Set options to object properties
      this.blkTblRec = rec;
      this.acTrans = tr;
      this.pathSection = section;
      this.postOrientation = side;
      this.layerName = layerName;

      // Set some defaults
      this.height = 42;
      this.tierHeight = new double[2] { height - 18, height };
      this.tiers = tierHeight.GetLength(0);
      this.postWidth = 1.5;
      this.railWidth = 1.5;
      this.defaultRailLength = 60;
      this.kickplateHeight = 4;
      this.kickplateWidth = 0.25;
      this.firstPost = true;
      this.lastPost = true;

      //Path horizontal orientation(phi in polar coordinates)
      this.pathOrientation = Utils.PolarAnglePhi(pathSection);

      //Path vertical orientation(theta in polar coordinates)
      this.pathVerticalAngle = Utils.PolarAngleTheta(pathSection);

      // Kickplate defaults (only in flat sections)
      if (Math.Round(pathVerticalAngle, 1) == 0) this.kickPlate = true;
      else this.kickPlate = false;

    }

    private void ErrorCheck() {

      // Error checks for section length (must be at least 2 post widths)
      // Post width corrected for vertical angle
      if (pathSection.Length < (2 * (postWidth / Math.Sin(pathVerticalAngle))))
      {
        throw new System.Exception("Invalid railing section: too short." +
            "\nFrom (" + pathSection.StartPoint.X + ", " + pathSection.StartPoint.Y + ", " + pathSection.StartPoint.Z + ") " +
            "\to (" + pathSection.EndPoint.X + ", " + pathSection.EndPoint.Y + ", " + pathSection.EndPoint.Z + ")");
      }

      // Check that section is not too steep (45 degrees or PI / 4 radians from horizontal)
      if (pathVerticalAngle < (Math.PI / 4) && pathVerticalAngle > (3 * Math.PI / 4))
      {
        throw new System.Exception("Invalid railing section: too steep (must be < 45)." +
            "\nFrom (" + pathSection.StartPoint.X + ", " + pathSection.StartPoint.Y + ", " + pathSection.StartPoint.Z + ") " +
            "\to (" + pathSection.EndPoint.X + ", " + pathSection.EndPoint.Y + ", " + pathSection.EndPoint.Z + ")" +
            "\nAngle (from horizontal): " + (90 - pathVerticalAngle.ToDegrees()));
      }

      // Check that post orientation is valid
      if (Math.Abs(postOrientation) != 1)
      {
        throw new System.Exception("Invalid argument in CreatePost..." +
          "\nside must equal 1 || -1" +
          "\nside = " + postOrientation);
      }

    }

    private void Calculations() {

      // Calculate rail sections and length (per section)
      this.sectionRails = Math.Max(1, Math.Round(pathSection.Length / (defaultRailLength - postWidth)));
      this.railLength = ((pathSection.Length - ((sectionRails + 1) *
          (postWidth / Math.Sin(pathVerticalAngle)))) / sectionRails);

    }

    public void Build()
    {
      // Declare variables
      Solid3d post;
      Solid3d rail;
      Solid3d kickplate;
      Vector3d startVector;
      Entity tempEnt;
      double numPosts = 0;

      // Error checking
      try
      {
        ErrorCheck();
      }
      catch (System.Exception ex)
      {
        Application.ShowAlertDialog(ex.Message);
        return;
      }

      // Run calculations
      Calculations();

      // Create post
      post = CreatePost(height, postWidth, pathVerticalAngle, pathOrientation, postOrientation);
      post.Layer = layerName;

      // Create rail section
      rail = CreateRail(railLength, railWidth, pathVerticalAngle, pathOrientation, postOrientation);
      rail.Layer = layerName;

      // Create kickplate
      kickplate = CreateKickplate(railLength, kickplateHeight, kickplateWidth, pathVerticalAngle,
        pathOrientation, postOrientation);
      kickplate.Layer = layerName;

      // Position first post (if necessary) and add to model and transaction
      startVector = pathSection.StartPoint - Point3d.Origin;
      if (firstPost)
      {
        tempEnt = post.GetTransformedCopy(Matrix3d.Displacement(startVector));
        numPosts++;
        blkTblRec.AppendEntity(tempEnt);
        acTrans.AddNewlyCreatedDBObject(tempEnt, true);
      }
      else
      {
        numPosts++;
      }

      // Loop over other rail sections
      for (int i = 1; i <= sectionRails; i++)
      {
        // Position rails and add to model and transaction
        foreach (double t in tierHeight)
        {
          // Pass true to RailLocation if no first post
          tempEnt = rail.GetTransformedCopy(Matrix3d.Displacement(
            startVector.Add(RailLocation(t, postWidth, pathVerticalAngle,
            pathOrientation))));
          blkTblRec.AppendEntity(tempEnt);
          acTrans.AddNewlyCreatedDBObject(tempEnt, true);
        }

        // Postition kickplate and add to model and transaction
        tempEnt = kickplate.GetTransformedCopy(Matrix3d.Displacement(
          startVector.Add(KickplateLocation(postWidth, kickplateHeight,
          pathVerticalAngle, pathOrientation))));
        blkTblRec.AppendEntity(tempEnt);
        acTrans.AddNewlyCreatedDBObject(tempEnt, true);

        // Account for lastPost
        if (i == sectionRails && !lastPost) continue;
        else
        {
          // Determine base point for next post
          startVector = pathSection.GetPointAtDist((railLength * i) +
            ((postWidth / Math.Sin(pathVerticalAngle)) * numPosts)) - Point3d.Origin;

          // Position post and add to model and transaction
          tempEnt = post.GetTransformedCopy(Matrix3d.Displacement(startVector));
          numPosts++;
          blkTblRec.AppendEntity(tempEnt);
          acTrans.AddNewlyCreatedDBObject(tempEnt, true);

        }

      }

    }

    private static Solid3d CreatePost(double height,
        double width,
        double vAngle,
        double oAngle,
        double side)
    {

      // Calculate angle offset
      double offset = width / Math.Tan(vAngle);

      // Create polyline of rail post profile
      Polyline postPoly = new Polyline();
      postPoly.AddVertexAt(0, new Point2d(0, 0), 0, 0, 0);
      postPoly.AddVertexAt(1, new Point2d(width, offset), 0, 0, 0);
      postPoly.AddVertexAt(2, new Point2d(width, height + offset), 0, 0, 0);
      postPoly.AddVertexAt(3, new Point2d(0, height), 0, 0, 0);
      postPoly.Closed = true;

      // Create the post
      Solid3d post = Utils.ExtrudePolyline(postPoly, -side * width);

      // Position the post vertically
      post.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.XAxis, Point3d.Origin));

      // Rotate the post
      post.TransformBy(Matrix3d.Rotation(oAngle, Vector3d.ZAxis, Point3d.Origin));

      return post;
    }

    private static Solid3d CreateRail(double length,
        double width,
        double vAngle,
        double oAngle,
        double side)
    {

      // Calculate angle offset
      double vOffset = length * Math.Cos(vAngle);

      // Create polyline of rail profile
      Polyline railPoly = new Polyline();
      railPoly.AddVertexAt(0, new Point2d(0, 0), 0, 0, 0);
      railPoly.AddVertexAt(1, new Point2d(0, width), 0, 0, 0);
      railPoly.AddVertexAt(2, new Point2d(length * Math.Sin(vAngle), width - vOffset), 0, 0, 0);
      railPoly.AddVertexAt(3, new Point2d(length * Math.Sin(vAngle), -vOffset), 0, 0, 0);
      railPoly.Closed = true;

      // Create the rail
      Solid3d rail = Utils.ExtrudePolyline(railPoly, side * width);

      // Position the rail vertically
      rail.TransformBy(Matrix3d.Rotation(-Math.PI / 2, Vector3d.XAxis, Point3d.Origin));

      // Rotate the rail
      rail.TransformBy(Matrix3d.Rotation(oAngle, Vector3d.ZAxis, Point3d.Origin));

      return rail;
    }

    private static Vector3d RailLocation(double height,
        double width,
        double vAngle,
        double oAngle)
    {
      // Declare variables
      double x = 0;
      double y = 0;
      double z = 0;

      // Calculate angle offset
      double offset = width / Math.Tan(vAngle);

      // Point coordinates
      x = width * Math.Cos(oAngle);
      y = width * Math.Sin(oAngle);
      z = offset + height;

      return new Vector3d(x, y, z);
    }

    private static Solid3d CreateKickplate(double length,
        double height,
        double width,
        double vAngle,
        double oAngle,
        double side)
    {

      // Calculate angle offset
      double vOffset = length * Math.Cos(vAngle);

      // Create polyline of rail profile
      Polyline kickplatePoly = new Polyline();
      kickplatePoly.AddVertexAt(0, new Point2d(0, 0), 0, 0, 0);
      kickplatePoly.AddVertexAt(1, new Point2d(0, height), 0, 0, 0);
      kickplatePoly.AddVertexAt(2, new Point2d(length * Math.Sin(vAngle), height - vOffset), 0, 0, 0);
      kickplatePoly.AddVertexAt(3, new Point2d(length * Math.Sin(vAngle), -vOffset), 0, 0, 0);
      kickplatePoly.Closed = true;

      // Create the rail
      Solid3d kickplate = Utils.ExtrudePolyline(kickplatePoly, side * width);

      // Position the rail vertically
      kickplate.TransformBy(Matrix3d.Rotation(-Math.PI / 2, Vector3d.XAxis, Point3d.Origin));

      // Rotate the rail
      kickplate.TransformBy(Matrix3d.Rotation(oAngle, Vector3d.ZAxis, Point3d.Origin));

      return kickplate;
    }

    private static Vector3d KickplateLocation(double postWidth,
        double height,
        double vAngle,
        double oAngle)
    {
      // Declare variables
      double x;
      double y;
      double z;

      // Calculate height offset
      z = height + postWidth / Math.Tan(vAngle);

      // Point coordinates
      x = postWidth * Math.Cos(oAngle);
      y = postWidth * Math.Sin(oAngle);

      return new Vector3d(x, y, z);
    }

  }

  class Stair
  {
    // Auto-impl class properties
    public Database db { get; private set; }
    public double height { get; private set; }
    public double width { get; private set; }
    public double stairHeight { get; private set; }
    public double length { get; private set; }
    public double stringerWidth { get; set; }
    public double stringerDepth { get; set; }
    public double treadHeight { get; set; }

    // Vector in stair "right" direction (looking up stairs)
    public Vector3d widthVector { get; set; }
    public string layerName { get; set; }

    // Other properties
    private int _nRisers;
    private double _defaultRiserHeight;
    private double _stairDepth;
    private double _treadOverlap;
    private Point3d _basePoint;
    private Point3d _topPoint;
    private string _definePoint;
    private Vector3d _lengthVector;

    public int numRisers
    {
      get { return _nRisers; }
      set
      {
        this._nRisers = value;

        // Calculate stair height
        this.stairHeight = height / value;

        // Calculate stair length
        this.length = (value - 1) * stairDepth - ((value - 2) * treadOverlap);

        // Move base point if defining point is top
        if (_definePoint == "top")
        {
          this._basePoint = _topPoint.Add(new Vector3d(
              (length * stairLengthVector.X),
              (length * stairLengthVector.Y),
              -height));
        }
      }
    }

    public double stairDepth
    {
      get { return _stairDepth; }
      set
      {
        this._stairDepth = value;

        // Calculate stair length
        this.length = (numRisers - 1) * value - ((numRisers - 2) * treadOverlap);

        // Move base point if defining point is top
        if (_definePoint == "top")
        {
          this._basePoint = _topPoint.Add(new Vector3d(
              (length * stairLengthVector.X),
              (length * stairLengthVector.Y),
              -height));
        }
      }
    }

    public double defaultRiserHeight
    {
      get { return _defaultRiserHeight; }
      set
      {
        this._defaultRiserHeight = value;

        // Calculate number of stairs
        this.numRisers = Convert.ToInt32(
            Math.Max(1, Math.Round(height / value))
            );
      }
    }

    public double treadOverlap
    {
      get { return _treadOverlap; }
      set
      {
        this._treadOverlap = value;

        // Calculate stair length
        this.length = (numRisers - 1) * stairDepth - ((numRisers - 2) * value);

        // Move base point if defining point is top
        if (_definePoint == "top")
        {
          this._basePoint = _topPoint.Add(new Vector3d(
              (length * stairLengthVector.X),
              (length * stairLengthVector.Y),
              -height));
        }
      }
    }

    public Point3d stairTopPoint
    {
      get { return _topPoint; }
      set
      {
        this._topPoint = value;
        this._definePoint = "top";
        this._basePoint = _topPoint.Add(new Vector3d(
            (length * stairLengthVector.X),
            (length * stairLengthVector.Y),
            -height));
      }
    }

    public Point3d stairBasePoint
    {
      get { return _basePoint; }
      set
      {
        this._basePoint = value;
        this._definePoint = "base";
        this._topPoint = _basePoint.Add(new Vector3d(
            (length * -stairLengthVector.X),
            (length * -stairLengthVector.Y),
            height));
      }
    }

    // Vector in stair "down" direction
    public Vector3d stairLengthVector
    {
      get { return _lengthVector; }
      set
      {
        this._lengthVector = value.UnitVector();
        this._basePoint = _topPoint.Add(new Vector3d(
            (length * stairLengthVector.X),
            (length * stairLengthVector.Y),
            -height));
        this.widthVector = new Vector3d(
            Math.Round(stairLengthVector.RotateBy(Math.PI / 2, Vector3d.ZAxis).X, 0),
            Math.Round(stairLengthVector.RotateBy(Math.PI / 2, Vector3d.ZAxis).Y, 0),
            Math.Round(stairLengthVector.RotateBy(Math.PI / 2, Vector3d.ZAxis).Z, 0));
      }
    }

    // Public constructor
    public Stair(Database db,
        double height,
        double width)
    {
      Color layerColor;
      this.db = db;
      this.height = height;
      this.width = width;

      // Set some defaults
      this.stairDepth = 11;
      this.defaultRiserHeight = 7;
      this.stringerWidth = 1.5;
      this.stringerDepth = 12;
      this.treadHeight = 1;
      this.treadOverlap = 0;
      this._lengthVector = -Vector3d.XAxis;
      this.stairBasePoint = Point3d.Origin;
      this.widthVector = -Vector3d.YAxis;

      // Create beam layer (if necessary)
      this.layerName = "3D-Mezz-Egress";
      layerColor = Utils.ChooseColor("teal");
      Utils.CreateLayer(db, layerName, layerColor);

    }

    public void Build()
    {
      // Declare variables
      Solid3d stringer;
      Solid3d tread;
      Vector3d locationVector;
      Entity tempEnt;

      using (Transaction acTrans = db.TransactionManager.StartTransaction())
      {
        // Open the Block table for read
        BlockTable acBlkTbl;
        acBlkTbl = acTrans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;

        // Open the Block table record Model space for write
        BlockTableRecord modelBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
          OpenMode.ForWrite) as BlockTableRecord;

        // Create stringer
        stringer = CreateStringer(height, length, stairHeight,
          stairDepth, stringerWidth, stringerDepth, stairLengthVector, widthVector);
        stringer.Layer = layerName;

        // Create stair tread
        tread = CreateTread(width, stringerWidth, stairDepth, treadHeight,
          stairLengthVector, widthVector);
        tread.Layer = layerName;

        // Add first stringer to model and transaction
        locationVector = stairBasePoint - Point3d.Origin;
        tempEnt = stringer.GetTransformedCopy(Matrix3d.Displacement(locationVector));
        modelBlkTblRec.AppendEntity(tempEnt);
        acTrans.AddNewlyCreatedDBObject(tempEnt, true);

        // Add second stringer to model and transaction
        tempEnt = stringer.GetTransformedCopy(Matrix3d.Displacement(locationVector.Add(
          new Vector3d(((width - stringerWidth) * widthVector.X),
            ((width - stringerWidth) * widthVector.Y),
            0))));
        modelBlkTblRec.AppendEntity(tempEnt);
        acTrans.AddNewlyCreatedDBObject(tempEnt, true);

        // Loop over stair treads
        for (int i = 1; i < numRisers; i++)
        {
          tempEnt = tread.GetTransformedCopy(Matrix3d.Displacement(locationVector.Add(
            new Vector3d(
              (stringerWidth * widthVector.X) +
              ((i - 1) * (stairDepth - treadOverlap) * -stairLengthVector.X) +
              ((stairLengthVector.CrossProduct(widthVector).X) * (i * stairHeight)),
              (stringerWidth * widthVector.Y) +
              ((i - 1) * (stairDepth - treadOverlap) * -stairLengthVector.Y) +
              ((stairLengthVector.CrossProduct(widthVector).X) * (i * stairHeight)),
              (stringerWidth * widthVector.Z) +
              ((i - 1) * (stairDepth - treadOverlap) * -stairLengthVector.Z) +
              ((stairLengthVector.CrossProduct(widthVector).Z) * (i * stairHeight))))));
          modelBlkTblRec.AppendEntity(tempEnt);
          acTrans.AddNewlyCreatedDBObject(tempEnt, true);
        }

        // Save the transaction
        acTrans.Commit();

      }

    }

    private static Solid3d CreateStringer(double height,
      double length,
      double stairHeight,
      double stairDepth,
      double stringerWidth,
      double stringerDepth,
      Vector3d lengthVector,
      Vector3d widthVector)
    {
      // Calculate stair angle and stringer offsets
      double stairAngle = Math.Atan2(height, length + stairDepth);
      double vOffset = (stringerDepth - (stairDepth * Math.Sin(stairAngle))) / Math.Cos(stairAngle);
      double hOffset = (stringerDepth - (stairHeight * Math.Cos(stairAngle))) / Math.Sin(stairAngle);

      // Create polyline of stringer profile (laying down in sector 1 of X/Y)
      Polyline stringerPoly = new Polyline();
      stringerPoly.AddVertexAt(0, new Point2d(0, 0), 0, 0, 0);
      stringerPoly.AddVertexAt(1, new Point2d(0, stairHeight), 0, 0, 0);
      stringerPoly.AddVertexAt(2, new Point2d(length, height), 0, 0, 0);
      stringerPoly.AddVertexAt(3, new Point2d(length, height - stairHeight - vOffset), 0, 0, 0);
      stringerPoly.AddVertexAt(4, new Point2d(hOffset, 0), 0, 0, 0);
      stringerPoly.Closed = true;

      // Create the stringer (into positive Z)
      Solid3d stringer = Utils.ExtrudePolyline(stringerPoly, stringerWidth);

      // Position the stringer vertically
      stringer.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.XAxis, Point3d.Origin));

      // Position stringer based on width vector (if width vector is not default/right)
      if ((lengthVector.CrossProduct(widthVector)) == -Vector3d.ZAxis)
      {
        stringer.TransformBy(Matrix3d.Displacement(new Vector3d(0, stringerWidth, 0)));
      }

      // Rotate the stringer
      stringer.TransformBy(Matrix3d.Rotation(
          Vector3d.XAxis.GetAngleTo(-lengthVector, Vector3d.ZAxis),
          lengthVector.CrossProduct(widthVector),
          Point3d.Origin));

      return stringer;
    }

    private static Solid3d CreateTread(double width,
      double stringerWidth,
      double stairDepth,
      double treadHeight,
      Vector3d lengthVector,
      Vector3d widthVector)
    {
      // Calculate tread width
      double treadWidth = width - (2 * stringerWidth);

      // Create polyline of tread profile
      Polyline treadPoly = new Polyline();
      treadPoly.AddVertexAt(0, new Point2d(0, 0), 0, 0, 0);
      treadPoly.AddVertexAt(1, new Point2d(0, -treadHeight), 0, 0, 0);
      treadPoly.AddVertexAt(2, new Point2d(stairDepth, -treadHeight), 0, 0, 0);
      treadPoly.AddVertexAt(3, new Point2d(stairDepth, 0), 0, 0, 0);
      treadPoly.Closed = true;

      // Create the tread
      Solid3d tread = Utils.ExtrudePolyline(treadPoly, treadWidth);

      // Position the tread vertically
      tread.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.XAxis, Point3d.Origin));

      // Position tread based on width vector
      if (lengthVector.CrossProduct(widthVector) == -Vector3d.ZAxis)
      {
        tread.TransformBy(Matrix3d.Displacement(new Vector3d(0, treadWidth, 0)));
      }

      // Rotate the tread
      tread.TransformBy(Matrix3d.Rotation(
          Vector3d.XAxis.GetAngleTo(-lengthVector, Vector3d.ZAxis), Vector3d.ZAxis, Point3d.Origin));

      return tread;
    }

    [CommandMethod("StairPrompt")]
    public static void StairPrompt()
    {
      // Get the current document and database, and start a transaction
      Document doc = Application.DocumentManager.MdiActiveDocument;
      Database db = doc.Database;

      // Prepare prompt for the stair height
      PromptDoubleResult heightRes;
      PromptDistanceOptions heightOpts = new PromptDistanceOptions("");
      heightOpts.Message = "\nEnter the stair height: ";
      heightOpts.DefaultValue = 108;

      // Prepare prompt for the stair width
      PromptDoubleResult widthRes;
      PromptDistanceOptions widthOpts = new PromptDistanceOptions("");
      widthOpts.Message = "\nEnter the stair width: ";
      widthOpts.DefaultValue = 36;

      // Prepare prompt for other options
      PromptResult othersRes;
      PromptKeywordOptions othersOpts = new PromptKeywordOptions("");
      othersOpts.Message = "\nOptions: ";
      othersOpts.Keywords.Add("TopPoint");
      othersOpts.Keywords.Add("BottomPoint");
      othersOpts.Keywords.Add("TreadDepth");
      othersOpts.Keywords.Add("RiserHeight");
      othersOpts.Keywords.Add("Overlap");
      othersOpts.AllowArbitraryInput = false;
      othersOpts.AllowNone = true;

      // Prepare prompt for top point
      PromptPointResult topPointResult;
      PromptPointOptions topPointOpts = new PromptPointOptions("");
      topPointOpts.Message = "\nSelect top point: ";
      topPointOpts.AllowNone = true;

      // Prepare prompt for bottom point
      PromptPointResult bottomPointResult;
      PromptPointOptions bottomPointOpts = new PromptPointOptions("");
      bottomPointOpts.Message = "\nSelect bottom point: ";
      bottomPointOpts.AllowNone = true;

      // Prepare prompt for bottom point
      PromptPointResult vectorPointResult;
      PromptPointOptions vectorPointOpts = new PromptPointOptions("");
      vectorPointOpts.Message = "\nSelect stair orientation: ";
      vectorPointOpts.AllowNone = true;
      vectorPointOpts.UseBasePoint = true;
      vectorPointOpts.UseDashedLine = true;

      // Prepare prompt for the default riser height
      PromptDoubleResult treadDepthResult;
      PromptDistanceOptions treadDepthOpts = new PromptDistanceOptions("");
      treadDepthOpts.Message = "\nEnter the target tread depth: ";
      treadDepthOpts.DefaultValue = 11;
      treadDepthOpts.AllowNegative = false;
      treadDepthOpts.AllowNone = false;
      treadDepthOpts.AllowZero = false;

      // Prepare prompt for the default riser height
      PromptDoubleResult riserHeightResult;
      PromptDistanceOptions riserHeightOpts = new PromptDistanceOptions("");
      riserHeightOpts.Message = "\nEnter the target riser height: ";
      riserHeightOpts.DefaultValue = 7;
      riserHeightOpts.AllowNegative = false;
      riserHeightOpts.AllowNone = false;
      riserHeightOpts.AllowZero = false;

      // Prepare prompt for the stair tread overlap
      PromptDoubleResult treadOverlapResult;
      PromptDistanceOptions treadOverlapOpts = new PromptDistanceOptions("");
      treadOverlapOpts.Message = "\nEnter the tread overlap distance: ";
      treadOverlapOpts.DefaultValue = 0;
      treadOverlapOpts.AllowNegative = false;
      treadOverlapOpts.AllowNone = false;
      treadOverlapOpts.AllowZero = true;

      // Prompt for stair height
      heightRes = doc.Editor.GetDistance(heightOpts);
      if (heightRes.Status != PromptStatus.OK) return;
      double height = heightRes.Value;

      // Prompt for stair width
      widthRes = doc.Editor.GetDistance(widthOpts);
      if (widthRes.Status != PromptStatus.OK) return;
      double width = widthRes.Value;

      // Create stair object
      Stair staircase = new Stair(db,
          height,
          width);

      // Prompt for other options
      do
      {
        othersRes = doc.Editor.GetKeywords(othersOpts);
        if (othersRes.Status == PromptStatus.Cancel) return;
        if (othersRes.Status == PromptStatus.OK)
        {
          switch (othersRes.StringResult)
          {
            case "TopPoint":
              topPointResult = doc.Editor.GetPoint(topPointOpts);
              staircase.stairTopPoint = topPointResult.Value;
              vectorPointOpts.BasePoint = topPointResult.Value;
              vectorPointResult = doc.Editor.GetPoint(vectorPointOpts);
              staircase.stairLengthVector = vectorPointResult.Value - topPointResult.Value;
              if (!Utils.VerifyOrthogonalAngle(Utils.PolarAnglePhi(staircase._lengthVector)))
              {
                Application.ShowAlertDialog("Invalid length vector: must be orthogonal.");
                return;
              }
              othersOpts = new PromptKeywordOptions("");
              othersOpts.Message = "\nOptions: ";
              othersOpts.Keywords.Add("RiserHeight");
              othersOpts.Keywords.Add("TreadDepth");
              othersOpts.Keywords.Add("Overlap");
              othersOpts.AllowArbitraryInput = false;
              othersOpts.AllowNone = true;
              break;
            case "BottomPoint":
              bottomPointResult = doc.Editor.GetPoint(bottomPointOpts);
              staircase.stairBasePoint = bottomPointResult.Value;
              vectorPointOpts.BasePoint = bottomPointResult.Value;
              vectorPointResult = doc.Editor.GetPoint(vectorPointOpts);
              staircase.stairLengthVector = bottomPointResult.Value - vectorPointResult.Value;
              if (!Utils.VerifyOrthogonalAngle(Utils.PolarAnglePhi(staircase._lengthVector)))
              {
                Application.ShowAlertDialog("Invalid length vector: must be orthogonal.");
                return;
              }
              othersOpts = new PromptKeywordOptions("");
              othersOpts.Message = "\nOptions: ";
              othersOpts.Keywords.Add("RiserHeight");
              othersOpts.Keywords.Add("TreadDepth");
              othersOpts.Keywords.Add("Overlap");
              othersOpts.AllowArbitraryInput = false;
              othersOpts.AllowNone = true;
              break;
            case "TreadDepth":
              treadDepthResult = doc.Editor.GetDistance(treadDepthOpts);
              if (treadDepthResult.Value < 8)
              {
                Application.ShowAlertDialog("Invalid tread depth: too narrow.");
                return;
              }
              else if (treadDepthResult.Value > 12)
              {
                Application.ShowAlertDialog("Invalid tread depth: too deep.");
                return;
              }
              else
              {
                staircase.stairDepth = treadDepthResult.Value;
              }
              break;
            case "RiserHeight":
              riserHeightResult = doc.Editor.GetDistance(riserHeightOpts);
              if (riserHeightResult.Value < 5)
              {
                Application.ShowAlertDialog("Invalid riser height: too shallow.");
                return;
              }
              else if (riserHeightResult.Value > 10)
              {
                Application.ShowAlertDialog("Invalid riser height: too steep.");
                return;
              }
              else
              {
                staircase.defaultRiserHeight = riserHeightResult.Value;
              }
              break;
            case "Overlap":
              treadOverlapResult = doc.Editor.GetDistance(treadOverlapOpts);
              if (treadOverlapResult.Value > staircase.stairDepth)
              {
                Application.ShowAlertDialog("Invalid overlap: too large.");
                return;
              }
              else
              {
                staircase.treadOverlap = treadOverlapResult.Value;
              }
              break;
            default:
              Application.ShowAlertDialog("Invalid Keyword");
              break;
          }
        }
      } while (othersRes.Status != PromptStatus.None);

      // Build stairs
      staircase.Build();
    }

  }

  class Ladder
  {
    // Auto-impl class properties
    public Database db { get; private set; }
    public double height { get; private set; }
    public double width { get; private set; }
    public double riserHeight { get; private set; }
    public double treadDepth { get; set; }
    public double length { get; private set; }
    public double stringerWidth { get; set; }
    public double stringerDepth { get; set; }
    public double treadHeight { get; set; }

    // Vector in stair "right" direction (looking up stairs)
    public Vector3d widthVector { get; set; }
    public string layerName { get; set; }

    // Other properties
    private int _nRisers;
    private double _defaultRiserHeight;
    private double _ladderAngle;
    private Point3d _basePoint;
    private Point3d _topPoint;
    private Vector3d _lengthVector;

    public int numRisers
    {
      get { return _nRisers; }
      set
      {
        this._nRisers = value;

        // Calculate stair height
        this.riserHeight = height / value;

      }
    }

    public double defaultRiserHeight
    {
      get { return _defaultRiserHeight; }
      set
      {
        this._defaultRiserHeight = value;

        // Calculate number of stairs
        Application.ShowAlertDialog("Height: " + height + "\n" +
            "Riser Default: " + value);
        this.numRisers = Convert.ToInt32(
            Math.Max(1, Math.Round(height / value))
            );
      }
    }

    public double angle
    {
      get { return Utils.ToDegrees(_ladderAngle); }
      set
      {
        this._ladderAngle = Utils.ToRadians(value);

        // Calculate ladder horizontal length
        this.length = height / Math.Tan(_ladderAngle);
      }
    }

    public Point3d ladderTopPoint
    {
      get { return _topPoint; }
      set
      {
        this._topPoint = value;
        this._basePoint = _topPoint.Add(new Vector3d(
            (length * ladderLengthVector.X),
            (length * ladderLengthVector.Y),
            -height));
      }
    }

    public Point3d ladderBasePoint
    {
      get { return _basePoint; }
      set
      {
        this._basePoint = value;
        this._topPoint = _basePoint.Add(new Vector3d(
            (length * -ladderLengthVector.X),
            (length * -ladderLengthVector.Y),
            height));
      }
    }

    // Vector in stair "down" direction
    public Vector3d ladderLengthVector
    {
      get { return _lengthVector; }
      set
      {
        this._lengthVector = value.UnitVector();
        this._basePoint = _topPoint.Add(new Vector3d(
            (length * ladderLengthVector.X),
            (length * ladderLengthVector.Y),
            -height));
        this.widthVector = new Vector3d(
            Math.Round(ladderLengthVector.RotateBy(Math.PI / 2, Vector3d.ZAxis).X, 0),
            Math.Round(ladderLengthVector.RotateBy(Math.PI / 2, Vector3d.ZAxis).Y, 0),
            Math.Round(ladderLengthVector.RotateBy(Math.PI / 2, Vector3d.ZAxis).Z, 0));
      }
    }

    // Public constructor
    public Ladder(Database db,
        double height,
        double width)
    {
      Color layerColor;
      this.db = db;
      this.height = height;
      this.width = width;

      // Set some defaults
      this.treadDepth = 5;
      this.defaultRiserHeight = 10;
      this.stringerWidth = 1;
      this.stringerDepth = 6;
      this.treadHeight = 0.75;
      this.angle = 75;
      this._lengthVector = -Vector3d.XAxis;
      this.ladderBasePoint = Point3d.Origin;
      this.widthVector = -Vector3d.YAxis;

      // Create beam layer (if necessary)
      this.layerName = "3D-Mezz-Egress";
      layerColor = Utils.ChooseColor("teal");
      Utils.CreateLayer(db, layerName, layerColor);

    }

    public void Build()
    {
      // Declare variables
      Solid3d stringer;
      Solid3d tread;
      Vector3d locationVector;
      Entity tempEnt;

      using (Transaction acTrans = db.TransactionManager.StartTransaction())
      {
        // Open the Block table for read
        BlockTable acBlkTbl;
        acBlkTbl = acTrans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;

        // Open the Block table record Model space for write
        BlockTableRecord modelBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
            OpenMode.ForWrite) as BlockTableRecord;

        // Create stringer
        stringer = CreateStringer(height, length, _ladderAngle, riserHeight,
            stringerWidth, stringerDepth, ladderLengthVector, widthVector);
        stringer.Layer = layerName;

        // Create stair tread
        tread = CreateTread(width, stringerWidth, treadDepth, treadHeight,
            ladderLengthVector, widthVector);
        tread.Layer = layerName;

        // Add first stringer to model and transaction
        locationVector = ladderBasePoint - Point3d.Origin;
        tempEnt = stringer.GetTransformedCopy(Matrix3d.Displacement(locationVector));
        modelBlkTblRec.AppendEntity(tempEnt);
        acTrans.AddNewlyCreatedDBObject(tempEnt, true);

        // Add second stringer to model and transaction
        tempEnt = stringer.GetTransformedCopy(Matrix3d.Displacement(locationVector.Add(
            new Vector3d(((width - stringerWidth) * widthVector.X),
                ((width - stringerWidth) * widthVector.Y),
                0))));
        modelBlkTblRec.AppendEntity(tempEnt);
        acTrans.AddNewlyCreatedDBObject(tempEnt, true);

        Application.ShowAlertDialog("Risers: " + numRisers + "\n" +
            "height: " + height + "\n" +
            "riser height: " + riserHeight + "\n" +
            "step: " + ((height - riserHeight) / Math.Tan(_ladderAngle)));

        // Loop over stair treads
        for (int i = 1; i < numRisers; i++)
        {
          tempEnt = tread.GetTransformedCopy(Matrix3d.Displacement(locationVector.Add(
              new Vector3d(
                  (stringerWidth * widthVector.X) +
                  ((i - 1) * ((height - riserHeight) / Math.Tan(_ladderAngle) / (numRisers - 1)) * -ladderLengthVector.X) +
                  ((ladderLengthVector.CrossProduct(widthVector).X) * (i * riserHeight)),
                  (stringerWidth * widthVector.Y) +
                  ((i - 1) * ((height - riserHeight) / Math.Tan(_ladderAngle) / (numRisers - 1)) * -ladderLengthVector.Y) +
                  ((ladderLengthVector.CrossProduct(widthVector).Y) * (i * riserHeight)),
                  (stringerWidth * widthVector.Z) +
                  ((i - 1) * ((height - riserHeight) / Math.Tan(_ladderAngle) / (numRisers - 1)) * -ladderLengthVector.Z) +
                  ((ladderLengthVector.CrossProduct(widthVector).Z) * (i * riserHeight))))));
          modelBlkTblRec.AppendEntity(tempEnt);
          acTrans.AddNewlyCreatedDBObject(tempEnt, true);
        }

        // Save the transaction
        acTrans.Commit();

      }

    }

    private static Solid3d CreateStringer(double height,
        double length,
        double stairAngle,
        double stairHeight,
        double stringerWidth,
        double stringerDepth,
        Vector3d lengthVector,
        Vector3d widthVector)
    {
      // Create polyline of stringer profile (laying down in sector 1 of X/Y)
      Polyline stringerPoly = new Polyline();
      stringerPoly.AddVertexAt(0, new Point2d(0, 0), 0, 0, 0);
      stringerPoly.AddVertexAt(1, new Point2d(0, stairHeight), 0, 0, 0);
      stringerPoly.AddVertexAt(2, new Point2d((height - stairHeight) / Math.Tan(stairAngle), height), 0, 0, 0);
      stringerPoly.AddVertexAt(3, new Point2d(length, height), 0, 0, 0);
      stringerPoly.AddVertexAt(4, new Point2d(length, height - stringerDepth), 0, 0, 0);
      stringerPoly.AddVertexAt(5, new Point2d((height - stringerDepth) / Math.Tan(stairAngle), height - stringerDepth), 0, 0, 0);
      stringerPoly.AddVertexAt(6, new Point2d(stringerDepth, stringerDepth * Math.Tan(stairAngle)), 0, 0, 0);
      stringerPoly.AddVertexAt(7, new Point2d(stringerDepth, 0), 0, 0, 0);
      stringerPoly.Closed = true;

      // Create the stringer (into positive Z)
      Solid3d stringer = Utils.ExtrudePolyline(stringerPoly, stringerWidth);

      // Position the stringer vertically
      stringer.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.XAxis, Point3d.Origin));

      // Position stringer based on width vector (if width vector is not default/right)
      if ((lengthVector.CrossProduct(widthVector)) == -Vector3d.ZAxis)
      {
        stringer.TransformBy(Matrix3d.Displacement(new Vector3d(0, stringerWidth, 0)));
      }

      // Rotate the stringer
      stringer.TransformBy(Matrix3d.Rotation(
          Vector3d.XAxis.GetAngleTo(-lengthVector, Vector3d.ZAxis),
          lengthVector.CrossProduct(widthVector),
          Point3d.Origin));

      return stringer;
    }

    private static Solid3d CreateTread(double width,
        double stringerWidth,
        double stairDepth,
        double treadHeight,
        Vector3d lengthVector,
        Vector3d widthVector)
    {
      // Calculate tread width
      double treadWidth = width - (2 * stringerWidth);

      // Create polyline of tread profile
      Polyline treadPoly = new Polyline();
      treadPoly.AddVertexAt(0, new Point2d(0, 0), 0, 0, 0);
      treadPoly.AddVertexAt(1, new Point2d(0, -treadHeight), 0, 0, 0);
      treadPoly.AddVertexAt(2, new Point2d(stairDepth, -treadHeight), 0, 0, 0);
      treadPoly.AddVertexAt(3, new Point2d(stairDepth, 0), 0, 0, 0);
      treadPoly.Closed = true;

      // Create the tread
      Solid3d tread = Utils.ExtrudePolyline(treadPoly, treadWidth);

      // Position the tread vertically
      tread.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.XAxis, Point3d.Origin));

      // Position tread based on width vector
      if (lengthVector.CrossProduct(widthVector) == -Vector3d.ZAxis)
      {
        tread.TransformBy(Matrix3d.Displacement(new Vector3d(0, treadWidth, 0)));
      }

      // Rotate the tread
      tread.TransformBy(Matrix3d.Rotation(
          Vector3d.XAxis.GetAngleTo(-lengthVector, Vector3d.ZAxis), Vector3d.ZAxis, Point3d.Origin));

      return tread;
    }

    //private static Solid3d CreateRail(double height,
    //    double length,
    //    double stairAngle,
    //    double stairHeight,
    //    double stringerWidth,
    //    double stringerDepth,
    //    Vector3d lengthVector,
    //    Vector3d widthVector)
    //{
    //    // Create polyline of main rail profile (laying down in sector 1 of X/Y)
    //    Polyline stringerPoly = new Polyline();
    //    stringerPoly.AddVertexAt(0, new Point2d(0, 0), 0, 0, 0);
    //    stringerPoly.AddVertexAt(1, new Point2d(0, stairHeight), 0, 0, 0);
    //    stringerPoly.AddVertexAt(2, new Point2d((height - stairHeight) / Math.Tan(stairAngle), height), 0, 0, 0);
    //    stringerPoly.AddVertexAt(3, new Point2d(length, height), 0, 0, 0);
    //    stringerPoly.AddVertexAt(4, new Point2d(length, height - stringerDepth), 0, 0, 0);
    //    stringerPoly.AddVertexAt(5, new Point2d((height - stringerDepth) / Math.Tan(stairAngle), height - stringerDepth), 0, 0, 0);
    //    stringerPoly.AddVertexAt(6, new Point2d(stringerDepth, stringerDepth * Math.Tan(stairAngle)), 0, 0, 0);
    //    stringerPoly.AddVertexAt(7, new Point2d(stringerDepth, 0), 0, 0, 0);
    //    stringerPoly.Closed = true;

    //    // Create the stringer (into positive Z)
    //    Solid3d stringer = Utils.ExtrudePolyline(stringerPoly, stringerWidth);

    //    // Position the stringer vertically
    //    stringer.TransformBy(Matrix3d.Rotation(Math.PI / 2, Vector3d.XAxis, Point3d.Origin));

    //    // Position stringer based on width vector (if width vector is not default/right)
    //    if ((lengthVector.CrossProduct(widthVector)) == -Vector3d.ZAxis)
    //    {
    //        stringer.TransformBy(Matrix3d.Displacement(new Vector3d(0, stringerWidth, 0)));
    //    }

    //    // Rotate the stringer
    //    stringer.TransformBy(Matrix3d.Rotation(
    //        Vector3d.XAxis.GetAngleTo(-lengthVector, Vector3d.ZAxis),
    //        lengthVector.CrossProduct(widthVector),
    //        Point3d.Origin));

    //    return stringer;
    //}

    [CommandMethod("LadderPrompt")]
    public static void LadderPrompt()
    {
      // Get the current document and database, and start a transaction
      Document doc = Application.DocumentManager.MdiActiveDocument;
      Database db = doc.Database;

      // Prepare prompt for the stair height
      PromptDoubleResult heightRes;
      PromptDistanceOptions heightOpts = new PromptDistanceOptions("");
      heightOpts.Message = "\nEnter the stair height: ";
      heightOpts.DefaultValue = 108;

      // Prepare prompt for the stair width
      PromptDoubleResult widthRes;
      PromptDistanceOptions widthOpts = new PromptDistanceOptions("");
      widthOpts.Message = "\nEnter the stair width: ";
      widthOpts.DefaultValue = 24;

      // Prepare prompt for other options
      PromptResult othersRes;
      PromptKeywordOptions othersOpts = new PromptKeywordOptions("");
      othersOpts.Message = "\nOptions: ";
      othersOpts.Keywords.Add("TopPoint");
      othersOpts.Keywords.Add("BottomPoint");
      othersOpts.Keywords.Add("RiserHeight");
      othersOpts.Keywords.Add("Angle");
      othersOpts.AllowArbitraryInput = false;
      othersOpts.AllowNone = true;

      // Prepare prompt for top point
      PromptPointResult topPointResult;
      PromptPointOptions topPointOpts = new PromptPointOptions("");
      topPointOpts.Message = "\nSelect top point: ";
      topPointOpts.AllowNone = true;

      // Prepare prompt for bottom point
      PromptPointResult bottomPointResult;
      PromptPointOptions bottomPointOpts = new PromptPointOptions("");
      bottomPointOpts.Message = "\nSelect bottom point: ";
      bottomPointOpts.AllowNone = true;

      // Prepare prompt for bottom point
      PromptPointResult vectorPointResult;
      PromptPointOptions vectorPointOpts = new PromptPointOptions("");
      vectorPointOpts.Message = "\nSelect stair orientation: ";
      vectorPointOpts.AllowNone = true;
      vectorPointOpts.UseBasePoint = true;
      vectorPointOpts.UseDashedLine = true;

      // Prepare prompt for the default riser height
      PromptDoubleResult riserHeightResult;
      PromptDistanceOptions riserHeightOpts = new PromptDistanceOptions("");
      riserHeightOpts.Message = "\nEnter the target riser height: ";
      riserHeightOpts.DefaultValue = 7;
      riserHeightOpts.AllowNegative = false;
      riserHeightOpts.AllowNone = false;
      riserHeightOpts.AllowZero = false;

      // Prepare prompt for the stair tread overlap
      PromptDoubleResult angleResult;
      PromptDoubleOptions angleOpts = new PromptDoubleOptions("");
      angleOpts.Message = "\nEnter the ladder angle: ";
      angleOpts.DefaultValue = 75;
      angleOpts.AllowNegative = false;
      angleOpts.AllowNone = false;
      angleOpts.AllowZero = false;

      // Prompt for stair height
      heightRes = doc.Editor.GetDistance(heightOpts);
      if (heightRes.Status != PromptStatus.OK) return;
      double height = heightRes.Value;

      // Prompt for stair width
      widthRes = doc.Editor.GetDistance(widthOpts);
      if (widthRes.Status != PromptStatus.OK) return;
      double width = widthRes.Value;

      // Create stair object
      Ladder shipladder = new Ladder(db,
          height,
          width);

      // Prompt for other options
      do
      {
        othersRes = doc.Editor.GetKeywords(othersOpts);
        if (othersRes.Status == PromptStatus.Cancel) return;
        if (othersRes.Status == PromptStatus.OK)
        {
          switch (othersRes.StringResult)
          {
            case "TopPoint":
              topPointResult = doc.Editor.GetPoint(topPointOpts);
              shipladder.ladderTopPoint = topPointResult.Value;
              vectorPointOpts.BasePoint = topPointResult.Value;
              vectorPointResult = doc.Editor.GetPoint(vectorPointOpts);
              shipladder.ladderLengthVector = vectorPointResult.Value - topPointResult.Value;
              if (!Utils.VerifyOrthogonalAngle(Utils.PolarAnglePhi(shipladder._lengthVector)))
              {
                Application.ShowAlertDialog("Invalid length vector: must be orthogonal.");
                return;
              }
              othersOpts = new PromptKeywordOptions("");
              othersOpts.Message = "\nOptions: ";
              othersOpts.Keywords.Add("RiserHeight");
              othersOpts.Keywords.Add("TreadOverlap");
              othersOpts.AllowArbitraryInput = false;
              othersOpts.AllowNone = true;
              break;
            case "BottomPoint":
              bottomPointResult = doc.Editor.GetPoint(bottomPointOpts);
              shipladder.ladderBasePoint = bottomPointResult.Value;
              vectorPointOpts.BasePoint = bottomPointResult.Value;
              vectorPointResult = doc.Editor.GetPoint(vectorPointOpts);
              shipladder.ladderLengthVector = bottomPointResult.Value - vectorPointResult.Value;
              if (!Utils.VerifyOrthogonalAngle(Utils.PolarAnglePhi(shipladder._lengthVector)))
              {
                Application.ShowAlertDialog("Invalid length vector: must be orthogonal.");
                return;
              }
              othersOpts = new PromptKeywordOptions("");
              othersOpts.Message = "\nOptions: ";
              othersOpts.Keywords.Add("RiserHeight");
              othersOpts.Keywords.Add("TreadOverlap");
              othersOpts.AllowArbitraryInput = false;
              othersOpts.AllowNone = true;
              break;
            case "RiserHeight":
              riserHeightResult = doc.Editor.GetDistance(riserHeightOpts);
              if (riserHeightResult.Value < 9)
              {
                Application.ShowAlertDialog("Invalid riser height: too shallow.");
                return;
              }
              else if (riserHeightResult.Value > 11)
              {
                Application.ShowAlertDialog("Invalid riser height: too steep.");
                return;
              }
              else
              {
                shipladder.defaultRiserHeight = riserHeightResult.Value;
              }
              break;
            case "Angle":
              angleResult = doc.Editor.GetDouble(angleOpts);
              if (angleResult.Value > 80)
              {
                Application.ShowAlertDialog("Invalid angle: too steep.");
                return;
              }
              else if (angleResult.Value < 50)
              {
                Application.ShowAlertDialog("Invalid angle: too shallow.");
              }
              else
              {
                shipladder.angle = angleResult.Value;
              }
              break;
            default:
              Application.ShowAlertDialog("Invalid Keyword");
              break;
          }
        }
      } while (othersRes.Status != PromptStatus.None);

      // Build stairs
      shipladder.Build();
    }

  }


}
