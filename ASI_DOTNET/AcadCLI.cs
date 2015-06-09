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

[assembly: CommandClass(typeof(ASI_DOTNET.AcadCLI))]

namespace ASI_DOTNET
{
    class AcadCLI
    {
        [CommandMethod("RackPrompt")]
        public static void RackPrompt()
        {
            // Get the current document and database, and start a transaction
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            Editor acEd = acDoc.Editor;

            // Create rack instance
            Rack rack = new Rack(acCurDb);

            // Get rack details
            if (RackOptions(acEd, rack) == false) { return; }

            // Prompt for other rack options
            if (RackOtherOptions(acEd, rack) == false) { return; }

            // Build rack
            rack.Build();
        }

        public static bool RackOptions(Editor ed, Rack rack)
        {
            // Prepare prompt for rack type
            PromptResult typeRes;
            PromptKeywordOptions typeOpts = new PromptKeywordOptions("");
            typeOpts.Message = "\nRack Type: ";
            typeOpts.Keywords.Add("PAllet");
            typeOpts.Keywords.Add("Flow");
            typeOpts.Keywords.Add("PUshback");
            typeOpts.Keywords.Add("DriveIn");
            typeOpts.Keywords.Default = "PAllet";
            typeOpts.AllowArbitraryInput = false;
            typeOpts.AllowNone = true;

            // Prepare prompt for rack depth
            PromptIntegerResult deepRes;
            PromptIntegerOptions deepOpts = new PromptIntegerOptions("");
            deepOpts.Message = "\nEnter the rack depth: ";
            deepOpts.DefaultValue = 1;
            deepOpts.AllowZero = false;
            deepOpts.AllowNegative = false;

            // Prepare prompt for rack bays
            PromptIntegerResult baysRes;
            PromptIntegerOptions baysOpts = new PromptIntegerOptions("");
            baysOpts.Message = "\nEnter the number of rack bays: ";
            baysOpts.DefaultValue = 1;
            baysOpts.AllowZero = false;
            baysOpts.AllowNegative = false;

            // Prepare prompt for rack levels
            PromptIntegerResult highRes;
            PromptIntegerOptions highOpts = new PromptIntegerOptions("");
            highOpts.Message = "\nEnter the number of rack levels: ";
            highOpts.DefaultValue = 1;
            highOpts.AllowZero = false;
            highOpts.AllowNegative = false;

            // Prompt for rack type
            typeRes = ed.GetKeywords(typeOpts);
            if (typeRes.Status == PromptStatus.Cancel) return false;
            rack.rackType = typeRes.StringResult;

            // Prompt for rack depth
            if (rack.rackType != "pallet")
            {
                deepRes = ed.GetInteger(deepOpts);
                if (deepRes.Status != PromptStatus.OK) return false;
                rack.rackCountDeep = deepRes.Value;
            }

            // Prompt for rack bays
            baysRes = ed.GetInteger(baysOpts);
            if (baysRes.Status != PromptStatus.OK) return false;
            rack.rackCountBays = baysRes.Value;

            // Prompt for rack levels
            highRes = ed.GetInteger(highOpts);
            if (highRes.Status != PromptStatus.OK) return false;
            rack.rackCountHigh = highRes.Value;

            return true;
        }

        public static bool RackOtherOptions(Editor ed, Rack rack)
        {
            // Prepare prompt for other options
            PromptResult res;
            PromptKeywordOptions opts = new PromptKeywordOptions("");
            opts.Message = "\nOptions: ";
            opts.Keywords.Add("Frame");
            opts.Keywords.Add("Beam");
            opts.Keywords.Add("Spacer");
            opts.AllowArbitraryInput = false;
            opts.AllowNone = true;

            // Prompt for other options
            res = ed.GetKeywords(opts);
            if (res.Status == PromptStatus.Cancel) return false;

            while (res.Status == PromptStatus.OK)
            {
                switch (res.StringResult)
                {
                    case "Frame":
                        if (RackFrameOptions(ed, rack) == false) { return false; }
                        break;
                    case "Beam":
                        if (RackBeamOptions(ed, rack, "beam") == false) { return false; }
                        break;
                    case "Spacer":
                        if (RackBeamOptions(ed, rack, "spacer") == false) { return false; }
                        break;
                    default:
                        Application.ShowAlertDialog("Invalid Keyword");
                        break;
                }

                // Re-prompt for keywords
                res = ed.GetKeywords(opts);
                if (res.Status == PromptStatus.Cancel) return false;
            }

            return true;
        }

        public static bool RackFrameOptions(Editor ed, Rack rack)
        {
            // Prepare prompt for the frame height
            PromptDoubleResult heightRes;
            PromptDistanceOptions heightOpts = new PromptDistanceOptions("");
            heightOpts.Message = "\nEnter the frame height: ";
            heightOpts.DefaultValue = 96;

            // Prepare prompt for the frame width
            PromptDoubleResult widthRes;
            PromptDistanceOptions widthOpts = new PromptDistanceOptions("");
            widthOpts.Message = "\nEnter the frame width: ";
            widthOpts.DefaultValue = 36;

            // Prepare prompt for the frame diameter
            PromptDoubleResult diameterRes;
            PromptDistanceOptions diameterOpts = new PromptDistanceOptions("");
            diameterOpts.Message = "\nEnter the frame width: ";
            diameterOpts.DefaultValue = 3.0;

            // Prompt for the frame height
            heightRes = ed.GetDistance(heightOpts);
            rack.frameHeight = heightRes.Value;
            if (heightRes.Status != PromptStatus.OK) return false;

            // Prompt for the frame width
            widthRes = ed.GetDistance(widthOpts);
            rack.frameWidth = widthRes.Value;
            if (widthRes.Status != PromptStatus.OK) return false;

            // Prompt for the frame diameter
            diameterRes = ed.GetDistance(diameterOpts);
            rack.frameDiameter = diameterRes.Value;
            if (diameterRes.Status != PromptStatus.OK) return false;

            return true;
        }

        public static bool RackBeamOptions(Editor ed, Rack rack, string type)
        {
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

            // Prepare prompt for the beam style
            string style = "Step";
            PromptResult bStyleRes;
            PromptKeywordOptions bStyleOpts = new PromptKeywordOptions("");
            bStyleOpts.Message = "\nEnter beam style: ";
            bStyleOpts.Keywords.Add("Step");
            bStyleOpts.Keywords.Add("Box");
            bStyleOpts.Keywords.Add("IBeam");
            bStyleOpts.Keywords.Add("CChannel");
            bStyleOpts.Keywords.Default = "Step";
            bStyleOpts.AllowArbitraryInput = false;

            // Prompt for beam length
            lengthRes = ed.GetDistance(lengthOpts);
            if (lengthRes.Status != PromptStatus.OK) return false;
            double length = lengthRes.Value;

            // Prompt for beam height
            heightRes = ed.GetDistance(heightOpts);
            if (heightRes.Status != PromptStatus.OK) return false;
            double height = heightRes.Value;

            // Prompt for beam width
            widthRes = ed.GetDistance(widthOpts);
            if (widthRes.Status != PromptStatus.OK) return false;
            double width = widthRes.Value;

            // Prompt for beam style
            if (type == "beam") {
                bStyleRes = ed.GetKeywords(bStyleOpts);
                style = bStyleRes.StringResult;
                if (bStyleRes.Status == PromptStatus.Cancel) return false;
            }

            if (type == "beam")
            {
                rack.beamLength = length;
                rack.beamHeight = height;
                rack.beamWidth = width;
                rack.beamType = style;
            }
            else if (type == "spacer")
            {
                rack.spacerLength = length;
                rack.spacerHeight = height;
                rack.spacerWidth = width;
                rack.spacerType = "Box";
            }
            else if (type == "tie")
            {
                rack.tieLength = length;
                rack.tieHeight = height;
                rack.tieWidth = width;
                rack.tieType = "Box";
            }
            else
            {
                return false;
            }

            return true;
        }
    }
}
