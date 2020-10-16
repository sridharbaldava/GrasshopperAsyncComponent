using Grasshopper.Kernel;
using Rhino.Geometry;
using SAP2000v1;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Grasshopper.GUI;
using Grasshopper.GUI.Canvas;
using Grasshopper.Kernel.Attributes;
using System.Drawing;
using System.Windows.Forms;


namespace GrasshopperAsyncComponent.SampleImplementations
{
    public class Sample_SAPReadAreaLoads : GH_AsyncComponent
    {
        public override Guid ComponentGuid { get => new Guid("135E67C4-FFED-4DB5-8B5E-016328CA71E1"); }

        protected override System.Drawing.Bitmap Icon { get => null; }

        public override GH_Exposure Exposure => GH_Exposure.primary;

        public override void CreateAttributes()
        {
            m_attributes = new RunButton(this, "Run");
        }

        public Sample_SAPReadAreaLoads() : base("Read SAP Area Loads", "RSAL", "Read SAP Area Loads", "WPM", "SAP2000v1")
        {
            BaseWorker = new SAPReadAreaLoadWorker();
        }
        
        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddBooleanParameter("Execute", "E", "Refresh SAP Data", GH_ParamAccess.item);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddPointParameter("Points", "Pt", "Names of Points Retrieved", GH_ParamAccess.list);
            pManager.AddTextParameter("Area Names", "An", "Names of Area Objects Retrieved", GH_ParamAccess.list);
            pManager.AddCurveParameter("Area Boundaries", "Pl", "Boundary curves of Area Objects Retrieved", GH_ParamAccess.list);
            pManager.AddTextParameter("Load Values", "Lv", "Applied loads to Area Objects Retrieved", GH_ParamAccess.list);
        }
    }

    public class SAPReadAreaLoadWorker : WorkerInstance
    {
        cOAPI sapObject = null;
        cSapModel sapModel = null;
        bool run = false;
        bool attached = false;
        readonly Dictionary<string, Point3d> Joints = new Dictionary<string, Point3d>();
        string[] ObjectNames = null;
        readonly List<string> LoadValues = new List<string>();
        readonly List<Polyline> areaBoundary = new List<Polyline>();
        readonly List<string> AreaNames = new List<string>();

        public override void DoWork(Action<string, double> ReportProgress, Action<string, GH_RuntimeMessageLevel> ReportError, Action Done)
        {
            if ( run )
            {
                ReportProgress("Start", 0.0);
                if ( CancellationToken.IsCancellationRequested ) return;
                AttachToSAP(ReportError);
                
                if ( CancellationToken.IsCancellationRequested ) return;
                ReadJoints();                

                if ( CancellationToken.IsCancellationRequested ) return;
                ReadSAPAreaObjects();
                
                if ( CancellationToken.IsCancellationRequested ) return;
                ReadAreaLoads(ReportProgress);
                Done();
            }            
        }

        public override WorkerInstance Duplicate() => new SAPReadAreaLoadWorker();

        public override void GetData(IGH_DataAccess DA, GH_ComponentParamServer Params)
        {
            bool execute = false;
            DA.GetData(0, ref execute);
            run = execute;
        }

        public override void SetData(IGH_DataAccess DA)
        {
            DA.SetDataList(0, Joints.Values);
            DA.SetDataList(1, AreaNames);
            DA.SetDataList(2, areaBoundary);
            DA.SetDataList(3, LoadValues);
        }

        private void AttachToSAP(Action<string, GH_RuntimeMessageLevel> ReportError)
        {
            try
            {
                sapObject = (cOAPI)Marshal.GetActiveObject("CSI.SAP2000.API.SapObject");
                sapModel = sapObject.SapModel;
                attached = true;
                //Print("Attached to active SAP instance");
            }
            catch ( COMException )
            {
                ReportError("No running instance of the program found or failed to attach.", GH_RuntimeMessageLevel.Error);
                return;
                //throw;
            }
        }

        private void ReadSAPAreaObjects()
        {
            if ( !attached )
                return;

            areaBoundary.Clear();
            AreaNames.Clear();
            int NumberItems = 0;
            sapModel.AreaObj.GetNameList(ref NumberItems, ref ObjectNames);

            foreach ( var objectName in ObjectNames )
            {
                string[] Points = null;
                int numberPoints = 0;
                sapModel.AreaObj.GetPoints(objectName, ref numberPoints, ref Points);
                var boundaryCurve = new Polyline();
                for ( int i = 0; i < numberPoints; i++ )
                {
                    boundaryCurve.Add(Joints[Points[i]]);
                }
                areaBoundary.Add(boundaryCurve);
                AreaNames.Add(objectName);
            }
        }

        private void ReadAreaLoads(Action<string, double> ReportProgress)
        {
            if ( !attached )
                return;
            int NumberItems = 0;
            string[] AreaName = null;
            string[] LoadPat = null;
            string[] CSys = null;
            int[] Dir = { 0 };
            double[] Value = { 0 };
            //int[] DistType = { 0 };
            //sapModel.AreaObj.GetLoadUniformToFrame("ALL", ref NumberItems, ref AreaName, ref LoadPat, ref CSys, ref Dir, ref Value, ref DistType, eItemType.Group);

            LoadValues.Clear();

            for ( int j = 0; j < AreaNames.Count; j++ )
            {
                if ( CancellationToken.IsCancellationRequested ) return;
                string areaName = AreaNames[j];
                sapModel.AreaObj.GetLoadUniform(areaName, ref NumberItems, ref AreaName, ref LoadPat, ref CSys, ref Dir, ref Value, eItemType.Objects);
                for ( int i = 0; i < NumberItems; i++ )
                {
                    LoadValues.Add(string.Join(",", AreaName[i], LoadPat[i], CSys[i], Dir[i], Value[i]));                    
                }
                ReportProgress(AreaNames[j], ((double)j) / AreaNames.Count);
            }
        }

        private void ReadJoints()
        {
            if ( !attached )
                return;
            Joints.Clear();
            int numberJoints = 0;
            string[] jointNames = null;
            sapModel.PointObj.GetNameList(ref numberJoints, ref jointNames);
            foreach ( var jointName in jointNames )
            {
                double x = double.NaN, y = double.NaN, z = double.NaN;
                sapModel.PointObj.GetCoordCartesian(jointName, ref x, ref y, ref z);
                var coords = new Point3d(x, y, z);
                Joints.Add(jointName, coords);
            }
        }
    }

    public class RunButton : GH_ComponentAttributes
    {

        private string _text;

        public bool Activate { get; set; }

        public string DisplayText
        {
            get
            {
                if ( string.IsNullOrEmpty(this._text) )
                    return "Export";
                return this._text;
            }
            set
            {
                this._text = value;
            }
        }

        public RunButton(GH_Component owner, string displayText)
          : base((IGH_Component)owner)
        {
            this.Activate = false;
            this.DisplayText = displayText;
        }

        protected override void Layout()
        {
            base.Layout();
            Rectangle rectangle1 = GH_Convert.ToRectangle(this.Bounds);
            rectangle1.Height += 22;
            Rectangle rectangle2 = rectangle1;
            rectangle2.Y = rectangle2.Bottom - 22;
            rectangle2.Height = 22;
            rectangle2.Inflate(-2, -2);
            this.Bounds = (RectangleF)rectangle1;
            this.ButtonBounds = rectangle2;
        }

        private Rectangle ButtonBounds { get; set; }

        protected override void Render(GH_Canvas canvas, Graphics graphics, GH_CanvasChannel channel)
        {
            base.Render(canvas, graphics, channel);
            if ( channel != GH_CanvasChannel.Objects )
                return;
            GH_Capsule ghCapsule = this.Activate ? GH_Capsule.CreateTextCapsule(this.ButtonBounds, this.ButtonBounds, GH_Palette.Grey, this.DisplayText, 2, 0) : GH_Capsule.CreateTextCapsule(this.ButtonBounds, this.ButtonBounds, GH_Palette.Black, this.DisplayText, 2, 0);
            ghCapsule.Render(graphics, this.Selected, this.Owner.Locked, false);
            ghCapsule.Dispose();
        }

        public override GH_ObjectResponse RespondToMouseUp(GH_Canvas sender, GH_CanvasMouseEvent e)
        {
            if ( e.Button != MouseButtons.Left || !((RectangleF)this.ButtonBounds).Contains(e.CanvasLocation) )
                return base.RespondToMouseUp(sender, e);
            this.Owner.ExpireSolution(true);
            this.Activate = false;
            sender.Refresh();
            return GH_ObjectResponse.Release;
        }

        public override GH_ObjectResponse RespondToMouseDown(GH_Canvas sender, GH_CanvasMouseEvent e)
        {
            if ( e.Button != MouseButtons.Left || !((RectangleF)this.ButtonBounds).Contains(e.CanvasLocation) )
                return base.RespondToMouseDown(sender, e);
            this.Activate = true;
            sender.Refresh();
            return GH_ObjectResponse.Capture;
        }
    }

}
