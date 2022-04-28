using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.IO;
using Tekla.Structures.Filtering;
using Tekla.Structures.Filtering.Categories;
using Tekla.Structures.Model;
using Tekla.Structures.Model.Operations;
using Tekla.Structures.Solid;
using Tekla.Structures.Geometry3d;
using ZeroTouchTekla.Profiles;



namespace ZeroTouchTekla
{
    class RebarCreator
    {
        public static void Test()
        {
            //Store current plane
            TransformationPlane currentPlane = Program.ActiveModel.GetWorkPlaneHandler().GetCurrentTransformationPlane();
            //Pick part
            Tekla.Structures.Model.UI.Picker picker = new Tekla.Structures.Model.UI.Picker();
            ModelObject modelObject = picker.PickObject(Tekla.Structures.Model.UI.Picker.PickObjectEnum.PICK_ONE_PART, "Pick part");
            Part part = modelObject as Part;
            TransformationPlane localPlane = new TransformationPlane(part.GetCoordinateSystem());
            Program.ActiveModel.GetWorkPlaneHandler().SetCurrentTransformationPlane(localPlane);

            Face[] faces = TeklaUtils.GetPartEndFaces(modelObject as Part);
            List<List<Point>> points = TeklaUtils.GetPointsFromFaces(faces);
            points = TeklaUtils.SortPoints(points);
            Tekla.Structures.Model.UI.GraphicsDrawer graphicsDrawer = new Tekla.Structures.Model.UI.GraphicsDrawer();
            for(int i=0;i< points.Count;i++)
            {
                List<Point> currentList = points[i];
                for(int j=0;j<currentList.Count;j++)
                {
                    string text = i.ToString() + j.ToString();
                    Tekla.Structures.Model.UI.Color color = new Tekla.Structures.Model.UI.Color();
                    graphicsDrawer.DrawText(currentList[j], text, color);
                }
            }

            //Restore user's plane
            Program.ActiveModel.GetWorkPlaneHandler().SetCurrentTransformationPlane(currentPlane);
            Program.ActiveModel.CommitChanges();
        }
        public static void CreateForPart(Element.ProfileType profileType)
        {
            //Store current work plane
            TransformationPlane currentPlane = Program.ActiveModel.GetWorkPlaneHandler().GetCurrentTransformationPlane();
            Part[] pickedParts = PickParts();
            Element element = Element.Create(pickedParts);

            //Restore user work plane
            Program.ActiveModel.GetWorkPlaneHandler().SetCurrentTransformationPlane(currentPlane);
            Program.ActiveModel.CommitChanges();

            ChangeLayer(Program.ActiveModel, element);
            ChangeLayer(Program.ActiveModel, element);
        }
        static Part[] PickParts()
        {
            Tekla.Structures.Model.UI.Picker picker = new Tekla.Structures.Model.UI.Picker();
            ModelObjectEnumerator modelObjects = picker.PickObjects(Tekla.Structures.Model.UI.Picker.PickObjectsEnum.PICK_N_PARTS, "Pick parts");

            Part[] parts = new Part[modelObjects.GetSize()];

            for (int i = 0; i < parts.Length; i++)
            {
                modelObjects.MoveNext();
                parts[i] = modelObjects.Current as Part;
            }
            return parts;
        }
        public static void CreateForComponent(Element.ProfileType profileType)
        {
            Model model = new Model();
            ModelInfo info = model.GetInfo();
            Tekla.Structures.Model.UI.Picker picker = new Tekla.Structures.Model.UI.Picker();
            Tekla.Structures.Model.UI.Picker.PickObjectEnum pickObjectEnum = Tekla.Structures.Model.UI.Picker.PickObjectEnum.PICK_ONE_OBJECT;
            ModelObject modelObject = picker.PickObject(pickObjectEnum);
            try
            {
                Beam part = modelObject as Beam;
                FatherID = part.Identifier.ID;
                if (part != null)
                {
                    //Store current work plane
                    TransformationPlane currentPlane = model.GetWorkPlaneHandler().GetCurrentTransformationPlane();
                    //Get beam local plane
                    TransformationPlane localPlane = new TransformationPlane(part.GetCoordinateSystem());
                    model.GetWorkPlaneHandler().SetCurrentTransformationPlane(localPlane);
                    Element element;
                    switch (profileType)
                    {
                        case Element.ProfileType.WING:
                            element = new WING(part);
                            break;
                        default:
                            throw new Exception("Profile type doesn't match");
                    }
                    element.Create();
                    //Restore user work plane
                    model.GetWorkPlaneHandler().SetCurrentTransformationPlane(currentPlane);
                    model.CommitChanges();
                    ChangeLayer(model, element);
                    ChangeLayer(model, element);
                }
            }
            catch (System.ApplicationException)
            {
                Operation.DisplayPrompt("User interrupted!");
            }
        }
        public static void RecreateRebar()
        {

            Tekla.Structures.Model.UI.Picker partPicker = new Tekla.Structures.Model.UI.Picker();
            Tekla.Structures.Model.UI.Picker.PickObjectEnum partPickObjectEnum = Tekla.Structures.Model.UI.Picker.PickObjectEnum.PICK_ONE_PART;

            Tekla.Structures.Model.UI.Picker rebarPicker = new Tekla.Structures.Model.UI.Picker();
            Tekla.Structures.Model.UI.Picker.PickObjectEnum rebarPickObjectEnum = Tekla.Structures.Model.UI.Picker.PickObjectEnum.PICK_ONE_REINFORCEMENT;

            try
            {
                Beam beam = partPicker.PickObject(partPickObjectEnum) as Beam;

                RebarSet rebarSet = rebarPicker.PickObject(rebarPickObjectEnum) as RebarSet;
                rebarSet.GetUserProperty(RebarCreator.FATHER_ID_NAME, ref FatherID);
                string rebarName = rebarSet.RebarProperties.Name;

                string hostName = beam.Profile.ProfileString;
                Element.ProfileType profileType = Element.GetProfileType(hostName);
                if (profileType != Element.ProfileType.None)
                {
                    Type[] Types = new Type[] { typeof(RebarSet) };
                    ModelObjectEnumerator moe = Program.ActiveModel.GetModelObjectSelector().GetAllObjectsWithType(Types);
                    var rebarList = Utility.ToList(moe);


                    List<RebarSet> selectedRebars = (from RebarSet r in rebarList
                                                     where Utility.GetUserProperty(r, FATHER_ID_NAME) == beam.Identifier.ID
                                                     select r).ToList();
                    Dictionary<int, int[]> currentLayerDictionary = new Dictionary<int, int[]>();
                    foreach (RebarSet rs in selectedRebars)
                    {
                        List<RebarLegFace> rebarLegFaces = rs.LegFaces;
                        int[] layers = new int[rebarLegFaces.Count];
                        for (int i = 0; i < rebarLegFaces.Count; i++)
                        {
                            layers[i] = rebarLegFaces[i].LayerOrderNumber;
                        }

                        currentLayerDictionary.Add(rs.Identifier.ID, layers);
                    }

                    List<RebarSet> barsToDelete = (from RebarSet rs in selectedRebars
                                                   where rs.RebarProperties.Name == rebarName
                                                   select rs).ToList();

                    foreach (RebarSet rs in barsToDelete)
                    {
                        bool deleted = rs.Delete();
                    }
                    Program.ActiveModel.CommitChanges();

                    // RebarCreator rebarCreator = new RebarCreator();
                    //Store current work plane
                    TransformationPlane currentPlane = Program.ActiveModel.GetWorkPlaneHandler().GetCurrentTransformationPlane();
                    //Get beam local plane
                    TransformationPlane localPlane = new TransformationPlane(beam.GetCoordinateSystem());
                    Program.ActiveModel.GetWorkPlaneHandler().SetCurrentTransformationPlane(localPlane);

                    Element element;
                    switch (profileType)
                    {
                        case Element.ProfileType.FTG:
                            element = new FTG(beam);
                            break;
                        case Element.ProfileType.RTW:
                            element = new RTW(beam);
                            break;
                        case Element.ProfileType.RCLMN:
                            element = new CLMN(beam);
                            break;
                        default:
                            throw new Exception("Profile type doesn't match");
                    }

                    element.CreateSingle(rebarName);
                    //Restore user work plane
                    Program.ActiveModel.GetWorkPlaneHandler().SetCurrentTransformationPlane(currentPlane);
                    Program.ActiveModel.CommitChanges();

                    ChangeLayer(Program.ActiveModel, element);
                    ChangeLayer(Program.ActiveModel, element);
                }
            }
            catch (System.ApplicationException)
            {
                Operation.DisplayPrompt("User interrupted!");
            }

        }
        static void ChangeLayer(Model model, Element element)
        {
            Type[] Types = new Type[] { typeof(RebarSet) };
            ModelObjectEnumerator Enum = model.GetModelObjectSelector().GetAllObjectsWithType(Types);
            var rebarList = Utility.ToList(Enum);

            foreach (KeyValuePair<int, int[]> keyValuePair in element.LayerDictionary)
            {
                RebarSet rs = (from RebarSet r in rebarList
                               where r.Identifier.ID == keyValuePair.Key
                               select r).FirstOrDefault();

                if (rs != null)
                {
                    bool modify = false;
                    int[] layerNumbers = keyValuePair.Value;
                    for (int i = 0; i < rs.LegFaces.Count; i++)
                    {
                        int newLayer = layerNumbers[i];
                        int currentLayer = rs.LegFaces[i].LayerOrderNumber;
                        if (currentLayer != newLayer)
                        {
                            rs.LegFaces[i].LayerOrderNumber = newLayer;
                            modify = true;
                        }
                    }
                    if (modify)
                    {
                        rs.Modify();
                        model.CommitChanges();
                    }
                }
            }
        }



        public static int FatherID;

        public const string FATHER_ID_NAME = "ZTB_FatherIDName";
        public const string METHOD_NAME = "ZTB_MethodName";
        public const string MethodInput = "ZTB_MethodInput";
        public static int MinLengthCoefficient = 20;
    }
}
