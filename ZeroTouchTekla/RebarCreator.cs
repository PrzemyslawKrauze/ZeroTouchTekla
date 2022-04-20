using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.IO;
using Tekla.Structures.Filtering;
using Tekla.Structures.Filtering.Categories;
using Tekla.Structures.Model;
using Tekla.Structures.Model.Operations;
using ZeroTouchTekla.Profiles;


namespace ZeroTouchTekla
{
    class RebarCreator
    {
        public static void Test()
        {
            Model model = new Model();
            ModelInfo info = model.GetInfo();

            // Creates the filter expressions
            Tekla.Structures.Model.UI.Picker picker = new Tekla.Structures.Model.UI.Picker();
            Tekla.Structures.Model.UI.Picker.PickObjectsEnum pickObjectEnum = Tekla.Structures.Model.UI.Picker.PickObjectsEnum.PICK_N_PARTS;
            ModelObjectEnumerator modelObject = picker.PickObjects(pickObjectEnum);

        }
        public static void CreateForPart(ProfileType profileType)
        {
            Model model = new Model();

            //Store current work plane
            TransformationPlane currentPlane = model.GetWorkPlaneHandler().GetCurrentTransformationPlane();
            Element element;

            switch (profileType)
            {
                case ProfileType.FTG:
                    element = new FTG(PickPart());
                    break;
                case ProfileType.RTW:
                    element = new RTW(PickPart());
                    break;
                case ProfileType.DRTW:
                    element = new DRTW(PickParts());
                    break;
                case ProfileType.RTWS:
                    element = new RTWS(PickPart());
                    break;
                case ProfileType.CLMN:
                    element = new CLMN(PickPart());
                    break;
                case ProfileType.ABT:
                    element = new ABT(PickParts());
                    break;
                case ProfileType.APS:
                    element = new APS(PickParts());
                    break;
                default:
                    throw new Exception("Profile type doesn't match");
            }
            element.Create();
            //Restore user work plane
            model.GetWorkPlaneHandler().SetCurrentTransformationPlane(currentPlane);
            model.CommitChanges();

            ChangeLayer(model,element);
            ChangeLayer(model,element);
        }
        static Part PickPart()
        {
            Tekla.Structures.Model.UI.Picker picker = new Tekla.Structures.Model.UI.Picker();
            Tekla.Structures.Model.UI.Picker.PickObjectEnum pickObjectEnum = Tekla.Structures.Model.UI.Picker.PickObjectEnum.PICK_ONE_PART;
            Part modelObject = picker.PickObject(pickObjectEnum) as Part;
            return modelObject;
        }
        static List<Part> PickParts()
        {
            Tekla.Structures.Model.UI.Picker picker = new Tekla.Structures.Model.UI.Picker();
            Tekla.Structures.Model.UI.Picker.PickObjectsEnum pickObjectEnum = Tekla.Structures.Model.UI.Picker.PickObjectsEnum.PICK_N_PARTS;
            ModelObjectEnumerator modelObjects = picker.PickObjects(pickObjectEnum);
            List<ModelObject> modelObjectList = Utility.ToList(modelObjects);
            List<Part> beamList = new List<Part>();
            foreach (ModelObject mo in modelObjectList)
            {
                beamList.Add(mo as Part);
            }
            return beamList;
        }
        public static void CreateForComponent(ProfileType profileType)
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
                        case ProfileType.WING:
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
            Model model = new Model();
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
                ProfileType profileType = GetProfileType(hostName);
                if (profileType != ProfileType.None)
                {
                    Type[] Types = new Type[] { typeof(RebarSet) };
                    ModelObjectEnumerator moe = model.GetModelObjectSelector().GetAllObjectsWithType(Types);
                    var rebarList = Utility.ToList(moe);


                    List<RebarSet> selectedRebars = (from RebarSet r in rebarList
                                                     where Utility.GetUserProperty(r, FATHER_ID_NAME) == beam.Identifier.ID
                                                     select r).ToList();
                    Dictionary<int,int[]> currentLayerDictionary = new Dictionary<int,int[]>();
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
                    model.CommitChanges();

                    // RebarCreator rebarCreator = new RebarCreator();
                    //Store current work plane
                    TransformationPlane currentPlane = model.GetWorkPlaneHandler().GetCurrentTransformationPlane();
                    //Get beam local plane
                    TransformationPlane localPlane = new TransformationPlane(beam.GetCoordinateSystem());
                    model.GetWorkPlaneHandler().SetCurrentTransformationPlane(localPlane);

                    Element element;
                    switch (profileType)
                    {
                        case ProfileType.FTG:
                            element = new FTG(beam);
                            break;
                        case ProfileType.RTW:
                           element= new RTW(beam);
                            break;
                        case ProfileType.CLMN:
                            element = new CLMN(beam);
                            break;
                        default:
                            throw new Exception("Profile type doesn't match");
                    }

                    element.CreateSingle(rebarName);
                    //Restore user work plane
                    model.GetWorkPlaneHandler().SetCurrentTransformationPlane(currentPlane);
                    model.CommitChanges();

                    ChangeLayer(model,element);
                    ChangeLayer(model,element);
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

       
        public enum ProfileType
        {
            None,
            FTG,
            RTW,
            DRTW,
            RTWS,
            CLMN,
            ABT,
            TABT,
            WING,
            APS
        }
        public static int FatherID;
        static ProfileType GetProfileType(string profileString)
        {
            if (profileString.Contains("FTG"))
            {
                return ProfileType.FTG;
            }
            else
            {
                if (profileString.Contains("RTW"))
                {
                    return ProfileType.RTW;
                }
                else
                {
                    if (profileString.Contains("CLMN"))
                    {
                        return ProfileType.CLMN;
                    }
                }
            }
            return ProfileType.None;
        }
        public const string FATHER_ID_NAME = "ZTB_FatherIDName";
        public const string METHOD_NAME = "ZTB_MethodName";
        public const string MethodInput = "ZTB_MethodInput";
        public static int MinLengthCoefficient = 20;
    }
}
