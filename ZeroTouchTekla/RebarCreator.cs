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

            switch (profileType)
            {
                case ProfileType.FTG:
                    FTG ftg = new FTG(PickPart());
                    ftg.Create();
                    break;
                case ProfileType.RTW:
                    RTW rtw = new RTW(PickPart());
                    rtw.Create();
                    break;
                case ProfileType.DRTW:
                    DRTW drtw = new DRTW(PickParts());
                    drtw.Create();
                    break;
                case ProfileType.RTWS:
                    RTWS rtws = new RTWS(PickPart());
                    rtws.Create();
                    break;
                case ProfileType.CLMN:
                    CLMN clmn = new CLMN(PickPart());
                    clmn.Create();
                    break;
                case ProfileType.ABT:
                    ABT dabt = new ABT(PickParts());
                    dabt.Create();
                    break;
                case ProfileType.APS:
                    APS aps = new APS(PickParts());
                    aps.Create();
                    break;
            }

            //Restore user work plane
            model.GetWorkPlaneHandler().SetCurrentTransformationPlane(currentPlane);
            model.CommitChanges();

            ChangeLayer(model);
            ChangeLayer(model);
            LayerDictionary = new Dictionary<int, int[]>();

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

                    switch (profileType)
                    {
                        case ProfileType.WING:
                            WING wing = new WING(part);
                            wing.Create();
                            break;

                    }

                    //Restore user work plane
                    model.GetWorkPlaneHandler().SetCurrentTransformationPlane(currentPlane);
                    model.CommitChanges();
                }

                ChangeLayer(model);
                ChangeLayer(model);
                LayerDictionary = new Dictionary<int, int[]>();
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
                rebarSet.GetUserProperty(RebarCreator.FatherIDName, ref FatherID);
                string rebarName = rebarSet.RebarProperties.Name;

                string hostName = beam.Profile.ProfileString;
                ProfileType profileType = GetProfileType(hostName);
                if (profileType != ProfileType.None)
                {
                    Type[] Types = new Type[] { typeof(RebarSet) };
                    ModelObjectEnumerator moe = model.GetModelObjectSelector().GetAllObjectsWithType(Types);
                    var rebarList = Utility.ToList(moe);


                    List<RebarSet> selectedRebars = (from RebarSet r in rebarList
                                                     where Utility.GetUserProperty(r, FatherIDName) == beam.Identifier.ID
                                                     select r).ToList();

                    foreach (RebarSet rs in selectedRebars)
                    {
                        List<RebarLegFace> rebarLegFaces = rs.LegFaces;
                        int[] layers = new int[rebarLegFaces.Count];
                        for (int i = 0; i < rebarLegFaces.Count; i++)
                        {
                            layers[i] = rebarLegFaces[i].LayerOrderNumber;
                        }
                        LayerDictionary.Add(rs.Identifier.ID, layers);
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

                    switch (profileType)
                    {
                        case ProfileType.FTG:
                            FTG ftg = new FTG(beam);
                            ftg.CreateSingle(rebarName);
                            break;
                        case ProfileType.RTW:
                            RTW rtw = new RTW(beam);
                            rtw.CreateSingle(rebarName);
                            break;
                        case ProfileType.CLMN:
                            CLMN clmn = new CLMN(beam);
                            clmn.CreateSingle(rebarName);
                            break;
                    }

                    //Restore user work plane
                    model.GetWorkPlaneHandler().SetCurrentTransformationPlane(currentPlane);
                    model.CommitChanges();

                    ChangeLayer(model);
                    ChangeLayer(model);
                    LayerDictionary = new Dictionary<int, int[]>();
                }
            }
            catch (System.ApplicationException)
            {
                Operation.DisplayPrompt("User interrupted!");
            }

        }
        static void ChangeLayer(Model model)
        {
            Type[] Types = new Type[] { typeof(RebarSet) };
            ModelObjectEnumerator Enum = model.GetModelObjectSelector().GetAllObjectsWithType(Types);
            var rebarList = Utility.ToList(Enum);

            foreach (KeyValuePair<int, int[]> keyValuePair in LayerDictionary)
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

        public static Dictionary<int, int[]> LayerDictionary = new Dictionary<int, int[]>();
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
        public static string FatherIDName = "ZTB_FatherIDName";
        public static string MethodName = "ZTB_MethodName";
        public static string MethodInput = "ZTB_MethodInput";
        public static int MinLengthCoefficient = 20;
    }
}
