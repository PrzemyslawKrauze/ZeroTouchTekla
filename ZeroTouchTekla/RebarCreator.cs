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
            Tekla.Structures.Model.UI.Picker.PickObjectEnum pickObjectEnum = Tekla.Structures.Model.UI.Picker.PickObjectEnum.PICK_ONE_OBJECT;
            ModelObject modelObject = picker.PickObject(pickObjectEnum);

        }
        public static void CreateForPart(ProfileType profileType)
        {
            Model model = new Model();
            Tekla.Structures.Model.UI.Picker picker = new Tekla.Structures.Model.UI.Picker();
            Tekla.Structures.Model.UI.Picker.PickObjectEnum pickObjectEnum = Tekla.Structures.Model.UI.Picker.PickObjectEnum.PICK_ONE_PART;
            try
            {
                Beam part = picker.PickObject(pickObjectEnum) as Beam;
                FatherID = part.Identifier.ID;
                if (part != null)
                {
                    //Store current work plane
                    TransformationPlane currentPlane = model.GetWorkPlaneHandler().GetCurrentTransformationPlane();
                    //Get beam local plane
                    TransformationPlane localPlane = new TransformationPlane(part.GetCoordinateSystem());
                    model.GetWorkPlaneHandler().SetCurrentTransformationPlane(localPlane);

                    Tekla.Structures.Model.UI.Picker secondPicker;
                    Tekla.Structures.Model.UI.Picker.PickObjectEnum secondPickObjectEnum;
                    Beam secondPart;
                    Tekla.Structures.Model.UI.Picker thirdPicker;
                    Tekla.Structures.Model.UI.Picker.PickObjectEnum thirdPickObjectEnum;
                    Beam thirdPart;
                    switch (profileType)
                    {
                        case ProfileType.FTG:
                            FTG ftg = new FTG(part);
                            ftg.Create();
                            break;
                        case ProfileType.RTW:
                            RTW rtw = new RTW(part);
                            rtw.Create();
                            break;
                        case ProfileType.DRTW:
                            secondPicker = new Tekla.Structures.Model.UI.Picker();
                            secondPickObjectEnum = Tekla.Structures.Model.UI.Picker.PickObjectEnum.PICK_ONE_PART;
                            secondPart = picker.PickObject(secondPickObjectEnum) as Beam;
                            DRTW drtw = new DRTW(part, secondPart);
                            drtw.Create();
                            break;
                        case ProfileType.RTWS:
                            RTWS rtws = new RTWS(part);
                            rtws.Create();
                            break;
                        case ProfileType.CLMN:
                            CLMN clmn = new CLMN(part);
                            clmn.Create();
                            break;
                        case ProfileType.ABT:
                            // ABT abt = new ABT(part);
                            // abt.Create();
                            DABT dabt1 = new DABT(part);
                            dabt1.Create();
                            break;
                        case ProfileType.DABT:
                            secondPicker = new Tekla.Structures.Model.UI.Picker();
                            secondPickObjectEnum = Tekla.Structures.Model.UI.Picker.PickObjectEnum.PICK_ONE_PART;
                            secondPart = picker.PickObject(secondPickObjectEnum) as Beam;
                            DABT dabt = new DABT(part, secondPart);
                            dabt.Create();
                            break;
                        case ProfileType.TABT:
                            secondPicker = new Tekla.Structures.Model.UI.Picker();
                            secondPickObjectEnum = Tekla.Structures.Model.UI.Picker.PickObjectEnum.PICK_ONE_PART;
                            secondPart = picker.PickObject(secondPickObjectEnum) as Beam;
                            thirdPicker = new Tekla.Structures.Model.UI.Picker();
                            thirdPickObjectEnum = Tekla.Structures.Model.UI.Picker.PickObjectEnum.PICK_ONE_PART;
                            thirdPart = picker.PickObject(thirdPickObjectEnum) as Beam;
                            DABT dabt3 = new DABT(part, secondPart, thirdPart);
                            dabt3.Create();
                            break;
                        case ProfileType.APS:
                            APS aps = new APS(part);
                            aps.Create();
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
            DABT,
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
