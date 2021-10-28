using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

using Tekla.Structures;
using Tekla.Structures.Model;
using Tekla.Structures.Geometry3d;
using Tekla.Structures.Model.Operations;


namespace ZeroTouchTekla
{
    class RebarCreator
    {
        public static void Test()
        {
            Model model = new Model();
            Tekla.Structures.Model.UI.Picker picker = new Tekla.Structures.Model.UI.Picker();
            Tekla.Structures.Model.UI.Picker.PickObjectEnum pickObjectEnum = Tekla.Structures.Model.UI.Picker.PickObjectEnum.PICK_ONE_REINFORCEMENT;
            RebarSet rebarSet = picker.PickObject(pickObjectEnum) as RebarSet;
            int mlsi = 90;
            double mlsd = 90;
            string mlss = "";
            rebarSet.GetUserProperty(RebarCreator.FatherIDName, ref mlsi);
            rebarSet.GetUserProperty(RebarCreator.FatherIDName, ref mlsd);
            rebarSet.GetUserProperty(RebarCreator.FatherIDName, ref mlss);

            rebarSet.SetUserProperty(RebarCreator.FatherIDName, "asd");
        }
        public static void Test2()
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

                    RTW rtw = new RTW(part);

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
        public static void Create(ProfileType profileType)
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
                            Tekla.Structures.Model.UI.Picker secondPicker = new Tekla.Structures.Model.UI.Picker();
                            Tekla.Structures.Model.UI.Picker.PickObjectEnum secondPickObjectEnum = Tekla.Structures.Model.UI.Picker.PickObjectEnum.PICK_ONE_PART;
                            Beam secondPart = picker.PickObject(secondPickObjectEnum) as Beam;
                            DRTW drtw = new DRTW(part,secondPart);
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
                    }

                    //Restore user work plane
                    model.GetWorkPlaneHandler().SetCurrentTransformationPlane(currentPlane);
                    model.CommitChanges();
                }

                ChangeLayer(model);
                ChangeLayer(model);
                LayerDictionary = new Dictionary<int, int[]>();
            }
            catch(System.ApplicationException)
            {
                Operation.DisplayPrompt("User interrupted!");
            }
        }
        public static void RecreateRebar()
        {
            Model model = new Model();
            Tekla.Structures.Model.UI.Picker picker = new Tekla.Structures.Model.UI.Picker();
            Tekla.Structures.Model.UI.Picker.PickObjectEnum pickObjectEnum = Tekla.Structures.Model.UI.Picker.PickObjectEnum.PICK_ONE_REINFORCEMENT;

            try
            {
                RebarSet rebarSet = picker.PickObject(pickObjectEnum) as RebarSet;
                rebarSet.GetUserProperty(RebarCreator.FatherIDName, ref FatherID);
                string rebarName = rebarSet.RebarProperties.Name;

                var singleRebars = Utility.ToList(rebarSet.GetReinforcements());
                Type[] fatherTypes = new Type[] { typeof(Beam) };
                ModelObjectEnumerator fatherEnumerator = model.GetModelObjectSelector().GetAllObjectsWithType(fatherTypes);
                var fatherList = Utility.ToList(fatherEnumerator);
                Beam beam = (from Beam b in fatherList
                             where b.Identifier.ID == FatherID
                             select b).FirstOrDefault();

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
            catch(System.ApplicationException)
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
            CLMN
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
                    if(profileString.Contains("CLMN"))
                    {
                        return ProfileType.CLMN;
                    }
                }
            }
            return ProfileType.None;
        }
        public static string FatherIDName = "USER_FIELD_1";
    }
}
