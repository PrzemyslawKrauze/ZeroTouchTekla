using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

using Tekla.Structures;
using Tekla.Structures.Model;
using Tekla.Structures.Geometry3d;


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
            rebarSet.GetUserProperty("__MIN_BAR_LENTYPE", ref mlsi);
            rebarSet.GetUserProperty("__MIN_BAR_LENTYPE", ref mlsd);
            rebarSet.GetUserProperty("__MIN_BAR_LENTYPE", ref mlss);

            int a = 1;
        }
        public static void Create(ProfileType profileType)
        {
            Model model = new Model();
            Tekla.Structures.Model.UI.Picker picker = new Tekla.Structures.Model.UI.Picker();
            Tekla.Structures.Model.UI.Picker.PickObjectEnum pickObjectEnum = Tekla.Structures.Model.UI.Picker.PickObjectEnum.PICK_ONE_PART;
            Beam part = picker.PickObject(pickObjectEnum) as Beam;
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
                }

                //Restore user work plane
                model.GetWorkPlaneHandler().SetCurrentTransformationPlane(currentPlane);
                model.CommitChanges();
            }

            ChangeLayer(model);
        }
        public static void RecreateRebar()
        {
            Model model = new Model();
            Tekla.Structures.Model.UI.Picker picker = new Tekla.Structures.Model.UI.Picker();
            Tekla.Structures.Model.UI.Picker.PickObjectEnum pickObjectEnum = Tekla.Structures.Model.UI.Picker.PickObjectEnum.PICK_ONE_REINFORCEMENT;
            RebarSet rebarSet = picker.PickObject(pickObjectEnum) as RebarSet;

            string rebarName = rebarSet.RebarProperties.Name;

            var singleRebars = Utility.ToList(rebarSet.GetReinforcements());
            ModelObject father = (singleRebars.FirstOrDefault() as SingleRebar).Father;
            Beam beam = father as Beam;
            string hostName = beam.Profile.ProfileString;
            ProfileType profileType = GetProfileType(hostName);
            if (profileType != ProfileType.None)
            {

                Type[] Types = new Type[] { typeof(RebarSet) };
                ModelObjectEnumerator moe = model.GetModelObjectSelector().GetAllObjectsWithType(Types);
                var rebarList = Utility.ToList(moe);

                List<RebarSet> allHostRebar = (from RebarSet r in rebarList
                                               where (Utility.ToList(r.GetReinforcements()).FirstOrDefault() as SingleRebar).Father.Identifier.ID == father.Identifier.ID
                                               select r).ToList();

                foreach (RebarSet rs in allHostRebar)
                {
                    List<RebarLegFace> rebarLegFaces = rs.LegFaces;
                    int[] layers = new int[rebarLegFaces.Count];
                    for (int i = 0; i < rebarLegFaces.Count; i++)
                    {
                        layers[i] = rebarLegFaces[i].LayerOrderNumber;
                    }
                    LayerDictionary.Add(rs.Identifier.ID, layers);
                }

                List<RebarSet> rebarsToRecreate = (from RebarSet r in allHostRebar
                                                   where r.RebarProperties.Name == rebarName
                                                   select r).ToList();

                foreach (RebarSet rs in rebarsToRecreate)
                {
                    rs.Delete();
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
                }

                //Restore user work plane
                model.GetWorkPlaneHandler().SetCurrentTransformationPlane(currentPlane);
                model.CommitChanges();

                ChangeLayer(model);
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
                    int[] layerNumbers = keyValuePair.Value;
                    for (int i = 0; i < rs.LegFaces.Count; i++)
                    {
                        rs.LegFaces[i].LayerOrderNumber = layerNumbers[i];
                    }
                    rs.Modify();
                    model.CommitChanges();
                }
            }
            LayerDictionary = new Dictionary<int, int[]>();
        }

        public static Dictionary<int, int[]> LayerDictionary = new Dictionary<int, int[]>();
        public enum ProfileType
        {
            None,
            FTG,
            RTW
        }
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
            }
            return ProfileType.None;
        }
    }
}
