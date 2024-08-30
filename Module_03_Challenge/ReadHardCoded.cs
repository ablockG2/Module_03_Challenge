using Autodesk.Revit.DB.Architecture;
using System.Drawing.Text;

namespace Module_03_Challenge
{
    [Transaction(TransactionMode.Manual)]
    public class ReadHardCoded : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            //this is a variable for the current Revit model

            Document doc = uiapp.ActiveUIDocument.Document;

            //string furnSetsPath = "C:\\Users\\rbachina\\Downloads\\RAB_Module_03_Challenge_Files\\RAB_Module 03_Furniture Sets.csv";

            //List<string[]> furnSetsData = new List<string[]>();

            //string[] furnSetsArray = System.IO.File.ReadAllLines(furnSetsPath);

            //foreach (string furnSetsString in furnSetsArray)
            //{
            //    string[] rowArray = furnSetsString.Split(',');
            //    furnSetsData.Add(rowArray);
            //}

            //furnSetsData.RemoveAt(0);

            //string furnTypesPath = "C:\\Users\\rbachina\\Downloads\\RAB_Module_03_Challenge_Files\\RAB_Module 03_Furniture Types.csv";

            //List<string[]> furnTypesData = new List<string[]>();

            //string[] furnTypesArray = System.IO.File.ReadAllLines(furnTypesPath);

            //foreach (string furnTypesString in furnTypesArray)
            //{
            //    string[] rowArray = furnTypesString.Split(',');
            //    furnTypesData.Add(rowArray);
            //}

            //furnTypesData.RemoveAt(0);

            List<string[]> furnTypesData = GetFurnitureTypes();
            List<string[]> furnSetsData = GetFurnitureSets();
            furnTypesData.RemoveAt(0);
            furnSetsData.RemoveAt(0);


            List<FurnitureTypes> furnitureTypesList = new List<FurnitureTypes>();
            foreach (string[] row in furnTypesData)
            {
                FurnitureTypes furnTypes = new FurnitureTypes(row[0], row[1], row[2]);
                furnitureTypesList.Add(furnTypes);
            }

            List<FurnitureSets> furnitureSetsList = new List<FurnitureSets>();
            foreach (string[] row in furnSetsData)
            {
                FurnitureSets furnSets = new FurnitureSets(row[0], row[1], row[2]);
                furnitureSetsList.Add(furnSets);
            }

            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(BuiltInCategory.OST_Rooms);

            //FamilySymbol curFS = Utils.Common.GetFamilySymbolByName(doc, furnitureTypesList[1].RevitFamilyName, furnitureTypesList[1].RevitFamilyType);

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Insert family into room");
                int counter = 0;

                foreach (SpatialElement room in collector)
                {
                    LocationPoint loc = room.Location as LocationPoint;
                    XYZ roomPoint = loc.Point as XYZ;


                    //FamilyInstance curFI = doc.Create.NewFamilyInstance(roomPoint, curFS, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                    // Get parameter value
                    string furnSetName = Utils.Common.GetParameterValueAsString(room, "Furniture Set");

                    foreach (FurnitureSets furnSet in furnitureSetsList)
                    {
                        if (furnSet.FurnitureSet == furnSetName)
                        {
                            foreach (string furnItem in furnSet.IncludedFurniture)
                            {
                                FamilySymbol furnSymbol = GetFurnitureByName(doc, furnitureTypesList, furnItem);

                                if (furnSymbol != null)
                                {
                                    FamilyInstance furnFI = doc.Create.NewFamilyInstance(roomPoint, furnSymbol, room,
                                        Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                                    counter++;
                                }
                            }
                        }

                        Utils.Common.SetParameterValue(room, "Furniture Set", furnSet.GetFurnitureCount());
                    }
                }
                t.Commit();

                TaskDialog.Show("Result", "Furniture added to " + counter + " rooms.");

                return Result.Succeeded;
            }
        }

        private FamilySymbol GetFurnitureByName(Document doc, List<FurnitureTypes> furnitureTypesList, string tmpfurnItem)
        {
            string furnItem = tmpfurnItem.Trim();
            foreach (FurnitureTypes furnType in furnitureTypesList)
            {
                if (furnType.FurnitureName == furnItem)
                {
                    FamilySymbol furnFS = Utils.Common.GetFamilySymbolByName(doc, furnType.RevitFamilyName, furnType.RevitFamilyType);

                    if (furnFS != null)
                    {
                        if (furnFS.IsActive == false)
                        {
                            furnFS.Activate();
                        }
                    }

                    return furnFS;
                }
            }

            return null;
        }

        public class FurnitureTypes
        {
            public string FurnitureName { get; set; }
            public string RevitFamilyName { get; set; }
            public string RevitFamilyType { get; set; }

            public FurnitureTypes(string _furniturename, string _revitfamilyname, string _revitfamilytype)
            {
                FurnitureName = _furniturename;
                RevitFamilyName = _revitfamilyname;
                RevitFamilyType = _revitfamilytype;
            }
        }

        public class FurnitureSets
        {
            public string FurnitureSet { get; set; }
            public string RoomType { get; set; }
            public string[] IncludedFurniture { get; set; }

            public FurnitureSets(string _furnitureset, string _roomtype, string _includedfurniture)
            {
                FurnitureSet = _furnitureset;
                RoomType = _roomtype;
                IncludedFurniture = _includedfurniture.Split(',');
            }

            public int GetFurnitureCount()
            {
                return FurnitureSet.Count();
            }
        }

        private List<string[]> GetFurnitureTypes()
        {
            List<string[]> returnList = new List<string[]>();
            returnList.Add(new string[] { "Furniture Name", "Revit Family Name", "Revit Family Type" });
            returnList.Add(new string[] { "desk", "Desk", "60in x 30in" });
            returnList.Add(new string[] { "task chair", "Chair-Task", "Chair-Task" });
            returnList.Add(new string[] { "side chair", "Chair-Breuer", "Chair-Breuer" });
            returnList.Add(new string[] { "bookcase", "Shelving", "96in x 12in x 84in" });
            returnList.Add(new string[] { "loveseat", "Sofa", "54in" });
            returnList.Add(new string[] { "teacher desk", "Table-Rectangular", "48in x 30in" });
            returnList.Add(new string[] { "student desk", "Desk", "60in x 30in Student" });
            returnList.Add(new string[] { "computer desk", "Table-Rectangular", "48in x 30in" });
            returnList.Add(new string[] { "lab desk", "Table-Rectangular", "72in x 30in" });
            returnList.Add(new string[] { "lounge chair", "Chair-Corbu", "Chair-Corbu" });
            returnList.Add(new string[] { "coffee table", "Table-Coffee", "30in x 60in x 18in" });
            returnList.Add(new string[] { "sofa", "Sofa-Corbu", "Sofa-Corbu" });
            returnList.Add(new string[] { "dining table", "Table-Dining", "30in x 84in x 22in" });
            returnList.Add(new string[] { "dining chair", "Chair-Breuer", "Chair-Breuer" });
            returnList.Add(new string[] { "stool", "Chair-Task", "Chair-Task" });

            return returnList;
        }

        private List<string[]> GetFurnitureSets()
        {
            List<string[]> returnList = new List<string[]>();
            returnList.Add(new string[] { "Furniture Set", "Room Type", "Included Furniture" });
            returnList.Add(new string[] { "A", "Office", "desk, task chair, side chair, bookcase" });
            returnList.Add(new string[] { "A2", "Office", "desk, task chair, side chair, bookcase, loveseat" });
            returnList.Add(new string[] { "B", "Classroom - Large", "teacher desk, task chair, student desk, student desk, student desk, student desk, student desk, student desk, student desk, student desk, student desk, student desk, student desk, student desk" });
            returnList.Add(new string[] { "B2", "Classroom - Medium", "teacher desk, task chair, student desk, student desk, student desk, student desk, student desk, student desk, student desk, student desk" });
            returnList.Add(new string[] { "C", "Computer Lab", "computer desk, computer desk, computer desk, computer desk, computer desk, computer desk, task chair, task chair, task chair, task chair, task chair, task chair" });
            returnList.Add(new string[] { "D", "Lab", "teacher desk, task chair, lab desk, lab desk, lab desk, lab desk, lab desk, lab desk, lab desk, stool, stool, stool, stool, stool, stool, stool" });
            returnList.Add(new string[] { "E", "Student Lounge", "lounge chair, lounge chair, lounge chair, sofa, coffee table, bookcase" });
            returnList.Add(new string[] { "F", "Teacher's Lounge", "lounge chair, lounge chair, sofa, coffee table, dining table, dining chair, dining chair, dining chair, dining chair, bookcase" });
            returnList.Add(new string[] { "G", "Waiting Room", "lounge chair, lounge chair, sofa, coffee table" });

            return returnList;
        }

        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand1";
            string buttonTitle = "Button 1";

            Utils.ButtonDataClass myButtonData1 = new Utils.ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "This is a tooltip for Button 1");

            return myButtonData1.Data;
        }
    }

}
