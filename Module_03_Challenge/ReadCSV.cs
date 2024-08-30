using Autodesk.Revit.DB.Architecture;
using System.Drawing.Text;

namespace Module_03_Challenge
{
    [Transaction(TransactionMode.Manual)]
    public class ReadCSV : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            // Read Furniture Types
            string furnTypesPath = "C:\\Users\\rbachina\\Downloads\\RAB_Module_03_Challenge_Files\\RAB_Module 03_Furniture Types.csv";

            List<string[]> furnTypesData = new List<string[]>();

            string[] furnTypesArray = System.IO.File.ReadAllLines(furnTypesPath);

            foreach (string furnTypesString in furnTypesArray)
            {
                string[] rowArray = furnTypesString.Split(',');
                furnTypesData.Add(rowArray);
            }

            furnTypesData.RemoveAt(0);

            // Read Furniture Sets
            string furnSetsPath = "C:\\Users\\rbachina\\Downloads\\RAB_Module_03_Challenge_Files\\RAB_Module 03_Furniture Sets.csv";

            List<string[]> furnSetsData = ParseCsv(furnSetsPath);
            furnSetsData.RemoveAt(0);

            // Add data into Classes
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

            // Filter out Rooms
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
                        if(furnSet.FurnitureSet == furnSetName)
                        {
                            foreach(string furnItem in furnSet.IncludedFurniture)
                            {
                                FamilySymbol furnSymbol = GetFurnitureByName(doc, furnitureTypesList, furnItem);

                                if(furnSymbol != null)
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

                private List<string[]> ParseCsv(string path)
        {
            List<string[]> data = new List<string[]>();

            using (var reader = new StreamReader(path))
            {
                while (!reader.EndOfStream)
                {
                    string line = reader.ReadLine();
                    string[] values = ParseCsvLine(line);
                    data.Add(values);
                }
            }

            return data;
        }

        private string[] ParseCsvLine(string line)
        {
            List<string> values = new List<string>();
            bool inQuotes = false;
            string currentValue = string.Empty;

            foreach (char c in line)
            {
                if (c == '\"')
                {
                    inQuotes = !inQuotes; // Toggle quote state
                }
                else if (c == ',' && !inQuotes)
                {
                    values.Add(currentValue);
                    currentValue = string.Empty;
                }
                else
                {
                    currentValue += c;
                }
            }

            values.Add(currentValue); // Add the last value
            return values.ToArray();
        }

        private FamilySymbol GetFurnitureByName(Document doc, List<FurnitureTypes> furnitureTypesList, string tmpfurnItem)
        {
            string furnItem = tmpfurnItem.Trim();
            foreach(FurnitureTypes furnType in furnitureTypesList)
            {
                if (furnType.FurnitureName == furnItem)
                {
                    FamilySymbol furnFS = Utils.Common.GetFamilySymbolByName(doc, furnType.RevitFamilyName, furnType.RevitFamilyType);

                    if(furnFS != null)
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
