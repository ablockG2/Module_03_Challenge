namespace Module_03_Challenge.Utils
{
    internal static class Common
    {
        internal static string GetParameterValueAsString(Element element, string paramName)
        {
            IList<Parameter> paramList = element.GetParameters(paramName);
            Parameter myParam = paramList.First();

            return myParam.AsString();
        }

        internal static double GetParameterValueAsDouble(Element element, string paramName)
        {
            IList<Parameter> paramList = element.GetParameters(paramName);
            Parameter myParam = paramList.First();

            return myParam.AsDouble();
        }

        internal static void SetParameterValue(Element element, string paramName, string value)
        {
            IList<Parameter> paramList = element.GetParameters(paramName);
            Parameter myParam = paramList.First();

            myParam.Set(value);
        }

        internal static void SetParameterValue(Element element, string paramName, double value)
        {
            IList<Parameter> paramList = element.GetParameters(paramName);
            Parameter myParam = paramList.First();

            myParam.Set(value);
        }

        internal static FamilySymbol GetFamilySymbolByName(Document doc, string famName, string fsName)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(FamilySymbol));


            foreach (FamilySymbol fs in collector)
            {
                if (fs.Name == fsName && fs.FamilyName == famName)
                    return fs;
            }

            return null;
        }

        internal static RibbonPanel CreateRibbonPanel(UIControlledApplication app, string tabName, string panelName)
        {
            RibbonPanel currentPanel = GetRibbonPanelByName(app, tabName, panelName);

            if (currentPanel == null)
                currentPanel = app.CreateRibbonPanel(tabName, panelName);

            return currentPanel;
        }

        internal static RibbonPanel? GetRibbonPanelByName(UIControlledApplication app, string tabName, string panelName)
        {
            foreach (RibbonPanel tmpPanel in app.GetRibbonPanels(tabName))
            {
                if (tmpPanel.Name == panelName)
                    return tmpPanel;
            }

            return null;
        }
    }
}
