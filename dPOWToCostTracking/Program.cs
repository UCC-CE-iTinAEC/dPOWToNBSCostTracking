using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Xbim.DPoW;

namespace dPOWToCostTracking
{
	class Program
	{
		private static int _assetCount = 1;
		public static int AssetCount()
		{
			return _assetCount++;
		}

		static void Main(string[] args)
		{
			// import plan of work, isolate assetTypes
			var planofwork = Xbim.DPoW.PlanOfWork.OpenJson("..\\..\\ExampledPOWFiles\\007-NBS_Lakeside_Restaurant.dpow");
            // example of exporting stage 4 (design) deliverables to NBS cost tracking
			var assetTypes = planofwork.ProjectStages[4].AssetTypes;

			// generate xml tree with appropriate data from plan of work
			XDocument output = new XDocument(new XDeclaration("1.0", "Windows-1252", null),
			new XElement("CostTrackingSpreadSheet",
				new XElement("Schedule",
					new XAttribute("title", planofwork.Project.Name),
					new XAttribute("currency", GetCurrentCode(planofwork.Project.CurrencyUnits)),
					new XAttribute("provSumsCosts", "0.00"),
					new XAttribute("retentionPer", "0"),
					new XAttribute("tenderSum", "0.00"),
					new XAttribute("contractSum", "0.00"),
					new XAttribute("scheduleOfWorkLocation", ""),
					new XAttribute("scheduleOfWorkTitle", planofwork.Project.Name),
					new XAttribute("caProjectLocation", ""),
					new XAttribute("caProjectName", ""),
					new XAttribute("isSpecifierSpreadSheet", "True"),
					new XElement("Constructions",
						from asset in assetTypes
						select new XElement("Construction",
							new XAttribute("key", AssetCount()),
							new XAttribute("visKey", GetUniclassCode(asset, planofwork)),
							new XAttribute("title", asset.Name),
							new XAttribute("contractCost", "0.00"),
							new XAttribute("headingtitle", ""),
							new XElement("ProgressDetails"))),
					new XElement("Prelims",
						new XAttribute("key", "00"),
						new XAttribute("title", "Prelims (excl prov sums)"),
						new XAttribute("contractCost", "0.00"),
						new XAttribute("headingTitle", ""),
						new XElement("ProgressDetails")),
					new XElement("ProvSums",
						new XAttribute("key", "PS"),
						new XAttribute("title", "Provisional Sums"),
						new XAttribute("contractCost", "0.00"),
						new XAttribute("headingTitle", ""),
						new XElement("ProgressDetails")),
					new XElement("Variations"),
					new XElement("AdditionalPayments"),
					new XElement("ProgressPeriods"))));

			output.Save("..\\..\\Output\\" + planofwork.Project.Name + ".sdlx", SaveOptions.DisableFormatting);
		}

		private static string GetUniclassCode(DPoWObject asset, PlanOfWork pow)
		{
			var uniclass = pow.ClassificationSystems.FirstOrDefault(c => c.Name == "Uniclass2015");
			var classification = asset.GetClassificationReferences(pow);

			foreach(var ci in classification)
			{
				var item = uniclass.ClassificationReferences.FirstOrDefault(c => c.Id == ci.Id);
				if (item != null) return item.ClassificationCode;
			}
			return string.Empty;
		}

        private static string GetCurrentCode(CurrencyUnits units)
        {
            switch (units)
            {
                case CurrencyUnits.GBP:
                    // Bug in NBS Cost Tracking - GBP is GPB
                    return "GPB";

                default:
                    return units.ToString();
            }
        }
	}
}