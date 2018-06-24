using System.Collections.Generic;

namespace DocxFromDotx.StatementGenerator.Types
{
	public class ReportBlock
		: IReportItem
	{
		public List<ReportItems> Rows { get; }
		public bool ShowFromNewPage { get; set; }

		public ReportBlock(string name)
			: this(name, false)
		{
				
		}

		public ReportBlock(string name, bool showFromNewPage)
		{
			Name = name;
			ShowFromNewPage = showFromNewPage;
			Rows = new List<ReportItems>();
		}

		#region Implementation of IReportItem

		public string Name { get; set; }

		#endregion
	}
}