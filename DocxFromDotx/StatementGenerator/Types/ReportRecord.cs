namespace DocxFromDotx.StatementGenerator.Types
{
	public class ReportRecord
		: IReportItem
	{
		public ReportRecord(string name, string value)
		{
			Name = name;
			Value = value;
		}

		public string Value { get; set; }

		#region Implementation of IReportItem

		public string Name { get; set; }

		#endregion
	}
}
