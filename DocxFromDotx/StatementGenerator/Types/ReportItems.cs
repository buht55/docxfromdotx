using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace DocxFromDotx.StatementGenerator.Types
{
	public class ReportItems
		: ICollection<IReportItem>
	{
		private readonly List<IReportItem> _list;

		public ReportItems()
		{
			_list = new List<IReportItem>();
		}

		public ReportItems(params IReportItem[] items)
			: this()
		{
			_list.AddRange(items);
		}
		public IReportItem this[string name]
		{
			get
			{
				return _list.FirstOrDefault(x => string.Equals(x.Name, name, StringComparison.OrdinalIgnoreCase));
			}
		}

		public ReportRecord AddReportRecord(string name, string value)
		{
			var item = new ReportRecord(name, value);
			Add(item);
			return item;
		}

		public ReportBlock AddReportBlock(string name, bool showFromNewPage = false)
		{
			var item = new ReportBlock(name, showFromNewPage);
			Add(item);
			return item;
		}

		public ReportBlock AddReportBlock(ReportBlock block)
		{
			Add(block);
			return block;
		}

		public ReportRecord GetReportRecord(string name)
		{
			var item = this[name];
			return item as ReportRecord;
		}

		public ReportBlock GetReportBlock(string name)
		{
			var item = this[name];
			return item as ReportBlock;
		}

		#region Implementation of IEnumerable

		public IEnumerator<IReportItem> GetEnumerator()
		{
			return _list.GetEnumerator();
		}

		IEnumerator IEnumerable.GetEnumerator()
		{
			return GetEnumerator();
		}

		#endregion

		#region Implementation of ICollection<IReportItem>

		public void Add(IReportItem item)
		{
			if (item == null)
				throw new ArgumentNullException(nameof(item));
			if (string.IsNullOrEmpty(item.Name))
				throw new ArgumentNullException(nameof(item), "Не в объекте IReportItem заполнено имя RusName");
			if (_list.Any(x => string.Equals(x.Name, item.Name, StringComparison.OrdinalIgnoreCase)))
				throw new ArgumentException("Поле RusName объекта IReportItem должно быть уникальным");
			_list.Add(item);
		}

		public void Clear()
		{
			_list.Clear();
		}

		public bool Contains(IReportItem item)
		{
			return _list.Contains(item);
		}

		public void CopyTo(IReportItem[] array, int arrayIndex)
		{
			_list.CopyTo(array, arrayIndex);
		}

		public bool Remove(IReportItem item)
		{
			return _list.Remove(item);
		}

		public int Count => _list.Count;
		public bool IsReadOnly => false;

		#endregion
	}
}