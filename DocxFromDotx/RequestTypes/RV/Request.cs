using System.Collections.Generic;

namespace Gems.Smev.Rpgu.Service.RecieveStatement.Facade.RV
{
    public interface IRequest
    {
		createSmvGusPodgRazreshVvodExplV2Request Request { get; }

		string RawRequestXml { get; }

	}
}
