using System.Collections.Generic;
using DocxFromDotx.StatementGenerator.Types;
using Gems.Smev.Rpgu.Service.RecieveStatement.Facade.RV;

namespace DocxFromDotx
{
    public class RvStatementDocxGus : StatementGenerator.StatementGenerator
    {
        private readonly smvGusPodgRazreshVvodExplV2 _request;
        public RvStatementDocxGus(string name, smvGusPodgRazreshVvodExplV2 request) : base(name)
        {
            _request = request;
        }

        protected override void FillMarks(ReportItems marks)
        {
            marks.AddReportRecord("OrganizationName", _request.OrganizationName);
            marks.AddReportRecord("LegalAddress", _request.LegalAddress);
            marks.AddReportRecord("LastName", _request.LastName);
            marks.AddReportRecord("FirstName", _request.FirstName);
            marks.AddReportRecord("SecondName", _request.SecondName);
            marks.AddReportRecord("RegisteredAddress", _request.RegisteredAddress);
            marks.AddReportRecord("PhoneNumber", _request.PhoneNumber);
            marks.AddReportRecord("PhoneNumber1", _request.PhoneNumber1);
            marks.AddReportRecord("Email", _request.EMail);
            marks.AddReportRecord("EMail1", _request.EMail1);
            marks.AddReportRecord("OName", _request.OName);
            marks.AddReportRecord("OBuildingAddress", _request.OBuildingAddress);
            marks.AddReportRecord("BUPDocumentDate", _request.BUPDocumentDateSpecified ? _request.BUPDocumentDate.ToString() : "");
            marks.AddReportRecord("BUPDocumentNumber", _request.BUPDocumentNumber);

            //todo нужно как-то заполнить это блоком. 
            var zakDocBlock = new ReportBlock("zakDocBlock");
            foreach (var zakDoc in _request.ZAKAppliedDocuments)
            {
                zakDocBlock.Rows.Add(new ReportItems(
                    new ReportRecord("ZAKDocumentNumber", zakDoc.ZAKDocumentNumber),
                    new ReportRecord("ZAKNameIssuingAuthority", zakDoc.ZAKNameIssuingAuthority),
                    new ReportRecord("ZAKDocumentDate", zakDoc.ZAKDocumentDateSpecified ? zakDoc.ZAKDocumentDate.ToString() : "")
                    ));
            }

            marks.AddReportRecord("ZAKDocumentNumber", "");
            marks.AddReportRecord("ZAKNameIssuingAuthority", "");
            marks.AddReportRecord("ZAKDocumentDate", "");

            var gettingResultWay = new List<string> { "", "", "", "" };
            var mfc = "";
            switch (int.Parse(_request.GettingResultWay))
            {
                case 1:
                    gettingResultWay[0] = "V";
                    break;
                case 2:
                    gettingResultWay[1] = "V";
                    break;
                case 3:
                    gettingResultWay[2] = "V";
                    break;
                case 4:
                    gettingResultWay[3] = "V";
                    mfc = _request.MFC;
                    break;
            }

            marks.AddReportRecord("GettingResultWay4", gettingResultWay[3]);
            marks.AddReportRecord("GettingResultWay1", gettingResultWay[0]);
            marks.AddReportRecord("GettingResultWay3", gettingResultWay[2]);
            marks.AddReportRecord("MFC", mfc);
            marks.AddReportRecord("GettingResultWay2", gettingResultWay[1]);


        }
    }
}
