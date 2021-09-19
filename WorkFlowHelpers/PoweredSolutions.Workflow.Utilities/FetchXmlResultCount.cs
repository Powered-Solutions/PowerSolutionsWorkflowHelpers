using System;
using System.Activities;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;


namespace PoweredSolutions.Workflow.Utilities
{
    public class FetchXmlResultCount : CodeActivity
    {
        [RequiredArgument]
        [Input("FetchXml")]
        public InArgument<string> Source { get; set; }

        [Input("FetchXml Parameter 1")]
        public InArgument<string> Parameter1 { get; set; }

        [Input("FetchXml Parameter 2")]
        public InArgument<string> Parameter2 { get; set; }

        [Input("FetchXml Parameter 3")]
        public InArgument<string> Parameter3 { get; set; }

        [Input("FetchXml Parameter 4")]
        public InArgument<string> Parameter4 { get; set; }

        [Input("FetchXml Parameter 5")]
        public InArgument<string> Parameter5 { get; set; }

        [Output("Count")]
        public OutArgument<int> Count { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            //Create the tracing service
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();
            tracingService.Trace("obtained tracing service successfully.\n");

            //Create the context
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            var fetchXmlQuery = Source.Get(executionContext);

            var parameter1 = FormatParameter(Parameter1.Get(executionContext));
            var parameter2 = FormatParameter(Parameter2.Get(executionContext));
            var parameter3 = FormatParameter(Parameter3.Get(executionContext));
            var parameter4 = FormatParameter(Parameter4.Get(executionContext));
            var parameter5 = FormatParameter(Parameter5.Get(executionContext));

            var formattedFetchXmlQuery = String.Format(fetchXmlQuery, parameter1, parameter2, parameter3, parameter4, parameter5);

            var fetchXmlQueryResults = service.RetrieveMultiple(new FetchExpression(formattedFetchXmlQuery));
            Count.Set(executionContext, fetchXmlQueryResults.Entities.Count);
        }

        private static string FormatParameter(string input)
        {
            var result = input;
            Uri uriRecord;
            DateTime dateValue;

            if (Uri.TryCreate(input, UriKind.Absolute, out uriRecord))
            {
                result = GetRecordId(input)?.ToString();
            }
            else if (DateTime.TryParse(input, out dateValue))
            {
                result = dateValue.ToString("yyyy-MM-dd");
            }

            return result ?? string.Empty;
        }

        public static Guid? GetRecordId(string recordUrl)
        {
            var queryString = new UriBuilder(recordUrl).Query;

            var idString = GetQueryStringParameter(queryString, "id");
            if (idString != null)
            {
                return Guid.Parse(idString);
            }

            return null;
        }

        private static string GetQueryStringParameter(string queryString, string parameterName)
        {
            queryString = queryString.TrimStart(new[] { '?' });
            var qsParams = queryString.Split(new[] { '&' });
            foreach (var item in qsParams)
            {
                var itemArr = item.Split(new[] { '=' });
                if (itemArr[0] == parameterName)
                {
                    return itemArr[1];
                }
            }

            return null;
        }
    }
}
