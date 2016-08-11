using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MultipleExchangeRates
{
    public class RetrieveRate : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            if (context.InputParameters["TransactionCurrencyId"] != null && context.InputParameters["TransactionCurrencyId"] is Guid)
            {
                Guid currencyId = (Guid)context.InputParameters["TransactionCurrencyId"];
                Decimal exchangeRate = GetCurrentExchangeRate(service, currencyId);

                if (exchangeRate != 0)
                {
                    context.OutputParameters["ExchangeRate"] = exchangeRate;
                }
            }
        }

        private decimal GetCurrentExchangeRate(IOrganizationService service, Guid currencyId)
        {
            QueryExpression query = new QueryExpression("fjo_exchangerate");
            query.ColumnSet = new ColumnSet("fjo_rate", "fjo_currencyid");
            query.Criteria.AddCondition("fjo_currencyid", ConditionOperator.Equal, currencyId);
            query.Criteria.AddCondition("fjo_datefrom", ConditionOperator.LessEqual, DateTime.Today);
            query.Criteria.AddCondition("fjo_dateto", ConditionOperator.GreaterEqual, DateTime.Today);

            EntityCollection exchangeRates = service.RetrieveMultiple(query);

            return (exchangeRates.Entities.Count == 1 ?
                exchangeRates[0].GetAttributeValue<decimal>("fjo_rate") :
                0);
        }
    }
}
