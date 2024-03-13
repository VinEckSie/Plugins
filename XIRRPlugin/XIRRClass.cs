using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.UI.WebControls;
using Microsoft.Crm.Sdk;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using static System.IdentityModel.Tokens.SecurityTokenHandlerCollectionManager;

namespace XIRRPlugin
{
    public class XIRRClass : IPlugin
    {
        private object conditionoperator;
        private object loanid;

        void IPlugin.Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext) serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService organizationService = serviceFactory.CreateOrganizationService(context.UserId);

            ITracingService tracingService = (ITracingService) serviceProvider.GetService(typeof(ITracingService));


            if (context != null)
            {
                try
                {
                    if (context.PreEntityImages.Contains("PreImageXIRR") && context.PreEntityImages["PreImageXIRR"] is Entity)
                    {
                        Entity updatedTransaction = (Entity)context.PreEntityImages["PreImageXIRR"];

                        // Check if the updated record has the necessary attributes
                        if (updatedTransaction.Attributes.Contains("cr471_loanid"))
                        {
                            // Retrieve the loanId from the updated record
                            EntityReference loanRef = (EntityReference)updatedTransaction.Attributes["cr471_loanid"];

                            // retrieve all cr471_transactions records with the same loanid
                            QueryExpression query = new QueryExpression("cr471_transaction");
                            query.ColumnSet = new ColumnSet("cr471_cashflow","cr471_date"); // add other required fields
                            query.Criteria.AddCondition("cr471_loanid", ConditionOperator.Equal, loanRef.Id);

                            EntityCollection transactions = organizationService.RetrieveMultiple(query);

                            // Create lists to store date and cashflow values
                            List<DateTime> dates = new List<DateTime>();
                            List<decimal> cashflows = new List<decimal>();

                            // Loop through the records
                            foreach (Entity transaction in transactions.Entities)
                            {
                                // Access the date and cashflow fields from each transaction
                                DateTime date = (DateTime)transaction["cr471_date"];
                                decimal cashflow = transaction.GetAttributeValue<decimal>("cr471_cashflow");

                                // Add values to the lists
                                dates.Add(date);
                                cashflows.Add(cashflow);
                            }

                            //double[] cashFlows = { -1000, 200, 300, 400, 500 };
                            //double[] dates = { 0, 30, 60, 90, 120 }; // Days from the start of the investment

                            decimal xirr = (decimal)CalculateXIRR(cashflows, dates);

                            // Do something with the xirrResult, such as updating a field on the loan record
                            //...
                        }
                    }
                }
                catch (Exception)
                {

                    throw new InvalidPluginExecutionException("Error during the XIRR plugin execution");
                }


            }
        }

        public static double CalculateXIRR(List<decimal> cashFlows, List<DateTime> dates, double guess = 0.1, double tolerance = 0.001)
        {
            int maxIterations = 100;
            double x0 = guess;
            double x1;

            for (int i = 0; i < maxIterations; i++)
            {
                double fValue = CalculateXIRREquation(cashFlows, dates, x0);
                double fPrimeValue = CalculateXIRRDerivative(cashFlows, dates, x0);

                x1 = x0 - fValue / fPrimeValue;

                if (Math.Abs(x1 - x0) < tolerance)
                {
                    return x1;
                }

                x0 = x1;
            }

            throw new InvalidOperationException("XIRR calculation did not converge to a solution.");
        }

        private static double CalculateXIRREquation(List<decimal> cashFlows, List<DateTime> dates, double rate)
        {
            double result = 0;

            for (int i = 0; i < cashFlows.Count; i++)
            {
                result += (double)cashFlows[i] / Math.Pow(1 + rate, (dates[i] - dates[0]).TotalDays / 365.25);
            }

            return result;
        }

        private static double CalculateXIRRDerivative(List<decimal> cashFlows, List<DateTime> dates, double rate)
        {
            double result = 0;

            for (int i = 0; i < cashFlows.Count; i++)
            {
                result -= (double)cashFlows[i] * (dates[i] - dates[0]).TotalDays / (365.25 * Math.Pow(1 + rate, 2 * (dates[i] - dates[0]).TotalDays / 365.25));
            }

            return result;
        }
    }
}
