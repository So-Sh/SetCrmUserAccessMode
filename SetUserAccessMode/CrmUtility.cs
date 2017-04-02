using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Microsoft.Xrm.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Client.Services;
using Microsoft.Xrm.Sdk.Query;
using System.Globalization;
using System.Configuration;

namespace SetUserAccessMode
{
    public static class CrmUtility
    {
        public static string CRM_USERNAME
        {
            get
            {
                return ConfigurationManager.AppSettings["CRM_USERDOMAIN"] + @"\" + ConfigurationManager.AppSettings["CRM_USERNAME"];
            }
        }
        public static string CRM_PASSWORD
        {
            get
            {
                return ConfigurationManager.AppSettings["CRM_PASSWORD"]; ;
            }
        }
        public static string CRM_URL
        {
            get
            {
                return ConfigurationManager.AppSettings["CRM_URL"]; ;
            }
        }


        internal static CrmConnection GetCRMConnection(string crmUrl, string username, string password)
        {
            try
            {
                string CRMServiceURL = crmUrl;
                string Username = username;
                string Password = password;


                CrmConnection connection = CrmConnection.Parse("ServiceUri=" + CRMServiceURL +
                  "; Username=" + Username + "; Password=" + Password);


                return connection;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }

        }

        internal static string GetCrmAttributeValue(Entity entity, string attribute)
        {

            if (entity.Attributes.Contains(attribute) && entity.Attributes[attribute] != null)
            {
                return (string)entity.Attributes[attribute];
            }
            else
            {
                return "N/A";
            }
        }

        internal static EntityReference GetCrmUserReference(CrmConnection connection, string username)
        {

            EntityReference entRefUser;
            using (OrganizationService _orgService = new OrganizationService(connection))
            {
                try
                {

                    //Create a Query Expression to get the user 
                    FilterExpression filter = new FilterExpression();
                    filter.AddCondition("domainname", ConditionOperator.EndsWith, username);
                    QueryExpression queryUser = new QueryExpression
                    {
                        EntityName = "systemuser",
                        ColumnSet = new ColumnSet(true),
                        Criteria = filter,
                    };

                    EntityCollection entColUser = _orgService.RetrieveMultiple(queryUser);
                    // Return null if no Substation is available
                    if (entColUser.Entities.Count == 0)
                    {
                        return null;
                    }

                    entRefUser = entColUser[0].ToEntityReference();
                    return entRefUser;

                }
                catch (Exception ex)
                {
                    var str = ex.Message;
                    return null;
                }
            }
        }

  
        internal static EntityReference GetUserBusinessUnit(CrmConnection connection, string buName)
        {

            EntityReference entRefBU;
            using (OrganizationService _orgService = new OrganizationService(connection))
            {
                try
                {

                    //Create a Query Expression to get the user 
                    FilterExpression filter = new FilterExpression();
                    filter.AddCondition("name", ConditionOperator.Equal, buName);
                    QueryExpression queryUser = new QueryExpression
                    {
                        EntityName = "businessunit",
                        ColumnSet = new ColumnSet(true),
                        Criteria = filter,
                    };

                    EntityCollection entColUser = _orgService.RetrieveMultiple(queryUser);
                    // Return null if no Substation is available
                    if (entColUser.Entities.Count == 0)
                    {
                        return null;
                    }

                    entRefBU = entColUser[0].ToEntityReference();
                    return entRefBU;

                }
                catch (Exception ex)
                {
                    var str = ex.Message;
                    return null;
                }
            }
        }

        public static string ParseSpecialStringAsDecimal(string value)
        {
            if (value.Contains("#") || value.Equals("#N/A") || value.Equals("#REF!") || value.Length == 0)
            {
                return "N/A";
            }
            else
            {
                if (value.Contains("E"))
                {
                    decimal decTemp = Decimal.Parse(value.Replace(" ", ""), NumberStyles.Any | NumberStyles.AllowCurrencySymbol
                        | NumberStyles.AllowThousands | NumberStyles.AllowDecimalPoint);

                    if (Convert.ToString(decTemp.ToString("#.#####")).StartsWith("."))
                    {
                        return "0" + Convert.ToString(decTemp.ToString("#.#####"));
                    }
                    else
                    {
                        return Convert.ToString(decTemp.ToString("#.#####"));
                    }
                }
                else
                {
                    double objD = Convert.ToDouble(value);
                    return objD.ToString();

                }
            }

        }

    }
}