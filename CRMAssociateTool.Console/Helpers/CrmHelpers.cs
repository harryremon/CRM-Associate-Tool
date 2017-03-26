using System.Collections.Generic;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Metadata.Query;
using Microsoft.Xrm.Sdk.Query;

namespace CRMAssociateTool.Console.Helpers
{
    public static class CrmHelpers
    {
        public enum RelationType
        {
            OneToManyRelationships,
            ManyToOneRelationships,
            ManyToManyRelationships
        }

        public static List<RelationshipMetadataBase> GetCustomRelationships(IOrganizationService service,
            string entityName, RelationType[] types)
        {
            var entityProperties = new MetadataPropertiesExpression
            {
                AllProperties = false
            };
            entityProperties.PropertyNames.AddRange(types.Select(type => type.ToString()));

            var entityFilter = new MetadataFilterExpression(LogicalOperator.And);
            entityFilter.Conditions.Add(
                new MetadataConditionExpression("LogicalName", MetadataConditionOperator.Equals, entityName));

            var relationFilter = new MetadataFilterExpression(LogicalOperator.And);
            relationFilter.Conditions
                .Add(new MetadataConditionExpression("IsCustomRelationship", MetadataConditionOperator.Equals, true));


            var entityQueryExpression = new EntityQueryExpression
            {
                Criteria = entityFilter,
                Properties = entityProperties,
                RelationshipQuery = new RelationshipQueryExpression
                {
                    Criteria = relationFilter
                }
            };

            var retrieveMetadataChangesRequest = new RetrieveMetadataChangesRequest
            {
                Query = entityQueryExpression,
                ClientVersionStamp = null
            };

            var response = ((RetrieveMetadataChangesResponse)service.Execute(retrieveMetadataChangesRequest)).EntityMetadata;

            var relationMetadata = new List<RelationshipMetadataBase>();

            if (response != null)
            {
                var metadata = response.FirstOrDefault();

                if (metadata != null)
                {
                    if (metadata.OneToManyRelationships != null)
                    {
                        relationMetadata.AddRange(metadata.OneToManyRelationships);
                    }

                    if (metadata.ManyToOneRelationships != null)
                    {
                        relationMetadata.AddRange(metadata.ManyToOneRelationships);
                    }

                    if (metadata.ManyToManyRelationships != null)
                    {
                        relationMetadata.AddRange(metadata.ManyToManyRelationships);
                    }
                }
            }

            return relationMetadata;
        }
    }
}