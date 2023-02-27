using System;
using Microsoft.Xrm.Sdk;

namespace CreateEmail
{
    public class CreateEmail : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            try
            {
                var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

                if (!context.InputParameters.Contains("Target")) return;

                bool isDelete = context.MessageName == "Delete";

                Entity entity = !isDelete ? (Entity)context.InputParameters["Target"] : context.PreEntityImages["PreImage"];
                EntityReference user = new EntityReference("systemuser", context.InitiatingUserId);

                if (entity.LogicalName != "contact") return;
                if (!isDelete && !entity.Contains("emailaddress1")) return;

                var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                var service = serviceFactory.CreateOrganizationService(context.UserId);

                Entity Email = new Entity("email");

                Entity From = new Entity("activityparty");
                From["partyid"] = new EntityReference("systemuser", user.Id);
                Email.Attributes["from"] = new Entity[] { From };

                if (!isDelete)
                {
                    Email["regardingobjectid"] = new EntityReference(entity.LogicalName, entity.Id);

                    Entity To = new Entity("activityparty");
                    To["partyid"] = new EntityReference("contact", entity.Id);
                    Email.Attributes["to"] = new Entity[] { To };
                }

                var preImage = context.PreEntityImages.Contains("PreImage") ? context.PreEntityImages["PreImage"] : new Entity();

                Email.Attributes["subject"] = GetSubject(context.MessageName, entity);
                Email.Attributes["description"] = GetDescription(context.MessageName, entity, preImage);
                 
                service.Create(Email);
            }
            catch (Exception ex) {
                throw new InvalidPluginExecutionException(ex.Message + "\n" + ex.StackTrace);
            }
        }

        private string GetDescription(string messageType, Entity contact, Entity preImage)
        {
            var message = "";
            switch (messageType)
            {
                case "Create":
                    message = "New contact created";
                    break;
                case "Update":
                    message = "Old email address - " + preImage.GetAttributeValue<string>("emailaddress1") + "\n" + "New email address " + contact.GetAttributeValue<string>("emailaddress1");
                    break;
                case "Delete":
                    message = "Contact was deleted!";
                    break;
            }

            return message;
        }
        
        private string GetSubject(string messageType, Entity contact)
        {
            var name = contact.GetAttributeValue<string>("fullname");

            var message = "";
            switch (messageType)
            {
                case "Create":
                    message = "New Contact " + name + " created " + contact.GetAttributeValue<DateTime>("createdon");
                    break;
                case "Update":
                    message = "Contact " + name + " email address changed " + contact.GetAttributeValue<DateTime>("modifiedon");
                    break;
                case "Delete":
                    message = "Contact " + name + " was deleted " + contact.GetAttributeValue<DateTime>("modifiedon");
                    break;
            }

            return message;
        }
    }
}
