//Проверка технологического ИИ на заполнение даты изменения в ERP 
using System;
using Intermech.Interfaces;
using Intermech.Interfaces.Workflow;

public class Script
{
    public ICSharpScriptContext ScriptContext { get; private set; }

    const int techIIType = 2761;
    const int techSvIIType = 2762;
    public void Execute(IActivity activity)
    {
        string error = string.Empty;
        foreach (IAttachment attachment in activity.Attachments)
        {
            if (attachment.ObjectType == techIIType || attachment.ObjectType == techSvIIType)
            {
                var attr1 = attachment.Object.GetAttributeByGuid(new Guid("778348ff-f60b-4e5f-ad42-9205f60cdb27" /*Дата изменения в ERP БМЗ*/));
                var attr2 = attachment.Object.GetAttributeByGuid(new Guid("ce88b31f-b918-4154-9c9f-7798049dda8f" /*Указание о внедрении*/));
                bool attr1IsNotExists = attr1 == null || attr1.AsDateTime == null || attr1.AsString == string.Empty;
                bool attr2IsNotExists = attr2 == null || attr2.AsDateTime == null || attr2.AsString == string.Empty;

                if (attr1IsNotExists || attr2IsNotExists)
                    error += string.Format("\nВ ИИ {0} нет информации об указании о внедрении или дате изменения в ERP\n", attachment.Object.Caption);
            }
        }

        if (error.Length > 0)
            throw new NotificationException(error);
    }
}