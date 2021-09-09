//Проверка технологического ИИ на заполнение даты изменения в ERP 
using System;
using Intermech.Interfaces;
using Intermech.Interfaces.Workflow;

public class Script
{
    public ICSharpScriptContext ScriptContext { get; private set; }

    const int techIIType = 2761;
    const int techSvIIType = 2762;
    const string processVarName = "БМЗ_ДП";
    public void Execute(IActivity activity)
    {
        //для первого варианта раскомментировать:

        //IVariable procesVar = activity.Variables.Find(processVarName);
        //if ((bool)procesVar.TypedValue)
        //    return;

        string error = string.Empty;
        foreach (IAttachment attachment in activity.Attachments)
        {
            if(attachment.ObjectType == techIIType || attachment.ObjectType == techSvIIType)
            {
                var attr = attachment.Object.GetAttributeByGuid(new Guid("04436034-ae68-4eb6-b33f-c48b760f0e7a" /*Дата расчета МН по ИИ БМЗ*/));
                if (attr == null || attr.AsDateTime == null || attr.AsString == string.Empty)
                    error += string.Format("\nВ ИИ {0} не проставлена дата изменения в ERP\n", attachment.Object.Caption);
            }
        }

        if (error.Length > 0)
            throw new NotificationException(error);
    }
}