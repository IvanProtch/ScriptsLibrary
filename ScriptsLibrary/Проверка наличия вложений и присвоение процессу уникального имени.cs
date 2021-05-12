using System;
using Intermech.Interfaces;
using Intermech.Interfaces.Workflow;
using Intermech;
using System.Diagnostics;

public class Script
{
    public ICSharpScriptContext ScriptContext { get; private set; }

    public void Execute(IActivity activity)
    {
        if (Debugger.IsAttached)
            Debugger.Break();
        // Если нет вложений
        if (activity.Attachments.Count == 0)
            throw new NotificationException(String.Format("Вложения для процесса '{0}' не добавлены. Нажмите 'назад' и повторите выбор.", activity.Process.Caption));

        IScheme scheme = activity.Process;

        bool c1 = (bool)activity.Variables.Find("БМЗ_ИИ_Т_БСбтепл").TypedValue;
        bool c2 = (bool)activity.Variables.Find("БМЗ_ИИ_Т_БСбтележек").TypedValue;
        bool c3 = (bool)activity.Variables.Find("БМЗ_ИИ_Т_БЧПУ").TypedValue;
        bool c4 = (bool)activity.Variables.Find("БМЗ_ИИ_Т_БЗПиН").TypedValue;
        bool c5 = (bool)activity.Variables.Find("БМЗ_ИИ_Т_БТо").TypedValue;
        bool c6 = (bool)activity.Variables.Find("БМЗ_ИИ_Т_согласов. УГТ не требуется").TypedValue;

        if (((c1 || c2 || c3 || c4 || c5) && c6) || (c1 && c2 && c3 && c4 && c5 && c6) || !(c1 || c2 || c3 || c4 || c5 || c6))
            throw new NotificationException("Форма старта заполнена некорректно.");

        //Если есть, формируем наименование
        SetUniqueProcessCaption(activity);
    }

    /// <summary>
    /// Формирует уникальное наименование для запущенного процесса
    /// </summary>
    /// <param name="activity"></param>
    /// <param name="addTemplateNameAsPrefix">Добавлять ли в начале название шаблона процесса</param>
    private void SetUniqueProcessCaption(IActivity activity, bool addTemplateNameAsPrefix = true)
    {
        IAttachments attachments = activity.Attachments;
        if (attachments.Count == 0)
            return;

        string newCaption = string.Empty;
        if (addTemplateNameAsPrefix)
        {
            IDBObject process = activity.Session.GetObject(activity.Process.ObjectID);
            IDBAttribute templateName = process.GetAttributeByName("Родительский шаблон");
            newCaption = templateName.AsString + "_";
        }

        if (attachments.Count == 1)
            newCaption += String.Format("{0}", attachments[0].Object.Caption);
        else
        {
            string manyAttachs = string.Empty;
            for (int i = 0; i < attachments.Count; i++)
            {
                manyAttachs += attachments[i].Object.Caption;
                if (i != attachments.Count - 1)
                    manyAttachs += "; ";
            }
            newCaption += String.Format("[{0}]", manyAttachs);
        }
        if (newCaption.Length > 250)
        {
            newCaption = newCaption.Remove(249);
            newCaption = newCaption.Remove(newCaption.LastIndexOf(';'));
        }

        activity.Process.Caption = newCaption;
    }
}
