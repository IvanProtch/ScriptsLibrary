//Выбор исполнителя из группы
using System;
using Intermech.Interfaces;
using Intermech.Interfaces.Workflow;
using Intermech.Workflow;
using System.Xml.Linq;
using System.Xml;
public class Script
{

    public ICSharpScriptContext ScriptContext { get; private set; }

    /// <summary>
    /// Производит выбор одного пользователя из группы
    /// </summary>
    /// <param name="activity"></param>
    /// <param name="groupVarName">Название группы пользоватей (переменная процесса)</param>
    /// <param name="procAttributeSelector">Атрибут для выбора типа пользователя</param>
    private void SetOneUserFromGroup(IActivity activity, string groupVarName, string procAttributeSelector = "Инициатор процесса")
    {
        ParticipantList list = new ParticipantList();
        IVariable groupVar = activity.Variables.Find(groupVarName);
        IDBAttribute selector = activity.Process.Attributes.FindByName(procAttributeSelector);
        //		if(groupVar == null || selector == null)
        //			return;

        XElement listXml = XElement.Parse(selector.Value.ToString());
        //list.AddParticipant(ParticipantKind.User, activity.ParticipantID);
        list.AddParticipant(ParticipantKind.User, Convert.ToInt64(listXml.Value));
        // По умолчанию "получатель" записывается в группу пользователей Convert.ToInt64(listXml.Value)
        // ??? еще был вариант записывать текущего исполнителя activity.ParticipantID
        groupVar.Value = list.AsString;
        //MessageBox.Show(groupVar.Value);
    }

    public void Execute(IActivity activity)
    {

        string varible = "САПР (процесс) (БМЗ)";
        SetOneUserFromGroup(activity, varible);

        //MessageBox.Show(XElement.Parse(activity.Process.Attributes.FindByName("Получатель").Value.ToString()).Value.ToString());

    }
}