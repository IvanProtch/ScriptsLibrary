using System;
using Intermech.Interfaces;
using Intermech.Interfaces.Workflow;
using Intermech;
using System.Diagnostics;
using Intermech.Interfaces.Compositions;
using System.Collections.Generic;
using Intermech.Kernel.Search;
using System.Data;
using System.Linq;

public class Script
{
	public ICSharpScriptContext ScriptContext { get; private set; }

	public void Execute(IActivity activity)
	{
		// Если нет вложений
		if (activity.Attachments.Count == 0)
			throw new NotificationException(String.Format("Вложения для процесса '{0}' не добавлены. Нажмите 'назад' и повторите выбор.", activity.Process.Caption));

		// Проверяет состав любого объекта, если удается найти объект взятый на редактирование, формирует сообщение об ошибке.
		string error = CheckAllConsistanceCheckout(activity);
		if (error.Length > 0)
			throw new ClientException(error);
		
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

	/// <summary>
	/// true - есть хотя бы один объект взятый на редактирование
	/// </summary>
	/// <param name="activity"></param>
	/// <param name="error"></param>
	/// <returns></returns>
	private string CheckAllConsistanceCheckout(IActivity activity)
    {
		List<int> allRelations = MetaDataHelper.GetRelationTypesList().Select(rel => rel.RelationTypeID).ToList();
	    string error = string.Empty;
		foreach (IAttachment attachment in activity.Attachments)
		{
			List<IDBObject> consistance = GetObjectConsistance(activity.Session, attachment.ObjectID, allRelations, null, -1);
			foreach (IDBObject item in consistance)
			{
				if (item.CheckoutBy > 0)
                {
					error += string.Format("У объекта '{0}' не может быть изменен шаг жизненного цикла, пока объект '{1}' находится на редактировании.\r\n", attachment.Object.NameInMessages, item.NameInMessages);
				}
			}
		}
		return error;
	}

	/// <summary>
	/// Получение списка состава/применяемости
	/// </summary>
	/// <param name="session"></param>
	/// <param name="objID"></param>
	/// <param name="relIDs"></param>
	/// <param name="childObjIDs"></param>
	/// <param name="lv">Количество уровней для обработки (-1 = загрузка полного состава / применяемости)</param>
	/// <returns></returns>
	private List<IDBObject> GetObjectConsistance(IUserSession session, long objID, List<int> relIDs, List<int> childObjIDs, int lv)
	{
		ICompositionLoadService _loadService = session.GetCustomService(typeof(ICompositionLoadService)) as ICompositionLoadService;
		List<ColumnDescriptor> col = new List<ColumnDescriptor>
			{
				new ColumnDescriptor(ObligatoryObjectAttributes.F_OBJECT_ID, AttributeSourceTypes.Object, ColumnContents.ID, ColumnNameMapping.Index, SortOrders.NONE, 0),
			};
		session.EditingContextID = objID;
		DataTable dt = _loadService.LoadComplexCompositions(
		session.SessionGUID,
		new List<ObjInfoItem> { new ObjInfoItem(objID) },
		relIDs,
		childObjIDs,
		col,
		true, // Режим получения данных (true - состав, false - применяемость)
		false, // Флаг группировки данных
		null, // Правило подбора версий
		null, // Условия на объекты при получении состава / применяемости
			  //SystemGUIDs.filtrationLatestVersions, // Настройки фильтрации объектов (последние)
			  //SystemGUIDs.filtrationBaseVersions, // Настройки фильтрации объектов (базовые)
		Intermech.SystemGUIDs.filtrationLatestVersions,
		null,
		lv // Количество уровней для обработки (-1 = загрузка полного состава / применяемости)
		);
		if (dt == null)
			return null;
		else
			return dt.Rows.OfType<DataRow>()
		.Select(element => session.GetObject((long)element[0])).ToList();
	}
}
