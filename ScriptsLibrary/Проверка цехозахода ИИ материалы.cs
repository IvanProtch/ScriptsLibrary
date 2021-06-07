//Проверка цехозахода ИИ материалы
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using Intermech;
using Intermech.Interfaces;
using Intermech.Interfaces.Compositions;
using Intermech.Interfaces.Workflow;
using Intermech.Kernel.Search;

public class Script
{
    public ICSharpScriptContext ScriptContext { get; set; }


    const int TechIIType = 2761;
    const int TechSvIIType = 2762;

    #region Типы связей
    const int relTypeChangingByII = 1007;//Изменяется по извещению
    const int relTypeDocNaIzd = 1004; //Документация на изделие
    const int relTypeConsist = 1; //Состав изделия
    const int relTypeTechConsist = 1002; //Технологический состав
    const int relTypeTechData = 1011; //Собираемые технологические данные
    const int relTypeTechConectScetch = 1012; //Технологическая связь с эскизом
    const int relTypeTP = 1036; //Сквозной ТП
    const int relTypeMO = 1060; //МО обработки изделия
    const int relTypeConection = 0; //Простая связь
    #endregion

    /// <summary>
    /// Получение списка ID состава/применяемости
    /// </summary>
    /// <param name="session"></param>
    /// <param name="objID"></param>
    /// <param name="relIDs"></param>
    /// <param name="childObjIDs"></param>
    /// <param name="lv">Количество уровней для обработки (-1 = загрузка полного состава / применяемости)</param>
    /// <returns></returns>
    private List<long> LoadItemIds(IUserSession session, long objID, List<int> relIDs, List<int> childObjIDs, int level, bool isConsist = true)
    {
        ICompositionLoadService _loadService = session.GetCustomService(typeof(ICompositionLoadService)) as ICompositionLoadService;
        List<ColumnDescriptor> col = new List<ColumnDescriptor>
        {
            new ColumnDescriptor(ObligatoryObjectAttributes.F_OBJECT_ID, AttributeSourceTypes.Object, ColumnContents.ID, ColumnNameMapping.Index, SortOrders.NONE, 0),
        };
        List<long> reslt = new List<long>();
        session.EditingContextID = objID;
        DataTable dt = _loadService.LoadComplexCompositions(
        session.SessionGUID,
        new List<ObjInfoItem> { new ObjInfoItem(objID) },
        relIDs,
        childObjIDs,
        col,
        isConsist, // Режим получения данных (true - состав, false - применяемость)
        false, // Флаг группировки данных
        null, // Правило подбора версий
        null, // Условия на объекты при получении состава / применяемости
        //SystemGUIDs.filtrationLatestVersions, // Настройки фильтрации объектов (последние)
        //SystemGUIDs.filtrationBaseVersions, // Настройки фильтрации объектов (базовые)
        Intermech.SystemGUIDs.filtrationLatestVersions,
        null,
        level // Количество уровней для обработки (-1 = загрузка полного состава / применяемости)
        );
        if (dt == null)
            return reslt;
        else
            return dt.Rows.OfType<DataRow>()
        .Select(element => (long)element[0]).ToList<long>();
    }

    public void Execute(IActivity activity)
    {
        IUserSession UserSession = activity.Session;
        string errorMessage = string.Empty;

        int materialGroupType = UserSession.GetObjectType("Группа материалов").ObjectType;

        foreach (IAttachment attachment in activity.Attachments)
        {
            if (MetaDataHelper.IsObjectTypeChildOf(attachment.ObjectType, TechIIType)
                || MetaDataHelper.IsObjectTypeChildOf(attachment.ObjectType, TechSvIIType))
            {
                List<long> consistance = LoadItemIds(UserSession, attachment.ObjectID,
                new List<int> { relTypeChangingByII, relTypeConsist },
                new List<int>() { materialGroupType }, -1).Distinct().ToList<long>();

                foreach (long item in consistance)
                {
                    IDBObject materialGroupObj = UserSession.GetObject(item);
                    IDBAttribute workShopAtr = materialGroupObj.GetAttributeByName("Код цеха");
                    if (workShopAtr == null || workShopAtr.Value.ToString() == string.Empty)
                        errorMessage += string.Format("Для объекта {0} не указано значение атрибута 'Цех-поставщик'.\n\r", materialGroupObj.NameInMessages);

                }
            }
        }

        if (errorMessage.Length > 0)
        {
            throw new NotificationException(errorMessage);
        }

    }
}