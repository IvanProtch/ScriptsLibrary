using Intermech;
using Intermech.Interfaces;
using Intermech.Interfaces.Compositions;
using Intermech.Interfaces.Contexts;
using Intermech.Interfaces.Workflow;
using Intermech.Kernel.Search;
using Intermech.Workflow;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Xml;
public class Script
{
    public ICSharpScriptContext ScriptContext { get; private set; }

    /// <summary>
    /// Список администраторов для отправки письма со специальными ошибками
    /// </summary>
    const string processVar_listOfAdmins = "список САПР";

    #region Константы
    /// <summary>
    /// Технологический объект
    /// </summary>
    const int techObjType = 1012;
    /// <summary>
    /// Техпроцесс базовый
    /// </summary>
    const int baseTPType = 1117;
    /// <summary>
    /// Материалы
    /// </summary>
    const int materialObjType = 1021;
    /// <summary>
    /// Тип извещения корневой
    /// </summary>
    const int IIType = 1066;
    /// <summary>
    /// Тип извещения конструкторское
    /// </summary>
    const int ConstrIIType = 2760;
    /// <summary>
    /// Взаимосвязанный контекст для ИИ
    /// </summary>
    const int connectedContectId = 12677;

    List<int> techObjTypeListAll;
    List<int> materialObjTypeListAll;
    List<int> TPListAll;
    List<int> IIListAll;
    List<int> TechIITypes;

    const int TechIIType = 2761;
    const int TechSvIIType = 2762;

    const int lcStep_Sogl = 1023;
    const int lcStep_Final = 1024;
    const int lcStep_Development = 1021;
    #endregion

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
    /// Отправка сообщения на почту пользователям из переменной
    /// </summary>
    /// <param name="activity"></param>
    /// <param name="UserSession"></param>
    /// <param name="message"></param>
    /// <param name="subject"></param>

    private IDBObject[] SendMessage(IActivity activity, IUserSession UserSession, string message, string subject)
    {
        //группа администраторов БМЗ САПР
        //3450327
        //3450583
        //51448525
        //8177509
        if (activity.Variables.Find(processVar_listOfAdmins) == null)
        {
            long[] toUsers_Default = new long[] { 3450327, 3450583, 51448525, 8177509 };
            return SendMessage(activity, UserSession, message, subject, toUsers_Default);
        }
        IRouterService router = UserSession.GetCustomService(typeof(IRouterService)) as IRouterService;

        ParticipantList users = new ParticipantList();
        users.AsString = activity.Variables.Find(processVar_listOfAdmins).Value;
        long[] toUsers = users.ObjectIDs.ToArray();

        IDBObject[] mailMesseges = router.CreateMessage(UserSession.SessionGUID, toUsers, subject, message, UserSession.UserID); ;

        foreach (IDBObject mail in mailMesseges)
            mail.GetAttributeByName("Процесс").AsInteger = activity.Process.ObjectID;

        return mailMesseges;
    }

    private IDBObject[] SendMessage(IActivity activity, IUserSession UserSession, string message, string subject, long[] toUsers)
    {
        IRouterService router = UserSession.GetCustomService(typeof(IRouterService)) as IRouterService;
        return router.CreateMessage(UserSession.SessionGUID, toUsers, subject, message, UserSession.UserID);
    }

    /// <summary>
    /// Найти объект, который есть в контексте, но не найден в составе объекта II
    /// </summary>
    /// <param name="UserSession"></param>
    /// <param name="II">Родительский объект (извещение об изменении)</param>
    /// <param name="consistance">Его состав</param>
    /// <returns></returns>

    private List<long> NotConsistedInII_ObjIDs(IUserSession UserSession, long II, List<long> consistance)
    {
        // сервис для работы с контекстами
        IDBEditingContextsService contextService = UserSession.GetCustomService(typeof(IDBEditingContextsService)) as IDBEditingContextsService;

        EditingContextsObjectContainer contextContainer = contextService.GetEditingContextsObject(UserSession.SessionGUID, II, true, true);

        consistance = consistance
        .Where(element => UserSession.GetObject(element).ModificationID == contextContainer.ModificationID)
        .ToList();

        List<long> context = contextContainer.GetVersionsID(false, UserSession.UserID);

        List<long> result = context
        .Where(element => !consistance.Contains(element))
        .ToList();

        return result;

    }

    /// <summary>
    /// Производит выбор объектов из БД по значению атрибута-селектора
    /// </summary>
    /// <param name="mainType">Тип искомого объекта</param>
    /// <param name="attrSelector"></param>
    /// <param name="attrSelValue"></param>
    /// <returns></returns>

    private List<long> SelectObjects(IUserSession UserSession, int mainType,
    int attrSelector, long attrSelValue)
    {
        IDBObjectCollection objects = UserSession.GetObjectCollection(mainType);

        ColumnDescriptor[] columns =
        {
                new ColumnDescriptor(ObligatoryObjectAttributes.F_OBJECT_ID,
                AttributeSourceTypes.Object, ColumnContents.ID, ColumnNameMapping.ID, SortOrders.ASC, 0),
                new ColumnDescriptor(ObligatoryObjectAttributes.F_OBJECT_TYPE,
                AttributeSourceTypes.Object, ColumnContents.ID, ColumnNameMapping.ID, SortOrders.NONE, 0)
            };
        ConditionStructure[] conds =
        {
                //new ConditionStructure((int)ObligatoryObjectAttributes.F_OBJECT_TYPE, RelationalOperators.In, types.ToArray(),
                //LogicalOperators.AND, 0, true),
            
                new ConditionStructure((int)attrSelector, RelationalOperators.Equal, (int)attrSelValue, LogicalOperators.NONE, 0, true)
            };
        DBRecordSetParams dbrsp = new DBRecordSetParams(conds, columns, 0, null, QueryConsts.All);

        DataTable tableObjs = objects.Select(dbrsp);
        List<long> idsL = tableObjs.Rows.
        OfType<DataRow>().
        Select(a => (long)a[0]).
        ToList();

        return idsL;
    }

    /// <summary>
    /// Получает связанные ИИ с выбранным ИИ
    /// </summary>
    /// <param name="UserSession"></param>
    /// <param name="constrIIid"></param>
    /// <returns></returns>

    private List<long> GetConnectedTechIIs(IUserSession UserSession, long IIid)
    {
        IDBObject mainII = UserSession.GetObject(IIid);
        IDBAttribute connectedContextAtr = mainII.GetAttributeByID(connectedContectId);

        List<long> techIIs = SelectObjects(UserSession, IIType,
        connectedContectId, Convert.ToInt64(connectedContextAtr.Value));

        techIIs = techIIs
        .Where(ii => TechIITypes.Contains(UserSession.GetObject(ii).TypeID))
        .Where(ii => ii != IIid)
        .ToList();

        return techIIs;
    }

    /// <summary>
    /// Получение списка ID состава/применяемости
    /// </summary>
    /// <param name="session"></param>
    /// <param name="objID"></param>
    /// <param name="relIDs"></param>
    /// <param name="childObjIDs"></param>
    /// <param name="lv">Количество уровней для обработки (-1 = загрузка полного состава / применяемости)</param>
    /// <returns></returns>

    private List<long> LoadItemIds(IUserSession session, long objID, List<int> relIDs, List<int> childObjIDs, int lv)
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
            return reslt;
        else
            return dt.Rows.OfType<DataRow>()
        .Select(element => (long)element[0]).ToList<long>();
    }

    /// <summary>
    /// Проверка атрибута "организация-источник БМЗ"
    /// </summary>
    private bool? FromBMZ(IUserSession UserSession, long elem)
    {
        const long idBMZ = 3251710; // id ообъекта организации "БМЗ"
        IDBObject obj = UserSession.GetObject(elem);
        IDBAttribute orgIstAtr = obj.GetAttributeByName("Организация-источник");
        if (orgIstAtr.IsNull)
        {
            return null;
        }
        long org = (long)orgIstAtr.Value;
        if (org == idBMZ)
            return true;
        else
            return false;
    }

    public void Execute(IActivity activity)
    {
        IUserSession UserSession = activity.Session;
        string FinalMessage = string.Empty;
        string mailToAdmins = string.Empty;
        // Инициализация списков типов объектов, используемых в запросах

        techObjTypeListAll = MetaDataHelper.GetObjectTypeChildrenIDRecursive(techObjType);
        materialObjTypeListAll = MetaDataHelper.GetObjectTypeChildrenIDRecursive(materialObjType);
        IIListAll = MetaDataHelper.GetObjectTypeChildrenIDRecursive(IIType);
        TPListAll = MetaDataHelper.GetObjectTypeChildrenIDRecursive(baseTPType);

        // Набор всех типов
        List<int> allObjectTypes = new List<int>();
        allObjectTypes.AddRange(techObjTypeListAll);
        allObjectTypes.AddRange(materialObjTypeListAll);
        allObjectTypes.AddRange(IIListAll);
        allObjectTypes.AddRange(TPListAll);

        TechIITypes = new List<int>();
        TechIITypes.Add(TechIIType);
        TechIITypes.Add(TechSvIIType);


        foreach (IAttachment attachment in activity.Attachments)
        {
            //Проходим только по конструкторским или технологическим ИИ
            if (MetaDataHelper.IsObjectTypeChildOf(attachment.ObjectType, ConstrIIType)
                || MetaDataHelper.IsObjectTypeChildOf(attachment.ObjectType, TechIIType)
            || MetaDataHelper.IsObjectTypeChildOf(attachment.ObjectType, TechSvIIType))
            {

                long mainII = attachment.ObjectID;
                //Получаем список, связанных с конструкторским ИИ, техн ИИ
                List<long> techIIs = GetConnectedTechIIs(UserSession, mainII);
                List<List<long>> techIIsAllConsistance = new List<List<long>>();

                //Получаем состав техн ИИ и заносим в список составов
                foreach (long techII in techIIs)
                {
                    List<long> consistance = LoadItemIds(UserSession, techII,
                    new List<int> { relTypeChangingByII, relTypeTechConsist },
                    allObjectTypes, -1).Distinct().ToList<long>();

                    techIIsAllConsistance.Add(consistance);
                }

                //Выполняем основные проверки ИИ
                if (techIIs.Count > 0)
                {
                    //Проходим по всем технологическим ИИ
                    for (int i = 0; i < techIIs.Count; i++)
                    {
                        //Прошло ли ИИ проверку?
                        bool isChecked = true;
                        IDBObject II = UserSession.GetObject(techIIs[i]);

                        //Проверяем редактируется ли само ИИ
                        if (II.CheckoutBy > 0)
                        {
                            IDBObject redactor = UserSession.GetObject(II.CheckoutBy);

                            FinalMessage += string.Format("\r\n[{0}] взял объект [{1}]|[{2}] на редактирование.\r\n",
                            redactor.NameInMessages, II.NameInMessages, II.ObjectID);
                            isChecked = false;
                        }

                        //Проверка принадлежности организации БМЗ
                        if (FromBMZ(UserSession, techIIs[i]) == null)
                        {
                            mailToAdmins += string.Format("\r\n[{0}]|[{1}] значение атрибута 'организация-источник' пустое," +
                            " проверьте принадлежность ИИ БМЗ и укажите значение атрибута и актуализируйте.\r\n",
                            II.NameInMessages, II.ObjectID);
                            isChecked = false;
                            continue;
                        }
                        //Пропускаем дальнейшую проверку для чужих ИИ
                        if (FromBMZ(UserSession, techIIs[i]) == false)
                            continue;

                        //Состав ИИ
                        List<long> consistance = techIIsAllConsistance[i];

                        //Проходим по составу
                        for (int ii_ind = 0; ii_ind < consistance.Count; ii_ind++)
                        {
                            IDBObject obj = UserSession.GetObject(consistance[ii_ind]);

                            IDBRelation relationII = UserSession.GetRelation(II.ObjectID, obj.ID);

                            if (relationII != null)
                            {
                                IDBAttribute reasonToEnterObj = relationII.GetAttributeByName("Цель включения объекта");
                                IDBAttribute customIDLC = relationII.GetAttributeByGuid(new Guid("cad01483-306c-11d8-b4e9-00304f19f545"
                                    /*Идентификатор шага ЖЦ, на который будет переведен объект*/));

                                if ((reasonToEnterObj.Description == "Создание" || reasonToEnterObj.Description == "Изменение")
                                    && Convert.ToInt32(customIDLC.Value) != 1058/*Шаг 'Согласование'*/)
                                {
                                    isChecked = false;
                                    FinalMessage += string.Format("Нельзя перевести объект {0} из {1} на шаг 'В разработке'\n",
                                        obj.NameInMessages, II.NameInMessages);
                                }
                            }

                            if (obj.CheckoutBy > 0)
                            {
                                IDBObject redactor = UserSession.GetObject(obj.CheckoutBy);

                                FinalMessage += string.Format("\r\n[{0}] взял объект [{1}]|[{3}] из [{2}]|[{4}] на редактирование.\r\n",
                                redactor.NameInMessages, obj.NameInMessages, II.NameInMessages, obj.ObjectID, II.ObjectID);

                                isChecked = false;
                            }
                        }

                        //Получаем список объектов не найденных в составе, но присутствующих в контексте ИИ
                        List<long> nConsistedObjIds = NotConsistedInII_ObjIDs(UserSession, techIIs[i], techIIsAllConsistance[i]);
                        foreach (long nConsistedObj in nConsistedObjIds)
                        {
                            IDBObject obj = UserSession.GetObject(nConsistedObj);

                            FinalMessage += string.Format("\r\nОбъект [{0}]|[{2}] присутствует в контексте редактирования [{1}]|[{3}]," +
                            " но не обнаружен в составе.\r\n",
                            obj.NameInMessages, II.NameInMessages, obj.ObjectID, II.ObjectID);

                            if (obj.CheckoutBy > 0)
                            {
                                IDBObject redactor = UserSession.GetObject(obj.CheckoutBy);

                                FinalMessage += string.Format("[{0}] взял объект [{1}]|[{2}] на редактирование.\r\n",
                                redactor.NameInMessages, obj.NameInMessages, obj.ObjectID);
                            }
                            isChecked = false;
                        }

                        //Переводим ИИ на шаг 'В разработке'
                        if (II.LCStep == lcStep_Sogl)
                        {

                            if (isChecked)
                            {
                                try
                                {
                                    II.LCStep = lcStep_Development;
                                }
                                catch (KernelException exc)
                                {
                                    mailToAdmins += string.Format("При переводе {0} на шаг 'В разработке' возникла системная ошибка: \r\n" +
                                    "Источник: {1}; \r\n" +
                                    "Сообщение: {2}; \r\n",
                                    II.NameInMessages, exc.Source, exc.Message);
                                }
                            }
                        }
                        else if ((II.LCStep != lcStep_Development) && (II.LCStep != lcStep_Sogl))
                        {
                            FinalMessage += string.Format("\r\nДля перевода на шаг 'В разработке' [{0}]|[{1}] необходимо," +
                            " чтобы ИИ находилось на шаге 'Согласование и утверждение'.\r\n",
                            II.NameInMessages, II.ObjectID);
                            isChecked = false;
                        }
                    }
                }

            }
        }
        if (FinalMessage.Length > 0)
        {
            throw new NotificationException(FinalMessage);
        }

        if (mailToAdmins.Length > 0)
        {
            SendMessage(activity, UserSession, mailToAdmins, "Ошибка перевода на шаг 'В разработке' технологических ИИ");
        }
    }

}
